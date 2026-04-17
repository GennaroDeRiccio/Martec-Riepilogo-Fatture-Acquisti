from __future__ import annotations

import csv
import copy
import html
import io
import json
import re
import sqlite3
import uuid
import zlib
from dataclasses import dataclass
from datetime import datetime, timedelta
from decimal import Decimal, InvalidOperation
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from zipfile import ZIP_DEFLATED, ZipFile
from xml.etree import ElementTree as ET


ROOT = Path(__file__).resolve().parent
STATIC_DIR = ROOT / "static"
DATA_DIR = ROOT / "data"
UPLOAD_DIR = DATA_DIR / "uploads"
DB_FILE = DATA_DIR / "martec.sqlite3"
LEGACY_RECORDS_FILE = DATA_DIR / "records.json"
BRAND_LOGO = Path("/Users/gennydericcio/Desktop/Hub/Martec/Documenti/SCR-20250724-jpqy.png")
EXCEL_TEMPLATE = Path("/Users/gennydericcio/Desktop/RIEPILOGO FATTURE MARTEC agg. 27-02-2026 formula cambio corretta.xlsx")

EXCEL_COLUMNS = [
    "Num.",
    "Cliente",
    "Data",
    "Fattura",
    "Valore in USD",
    "Valore in Euro",
    "Valore",
    "IVA",
    "Totale",
    "Entrate",
    "Delta incasso",
    "Data fattura",
    "Scadenza",
    "Incasso avvenuto",
    "BANCA - C/C",
    "Termini pagamento fattura",
    "Note",
]

EXPORT_TEMPLATE_COLUMNS = [
    "Num.",
    "Cliente",
    "Data",
    "Fattura",
    "Valore in USD",
    "Valore in Euro",
    "Valore",
    "IVA",
    "Totale",
    "Entrate",
    "Delta incasso",
    "Fattura anno 2019 - 2020-2021-2022 - 2023",
    "Scadenza",
    "Incasso avvenuto",
    "BANCA - C/C",
    "Termini pagamento fattura",
    "NOTE VARIE",
    "Ulteriori Note",
    "Cambio",
]

SHEET_MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
SHEET_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ET.register_namespace("", SHEET_MAIN_NS)
ET.register_namespace("r", SHEET_REL_NS)
ET.register_namespace("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")
ET.register_namespace("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main")
ET.register_namespace("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision")
ET.register_namespace("xr6", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6")
ET.register_namespace("xr10", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10")
ET.register_namespace("xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2")

NS = {"a": SHEET_MAIN_NS, "r": SHEET_REL_NS}

STATUS_OK = "Pagato"
STATUS_VERIFY = "Da pagare"
OLD_DATE_COLUMN = "Fattura anno 2019 - 2020-2021-2022 - 2023"
OLD_NOTE_COLUMN = "NOTE VARIE"
OLD_EXTRA_NOTE_COLUMN = "Ulteriori Note"


@dataclass
class TextItem:
    y: float
    x: float
    text: str


def ensure_dirs() -> None:
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)


def connect() -> sqlite3.Connection:
    ensure_dirs()
    db = sqlite3.connect(DB_FILE)
    db.row_factory = sqlite3.Row
    db.execute("PRAGMA foreign_keys = ON")
    return db


def init_db() -> None:
    with connect() as db:
        db.executescript(
            """
            CREATE TABLE IF NOT EXISTS records (
                id TEXT PRIMARY KEY,
                created_at TEXT NOT NULL,
                row_json TEXT NOT NULL,
                invoice_json TEXT NOT NULL,
                transfer_json TEXT NOT NULL,
                checks_json TEXT NOT NULL,
                source TEXT NOT NULL DEFAULT 'upload',
                invoice_key TEXT,
                status TEXT NOT NULL DEFAULT 'Da verificare'
            );

            CREATE UNIQUE INDEX IF NOT EXISTS idx_records_invoice_key
            ON records(invoice_key)
            WHERE invoice_key IS NOT NULL AND invoice_key <> '';

            CREATE TABLE IF NOT EXISTS suppliers (
                id TEXT PRIMARY KEY,
                name TEXT NOT NULL,
                vat TEXT,
                iban TEXT,
                swift TEXT,
                bank TEXT,
                currency TEXT,
                notes TEXT,
                updated_at TEXT NOT NULL
            );

            CREATE UNIQUE INDEX IF NOT EXISTS idx_suppliers_name
            ON suppliers(upper(name));
            """
        )
    migrate_legacy_json()


def migrate_legacy_json() -> None:
    if not LEGACY_RECORDS_FILE.exists():
        return
    with connect() as db:
        count = db.execute("SELECT COUNT(*) FROM records").fetchone()[0]
        if count:
            return
        try:
            records = json.loads(LEGACY_RECORDS_FILE.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            return
        for record in records:
            save_record(db, record, allow_duplicate=False)


def json_dump(value: Any) -> str:
    return json.dumps(value, ensure_ascii=False, separators=(",", ":"))


def row_from_db(db_row: sqlite3.Row) -> dict[str, Any]:
    row = normalize_row(json.loads(db_row["row_json"]))
    status = normalize_status(db_row["status"])
    return {
        "id": db_row["id"],
        "createdAt": db_row["created_at"],
        "row": row,
        "invoice": json.loads(db_row["invoice_json"]),
        "transfer": json.loads(db_row["transfer_json"]),
        "checks": json.loads(db_row["checks_json"]),
        "source": db_row["source"],
        "status": status,
    }


def normalize_status(value: str | None) -> str:
    return STATUS_OK if value == STATUS_OK else STATUS_VERIFY


def normalize_paid_value(value: str | None) -> str:
    text = str(value or "").strip().upper()
    return "✅" if text in {"X", "XX", "✅", "SI", "SÌ", "PAGATO", "PAGATA"} else "❌"


def normalize_row(row: dict[str, Any]) -> dict[str, str]:
    row = {str(key): "" if value is None else str(value) for key, value in row.items()}
    if "Data fattura" not in row:
        row["Data fattura"] = row.get(OLD_DATE_COLUMN, "")
    if "Note" not in row:
        notes = [row.get(OLD_NOTE_COLUMN, ""), row.get(OLD_EXTRA_NOTE_COLUMN, "")]
        row["Note"] = " - ".join(note for note in notes if note)
    row["Incasso avvenuto"] = normalize_paid_value(row.get("Incasso avvenuto"))
    normalized = {column: row.get(column, "") for column in EXCEL_COLUMNS}
    for column in EURO_COLUMNS:
        normalized[column] = format_currency(normalized.get(column, ""), "€")
    normalized["Valore in USD"] = format_currency(normalized.get("Valore in USD", ""), "$")
    return normalized


def load_records() -> list[dict[str, Any]]:
    init_db()
    with connect() as db:
        rows = db.execute("SELECT * FROM records ORDER BY CAST(json_extract(row_json, '$.\"Num.\"') AS INTEGER), created_at").fetchall()
        return [row_from_db(row) for row in rows]


def load_suppliers() -> list[dict[str, str]]:
    init_db()
    with connect() as db:
        rows = db.execute("SELECT * FROM suppliers ORDER BY name").fetchall()
        return [dict(row) for row in rows]


def next_number(db: sqlite3.Connection) -> int:
    rows = db.execute("SELECT row_json FROM records").fetchall()
    numbers = []
    for row in rows:
        try:
            numbers.append(int(json.loads(row["row_json"]).get("Num.", 0)))
        except (TypeError, ValueError, json.JSONDecodeError):
            pass
    return max(numbers, default=0) + 1


def invoice_key(record: dict[str, Any]) -> str:
    row = record.get("row", {})
    supplier = normalize_key(row.get("Cliente", ""))
    number = normalize_key(row.get("Fattura", ""))
    date = normalize_key(row.get("Data", ""))
    total = normalize_key(row.get("Totale", ""))
    if not supplier or not number:
        return ""
    return f"{supplier}|{number}|{date}|{total}"


def save_record(db: sqlite3.Connection, record: dict[str, Any], allow_duplicate: bool = False) -> tuple[bool, str]:
    key = invoice_key(record)
    if key and not allow_duplicate:
        existing = db.execute("SELECT id FROM records WHERE invoice_key = ?", (key,)).fetchone()
        if existing:
            return False, existing["id"]

    row = record.get("row", {})
    checks = record.get("checks", [])
    row = normalize_row(row)
    status = record.get("status") or (STATUS_OK if row.get("Incasso avvenuto") == "✅" else STATUS_VERIFY)
    db.execute(
        """
        INSERT INTO records (id, created_at, row_json, invoice_json, transfer_json, checks_json, source, invoice_key, status)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            record.get("id") or str(uuid.uuid4()),
            record.get("createdAt") or datetime.now().isoformat(timespec="seconds"),
            json_dump(row),
            json_dump(record.get("invoice", {})),
            json_dump(record.get("transfer", {})),
            json_dump(checks),
            record.get("source", "upload"),
            key,
            status,
        ),
    )
    upsert_supplier(db, record)
    return True, record["id"]


def update_record_row(record_id: str, row: dict[str, str]) -> dict[str, Any] | None:
    init_db()
    with connect() as db:
        existing = db.execute("SELECT * FROM records WHERE id = ?", (record_id,)).fetchone()
        if not existing:
            return None
        record = row_from_db(existing)
        merged = normalize_row({column: row.get(column, record["row"].get(column, "")) for column in EXCEL_COLUMNS})
        checks = checks_from_row(merged)
        status = normalize_status(row.get("Stato") or payload_status_from_row(row) or (STATUS_OK if merged.get("Incasso avvenuto") == "✅" else STATUS_VERIFY))
        db.execute(
            """
            UPDATE records
            SET row_json = ?, checks_json = ?, status = ?, invoice_key = ?
            WHERE id = ?
            """,
            (json_dump(merged), json_dump(checks), status, invoice_key({"row": merged}), record_id),
        )
        updated = db.execute("SELECT * FROM records WHERE id = ?", (record_id,)).fetchone()
        return row_from_db(updated) if updated else None


def clear_records() -> None:
    init_db()
    with connect() as db:
        db.execute("DELETE FROM records")


def payload_status_from_row(row: dict[str, str]) -> str:
    return normalize_status(row.get("Stato")) if row.get("Stato") else ""


def upsert_supplier(db: sqlite3.Connection, record: dict[str, Any]) -> None:
    row = record.get("row", {})
    invoice = record.get("invoice", {})
    transfer = record.get("transfer", {})
    name = row.get("Cliente") or invoice.get("supplier") or transfer.get("beneficiary")
    if not name:
        return
    now = datetime.now().isoformat(timespec="seconds")
    existing = db.execute("SELECT id FROM suppliers WHERE upper(name) = upper(?)", (name,)).fetchone()
    payload = {
        "id": existing["id"] if existing else str(uuid.uuid4()),
        "name": name,
        "vat": invoice.get("supplierVat", ""),
        "iban": invoice.get("iban") or transfer.get("beneficiaryIban", ""),
        "swift": transfer.get("swift", ""),
        "bank": row.get("BANCA - C/C", ""),
        "currency": "USD" if row.get("Valore in USD") else "EUR",
        "notes": "",
        "updated_at": now,
    }
    db.execute(
        """
        INSERT INTO suppliers (id, name, vat, iban, swift, bank, currency, notes, updated_at)
        VALUES (:id, :name, :vat, :iban, :swift, :bank, :currency, :notes, :updated_at)
        ON CONFLICT(upper(name)) DO UPDATE SET
            vat = excluded.vat,
            iban = excluded.iban,
            swift = excluded.swift,
            bank = excluded.bank,
            currency = excluded.currency,
            updated_at = excluded.updated_at
        """,
        payload,
    )


def decimal_from_it(value: str | None) -> Decimal | None:
    if not value:
        return None
    cleaned = value.strip().replace("EUR", "").replace("USD", "").replace("€", "").replace("$", "").replace(" ", "")
    cleaned = cleaned.replace(".", "").replace(",", ".")
    try:
        return Decimal(cleaned)
    except InvalidOperation:
        return None


def decimal_to_it(value: Decimal | None) -> str:
    if value is None:
        return ""
    quantized = value.quantize(Decimal("0.01"))
    return f"{quantized:.2f}".replace(".", ",")


EURO_COLUMNS = {"Valore in Euro", "Valore", "IVA", "Totale", "Entrate", "Delta incasso"}


def format_currency(value: str | Decimal | None, symbol: str) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if not text:
        return ""
    amount = decimal_from_it(text)
    if amount is None:
        return text
    return f"{decimal_to_it(amount)} {symbol}"


def normalize_date(value: str | None) -> str:
    if not value:
        return ""
    text = value.strip()
    for fmt in ("%d-%m-%Y", "%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt).strftime("%d/%m/%Y")
        except ValueError:
            pass
    return text


def add_days(date_text: str, days: int) -> str:
    try:
        date = datetime.strptime(normalize_date(date_text), "%d/%m/%Y")
    except ValueError:
        return ""
    return (date + timedelta(days=days)).strftime("%d/%m/%Y")


def normalize_key(value: str) -> str:
    return re.sub(r"[^A-Z0-9]+", "", str(value).upper())


def pdf_streams(data: bytes) -> list[bytes]:
    streams: list[bytes] = []
    for raw in re.findall(rb"stream\r?\n(.*?)\r?\nendstream", data, re.S):
        try:
            streams.append(zlib.decompress(raw))
        except zlib.error:
            streams.append(raw)
    return streams


def parse_cmap(stream: bytes) -> dict[int, str]:
    text = stream.decode("latin1", errors="ignore")
    cmap: dict[int, str] = {}
    for start, end, dest in re.findall(r"<([0-9A-Fa-f]{4})><([0-9A-Fa-f]{4})><([0-9A-Fa-f]{4})>", text):
        start_i = int(start, 16)
        end_i = int(end, 16)
        dest_i = int(dest, 16)
        for code in range(start_i, end_i + 1):
            cmap[code] = chr(dest_i + code - start_i)
    return cmap


def read_pdf_literal(buffer: bytes, start: int) -> tuple[bytes, int]:
    i = start + 1
    depth = 1
    out = bytearray()
    while i < len(buffer) and depth:
        byte = buffer[i]
        if byte == 92:
            i += 1
            if i >= len(buffer):
                break
            escaped = buffer[i]
            replacements = {
                ord("n"): 10,
                ord("r"): 13,
                ord("t"): 9,
                ord("b"): 8,
                ord("f"): 12,
                ord("("): 40,
                ord(")"): 41,
                ord("\\"): 92,
            }
            if escaped in replacements:
                out.append(replacements[escaped])
                i += 1
            elif 48 <= escaped <= 55:
                octal = bytes([escaped])
                i += 1
                for _ in range(2):
                    if i < len(buffer) and 48 <= buffer[i] <= 55:
                        octal += bytes([buffer[i]])
                        i += 1
                out.append(int(octal, 8))
            elif escaped in (10, 13):
                if escaped == 13 and i + 1 < len(buffer) and buffer[i + 1] == 10:
                    i += 2
                else:
                    i += 1
            else:
                out.append(escaped)
                i += 1
        elif byte == 40:
            depth += 1
            out.append(byte)
            i += 1
        elif byte == 41:
            depth -= 1
            if depth:
                out.append(byte)
            i += 1
        else:
            out.append(byte)
            i += 1
    return bytes(out), i


def decode_literal(raw: bytes, cmap: dict[int, str]) -> str:
    if cmap and len(raw) >= 2 and raw[0] == 0:
        chars = []
        for idx in range(0, len(raw) - 1, 2):
            code = (raw[idx] << 8) + raw[idx + 1]
            chars.append(cmap.get(code, ""))
        return "".join(chars)
    return raw.decode("latin1", errors="replace")


def extract_pdf_items(file_data: bytes) -> list[TextItem]:
    streams = pdf_streams(file_data)
    if not streams:
        return []

    cmap: dict[int, str] = {}
    for stream in streams:
        if b"beginbfrange" in stream:
            cmap.update(parse_cmap(stream))

    content = streams[0]
    items: list[TextItem] = []
    current_x = 0.0
    current_y = 0.0
    token_re = re.compile(rb"([\d\-.]+) ([\d\-.]+) ([\d\-.]+) ([\d\-.]+) ([\d\-.]+) ([\d\-.]+) Tm|\(")
    for match in token_re.finditer(content):
        if match.group(0).endswith(b"Tm"):
            current_x = float(match.group(5))
            current_y = float(match.group(6))
            continue
        raw, end = read_pdf_literal(content, match.start())
        if re.match(rb"\s*Tj", content[end : end + 20]):
            decoded = decode_literal(raw, cmap).strip()
            if decoded:
                items.append(TextItem(y=current_y, x=current_x, text=decoded))
    return sorted(items, key=lambda item: (-item.y, item.x))


def clean_text(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def lines_from_items(items: list[TextItem], tolerance: float = 1.2) -> list[str]:
    rows: list[list[TextItem]] = []
    for item in items:
        if not rows or abs(rows[-1][0].y - item.y) > tolerance:
            rows.append([item])
        else:
            rows[-1].append(item)
    lines = []
    for row in rows:
        parts = [clean_text(item.text) for item in sorted(row, key=lambda item: item.x)]
        line = clean_text(" ".join(part for part in parts if part))
        if line:
            lines.append(line)
    return lines


def right_value(items: list[TextItem], label: str, min_x: float | None = None) -> str:
    label_lower = label.lower()
    for item in items:
        if label_lower in item.text.lower():
            same_row = [
                other
                for other in items
                if abs(other.y - item.y) <= 1.2 and other.x > item.x + 20 and other.text.strip() != item.text.strip()
            ]
            if min_x is not None:
                same_row = [other for other in same_row if other.x >= min_x]
            if same_row:
                values = [other.text for other in sorted(same_row, key=lambda other: other.x)]
                filtered = []
                for value in values:
                    cleaned = clean_text(value)
                    if cleaned and not any(cleaned != other and cleaned in other for other in values):
                        filtered.append(cleaned)
                return clean_text(" ".join(filtered))
    return ""


def value_on_same_row(
    items: list[TextItem],
    label: str,
    *,
    label_max_x: float | None = None,
    value_min_x: float | None = None,
    value_max_x: float | None = None,
) -> str:
    label_lower = label.lower()
    for item in items:
        if label_lower not in item.text.lower():
            continue
        if label_max_x is not None and item.x > label_max_x:
            continue
        same_row = [
            other
            for other in items
            if abs(other.y - item.y) <= 1.2
            and other.x > item.x + 8
            and (value_min_x is None or other.x >= value_min_x)
            and (value_max_x is None or other.x <= value_max_x)
        ]
        if same_row:
            return clean_text(" ".join(other.text for other in sorted(same_row, key=lambda other: other.x)))
    return ""


def money_below_label(items: list[TextItem], label: str, *, x_min: float, x_max: float) -> str:
    label_items = [item for item in items if label.lower() in item.text.lower()]
    if not label_items:
        return ""
    label_y = label_items[0].y
    candidates = [
        item
        for item in items
        if item.y < label_y
        and x_min <= item.x <= x_max
        and re.fullmatch(r"-?\d+(?:\.\d{3})*,\d{2}", item.text.strip())
    ]
    if not candidates:
        return ""
    candidates.sort(key=lambda item: (label_y - item.y, abs(((x_min + x_max) / 2) - item.x)))
    return candidates[0].text.strip()


def first_match(pattern: str, text: str, flags: int = re.I) -> str:
    match = re.search(pattern, text, flags)
    return clean_text(match.group(1)) if match else ""


def parse_invoice(file_data: bytes) -> dict[str, str]:
    items = extract_pdf_items(file_data)
    full_text = "\n".join(lines_from_items(items))

    supplier = value_on_same_row(items, "Denominazione:", label_max_x=80, value_min_x=70, value_max_x=260)
    vat_ids = re.findall(r"Identificativo fiscale ai fini IVA:\s+([A-Z]{2}\d+)", full_text)
    supplier_vat = vat_ids[0] if vat_ids else ""
    invoice_number = first_match(r"Numero documento\s+Data documento.*?\n.*?([A-Z0-9./ -]{2,})\s+\d{2}-\d{2}-\d{4}", full_text, re.I | re.S)
    invoice_date = first_match(r"Numero documento\s+Data documento.*?\n.*?[A-Z0-9./ -]{2,}\s+(\d{2}-\d{2}-\d{4})", full_text, re.I | re.S)
    taxable = money_below_label(items, "Totale imponibile", x_min=430, x_max=510)
    vat = money_below_label(items, "Totale imposta", x_min=525, x_max=585)
    total = money_below_label(items, "Totale documento", x_min=520, x_max=585)
    if not taxable:
        taxable = first_match(r"Totale imponibile\s+Totale imposta.*?\n.*?\d+,\d{2}\s+(\d+,\d{2})\s+\d+,\d{2}", full_text, re.I | re.S)
    if not vat:
        vat = first_match(r"Totale imponibile\s+Totale imposta.*?\n.*?\d+,\d{2}\s+\d+,\d{2}\s+(\d+,\d{2})", full_text, re.I | re.S)
    if not total:
        total = right_value(items, "Totale documento")

    iban = first_match(r"\b(IT\d{2}[A-Z]\d{22})\b", full_text)
    due_date = first_match(r"Data scadenza\s+Importo.*?\n.*?(\d{2}-\d{2}-\d{4})", full_text, re.I | re.S)
    payment_terms = "Bonifico" if re.search(r"\bBonifico\b", full_text, re.I) else right_value(items, "Modalità pagamento")

    return {
        "type": "invoice",
        "supplier": supplier,
        "supplierVat": supplier_vat,
        "invoiceNumber": invoice_number,
        "invoiceDate": normalize_date(invoice_date),
        "taxable": taxable,
        "vat": vat,
        "total": total,
        "iban": iban,
        "dueDate": normalize_date(due_date),
        "paymentTerms": payment_terms,
        "rawText": full_text,
    }


def parse_transfer(file_data: bytes) -> dict[str, str]:
    items = extract_pdf_items(file_data)
    full_text = "\n".join(lines_from_items(items))

    document_date = first_match(r"Data:\s+(\d{2}\.\d{2}\.\d{4})", full_text)
    bank = first_match(r"(INTESA SANPAOLO S\.P\.A\.)", full_text)
    payer = right_value(items, "Ragione Sociale:", min_x=130)
    execution_date = first_match(r"Data esecuzione:\s+(\d{2}\.\d{2}\.\d{4})", full_text)
    total = first_match(r"Totale:\s+(\d+,\d{2})\s+EUR", full_text)
    beneficiary = right_value(items, "Beneficiario", min_x=150)
    beneficiary_iban = first_match(r"Conto beneficiario\s+(IT\d{2}[A-Z]\d{22})", full_text)
    swift = first_match(r"Codice SWIFT\s+([A-Z0-9]+)", full_text)
    reason = first_match(r"Informazioni aggiuntive \(max\s+(.+)", full_text)

    return {
        "type": "transfer",
        "documentDate": normalize_date(document_date),
        "bank": bank,
        "payer": payer,
        "executionDate": normalize_date(execution_date),
        "total": total,
        "beneficiary": beneficiary,
        "beneficiaryIban": beneficiary_iban,
        "swift": swift,
        "reason": reason,
        "rawText": full_text,
    }


def classify_pdf(file_data: bytes) -> dict[str, str]:
    invoice = parse_invoice(file_data)
    transfer = parse_transfer(file_data)
    invoice_score = sum(bool(invoice.get(key)) for key in ("supplier", "invoiceNumber", "invoiceDate", "total"))
    transfer_score = sum(bool(transfer.get(key)) for key in ("beneficiary", "beneficiaryIban", "reason", "executionDate", "total"))
    return transfer if transfer_score > invoice_score else invoice


def checks_from_row(row: dict[str, str]) -> list[dict[str, Any]]:
    total = decimal_from_it(row.get("Totale"))
    paid = decimal_from_it(row.get("Entrate"))
    delta = total - paid if total is not None and paid is not None else None
    row["Delta incasso"] = decimal_to_it(delta)
    return [
        {"label": "Importo", "ok": bool(delta is not None and delta == Decimal("0.00"))},
        {"label": "Pagamento", "ok": row.get("Incasso avvenuto", "") == "✅"},
    ]


def build_record(invoice: dict[str, str], transfer: dict[str, str] | None, index: int) -> dict[str, Any]:
    transfer = transfer or {}
    invoice_total = decimal_from_it(invoice.get("total"))
    transfer_total = decimal_from_it(transfer.get("total"))
    taxable = decimal_from_it(invoice.get("taxable"))
    delta = invoice_total - transfer_total if invoice_total is not None and transfer_total is not None else None
    paid = bool(delta is not None and delta == Decimal("0.00"))
    iban_matches = bool(invoice.get("iban") and invoice.get("iban") == transfer.get("beneficiaryIban"))
    invoice_in_reason = bool(invoice.get("invoiceNumber") and invoice.get("invoiceNumber") in transfer.get("reason", ""))

    checks = [
        {"label": "Importo", "ok": paid},
        {"label": "IBAN", "ok": iban_matches},
        {"label": "Numero fattura in causale", "ok": invoice_in_reason},
    ]

    row = {
        "Num.": str(index),
        "Cliente": invoice.get("supplier") or transfer.get("beneficiary", ""),
        "Data": invoice.get("invoiceDate", ""),
        "Fattura": invoice.get("invoiceNumber", ""),
        "Valore in USD": "",
        "Valore in Euro": decimal_to_it(taxable),
        "Valore": decimal_to_it(taxable),
        "IVA": invoice.get("vat", ""),
        "Totale": invoice.get("total", ""),
        "Entrate": transfer.get("total", ""),
        "Delta incasso": decimal_to_it(delta),
        "Data fattura": invoice.get("invoiceDate", ""),
        "Scadenza": invoice.get("dueDate") or add_days(invoice.get("invoiceDate", ""), 30),
        "Incasso avvenuto": "✅" if paid else "❌",
        "BANCA - C/C": transfer.get("bank", ""),
        "Termini pagamento fattura": invoice.get("paymentTerms", ""),
        "Note": "",
    }

    return {
        "id": str(uuid.uuid4()),
        "createdAt": datetime.now().isoformat(timespec="seconds"),
        "row": row,
        "invoice": invoice,
        "transfer": transfer,
        "checks": checks,
        "source": "upload",
    }


def match_score(invoice: dict[str, str], transfer: dict[str, str]) -> int:
    score = 0
    if invoice.get("invoiceNumber") and invoice["invoiceNumber"] in transfer.get("reason", ""):
        score += 100
    if invoice.get("iban") and invoice["iban"] == transfer.get("beneficiaryIban"):
        score += 40
    if decimal_from_it(invoice.get("total")) == decimal_from_it(transfer.get("total")):
        score += 40
    supplier = normalize_key(invoice.get("supplier", ""))
    beneficiary = normalize_key(transfer.get("beneficiary", ""))
    if supplier and beneficiary and (supplier in beneficiary or beneficiary in supplier):
        score += 20
    return score


def pair_documents(invoices: list[dict[str, str]], transfers: list[dict[str, str]]) -> list[tuple[dict[str, str], dict[str, str] | None]]:
    pairs = []
    used: set[int] = set()
    for invoice in invoices:
        best_index = None
        best_score = -1
        for idx, transfer in enumerate(transfers):
            if idx in used:
                continue
            score = match_score(invoice, transfer)
            if score > best_score:
                best_index = idx
                best_score = score
        if best_index is not None and best_score >= 40:
            used.add(best_index)
            pairs.append((invoice, transfers[best_index]))
        else:
            pairs.append((invoice, None))
    return pairs


def parse_multipart(body: bytes, content_type: str) -> list[dict[str, Any]]:
    match = re.search(r"boundary=(.+)", content_type)
    if not match:
        return []
    boundary = ("--" + match.group(1).strip().strip('"')).encode()
    files = []
    for part in body.split(boundary):
        if b"Content-Disposition" not in part:
            continue
        header, _, data = part.partition(b"\r\n\r\n")
        disposition = header.decode("latin1", errors="ignore")
        name_match = re.search(r'name="([^"]+)"', disposition)
        filename_match = re.search(r'filename="([^"]*)"', disposition)
        if not name_match or not filename_match or not filename_match.group(1):
            continue
        files.append(
            {
                "field": name_match.group(1),
                "filename": filename_match.group(1),
                "data": data.rstrip(b"\r\n-"),
            }
        )
    return files


def excel_serial_to_date(value: str) -> str:
    try:
        number = float(value)
    except ValueError:
        return value
    return (datetime(1899, 12, 30) + timedelta(days=number)).strftime("%d/%m/%Y")


def parse_xlsx_rows(file_data: bytes) -> list[dict[str, str]]:
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main", "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships"}
    with ZipFile(io.BytesIO(file_data)) as zf:
        shared: list[str] = []
        if "xl/sharedStrings.xml" in zf.namelist():
            root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
            for si in root.findall("m:si", ns):
                shared.append("".join(t.text or "" for t in si.findall(".//m:t", ns)))
        workbook = ET.fromstring(zf.read("xl/workbook.xml"))
        rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
        relmap = {rel.attrib["Id"]: rel.attrib["Target"] for rel in rels}
        first_sheet = workbook.find(".//m:sheet", ns)
        if first_sheet is None:
            return []
        rid = first_sheet.attrib["{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"]
        target = relmap[rid]
        path = "xl/" + target.lstrip("/") if not target.startswith("xl/") else target
        sheet = ET.fromstring(zf.read(path))

    def col_index(ref: str) -> int:
        match = re.match(r"([A-Z]+)", ref)
        if not match:
            return 0
        total = 0
        for char in match.group(1):
            total = total * 26 + ord(char) - 64
        return total

    sheet_rows: dict[int, dict[int, str]] = {}
    for cell in sheet.findall(".//m:c", ns):
        ref = cell.attrib.get("r", "")
        row_match = re.search(r"\d+", ref)
        if not row_match:
            continue
        row_number = int(row_match.group(0))
        value_node = cell.find("m:v", ns)
        if value_node is None or value_node.text is None:
            continue
        value = shared[int(value_node.text)] if cell.attrib.get("t") == "s" else value_node.text
        sheet_rows.setdefault(row_number, {})[col_index(ref)] = value

    imported = []
    for row_number in sorted(sheet_rows):
        if row_number <= 2:
            continue
        row_values = sheet_rows[row_number]
        if not any(row_values.get(col) for col in range(2, 7)):
            continue
        row = {}
        for idx, column in enumerate(EXCEL_COLUMNS, start=1):
            value = row_values.get(idx, "")
            if column in {"Data", "Data fattura", "Scadenza"} and re.fullmatch(r"\d+(?:\.\d+)?", value):
                value = excel_serial_to_date(value)
            row[column] = value
        extra_note = row_values.get(18, "")
        if extra_note:
            row["Note"] = " - ".join(note for note in [row.get("Note", ""), extra_note] if note)
        imported.append(row)
    return imported


def record_from_imported_row(row: dict[str, str], index: int) -> dict[str, Any]:
    full_row = normalize_row(row)
    if not full_row["Num."]:
        full_row["Num."] = str(index)
    checks = checks_from_row(full_row)
    return {
        "id": str(uuid.uuid4()),
        "createdAt": datetime.now().isoformat(timespec="seconds"),
        "row": full_row,
        "invoice": {},
        "transfer": {},
        "checks": checks,
        "source": "excel",
    }


def cell_ref(row: int, col: int) -> str:
    letters = ""
    while col:
        col, rem = divmod(col - 1, 26)
        letters = chr(65 + rem) + letters
    return f"{letters}{row}"


def xlsx_escape(value: Any) -> str:
    return html.escape(str(value or ""), quote=True)


def build_xlsx(records: list[dict[str, Any]]) -> bytes:
    if not EXCEL_TEMPLATE.exists():
        raise FileNotFoundError(f"Template Excel non trovato: {EXCEL_TEMPLATE}")
    output = io.BytesIO()
    with ZipFile(EXCEL_TEMPLATE) as template_zip:
        styles_xml = template_zip.read("xl/styles.xml")
        theme_xml = template_zip.read("xl/theme/theme1.xml")
        core_xml = template_zip.read("docProps/core.xml")
        app_xml = template_zip.read("docProps/app.xml")
    with ZipFile(output, "w", ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", build_content_types_xml())
        zf.writestr("_rels/.rels", build_root_rels_xml())
        zf.writestr("docProps/core.xml", core_xml)
        zf.writestr("docProps/app.xml", app_xml)
        zf.writestr("xl/workbook.xml", build_workbook_xml(records))
        zf.writestr("xl/_rels/workbook.xml.rels", build_workbook_rels_xml())
        zf.writestr("xl/worksheets/sheet1.xml", build_template_sheet(records))
        zf.writestr("xl/styles.xml", styles_xml)
        zf.writestr("xl/theme/theme1.xml", theme_xml)
    return output.getvalue()


def excel_serial_from_date(value: str | None) -> int | None:
    text = normalize_date(value)
    if not text:
        return None
    try:
        date = datetime.strptime(text, "%d/%m/%Y")
    except ValueError:
        return None
    base = datetime(1899, 12, 30)
    return (date - base).days


def decimal_as_excel(value: str | None) -> str:
    amount = decimal_from_it(value)
    if amount is None:
        return ""
    normalized = amount.normalize()
    text = format(normalized, "f")
    return text.rstrip("0").rstrip(".") if "." in text else text


def export_row_from_record(record: dict[str, Any]) -> list[str]:
    row = record.get("row", {})
    paid = "X" if row.get("Incasso avvenuto") == "✅" else ""
    return [
        row.get("Num.", ""),
        row.get("Cliente", ""),
        row.get("Data", ""),
        row.get("Fattura", ""),
        row.get("Valore in USD", ""),
        row.get("Valore in Euro", ""),
        row.get("Valore", ""),
        row.get("IVA", ""),
        row.get("Totale", ""),
        row.get("Entrate", ""),
        row.get("Delta incasso", ""),
        row.get("Data fattura", ""),
        row.get("Scadenza", ""),
        paid,
        row.get("BANCA - C/C", ""),
        row.get("Termini pagamento fattura", ""),
        row.get("Note", ""),
        "",
        "",
    ]


def make_text_cell(ref: str, style: int, value: str) -> ET.Element:
    cell = ET.Element(f"{{{SHEET_MAIN_NS}}}c", {"r": ref, "s": str(style), "t": "inlineStr"})
    inline = ET.SubElement(cell, f"{{{SHEET_MAIN_NS}}}is")
    text = ET.SubElement(inline, f"{{{SHEET_MAIN_NS}}}t")
    if value[:1].isspace() or value[-1:].isspace():
        text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    text.text = value
    return cell


def make_number_cell(ref: str, style: int, value: str) -> ET.Element:
    cell = ET.Element(f"{{{SHEET_MAIN_NS}}}c", {"r": ref, "s": str(style)})
    ET.SubElement(cell, f"{{{SHEET_MAIN_NS}}}v").text = value
    return cell


def make_blank_cell(ref: str, style: int) -> ET.Element:
    return ET.Element(f"{{{SHEET_MAIN_NS}}}c", {"r": ref, "s": str(style)})


def build_totals_row(end_row: int, totals: dict[str, Decimal]) -> ET.Element:
    row = ET.Element(
        f"{{{SHEET_MAIN_NS}}}row",
        {"r": "2", "spans": "1:19", "ht": "22.5", "customHeight": "1", "thickBot": "1"},
    )
    static_styles = [119, 109, 117, 117, 121, 117]
    for idx, style in enumerate(static_styles, start=1):
        row.append(make_blank_cell(cell_ref(2, idx), style))

    formula_styles = {7: 3, 8: 3, 9: 3, 10: 4, 11: 4}
    for col_idx, key in zip(range(7, 12), ["Valore", "IVA", "Totale", "Entrate", "Delta incasso"], strict=False):
        cell = ET.Element(f"{{{SHEET_MAIN_NS}}}c", {"r": cell_ref(2, col_idx), "s": str(formula_styles[col_idx])})
        ET.SubElement(cell, f"{{{SHEET_MAIN_NS}}}f").text = f"SUBTOTAL(9,{cell_ref(3, col_idx)}:{cell_ref(end_row, col_idx)})"
        ET.SubElement(cell, f"{{{SHEET_MAIN_NS}}}v").text = decimal_as_excel(decimal_to_it(totals[key])) or "0"
        row.append(cell)

    for col_idx, style in zip(range(12, 20), [107, 109, 111, 98, 113, 115, 104, 105], strict=False):
        row.append(make_blank_cell(cell_ref(2, col_idx), style))
    return row


def build_data_row(row_number: int, values: list[str]) -> ET.Element:
    row = ET.Element(
        f"{{{SHEET_MAIN_NS}}}row",
        {"r": str(row_number), "spans": "1:19", "ht": "20.100000000000001", "customHeight": "1"},
    )
    styles = [18, 19, 20, 21, 22, 23, 10, 24, 25, 26, 23, 27, 27, 28, 100, 17, 63, 35, 80]
    numeric_columns = {1, 5, 6, 7, 8, 9, 10, 11, 19}
    date_columns = {3, 12, 13}
    for col_idx, (style, value) in enumerate(zip(styles, values, strict=False), start=1):
        ref = cell_ref(row_number, col_idx)
        if col_idx in date_columns:
            serial = excel_serial_from_date(value)
            row.append(make_number_cell(ref, style, str(serial)) if serial is not None else make_blank_cell(ref, style) if not value else make_text_cell(ref, style, value))
            continue
        if col_idx in numeric_columns:
            number = decimal_as_excel(value)
            row.append(make_number_cell(ref, style, number) if number else make_blank_cell(ref, style))
            continue
        row.append(make_blank_cell(ref, style) if not value else make_text_cell(ref, style, value))
    return row


def build_template_sheet(records: list[dict[str, Any]]) -> bytes:
    with ZipFile(EXCEL_TEMPLATE) as template_zip:
        template_sheet = ET.fromstring(template_zip.read("xl/worksheets/sheet1.xml"))
    worksheet = ET.Element(f"{{{SHEET_MAIN_NS}}}worksheet")
    dimension = ET.Element(f"{{{SHEET_MAIN_NS}}}dimension")
    sheet_views = copy.deepcopy(template_sheet.find("a:sheetViews", NS))
    sheet_format = copy.deepcopy(template_sheet.find("a:sheetFormatPr", NS))
    cols = copy.deepcopy(template_sheet.find("a:cols", NS))
    sheet_data = ET.Element(f"{{{SHEET_MAIN_NS}}}sheetData")
    auto_filter = ET.Element(f"{{{SHEET_MAIN_NS}}}autoFilter")
    merge_cells = copy.deepcopy(template_sheet.find("a:mergeCells", NS))
    page_margins = copy.deepcopy(template_sheet.find("a:pageMargins", NS))

    for node in [dimension, sheet_views, sheet_format, cols, sheet_data, auto_filter, merge_cells, page_margins]:
        if node is not None:
            worksheet.append(node)

    sheet_data = worksheet.find("a:sheetData", NS)
    if sheet_data is None:
        raise ValueError("Errore durante la costruzione del foglio Excel.")

    header_row = build_header_row()
    totals: dict[str, Decimal] = {key: Decimal("0") for key in ["Valore", "IVA", "Totale", "Entrate", "Delta incasso"]}
    export_rows = [export_row_from_record(record) for record in records]
    for row in export_rows:
        for key, idx in [("Valore", 6), ("IVA", 7), ("Totale", 8), ("Entrate", 9), ("Delta incasso", 10)]:
            amount = decimal_from_it(row[idx])
            if amount is not None:
                totals[key] += amount

    end_row = max(3, len(export_rows) + 2)
    sheet_data.append(header_row)
    sheet_data.append(build_totals_row(end_row, totals))
    for row_number, values in enumerate(export_rows, start=3):
        sheet_data.append(build_data_row(row_number, values))

    dimension = worksheet.find("a:dimension", NS)
    if dimension is not None:
        dimension.set("ref", f"A1:S{max(2, len(export_rows) + 2)}")

    auto_filter = worksheet.find("a:autoFilter", NS)
    if auto_filter is not None:
        auto_filter.attrib.clear()
        auto_filter.set("ref", f"A2:S{max(2, len(export_rows) + 2)}")

    return ET.tostring(worksheet, encoding="utf-8", xml_declaration=True)

def build_header_row() -> ET.Element:
    row = ET.Element(f"{{{SHEET_MAIN_NS}}}row", {"r": "1", "spans": "1:19", "ht": "58.8", "customHeight": "1"})
    header_styles = [118, 108, 116, 116, 120, 116, 1, 1, 2, 2, 2, 106, 108, 110, 97, 112, 114, 104, 105]
    for idx, (style, label) in enumerate(zip(header_styles, EXPORT_TEMPLATE_COLUMNS, strict=False), start=1):
        row.append(make_text_cell(cell_ref(1, idx), style, label))
    return row


def build_content_types_xml() -> bytes:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        '<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
        '</Types>'
    ).encode("utf-8")


def build_root_rels_xml() -> bytes:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'
        '</Relationships>'
    ).encode("utf-8")


def build_workbook_xml(records: list[dict[str, Any]]) -> bytes:
    last_row = max(2, len(records) + 2)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<bookViews><workbookView xWindow="-108" yWindow="-108" windowWidth="23256" windowHeight="12456"/></bookViews>'
        '<sheets><sheet name="MARTEC" sheetId="1" r:id="rId1"/></sheets>'
        f'<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">MARTEC!$A$2:$S${last_row}</definedName></definedNames>'
        '<calcPr calcId="181029"/>'
        '</workbook>'
    ).encode("utf-8")


def build_workbook_rels_xml() -> bytes:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
        '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'
        '</Relationships>'
    ).encode("utf-8")


class AppHandler(BaseHTTPRequestHandler):
    server_version = "MartecRiepilogo/0.2"

    def log_message(self, format: str, *args: Any) -> None:
        print(f"{self.address_string()} - {format % args}")

    def send_json(self, payload: Any, status: HTTPStatus = HTTPStatus.OK) -> None:
        encoded = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(encoded)))
        self.end_headers()
        self.wfile.write(encoded)

    def read_json(self) -> dict[str, Any]:
        length = int(self.headers.get("Content-Length", "0"))
        if not length:
            return {}
        return json.loads(self.rfile.read(length).decode("utf-8"))

    def do_HEAD(self) -> None:
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.end_headers()

    def do_GET(self) -> None:
        if self.path == "/api/records":
            self.send_json({"columns": EXCEL_COLUMNS, "records": load_records(), "suppliers": load_suppliers()})
            return
        if self.path == "/api/suppliers":
            self.send_json({"suppliers": load_suppliers()})
            return
        if self.path == "/brand-logo":
            self.serve_brand_logo()
            return
        if self.path == "/api/export.csv":
            self.export_csv()
            return
        if self.path == "/api/export.xlsx":
            self.export_xlsx()
            return
        path = STATIC_DIR / ("index.html" if self.path in ("/", "") else self.path.lstrip("/"))
        if not path.resolve().is_relative_to(STATIC_DIR.resolve()) or not path.exists():
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        content = path.read_bytes()
        content_type = "text/html; charset=utf-8"
        if path.suffix == ".css":
            content_type = "text/css; charset=utf-8"
        elif path.suffix == ".js":
            content_type = "application/javascript; charset=utf-8"
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(content)))
        self.end_headers()
        self.wfile.write(content)

    def do_POST(self) -> None:
        if self.path == "/api/upload":
            self.handle_upload()
            return
        if self.path == "/api/import-xlsx":
            self.handle_import_xlsx()
            return
        self.send_error(HTTPStatus.NOT_FOUND)

    def do_PATCH(self) -> None:
        match = re.fullmatch(r"/api/records/([^/]+)", self.path)
        if not match:
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        payload = self.read_json()
        updated = update_record_row(match.group(1), payload.get("row", {}))
        if not updated:
            self.send_json({"error": "Riga non trovata"}, HTTPStatus.NOT_FOUND)
            return
        self.send_json({"record": updated, "records": load_records(), "suppliers": load_suppliers()})

    def do_DELETE(self) -> None:
        if self.path == "/api/records":
            clear_records()
            self.send_json({"ok": True, "records": []})
            return
        self.send_error(HTTPStatus.NOT_FOUND)

    def handle_upload(self) -> None:
        files = parse_multipart(self.rfile.read(int(self.headers.get("Content-Length", "0"))), self.headers.get("Content-Type", ""))
        pdf_files = [file for file in files if file["filename"].lower().endswith(".pdf")]
        if not pdf_files:
            self.send_json({"error": "Carica almeno un PDF."}, HTTPStatus.BAD_REQUEST)
            return

        invoices: list[dict[str, str]] = []
        transfers: list[dict[str, str]] = []
        batch_id = str(uuid.uuid4())
        for file in pdf_files:
            safe_name = Path(file["filename"]).name
            (UPLOAD_DIR / f"{batch_id}-{safe_name}").write_bytes(file["data"])
            parsed = classify_pdf(file["data"])
            parsed["filename"] = safe_name
            if parsed["type"] == "transfer":
                transfers.append(parsed)
            else:
                invoices.append(parsed)

        if not invoices:
            self.send_json({"error": "Nessuna fattura riconosciuta nei file caricati."}, HTTPStatus.BAD_REQUEST)
            return

        added = []
        duplicates = []
        with connect() as db:
            index = next_number(db)
            for invoice, transfer in pair_documents(invoices, transfers):
                record = build_record(invoice, transfer, index)
                ok, existing_id = save_record(db, record, allow_duplicate=False)
                if ok:
                    added.append(record)
                    index += 1
                else:
                    duplicates.append({"invoice": invoice.get("invoiceNumber"), "supplier": invoice.get("supplier"), "existingId": existing_id})
        self.send_json({"columns": EXCEL_COLUMNS, "added": added, "duplicates": duplicates, "records": load_records(), "suppliers": load_suppliers()})

    def handle_import_xlsx(self) -> None:
        files = parse_multipart(self.rfile.read(int(self.headers.get("Content-Length", "0"))), self.headers.get("Content-Type", ""))
        excel_files = [file for file in files if file["filename"].lower().endswith(".xlsx")]
        if not excel_files:
            self.send_json({"error": "Carica un file .xlsx."}, HTTPStatus.BAD_REQUEST)
            return
        imported_rows = parse_xlsx_rows(excel_files[0]["data"])
        added = 0
        duplicates = 0
        with connect() as db:
            index = next_number(db)
            for row in imported_rows:
                record = record_from_imported_row(row, index)
                ok, _ = save_record(db, record, allow_duplicate=False)
                if ok:
                    added += 1
                    index += 1
                else:
                    duplicates += 1
        self.send_json({"added": added, "duplicates": duplicates, "records": load_records(), "suppliers": load_suppliers()})

    def export_csv(self) -> None:
        buffer = io.StringIO()
        writer = csv.DictWriter(buffer, fieldnames=EXCEL_COLUMNS, delimiter=";")
        writer.writeheader()
        for record in load_records():
            writer.writerow(record.get("row", {}))
        encoded = buffer.getvalue().encode("utf-8-sig")
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "text/csv; charset=utf-8")
        self.send_header("Content-Disposition", 'attachment; filename="riepilogo-fatture-martec.csv"')
        self.send_header("Content-Length", str(len(encoded)))
        self.end_headers()
        self.wfile.write(encoded)

    def export_xlsx(self) -> None:
        encoded = build_xlsx(load_records())
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        self.send_header("Content-Disposition", 'attachment; filename="riepilogo-fatture-martec.xlsx"')
        self.send_header("Content-Length", str(len(encoded)))
        self.end_headers()
        self.wfile.write(encoded)

    def serve_brand_logo(self) -> None:
        if not BRAND_LOGO.exists():
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        content = BRAND_LOGO.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "image/png")
        self.send_header("Content-Length", str(len(content)))
        self.end_headers()
        self.wfile.write(content)


def main() -> None:
    init_db()
    server = ThreadingHTTPServer(("127.0.0.1", 8000), AppHandler)
    print("Web app avviata su http://127.0.0.1:8000")
    server.serve_forever()


if __name__ == "__main__":
    main()
