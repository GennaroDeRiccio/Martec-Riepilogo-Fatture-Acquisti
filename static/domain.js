export const EXCEL_COLUMNS = [
  "Num.",
  "Fornitore",
  "Fattura",
  "Valore in USD",
  "Valore in Euro",
  "Imponibile",
  "IVA",
  "Totale",
  "Uscite",
  "Da pagare ancora",
  "Data fattura",
  "Scadenza",
  "BANCA - C/C",
  "Termini pagamento fattura",
  "Note",
];

export const EXPORT_TEMPLATE_COLUMNS = [
  "Num.",
  "Fornitore",
  "Fattura",
  "Valore in USD",
  "Valore in Euro",
  "Imponibile",
  "IVA",
  "Totale",
  "Uscite",
  "Da pagare ancora",
  "Data fattura",
  "Scadenza",
  "Stato",
  "BANCA - C/C",
  "Termini pagamento fattura",
  "Note",
];

export const STATUS_OK = "Pagato";
export const STATUS_VERIFY = "Da pagare";
export const MATCH_THRESHOLD = 40;

const OLD_DATE_COLUMN = "Fattura anno 2019 - 2020-2021-2022 - 2023";
const OLD_NOTE_COLUMN = "NOTE VARIE";
const OLD_EXTRA_NOTE_COLUMN = "Ulteriori Note";
const OLD_VALUE_COLUMN = "Valore";
const OLD_OUTGOING_COLUMN = "Entrate";
const OLD_OUTSTANDING_COLUMN = "Delta incasso";
const OLD_PAID_COLUMN = "Incasso avvenuto";
const EURO_COLUMNS = new Set(["Valore in Euro", "Imponibile", "IVA", "Totale", "Uscite"]);

export function cleanText(text = "") {
  return String(text).replace(/\s+/g, " ").trim();
}

export function decimalFromIt(value) {
  if (value === null || value === undefined || value === "") return null;
  const cleaned = String(value)
    .trim()
    .replace(/EUR|USD|€|\$/g, "")
    .replace(/\s+/g, "")
    .replace(/\./g, "")
    .replace(",", ".");
  const parsed = Number.parseFloat(cleaned);
  return Number.isFinite(parsed) ? parsed : null;
}

export function decimalToIt(value) {
  if (value === null || value === undefined || Number.isNaN(value)) return "";
  return Number(value).toLocaleString("it-IT", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

export function formatCurrency(value, symbol) {
  if (value === null || value === undefined || value === "") return "";
  const amount = decimalFromIt(value);
  return amount === null ? String(value) : `${decimalToIt(amount)} ${symbol}`;
}

export function normalizeDate(value) {
  if (!value) return "";
  const text = String(value).trim();
  const formats = [
    [/^(\d{2})-(\d{2})-(\d{4})$/, ([d, m, y]) => `${d}/${m}/${y}`],
    [/^(\d{2})\.(\d{2})\.(\d{4})$/, ([d, m, y]) => `${d}/${m}/${y}`],
    [/^(\d{2})\/(\d{2})\/(\d{4})$/, ([d, m, y]) => `${d}/${m}/${y}`],
    [/^(\d{4})-(\d{2})-(\d{2})$/, ([y, m, d]) => `${d}/${m}/${y}`],
  ];
  for (const [pattern, formatter] of formats) {
    const match = text.match(pattern);
    if (match) return formatter(match.slice(1));
  }
  return text;
}

export function addDays(dateText, days) {
  const normalized = normalizeDate(dateText);
  const match = normalized.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!match) return "";
  const date = new Date(Number(match[3]), Number(match[2]) - 1, Number(match[1]));
  if (Number.isNaN(date.getTime())) return "";
  date.setDate(date.getDate() + days);
  return `${String(date.getDate()).padStart(2, "0")}/${String(date.getMonth() + 1).padStart(2, "0")}/${date.getFullYear()}`;
}

export function normalizeKey(value) {
  return String(value || "").toUpperCase().replace(/[^A-Z0-9]+/g, "");
}

function normalizedWords(value) {
  const stopwords = new Set(["SRL", "SPA", "S", "RL", "ITALIA", "UNIPERSONALE", "SPA", "SOCIETA", "SOCIET", "LTD"]);
  return String(value || "")
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, " ")
    .split(/\s+/)
    .filter(Boolean)
    .filter((word) => word.length > 2 && !stopwords.has(word));
}

function overlapScore(a, b) {
  const left = normalizedWords(a);
  const right = normalizedWords(b);
  if (!left.length || !right.length) return { score: 0, overlap: [] };
  const overlap = left.filter((word) => right.includes(word));
  if (overlap.length >= 2) return { score: 30, overlap };
  if (overlap.length === 1) return { score: 15, overlap };
  return { score: 0, overlap: [] };
}

function normalizePaymentMethod(value = "") {
  const normalized = String(value || "").trim().toUpperCase();
  if (normalized.includes("PAYPAL")) return "PAYPAL";
  if (normalized.includes("CARTA DI CREDITO")) return "CARTA DI CREDITO";
  if (normalized.includes("RIBA")) return "RIBA";
  if (normalized.includes("BONIFICO")) return "BONIFICO";
  return normalized;
}

export function normalizePaidValue(value) {
  const text = String(value || "").trim().toUpperCase();
  return ["X", "XX", "✅", "SI", "SÌ", "PAGATO", "PAGATA"].includes(text) ? "✅" : "❌";
}

export function outstandingLabel(total, paid) {
  if (total === null) return "";
  const remaining = Math.max(0, total - (paid || 0));
  return remaining < 0.00001 ? "Pagato" : decimalToIt(remaining);
}

export function amountPaidLabel(total, outstanding) {
  if (total === null) return "";
  const remaining = String(outstanding || "").trim().toUpperCase() === "PAGATO" ? 0 : (decimalFromIt(outstanding) || 0);
  return decimalToIt(Math.max(0, total - remaining));
}

export function normalizeStatus(value) {
  return value === STATUS_OK ? STATUS_OK : STATUS_VERIFY;
}

export function normalizeRow(row = {}) {
  const stringRow = Object.fromEntries(Object.entries(row).map(([key, value]) => [String(key), value ?? ""]));
  if (!stringRow.Fornitore) stringRow.Fornitore = stringRow.Cliente || "";
  if (!stringRow["Data fattura"]) stringRow["Data fattura"] = stringRow[OLD_DATE_COLUMN] || "";
  if (!stringRow["Data fattura"]) stringRow["Data fattura"] = stringRow.Data || "";
  if (!stringRow.Imponibile) stringRow.Imponibile = stringRow[OLD_VALUE_COLUMN] || "";
  if (!stringRow.Uscite) stringRow.Uscite = stringRow[OLD_OUTGOING_COLUMN] || "";
  if (!stringRow["Da pagare ancora"]) stringRow["Da pagare ancora"] = stringRow[OLD_OUTSTANDING_COLUMN] || "";
  if (!stringRow["Da pagare ancora"] && stringRow[OLD_PAID_COLUMN]) {
    stringRow["Da pagare ancora"] = normalizePaidValue(stringRow[OLD_PAID_COLUMN]) === "✅" ? "Pagato" : "";
  }
  if (!stringRow.Note) {
    const notes = [stringRow[OLD_NOTE_COLUMN], stringRow[OLD_EXTRA_NOTE_COLUMN]].filter(Boolean);
    stringRow.Note = notes.join(" - ");
  }
  const normalized = Object.fromEntries(EXCEL_COLUMNS.map((column) => [column, String(stringRow[column] || "")]));
  for (const column of EURO_COLUMNS) normalized[column] = formatCurrency(normalized[column], "€");
  normalized["Valore in USD"] = formatCurrency(normalized["Valore in USD"], "$");
  normalized["Data fattura"] = normalizeDate(normalized["Data fattura"]);
  normalized.Scadenza = normalizeDate(normalized.Scadenza);
  return normalized;
}

export function checksFromRow(row) {
  const total = decimalFromIt(row?.Totale);
  const paid = decimalFromIt(row?.Uscite);
  const outstanding = String(row?.["Da pagare ancora"] || "").trim().toUpperCase() === "PAGATO"
    ? 0
    : decimalFromIt(row?.["Da pagare ancora"]);
  return [
    { label: "Importo", ok: total !== null && outstanding !== null && Math.abs(Math.max(0, total - (paid || 0)) - outstanding) < 0.00001 },
    { label: "Pagamento", ok: String(row?.["Da pagare ancora"] || "").trim().toUpperCase() === "PAGATO" },
  ];
}

export function invoiceKeyFromRow(row) {
  const normalized = normalizeRow(row);
  const supplier = normalizeKey(normalized.Fornitore);
  const number = normalizeKey(normalized.Fattura);
  const date = normalizeKey(normalized["Data fattura"]);
  const total = normalizeKey(normalized.Totale);
  return supplier && number ? `${supplier}|${number}|${date}|${total}` : "";
}

export function normalizeTransfers(transferData = null) {
  if (Array.isArray(transferData?.transfers)) return transferData.transfers;
  if (Array.isArray(transferData)) return transferData;
  if (transferData && Object.keys(transferData).length) return [transferData];
  return [];
}

function uniqueNonEmpty(values) {
  return [...new Set(values.filter(Boolean))];
}

function transferDateLabel(transfer = {}) {
  return normalizeDate(transfer.dueDate || transfer.executionDate || transfer.documentDate || "");
}

function bankAccountLabel(invoice, transfers) {
  const lines = uniqueNonEmpty(transfers.map((transfer) => {
    const bank = cleanText(transfer.bank || "");
    const iban = cleanText(transfer.payerIban || "");
    if (bank && iban) return `${bank}\n${iban}`;
    return bank || iban;
  }));
  return lines.join("\n");
}

function paymentTermsLabel(invoice, transfers) {
  if (isInstantPaidMethod(invoice.paymentTerms)) {
    return `${invoice.paymentTerms} in data ${invoice.invoiceDate || ""}`.trim();
  }
  const method = normalizePaymentMethod(invoice.paymentTerms || transfers[0]?.paymentType);
  const datedTransfer = transfers.find((transfer) => transferDateLabel(transfer));
  if (datedTransfer && method === "RIBA") return `RIBA in data ${transferDateLabel(datedTransfer)}`;
  if (datedTransfer) return `Bonifico in data ${transferDateLabel(datedTransfer)}`;
  if (invoice.paymentTerms) return invoice.paymentTerms;
  return "Bonifico";
}

function isInstantPaidMethod(paymentTerms = "") {
  const normalized = String(paymentTerms || "").trim().toUpperCase();
  return normalized === "PAYPAL" || normalized === "CARTA DI CREDITO";
}

export function recalculateRecord(invoice, transfers, index, matchDebug = null, source = "upload") {
  const safeTransfers = normalizeTransfers(transfers);
  const invoiceTotal = invoice.totalEur ?? decimalFromIt(invoice.total);
  const taxableEuro = invoice.taxableEur ?? decimalFromIt(invoice.taxable);
  const vatEuro = invoice.vatEur ?? decimalFromIt(invoice.vat);
  const taxableUsd = invoice.currency === "USD"
    ? (invoice.taxableUsd ?? decimalFromIt(invoice.taxable) ?? decimalFromIt(invoice.total))
    : null;
  const instantPaid = isInstantPaidMethod(invoice.paymentTerms);
  const transfersPaidAmount = safeTransfers.reduce((sum, transfer) => sum + (transfer.totalEur ?? decimalFromIt(transfer.total) ?? 0), 0);
  const paidAmount = instantPaid ? (invoiceTotal ?? transfersPaidAmount) : transfersPaidAmount;
  const remainingLabel = outstandingLabel(invoiceTotal, paidAmount);
  const row = normalizeRow({
    "Num.": String(index),
    Fornitore: invoice.supplier || safeTransfers[0]?.beneficiary || "",
    Fattura: invoice.invoiceNumber || "",
    "Valore in USD": taxableUsd !== null ? decimalToIt(taxableUsd) : "",
    "Valore in Euro": decimalToIt(taxableEuro),
    Imponibile: decimalToIt(taxableEuro),
    IVA: decimalToIt(vatEuro),
    Totale: decimalToIt(invoiceTotal),
    Uscite: decimalToIt(paidAmount),
    "Da pagare ancora": remainingLabel,
    "Data fattura": invoice.invoiceDate || "",
    Scadenza: invoice.dueDate || addDays(invoice.invoiceDate || "", 30),
    "BANCA - C/C": bankAccountLabel(invoice, safeTransfers),
    "Termini pagamento fattura": paymentTermsLabel(invoice, safeTransfers),
    Note: "",
  });
  const checks = [
    { label: "Importo", ok: String(remainingLabel).toUpperCase() === "PAGATO" },
    { label: "Pagamento registrato", ok: instantPaid || safeTransfers.length > 0 },
  ];
  return {
    id: crypto.randomUUID(),
    createdAt: new Date().toISOString(),
    row,
    invoice: { ...invoice, matchDebug: matchDebug || invoice.matchDebug || null },
    transfer: safeTransfers[0] || {},
    transfers: safeTransfers,
    checks,
    source,
    status: String(remainingLabel).toUpperCase() === "PAGATO" ? STATUS_OK : STATUS_VERIFY,
    invoiceKey: invoiceKeyFromRow(row),
  };
}

export function buildRecord(invoice, transfer, index, matchDebug = null) {
  return recalculateRecord(invoice, transfer ? [transfer] : [], index, matchDebug, "upload");
}

export function recordFromImportedRow(row, index) {
  const normalized = normalizeRow(row);
  if (!normalized["Num."]) normalized["Num."] = String(index);
  return {
    id: crypto.randomUUID(),
    createdAt: new Date().toISOString(),
    row: normalized,
    invoice: {},
    transfer: {},
    transfers: [],
    checks: checksFromRow(normalized),
    source: "excel",
    status: String(normalized["Da pagare ancora"] || "").trim().toUpperCase() === "PAGATO" ? STATUS_OK : STATUS_VERIFY,
    invoiceKey: invoiceKeyFromRow(normalized),
  };
}

export function matchScore(invoice, transfer) {
  let score = 0;
  const reasons = [];
  const issues = [];
  const invoiceMethod = normalizePaymentMethod(invoice.paymentTerms);
  const transferMethod = normalizePaymentMethod(transfer.paymentType || transfer.paymentTerms || transfer.method);

  if (invoiceMethod && transferMethod && invoiceMethod === transferMethod) {
    score += 20;
    reasons.push(`Metodo pagamento coerente (${invoiceMethod})`);
  } else if (transferMethod) {
    issues.push(`Metodo pagamento diverso (${transferMethod})`);
  }

  const invoiceNumberInReason = Boolean(invoice.invoiceNumber && normalizeKey(transfer.reason || "").includes(normalizeKey(invoice.invoiceNumber)));
  if (invoiceNumberInReason) {
    score += 100;
    reasons.push(`Numero fattura trovato in causale (${invoice.invoiceNumber})`);
  } else {
    issues.push("Numero fattura non trovato nella causale");
  }

  const ibanMatches = Boolean(invoice.iban && invoice.iban === transfer.beneficiaryIban);
  if (ibanMatches) {
    score += 40;
    reasons.push("IBAN fattura uguale al conto beneficiario");
  } else if (invoice.iban || transfer.beneficiaryIban) {
    issues.push("IBAN non coincidente");
  }

  const invoiceTotal = invoice.totalEur ?? decimalFromIt(invoice.total);
  const transferTotal = transfer.totalEur ?? decimalFromIt(transfer.total);
  const amountsMatch = invoiceTotal !== null && transferTotal !== null && Math.abs(invoiceTotal - transferTotal) < 0.00001;
  if (amountsMatch) {
    score += 40;
    reasons.push("Importo bonifico uguale al totale fattura");
  } else {
    issues.push("Importo diverso");
  }

  const dueDateMatches = Boolean(invoice.dueDate && transfer.dueDate && normalizeDate(invoice.dueDate) === normalizeDate(transfer.dueDate));
  if (dueDateMatches) {
    score += 40;
    reasons.push("Scadenza effetto uguale alla scadenza fattura");
  } else if (invoice.dueDate || transfer.dueDate) {
    issues.push("Scadenza diversa");
  }

  const supplier = normalizeKey(invoice.supplier || "");
  const beneficiary = normalizeKey(transfer.beneficiary || "");
  const directNameMatch = supplier && beneficiary && (supplier.includes(beneficiary) || beneficiary.includes(supplier));
  const overlap = overlapScore(invoice.supplier || "", transfer.beneficiary || "");
  if (directNameMatch) {
    score += 20;
    reasons.push("Nome fornitore coerente con beneficiario");
  } else if (overlap.score > 0) {
    score += overlap.score;
    reasons.push(`Nome simile tra fornitore e beneficiario (${overlap.overlap.join(", ")})`);
  } else {
    issues.push("Nome fornitore e beneficiario poco simili");
  }

  return {
    score,
    reasons,
    issues,
    summary: transfer.beneficiary
      ? `${transfer.beneficiary} | ${transfer.total || "-"} ${transfer.currency || "EUR"} | ${transfer.dueDate || transfer.executionDate || transfer.documentDate || ""}`.trim()
      : "Bonifico non identificato",
  };
}

export function pairDocuments(invoices, transfers) {
  const pairs = [];
  const used = new Set();
  for (const invoice of invoices) {
    const remainingStart = invoice.totalEur ?? decimalFromIt(invoice.total) ?? 0;
    const ranked = transfers
      .map((transfer, index) => ({ transfer, index, match: matchScore(invoice, transfer) }))
      .filter((entry) => !used.has(entry.index))
      .sort((left, right) => right.match.score - left.match.score);
    let bestMatch = ranked[0]?.match || null;
    const selected = [];
    let remaining = remainingStart;
    for (const entry of ranked) {
      if (entry.match.score < MATCH_THRESHOLD) continue;
      const transferAmount = entry.transfer.totalEur ?? decimalFromIt(entry.transfer.total) ?? 0;
      if (selected.length && transferAmount > remaining + 0.01 && entry.match.score < 100) continue;
      selected.push(entry);
      used.add(entry.index);
      remaining -= transferAmount;
      if (remaining <= 0.01) break;
    }
    if (selected.length) {
      const combinedReasons = [...new Set(selected.flatMap((entry) => entry.match.reasons))];
      const combinedIssues = remaining > 0.01
        ? [`Residuo da pagare: ${decimalToIt(remaining)} €`]
        : [];
      pairs.push({
        invoice,
        transfers: selected.map((entry) => entry.transfer),
        match: {
          score: selected.reduce((sum, entry) => sum + entry.match.score, 0),
          reasons: combinedReasons,
          issues: combinedIssues,
          summary: selected.map((entry) =>
            `${entry.transfer.beneficiary || "-"} | ${entry.transfer.total || "-"} ${entry.transfer.currency || "EUR"} | ${entry.transfer.dueDate || entry.transfer.executionDate || entry.transfer.documentDate || ""}`.trim(),
          ).join(" || "),
          matched: true,
          threshold: MATCH_THRESHOLD,
        },
      });
    } else {
      pairs.push({
        invoice,
        transfers: [],
        match: {
          ...(bestMatch || { score: 0, reasons: [], issues: ["Nessun bonifico compatibile trovato"], summary: "" }),
          matched: false,
          threshold: MATCH_THRESHOLD,
        },
      });
    }
  }
  return pairs;
}

export function exportRowFromRecord(record) {
  const sourceRow = normalizeRow(record.row || {});
  const transfers = normalizeTransfers(record.transfers || record.transfer || []);
  const row = normalizeRow({
    ...sourceRow,
    "BANCA - C/C": sourceRow["BANCA - C/C"] || bankAccountLabel(record.invoice || {}, transfers),
    "Termini pagamento fattura": sourceRow["Termini pagamento fattura"] || paymentTermsLabel(record.invoice || {}, transfers),
  });
  return [
    row["Num."],
    row.Fornitore,
    row.Fattura,
    row["Valore in USD"],
    row["Valore in Euro"],
    row.Imponibile,
    row.IVA,
    row.Totale,
    row.Uscite,
    row["Da pagare ancora"],
    row["Data fattura"],
    row.Scadenza,
    record.status,
    row["BANCA - C/C"],
    row["Termini pagamento fattura"],
    row.Note,
  ];
}
