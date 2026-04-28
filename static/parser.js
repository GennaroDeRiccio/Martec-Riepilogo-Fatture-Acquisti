import JSZip from "https://cdn.jsdelivr.net/npm/jszip@3.10.1/+esm";
import * as pdfjsLib from "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/build/pdf.min.mjs";
import {
  EXCEL_COLUMNS,
  EXPORT_TEMPLATE_COLUMNS,
  addDays,
  buildRecord,
  cleanText,
  decimalFromIt,
  decimalToIt,
  exportRowFromRecord,
  normalizeDate,
  normalizeRow,
  normalizeTransfers,
  pairDocuments,
  recordFromImportedRow,
} from "./domain.js";

pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdn.jsdelivr.net/npm/pdfjs-dist@4.10.38/build/pdf.worker.min.mjs";

const NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const templateCache = { workbook: null, parts: null };
const fxCache = new Map();
const APP_CONFIG_KEY = "martec-cloud-config";
const DEFAULT_GEMINI_MODEL = "gemini-2.5-flash";

function isFetchFailure(error) {
  return error instanceof TypeError || /Failed to fetch|Load failed|NetworkError/i.test(String(error?.message || error || ""));
}

function configFromWindow() {
  return window.APP_CONFIG || {};
}

function getAiConfig() {
  const stored = JSON.parse(localStorage.getItem(APP_CONFIG_KEY) || "{}");
  const config = { ...configFromWindow(), ...stored };
  return {
    supabaseUrl: String(config.supabaseUrl || "").trim(),
    supabaseAnonKey: String(config.supabaseAnonKey || "").trim(),
    geminiModel: String(config.geminiModel || DEFAULT_GEMINI_MODEL).trim() || DEFAULT_GEMINI_MODEL,
  };
}

function sleep(ms) {
  return new Promise((resolve) => window.setTimeout(resolve, ms));
}

function shouldFallbackFromGemini(status, text = "") {
  return [429, 500, 503, 504].includes(Number(status))
    || /high demand|try again later|overloaded|unavailable|quota/i.test(String(text || ""));
}

function geminiPermissionProblem(status, text = "") {
  return Number(status) === 403
    && /api key was reported as leaked|permission_denied|api key not valid|forbidden/i.test(String(text || ""));
}

function firstMatch(pattern, text, flags = "i") {
  const match = new RegExp(pattern, flags).exec(text);
  return match ? cleanText(match[1]) : "";
}

function valueOnSameRow(items, label, { labelMaxX = null, valueMinX = null, valueMaxX = null } = {}) {
  const labelLower = label.toLowerCase();
  for (const item of items) {
    if (!String(item.text || "").toLowerCase().includes(labelLower)) continue;
    if (labelMaxX !== null && item.x > labelMaxX) continue;
    const sameRow = items
      .filter((other) => Math.abs(other.y - item.y) <= 1.5 && other.x > item.x + 8)
      .filter((other) => valueMinX === null || other.x >= valueMinX)
      .filter((other) => valueMaxX === null || other.x <= valueMaxX);
    if (sameRow.length) return cleanText(sameRow.sort((a, b) => a.x - b.x).map((other) => other.text).join(" "));
  }
  return "";
}

function moneyBelowLabel(items, label, { xMin, xMax }) {
  const labelItem = items.find((item) => String(item.text || "").toLowerCase().includes(label.toLowerCase()));
  if (!labelItem) return "";
  const candidates = items
    .filter((item) => item.y < labelItem.y && item.x >= xMin && item.x <= xMax)
    .filter((item) => /^-?\d+(?:\.\d{3})*,\d{2}$/.test(String(item.text || "").trim()))
    .sort((a, b) => (labelItem.y - a.y) - (labelItem.y - b.y) || Math.abs((xMin + xMax) / 2 - a.x) - Math.abs((xMin + xMax) / 2 - b.x));
  return candidates[0]?.text?.trim() || "";
}

function moneyCandidates(text) {
  return [...String(text || "").matchAll(/\b\d{1,3}(?:\.\d{3})*,\d{2}\b/g)].map((match) => match[0]);
}

function detectDocumentCurrency(text, explicitValue = "") {
  const explicit = String(explicitValue || "").trim().toUpperCase();
  if (["USD", "EUR"].includes(explicit)) return explicit;
  if (/\bDIVISA\s+USD\b/i.test(text) || /\bUSD\b/.test(text) || /\$/u.test(text)) return "USD";
  return "EUR";
}

function shiftIsoDate(isoDate, days) {
  const match = String(isoDate || "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return "";
  const date = new Date(Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3])));
  date.setUTCDate(date.getUTCDate() + days);
  return date.toISOString().slice(0, 10);
}

async function fetchUsdEurRate(isoDate) {
  if (!isoDate) return null;
  if (fxCache.has(isoDate)) return fxCache.get(isoDate);
  const start = shiftIsoDate(isoDate, -7);
  const url = `https://data-api.ecb.europa.eu/service/data/EXR/D.USD.EUR.SP00.A?startPeriod=${start}&endPeriod=${isoDate}&format=csvdata`;
  let response;
  try {
    response = await fetch(url);
  } catch (error) {
    if (isFetchFailure(error)) {
      throw new Error("Connessione non riuscita mentre leggevo il cambio USD/EUR dalla BCE.");
    }
    throw error;
  }
  if (!response.ok) throw new Error("Impossibile leggere il cambio USD/EUR dalla BCE.");
  const csv = await response.text();
  const rows = csv.trim().split(/\r?\n/);
  if (rows.length < 2) {
    fxCache.set(isoDate, null);
    return null;
  }
  const headers = rows[0].split(",");
  const periodIndex = headers.indexOf("TIME_PERIOD");
  const valueIndex = headers.indexOf("OBS_VALUE");
  const parsed = rows.slice(1)
    .map((line) => line.split(","))
    .map((cols) => ({ date: cols[periodIndex], value: Number.parseFloat(cols[valueIndex]) }))
    .filter((entry) => entry.date && Number.isFinite(entry.value))
    .sort((a, b) => a.date.localeCompare(b.date));
  const latest = parsed[parsed.length - 1]?.value ?? null;
  fxCache.set(isoDate, latest);
  return latest;
}

async function normalizeCurrencyAmounts({ currency, date, taxable = "", vat = "", total = "" }) {
  const normalizedCurrency = detectDocumentCurrency("", currency);
  const taxableValue = decimalFromIt(taxable);
  const vatValue = decimalFromIt(vat);
  const totalValue = decimalFromIt(total);
  if (normalizedCurrency !== "USD") {
    return {
      currency: "EUR",
      exchangeRate: null,
      taxableUsd: null,
      taxableEur: taxableValue,
      vatEur: vatValue,
      totalEur: totalValue,
      totalUsd: null,
    };
  }
  const rate = await fetchUsdEurRate((date || "").split("/").reverse().join("-"));
  if (!rate) throw new Error(`Cambio USD/EUR non trovato per la data ${date || "documento"}.`);
  const convert = (value) => (value === null ? null : value / rate);
  return {
    currency: "USD",
    exchangeRate: rate,
    taxableUsd: taxableValue,
    taxableEur: convert(taxableValue),
    vatEur: convert(vatValue),
    totalEur: convert(totalValue),
    totalUsd: totalValue,
  };
}

function numberToIt(value) {
  if (value === null || value === undefined || value === "") return "";
  return decimalToIt(Number(value));
}

function documentIdForFile(fileName, index) {
  return `doc_${index + 1}_${String(fileName || "file").replace(/[^\w.-]+/g, "_")}`;
}

async function fileToBase64(file) {
  const buffer = await file.arrayBuffer();
  const bytes = new Uint8Array(buffer);
  let binary = "";
  const chunkSize = 0x8000;
  for (let index = 0; index < bytes.length; index += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(index, index + chunkSize));
  }
  return btoa(binary);
}

function existingRecordSummary(records = []) {
  return records.map((record) => ({
    recordId: record.id,
    supplier: record.row?.Fornitore || "",
    invoiceNumber: record.row?.Fattura || "",
    invoiceDate: record.row?.["Data fattura"] || "",
    dueDate: record.row?.Scadenza || "",
    total: record.row?.Totale || "",
    outstanding: record.row?.["Da pagare ancora"] || "",
    paymentTerms: record.row?.["Termini pagamento fattura"] || "",
    payments: normalizeTransfers(record.transfers || record.transfer || []).map((transfer) => ({
      paymentType: transfer.paymentType || "",
      total: transfer.total || "",
      beneficiary: transfer.beneficiary || "",
      reason: transfer.reason || "",
      noticeNumber: transfer.noticeNumber || "",
      dueDate: transfer.dueDate || "",
      executionDate: transfer.executionDate || "",
    })),
  }));
}

function existingPendingPaymentSummary(pendingPayments = []) {
  return pendingPayments.map((entry) => ({
    signature: entry.signature,
    status: entry.status,
    paymentType: entry.payment?.paymentType || "",
    supplier: entry.payment?.beneficiary || "",
    total: entry.payment?.total || "",
    dueDate: entry.payment?.dueDate || "",
    executionDate: entry.payment?.executionDate || "",
    documentDate: entry.payment?.documentDate || "",
    bank: entry.payment?.bank || "",
    payerIban: entry.payment?.payerIban || "",
    reason: entry.payment?.reason || "",
    noticeNumber: entry.payment?.noticeNumber || "",
    flowName: entry.payment?.flowName || "",
    notes: entry.notes || "",
  }));
}

function geminiPrompt(fileIds, existingRecords, pendingPayments) {
  return `
Sei il motore ufficiale di estrazione e matching contabile di Martec.

Hai in input PDF di fatture e PDF di pagamenti (bonifici, RIBA, PayPal, carta di credito).
Devi:
1. classificare ogni documento o effetto come invoice o payment;
2. estrarre i campi strutturati;
3. decidere gli abbinamenti ufficiali;
4. segnalare le fatture duplicate già presenti in archivio;
5. usare i record già presenti quando un pagamento del nuovo gruppo si riferisce a una fattura già esistente.
6. usare anche i pagamenti gia' in sospeso quando il nuovo upload contiene la fattura corretta.

Regole fondamentali:
- Per i RIBA, il campo più importante è "Riferimento Operazione".
- Un singolo pagamento può coprire più fatture.
- Una singola fattura può avere più pagamenti.
- Se un effetto RIBA contiene riferimenti a più fatture, crea più allocazioni.
- Non inventare abbinamenti se il riferimento non è abbastanza forte.
- Considera l'archivio esistente come fonte valida per collegare pagamenti che non trovano la fattura nel batch corrente.
- Considera anche i pagamenti gia' memorizzati come "in sospeso": possono essere il vero pagamento di una nuova fattura caricata adesso.
- Se una fattura caricata è già presente in archivio, non trattarla come nuova: inseriscila in duplicateInvoices.
- Usa importi numerici senza simboli valuta.
- Usa date nel formato DD/MM/YYYY quando possibile.

Documenti caricati in questa richiesta:
${JSON.stringify(fileIds, null, 2)}

Archivio già presente:
${JSON.stringify(existingRecordSummary(existingRecords), null, 2)}

Pagamenti gia' in sospeso:
${JSON.stringify(existingPendingPaymentSummary(pendingPayments), null, 2)}

Restituisci solo JSON conforme allo schema richiesto.
`;
}

async function callGeminiMatcher(files, existingRecords = [], pendingPayments = []) {
  const { supabaseUrl, supabaseAnonKey, geminiModel } = getAiConfig();
  if (!supabaseUrl || !supabaseAnonKey) return null;
  const fileIds = files.map((file, index) => ({ fileName: file.name, documentId: documentIdForFile(file.name, index) }));
  const documents = await Promise.all(files.map(async (file, index) => ({
    id: fileIds[index].documentId,
    fileName: file.name,
    mimeType: "application/pdf",
    data: await fileToBase64(file),
  })));
  let response;
  try {
    response = await fetch(`${supabaseUrl}/functions/v1/gemini-match`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${supabaseAnonKey}`,
        apikey: supabaseAnonKey,
      },
      body: JSON.stringify({
        model: geminiModel,
        prompt: geminiPrompt(fileIds, existingRecords, pendingPayments),
        documents,
      }),
    });
  } catch (error) {
    if (isFetchFailure(error)) {
      throw new Error("Connessione non riuscita verso la funzione AI di Supabase. Controlla internet e che la funzione sia pubblicata.");
    }
    throw error;
  }
  if (response.ok) {
    const json = await response.json();
    return { payload: json, fileIds, modelUsed: json.modelUsed || geminiModel };
  }
  const text = await response.text();
  if (geminiPermissionProblem(response.status, text)) {
    return {
      payload: null,
      fileIds,
      modelUsed: null,
      fallbackReason: "Gemini disattivato: la chiave server-side e' bloccata o revocata.",
    };
  }
  if (response.status === 404) {
    return {
      payload: null,
      fileIds,
      modelUsed: null,
      fallbackReason: "Funzione AI non pubblicata su Supabase",
    };
  }
  if (shouldFallbackFromGemini(response.status, text)) {
    await sleep(1200);
    return {
      payload: null,
      fileIds,
      modelUsed: null,
      fallbackReason: `Gemini temporaneamente non disponibile (${response.status})`,
    };
  }
  throw new Error(`La funzione AI non ha accettato la richiesta (${response.status}). ${text.slice(0, 180)}`);
}

function findTransferAmount(lines, fullText) {
  const directPatterns = [
    "Totale:\\s+(\\d{1,3}(?:\\.\\d{3})*,\\d{2})\\s+EUR",
    "Totale bonifico:?\\s+(\\d{1,3}(?:\\.\\d{3})*,\\d{2})",
    "Importo bonifico:?\\s+(\\d{1,3}(?:\\.\\d{3})*,\\d{2})",
    "Importo:?\\s+(\\d{1,3}(?:\\.\\d{3})*,\\d{2})\\s+EUR",
    "EUR\\s+(\\d{1,3}(?:\\.\\d{3})*,\\d{2})",
  ];
  for (const pattern of directPatterns) {
    const value = firstMatch(pattern, fullText, "i");
    if (value) return value;
  }

  const keywordLines = lines.filter((line) => /(totale|importo|bonifico|addebito|beneficiario)/i.test(line));
  for (const line of keywordLines) {
    const amounts = moneyCandidates(line);
    if (amounts.length) return amounts[amounts.length - 1];
  }

  const allAmounts = moneyCandidates(fullText)
    .map((value) => ({ value, numeric: Number.parseFloat(value.replace(/\./g, "").replace(",", ".")) }))
    .filter((entry) => Number.isFinite(entry.numeric))
    .sort((a, b) => b.numeric - a.numeric);
  return allAmounts[0]?.value || "";
}

function findTransferCurrency(fullText) {
  const explicit = firstMatch("Totale:?\\s+\\d{1,3}(?:\\.\\d{3})*,\\d{2}\\s+(EUR|USD)", fullText, "i");
  return detectDocumentCurrency(fullText, explicit);
}

function detectInvoicePaymentMethod(fullText, explicitValue = "") {
  const source = `${explicitValue}\n${fullText}`.toUpperCase();
  if (/PAYPAL/.test(source)) return "PayPal";
  if (/CARTA DI CREDITO|CARTA DI PAGAMENTO|MP08/.test(source)) return "Carta di credito";
  if (/\bRIBA\b|\bMP12\b/.test(source)) return "RIBA";
  if (/BONIFICO|CREDIT TRANSFER|SEPA|MP05/.test(source)) return "Bonifico";
  return cleanText(explicitValue || "");
}

export async function parseRibaEffects(file) {
  const items = await extractPdfItems(file);
  const fullText = linesFromItems(items).join("\n");
  if (!/(Pagamento Effetti|Dettaglio Presentazione)/i.test(fullText) || !/Dati effetto/i.test(fullText)) return [];

  const bank = firstMatch("(INTESA SANPAOLO S\\.P\\.A\\.)", fullText)
    || firstMatch("ABI:\\s*\\d+\\s*-\\s*([A-Z ]+)", fullText, "i");
  const payer = valueOnSameRow(items, "Ragione Sociale:", { valueMinX: 120 })
    || firstMatch("Ragione Sociale:\\s+(.+?)\\s+Codice SIA:", fullText, "is");
  const payerIban = firstMatch("Conto di addebito:\\s*(IT\\d{2}[A-Z]\\d{22})", fullText, "i");
  const documentDate = normalizeDate(firstMatch("Data:\\s+(\\d{2}\\.\\d{2}\\.\\d{4})", fullText, "i"));
  const flowName = firstMatch("Nome Flusso:\\s+(.+?)\\s+Data/Ora:", fullText, "is");

  const effects = [];
  const effectBlocks = fullText
    .split(/Dati effetto/gi)
    .map((block) => cleanText(block))
    .filter((block) => /Importo/i.test(block) && /Creditore/i.test(block));

  for (const block of effectBlocks) {
    const amount = firstMatch("Importo\\s+([\\d.]+,\\d{2})\\s+EUR", block, "i")
      || firstMatch("Importo\\s+([\\d.]+,\\d{2})", block, "i");
    const dueDate = firstMatch("Scadenza\\s+(\\d{2}\\.\\d{2}\\.\\d{4})", block, "i");
    const beneficiary = firstMatch("Creditore\\s+(.+?)\\s+Codice Fiscale \\/ Partita IVA", block, "is")
      || firstMatch("Creditore\\s+(.+?)\\s+Debitore su avviso", block, "is");
    const beneficiaryVat = firstMatch("Codice Fiscale \\/ Partita IVA\\s+([A-Z0-9]+)", block, "i");
    const noticeNumber = firstMatch("Numero avviso\\s+([0-9]+)", block, "i");
    const referenceOperation = firstMatch("Riferimento Operazione\\s+(.+)$", block, "is");
    if (!amount || !beneficiary) continue;
    effects.push({
      type: "transfer",
      paymentType: "RIBA",
      documentDate,
      executionDate: documentDate,
      dueDate: normalizeDate(dueDate),
      total: cleanText(amount),
      currency: "EUR",
      totalEur: decimalFromIt(amount),
      bank: cleanText(bank),
      payer: cleanText(payer),
      payerIban: cleanText(payerIban),
      beneficiary: cleanText(beneficiary),
      beneficiaryVat: cleanText(beneficiaryVat),
      beneficiaryIban: "",
      swift: "",
      reason: cleanText(referenceOperation),
      noticeNumber: cleanText(noticeNumber),
      flowName: cleanText(flowName),
      rawText: fullText,
    });
  }

  return effects;
}

export async function extractPdfItems(file) {
  const data = file instanceof ArrayBuffer ? file : await file.arrayBuffer();
  const loadingTask = pdfjsLib.getDocument({ data });
  const pdf = await loadingTask.promise;
  const items = [];
  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const content = await page.getTextContent();
    content.items.forEach((item) => {
      const [, , , , x, y] = item.transform;
      const text = cleanText(item.str || "");
      if (!text) return;
      items.push({
        x,
        y: y - (pageNumber - 1) * 1500,
        text,
      });
    });
  }
  return items.sort((a, b) => b.y - a.y || a.x - b.x);
}

export function linesFromItems(items, tolerance = 1.5) {
  const rows = [];
  items.forEach((item) => {
    if (!rows.length || Math.abs(rows[rows.length - 1][0].y - item.y) > tolerance) rows.push([item]);
    else rows[rows.length - 1].push(item);
  });
  return rows
    .map((row) => row.sort((a, b) => a.x - b.x).map((item) => cleanText(item.text)).filter(Boolean).join(" "))
    .map(cleanText)
    .filter(Boolean);
}

export async function parseInvoice(file) {
  const items = await extractPdfItems(file);
  const fullText = linesFromItems(items).join("\n");
  const supplier = valueOnSameRow(items, "Denominazione:", { labelMaxX: 80, valueMinX: 70, valueMaxX: 260 });
  const vatIds = [...fullText.matchAll(/Identificativo fiscale ai fini IVA:\s+([A-Z]{2}\d+)/gi)].map((match) => match[1]);
  const supplierVat = vatIds[0] || "";
  const invoiceNumber = firstMatch("Numero documento\\s+Data documento.*?\\n.*?([A-Z0-9./ -]{2,})\\s+\\d{2}-\\d{2}-\\d{4}", fullText, "is");
  const invoiceDate = firstMatch("Numero documento\\s+Data documento.*?\\n.*?[A-Z0-9./ -]{2,}\\s+(\\d{2}-\\d{2}-\\d{4})", fullText, "is");
  let taxable = moneyBelowLabel(items, "Totale imponibile", { xMin: 430, xMax: 510 });
  let vat = moneyBelowLabel(items, "Totale imposta", { xMin: 525, xMax: 585 });
  let total = moneyBelowLabel(items, "Totale documento", { xMin: 520, xMax: 585 });
  if (!taxable) taxable = firstMatch("Totale imponibile\\s+Totale imposta.*?\\n.*?\\d+,\\d{2}\\s+(\\d+,\\d{2})\\s+\\d+,\\d{2}", fullText, "is");
  if (!vat) vat = firstMatch("Totale imponibile\\s+Totale imposta.*?\\n.*?\\d+,\\d{2}\\s+\\d+,\\d{2}\\s+(\\d+,\\d{2})", fullText, "is");
  if (!total) total = firstMatch("Totale documento\\s+(\\d+,\\d{2})", fullText, "i");
  const iban = firstMatch("\\b(IT\\d{2}[A-Z]\\d{22})\\b", fullText, "i");
  const dueDate = firstMatch("Data scadenza\\s+Importo.*?\\n.*?(\\d{2}-\\d{2}-\\d{4})", fullText, "is");
  const explicitPaymentTerms = valueOnSameRow(items, "Modalità pagamento")
    || firstMatch("Modalità pagamento\\s+(.+)", fullText, "i");
  const paymentTerms = detectInvoicePaymentMethod(fullText, explicitPaymentTerms);
  const explicitCurrency = firstMatch("Divisa\\s+([A-Z]{3})", fullText, "i");
  const currency = detectDocumentCurrency(fullText, explicitCurrency);
  const amounts = await normalizeCurrencyAmounts({
    currency,
    date: normalizeDate(invoiceDate),
    taxable,
    vat,
    total,
  });
  return {
    type: "invoice",
    supplier,
    supplierVat,
    invoiceNumber,
    invoiceDate: normalizeDate(invoiceDate),
    taxable,
    vat,
    total,
    currency,
    ...amounts,
    iban,
    dueDate: normalizeDate(dueDate),
    paymentTerms,
    rawText: fullText,
  };
}

export async function parseTransfer(file) {
  const items = await extractPdfItems(file);
  const lines = linesFromItems(items);
  const fullText = lines.join("\n");
  const total = findTransferAmount(lines, fullText);
  const currency = findTransferCurrency(fullText);
  const executionDate = normalizeDate(firstMatch("Data esecuzione:\\s+(\\d{2}\\.\\d{2}\\.\\d{4})", fullText));
  const documentDate = normalizeDate(firstMatch("Data:\\s+(\\d{2}\\.\\d{2}\\.\\d{4})", fullText));
  const amounts = await normalizeCurrencyAmounts({
    currency,
    date: executionDate || documentDate,
    total,
  });
  return {
    type: "transfer",
    documentDate,
    bank: firstMatch("(INTESA SANPAOLO S\\.P\\.A\\.)", fullText),
    payer: valueOnSameRow(items, "Ragione Sociale:", { valueMinX: 130 }),
    payerIban: firstMatch("Conto ordinante\\s+(IT\\d{2}[A-Z]\\d{22})", fullText),
    executionDate,
    total,
    currency,
    ...amounts,
    beneficiary: valueOnSameRow(items, "Beneficiario", { valueMinX: 150 }),
    beneficiaryIban: firstMatch("Conto beneficiario\\s+(IT\\d{2}[A-Z]\\d{22})", fullText),
    swift: firstMatch("Codice SWIFT\\s+([A-Z0-9]+)", fullText),
    reason: firstMatch("Informazioni aggiuntive \\(max\\s+(.+)", fullText),
    rawText: fullText,
  };
}

export async function classifyPdf(file) {
  const ribaEffects = await parseRibaEffects(file);
  if (ribaEffects.length) return ribaEffects;
  const invoice = await parseInvoice(file);
  const transfer = await parseTransfer(file);
  const invoiceScore = ["supplier", "invoiceNumber", "invoiceDate", "total"].filter((key) => Boolean(invoice[key])).length;
  const transferScore = ["beneficiary", "beneficiaryIban", "reason", "executionDate", "total"].filter((key) => Boolean(transfer[key])).length;
  return [transferScore > invoiceScore ? transfer : invoice];
}

function columnIndex(ref) {
  const match = String(ref || "").match(/^([A-Z]+)/);
  if (!match) return 0;
  return match[1].split("").reduce((total, char) => total * 26 + char.charCodeAt(0) - 64, 0);
}

function excelSerialToDate(value) {
  const number = Number(value);
  if (!Number.isFinite(number)) return value;
  const base = new Date(Date.UTC(1899, 11, 30));
  base.setUTCDate(base.getUTCDate() + number);
  return `${String(base.getUTCDate()).padStart(2, "0")}/${String(base.getUTCMonth() + 1).padStart(2, "0")}/${base.getUTCFullYear()}`;
}

export async function parseXlsxRows(file) {
  const zip = await JSZip.loadAsync(file instanceof Blob ? await file.arrayBuffer() : file);
  const parser = new DOMParser();
  const workbook = parser.parseFromString(await zip.file("xl/workbook.xml").async("string"), "application/xml");
  const rels = parser.parseFromString(await zip.file("xl/_rels/workbook.xml.rels").async("string"), "application/xml");
  const firstSheet = workbook.getElementsByTagNameNS(NS_MAIN, "sheet")[0];
  if (!firstSheet) return [];
  const rid = firstSheet.getAttribute("r:id") || firstSheet.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
  const relationship = [...rels.getElementsByTagName("Relationship")].find((node) => node.getAttribute("Id") === rid);
  if (!relationship) return [];
  const sheetPath = `xl/${relationship.getAttribute("Target").replace(/^\/?/, "").replace(/^xl\//, "")}`;
  const sheet = parser.parseFromString(await zip.file(sheetPath).async("string"), "application/xml");
  const sharedStrings = zip.file("xl/sharedStrings.xml")
    ? parser.parseFromString(await zip.file("xl/sharedStrings.xml").async("string"), "application/xml")
    : null;
  const shared = sharedStrings
    ? [...sharedStrings.getElementsByTagNameNS(NS_MAIN, "si")].map((node) =>
        [...node.getElementsByTagNameNS(NS_MAIN, "t")].map((textNode) => textNode.textContent || "").join(""),
      )
    : [];
  const rows = new Map();
  [...sheet.getElementsByTagNameNS(NS_MAIN, "c")].forEach((cell) => {
    const ref = cell.getAttribute("r") || "";
    const rowNumber = Number((ref.match(/\d+/) || ["0"])[0]);
    const valueNode = cell.getElementsByTagNameNS(NS_MAIN, "v")[0];
    if (!rowNumber || !valueNode) return;
    const value = cell.getAttribute("t") === "s" ? shared[Number(valueNode.textContent || 0)] || "" : valueNode.textContent || "";
    rows.set(rowNumber, rows.get(rowNumber) || {});
    rows.get(rowNumber)[columnIndex(ref)] = value;
  });
  const imported = [];
  const headerValues = rows.get(1) || {};
  const headerByIndex = new Map(
    Object.entries(headerValues).map(([index, value]) => [Number(index), cleanText(value)]),
  );
  [...rows.keys()].sort((a, b) => a - b).forEach((rowNumber) => {
    if (rowNumber <= 2) return;
    const values = rows.get(rowNumber);
    if (![2, 3, 4, 5, 6].some((col) => values[col])) return;
    const row = {};
    Object.entries(values).forEach(([index, rawValue]) => {
      const column = headerByIndex.get(Number(index));
      if (!column) return;
      let value = rawValue || "";
      if (["Data", "Data fattura", "Scadenza"].includes(column) && /^\d+(\.\d+)?$/.test(value)) value = excelSerialToDate(value);
      row[column] = value;
    });
    imported.push(normalizeRow(row));
  });
  return imported;
}

function escapeXml(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function cellRef(row, col) {
  let letters = "";
  let current = col;
  while (current) {
    const rem = (current - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    current = Math.floor((current - 1) / 26);
  }
  return `${letters}${row}`;
}

function excelSerialFromDate(value) {
  const normalized = normalizeDate(value);
  const match = normalized.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!match) return "";
  const date = new Date(Date.UTC(Number(match[3]), Number(match[2]) - 1, Number(match[1])));
  const base = new Date(Date.UTC(1899, 11, 30));
  return String(Math.round((date - base) / 86400000));
}

function decimalAsExcel(value) {
  const amount = decimalFromIt(value);
  if (amount === null) return "";
  return Number(amount.toFixed(3)).toString().replace(/\.0+$/, "").replace(/(\.\d*[1-9])0+$/, "$1");
}

async function getTemplateParts() {
  if (templateCache.parts) return templateCache.parts;
  let response;
  try {
    response = await fetch(new URL("assets/martec-template.xlsx", window.location.href));
  } catch (error) {
    if (isFetchFailure(error)) {
      throw new Error("Impossibile caricare il template Excel. Apri la web app da http/https e verifica che il file assets/martec-template.xlsx sia pubblicato.");
    }
    throw error;
  }
  if (!response.ok) {
    throw new Error("Template Excel non trovato. Verifica che assets/martec-template.xlsx sia presente nella pubblicazione.");
  }
  const zip = await JSZip.loadAsync(await response.arrayBuffer());
  const parser = new DOMParser();
  const serializer = new XMLSerializer();
  const sheet = parser.parseFromString(await zip.file("xl/worksheets/sheet1.xml").async("string"), "application/xml");
  const part = (tagName) => {
    const node = sheet.getElementsByTagNameNS(NS_MAIN, tagName)[0];
    return node ? serializer.serializeToString(node) : "";
  };
  templateCache.parts = {
    styles: await zip.file("xl/styles.xml").async("uint8array"),
    theme: await zip.file("xl/theme/theme1.xml").async("uint8array"),
    core: await zip.file("docProps/core.xml").async("uint8array"),
    app: await zip.file("docProps/app.xml").async("uint8array"),
    sheetViews: part("sheetViews"),
    sheetFormatPr: part("sheetFormatPr"),
    cols: part("cols"),
    mergeCells: part("mergeCells"),
    pageMargins: part("pageMargins"),
  };
  return templateCache.parts;
}

function buildSheetXml(records, template) {
  const totals = { Imponibile: 0, IVA: 0, Totale: 0, Uscite: 0 };
  const rows = records.map(exportRowFromRecord);
  rows.forEach((row) => {
    totals.Imponibile += decimalFromIt(row[6]) || 0;
    totals.IVA += decimalFromIt(row[7]) || 0;
    totals.Totale += decimalFromIt(row[8]) || 0;
    totals.Uscite += decimalFromIt(row[9]) || 0;
  });
  const endRow = Math.max(3, rows.length + 2);
  const lastColumnRef = cellRef(1, EXPORT_TEMPLATE_COLUMNS.length).replace(/\d+/g, "");
  const columnCount = EXPORT_TEMPLATE_COLUMNS.length;
  const headerStyles = [118, 108, 116, 120, 116, 1, 1, 2, 2, 2, 106, 108, 110, 97, 112, 114];
  const totalsStyles = [119, 109, 117, 121, 117, 3, 3, 4, 4, 4, 107, 109, 111, 98, 113, 115];
  const dataStyles = [18, 19, 21, 22, 23, 10, 24, 25, 26, 23, 27, 27, 28, 100, 17, 63];

  const headerCells = EXPORT_TEMPLATE_COLUMNS.map((label, index) =>
    `<c r="${cellRef(1, index + 1)}" s="${headerStyles[index]}" t="inlineStr"><is><t>${escapeXml(label)}</t></is></c>`,
  ).join("");

  const totalsKeys = ["Imponibile", "IVA", "Totale", "Uscite", null];
  const totalsCells = totalsStyles.map((style, index) => {
    const ref = cellRef(2, index + 1);
    if (index < 5 || index > 9) return `<c r="${ref}" s="${style}" />`;
    const key = totalsKeys[index - 5];
    if (!key) return `<c r="${ref}" s="${style}" />`;
    return `<c r="${ref}" s="${style}"><f>SUBTOTAL(9,${cellRef(3, index + 1)}:${cellRef(endRow, index + 1)})</f><v>${decimalAsExcel(decimalToIt(totals[key])) || "0"}</v></c>`;
  }).join("");

  const bodyRows = rows.map((row, rowIndex) => {
    const r = rowIndex + 3;
    const cells = row.map((value, index) => {
      const ref = cellRef(r, index + 1);
      const style = dataStyles[index];
      if ([11, 12].includes(index + 1)) {
        const serial = excelSerialFromDate(value);
        return serial ? `<c r="${ref}" s="${style}"><v>${serial}</v></c>` : `<c r="${ref}" s="${style}" />`;
      }
      if ([1, 4, 5, 6, 7, 8, 9].includes(index + 1)) {
        const number = decimalAsExcel(value);
        return number ? `<c r="${ref}" s="${style}"><v>${number}</v></c>` : `<c r="${ref}" s="${style}" />`;
      }
      return value ? `<c r="${ref}" s="${style}" t="inlineStr"><is><t>${escapeXml(value)}</t></is></c>` : `<c r="${ref}" s="${style}" />`;
    }).join("");
    return `<row r="${r}" spans="1:${columnCount}" ht="28" customHeight="1">${cells}</row>`;
  }).join("");

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="${NS_MAIN}">
<dimension ref="A1:${lastColumnRef}${Math.max(2, rows.length + 2)}" />
${template.sheetViews}
${template.sheetFormatPr}
${template.cols}
<sheetData>
<row r="1" spans="1:${columnCount}" ht="58.8" customHeight="1">${headerCells}</row>
<row r="2" spans="1:${columnCount}" ht="22.5" customHeight="1" thickBot="1">${totalsCells}</row>
${bodyRows}
</sheetData>
<autoFilter ref="A2:${lastColumnRef}${Math.max(2, rows.length + 2)}" />
${template.mergeCells}
${template.pageMargins}
</worksheet>`;
}

export async function buildXlsx(records) {
  const template = await getTemplateParts();
  const zip = new JSZip();
  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`);
  zip.folder("_rels").file(".rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`);
  zip.folder("docProps").file("core.xml", template.core);
  zip.folder("docProps").file("app.xml", template.app);
  zip.folder("xl").file("workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="${NS_MAIN}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<bookViews><workbookView xWindow="-108" yWindow="-108" windowWidth="23256" windowHeight="12456"/></bookViews>
<sheets><sheet name="MARTEC" sheetId="1" r:id="rId1"/></sheets>
<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">MARTEC!$A$2:$P$${Math.max(2, records.length + 2)}</definedName></definedNames>
<calcPr calcId="181029"/>
</workbook>`);
  zip.folder("xl").folder("_rels").file("workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>`);
  zip.folder("xl").folder("worksheets").file("sheet1.xml", buildSheetXml(records.map((record) => ({ ...record, row: normalizeRow(record.row) })), template));
  zip.folder("xl").file("styles.xml", template.styles);
  zip.folder("xl").folder("theme").file("theme1.xml", template.theme);
  return zip.generateAsync({ type: "blob" });
}

async function normalizeAiDocument(raw) {
  if (raw.type === "invoice") {
    const amounts = await normalizeCurrencyAmounts({
      currency: raw.currency || "EUR",
      date: normalizeDate(raw.invoiceDate || ""),
      taxable: numberToIt(raw.taxable),
      vat: numberToIt(raw.vat),
      total: numberToIt(raw.total),
    });
    return {
      type: "invoice",
      aiId: raw.id,
      fileName: raw.fileName,
      supplier: cleanText(raw.supplier),
      supplierVat: cleanText(raw.supplierVat),
      invoiceNumber: cleanText(raw.invoiceNumber),
      invoiceDate: normalizeDate(raw.invoiceDate),
      taxable: numberToIt(raw.taxable),
      vat: numberToIt(raw.vat),
      total: numberToIt(raw.total),
      currency: detectDocumentCurrency("", raw.currency),
      ...amounts,
      iban: "",
      dueDate: normalizeDate(raw.dueDate),
      paymentTerms: detectInvoicePaymentMethod("", raw.paymentType),
      rawText: cleanText(raw.notes || ""),
    };
  }
  const amounts = await normalizeCurrencyAmounts({
    currency: raw.currency || "EUR",
    date: normalizeDate(raw.executionDate || raw.documentDate || ""),
    total: numberToIt(raw.total),
  });
  return {
    type: "transfer",
    aiId: raw.id,
    fileName: raw.fileName,
    paymentType: cleanText(raw.paymentType || "Bonifico"),
    documentDate: normalizeDate(raw.documentDate),
    executionDate: normalizeDate(raw.executionDate),
    dueDate: normalizeDate(raw.dueDate),
    total: numberToIt(raw.total),
    currency: detectDocumentCurrency("", raw.currency),
    ...amounts,
    bank: cleanText(raw.bank),
    payer: cleanText(raw.payer),
    payerIban: cleanText(raw.payerIban),
    beneficiary: cleanText(raw.beneficiary),
    beneficiaryVat: cleanText(raw.beneficiaryVat),
    beneficiaryIban: cleanText(raw.beneficiaryIban),
    swift: cleanText(raw.swift),
    reason: cleanText(raw.reason),
    noticeNumber: cleanText(raw.noticeNumber),
    flowName: cleanText(raw.flowName),
    rawText: cleanText(raw.notes || ""),
  };
}

async function classifyDocumentsWithGemini(files, existingRecords, pendingPayments) {
  const aiResult = await callGeminiMatcher(files, existingRecords, pendingPayments);
  if (!aiResult || !aiResult.payload) {
    return aiResult?.fallbackReason
      ? { aiUsed: false, aiFallbackReason: aiResult.fallbackReason }
      : null;
  }
  const { payload } = aiResult;
  const normalizedDocs = await Promise.all((payload.documents || []).map((doc) => normalizeAiDocument(doc)));
  const invoices = normalizedDocs.filter((doc) => doc.type === "invoice");
  const transfers = normalizedDocs.filter((doc) => doc.type === "transfer");
  const invoiceById = new Map(invoices.map((doc) => [doc.aiId, doc]));
  const transferById = new Map(transfers.map((doc) => [doc.aiId, doc]));
  const pendingBySignature = new Map((pendingPayments || []).map((entry) => [entry.signature, entry.payment || {}]));
  const matches = (payload.invoiceMatches || [])
    .map((entry) => {
      const invoice = invoiceById.get(entry.invoiceDocumentId);
      if (!invoice) return null;
      const matchedTransfers = (entry.paymentAllocations || [])
        .map((allocation) => {
          const transfer = transferById.get(allocation.paymentDocumentId);
          if (!transfer) return null;
          return {
            ...transfer,
            allocatedAmount: allocation.allocatedAmount,
            total: numberToIt(allocation.allocatedAmount),
            totalEur: allocation.allocatedAmount,
            allocationConfidence: entry.confidence,
          };
        })
        .filter(Boolean);
      return {
        invoice,
        transfers: matchedTransfers,
        match: {
          score: Math.round(Number(entry.confidence || 0) * 100),
          threshold: 70,
          matched: matchedTransfers.length > 0,
          reasons: [cleanText(entry.rationale)].filter(Boolean),
          issues: matchedTransfers.length ? [] : ["Gemini non ha trovato un pagamento affidabile"],
          summary: matchedTransfers.map((transfer) =>
            `${transfer.beneficiary || "-"} | ${transfer.total || "-"} ${transfer.currency || "EUR"} | ${transfer.dueDate || transfer.executionDate || transfer.documentDate || ""}`.trim(),
          ).join(" || "),
        },
      };
    })
    .filter(Boolean);
  const existingRecordMatches = (payload.existingRecordMatches || []).map((entry) => ({
    recordId: entry.recordId,
    confidence: Number(entry.confidence || 0),
    rationale: cleanText(entry.rationale),
    transfers: (entry.paymentAllocations || [])
      .map((allocation) => {
        const transfer = transferById.get(allocation.paymentDocumentId);
        if (!transfer) return null;
        return {
          ...transfer,
          allocatedAmount: allocation.allocatedAmount,
          total: numberToIt(allocation.allocatedAmount),
          totalEur: allocation.allocatedAmount,
        };
      })
      .filter(Boolean),
  }));
  const pendingPaymentMatches = (payload.existingPendingPaymentMatches || [])
    .map((entry) => {
      const invoice = invoiceById.get(entry.invoiceDocumentId);
      if (!invoice) return null;
      const matchedTransfers = (entry.pendingAllocations || [])
        .map((allocation) => {
          const transfer = pendingBySignature.get(allocation.pendingSignature);
          if (!transfer) return null;
          return {
            ...transfer,
            allocatedAmount: allocation.allocatedAmount,
            total: numberToIt(allocation.allocatedAmount),
            totalEur: allocation.allocatedAmount,
          };
        })
        .filter(Boolean);
      return {
        invoiceDocumentId: entry.invoiceDocumentId,
        confidence: Number(entry.confidence || 0),
        rationale: cleanText(entry.rationale),
        transfers: matchedTransfers,
      };
    })
    .filter(Boolean);
  const duplicateInvoices = (payload.duplicateInvoices || []).map((entry) => ({
    invoiceDocumentId: entry.invoiceDocumentId,
    existingRecordId: entry.existingRecordId,
    rationale: cleanText(entry.rationale),
  }));
  const handledInvoiceIds = new Set([
    ...matches.map((entry) => entry.invoice.aiId),
    ...duplicateInvoices.map((entry) => entry.invoiceDocumentId),
  ]);
  for (const invoice of invoices) {
    if (handledInvoiceIds.has(invoice.aiId)) continue;
    matches.push({
      invoice,
      transfers: [],
      match: {
        score: 0,
        threshold: 70,
        matched: false,
        reasons: [],
        issues: ["Gemini non ha associato alcun pagamento a questa fattura"],
        summary: "",
      },
    });
  }
  return {
    aiUsed: true,
    aiModelUsed: aiResult.modelUsed,
    invoices,
    transfers,
    matches,
    existingRecordMatches,
    pendingPaymentMatches,
    duplicateInvoices,
  };
}

export async function classifyDocuments(files, existingRecords = [], pendingPayments = []) {
  const aiResult = await classifyDocumentsWithGemini(files, existingRecords, pendingPayments);
  if (aiResult?.aiUsed) return aiResult;
  const parsed = [];
  for (const file of files) {
    const docs = await classifyPdf(file);
    docs.forEach((parsedDocument) => parsed.push({ file, parsed: parsedDocument }));
  }
  const invoices = parsed.filter((item) => item.parsed.type !== "transfer").map((item) => ({ ...item.parsed, fileName: item.file.name }));
  const transfers = [
    ...parsed.filter((item) => item.parsed.type === "transfer").map((item) => ({ ...item.parsed, fileName: item.file.name })),
    ...(pendingPayments || []).map((entry) => ({ ...(entry.payment || {}) })),
  ];
  return {
    aiUsed: false,
    aiFallbackReason: aiResult?.aiFallbackReason || "",
    invoices,
    transfers,
    matches: invoices.map((invoice) => ({
      invoice,
      transfers: [],
      match: {
        score: 0,
        threshold: 70,
        matched: false,
        reasons: [],
        issues: ["In attesa del matching AI"],
        summary: "",
      },
    })),
    existingRecordMatches: [],
    pendingPaymentMatches: [],
    duplicateInvoices: [],
  };
}

export function recordsFromImportedRows(rows, startingIndex) {
  return rows.map((row, offset) => recordFromImportedRow(row, startingIndex + offset));
}

export { addDays, buildRecord, pairDocuments };
