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

function alphaPrefix(value = "") {
  const match = normalizeKey(value).match(/^[A-Z]+/);
  return match ? match[0] : "";
}

function normalizeVat(value = "") {
  return String(value || "").toUpperCase().replace(/[^A-Z0-9]+/g, "");
}

function invoiceReferenceVariants(value = "") {
  const source = String(value || "").toUpperCase();
  const normalized = normalizeKey(source);
  const compact = normalized.replace(/^(FT|FAT|FATTURA|DOC|DOCUMENTO|NUMERO)+/, "");
  const digitRuns = [...source.matchAll(/\d{2,}/g)].map((match) => match[0]);
  const variants = new Set([normalized, compact]);
  digitRuns.forEach((digits) => {
    const stripped = digits.replace(/^0+/, "") || digits;
    variants.add(digits);
    variants.add(stripped);
  });
  if (digitRuns.length >= 2) {
    const joined = digitRuns.join("");
    variants.add(joined);
    variants.add(joined.replace(/^0+/, "") || joined);
  }
  if (digitRuns.length) {
    const stripped = digitRuns[0].replace(/^0+/, "") || digitRuns[0];
    const prefix = alphaPrefix(source);
    if (prefix) {
      variants.add(`${digitRuns[0]}${prefix}`);
      variants.add(`${stripped}${prefix}`);
      variants.add(`${prefix}${stripped}`);
    }
  }
  for (const match of normalized.matchAll(/([A-Z]+)0*(\d{2,})/g)) {
    const [, letters, digits] = match;
    const stripped = digits.replace(/^0+/, "") || digits;
    variants.add(`${letters}${digits}`);
    variants.add(`${letters}${stripped}`);
    variants.add(`${stripped}${letters}`);
  }
  for (const match of normalized.matchAll(/(\d{2,})0*([A-Z]+)/g)) {
    const [, digits, letters] = match;
    const stripped = digits.replace(/^0+/, "") || digits;
    variants.add(`${digits}${letters}`);
    variants.add(`${stripped}${letters}`);
    variants.add(`${letters}${stripped}`);
  }
  return [...variants].filter((entry) => entry && entry.length >= 3);
}

function reasonReferenceMatch(invoiceNumber = "", transfer = {}) {
  const reasonText = [
    transfer.reason,
    transfer.noticeNumber,
    transfer.flowName,
    transfer.rawText,
  ].filter(Boolean).join(" ");
  const normalizedReason = normalizeKey(reasonText);
  if (!invoiceNumber || !normalizedReason) return { matched: false, strong: false, fragment: "" };
  const variants = invoiceReferenceVariants(invoiceNumber);
  const fragment = variants.find((variant) => normalizedReason.includes(variant));
  if (!fragment) return { matched: false, strong: false, fragment: "" };
  const full = normalizeKey(invoiceNumber);
  const strongFragment = fragment === full
    || fragment.length >= Math.max(5, full.length - 2)
    || /\d{5,}/.test(fragment)
    || (fragment.length >= 5 && /[A-Z]/.test(fragment) && /\d/.test(fragment));
  return {
    matched: true,
    strong: strongFragment,
    fragment,
  };
}

function isRibaTransfer(transfer = {}) {
  return normalizePaymentMethod(transfer.paymentType || transfer.paymentTerms || transfer.method) === "RIBA";
}

function cloneTransferForInvoice(transfer, invoice, sharedTotal) {
  const invoiceTotal = payableInvoiceAmount(invoice) ?? invoice.totalEur ?? decimalFromIt(invoice.total) ?? 0;
  return {
    ...transfer,
    total: decimalToIt(invoiceTotal),
    totalEur: invoiceTotal,
    originalTotal: transfer.total,
    originalTotalEur: transfer.totalEur ?? decimalFromIt(transfer.total),
    splitInvoices: sharedTotal,
    splitFromRiba: true,
    splitInvoiceNumber: invoice.invoiceNumber || "",
  };
}

function invoicePaymentTarget(invoice = {}) {
  return payableInvoiceAmount(invoice) ?? invoice.totalEur ?? decimalFromIt(invoice.total) ?? 0;
}

function findMatchingInvoiceSubset(candidates = [], targetTotal = 0) {
  const meaningful = candidates
    .map((entry) => ({ ...entry, amount: invoicePaymentTarget(entry.invoice) }))
    .filter((entry) => entry.amount > 0);
  if (!meaningful.length || !targetTotal) return [];

  let best = [];
  const limit = Math.min(meaningful.length, 16);
  const visit = (index, selected, sum) => {
    if (Math.abs(sum - targetTotal) <= 0.01) {
      if (
        selected.length > best.length
        || (
          selected.length === best.length
          && selected.reduce((total, entry) => total + entry.match.score, 0) > best.reduce((total, entry) => total + entry.match.score, 0)
        )
      ) {
        best = [...selected];
      }
      return;
    }
    if (index >= limit || sum > targetTotal + 0.01) return;
    visit(index + 1, [...selected, meaningful[index]], sum + meaningful[index].amount);
    visit(index + 1, selected, sum);
  };
  visit(0, [], 0);
  return best;
}

function explicitReferenceCandidates(candidates = [], targetTotal = 0) {
  const referenced = candidates
    .filter((entry) => entry.ref?.strong)
    .map((entry) => ({ ...entry, amount: invoicePaymentTarget(entry.invoice) }))
    .filter((entry) => entry.amount > 0);
  if (referenced.length < 2) return [];
  const sum = referenced.reduce((total, entry) => total + entry.amount, 0);
  return sum <= targetTotal + 0.01 ? referenced : [];
}

function expandRibaTransfers(invoices, transfers) {
  const expanded = [];
  for (const transfer of transfers) {
    if (!isRibaTransfer(transfer)) {
      expanded.push(transfer);
      continue;
    }
    const candidates = invoices
      .map((invoice) => {
        const match = matchScore(invoice, transfer);
        const ref = reasonReferenceMatch(invoice.invoiceNumber, transfer);
        return { invoice, match, ref };
      })
      .filter((entry) => entry.ref.matched && entry.match.score >= MATCH_THRESHOLD)
      .sort((left, right) => right.match.score - left.match.score);

    if (candidates.length < 2) {
      expanded.push(transfer);
      continue;
    }

    const transferTotal = transfer.totalEur ?? decimalFromIt(transfer.total) ?? 0;
    const splitCandidates = findMatchingInvoiceSubset(candidates, transferTotal);
    const explicitCandidates = splitCandidates.length >= 2 ? splitCandidates : explicitReferenceCandidates(candidates, transferTotal);
    if (explicitCandidates.length < 2) {
      expanded.push(transfer);
      continue;
    }

    const splitInvoices = explicitCandidates
      .map((entry) => entry.invoice.invoiceNumber || "")
      .filter(Boolean)
      .join(", ");
    explicitCandidates.forEach((entry) => {
      expanded.push(cloneTransferForInvoice(transfer, entry.invoice, splitInvoices));
    });
  }
  return expanded;
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

function normalizeInvoicePaymentSplits(invoice = {}) {
  return Array.isArray(invoice.paymentSplits)
    ? invoice.paymentSplits
      .map((split) => ({
        paymentType: normalizePaymentMethod(split.paymentType || ""),
        dueDate: normalizeDate(split.dueDate || ""),
        amount: split.amount || "",
        amountEur: split.amountEur ?? decimalFromIt(split.amount),
        bankReference: cleanText(split.bankReference || ""),
      }))
      .filter((split) => split.amountEur !== null && split.amountEur !== undefined)
    : [];
}

function bestInvoiceSplitMatch(invoice = {}, transfer = {}) {
  const splits = normalizeInvoicePaymentSplits(invoice);
  if (!splits.length) return null;
  const transferAmount = transfer.totalEur ?? decimalFromIt(transfer.total);
  const transferMethod = normalizePaymentMethod(transfer.paymentType || transfer.paymentTerms || transfer.method);
  const transferDate = normalizeDate(transfer.dueDate || transfer.executionDate || transfer.documentDate || "");
  const ranked = splits.map((split) => {
    let score = 0;
    if (transferMethod && split.paymentType && transferMethod === split.paymentType) score += 50;
    else if (transferMethod === "BONIFICO" && split.paymentType === "RIBA") score -= 20;
    else if (transferMethod === "RIBA" && split.paymentType && split.paymentType !== "RIBA") score -= 20;
    if (transferAmount !== null && split.amountEur !== null && Math.abs(split.amountEur - transferAmount) < 0.00001) score += 80;
    if (transferDate && split.dueDate && transferDate === split.dueDate) score += 60;
    return { split, score };
  }).sort((left, right) => right.score - left.score);
  return ranked[0]?.score ? ranked[0] : null;
}

function sortSplitsByDueDate(splits = []) {
  return [...splits].sort((left, right) => {
    const leftDate = normalizeDate(left.dueDate || "");
    const rightDate = normalizeDate(right.dueDate || "");
    if (!leftDate && !rightDate) return 0;
    if (!leftDate) return 1;
    if (!rightDate) return -1;
    const [ld, lm, ly] = leftDate.split("/").map(Number);
    const [rd, rm, ry] = rightDate.split("/").map(Number);
    return new Date(ly, lm - 1, ld) - new Date(ry, rm - 1, rd);
  });
}

function splitLabel(split = {}) {
  const parts = [];
  if (split.paymentType) parts.push(split.paymentType);
  if (split.amountEur !== null && split.amountEur !== undefined) parts.push(`${decimalToIt(split.amountEur)} €`);
  if (split.dueDate) parts.push(split.dueDate);
  return parts.join(" - ");
}

function allocateTransfersToInvoiceSplits(invoice = {}, transfers = []) {
  const splits = sortSplitsByDueDate(normalizeInvoicePaymentSplits(invoice));
  if (!splits.length) {
    return {
      splits: [],
      paidSplits: [],
      unpaidSplits: [],
      partialSplits: [],
      matchedTransfers: [],
      summary: "",
    };
  }

  const remaining = splits.map((split, index) => ({
    ...split,
    index,
    remaining: split.amountEur ?? 0,
    paid: 0,
    transfers: [],
  }));

  const normalizedTransfers = normalizeTransfers(transfers)
    .map((transfer) => ({
      ...transfer,
      totalEur: transfer.totalEur ?? decimalFromIt(transfer.total),
    }))
    .filter((transfer) => transfer.totalEur !== null && transfer.totalEur !== undefined);

  const assignedTransfers = [];
  for (const transfer of normalizedTransfers) {
    const ranked = remaining
      .filter((split) => split.remaining > 0.00001)
      .map((split) => {
        let score = 0;
        const transferMethod = normalizePaymentMethod(transfer.paymentType || transfer.paymentTerms || transfer.method);
        if (split.paymentType && transferMethod && split.paymentType === transferMethod) score += 50;
        if (Math.abs((split.amountEur ?? 0) - transfer.totalEur) < 0.00001) score += 80;
        if (Math.abs((split.remaining ?? 0) - transfer.totalEur) < 0.00001) score += 70;
        if (split.dueDate && normalizeDate(split.dueDate) === transferDateLabel(transfer)) score += 60;
        return { split, score };
      })
      .sort((left, right) => right.score - left.score);

    const best = ranked[0];
    if (!best || best.score <= 0) continue;

    const allocation = Math.min(best.split.remaining, transfer.totalEur);
    best.split.remaining = Math.max(0, best.split.remaining - allocation);
    best.split.paid += allocation;
    best.split.transfers.push({
      ...transfer,
      allocatedAmount: allocation,
    });
    assignedTransfers.push({
      transfer,
      splitIndex: best.split.index,
      allocatedAmount: allocation,
    });
  }

  const paidSplits = remaining.filter((split) => split.remaining <= 0.00001);
  const partialSplits = remaining.filter((split) => split.paid > 0.00001 && split.remaining > 0.00001);
  const unpaidSplits = remaining.filter((split) => split.paid <= 0.00001 && split.remaining > 0.00001);
  const summaryParts = [
    paidSplits.length ? `${paidSplits.length} quota/e saldate` : "",
    partialSplits.length ? `${partialSplits.length} quota/e parziali` : "",
    unpaidSplits.length ? `${unpaidSplits.length} quota/e ancora aperte` : "",
  ].filter(Boolean);

  return {
    splits: remaining,
    paidSplits,
    partialSplits,
    unpaidSplits,
    matchedTransfers: assignedTransfers,
    summary: summaryParts.join(" | "),
  };
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

export function withholdingAmount(invoice = {}) {
  return invoice.withholdingEur
    ?? decimalFromIt(invoice.withholding)
    ?? decimalFromIt(invoice.withholdingAmount)
    ?? 0;
}

export function payableInvoiceAmount(invoice = {}) {
  const gross = invoice.totalEur ?? decimalFromIt(invoice.total);
  if (gross === null) return null;
  const withholding = withholdingAmount(invoice);
  return Math.max(0, gross - withholding);
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
  const splitPlan = allocateTransfersToInvoiceSplits(invoice, transfers);
  const splitMethods = uniqueNonEmpty(splitPlan.splits.map((split) => split.paymentType));
  if (isInstantPaidMethod(invoice.paymentTerms)) {
    return `${invoice.paymentTerms} in data ${invoice.invoiceDate || ""}`.trim();
  }
  if (splitMethods.length > 1) {
    return splitMethods.join(" + ");
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
  const splitPlan = allocateTransfersToInvoiceSplits(invoice, safeTransfers);
  const invoiceTotal = invoice.totalEur ?? decimalFromIt(invoice.total);
  const invoicePayable = payableInvoiceAmount(invoice);
  const taxableEuro = invoice.taxableEur ?? decimalFromIt(invoice.taxable);
  const vatEuro = invoice.vatEur ?? decimalFromIt(invoice.vat);
  const taxableUsd = invoice.currency === "USD"
    ? (invoice.taxableUsd ?? decimalFromIt(invoice.taxable) ?? decimalFromIt(invoice.total))
    : null;
  const instantPaid = isInstantPaidMethod(invoice.paymentTerms);
  const transfersPaidAmount = safeTransfers.reduce((sum, transfer) => sum + (transfer.totalEur ?? decimalFromIt(transfer.total) ?? 0), 0);
  const paidAmount = instantPaid ? (invoicePayable ?? invoiceTotal ?? transfersPaidAmount) : transfersPaidAmount;
  const remainingLabel = outstandingLabel(invoicePayable ?? invoiceTotal, paidAmount);
  const nextDueDate = splitPlan.unpaidSplits[0]?.dueDate
    || splitPlan.partialSplits[0]?.dueDate
    || invoice.dueDate
    || sortSplitsByDueDate(splitPlan.splits).at(-1)?.dueDate
    || addDays(invoice.invoiceDate || "", 30);
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
    Scadenza: nextDueDate,
    "BANCA - C/C": bankAccountLabel(invoice, safeTransfers),
    "Termini pagamento fattura": paymentTermsLabel(invoice, safeTransfers),
    Note: invoice.documentType === "Proforma" ? "Proforma" : "",
  });
  const missingInvoiceFields = [
    !taxableEuro && taxableEuro !== 0 ? "imponibile" : "",
    !vatEuro && vatEuro !== 0 ? "IVA" : "",
    !invoiceTotal && invoiceTotal !== 0 ? "totale" : "",
  ].filter(Boolean);
  const checks = [
    { label: "Importo", ok: String(remainingLabel).toUpperCase() === "PAGATO" },
    { label: "Pagamento registrato", ok: instantPaid || safeTransfers.length > 0 },
  ];
  const mergedMatchDebug = {
    ...(matchDebug || invoice.matchDebug || {}),
    missingInvoiceFields,
    paymentPlanSummary: splitPlan.summary,
    paymentPlanSplits: splitPlan.splits.map((split) => ({
      label: splitLabel(split),
      paid: split.paid,
      remaining: split.remaining,
    })),
  };
  return {
    id: crypto.randomUUID(),
    createdAt: new Date().toISOString(),
    row,
    invoice: { ...invoice, matchDebug: mergedMatchDebug },
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
  if (transfer.splitFromRiba && transfer.splitInvoiceNumber) {
    const expected = normalizeKey(transfer.splitInvoiceNumber);
    const actualVariants = invoiceReferenceVariants(invoice.invoiceNumber || "").map(normalizeKey);
    if (!actualVariants.includes(expected) && normalizeKey(invoice.invoiceNumber || "") !== expected) {
      return {
        score: 0,
        reasons: [],
        issues: [`Quota RIBA destinata alla fattura ${transfer.splitInvoiceNumber}`],
        summary: "Quota RIBA assegnata a un'altra fattura",
      };
    }
  }
  const splitMatch = bestInvoiceSplitMatch(invoice, transfer);
  const invoiceMethod = splitMatch?.split?.paymentType || normalizePaymentMethod(invoice.paymentTerms);
  const transferMethod = normalizePaymentMethod(transfer.paymentType || transfer.paymentTerms || transfer.method);
  const isRibaPair = invoiceMethod === "RIBA" || transferMethod === "RIBA";
  let anchors = 0;

  if (invoiceMethod && transferMethod && invoiceMethod === transferMethod) {
    score += 20;
    reasons.push(`Metodo pagamento coerente (${invoiceMethod})`);
  } else if (transferMethod) {
    issues.push(`Metodo pagamento diverso (${transferMethod})`);
  }

  const referenceMatch = reasonReferenceMatch(invoice.invoiceNumber, transfer);
  if (referenceMatch.matched) {
    score += referenceMatch.strong ? 120 : 90;
    anchors += 1;
    reasons.push(referenceMatch.strong
      ? `Riferimento documento trovato (${invoice.invoiceNumber})`
      : `Riferimento compatibile trovato (${referenceMatch.fragment})`);
  } else {
    issues.push(isRibaPair ? "Riferimento fattura non trovato nell'effetto RIBA" : "Numero fattura non trovato nella causale");
  }

  const ibanMatches = Boolean(invoice.iban && invoice.iban === transfer.beneficiaryIban);
  if (ibanMatches) {
    score += 40;
    reasons.push("IBAN fattura uguale al conto beneficiario");
  } else if (invoice.iban || transfer.beneficiaryIban) {
    issues.push("IBAN non coincidente");
  }

  const invoiceTotal = invoice.totalEur ?? decimalFromIt(invoice.total);
  const invoicePayable = payableInvoiceAmount(invoice);
  const splitAmount = splitMatch?.split?.amountEur ?? null;
  const transferTotal = transfer.totalEur ?? decimalFromIt(transfer.total);
  const amountsMatchGross = invoiceTotal !== null && transferTotal !== null && Math.abs(invoiceTotal - transferTotal) < 0.00001;
  const amountsMatchNet = invoicePayable !== null && transferTotal !== null && Math.abs(invoicePayable - transferTotal) < 0.00001;
  const amountsMatchSplit = splitAmount !== null && transferTotal !== null && Math.abs(splitAmount - transferTotal) < 0.00001;
  if (amountsMatchGross || amountsMatchNet || amountsMatchSplit) {
    score += 40;
    anchors += 1;
    reasons.push(
      amountsMatchSplit
        ? `Importo pagamento uguale a una quota del piano pagamenti (${decimalToIt(splitAmount)} €)`
        : amountsMatchNet && !amountsMatchGross
          ? `Importo pagamento uguale al netto fattura dopo ritenuta (${decimalToIt(invoicePayable)} €)`
          : "Importo bonifico uguale al totale fattura",
    );
  } else {
    issues.push(splitMatch?.split
      ? "Importo diverso da quota prevista e da totale fattura"
      : "Importo diverso");
    if (isRibaPair) score -= 25;
  }

  const dueDateBaseline = splitMatch?.split?.dueDate || invoice.dueDate;
  const dueDateMatches = Boolean(dueDateBaseline && transfer.dueDate && normalizeDate(dueDateBaseline) === normalizeDate(transfer.dueDate));
  if (dueDateMatches) {
    score += 40;
    anchors += 1;
    reasons.push(splitMatch?.split?.dueDate ? "Scadenza effetto uguale a una quota del piano pagamenti" : "Scadenza effetto uguale alla scadenza fattura");
  } else if (dueDateBaseline || transfer.dueDate) {
    issues.push("Scadenza diversa");
    if (isRibaPair) score -= 25;
  }

  const supplierVatMatches = Boolean(invoice.supplierVat && transfer.beneficiaryVat && normalizeVat(invoice.supplierVat) === normalizeVat(transfer.beneficiaryVat));
  if (supplierVatMatches) {
    score += 90;
    anchors += 1;
    reasons.push("Partita IVA fornitore uguale al creditore RIBA");
  } else if (invoice.supplierVat && transfer.beneficiaryVat) {
    issues.push("Partita IVA non coincidente");
    if (isRibaPair) score -= 35;
  }

  const supplier = normalizeKey(invoice.supplier || "");
  const beneficiary = normalizeKey(transfer.beneficiary || "");
  const directNameMatch = supplier && beneficiary && (supplier.includes(beneficiary) || beneficiary.includes(supplier));
  const overlap = overlapScore(invoice.supplier || "", transfer.beneficiary || "");
  if (directNameMatch) {
    score += 20;
    anchors += 1;
    reasons.push("Nome fornitore coerente con beneficiario");
  } else if (overlap.score > 0) {
    score += overlap.score;
    reasons.push(`Nome simile tra fornitore e beneficiario (${overlap.overlap.join(", ")})`);
  } else {
    issues.push("Nome fornitore e beneficiario poco simili");
    if (isRibaPair) score -= 30;
  }

  if (isRibaPair) {
    if (anchors === 0) {
      score = Math.min(score, 15);
      issues.push("Effetto RIBA senza segnali forti di collegamento");
    } else if (anchors === 1 && !referenceMatch.matched && !supplierVatMatches) {
      score = Math.min(score, 35);
      issues.push("Effetto RIBA con un solo segnale debole");
    }
  }

  return {
    score: Math.max(0, score),
    reasons,
    issues,
    summary: transfer.beneficiary
      ? `${transfer.beneficiary} | ${transfer.total || "-"} ${transfer.currency || "EUR"} | ${transfer.dueDate || transfer.executionDate || transfer.documentDate || ""}${transfer.noticeNumber ? ` | Avviso ${transfer.noticeNumber}` : ""}`.trim()
      : "Bonifico non identificato",
  };
}

export function pairDocuments(invoices, transfers) {
  const pairs = [];
  const used = new Set();
  const effectiveTransfers = expandRibaTransfers(invoices, transfers);
  const queue = invoices
    .map((invoice) => {
      const ranked = effectiveTransfers
        .map((transfer, index) => ({ transfer, index, match: matchScore(invoice, transfer) }))
        .sort((left, right) => right.match.score - left.match.score);
      const first = ranked[0]?.match.score || 0;
      const second = ranked[1]?.match.score || 0;
      return { invoice, ranked, first, second };
    })
    .sort((left, right) => right.first - left.first || (right.first - right.second) - (left.first - left.second));

  for (const item of queue) {
    const { invoice } = item;
    const remainingStart = payableInvoiceAmount(invoice) ?? invoice.totalEur ?? decimalFromIt(invoice.total) ?? 0;
    const ranked = effectiveTransfers
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
