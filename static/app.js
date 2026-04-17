import {
  buildXlsx,
  classifyDocuments,
  parseXlsxRows,
  recordsFromImportedRows,
} from "./parser.js";
import {
  MATCH_THRESHOLD,
  matchScore,
  normalizeTransfers,
  recalculateRecord,
} from "./domain.js";
import {
  deleteRecord,
  fetchRecords,
  getColumns,
  insertRecords,
  isCloudConfigured,
  nextRecordNumber,
  replaceImportedRows,
  resetCloudClient,
  saveCloudConfig,
  showSetupState,
  subscribeToChanges,
  updateRecord,
  uploadFilesToStorage,
} from "./cloud.js";

const table = document.querySelector("#recordsTable");
const uploadForm = document.querySelector("#uploadForm");
const importForm = document.querySelector("#importForm");
const documentsInput = document.querySelector("#documentsInput");
const excelInput = document.querySelector("#excelInput");
const documentsName = document.querySelector("#documentsName");
const excelName = document.querySelector("#excelName");
const toast = document.querySelector("#toast");
const searchInput = document.querySelector("#searchInput");
const statusFilter = document.querySelector("#statusFilter");
const invoiceDateFrom = document.querySelector("#invoiceDateFrom");
const invoiceDateTo = document.querySelector("#invoiceDateTo");
const dueDateFrom = document.querySelector("#dueDateFrom");
const dueDateTo = document.querySelector("#dueDateTo");
const exportExcelButton = document.querySelector("#exportExcelButton");
const cloudSetup = document.querySelector("#cloudSetup");
const supabaseUrlInput = document.querySelector("#supabaseUrlInput");
const supabaseAnonKeyInput = document.querySelector("#supabaseAnonKeyInput");
const saveCloudConfigButton = document.querySelector("#saveCloudConfigButton");

let columns = getColumns();
let records = [];
let unsubscribe = null;

const compactColumns = new Set([
  "Num.",
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
]);

const numericColumns = new Set([
  "Valore in USD",
  "Valore in Euro",
  "Imponibile",
  "IVA",
  "Totale",
  "Uscite",
]);

const multilineColumns = new Set(["Fornitore", "BANCA - C/C", "Termini pagamento fattura", "Note", "Match"]);

function showToast(message) {
  toast.textContent = message;
  toast.classList.add("visible");
  clearTimeout(showToast.timeout);
  showToast.timeout = setTimeout(() => toast.classList.remove("visible"), 3600);
}

function parseAmount(value) {
  if (!value) return 0;
  const cleaned = String(value).replace(/EUR|USD|€|\$/g, "").replace(/\./g, "").replace(",", ".").trim();
  const parsed = Number.parseFloat(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
}

function formatAmount(value) {
  return value.toLocaleString("it-IT", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function parseDate(value) {
  const match = String(value || "").match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  return match ? `${match[3]}-${match[2]}-${match[1]}` : null;
}

function updateFileLabels() {
  const docs = [...documentsInput.files].map((file) => file.name);
  documentsName.textContent = docs.length ? `${docs.length} file selezionati` : "Seleziona uno o più file";
  excelName.textContent = excelInput.files[0]?.name || "Seleziona .xlsx";
}

function isOk(record) {
  return (record.status || "").toLowerCase() === "pagato";
}

function invoiceFromRecord(record) {
  if (record.invoice && Object.keys(record.invoice).length) return record.invoice;
  return {
    supplier: record.row?.Fornitore || record.row?.Cliente || "",
    invoiceNumber: record.row?.Fattura || "",
    invoiceDate: record.row?.["Data fattura"] || record.row?.Data || "",
    taxable: record.row?.Imponibile || record.row?.["Valore in Euro"] || "",
    taxableEur: parseAmount(record.row?.Imponibile || record.row?.["Valore in Euro"]),
    vat: record.row?.IVA || "",
    vatEur: parseAmount(record.row?.IVA),
    total: record.row?.Totale || "",
    totalEur: parseAmount(record.row?.Totale),
    dueDate: record.row?.Scadenza || "",
    paymentTerms: record.row?.["Termini pagamento fattura"] || "",
    currency: record.row?.["Valore in USD"] ? "USD" : "EUR",
    exchangeRate: record.invoice?.exchangeRate || null,
  };
}

function displayTransferDate(record) {
  const transfers = normalizeTransfers(record.transfers || record.transfer || []);
  const dated = transfers.find((transfer) => transfer.executionDate || transfer.documentDate);
  return dated ? (dated.executionDate || dated.documentDate) : "";
}

function displayBankCell(record) {
  const transfers = normalizeTransfers(record.transfers || record.transfer || []);
  if (!transfers.length) return "";
  const pairs = [...new Set(transfers.map((transfer) => {
    const bank = transfer.bank || "";
    const iban = transfer.beneficiaryIban || "";
    return [bank, iban].filter(Boolean).join("\n");
  }).filter(Boolean))];
  return pairs.join("\n");
}

function displayPaymentTerms(record) {
  const rowValue = record.row?.["Termini pagamento fattura"] || "";
  if (rowValue.toLowerCase().includes(" in data ")) return rowValue;
  const date = displayTransferDate(record);
  if (date) return `Bonifico in data ${date}`;
  return rowValue || "Bonifico";
}

function cellValue(record, column) {
  if (column === "BANCA - C/C") return displayBankCell(record);
  if (column === "Termini pagamento fattura") return displayPaymentTerms(record);
  return record.row?.[column] || "";
}

function filteredRecords() {
  const query = searchInput.value.trim().toLowerCase();
  const status = statusFilter.value;
  const invoiceStart = invoiceDateFrom.value;
  const invoiceEnd = invoiceDateTo.value;
  const dueStart = dueDateFrom.value;
  const dueEnd = dueDateTo.value;
  return records.filter((record) => {
    const haystack = columns.map((column) => cellValue(record, column)).join(" ").toLowerCase();
    const invoiceDate = parseDate(record.row?.["Data fattura"]);
    const dueDate = parseDate(record.row?.Scadenza);
    if (query && !haystack.includes(query)) return false;
    if (status === "ok" && !isOk(record)) return false;
    if (status === "verify" && isOk(record)) return false;
    if (invoiceStart && invoiceDate && invoiceDate < invoiceStart) return false;
    if (invoiceEnd && invoiceDate && invoiceDate > invoiceEnd) return false;
    if (dueStart && dueDate && dueDate < dueStart) return false;
    if (dueEnd && dueDate && dueDate > dueEnd) return false;
    return true;
  });
}

function makeStatusSelect(record) {
  const select = document.createElement("select");
  const current = record.status || (isOk(record) ? "Pagato" : "Da pagare");
  select.className = `cell-select status-select ${current === "Pagato" ? "paid" : "unpaid"}`;
  for (const optionValue of ["Pagato", "Da pagare"]) {
    const option = document.createElement("option");
    option.value = optionValue;
    option.textContent = optionValue;
    option.selected = optionValue === current;
    select.appendChild(option);
  }
  select.addEventListener("change", () => saveStatus(record.id, select.value));
  return select;
}

function updateMetrics() {
  const visible = filteredRecords();
  const invoiceTotal = visible.reduce((sum, record) => sum + parseAmount(record.row?.Totale), 0);
  const transferTotal = visible.reduce((sum, record) => sum + parseAmount(record.row?.Uscite), 0);
  const issues = visible.filter((record) => !isOk(record)).length;
  document.querySelector("#rowsCount").textContent = visible.length;
  document.querySelector("#invoiceTotal").textContent = `€ ${formatAmount(invoiceTotal)}`;
  document.querySelector("#transferTotal").textContent = `€ ${formatAmount(transferTotal)}`;
  document.querySelector("#issuesCount").textContent = issues;
}

function matchTooltip(record) {
  const debug = record.matchDebug;
  if (!debug) return "Nessun dettaglio match disponibile";
  const lines = [];
  lines.push(`Bonifico proposto: ${debug.summary || "Nessuno"}`);
  lines.push(`Punteggio match: ${debug.score ?? 0}/${debug.threshold ?? 0}`);
  if (debug.reasons?.length) lines.push(`Segnali positivi: ${debug.reasons.join(" | ")}`);
  if (debug.issues?.length) lines.push(`Criticita: ${debug.issues.join(" | ")}`);
  return lines.join("\n");
}

function makeMatchBadge(record) {
  const badge = document.createElement("span");
  const debug = record.matchDebug;
  if (!debug) {
    badge.className = "badge warn";
    badge.textContent = "Nessun debug";
    return badge;
  }
  badge.className = `badge ${debug.matched ? "good" : "warn"} match-badge`;
  badge.textContent = debug.matched ? `${debug.score} pt` : `No match (${debug.score})`;
  badge.title = matchTooltip(record);
  return badge;
}

function renderTable() {
  const thead = table.querySelector("thead");
  const tbody = table.querySelector("tbody");
  thead.innerHTML = "";
  tbody.innerHTML = "";

  const headerRow = document.createElement("tr");
  for (const column of [...columns, "Stato", "Match", "Azioni"]) {
    const th = document.createElement("th");
    th.textContent = column;
    if (compactColumns.has(column)) th.classList.add("is-compact");
    if (numericColumns.has(column)) th.classList.add("is-numeric");
    if (column === "Stato") th.classList.add("is-status");
    if (column === "Match") th.classList.add("is-match");
    if (column === "Azioni") th.classList.add("is-actions");
    headerRow.appendChild(th);
  }
  thead.appendChild(headerRow);

  const visible = filteredRecords();
  if (!visible.length) {
    const row = document.createElement("tr");
    row.className = "empty-row";
    const cell = document.createElement("td");
    cell.colSpan = columns.length + 3;
    cell.textContent = records.length ? "Nessuna riga corrisponde ai filtri" : "Nessuna fattura caricata";
    row.appendChild(cell);
    tbody.appendChild(row);
    updateMetrics();
    return;
  }

  for (const record of visible) {
    const row = document.createElement("tr");
    row.dataset.id = record.id;
    for (const column of columns) {
      const td = document.createElement("td");
      td.dataset.column = column;
      if (compactColumns.has(column)) td.classList.add("is-compact");
      if (numericColumns.has(column)) td.classList.add("is-numeric");
      if (multilineColumns.has(column)) td.classList.add("is-multiline");
      td.textContent = cellValue(record, column);
      td.contentEditable = "true";
      td.title = "Modifica e poi esci dalla cella per salvare";
      td.addEventListener("blur", () => saveCell(record.id, column, td.textContent));
      row.appendChild(td);
    }
    const status = document.createElement("td");
    status.className = "is-status";
    status.appendChild(makeStatusSelect(record));
    status.title = (record.checks || []).map((check) => `${check.label}: ${check.ok ? "OK" : "verifica"}`).join("\n");
    row.appendChild(status);

    const match = document.createElement("td");
    match.className = "is-match";
    match.appendChild(makeMatchBadge(record));
    const matchMeta = document.createElement("div");
    matchMeta.className = "match-meta";
    matchMeta.textContent = record.matchDebug?.summary || "Nessun bonifico proposto";
    match.title = matchTooltip(record);
    match.appendChild(matchMeta);
    row.appendChild(match);

    const actions = document.createElement("td");
    actions.className = "is-actions";
    const deleteButton = document.createElement("button");
    deleteButton.type = "button";
    deleteButton.className = "row-delete";
    deleteButton.textContent = "Elimina";
    deleteButton.addEventListener("click", () => removeRecord(record));
    actions.appendChild(deleteButton);
    row.appendChild(actions);
    tbody.appendChild(row);
  }
  updateMetrics();
}

async function loadRecords() {
  records = await fetchRecords();
  renderTable();
}

async function saveCell(recordId, column, value) {
  const record = records.find((item) => item.id === recordId);
  if (!record || (record.row?.[column] || "") === value) return;
  const nextRow = { ...record.row, [column]: value };
  await updateRecord(recordId, nextRow);
  await loadRecords();
  showToast("Modifica salvata");
}

async function saveStatus(recordId, status) {
  const record = records.find((item) => item.id === recordId);
  if (!record || record.status === status) return;
  const nextRow = { ...record.row };
  await updateRecord(recordId, nextRow, { status });
  await loadRecords();
  showToast("Stato aggiornato");
}

async function attachTransfersToExistingRecords(transfers, uploadsByName) {
  let attached = 0;
  for (const transfer of transfers) {
    let bestRecord = null;
    let bestMatch = null;
    for (const record of records) {
      const invoice = invoiceFromRecord(record);
      const match = matchScore(invoice, transfer);
      if (match.score < MATCH_THRESHOLD) continue;
      if (!bestMatch || match.score > bestMatch.score) {
        bestMatch = match;
        bestRecord = record;
      }
    }
    if (!bestRecord || !bestMatch) continue;
    const nextTransfer = {
      ...transfer,
      storagePath: uploadsByName.get(transfer.fileName) || transfer.storagePath || "",
    };
    const currentTransfers = normalizeTransfers(bestRecord.transfers || bestRecord.transfer || []);
    const rebuilt = recalculateRecord(
      invoiceFromRecord(bestRecord),
      [...currentTransfers, nextTransfer],
      Number(bestRecord.row?.["Num."] || 0),
      bestMatch,
      bestRecord.source || "upload",
    );
    await updateRecord(bestRecord.id, rebuilt.row, {
      invoiceData: rebuilt.invoice,
      transferData: { transfers: rebuilt.transfers },
      checks: rebuilt.checks,
      status: rebuilt.status,
    });
    attached += 1;
    await loadRecords();
  }
  return attached;
}

async function removeRecord(record) {
  const label = [record.row?.Fornitore || record.row?.Cliente, record.row?.Fattura].filter(Boolean).join(" - ") || "questa riga";
  if (!window.confirm(`Vuoi eliminare ${label}?`)) return;
  try {
    await deleteRecord(record.id);
    await loadRecords();
    showToast("Riga eliminata");
  } catch (error) {
    showToast(error.message || "Eliminazione non riuscita");
  }
}

async function ensureCloudReady() {
  if (isCloudConfigured()) return true;
  showSetupState({ cloudSetup, urlInput: supabaseUrlInput, keyInput: supabaseAnonKeyInput }, true);
  showToast("Configura Supabase per usare il database condiviso.");
  return false;
}

async function initCloud() {
  if (!isCloudConfigured()) {
    showSetupState({ cloudSetup, urlInput: supabaseUrlInput, keyInput: supabaseAnonKeyInput }, true);
    return;
  }
  showSetupState({ cloudSetup, urlInput: supabaseUrlInput, keyInput: supabaseAnonKeyInput }, false);
  if (unsubscribe) unsubscribe();
  await loadRecords();
  unsubscribe = subscribeToChanges(async (tableName) => {
    if (tableName === "records") await loadRecords();
  });
}

uploadForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!(await ensureCloudReady())) return;
  const files = [...documentsInput.files];
  if (!files.length) {
    showToast("Seleziona almeno un PDF.");
    return;
  }
  const submitButton = uploadForm.querySelector("button[type='submit']");
  submitButton.disabled = true;
  submitButton.textContent = "Estrazione...";
  try {
    const uploads = await uploadFilesToStorage(files);
    const uploadsByName = new Map(uploads.map((upload) => [upload.fileName, upload.path]));
    const { matches, transfers } = await classifyDocuments(files);
    const firstIndex = nextRecordNumber(records);
    const builtRecords = matches.map(({ invoice, transfers: matchedTransfers, match }, offset) => {
      const nextInvoice = invoice ? { ...invoice, storagePath: uploadsByName.get(invoice.fileName) || "" } : invoice;
      const nextTransfers = matchedTransfers.map((transfer) => ({
        ...transfer,
        storagePath: uploadsByName.get(transfer.fileName) || "",
      }));
      return recalculateRecord(nextInvoice, nextTransfers, firstIndex + offset, match, "upload");
    });
    const result = await insertRecords(builtRecords);
    const matchedTransferNames = new Set(matches.flatMap((entry) => entry.transfers.map((transfer) => transfer.fileName)));
    const unmatchedTransfers = transfers.filter((transfer) => !matchedTransferNames.has(transfer.fileName));
    await loadRecords();
    const attached = await attachTransfersToExistingRecords(unmatchedTransfers, uploadsByName);
    uploadForm.reset();
    updateFileLabels();
    const duplicateText = result.duplicates.length ? `, ${result.duplicates.length} duplicati ignorati` : "";
    const attachedText = attached ? `, ${attached} bonifici aggiunti a fatture esistenti` : "";
    showToast(`${result.added.length} righe aggiunte${duplicateText}${attachedText}`);
  } catch (error) {
    showToast(error.message || "Upload non riuscito");
  } finally {
    submitButton.disabled = false;
    submitButton.textContent = "Estrai e abbina";
  }
});

importForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!(await ensureCloudReady())) return;
  if (!excelInput.files[0]) {
    showToast("Seleziona un file Excel");
    return;
  }
  const submitButton = importForm.querySelector("button[type='submit']");
  submitButton.disabled = true;
  submitButton.textContent = "Import...";
  try {
    const rows = await parseXlsxRows(excelInput.files[0]);
    const importedRecords = recordsFromImportedRows(rows, nextRecordNumber(records));
    const result = await replaceImportedRows(importedRecords);
    await loadRecords();
    importForm.reset();
    updateFileLabels();
    showToast(`${result.added.length} righe importate, ${result.duplicates.length} duplicati ignorati`);
  } catch (error) {
    showToast(error.message || "Import non riuscito");
  } finally {
    submitButton.disabled = false;
    submitButton.textContent = "Importa";
  }
});

exportExcelButton.addEventListener("click", async (event) => {
  event.preventDefault();
  if (!(await ensureCloudReady())) return;
  try {
    const blob = await buildXlsx(records);
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "riepilogo-fatture-martec.xlsx";
    link.click();
    URL.revokeObjectURL(url);
  } catch (error) {
    showToast(error.message || "Export non riuscito");
  }
});

saveCloudConfigButton.addEventListener("click", async () => {
  const url = supabaseUrlInput.value.trim();
  const key = supabaseAnonKeyInput.value.trim();
  if (!url || !key) {
    showToast("Inserisci URL e anon key.");
    return;
  }
  saveCloudConfig({ supabaseUrl: url, supabaseAnonKey: key });
  resetCloudClient();
  await initCloud();
  showToast("Connessione cloud salvata");
});

for (const element of [documentsInput, excelInput]) element.addEventListener("change", updateFileLabels);
for (const element of [searchInput, statusFilter, invoiceDateFrom, invoiceDateTo, dueDateFrom, dueDateTo]) {
  element.addEventListener("input", renderTable);
}

showSetupState({ cloudSetup, urlInput: supabaseUrlInput, keyInput: supabaseAnonKeyInput }, !isCloudConfigured());
updateFileLabels();
initCloud().catch((error) => showToast(error.message || "Connessione cloud non riuscita"));
