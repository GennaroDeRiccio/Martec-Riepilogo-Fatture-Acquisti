import {
  buildXlsx,
  classifyDocuments,
  parseXlsxRows,
  recordsFromImportedRows,
} from "./parser.js";
import {
  MATCH_THRESHOLD,
  invoiceKeyFromRow,
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
const geminiApiKeyInput = document.querySelector("#geminiApiKeyInput");
const geminiModelInput = document.querySelector("#geminiModelInput");
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
  const dated = transfers.find((transfer) => transfer.dueDate || transfer.executionDate || transfer.documentDate);
  return dated ? (dated.dueDate || dated.executionDate || dated.documentDate) : "";
}

function displayBankCell(record) {
  const transfers = normalizeTransfers(record.transfers || record.transfer || []);
  if (!transfers.length) return "";
  const pairs = [...new Set(transfers.map((transfer) => {
    const bank = transfer.bank || "";
    const iban = transfer.payerIban || "";
    return [bank, iban].filter(Boolean).join("\n");
  }).filter(Boolean))];
  return pairs.join("\n");
}

function displayPaymentTerms(record) {
  const rowValue = record.row?.["Termini pagamento fattura"] || "";
  if (rowValue.toLowerCase().includes(" in data ")) return rowValue;
  if (/^paypal$/i.test(rowValue) || /^carta di credito$/i.test(rowValue)) {
    const invoiceDate = record.row?.["Data fattura"] || "";
    return invoiceDate ? `${rowValue} in data ${invoiceDate}` : rowValue;
  }
  if (/^riba$/i.test(rowValue)) {
    const date = displayTransferDate(record);
    return date ? `RIBA in data ${date}` : rowValue;
  }
  const date = displayTransferDate(record);
  if (date) return `Bonifico in data ${date}`;
  return rowValue || "Bonifico";
}

function cellValue(record, column) {
  if (column === "BANCA - C/C") return displayBankCell(record);
  if (column === "Termini pagamento fattura") return displayPaymentTerms(record);
  return record.row?.[column] || "";
}

function transferIdentity(transfer) {
  return [
    transfer.fileName || "",
    transfer.paymentType || "",
    transfer.noticeNumber || "",
    transfer.beneficiary || "",
    transfer.originalTotal || transfer.total || "",
    transfer.dueDate || "",
    transfer.executionDate || "",
    transfer.reason || "",
  ].join("|").toUpperCase();
}

function existingTransferKeys(recordsList = records) {
  return new Set(
    recordsList.flatMap((record) => normalizeTransfers(record.transfers || record.transfer || []).map((transfer) => transferIdentity(transfer))),
  );
}

function existingRecordsByInvoiceKey(recordsList = records) {
  return new Map(
    recordsList
      .map((record) => [invoiceKeyFromRow(record.row || {}), record])
      .filter(([key]) => Boolean(key)),
  );
}

function remainingAmountForRecord(record) {
  const value = String(record.row?.["Da pagare ancora"] || "").trim();
  if (!value || value.toUpperCase() === "PAGATO") return 0;
  return parseAmount(value);
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
  const knownTransferKeys = existingTransferKeys();
  for (const transfer of transfers) {
    if (knownTransferKeys.has(transferIdentity(transfer))) continue;
    let bestRecord = null;
    let bestMatch = null;
    for (const record of records) {
      const invoice = invoiceFromRecord(record);
      const match = matchScore(invoice, transfer);
      if (match.score < MATCH_THRESHOLD) continue;
      const transferAmount = transfer.totalEur ?? parseAmount(transfer.total);
      const remainingAmount = remainingAmountForRecord(record);
      const isRiba = String(transfer.paymentType || "").toUpperCase() === "RIBA";
      if (remainingAmount > 0 && transferAmount > remainingAmount + 0.01 && match.score < 100) continue;
      if (isRiba && transferAmount > remainingAmount + 0.01 && !String(transfer.reason || "").trim()) continue;
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
    knownTransferKeys.add(transferIdentity(nextTransfer));
    attached += 1;
    await loadRecords();
  }
  return attached;
}

async function applyAiExistingMatches(existingRecordMatches, uploadsByName) {
  let updated = 0;
  for (const entry of existingRecordMatches) {
    const record = records.find((item) => item.id === entry.recordId);
    if (!record || !entry.transfers?.length) continue;
    const currentTransfers = normalizeTransfers(record.transfers || record.transfer || []);
    const nextTransfers = [
      ...currentTransfers,
      ...entry.transfers
        .map((transfer) => ({
          ...transfer,
          storagePath: uploadsByName.get(transfer.fileName) || transfer.storagePath || "",
        }))
        .filter((transfer) => !currentTransfers.some((current) => transferIdentity(current) === transferIdentity(transfer))),
    ];
    if (nextTransfers.length === currentTransfers.length) continue;
    const rebuilt = recalculateRecord(
      invoiceFromRecord(record),
      nextTransfers,
      Number(record.row?.["Num."] || 0),
      {
        score: Math.round((entry.confidence || 0) * 100),
        threshold: 70,
        matched: true,
        reasons: [entry.rationale].filter(Boolean),
        issues: [],
        summary: nextTransfers.map((transfer) => `${transfer.beneficiary || "-"} | ${transfer.total || "-"} ${transfer.currency || "EUR"}`).join(" || "),
      },
      record.source || "upload",
    );
    await updateRecord(record.id, rebuilt.row, {
      invoiceData: rebuilt.invoice,
      transferData: { transfers: rebuilt.transfers },
      checks: rebuilt.checks,
      status: rebuilt.status,
    });
    updated += 1;
    await loadRecords();
  }
  return updated;
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
  showSetupState({
    cloudSetup,
    urlInput: supabaseUrlInput,
    keyInput: supabaseAnonKeyInput,
    geminiKeyInput: geminiApiKeyInput,
    geminiModelInput: geminiModelInput,
  }, true);
  showToast("Configura Supabase per usare il database condiviso.");
  return false;
}

async function initCloud() {
  if (!isCloudConfigured()) {
    showSetupState({
      cloudSetup,
      urlInput: supabaseUrlInput,
      keyInput: supabaseAnonKeyInput,
      geminiKeyInput: geminiApiKeyInput,
      geminiModelInput: geminiModelInput,
    }, true);
    return;
  }
  showSetupState({
    cloudSetup,
    urlInput: supabaseUrlInput,
    keyInput: supabaseAnonKeyInput,
    geminiKeyInput: geminiApiKeyInput,
    geminiModelInput: geminiModelInput,
  }, false);
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
    const { matches, transfers, existingRecordMatches, duplicateInvoices, aiUsed, aiModelUsed, aiFallbackReason } = await classifyDocuments(files, records);
    const byInvoiceKey = existingRecordsByInvoiceKey(records);
    const knownTransferKeys = existingTransferKeys(records);
    const firstIndex = nextRecordNumber(records);
    const builtRecords = [];
    let offset = 0;
    let updatedExisting = 0;
    for (const { invoice, transfers: matchedTransfers, match } of matches) {
      const nextInvoice = invoice ? { ...invoice, storagePath: uploadsByName.get(invoice.fileName) || "" } : invoice;
      const nextTransfers = matchedTransfers.map((transfer) => ({
        ...transfer,
        storagePath: uploadsByName.get(transfer.fileName) || "",
      })).filter((transfer) => !knownTransferKeys.has(transferIdentity(transfer)));
      const candidate = recalculateRecord(nextInvoice, nextTransfers, firstIndex + offset, match, "upload");
      const existingRecord = byInvoiceKey.get(candidate.invoiceKey);
      if (existingRecord) {
        const currentTransfers = normalizeTransfers(existingRecord.transfers || existingRecord.transfer || []);
        const mergedTransfers = [
          ...currentTransfers,
          ...nextTransfers.filter((transfer) => !currentTransfers.some((current) => transferIdentity(current) === transferIdentity(transfer))),
        ];
        if (mergedTransfers.length !== currentTransfers.length) {
          const rebuilt = recalculateRecord(
            nextInvoice || invoiceFromRecord(existingRecord),
            mergedTransfers,
            Number(existingRecord.row?.["Num."] || 0),
            match,
            existingRecord.source || "upload",
          );
          await updateRecord(existingRecord.id, rebuilt.row, {
            invoiceData: rebuilt.invoice,
            transferData: { transfers: rebuilt.transfers },
            checks: rebuilt.checks,
            status: rebuilt.status,
          });
          mergedTransfers.forEach((transfer) => knownTransferKeys.add(transferIdentity(transfer)));
          updatedExisting += 1;
          await loadRecords();
        }
        continue;
      }
      nextTransfers.forEach((transfer) => knownTransferKeys.add(transferIdentity(transfer)));
      builtRecords.push(candidate);
      offset += 1;
    }
    const result = await insertRecords(builtRecords);
    let aiExistingUpdated = 0;
    if (aiUsed) {
      await loadRecords();
      aiExistingUpdated = await applyAiExistingMatches(existingRecordMatches, uploadsByName);
    }
    const matchedTransferKeys = new Set([
      ...matches.flatMap((entry) => entry.transfers.map((transfer) => transferIdentity(transfer))),
      ...(existingRecordMatches || []).flatMap((entry) => entry.transfers.map((transfer) => transferIdentity(transfer))),
    ]);
    const unmatchedTransfers = transfers.filter((transfer) =>
      !matchedTransferKeys.has(transferIdentity(transfer)) && !knownTransferKeys.has(transferIdentity(transfer)));
    await loadRecords();
    const attached = aiUsed ? 0 : await attachTransfersToExistingRecords(unmatchedTransfers, uploadsByName);
    uploadForm.reset();
    updateFileLabels();
    const duplicateCount = result.duplicates.length + (duplicateInvoices?.length || 0);
    const duplicateText = duplicateCount ? `, ${duplicateCount} duplicati ignorati` : "";
    const updatedText = updatedExisting ? `, ${updatedExisting} fatture esistenti aggiornate` : "";
    const aiUpdatedText = aiExistingUpdated ? `, ${aiExistingUpdated} pagamenti associati a righe già presenti` : "";
    const attachedText = attached ? `, ${attached} pagamenti aggiunti a fatture esistenti` : "";
    const modeText = aiUsed ? `Gemini attivo (${aiModelUsed || "modello AI"})` : "matching locale";
    const fallbackText = !aiUsed && aiFallbackReason ? ` - ${aiFallbackReason}, uso fallback locale` : "";
    showToast(`${result.added.length} righe aggiunte${duplicateText}${updatedText}${aiUpdatedText}${attachedText} (${modeText})${fallbackText}`);
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
  const geminiKey = geminiApiKeyInput?.value.trim() || "";
  const geminiModel = geminiModelInput?.value.trim() || "";
  if (!url || !key) {
    showToast("Inserisci URL e anon key.");
    return;
  }
  saveCloudConfig({ supabaseUrl: url, supabaseAnonKey: key, geminiApiKey: geminiKey, geminiModel });
  resetCloudClient();
  await initCloud();
  showToast("Configurazione cloud e AI salvata");
});

for (const element of [documentsInput, excelInput]) element.addEventListener("change", updateFileLabels);
for (const element of [searchInput, statusFilter, invoiceDateFrom, invoiceDateTo, dueDateFrom, dueDateTo]) {
  element.addEventListener("input", renderTable);
}

showSetupState({
  cloudSetup,
  urlInput: supabaseUrlInput,
  keyInput: supabaseAnonKeyInput,
  geminiKeyInput: geminiApiKeyInput,
  geminiModelInput: geminiModelInput,
}, !isCloudConfigured());
updateFileLabels();
initCloud().catch((error) => showToast(error.message || "Connessione cloud non riuscita"));
