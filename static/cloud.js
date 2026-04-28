import { createClient } from "https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2/+esm";
import {
  EXCEL_COLUMNS,
  STATUS_OK,
  STATUS_VERIFY,
  checksFromRow,
  invoiceKeyFromRow,
  normalizeRow,
  normalizeStatus,
  normalizeTransfers,
} from "./domain.js";

const CONFIG_KEY = "martec-cloud-config";
const RECORDS_TABLE = "records";
const SUPPLIERS_TABLE = "suppliers";

let client = null;
let currentConfigKey = "";

function normalizeCloudError(error) {
  const message = String(error?.message || error || "").trim();
  if (error instanceof TypeError || /Failed to fetch|Load failed|NetworkError/i.test(message)) {
    return new Error("Connessione cloud non riuscita. Controlla internet, URL/chiave Supabase e apri la web app da http/https.");
  }
  return error instanceof Error ? error : new Error(message || "Errore cloud sconosciuto.");
}

async function withCloudError(task) {
  try {
    return await task();
  } catch (error) {
    throw normalizeCloudError(error);
  }
}

function configFromWindow() {
  return window.APP_CONFIG || {};
}

export function getCloudConfig() {
  const stored = JSON.parse(localStorage.getItem(CONFIG_KEY) || "{}");
  const config = { ...configFromWindow(), ...stored };
  return {
    supabaseUrl: String(config.supabaseUrl || "").trim(),
    supabaseAnonKey: String(config.supabaseAnonKey || "").trim(),
    storageBucket: String(config.storageBucket || "documents").trim(),
    realtimeChannel: String(config.realtimeChannel || "martec-records").trim(),
    geminiModel: String(config.geminiModel || "gemini-2.5-flash").trim(),
  };
}

export function isCloudConfigured() {
  const config = getCloudConfig();
  return Boolean(config.supabaseUrl && config.supabaseAnonKey);
}

export function saveCloudConfig(config) {
  localStorage.setItem(CONFIG_KEY, JSON.stringify({
    supabaseUrl: String(config.supabaseUrl || "").trim(),
    supabaseAnonKey: String(config.supabaseAnonKey || "").trim(),
    geminiModel: String(config.geminiModel || "gemini-2.5-flash").trim(),
  }));
}

export function resetCloudClient() {
  client = null;
  currentConfigKey = "";
}

export function getSupabase() {
  const config = getCloudConfig();
  if (!config.supabaseUrl || !config.supabaseAnonKey) throw new Error("Configura prima Supabase.");
  const key = `${config.supabaseUrl}|${config.supabaseAnonKey}`;
  if (client && currentConfigKey === key) return client;
  client = createClient(config.supabaseUrl, config.supabaseAnonKey, {
    auth: { persistSession: false, autoRefreshToken: false },
  });
  currentConfigKey = key;
  return client;
}

function mapRecord(row) {
  const transfers = normalizeTransfers(row.transfer_data || {});
  return {
    id: row.id,
    createdAt: row.created_at,
    row: normalizeRow(row.row_data || {}),
    invoice: row.invoice_data || {},
    transfer: transfers[0] || {},
    transfers,
    checks: row.checks_data || [],
    source: row.source || "upload",
    status: normalizeStatus(row.status),
    matchDebug: row.invoice_data?.matchDebug || null,
  };
}

function mapSupplier(row) {
  return {
    id: row.id,
    name: row.name || "",
    vat: row.vat || "",
    iban: row.iban || "",
    swift: row.swift || "",
    bank: row.bank || "",
    currency: row.currency || "",
    notes: row.notes || "",
    updated_at: row.updated_at || "",
  };
}

export async function fetchRecords() {
  return withCloudError(async () => {
    const { data, error } = await getSupabase().from(RECORDS_TABLE).select("*").order("created_at", { ascending: true });
    if (error) throw error;
    return (data || []).map(mapRecord).sort((a, b) => Number(a.row["Num."] || 0) - Number(b.row["Num."] || 0));
  });
}

export async function fetchSuppliers() {
  return withCloudError(async () => {
    const { data, error } = await getSupabase().from(SUPPLIERS_TABLE).select("*").order("name", { ascending: true });
    if (error) throw error;
    return (data || []).map(mapSupplier);
  });
}

export function nextRecordNumber(records) {
  return records.reduce((max, record) => Math.max(max, Number(record.row?.["Num."] || 0)), 0) + 1;
}

async function upsertSupplierFromRecord(record) {
  const row = normalizeRow(record.row || {});
  const invoice = record.invoice || {};
  const transfer = normalizeTransfers(record.transfers || record.transfer || [])[0] || {};
  const name = row.Fornitore || row.Cliente || invoice.supplier || transfer.beneficiary;
  if (!name) return;
  const payload = {
    name,
    vat: invoice.supplierVat || "",
    iban: invoice.iban || transfer.beneficiaryIban || "",
    swift: transfer.swift || "",
    bank: row["BANCA - C/C"] || "",
    currency: row["Valore in USD"] ? "USD" : "EUR",
    notes: "",
    updated_at: new Date().toISOString(),
  };
  const { error } = await getSupabase().from(SUPPLIERS_TABLE).upsert(payload, { onConflict: "name" });
  if (error) throw error;
}

export async function insertRecords(records) {
  return withCloudError(async () => {
    const added = [];
    const duplicates = [];
    for (const record of records) {
      const normalizedRow = normalizeRow(record.row || {});
      const payload = {
        id: record.id,
        created_at: record.createdAt || new Date().toISOString(),
        row_data: normalizedRow,
        invoice_data: record.invoice || {},
        transfer_data: { transfers: normalizeTransfers(record.transfers || record.transfer || []) },
        checks_data: record.checks || checksFromRow(normalizedRow),
        source: record.source || "upload",
        invoice_key: invoiceKeyFromRow(normalizedRow),
        status: normalizeStatus(record.status || (String(normalizedRow["Da pagare ancora"] || "").trim().toUpperCase() === "PAGATO" ? STATUS_OK : STATUS_VERIFY)),
      };
      const { error } = await getSupabase().from(RECORDS_TABLE).insert(payload);
      if (error) {
        if (error.code === "23505") {
          duplicates.push({ invoice: normalizedRow.Fattura, supplier: normalizedRow.Fornitore || normalizedRow.Cliente });
          continue;
        }
        throw error;
      }
      await upsertSupplierFromRecord({ ...record, row: normalizedRow });
      added.push(payload);
    }
    return { added, duplicates };
  });
}

export async function updateRecord(recordId, nextRow, options = {}) {
  return withCloudError(async () => {
    const normalizedRow = normalizeRow(nextRow);
    const { data: existing, error: existingError } = await getSupabase()
      .from(RECORDS_TABLE)
      .select("*")
      .eq("id", recordId)
      .single();
    if (existingError) throw existingError;
    const payload = {
      row_data: normalizedRow,
      invoice_data: options.invoiceData || existing.invoice_data || {},
      transfer_data: options.transferData || existing.transfer_data || { transfers: normalizeTransfers(existing.transfer_data || {}) },
      checks_data: options.checks || checksFromRow(normalizedRow),
      status: normalizeStatus(options.status || nextRow.Stato || (String(normalizedRow["Da pagare ancora"] || "").trim().toUpperCase() === "PAGATO" ? STATUS_OK : STATUS_VERIFY)),
      invoice_key: invoiceKeyFromRow(normalizedRow),
    };
    const { error } = await getSupabase().from(RECORDS_TABLE).update(payload).eq("id", recordId);
    if (error) throw error;
    await upsertSupplierFromRecord({
      row: normalizedRow,
      invoice: existing.invoice_data || {},
      transfers: normalizeTransfers(options.transferData || existing.transfer_data || {}),
    });
  });
}

export async function deleteRecord(recordId) {
  return withCloudError(async () => {
    const { error } = await getSupabase().from(RECORDS_TABLE).delete().eq("id", recordId);
    if (error) throw error;
  });
}

export async function replaceImportedRows(records) {
  return insertRecords(records);
}

export async function uploadFilesToStorage(files) {
  return withCloudError(async () => {
    const config = getCloudConfig();
    const batchId = crypto.randomUUID();
    const uploads = [];
    for (const file of files) {
      const path = `${batchId}/${Date.now()}-${file.name.replace(/[^\w.-]+/g, "_")}`;
      const { error } = await getSupabase().storage.from(config.storageBucket).upload(path, file, { upsert: false });
      if (error) throw error;
      uploads.push({ fileName: file.name, path });
    }
    return uploads;
  });
}

export function subscribeToChanges(onChange) {
  const config = getCloudConfig();
  const channel = getSupabase()
    .channel(`${config.realtimeChannel}-${Date.now()}`)
    .on("postgres_changes", { event: "*", schema: "public", table: RECORDS_TABLE }, () => onChange("records"))
    .on("postgres_changes", { event: "*", schema: "public", table: SUPPLIERS_TABLE }, () => onChange("suppliers"))
    .subscribe();
  return () => getSupabase().removeChannel(channel);
}

export function showSetupState(elements, visible) {
  elements.cloudSetup?.classList.toggle("hidden", !visible);
  elements.urlInput && (elements.urlInput.value = getCloudConfig().supabaseUrl);
  elements.keyInput && (elements.keyInput.value = getCloudConfig().supabaseAnonKey);
  elements.geminiModelInput && (elements.geminiModelInput.value = getCloudConfig().geminiModel);
}

export function getColumns() {
  return [...EXCEL_COLUMNS];
}
