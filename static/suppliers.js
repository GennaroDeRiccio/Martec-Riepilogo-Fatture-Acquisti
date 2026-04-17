import {
  fetchSuppliers,
  isCloudConfigured,
  resetCloudClient,
  saveCloudConfig,
  showSetupState,
  subscribeToChanges,
} from "./cloud.js";

const suppliersTable = document.querySelector("#suppliersTable");
const supplierSearchInput = document.querySelector("#supplierSearchInput");
const suppliersCount = document.querySelector("#suppliersCount");
const toast = document.querySelector("#toast");
const cloudSetup = document.querySelector("#cloudSetup");
const supabaseUrlInput = document.querySelector("#supabaseUrlInput");
const supabaseAnonKeyInput = document.querySelector("#supabaseAnonKeyInput");
const saveCloudConfigButton = document.querySelector("#saveCloudConfigButton");

let suppliers = [];
let unsubscribe = null;

function showToast(message) {
  toast.textContent = message;
  toast.classList.add("visible");
  clearTimeout(showToast.timeout);
  showToast.timeout = setTimeout(() => toast.classList.remove("visible"), 3600);
}

function filteredSuppliers() {
  const query = supplierSearchInput.value.trim().toLowerCase();
  if (!query) return suppliers;
  return suppliers.filter((supplier) =>
    ["name", "vat", "iban", "swift", "bank", "currency"].map((key) => supplier[key] || "").join(" ").toLowerCase().includes(query),
  );
}

function renderSuppliers() {
  const tbody = suppliersTable.querySelector("tbody");
  const visible = filteredSuppliers();
  tbody.innerHTML = "";
  suppliersCount.textContent = visible.length;
  if (!visible.length) {
    const row = document.createElement("tr");
    row.className = "empty-row";
    const cell = document.createElement("td");
    cell.colSpan = 6;
    cell.textContent = suppliers.length ? "Nessun fornitore corrisponde alla ricerca" : "Archivio fornitori vuoto";
    row.appendChild(cell);
    tbody.appendChild(row);
    return;
  }
  visible.forEach((supplier) => {
    const row = document.createElement("tr");
    ["name", "vat", "iban", "swift", "bank", "currency"].forEach((key) => {
      const td = document.createElement("td");
      td.textContent = supplier[key] || "";
      row.appendChild(td);
    });
    tbody.appendChild(row);
  });
}

async function loadSuppliers() {
  suppliers = await fetchSuppliers();
  renderSuppliers();
}

async function initCloud() {
  if (!isCloudConfigured()) {
    showSetupState({ cloudSetup, urlInput: supabaseUrlInput, keyInput: supabaseAnonKeyInput }, true);
    return;
  }
  showSetupState({ cloudSetup, urlInput: supabaseUrlInput, keyInput: supabaseAnonKeyInput }, false);
  if (unsubscribe) unsubscribe();
  await loadSuppliers();
  unsubscribe = subscribeToChanges(async (tableName) => {
    if (tableName === "suppliers") await loadSuppliers();
  });
}

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

supplierSearchInput.addEventListener("input", renderSuppliers);
showSetupState({ cloudSetup, urlInput: supabaseUrlInput, keyInput: supabaseAnonKeyInput }, !isCloudConfigured());
initCloud().catch((error) => showToast(error.message || "Connessione cloud non riuscita"));
