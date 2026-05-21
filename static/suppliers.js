import {
  deleteSupplier,
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
const selectAllSuppliers = document.querySelector("#selectAllSuppliers");
const deleteSelectedSuppliersButton = document.querySelector("#deleteSelectedSuppliersButton");
const toast = document.querySelector("#toast");
const cloudSetup = document.querySelector("#cloudSetup");
const supabaseUrlInput = document.querySelector("#supabaseUrlInput");
const supabaseAnonKeyInput = document.querySelector("#supabaseAnonKeyInput");
const saveCloudConfigButton = document.querySelector("#saveCloudConfigButton");

let suppliers = [];
let unsubscribe = null;
const selectedSupplierIds = new Set();

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
  const visibleIds = new Set(visible.map((supplier) => supplier.id));
  [...selectedSupplierIds].forEach((id) => {
    if (!visibleIds.has(id) && !suppliers.some((supplier) => supplier.id === id)) selectedSupplierIds.delete(id);
  });
  tbody.innerHTML = "";
  suppliersCount.textContent = visible.length;
  updateSupplierSelectionUi(visible);
  if (!visible.length) {
    const row = document.createElement("tr");
    row.className = "empty-row";
    const cell = document.createElement("td");
    cell.colSpan = 8;
    cell.textContent = suppliers.length ? "Nessun fornitore corrisponde alla ricerca" : "Archivio fornitori vuoto";
    row.appendChild(cell);
    tbody.appendChild(row);
    return;
  }
  visible.forEach((supplier) => {
    const row = document.createElement("tr");
    const selection = document.createElement("td");
    selection.className = "is-selection";
    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.checked = selectedSupplierIds.has(supplier.id);
    checkbox.setAttribute("aria-label", `Seleziona ${supplier.name || "fornitore"}`);
    checkbox.addEventListener("change", () => {
      if (checkbox.checked) selectedSupplierIds.add(supplier.id);
      else selectedSupplierIds.delete(supplier.id);
      renderSuppliers();
    });
    selection.appendChild(checkbox);
    row.appendChild(selection);
    ["name", "vat", "iban", "swift", "bank", "currency"].forEach((key) => {
      const td = document.createElement("td");
      td.textContent = supplier[key] || "";
      row.appendChild(td);
    });
    const actions = document.createElement("td");
    const button = document.createElement("button");
    button.className = "table-action";
    button.type = "button";
    button.textContent = "Elimina";
    button.addEventListener("click", () => removeSupplier(supplier));
    actions.appendChild(button);
    row.appendChild(actions);
    tbody.appendChild(row);
  });
}

function updateSupplierSelectionUi(visible = filteredSuppliers()) {
  const selectedVisible = visible.filter((supplier) => selectedSupplierIds.has(supplier.id));
  if (selectAllSuppliers) {
    selectAllSuppliers.checked = visible.length > 0 && selectedVisible.length === visible.length;
    selectAllSuppliers.indeterminate = selectedVisible.length > 0 && selectedVisible.length < visible.length;
  }
  if (deleteSelectedSuppliersButton) deleteSelectedSuppliersButton.disabled = selectedSupplierIds.size === 0;
}

async function removeSupplier(supplier) {
  if (!window.confirm(`Vuoi eliminare il fornitore ${supplier.name || "selezionato"}?`)) return;
  try {
    await deleteSupplier(supplier.id);
    await loadSuppliers();
    showToast("Fornitore eliminato");
  } catch (error) {
    showToast(error.message || "Eliminazione fornitore non riuscita");
  }
}

async function removeSelectedSuppliers() {
  const selected = suppliers.filter((supplier) => selectedSupplierIds.has(supplier.id));
  if (!selected.length) return;
  if (!window.confirm(`Vuoi eliminare ${selected.length} fornitori selezionati?`)) return;
  try {
    for (const supplier of selected) {
      await deleteSupplier(supplier.id);
    }
    selectedSupplierIds.clear();
    await loadSuppliers();
    showToast(`${selected.length} fornitori eliminati`);
  } catch (error) {
    showToast(error.message || "Eliminazione fornitori non riuscita");
  }
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

selectAllSuppliers?.addEventListener("change", () => {
  filteredSuppliers().forEach((supplier) => {
    if (selectAllSuppliers.checked) selectedSupplierIds.add(supplier.id);
    else selectedSupplierIds.delete(supplier.id);
  });
  renderSuppliers();
});
deleteSelectedSuppliersButton?.addEventListener("click", removeSelectedSuppliers);
supplierSearchInput.addEventListener("input", renderSuppliers);
showSetupState({ cloudSetup, urlInput: supabaseUrlInput, keyInput: supabaseAnonKeyInput }, !isCloudConfigured());
initCloud().catch((error) => showToast(error.message || "Connessione cloud non riuscita"));
