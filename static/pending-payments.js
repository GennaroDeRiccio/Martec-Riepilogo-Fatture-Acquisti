import {
  fetchPendingPayments,
  isCloudConfigured,
  resetCloudClient,
  saveCloudConfig,
  showSetupState,
  subscribeToChanges,
} from "./cloud.js";

const pendingPaymentsTable = document.querySelector("#pendingPaymentsTable");
const pendingSearchInput = document.querySelector("#pendingSearchInput");
const pendingStatusFilter = document.querySelector("#pendingStatusFilter");
const pendingCount = document.querySelector("#pendingCount");
const toast = document.querySelector("#toast");
const cloudSetup = document.querySelector("#cloudSetup");
const toggleCloudSetupButton = document.querySelector("#toggleCloudSetupButton");
const supabaseUrlInput = document.querySelector("#supabaseUrlInput");
const supabaseAnonKeyInput = document.querySelector("#supabaseAnonKeyInput");
const geminiModelInput = document.querySelector("#geminiModelInput");
const saveCloudConfigButton = document.querySelector("#saveCloudConfigButton");

let pendingPayments = [];
let unsubscribe = null;

function showToast(message) {
  toast.textContent = message;
  toast.classList.add("visible");
  clearTimeout(showToast.timeout);
  showToast.timeout = setTimeout(() => toast.classList.remove("visible"), 3600);
}

function formatAmount(value) {
  const number = Number.parseFloat(String(value || "").replace(/[^\d,.-]/g, "").replace(/\./g, "").replace(",", "."));
  if (!Number.isFinite(number)) return String(value || "");
  return `${number.toLocaleString("it-IT", { minimumFractionDigits: 2, maximumFractionDigits: 2 })} €`;
}

function filterText(payment) {
  return [
    payment.status,
    payment.payment?.paymentType,
    payment.payment?.beneficiary,
    payment.payment?.total,
    payment.payment?.documentDate,
    payment.payment?.executionDate,
    payment.payment?.dueDate,
    payment.payment?.bank,
    payment.payment?.payerIban,
    payment.payment?.reason,
    payment.payment?.noticeNumber,
    payment.notes,
  ].join(" ").toLowerCase();
}

function visiblePendingPayments() {
  const query = pendingSearchInput.value.trim().toLowerCase();
  const status = pendingStatusFilter.value.trim();
  return pendingPayments.filter((payment) => {
    if (status && payment.status !== status) return false;
    if (query && !filterText(payment).includes(query)) return false;
    return true;
  });
}

function statusMeta(status) {
  if (status === "pending_invoice") {
    return { label: "In attesa fattura", className: "badge warn" };
  }
  if (status === "uncertain") {
    return { label: "Da verificare", className: "badge danger" };
  }
  return { label: status || "Sospeso", className: "badge warn" };
}

function bankCell(payment) {
  const bank = payment.bank || "";
  const iban = payment.payerIban || "";
  return [bank, iban].filter(Boolean).join("\n");
}

function paymentDate(payment) {
  return payment.executionDate || payment.documentDate || "";
}

function renderPendingPayments() {
  const tbody = pendingPaymentsTable.querySelector("tbody");
  const visible = visiblePendingPayments();
  tbody.innerHTML = "";
  pendingCount.textContent = visible.length;

  if (!visible.length) {
    const row = document.createElement("tr");
    row.className = "empty-row";
    const cell = document.createElement("td");
    cell.colSpan = 10;
    cell.textContent = pendingPayments.length
      ? "Nessun pagamento in sospeso corrisponde ai filtri"
      : "Nessun pagamento in sospeso memorizzato";
    row.appendChild(cell);
    tbody.appendChild(row);
    return;
  }

  visible.forEach((entry) => {
    const payment = entry.payment || {};
    const row = document.createElement("tr");

    const status = document.createElement("td");
    const statusBadge = document.createElement("span");
    const meta = statusMeta(entry.status);
    statusBadge.className = meta.className;
    statusBadge.textContent = meta.label;
    status.appendChild(statusBadge);

    [
      payment.paymentType || "",
      payment.beneficiary || "",
      formatAmount(payment.total || payment.totalEur || ""),
      paymentDate(payment),
      payment.dueDate || "",
      bankCell(payment),
      payment.reason || "",
      payment.noticeNumber || "",
      entry.notes || "",
    ].forEach((value, index) => {
      const td = document.createElement("td");
      td.textContent = value;
      if ([1, 5, 6, 8].includes(index)) td.classList.add("is-multiline");
      row.appendChild(td);
    });

    row.prepend(status);
    tbody.appendChild(row);
  });
}

async function loadPendingPayments() {
  pendingPayments = await fetchPendingPayments();
  renderPendingPayments();
}

async function initCloud() {
  if (!isCloudConfigured()) {
    showSetupState({
      cloudSetup,
      urlInput: supabaseUrlInput,
      keyInput: supabaseAnonKeyInput,
      geminiModelInput,
    }, true);
    return;
  }
  showSetupState({
    cloudSetup,
    urlInput: supabaseUrlInput,
    keyInput: supabaseAnonKeyInput,
    geminiModelInput,
  }, false);
  if (unsubscribe) unsubscribe();
  await loadPendingPayments();
  unsubscribe = subscribeToChanges(async (tableName) => {
    if (tableName === "pending_payments") await loadPendingPayments();
  });
}

saveCloudConfigButton.addEventListener("click", async () => {
  const url = supabaseUrlInput.value.trim();
  const key = supabaseAnonKeyInput.value.trim();
  const geminiModel = geminiModelInput.value.trim();
  if (!url || !key) {
    showToast("Inserisci URL e anon key.");
    return;
  }
  saveCloudConfig({ supabaseUrl: url, supabaseAnonKey: key, geminiModel });
  resetCloudClient();
  await initCloud();
  showToast("Configurazione cloud salvata");
});

toggleCloudSetupButton?.addEventListener("click", () => {
  const isHidden = cloudSetup.classList.contains("hidden");
  showSetupState({
    cloudSetup,
    urlInput: supabaseUrlInput,
    keyInput: supabaseAnonKeyInput,
    geminiModelInput,
  }, isHidden);
});

pendingSearchInput.addEventListener("input", renderPendingPayments);
pendingStatusFilter.addEventListener("input", renderPendingPayments);

showSetupState({
  cloudSetup,
  urlInput: supabaseUrlInput,
  keyInput: supabaseAnonKeyInput,
  geminiModelInput,
}, !isCloudConfigured());

initCloud().catch((error) => showToast(error.message || "Connessione cloud non riuscita"));
