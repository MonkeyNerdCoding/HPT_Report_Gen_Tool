const modeConfig = {
  oraclehc: {
    label: "OracleHC",
    outputName: "final_oraclehc_report.docx",
    sourceHelper: "OracleHC: ZIP containing the OracleHC HTML folder.",
    placeholders: [
      {
        group: "OracleHC Tables",
        items: [
          { label: "ASH Top SQL for Cluster for 9 days of history", key: "table_ash_top_sql_cluster_9_days", token: "{{table_ash_top_sql_cluster_9_days}}" },
          { label: "ASH CPU per Source", key: "table_ash_cpu_per_source", token: "{{table_ash_cpu_per_source}}" },
          { label: "Instance Summary", key: "table_instance_summary", token: "{{table_instance_summary}}" },
        ],
      },
      {
        group: "OracleHC Charts",
        items: [
          { label: "Database Time", key: "chart_database_time", token: "{{chart_database_time}}" },
        ],
      },
    ],
  },
  sqlhealthcheck: {
    label: "SQLHealthcheck",
    outputName: "final_healthcheck_report.docx",
    sourceHelper: "SQLHealthcheck: ZIP containing the SQLHealthcheck CSV folder.",
    placeholders: [
      {
        group: "SQLHealthcheck Tables",
        items: [
          { label: "Healthcheck Summary", key: "sql_healthcheck_summary", token: "{{sql_healthcheck_summary}}" },
          { label: "Database Findings", key: "sql_database_findings", token: "{{sql_database_findings}}" },
        ],
      },
    ],
  },
};

const mockHistory = [
  { fileName: "final_healthcheck_report.docx", mode: "sqlhealthcheck", template: "sql_template.docx", source: "sql_healthcheck_source.zip", status: "success", createdAt: "Today 10:24", duration: "21s" },
  { fileName: "final_oraclehc_report.docx", mode: "oraclehc", template: "oraclehc_template.docx", source: "oraclehc_source.zip", status: "ready", createdAt: "Yesterday 16:18", duration: "34s" },
];

const state = {
  activeTab: "tool",
  mode: "oraclehc",
  outputTouched: false,
  isGenerating: false,
  lastDownload: null,
  logs: [],
  history: [...mockHistory],
};

const elements = {
  navTabs: document.querySelectorAll("[data-tab-target]"),
  pages: document.querySelectorAll("[data-tab-page]"),
  form: document.querySelector("[data-generate-form]"),
  modeSelector: document.querySelector("[data-mode-selector]"),
  modeValue: document.querySelector("[data-mode-value]"),
  outputName: document.querySelector("[data-output-name]"),
  sourceHelper: document.querySelector("[data-source-helper]"),
  submitButton: document.querySelector("[data-submit-button]"),
  resetButton: document.querySelector("[data-reset-form]"),
  templateInput: document.querySelector("[data-file-input='template']"),
  sourceInput: document.querySelector("[data-file-input='source']"),
  templateName: document.querySelector("[data-file-name='template']"),
  sourceName: document.querySelector("[data-file-name='source']"),
  templateError: document.querySelector("[data-field-error='template']"),
  sourceError: document.querySelector("[data-field-error='source']"),
  outputError: document.querySelector("[data-field-error='output']"),
  placeholderSearch: document.querySelector("[data-placeholder-search]"),
  placeholderGroups: document.querySelector("[data-placeholder-groups]"),
  logList: document.querySelector("[data-log-list]"),
  clearLogs: document.querySelector("[data-clear-logs]"),
  toastRegion: document.querySelector("[data-toast-region]"),
  resultCard: document.querySelector("[data-result-card]"),
  resultFilename: document.querySelector("[data-result-filename]"),
  downloadAgain: document.querySelector("[data-download-again]"),
  generateAnother: document.querySelector("[data-generate-another]"),
  historySearch: document.querySelector("[data-history-search]"),
  historyMode: document.querySelector("[data-history-mode]"),
  historyStatus: document.querySelector("[data-history-status]"),
  historyBody: document.querySelector("[data-history-body]"),
  historyEmpty: document.querySelector("[data-history-empty]"),
};

function init() {
  renderModeSelector();
  applyMode("oraclehc", { forceOutput: true });
  renderPlaceholders();
  addLog("info", "Selected mode: OracleHC");
  renderLogs();
  renderHistory();
  bindEvents();
}

function bindEvents() {
  elements.navTabs.forEach((tab) => tab.addEventListener("click", () => showTab(tab.dataset.tabTarget)));
  elements.outputName.addEventListener("input", () => {
    state.outputTouched = true;
    elements.outputError.textContent = "";
  });
  elements.templateInput.addEventListener("change", () => handleFileChange("template"));
  elements.sourceInput.addEventListener("change", () => handleFileChange("source"));
  elements.form.addEventListener("submit", handleSubmit);
  elements.resetButton.addEventListener("click", () => resetForm({ clearLogs: false }));
  elements.clearLogs.addEventListener("click", clearLogs);
  elements.placeholderSearch.addEventListener("input", renderPlaceholders);
  elements.downloadAgain.addEventListener("click", downloadAgain);
  elements.generateAnother.addEventListener("click", () => resetForm({ clearLogs: false }));
  elements.historySearch.addEventListener("input", renderHistory);
  elements.historyMode.addEventListener("change", renderHistory);
  elements.historyStatus.addEventListener("change", renderHistory);
}

function renderModeSelector() {
  elements.modeSelector.innerHTML = Object.entries(modeConfig)
    .map(([value, config]) => `<button class="mode-option" type="button" data-mode="${value}">${config.label}</button>`)
    .join("");
  elements.modeSelector.querySelectorAll("[data-mode]").forEach((button) => {
    button.addEventListener("click", () => applyMode(button.dataset.mode));
  });
}

function applyMode(mode, options = {}) {
  const config = modeConfig[mode];
  if (!config) return;
  const previousDefault = modeConfig[state.mode]?.outputName;
  const canReplaceOutput = options.forceOutput || !state.outputTouched || elements.outputName.value === previousDefault;
  state.mode = mode;
  elements.modeValue.value = mode;
  elements.sourceHelper.textContent = config.sourceHelper;
  if (canReplaceOutput) {
    elements.outputName.value = config.outputName;
    state.outputTouched = false;
  }
  elements.modeSelector.querySelectorAll("[data-mode]").forEach((button) => {
    button.classList.toggle("is-active", button.dataset.mode === mode);
  });
  addLog("info", `Selected mode: ${config.label}`);
  renderPlaceholders();
  renderLogs();
}

function handleFileChange(type) {
  const input = type === "template" ? elements.templateInput : elements.sourceInput;
  const label = type === "template" ? elements.templateName : elements.sourceName;
  const error = type === "template" ? elements.templateError : elements.sourceError;
  const expected = type === "template" ? ".docx" : ".zip";
  const file = input.files[0];
  error.textContent = "";
  label.textContent = file ? file.name : type === "template" ? "No template selected" : "No source package selected";
  if (!file) return;
  if (!file.name.toLowerCase().endsWith(expected)) {
    error.textContent = `File must be ${expected}.`;
    addLog("error", `${type === "template" ? "Word template" : "Source package"} has invalid file type`);
    showToast(`Invalid ${type === "template" ? "template" : "source package"} file`, "error");
    renderLogs();
    return;
  }
  const message = type === "template" ? "Template uploaded successfully" : "Source package uploaded successfully";
  addLog("success", message);
  showToast(message, "success");
  renderLogs();
}

async function handleSubmit(event) {
  event.preventDefault();
  if (state.isGenerating) return;
  if (!validateForm()) return;
  state.isGenerating = true;
  setGenerating(true);
  elements.resultCard.hidden = true;
  addLog("info", `Start generating report for ${modeConfig[state.mode].label}`);
  renderLogs();
  const startedAt = performance.now();
  const formData = new FormData(elements.form);
  formData.set("mode", state.mode);
  try {
    // TODO: connect richer backend progress events when the generator exposes them.
    addLog("info", "Placeholder scan started");
    renderLogs();
    const response = await fetch(elements.form.action, { method: "POST", body: formData });
    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(extractErrorMessage(errorText) || "Failed to generate report");
    }
    addLog("success", "Placeholder scan completed");
    const blob = await response.blob();
    const fileName = getDownloadFilename(response) || normalizeDocxName(elements.outputName.value);
    state.lastDownload = { blob, fileName };
    triggerDownload(blob, fileName);
    addLog("success", "Report generated successfully");
    addLog("success", "File downloaded");
    showResult(fileName);
    showToast("Report generated successfully", "success");
    addHistoryItem(fileName, startedAt, "success");
  } catch (error) {
    addLog("error", error.message);
    showToast("Failed to generate report", "error");
    addHistoryItem(normalizeDocxName(elements.outputName.value), startedAt, "failed");
  } finally {
    state.isGenerating = false;
    setGenerating(false);
    renderLogs();
    renderHistory();
  }
}

function validateForm() {
  let valid = true;
  elements.templateError.textContent = "";
  elements.sourceError.textContent = "";
  elements.outputError.textContent = "";
  const template = elements.templateInput.files[0];
  const source = elements.sourceInput.files[0];
  const output = elements.outputName.value.trim();
  if (!template) {
    elements.templateError.textContent = "Please upload a Word template.";
    valid = false;
  } else if (!template.name.toLowerCase().endsWith(".docx")) {
    elements.templateError.textContent = "Word template must be a .docx file.";
    valid = false;
  }
  if (!source) {
    elements.sourceError.textContent = "Please upload a source package.";
    valid = false;
  } else if (!source.name.toLowerCase().endsWith(".zip")) {
    elements.sourceError.textContent = "Source package must be a .zip file.";
    valid = false;
  }
  if (!output) {
    elements.outputError.textContent = "Output filename is required.";
    valid = false;
  }
  if (!valid) {
    addLog("warning", "Validation failed. Check required inputs.");
    showToast("Please check required inputs", "warning");
    renderLogs();
  }
  return valid;
}

function setGenerating(isGenerating) {
  elements.submitButton.disabled = isGenerating;
  elements.submitButton.textContent = isGenerating ? "Generating..." : "Generate Report";
}

function resetForm(options = { clearLogs: false }) {
  elements.form.reset();
  elements.templateName.textContent = "No template selected";
  elements.sourceName.textContent = "No source package selected";
  elements.templateError.textContent = "";
  elements.sourceError.textContent = "";
  elements.outputError.textContent = "";
  state.outputTouched = false;
  applyMode(state.mode, { forceOutput: true });
  elements.resultCard.hidden = true;
  state.lastDownload = null;
  state.isGenerating = false;
  setGenerating(false);
  if (options.clearLogs) clearLogs();
  showToast("Form reset", "info");
}

function showResult(fileName) {
  elements.resultFilename.textContent = fileName;
  elements.resultCard.hidden = false;
}

function downloadAgain() {
  if (!state.lastDownload) return;
  triggerDownload(state.lastDownload.blob, state.lastDownload.fileName);
  addLog("success", "File downloaded again");
  showToast("Download started", "success");
  renderLogs();
}

function triggerDownload(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

function getDownloadFilename(response) {
  const disposition = response.headers.get("content-disposition") || "";
  const match = disposition.match(/filename="?([^"]+)"?/i);
  return match ? match[1] : "";
}

function normalizeDocxName(fileName) {
  const trimmed = fileName.trim() || modeConfig[state.mode].outputName;
  return trimmed.toLowerCase().endsWith(".docx") ? trimmed : `${trimmed}.docx`;
}

function extractErrorMessage(html) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(html, "text/html");
  const alert = doc.querySelector(".alert span");
  return alert ? alert.textContent.replace(/^Reason:\s*/, "") : "";
}

function addLog(status, message) {
  state.logs.push({ time: new Date().toLocaleTimeString("en-GB", { hour12: false }), status, message });
}

function renderLogs() {
  if (state.logs.length === 0) {
    elements.logList.innerHTML = `<div class="empty-inline">No logs yet.</div>`;
    return;
  }
  elements.logList.innerHTML = state.logs
    .map((log) => `<div class="log-row ${log.status}"><span class="log-icon">${statusIcon(log.status)}</span><time>[${log.time}]</time><span>${escapeHtml(log.message)}</span></div>`)
    .join("");
  elements.logList.scrollTop = elements.logList.scrollHeight;
}

function clearLogs() {
  state.logs = [];
  renderLogs();
}

function statusIcon(status) {
  return { info: "i", success: "ok", warning: "!", error: "x" }[status] || "i";
}

function renderPlaceholders() {
  const query = elements.placeholderSearch.value.trim().toLowerCase();
  const groups = modeConfig[state.mode].placeholders
    .map((group) => ({ ...group, items: group.items.filter((item) => `${item.label} ${item.key} ${item.token}`.toLowerCase().includes(query)) }))
    .filter((group) => group.items.length > 0);
  elements.placeholderGroups.innerHTML = groups
    .map((group, index) => `<details class="placeholder-group" ${index === 0 ? "open" : ""}><summary>${group.group}</summary><div class="placeholder-list">${group.items.map(renderPlaceholderItem).join("")}</div></details>`)
    .join("") || `<div class="empty-inline">No placeholders found.</div>`;
  elements.placeholderGroups.querySelectorAll("[data-copy-placeholder]").forEach((button) => {
    button.addEventListener("click", () => copyPlaceholder(button));
  });
}

function renderPlaceholderItem(item) {
  return `<div class="placeholder-item"><div><strong>${item.label}</strong><code>${item.token}</code><small>Mapping key: ${item.key}</small></div><button class="copy-button" type="button" data-copy-placeholder="${item.token}">Copy</button></div>`;
}

async function copyPlaceholder(button) {
  const token = button.dataset.copyPlaceholder;
  try {
    await navigator.clipboard.writeText(token);
  } catch (_error) {
    const temp = document.createElement("textarea");
    temp.value = token;
    document.body.appendChild(temp);
    temp.select();
    document.execCommand("copy");
    temp.remove();
  }
  button.textContent = "Copied";
  addLog("success", `Placeholder copied: ${token}`);
  showToast("Placeholder copied", "success");
  renderLogs();
  setTimeout(() => {
    button.textContent = "Copy";
  }, 1300);
}

function showTab(tabName) {
  state.activeTab = tabName;
  elements.navTabs.forEach((tab) => tab.classList.toggle("is-active", tab.dataset.tabTarget === tabName));
  elements.pages.forEach((page) => page.classList.toggle("is-active", page.dataset.tabPage === tabName));
}

function addHistoryItem(fileName, startedAt, status) {
  const duration = `${Math.max(1, Math.round((performance.now() - startedAt) / 1000))}s`;
  state.history.unshift({
    fileName,
    mode: state.mode,
    template: elements.templateInput.files[0]?.name || "-",
    source: elements.sourceInput.files[0]?.name || "-",
    status,
    createdAt: new Date().toLocaleString("en-GB", { day: "2-digit", month: "2-digit", hour: "2-digit", minute: "2-digit" }),
    duration,
  });
}

function renderHistory() {
  const search = elements.historySearch.value.trim().toLowerCase();
  const mode = elements.historyMode.value;
  const status = elements.historyStatus.value;
  const rows = state.history.filter((item) => {
    return item.fileName.toLowerCase().includes(search) && (mode === "all" || item.mode === mode) && (status === "all" || item.status === status);
  });
  elements.historyEmpty.hidden = rows.length > 0;
  elements.historyBody.innerHTML = rows
    .map((item) => `<tr><td><strong>${escapeHtml(item.fileName)}</strong></td><td>${modeConfig[item.mode]?.label || item.mode}</td><td>${escapeHtml(item.template)}</td><td>${escapeHtml(item.source)}</td><td><span class="status-badge ${item.status}">${capitalize(item.status)}</span></td><td>${item.createdAt}</td><td>${item.duration}</td><td><div class="table-actions"><button type="button">Download</button><button type="button">View Logs</button><button type="button">Regenerate</button><button type="button">Delete</button></div></td></tr>`)
    .join("");
}

function showToast(message, type = "info") {
  const toast = document.createElement("div");
  toast.className = `toast ${type}`;
  toast.textContent = message;
  elements.toastRegion.appendChild(toast);
  setTimeout(() => {
    toast.classList.add("is-leaving");
    setTimeout(() => toast.remove(), 220);
  }, 2600);
}

function escapeHtml(value) {
  return String(value).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#039;");
}

function capitalize(value) {
  return value.charAt(0).toUpperCase() + value.slice(1);
}

init();
