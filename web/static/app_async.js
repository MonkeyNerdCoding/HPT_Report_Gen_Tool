const modeConfig = {
  oraclehc: {
    label: "OracleHC",
    outputName: "final_oraclehc_report.docx",
    sourceHelper: "ZIP containing OracleHC HTML output.",
  },
  sqlhealthcheck: {
    label: "SQLHealthcheck",
    outputName: "final_healthcheck_report.docx",
    sourceHelper: "ZIP containing SQLHealthcheck CSV output.",
  },
};

const state = {
  mode: "oraclehc",
  outputTouched: false,
  isGenerating: false,
  activeJobId: null,
  pollTimer: null,
  lastDownloadJobId: null,
  logs: [],
  logFilter: "all",
  history: [],
  companies: [],
  activeCompanyId: "",
  pendingAssignJobId: "",
  placeholders: [],
  visiblePlaceholders: [],
  lastPlaceholderScan: null,
  editingPlaceholder: "",
  aiReviewRows: [],
  isAiReviewing: false,
};

const el = {
  navTabs: document.querySelectorAll("[data-tab-target]"),
  pages: document.querySelectorAll("[data-tab-page]"),
  form: document.querySelector("[data-generate-form]"),
  modeSelector: document.querySelector("[data-mode-selector]"),
  modeValue: document.querySelector("[data-mode-value]"),
  outputName: document.querySelector("[data-output-name]"),
  sourceHelper: document.querySelector("[data-source-helper]"),
  submitButton: document.querySelector("[data-submit-button]"),
  resetButton: document.querySelector("[data-reset-form]"),
  insertButton: document.querySelector("[data-insert-placeholders]"),
  templateInput: document.querySelector("[data-file-input='template']"),
  sourceInput: document.querySelector("[data-file-input='source']"),
  templateName: document.querySelector("[data-file-name='template']"),
  sourceName: document.querySelector("[data-file-name='source']"),
  templateMeta: document.querySelector("[data-file-meta='template']"),
  sourceMeta: document.querySelector("[data-file-meta='source']"),
  templateError: document.querySelector("[data-field-error='template']"),
  sourceError: document.querySelector("[data-field-error='source']"),
  outputError: document.querySelector("[data-field-error='output']"),
  loadingModal: document.querySelector("[data-loading-modal]"),
  loadingTitle: document.querySelector("[data-loading-title]"),
  loadingStep: document.querySelector("[data-loading-step]"),
  loadingBar: document.querySelector("[data-loading-bar]"),
  loadingPercent: document.querySelector("[data-loading-percent]"),
  logList: document.querySelector("[data-log-list]"),
  clearLogs: document.querySelector("[data-clear-logs]"),
  copyLogs: document.querySelector("[data-copy-logs]"),
  downloadLogs: document.querySelector("[data-download-logs]"),
  logFilters: document.querySelector("[data-log-filters]"),
  resultCard: document.querySelector("[data-result-card]"),
  resultTitle: document.querySelector("[data-result-title]"),
  resultFilename: document.querySelector("[data-result-filename]"),
  downloadAgain: document.querySelector("[data-download-again]"),
  generateAnother: document.querySelector("[data-generate-another]"),
  toastRegion: document.querySelector("[data-toast-region]"),
  historySearch: document.querySelector("[data-history-search]"),
  historyMode: document.querySelector("[data-history-mode]"),
  historyStatus: document.querySelector("[data-history-status]"),
  historyStats: document.querySelector("[data-history-stats]"),
  historyBody: document.querySelector("[data-history-body]"),
  historyEmpty: document.querySelector("[data-history-empty]"),
  companySearch: document.querySelector("[data-company-search]"),
  companyBody: document.querySelector("[data-company-body]"),
  companyEmpty: document.querySelector("[data-company-empty]"),
  newCompany: document.querySelector("[data-new-company]"),
  companyDetail: document.querySelector("[data-company-detail]"),
  companyDetailTitle: document.querySelector("[data-company-detail-title]"),
  companyReportType: document.querySelector("[data-company-report-type]"),
  companyQuarter: document.querySelector("[data-company-quarter]"),
  companyYear: document.querySelector("[data-company-year]"),
  companyReportStatus: document.querySelector("[data-company-report-status]"),
  companyReportBody: document.querySelector("[data-company-report-body]"),
  companyReportEmpty: document.querySelector("[data-company-report-empty]"),
  companyModal: document.querySelector("[data-company-modal]"),
  companyForm: document.querySelector("[data-company-form]"),
  companyModalTitle: document.querySelector("[data-company-modal-title]"),
  companyId: document.querySelector("[data-company-id]"),
  companyFields: document.querySelectorAll("[data-company-field]"),
  closeCompanyModal: document.querySelectorAll("[data-close-company-modal]"),
  assignModal: document.querySelector("[data-assign-modal]"),
  assignForm: document.querySelector("[data-assign-form]"),
  assignJobId: document.querySelector("[data-assign-job-id]"),
  assignSummary: document.querySelector("[data-assign-summary]"),
  assignFields: document.querySelectorAll("[data-assign-field]"),
  closeAssignModal: document.querySelectorAll("[data-close-assign-modal]"),
  placeholderSearch: document.querySelector("[data-placeholder-search]"),
  placeholderType: document.querySelector("[data-placeholder-type]"),
  placeholderStats: document.querySelector("[data-placeholder-stats]"),
  placeholderBody: document.querySelector("[data-placeholder-body]"),
  placeholderEmpty: document.querySelector("[data-placeholder-empty]"),
  refreshPlaceholders: document.querySelector("[data-refresh-placeholders]"),
  placeholderEditor: document.querySelector("[data-placeholder-editor]"),
  placeholderEditorMode: document.querySelector("[data-placeholder-editor-mode]"),
  placeholderEditForm: document.querySelector("[data-placeholder-edit-form]"),
  placeholderOriginal: document.querySelector("[data-placeholder-original]"),
  placeholderFields: document.querySelectorAll("[data-placeholder-field]"),
  placeholderNew: document.querySelector("[data-placeholder-new]"),
  pageScrollButtons: document.querySelectorAll("[data-page-scroll]"),
  scrollActions: document.querySelector("[data-scroll-actions]"),
  placeholderScanForm: document.querySelector("[data-placeholder-scan-form]"),
  placeholderFile: document.querySelector("[data-placeholder-file]"),
  placeholderFileName: document.querySelector("[data-placeholder-file-name]"),
  placeholderFileMeta: document.querySelector("[data-placeholder-file-meta]"),
  placeholderFileError: document.querySelector("[data-placeholder-file-error]"),
  scanButton: document.querySelector("[data-scan-placeholders]"),
  addNewPlaceholders: document.querySelector("[data-add-new-placeholders]"),
  scanResults: document.querySelector("[data-placeholder-scan-results]"),
  scanFound: document.querySelector("[data-scan-found]"),
  scanMissing: document.querySelector("[data-scan-missing]"),
  scanNew: document.querySelector("[data-scan-new]"),
  aiReviewForm: document.querySelector("[data-ai-review-form]"),
  aiHistorySource: document.querySelector("[data-ai-history-source]"),
  aiReviewType: document.querySelector("[data-ai-review-type]"),
  aiOutputStyle: document.querySelector("[data-ai-output-style]"),
  aiSourceFile: document.querySelector("[data-ai-source-file]"),
  aiSourceName: document.querySelector("[data-ai-source-name]"),
  aiSourceMeta: document.querySelector("[data-ai-source-meta]"),
  aiSourceError: document.querySelector("[data-ai-source-error]"),
  aiGenerate: document.querySelector("[data-ai-generate]"),
  aiCopy: document.querySelector("[data-ai-copy]"),
  aiExportDocx: document.querySelector("[data-ai-export-docx]"),
  aiExportXlsx: document.querySelector("[data-ai-export-xlsx]"),
  aiReviewStatus: document.querySelector("[data-ai-review-status]"),
  aiReviewResult: document.querySelector("[data-ai-review-result]"),
  aiReviewBody: document.querySelector("[data-ai-review-body]"),
  aiScanPanel: document.querySelector("[data-ai-scan-panel]"),
  aiLiveIndicator: document.querySelector("[data-ai-live-indicator]"),
  aiStepList: document.querySelector("[data-ai-step-list]"),
};

function init() {
  renderModeSelector();
  applyMode("oraclehc", { forceOutput: true });
  setLoadingState(false, 0, "Ready", "Waiting for input");
  bindEvents();
  addLog("info", "init", "Tool ready");
  loadHistory();
  loadCompanies();
  loadPlaceholders();
  renderAiHistoryOptions();
  updateGenerateAvailability();
  updateScrollActions();
}

function bindEvents() {
  el.navTabs.forEach((tab) => tab.addEventListener("click", () => showTab(tab.dataset.tabTarget)));
  el.outputName.addEventListener("input", () => {
    state.outputTouched = true;
    updateGenerateAvailability();
  });
  el.templateInput.addEventListener("change", () => handleFileChange("template"));
  el.sourceInput.addEventListener("change", () => handleFileChange("source"));
  el.form.addEventListener("submit", handleSubmit);
  el.resetButton.addEventListener("click", resetForm);
  el.insertButton.addEventListener("click", insertTemplatePlaceholders);
  el.clearLogs.addEventListener("click", clearLogs);
  el.copyLogs.addEventListener("click", copyLogs);
  el.downloadLogs.addEventListener("click", downloadLogs);
  el.logFilters.addEventListener("click", (event) => {
    const button = event.target.closest("[data-log-level]");
    if (!button) return;
    state.logFilter = button.dataset.logLevel;
    el.logFilters.querySelectorAll("button").forEach((item) => item.classList.toggle("is-active", item === button));
    renderLogs();
  });
  el.downloadAgain.addEventListener("click", () => {
    if (state.lastDownloadJobId) window.location.href = `/api/report/download/${state.lastDownloadJobId}`;
  });
  el.generateAnother.addEventListener("click", resetForm);
  el.historySearch.addEventListener("input", renderHistory);
  el.historyMode.addEventListener("change", renderHistory);
  el.historyStatus.addEventListener("change", renderHistory);
  el.companySearch.addEventListener("input", renderCompanies);
  el.newCompany.addEventListener("click", () => openCompanyModal());
  el.companyForm.addEventListener("submit", saveCompany);
  el.closeCompanyModal.forEach((button) => button.addEventListener("click", closeCompanyModal));
  el.companyReportType.addEventListener("change", loadCompanyReports);
  el.companyQuarter.addEventListener("change", loadCompanyReports);
  el.companyYear.addEventListener("input", loadCompanyReports);
  el.companyReportStatus.addEventListener("change", loadCompanyReports);
  el.assignForm.addEventListener("submit", saveAssignment);
  el.closeAssignModal.forEach((button) => button.addEventListener("click", closeAssignModal));
  el.placeholderSearch.addEventListener("input", renderPlaceholders);
  el.placeholderType.addEventListener("change", renderPlaceholders);
  el.refreshPlaceholders.addEventListener("click", (event) => {
    event.preventDefault();
    event.stopPropagation();
    refreshPlaceholders();
  });
  el.placeholderEditForm.addEventListener("submit", savePlaceholderFromForm);
  el.placeholderNew.addEventListener("click", resetPlaceholderForm);
  el.pageScrollButtons.forEach((button) => {
    button.addEventListener("click", (event) => scrollPage(event, button.dataset.pageScroll));
  });
  window.addEventListener("scroll", updateScrollActions, { passive: true });
  window.addEventListener("resize", updateScrollActions);
  el.placeholderFile.addEventListener("change", handlePlaceholderFileChange);
  el.placeholderScanForm.addEventListener("submit", scanPlaceholders);
  el.addNewPlaceholders.addEventListener("click", addNewPlaceholdersToYaml);
  el.aiReviewForm.addEventListener("submit", generateAiReview);
  el.aiHistorySource.addEventListener("change", handleAiSourceChoice);
  el.aiSourceFile.addEventListener("change", handleAiSourceFileChange);
  el.aiCopy.addEventListener("click", copyAiReviewTable);
  el.aiExportDocx.addEventListener("click", () => exportAiReview("docx"));
  el.aiExportXlsx.addEventListener("click", () => exportAiReview("xlsx"));
}

function renderModeSelector() {
  el.modeSelector.innerHTML = Object.entries(modeConfig)
    .map(([value, config]) => `<button class="mode-option" type="button" data-mode="${value}">${config.label}</button>`)
    .join("");
  el.modeSelector.querySelectorAll("[data-mode]").forEach((button) => {
    button.addEventListener("click", () => applyMode(button.dataset.mode));
  });
}

function applyMode(mode, options = {}) {
  const previousDefault = modeConfig[state.mode]?.outputName;
  const canReplaceOutput = options.forceOutput || !state.outputTouched || el.outputName.value === previousDefault;
  state.mode = mode;
  el.modeValue.value = mode;
  el.sourceHelper.textContent = modeConfig[mode].sourceHelper;
  if (canReplaceOutput) {
    el.outputName.value = modeConfig[mode].outputName;
    state.outputTouched = false;
  }
  el.modeSelector.querySelectorAll("[data-mode]").forEach((button) => {
    button.classList.toggle("is-active", button.dataset.mode === mode);
  });
  addLog("info", "select_mode", `Selected mode: ${modeConfig[mode].label}`);
  updateGenerateAvailability();
  loadPlaceholders();
}

function scrollPage(event, direction) {
  event.preventDefault();
  window.scrollTo({
    top: direction === "bottom" ? document.documentElement.scrollHeight : 0,
    behavior: "smooth",
  });
}

function updateScrollActions() {
  if (!el.scrollActions) return;
  const scrollTop = window.scrollY || document.documentElement.scrollTop;
  const maxScroll = Math.max(0, document.documentElement.scrollHeight - window.innerHeight);
  const canGoTop = scrollTop > 80;
  const canGoBottom = scrollTop < maxScroll - 80;

  el.pageScrollButtons.forEach((button) => {
    const direction = button.dataset.pageScroll;
    button.classList.toggle("is-visible", direction === "top" ? canGoTop : canGoBottom);
  });
}

function handleFileChange(type) {
  const input = fileInput(type);
  const file = input.files[0];
  const expected = type === "template" ? ".docx" : ".zip";
  const nameEl = type === "template" ? el.templateName : el.sourceName;
  const metaEl = type === "template" ? el.templateMeta : el.sourceMeta;
  const errorEl = type === "template" ? el.templateError : el.sourceError;
  errorEl.textContent = "";

  if (!file) {
    nameEl.textContent = type === "template" ? "No template selected" : "No source data selected";
    metaEl.textContent = "Waiting for upload";
    updateGenerateAvailability();
    return;
  }

  nameEl.textContent = file.name;
  metaEl.textContent = `${formatBytes(file.size)} | Ready`;
  if (!file.name.toLowerCase().endsWith(expected)) {
    errorEl.textContent = `File must be ${expected}.`;
    addLog("error", "validate_file", `${file.name} has invalid file type`);
    showToast(`Invalid ${type === "template" ? "template" : "source data"} file`, "error");
  } else {
    addLog("success", "upload_file", `${file.name} uploaded successfully`);
    showToast(type === "template" ? "Template uploaded successfully" : "Source data uploaded successfully", "success");
  }
  updateGenerateAvailability();
}

async function insertTemplatePlaceholders() {
  if (state.isGenerating) return;
  const file = el.templateInput.files[0];
  if (!file || !file.name.toLowerCase().endsWith(".docx")) {
    el.templateError.textContent = "Upload a valid .docx template before inserting placeholders.";
    showToast("Upload a Word template first", "warning");
    return;
  }
  state.isGenerating = true;
  state.activeJobId = null;
  updateGenerateAvailability();
  addLog("info", "insert_placeholders", "Inserting placeholders into Word template");
  setLoadingState(true, 8, "Inserting placeholders", "Applying mapping to Word template");
  const formData = new FormData();
  formData.set("template_file", file);
  try {
    const response = await fetch("/api/template/insert-job", { method: "POST", body: formData });
    const result = await response.json();
    if (!response.ok || !result.success) throw new Error(result.detail || "Placeholder insertion failed");
    state.activeJobId = result.job_id;
    addLog("success", "insert_created", `Insert job created: ${result.job_id}`);
    setLoadingState(true, 10, "Inserting placeholders", "Job created");
    pollJob(result.job_id, { downloadOnSuccess: true, successToast: "Placeholders inserted successfully" });
  } catch (error) {
    state.isGenerating = false;
    state.activeJobId = null;
    updateGenerateAvailability();
    setLoadingState(false, 0, "Ready", "Waiting for input");
    addLog("error", "insert_placeholders", error.message);
    showToast("Placeholder insertion failed", "error");
  }
}

async function handleSubmit(event) {
  event.preventDefault();
  if (state.isGenerating || !validateForm()) return;
  state.isGenerating = true;
  state.activeJobId = null;
  setGenerating(true);
  el.resultCard.hidden = true;
  setLoadingState(true, 0, "Generating report", "Creating job");
  addLog("info", "generate", "Creating background report job");

  const formData = new FormData(el.form);
  formData.set("mode", state.mode);
  formData.set("output_name", normalizeDocxName(el.outputName.value));

  try {
    const response = await fetch("/api/report/generate", { method: "POST", body: formData });
    const result = await response.json();
    if (!response.ok || !result.success) throw new Error(result.detail || "Failed to create report job");
    state.activeJobId = result.job_id;
    addLog("success", "job_created", `Job created: ${result.job_id}`);
    setLoadingState(true, 10, "Generating report", "Job created");
    pollJob(result.job_id, { downloadOnSuccess: true, successToast: "Report generated successfully" });
  } catch (error) {
    state.isGenerating = false;
    state.activeJobId = null;
    setGenerating(false);
    setLoadingState(false, 0, "Ready", "Waiting for input");
    addLog("error", "generate", error.message);
    showToast("Failed to generate report", "error");
  }
}

async function pollJob(jobId, options = {}) {
  const settings = { downloadOnSuccess: true, successToast: "Report generated successfully", ...options };
  clearTimeout(state.pollTimer);
  try {
    const response = await fetch(`/api/report/status/${jobId}`);
    const result = await response.json();
    if (state.activeJobId !== jobId) return;
    if (!response.ok || !result.success) throw new Error(result.detail || "Could not read job status");
    const job = result.job;
    state.logs = result.logs.map((log) => ({
      timestamp: log.timestamp,
      level: log.level,
      step: log.step || "",
      message: log.message,
    }));
    renderLogs();
    setLoadingState(true, job.progress || 0, capitalize(job.status), job.current_step || "Processing");

    if (job.status === "success") {
      state.isGenerating = false;
      state.activeJobId = null;
      state.lastDownloadJobId = job.id;
      setGenerating(false);
      setLoadingState(false, 0, "Ready", "Waiting for input");
      showResult(job.output_file_name, false);
      showToast(settings.successToast, "success");
      await loadHistory();
      if (settings.downloadOnSuccess) window.location.href = `/api/report/download/${job.id}`;
      return;
    }

    if (job.status === "failed") {
      state.isGenerating = false;
      state.activeJobId = null;
      setGenerating(false);
      setLoadingState(false, 0, "Failed", job.current_step || "Generation failed");
      showResult(job.output_file_name || "Report", true, job.error_message || "Generation failed");
      showToast("Failed to generate report", "error");
      await loadHistory();
      return;
    }

    state.pollTimer = setTimeout(() => pollJob(jobId, settings), 1200);
  } catch (error) {
    if (state.activeJobId !== jobId) return;
    state.isGenerating = false;
    state.activeJobId = null;
    setGenerating(false);
    addLog("error", "poll_status", error.message);
    showToast("Could not read job status", "error");
  }
}

function validateForm() {
  let valid = true;
  clearFieldErrors();
  const template = el.templateInput.files[0];
  const source = el.sourceInput.files[0];
  const output = el.outputName.value.trim();
  if (!template) {
    el.templateError.textContent = "Please upload a Word template.";
    valid = false;
  } else if (!template.name.toLowerCase().endsWith(".docx")) {
    el.templateError.textContent = "Word template must be a .docx file.";
    valid = false;
  }
  if (!source) {
    el.sourceError.textContent = "Please upload a source data package.";
    valid = false;
  } else if (!source.name.toLowerCase().endsWith(".zip")) {
    el.sourceError.textContent = "Source data package must be a .zip file.";
    valid = false;
  }
  if (!output) {
    el.outputError.textContent = "Output filename is required.";
    valid = false;
  }
  if (!valid) addLog("warning", "validation", "Validation failed. Check required inputs.");
  return valid;
}

function updateGenerateAvailability() {
  const template = el.templateInput.files[0];
  const source = el.sourceInput.files[0];
  const output = el.outputName.value.trim();
  const ready = template && source && output && template.name.toLowerCase().endsWith(".docx") && source.name.toLowerCase().endsWith(".zip");
  el.submitButton.disabled = state.isGenerating || !ready;
  el.insertButton.disabled = state.isGenerating || !template || !template.name.toLowerCase().endsWith(".docx");
}

function setLoadingState(visible, progress, title, step) {
  if (!el.loadingModal) return;
  el.loadingModal.hidden = !visible;
  el.loadingModal.setAttribute("aria-hidden", visible ? "false" : "true");
  el.loadingModal.classList.toggle("is-running", visible);
  if (el.loadingTitle) el.loadingTitle.textContent = title;
  if (el.loadingStep) el.loadingStep.textContent = step;
  if (el.loadingPercent) el.loadingPercent.textContent = `${Math.max(0, Math.min(100, Math.round(progress)))}%`;
  if (el.loadingBar) el.loadingBar.style.width = `${Math.max(0, Math.min(100, progress))}%`;
}

async function loadHistory() {
  try {
    const response = await fetch("/api/history");
    const result = await response.json();
    if (!response.ok || !result.success) throw new Error("Could not load history");
    state.history = result.jobs || [];
    renderHistory();
    renderAiHistoryOptions();
  } catch (error) {
    addLog("warning", "history", error.message);
  }
}

async function loadCompanies() {
  try {
    const response = await fetch("/api/companies");
    const result = await response.json();
    if (!response.ok || !result.success) throw new Error(result.detail || "Could not load companies");
    state.companies = result.companies || [];
    renderCompanies();
    renderAssignCompanyOptions();
  } catch (error) {
    addLog("warning", "company", error.message);
  }
}

function renderCompanies() {
  const query = el.companySearch.value.trim().toLowerCase();
  const rows = state.companies.filter((company) => {
    return [company.company_name, company.short_name, company.customer_code, company.contact_person]
      .join(" ")
      .toLowerCase()
      .includes(query);
  });
  el.companyEmpty.hidden = rows.length > 0;
  el.companyBody.innerHTML = rows.map((company) => `
    <tr>
      <td><strong>${escapeHtml(company.company_name)}</strong></td>
      <td>${escapeHtml(company.short_name)}</td>
      <td>${escapeHtml(company.contact_person || "-")}</td>
      <td>${Number(company.total_reports || 0)}</td>
      <td>${escapeHtml(company.latest_report || "-")}</td>
      <td>${formatDate(company.last_updated || company.updated_at)}</td>
      <td>
        <div class="table-actions">
          <button type="button" data-company-view="${company.id}">View</button>
          <button type="button" data-company-edit="${company.id}">Edit</button>
          <button type="button" data-company-delete="${company.id}">Delete</button>
        </div>
      </td>
    </tr>
  `).join("");
  el.companyBody.querySelectorAll("[data-company-view]").forEach((button) => {
    button.addEventListener("click", () => viewCompany(button.dataset.companyView));
  });
  el.companyBody.querySelectorAll("[data-company-edit]").forEach((button) => {
    button.addEventListener("click", () => {
      const company = state.companies.find((item) => item.id === button.dataset.companyEdit);
      if (company) openCompanyModal(company);
    });
  });
  el.companyBody.querySelectorAll("[data-company-delete]").forEach((button) => {
    button.addEventListener("click", () => deleteCompany(button.dataset.companyDelete));
  });
}

function openCompanyModal(company = null) {
  el.companyForm.reset();
  el.companyId.value = company?.id || "";
  el.companyModalTitle.textContent = company ? "Edit Company" : "New Company";
  el.companyFields.forEach((field) => {
    field.value = company?.[field.dataset.companyField] || "";
  });
  el.companyModal.hidden = false;
  el.companyModal.setAttribute("aria-hidden", "false");
}

function closeCompanyModal() {
  el.companyModal.hidden = true;
  el.companyModal.setAttribute("aria-hidden", "true");
}

async function saveCompany(event) {
  event.preventDefault();
  const payload = {};
  el.companyFields.forEach((field) => {
    payload[field.dataset.companyField] = field.value.trim();
  });
  const companyId = el.companyId.value;
  try {
    const response = await fetch(companyId ? `/api/companies/${companyId}` : "/api/companies", {
      method: companyId ? "PUT" : "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    const result = await response.json();
    if (!response.ok || !result.success) throw new Error(result.detail || "Could not save company");
    await loadCompanies();
    showToast("Company saved", "success");
    closeCompanyModal();
  } catch (error) {
    showToast(error.message, "error");
  }
}

async function deleteCompany(companyId) {
  const company = state.companies.find((item) => item.id === companyId);
  if (!company || !confirm(`Delete company ${company.company_name}? Reports will become Unassigned.`)) return;
  try {
    const response = await fetch(`/api/companies/${companyId}`, { method: "DELETE" });
    const result = await response.json();
    if (!response.ok || !result.success) throw new Error(result.detail || "Could not delete company");
    if (state.activeCompanyId === companyId) {
      state.activeCompanyId = "";
      el.companyDetail.hidden = true;
    }
    await loadCompanies();
    await loadHistory();
    showToast("Company deleted", "success");
  } catch (error) {
    showToast(error.message, "error");
  }
}

async function viewCompany(companyId) {
  const company = state.companies.find((item) => item.id === companyId);
  if (!company) return;
  state.activeCompanyId = companyId;
  el.companyDetail.hidden = false;
  el.companyDetailTitle.textContent = `Company: ${company.company_name}`;
  await loadCompanyReports();
  el.companyDetail.scrollIntoView({ behavior: "smooth", block: "start" });
}

async function loadCompanyReports() {
  if (!state.activeCompanyId) return;
  const params = new URLSearchParams({
    report_type: el.companyReportType.value || "all",
    quarter: el.companyQuarter.value || "all",
    year: el.companyYear.value.trim() || "all",
    report_status: el.companyReportStatus.value || "all",
  });
  try {
    const response = await fetch(`/api/companies/${state.activeCompanyId}/reports?${params}`);
    const result = await response.json();
    if (!response.ok || !result.success) throw new Error(result.detail || "Could not load company reports");
    renderCompanyReports(result.reports || []);
  } catch (error) {
    showToast(error.message, "error");
  }
}

function renderCompanyReports(rows) {
  el.companyReportEmpty.hidden = rows.length > 0;
  el.companyReportBody.innerHTML = rows.map((job) => `
    <tr>
      <td>${escapeHtml(job.report_type || "Unclassified")}</td>
      <td>${escapeHtml(reportPeriod(job))}</td>
      <td>${escapeHtml(job.month || "-")}</td>
      <td>${escapeHtml(job.year || "-")}</td>
      <td><span class="status-badge ${statusClass(job.report_status || job.status)}">${labelStatus(job.report_status || job.status)}</span></td>
      <td>${escapeHtml(job.note || "-")}</td>
      <td>${formatDate(job.created_at)}</td>
      <td>
        <div class="table-actions">
          <button type="button" data-history-download="${job.id}" ${job.status === "success" ? "" : "disabled"}>Download</button>
          <button type="button" data-history-assign="${job.id}">Assign/Edit</button>
        </div>
      </td>
    </tr>
  `).join("");
  bindHistoryActions();
}

async function loadPlaceholders() {
  try {
    const response = await fetch(`/api/placeholders?mode=${encodeURIComponent(state.mode)}`);
    const result = await response.json();
    if (!response.ok || !result.success) throw new Error(result.detail || "Could not load placeholders");
    state.placeholders = result.placeholders || [];
    renderPlaceholders();
  } catch (error) {
    addLog("warning", "placeholder", error.message);
    renderPlaceholders();
  }
}

async function refreshPlaceholders() {
  el.placeholderSearch.value = "";
  el.placeholderType.value = "all";
  await loadPlaceholders();
  showToast("Placeholder list refreshed", "success");
}

function renderPlaceholders() {
  const query = el.placeholderSearch.value.trim().toLowerCase();
  const type = el.placeholderType.value;
  const rows = state.placeholders.filter((item) => {
    const haystack = [
      item.placeholder,
      item.source_key,
      item.description,
      item.source_html,
      item.report_section,
      item.source_file,
      item.section,
      item.template,
      item.location,
      item.purpose,
      item.content_type,
    ].join(" ").toLowerCase();
    return haystack.includes(query) && (type === "all" || item.content_type === type);
  });
  state.visiblePlaceholders = rows;
  renderPlaceholderStats(rows);
  el.placeholderEmpty.hidden = rows.length > 0;
  el.placeholderBody.innerHTML = rows.map((item, index) => `
    <tr>
      <td><code>${escapeHtml(item.placeholder || "")}</code><small>${escapeHtml(item.source_key || "")}</small></td>
      <td>${escapeHtml(item.description || "-")}</td>
      <td><span class="status-badge ${statusClass(item.content_type)}">${escapeHtml(item.content_type || "-")}</span></td>
      <td>${escapeHtml(item.source_html || item.source_file || item.source_key || "-")}</td>
      <td>${escapeHtml(item.template || item.source_file || "-")}</td>
      <td>${escapeHtml(item.location || item.report_section || item.section || "-")}</td>
      <td>${escapeHtml(item.purpose || "-")}</td>
      <td>
        <div class="placeholder-row-actions">
          <button class="secondary-button" type="button" data-placeholder-edit="${index}">Edit</button>
          <button class="secondary-button danger" type="button" data-placeholder-delete="${index}">Delete</button>
        </div>
      </td>
    </tr>
  `).join("");
  el.placeholderBody.querySelectorAll("[data-placeholder-edit]").forEach((button) => {
    button.addEventListener("click", () => editPlaceholder(Number(button.dataset.placeholderEdit)));
  });
  el.placeholderBody.querySelectorAll("[data-placeholder-delete]").forEach((button) => {
    button.addEventListener("click", () => deletePlaceholder(Number(button.dataset.placeholderDelete)));
  });
}

function resetPlaceholderForm() {
  state.editingPlaceholder = "";
  el.placeholderOriginal.value = "";
  el.placeholderFields.forEach((field) => {
    field.value = field.dataset.placeholderField === "content_type" ? "table" : "";
  });
  if (el.placeholderEditorMode) el.placeholderEditorMode.textContent = "New";
  if (el.placeholderEditor) el.placeholderEditor.open = true;
}

function editPlaceholder(index) {
  const item = state.visiblePlaceholders[index];
  if (!item) return;
  state.editingPlaceholder = item.placeholder || "";
  el.placeholderOriginal.value = state.editingPlaceholder;
  if (el.placeholderEditorMode) el.placeholderEditorMode.textContent = "Edit";
  if (el.placeholderEditor) el.placeholderEditor.open = true;
  el.placeholderFields.forEach((field) => {
    const key = field.dataset.placeholderField;
    field.value = item[key] || "";
  });
  el.placeholderEditor.scrollIntoView({ behavior: "smooth", block: "start" });
}

async function savePlaceholderFromForm(event) {
  event.preventDefault();
  const item = {};
  el.placeholderFields.forEach((field) => {
    if (field.value.trim()) item[field.dataset.placeholderField] = field.value.trim();
  });
  if (!item.placeholder || !item.source_key) {
    showToast("Placeholder and source key are required", "warning");
    return;
  }
  try {
    const response = await fetch("/api/placeholders", {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        mode: state.mode,
        original_placeholder: el.placeholderOriginal.value || "",
        placeholder: item,
      }),
    });
    const result = await response.json();
    if (!response.ok || !result.success) throw new Error(result.detail || "Could not save placeholder");
    addLog("success", "placeholder_yaml", `Saved ${item.placeholder}`);
    showToast("Placeholder saved", "success");
    resetPlaceholderForm();
    await loadPlaceholders();
  } catch (error) {
    addLog("error", "placeholder_yaml", error.message);
    showToast("Could not save placeholder", "error");
  }
}

async function deletePlaceholder(index) {
  const item = state.visiblePlaceholders[index];
  if (!item?.placeholder) return;
  if (!confirm(`Delete ${item.placeholder}?`)) return;
  try {
    const response = await fetch("/api/placeholders", {
      method: "DELETE",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ mode: state.mode, placeholder: item.placeholder }),
    });
    const result = await response.json();
    if (!response.ok || !result.success) throw new Error(result.detail || "Could not delete placeholder");
    addLog("success", "placeholder_yaml", `Deleted ${item.placeholder}`);
    showToast("Placeholder deleted", "success");
    if (state.editingPlaceholder === item.placeholder) resetPlaceholderForm();
    await loadPlaceholders();
  } catch (error) {
    addLog("error", "placeholder_yaml", error.message);
    showToast("Could not delete placeholder", "error");
  }
}

function renderPlaceholderStats(rows) {
  const total = rows.length;
  const tables = rows.filter((item) => item.content_type === "table").length;
  const charts = rows.filter((item) => item.content_type === "chart").length;
  const other = total - tables - charts;
  el.placeholderStats.innerHTML = `
    <div class="history-stat"><span>Total</span><strong>${total}</strong><em>visible placeholders</em></div>
    <div class="history-stat success"><span>Tables</span><strong>${tables}</strong><em>table mappings</em></div>
    <div class="history-stat warning"><span>Charts</span><strong>${charts}</strong><em>chart mappings</em></div>
    <div class="history-stat failed"><span>Other</span><strong>${other}</strong><em>image or text</em></div>
  `;
}

function handlePlaceholderFileChange() {
  const file = el.placeholderFile.files[0];
  el.placeholderFileError.textContent = "";
  el.addNewPlaceholders.disabled = true;
  state.lastPlaceholderScan = null;
  if (!file) {
    el.placeholderFileName.textContent = "No scan file selected";
    el.placeholderFileMeta.textContent = "Waiting for upload";
    return;
  }
  el.placeholderFileName.textContent = file.name;
  el.placeholderFileMeta.textContent = `${formatBytes(file.size)} | Ready`;
  if (!/\.(docx|pdf|html|htm)$/i.test(file.name)) {
    el.placeholderFileError.textContent = "File must be .docx, .pdf, .html, or .htm.";
  }
}

async function scanPlaceholders(event) {
  event.preventDefault();
  const file = el.placeholderFile.files[0];
  if (!file || !/\.(docx|pdf|html|htm)$/i.test(file.name)) {
    el.placeholderFileError.textContent = "Upload a valid scan file first.";
    showToast("Upload a valid scan file first", "warning");
    return;
  }
  el.scanButton.disabled = true;
  el.addNewPlaceholders.disabled = true;
  addLog("info", "placeholder_scan", `Scanning ${file.name}`);
  const formData = new FormData();
  formData.set("mode", state.mode);
  formData.set("template_file", file);
  try {
    const response = await fetch("/api/placeholders/scan", { method: "POST", body: formData });
    const result = await response.json();
    if (!response.ok || !result.success) throw new Error(result.detail || "Placeholder scan failed");
    state.lastPlaceholderScan = result;
    renderPlaceholderScan(result);
    addLog("success", "placeholder_scan", `Found ${result.summary.found} placeholders, ${result.summary.new_in_file} new`);
    showToast("Placeholder scan completed", "success");
  } catch (error) {
    addLog("error", "placeholder_scan", error.message);
    showToast("Placeholder scan failed", "error");
  } finally {
    el.scanButton.disabled = false;
  }
}

function renderPlaceholderScan(result) {
  el.scanResults.hidden = false;
  renderScanRows(el.scanFound, result.found || [], "No placeholders found in file.");
  renderScanRows(el.scanMissing, result.missing_in_file || [], "Every YAML placeholder was found in file.");
  renderScanRows(el.scanNew, result.new_in_file || [], "No new placeholders in file.");
  el.addNewPlaceholders.disabled = !(result.new_in_file || []).length;
}

function renderScanRows(target, rows, emptyText) {
  target.innerHTML = rows.length
    ? rows.map((item) => `
      <tr>
        <td><code>${escapeHtml(item.placeholder)}</code><small>${escapeHtml(item.source_key || "")}</small></td>
        <td>${escapeHtml(item.description || item.content_type || "-")}</td>
        <td><span class="status-badge ${statusClass(item.status)}">${labelStatus(item.status)}</span></td>
      </tr>
    `).join("")
    : `<tr><td class="muted-cell" colspan="3">${emptyText}</td></tr>`;
}

async function addNewPlaceholdersToYaml() {
  const rows = state.lastPlaceholderScan?.new_in_file || [];
  if (!rows.length) return;
  el.addNewPlaceholders.disabled = true;
  try {
    const response = await fetch("/api/placeholders/add", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ mode: state.mode, placeholders: rows }),
    });
    const result = await response.json();
    if (!response.ok || !result.success) throw new Error(result.detail || "Could not add placeholders");
    addLog("success", "placeholder_yaml", `Added ${result.added_count} placeholders to YAML`);
    showToast(`Added ${result.added_count} placeholders to YAML`, "success");
    await loadPlaceholders();
    if (state.lastPlaceholderScan) {
      state.lastPlaceholderScan.new_in_file = [];
      state.lastPlaceholderScan.missing_in_file = state.placeholders.filter(
        (item) => !new Set((state.lastPlaceholderScan.found || []).map((row) => row.placeholder)).has(item.placeholder)
      );
      renderPlaceholderScan(state.lastPlaceholderScan);
    }
  } catch (error) {
    el.addNewPlaceholders.disabled = false;
    addLog("error", "placeholder_yaml", error.message);
    showToast("Could not add placeholders", "error");
  }
}

function renderAiHistoryOptions() {
  if (!el.aiHistorySource) return;
  const selected = el.aiHistorySource.value;
  const rows = state.history.filter((job) => job.mode === "oraclehc" && job.status === "success");
  el.aiHistorySource.innerHTML = `<option value="">Upload new source</option>` + rows.map((job) => {
    const label = `${job.output_file_name || job.source_package_name || job.id} | ${formatDate(job.created_at)}`;
    return `<option value="${job.id}">${escapeHtml(label)}</option>`;
  }).join("");
  if (rows.some((job) => job.id === selected)) el.aiHistorySource.value = selected;
}

function handleAiSourceChoice() {
  const usingHistory = Boolean(el.aiHistorySource.value);
  el.aiSourceFile.disabled = usingHistory;
  el.aiSourceError.textContent = "";
  if (usingHistory) {
    const job = state.history.find((item) => item.id === el.aiHistorySource.value);
    el.aiSourceName.textContent = job?.source_package_name || "History source package";
    el.aiSourceMeta.textContent = "Using selected History package";
  } else {
    handleAiSourceFileChange();
  }
}

function handleAiSourceFileChange() {
  const file = el.aiSourceFile.files[0];
  el.aiSourceError.textContent = "";
  if (!file) {
    el.aiSourceName.textContent = "No source selected";
    el.aiSourceMeta.textContent = "Waiting for upload or History selection";
    return;
  }
  el.aiSourceName.textContent = file.name;
  el.aiSourceMeta.textContent = `${formatBytes(file.size)} | Ready`;
  if (!/\.(zip|docx|html|htm)$/i.test(file.name)) {
    el.aiSourceError.textContent = "File must be .zip, .docx, .html, or .htm.";
  }
}

async function generateAiReview(event) {
  event.preventDefault();
  if (state.isAiReviewing) return;
  const historyJobId = el.aiHistorySource.value;
  const file = el.aiSourceFile.files[0];
  el.aiSourceError.textContent = "";
  if (!historyJobId && (!file || !/\.(zip|docx|html|htm)$/i.test(file.name))) {
    el.aiSourceError.textContent = "Select a History package or upload a valid source file.";
    showToast("Select source for AI Review", "warning");
    return;
  }

  state.isAiReviewing = true;
  setAiReviewBusy(true, "Parsing source and generating review...");
  runAiScanSteps();
  addLog("info", "ai_review", "Parsing source package before calling AI");
  const formData = new FormData();
  formData.set("review_type", el.aiReviewType.value);
  formData.set("output_style", el.aiOutputStyle.value);
  formData.set("history_job_id", historyJobId);
  if (!historyJobId && file) formData.set("source_file", file);

  try {
    const response = await fetch("/api/ai-review/generate", { method: "POST", body: formData });
    const result = await response.json();
    if (!response.ok || !result.success) throw new Error(result.detail || "AI Review failed");
    state.aiReviewRows = result.rows || [];
    renderAiReviewTable();
    const providerLabel = result.used_ai ? result.provider : "local rules";
    el.aiReviewStatus.textContent = `Generated from ${result.source_name || "source"} using ${providerLabel}.`;
    completeAiScanSteps();
    addLog("success", "ai_review", `Generated ${state.aiReviewRows.length} review rows`);
    showToast("AI Review generated", "success");
  } catch (error) {
    addLog("error", "ai_review", error.message);
    el.aiReviewStatus.textContent = error.message;
    failAiScanSteps();
    showToast("AI Review failed", "error");
  } finally {
    state.isAiReviewing = false;
    setAiReviewBusy(false);
  }
}

function setAiReviewBusy(isBusy, status = "Ready") {
  el.aiGenerate.disabled = isBusy;
  el.aiGenerate.classList.toggle("is-loading", isBusy);
  const buttonLabel = isBusy ? "Generating..." : "Scan & Generate AI Review";
  const labelSpan = el.aiGenerate.querySelector("span:last-child");
  if (labelSpan) {
    labelSpan.replaceChildren(document.createTextNode(buttonLabel));
  } else {
    el.aiGenerate.textContent = buttonLabel;
  }
  if (isBusy) el.aiReviewStatus.textContent = status;
  if (el.aiScanPanel) el.aiScanPanel.classList.toggle("is-running", isBusy);
  if (el.aiLiveIndicator) el.aiLiveIndicator.textContent = isBusy ? "Analyzing" : "Idle";
}

function renderAiReviewTable() {
  el.aiReviewResult.hidden = state.aiReviewRows.length === 0;
  el.aiReviewBody.innerHTML = state.aiReviewRows.map((row, index) => `
    <tr>
      <td><span class="ai-status-badge ${aiReviewStatusClass(row)}">${escapeHtml(aiReviewStatusLabel(row))}</span><span contenteditable="true" data-ai-cell="${index}:section">${escapeHtml(row.section)}</span></td>
      <td contenteditable="true" data-ai-cell="${index}:assessment">${escapeHtml(row.assessment)}</td>
      <td contenteditable="true" data-ai-cell="${index}:recommendation">${escapeHtml(row.recommendation)}</td>
    </tr>
  `).join("");
  el.aiReviewBody.querySelectorAll("[data-ai-cell]").forEach((cell) => {
    cell.addEventListener("input", () => {
      const [index, key] = cell.dataset.aiCell.split(":");
      if (state.aiReviewRows[Number(index)]) state.aiReviewRows[Number(index)][key] = cell.textContent.trim();
    });
  });
  const hasRows = state.aiReviewRows.length > 0;
  el.aiCopy.disabled = !hasRows;
  el.aiExportDocx.disabled = !hasRows;
  el.aiExportXlsx.disabled = !hasRows;
}

function runAiScanSteps() {
  if (!el.aiStepList) return;
  el.aiStepList.querySelectorAll("[data-ai-step]").forEach((step, index) => {
    step.classList.remove("is-done", "is-active", "is-error");
    if (index === 0) step.classList.add("is-active");
  });
  [1, 2, 3, 4, 5].forEach((stepIndex) => {
    setTimeout(() => {
      if (!state.isAiReviewing || !el.aiStepList) return;
      el.aiStepList.querySelectorAll("[data-ai-step]").forEach((step, index) => {
        step.classList.toggle("is-done", index < stepIndex);
        step.classList.toggle("is-active", index === stepIndex);
      });
    }, stepIndex * 520);
  });
}

function completeAiScanSteps() {
  if (!el.aiStepList) return;
  el.aiStepList.querySelectorAll("[data-ai-step]").forEach((step) => {
    step.classList.remove("is-active", "is-error");
    step.classList.add("is-done");
  });
  const last = el.aiStepList.querySelector("[data-ai-step='6']");
  if (last) last.classList.add("is-active");
  if (el.aiLiveIndicator) el.aiLiveIndicator.textContent = "Completed";
}

function failAiScanSteps() {
  if (!el.aiStepList) return;
  const active = el.aiStepList.querySelector(".is-active") || el.aiStepList.querySelector("[data-ai-step='4']");
  if (active) active.classList.add("is-error");
  if (el.aiLiveIndicator) el.aiLiveIndicator.textContent = "Needs review";
}

function aiReviewStatusLabel(row) {
  const text = `${row.assessment || ""} ${row.recommendation || ""}`.toLowerCase();
  if (text.includes("không đủ") || text.includes("missing") || text.includes("no data")) return "Missing Data";
  if ((row.recommendation || "").trim() && !(row.recommendation || "").toLowerCase().includes("không cần")) return "Recommendation";
  if (text.includes("cần") || text.includes("review") || text.includes("cao")) return "Review";
  return "OK";
}

function aiReviewStatusClass(row) {
  return aiReviewStatusLabel(row).toLowerCase().replace(/\s+/g, "-");
}

async function copyAiReviewTable() {
  const text = [["Mục", "Đánh giá", "Khuyến nghị"], ...state.aiReviewRows.map((row) => [
    row.section,
    row.assessment,
    row.recommendation,
  ])].map((row) => row.join("\t")).join("\n");
  await navigator.clipboard.writeText(text);
  showToast("AI Review table copied", "success");
}

async function exportAiReview(type) {
  if (!state.aiReviewRows.length) return;
  try {
    const response = await fetch(`/api/ai-review/export/${type}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ rows: state.aiReviewRows }),
    });
    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(errorText || "Export failed");
    }
    const blob = await response.blob();
    downloadBlob(blob, downloadNameFromResponse(response) || `ai_review.${type}`);
    showToast(`AI Review ${type.toUpperCase()} exported`, "success");
  } catch (error) {
    showToast(error.message, "error");
  }
}

function renderHistory() {
  const search = el.historySearch.value.trim().toLowerCase();
  const mode = el.historyMode.value;
  const status = el.historyStatus.value;
  const rows = state.history.filter((job) => {
    return (job.output_file_name || "").toLowerCase().includes(search)
      && (mode === "all" || job.mode === mode)
      && (status === "all" || job.status === status);
  });
  renderHistoryStats(rows);
  el.historyEmpty.hidden = rows.length > 0;
  el.historyBody.innerHTML = rows.map((job) => `
    <tr>
      <td><strong>${escapeHtml(job.output_file_name || "")}</strong></td>
      <td>${escapeHtml(job.company_name || "Unassigned")}</td>
      <td>${escapeHtml(job.report_type || "Unclassified")}</td>
      <td>${escapeHtml(reportPeriod(job))}</td>
      <td>${escapeHtml(job.year || "-")}</td>
      <td>${modeConfig[job.mode]?.label || job.mode}</td>
      <td>${escapeHtml(job.template_file_name || "-")}</td>
      <td>${escapeHtml(job.source_package_name || "-")}</td>
      <td><span class="status-badge ${statusClass(job.status)}">${labelStatus(job.status)}</span></td>
      <td>
        <span class="history-progress"><span style="width: ${Math.max(0, Math.min(100, Number(job.progress || 0)))}%"></span></span>
        <small>${job.progress || 0}%</small>
      </td>
      <td>${formatDate(job.created_at)}</td>
      <td>${job.duration_seconds ? `${Number(job.duration_seconds).toFixed(1)}s` : "-"}</td>
      <td>
        <div class="table-actions">
          <button type="button" data-history-download="${job.id}" ${job.status === "success" ? "" : "disabled"}>Download</button>
          <button type="button" data-history-assign="${job.id}">Assign/Edit</button>
          <button type="button" data-history-logs="${job.id}">View Logs</button>
          <button type="button" data-history-delete="${job.id}">Delete</button>
        </div>
      </td>
    </tr>
  `).join("");
  bindHistoryActions();
}

function renderHistoryStats(rows) {
  if (!el.historyStats) return;
  const total = rows.length;
  const success = rows.filter((job) => job.status === "success").length;
  const failed = rows.filter((job) => job.status === "failed").length;
  const processing = rows.filter((job) => job.status === "processing").length;
  const completed = total ? Math.round((success / total) * 100) : 0;
  el.historyStats.innerHTML = `
    <div class="history-stat">
      <span>Total Reports</span>
      <strong>${total}</strong>
      <em>${completed}% success rate</em>
    </div>
    <div class="history-stat success">
      <span>Success</span>
      <strong>${success}</strong>
      <em>Ready to download</em>
    </div>
    <div class="history-stat warning">
      <span>Processing</span>
      <strong>${processing}</strong>
      <em>Running jobs</em>
    </div>
    <div class="history-stat failed">
      <span>Failed</span>
      <strong>${failed}</strong>
      <em>Needs review</em>
    </div>
  `;
}

function bindHistoryActions() {
  document.querySelectorAll("[data-history-download]").forEach((button) => {
    button.addEventListener("click", () => { window.location.href = `/api/report/download/${button.dataset.historyDownload}`; });
  });
  document.querySelectorAll("[data-history-assign]").forEach((button) => {
    button.addEventListener("click", () => openAssignModal(button.dataset.historyAssign));
  });
  document.querySelectorAll("[data-history-logs]").forEach((button) => {
    button.addEventListener("click", async () => {
      const response = await fetch(`/api/history/${button.dataset.historyLogs}/logs`);
      const result = await response.json();
      state.logs = (result.logs || []).map((log) => ({ timestamp: log.timestamp, level: log.level, step: log.step || "", message: log.message }));
      renderLogs();
      showTab("tool");
    });
  });
  document.querySelectorAll("[data-history-delete]").forEach((button) => {
    button.addEventListener("click", async () => {
      await fetch(`/api/history/${button.dataset.historyDelete}`, { method: "DELETE" });
      await loadHistory();
      await loadCompanies();
      if (state.activeCompanyId) await loadCompanyReports();
    });
  });
}

function openAssignModal(jobId) {
  const job = state.history.find((item) => item.id === jobId) || {};
  state.pendingAssignJobId = jobId;
  el.assignForm.reset();
  el.assignJobId.value = jobId;
  el.assignSummary.innerHTML = `
    <div><span>File Name</span><strong>${escapeHtml(job.output_file_name || "-")}</strong></div>
    <div><span>Source Package</span><strong>${escapeHtml(job.source_package_name || "-")}</strong></div>
    <div><span>Created At</span><strong>${formatDate(job.created_at)}</strong></div>
  `;
  renderAssignCompanyOptions();
  setAssignValue("company_id", job.company_id || "");
  setAssignValue("report_type", job.report_type || "Unclassified");
  setAssignValue("period_type", job.period_type || "Quarter");
  setAssignValue("quarter", job.quarter || "");
  setAssignValue("month", job.month || "");
  setAssignValue("year", job.year || new Date().getFullYear());
  setAssignValue("report_status", job.report_status || "Generated");
  setAssignValue("note", job.note || "");
  el.assignModal.hidden = false;
  el.assignModal.setAttribute("aria-hidden", "false");
}

function closeAssignModal() {
  el.assignModal.hidden = true;
  el.assignModal.setAttribute("aria-hidden", "true");
  state.pendingAssignJobId = "";
}

function renderAssignCompanyOptions() {
  const select = assignField("company_id");
  if (!select) return;
  const selected = select.value;
  select.innerHTML = `<option value="">Unassigned</option>` + state.companies
    .map((company) => `<option value="${company.id}">${escapeHtml(company.company_name)} (${escapeHtml(company.short_name)})</option>`)
    .join("");
  select.value = selected;
}

async function saveAssignment(event) {
  event.preventDefault();
  const jobId = el.assignJobId.value;
  const payload = {};
  el.assignFields.forEach((field) => {
    payload[field.dataset.assignField] = field.value.trim();
  });
  try {
    const response = await fetch(`/api/history/${jobId}/assignment`, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    const result = await response.json();
    if (!response.ok || !result.success) throw new Error(result.detail || "Could not save assignment");
    closeAssignModal();
    await loadHistory();
    await loadCompanies();
    if (state.activeCompanyId) await loadCompanyReports();
    showToast("Report assignment saved", "success");
  } catch (error) {
    showToast(error.message, "error");
  }
}

function assignField(name) {
  return Array.from(el.assignFields).find((field) => field.dataset.assignField === name);
}

function setAssignValue(name, value) {
  const field = assignField(name);
  if (field) field.value = value ?? "";
}

function reportPeriod(job) {
  if (job.period_type === "Month" && job.month) return `Month ${job.month}`;
  if (job.period_type === "Custom") return "Custom";
  return job.quarter || job.month || "-";
}

function addLog(level, step, message) {
  state.logs.push({ timestamp: new Date().toISOString().slice(0, 19), level, step, message });
  renderLogs();
}

function renderLogs() {
  const rows = state.logs.filter((log) => state.logFilter === "all" || log.level === state.logFilter);
  el.logList.innerHTML = rows.length
    ? rows.map((log) => `
      <div class="log-row ${log.level}">
        <span class="log-icon">${log.level}</span>
        <time>${formatDate(log.timestamp)}</time>
        <span><strong>${escapeHtml(log.step || "-")}</strong> ${escapeHtml(log.message)}</span>
      </div>
    `).join("")
    : `<div class="empty-inline">No logs yet.</div>`;
  el.logList.scrollTop = el.logList.scrollHeight;
}

function clearLogs() {
  state.logs = [];
  renderLogs();
}

async function copyLogs() {
  const text = state.logs.map((log) => `[${log.timestamp}] [${log.level}] [${log.step}] ${log.message}`).join("\n");
  await navigator.clipboard.writeText(text);
  showToast("Logs copied", "success");
}

function downloadLogs() {
  const text = state.logs.map((log) => `[${log.timestamp}] [${log.level}] [${log.step}] ${log.message}`).join("\n");
  const blob = new Blob([text], { type: "text/plain" });
  downloadBlob(blob, "runtime.log");
}

function resetForm() {
  clearTimeout(state.pollTimer);
  el.form.reset();
  clearFieldErrors();
  state.isGenerating = false;
  state.outputTouched = false;
  state.activeJobId = null;
  state.lastDownloadJobId = null;
  el.templateName.textContent = "No template selected";
  el.sourceName.textContent = "No source data selected";
  el.templateMeta.textContent = "Waiting for upload";
  el.sourceMeta.textContent = "Waiting for upload";
  el.resultCard.hidden = true;
  setLoadingState(false, 0, "Ready", "Waiting for input");
  applyMode(state.mode, { forceOutput: true });
  updateGenerateAvailability();
}

function showResult(fileName, failed = false, error = "") {
  el.resultCard.hidden = false;
  el.resultTitle.textContent = failed ? "Generation failed" : "Report is ready";
  el.resultFilename.textContent = failed ? error : fileName;
  el.resultCard.classList.toggle("error", failed);
}

function setGenerating(value) {
  el.submitButton.textContent = value ? "Generating..." : "Generate Report";
  updateGenerateAvailability();
}

function downloadNameFromResponse(response) {
  const disposition = response.headers.get("content-disposition") || "";
  const match = disposition.match(/filename\*=UTF-8''([^;]+)|filename="?([^";]+)"?/i);
  if (!match) return "";
  return decodeURIComponent(match[1] || match[2] || "");
}

function stripDocxExtension(fileName) {
  return fileName.replace(/\.docx$/i, "");
}

function showTab(tabName) {
  el.navTabs.forEach((tab) => tab.classList.toggle("is-active", tab.dataset.tabTarget === tabName));
  el.pages.forEach((page) => page.classList.toggle("is-active", page.dataset.tabPage === tabName));
  requestAnimationFrame(updateScrollActions);
  if (tabName === "history") loadHistory();
  if (tabName === "ai-review") renderAiHistoryOptions();
  if (tabName === "company") loadCompanies();
}

function fileInput(type) {
  return type === "template" ? el.templateInput : el.sourceInput;
}

function clearFieldErrors() {
  el.templateError.textContent = "";
  el.sourceError.textContent = "";
  el.outputError.textContent = "";
}

function detectSourceType() {
  return state.mode === "sqlhealthcheck" ? "Detected source type: SQLHealthcheck CSV" : "Detected source type: OracleHC HTML";
}

function normalizeDocxName(value) {
  const clean = value.trim() || modeConfig[state.mode].outputName;
  return clean.toLowerCase().endsWith(".docx") ? clean : `${clean}.docx`;
}

function formatBytes(bytes) {
  if (!bytes) return "0 B";
  const units = ["B", "KB", "MB", "GB"];
  const index = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
  return `${(bytes / Math.pow(1024, index)).toFixed(index ? 1 : 0)} ${units[index]}`;
}

function formatDate(value) {
  if (!value) return "-";
  return value.replace("T", " ");
}

function statusClass(status) {
  return {
    success: "success",
    failed: "failed",
    processing: "processing",
    Draft: "processing",
    Generated: "success",
    Reviewed: "ready",
    Sent: "success",
    table: "success",
    chart: "processing",
    image: "ready",
    text: "ready",
    mapped: "success",
    missing_mapping: "failed",
    duplicate: "processing",
    missing_in_file: "failed",
    new: "processing",
    unsupported: "failed",
  }[status] || "ready";
}

function labelStatus(status) {
  return String(status || "").replace(/_/g, " ").replace(/\b\w/g, (char) => char.toUpperCase());
}

function capitalize(value) {
  return String(value || "").charAt(0).toUpperCase() + String(value || "").slice(1);
}

function showToast(message, type = "info") {
  const toast = document.createElement("div");
  toast.className = `toast ${type}`;
  toast.textContent = message;
  el.toastRegion.appendChild(toast);
  setTimeout(() => {
    toast.classList.add("is-leaving");
    setTimeout(() => toast.remove(), 220);
  }, 2600);
}

function downloadBlob(blob, fileName) {
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  document.body.appendChild(anchor);
  anchor.click();
  anchor.remove();
  URL.revokeObjectURL(url);
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

init();
