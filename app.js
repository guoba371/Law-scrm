"use strict";

function createDeferredRunner(run, options = {}) {
  const schedule = options.schedule || (callback => setTimeout(callback, 0));
  const cancel = options.cancel || (handle => clearTimeout(handle));

  let pendingHandle = null;
  let pendingArgs = [];

  return function deferredRunner(...args) {
    pendingArgs = args;
    if (pendingHandle !== null) {
      cancel(pendingHandle);
    }
    pendingHandle = schedule(() => {
      pendingHandle = null;
      run(...pendingArgs);
    });
  };
}

const CUSTOMER_GRADE = {
  closed: "成交客户",
  valuable: "有价值潜在客户",
  unclosed: "潜在客户未成交",
  unknown: "未识别"
};

const TRACKING_FIELDS = [
  { key: "caseName", label: "案件名称", type: "text" },
  { key: "filingDate", label: "立案时间", type: "date" },
  { key: "arbitrationHearingDate", label: "仲裁开庭时间", type: "date" },
  { key: "firstTrialDate", label: "一审开庭时间", type: "date" },
  { key: "secondTrialDate", label: "二审开庭时间", type: "date" },
  {
    key: "civilResult",
    label: "民事结果",
    type: "select",
    options: ["", "仲裁裁决", "法院判决"]
  },
  { key: "detentionDate", label: "拘留时间", type: "date" },
  { key: "procuratorateTransferDate", label: "移送检察院时间", type: "date" },
  { key: "criminalCourtDate", label: "法院开庭", type: "date" },
  {
    key: "criminalResult",
    label: "刑事结果",
    type: "select",
    options: ["", "取保候审", "不予逮捕", "法院判决"]
  },
  { key: "caseClosed", label: "案件是否结案", type: "checkbox" },
  { key: "followUpNotes", label: "补充跟踪", type: "textarea" }
];

const CASE_TYPE_HEADER = { label: "案件类型" };
const HEADER_KEYWORD_RE = /客户|姓名|来源|电话|案件|备注|金额|日期|时间|律师|助理|跟进|合同|代理|结果|开庭|立案|发生地|案由/;

const DEFAULT_SOURCE_FILE = "2026年每月客户统计.xlsx";
const DEFAULT_THEME = "light";
const WORKSPACE_STORAGE_KEY = "law-scrm-workspace-v1";
const DEAL_AMOUNT_HEADER = "成交金额";
const COLUMN_WIDTH_MIN = 80;
const COLUMN_WIDTH_MAX = 520;
const ROW_HEIGHT_MIN = 32;
const ROW_HEIGHT_MAX = 220;
const OVERDUE_CASE_MONTHS = 3;
const appState = {
  fileName: "",
  workbook: null,
  customers: [],
  closedCustomers: [],
  sourceHeaders: [],
  columnWidths: [],
  rowHeights: [],
  caseColumnWidths: [],
  caseRowHeights: [],
  trackingById: new Map(),
  activeView: "all",
  uploadProgress: 0,
  uploadMessage: "准备上传",
  theme: loadStoredTheme(),
  rowSpacing: 42,
  columnWidth: 148,
  activeResize: null,
  activeModalEditor: null,
  overdueNoticeShownFor: ""
};

const scheduleExportLengthRefresh = createDeferredRunner(() => {
  document.documentElement.dataset.exportLength = String(appState.customers.length);
});

const scheduleWorkspacePersist = createDeferredRunner(() => {
  persistWorkspace();
});

window.lawScrmDebug = {
  state: appState,
  buildManagedWorkbook
};

const els = {
  fileInput: document.getElementById("fileInput"),
  downloadBtn: document.getElementById("downloadBtn"),
  themeToggleBtn: document.getElementById("themeToggleBtn"),
  totalCustomerCount: document.getElementById("totalCustomerCount"),
  openCaseCount: document.getElementById("openCaseCount"),
  closedCustomerCount: document.getElementById("closedCustomerCount"),
  completedCaseCount: document.getElementById("completedCaseCount"),
  sourceHint: document.getElementById("sourceHint"),
  allCustomersHint: document.getElementById("allCustomersHint"),
  allCustomersInlineStats: document.getElementById("allCustomersInlineStats"),
  emptyState: document.getElementById("emptyState"),
  caseTableToolbar: document.getElementById("caseTableToolbar"),
  caseTableWrap: document.getElementById("caseTableWrap"),
  caseHead: document.getElementById("caseHead"),
  caseBody: document.getElementById("caseBody"),
  allCustomersPage: document.getElementById("allCustomersPage"),
  closedCustomersPage: document.getElementById("closedCustomersPage"),
  allCustomersTabBtn: document.getElementById("allCustomersTabBtn"),
  closedCustomersTabBtn: document.getElementById("closedCustomersTabBtn"),
  allCustomersPanel: document.getElementById("allCustomersPanel"),
  customerTable: document.getElementById("customerTable"),
  customerHead: document.getElementById("customerHead"),
  customerBody: document.getElementById("customerBody"),
  rowSpacingRange: document.getElementById("rowSpacingRange"),
  columnWidthRange: document.getElementById("columnWidthRange"),
  addCustomerRowBtn: document.getElementById("addCustomerRowBtn"),
  addCustomerColumnBtn: document.getElementById("addCustomerColumnBtn"),
  addCaseRowBtn: document.getElementById("addCaseRowBtn"),
  addCaseColumnBtn: document.getElementById("addCaseColumnBtn"),
  uploadStatus: document.getElementById("uploadStatus"),
  uploadProgressBar: document.getElementById("uploadProgressBar"),
  uploadProgressText: document.getElementById("uploadProgressText"),
  noticeBtn: document.getElementById("noticeBtn"),
  noticeBadge: document.getElementById("noticeBadge"),
  noticeModal: document.getElementById("noticeModal"),
  noticeModalMeta: document.getElementById("noticeModalMeta"),
  noticeModalBody: document.getElementById("noticeModalBody"),
  noticeModalCloseBtn: document.getElementById("noticeModalCloseBtn"),
  editorModal: document.getElementById("editorModal"),
  editorModalMeta: document.getElementById("editorModalMeta"),
  editorModalTextarea: document.getElementById("editorModalTextarea"),
  editorModalCloseBtn: document.getElementById("editorModalCloseBtn"),
  editorModalSaveBtn: document.getElementById("editorModalSaveBtn")
};

els.fileInput.addEventListener("change", async event => {
  const file = event.target.files?.[0];
  if (!file) return;
  updateUploadProgress(8, `正在读取 ${file.name}`);
  const buffer = await file.arrayBuffer();
  updateUploadProgress(34, "正在解析 Excel");
  await loadWorkbook(file.name, buffer);
  updateUploadProgress(100, "导入完成");
  window.setTimeout(() => {
    appState.uploadProgress = 0;
    appState.uploadMessage = "准备上传";
    syncUploadProgress();
  }, 900);
});

els.downloadBtn.addEventListener("click", () => {
  const bytes = buildManagedWorkbook();
  const blob = new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
  const stamp = new Date().toISOString().slice(0, 10);
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = `客户案件管理-${stamp}.xlsx`;
  link.click();
  URL.revokeObjectURL(link.href);
});

els.themeToggleBtn.addEventListener("click", () => {
  appState.theme = appState.theme === "dark" ? "light" : "dark";
  persistTheme(appState.theme);
  applyTheme();
});

els.allCustomersTabBtn.addEventListener("click", () => {
  appState.activeView = "all";
  render();
});

els.closedCustomersTabBtn.addEventListener("click", () => {
  appState.activeView = "closed";
  render();
});

els.rowSpacingRange.addEventListener("input", event => {
  appState.rowSpacing = Number(event.currentTarget.value);
  appState.rowHeights = appState.customers.map(() => appState.rowSpacing);
  appState.caseRowHeights = appState.closedCustomers.map(() => appState.rowSpacing);
  syncTableMetrics();
  scheduleWorkspacePersist();
});

els.columnWidthRange.addEventListener("input", event => {
  appState.columnWidth = Number(event.currentTarget.value);
  appState.columnWidths = Array.from(
    { length: appState.sourceHeaders.length + 1 },
    () => appState.columnWidth
  );
  appState.caseColumnWidths = Array.from(
    { length: getCaseColumnCount() },
    () => appState.columnWidth
  );
  syncTableMetrics();
  scheduleWorkspacePersist();
});

els.addCustomerRowBtn.addEventListener("click", () => {
  addCustomerRow();
});

els.addCustomerColumnBtn.addEventListener("click", () => {
  addCustomerColumn();
});

els.addCaseRowBtn?.addEventListener("click", () => {
  addCaseRow();
});

els.addCaseColumnBtn?.addEventListener("click", () => {
  addCaseColumn();
});

els.noticeBtn?.addEventListener("click", () => {
  openNoticeModal(false);
});
els.noticeModalCloseBtn?.addEventListener("click", closeNoticeModal);
els.noticeModal?.addEventListener("click", event => {
  if (event.target instanceof HTMLElement && event.target.dataset.closeNotice === "true") {
    closeNoticeModal();
  }
});
els.editorModalCloseBtn?.addEventListener("click", closeEditorModal);
els.editorModalSaveBtn?.addEventListener("click", saveEditorModal);
els.editorModal?.addEventListener("click", event => {
  if (event.target instanceof HTMLElement && event.target.dataset.closeEditor === "true") {
    closeEditorModal();
  }
});
window.addEventListener("keydown", event => {
  if (event.key === "Escape" && !els.noticeModal?.classList.contains("hidden")) {
    closeNoticeModal();
    return;
  }
  if (event.key === "Escape" && appState.activeModalEditor) {
    closeEditorModal();
  }
});
document.addEventListener("copy", handleTableCopy);

applyTheme();
syncUploadProgress();
syncTableMetrics();
if (!restoreStoredWorkspace()) {
  bootstrapDefaultWorkbook();
}

async function bootstrapDefaultWorkbook() {
  try {
    const response = await fetch(encodeURI(DEFAULT_SOURCE_FILE));
    if (!response.ok) return;
    await loadWorkbook(DEFAULT_SOURCE_FILE, await response.arrayBuffer());
  } catch {
    // Opening the HTML directly will not always allow local fetch. Manual upload still works.
  }
}

async function loadWorkbook(fileName, buffer) {
  appState.fileName = fileName;
  const workbook = await parseXlsx(buffer);
  updateUploadProgress(66, "正在整理客户字段");
  appState.workbook = workbook;
  const imported = extractCustomers(workbook);
  appState.customers = imported.customers;
  appState.sourceHeaders = ensureManagedTailHeaders(imported.headers);
  ensureCustomerFields(appState.customers);
  resetCustomerMetrics();
  syncClosedCustomers();
  seedTrackingRows();
  updateUploadProgress(88, "正在生成管理页面");
  render();
}

function extractCustomers(workbook) {
  const preferredSheet = workbook.sheets.find(sheet => /2026年每月客户情况统计/.test(sheet.name));
  const customerSheets = preferredSheet
    ? [preferredSheet]
    : workbook.sheets.filter(sheet => /客户/.test(sheet.name) && !/报销/.test(sheet.name));
  const customers = [];
  let canonicalHeaders = [];

  for (const sheet of customerSheets) {
    const headerInfo = findHeaderRow(sheet.rows);
    if (!headerInfo) continue;
    const headers = normalizeHeaders(headerInfo.headers);
    if (headers.length > canonicalHeaders.length) canonicalHeaders = headers;

    for (const row of sheet.rows.slice(headerInfo.index + 1)) {
      const values = headers.map((_, index) => row.cells[index]?.value ?? "");
      if (values.every(value => String(value).trim() === "")) continue;
      const gradeKey = inferGrade(row);
      const data = {};
      headers.forEach((header, index) => {
        data[header] = values[index] ?? "";
      });
      customers.push({
        id: `${sheet.name}-${row.number}-${customers.length}`,
        sheetName: sheet.name,
        rowNumber: row.number,
        gradeKey,
        grade: CUSTOMER_GRADE[gradeKey],
        data,
        values
      });
    }
  }

  return {
    headers: uniqueHeaders(canonicalHeaders, customers),
    customers
  };
}

function getRawCustomerRows() {
  return appState.customers.map(customer =>
    appState.sourceHeaders.map(header => customer.data[header] ?? "")
  );
}

function ensureManagedTailHeaders(headers) {
  const nextHeaders = headers.filter(header => header !== DEAL_AMOUNT_HEADER);
  nextHeaders.push(DEAL_AMOUNT_HEADER);
  return nextHeaders;
}

function ensureCustomerFields(customers) {
  for (const customer of customers) {
    for (const header of appState.sourceHeaders) {
      if (!(header in customer.data)) {
        customer.data[header] = "";
      }
    }
  }
}

function findHeaderRow(rows) {
  const candidates = rows
    .map((row, index) => {
      const values = getRowTextValues(row);
      const uniqueCount = new Set(values).size;
      const keywordHits = countHeaderKeywords(values);
      const repetitivePenalty = Math.max(0, values.length - uniqueCount);
      const numericLikeCount = values.filter(isMostlyNumeric).length;
      const score =
        uniqueCount * 2 +
        keywordHits * 6 +
        (values.length >= 4 ? 2 : 0) -
        repetitivePenalty -
        numericLikeCount * 2;
      return { row, index, score, values, uniqueCount, keywordHits };
    })
    .filter(item => item.score >= 3)
    .sort((a, b) => b.score - a.score || a.index - b.index);

  const picked = candidates[0];
  if (!picked) return null;
  const headerRows = collectHeaderRows(rows, picked.index);
  const lastCell = Math.max(
    ...headerRows.flatMap(item => item.row.cells.map(cell => cell.index)),
    0
  );
  const headers = Array.from({ length: lastCell + 1 }, (_, index) =>
    buildHeaderLabel(headerRows.map(item => item.row), index)
  );
  return { index: headerRows[headerRows.length - 1]?.index ?? picked.index, headers };
}

function normalizeHeaders(rawHeaders) {
  const seen = new Map();
  return rawHeaders.map((header, index) => {
    const fallback = `${indexToColumnName(index)}列`;
    const clean = String(header || "").trim() || fallback;
    const count = seen.get(clean) || 0;
    seen.set(clean, count + 1);
    return count ? `${clean}_${count + 1}` : clean;
  });
}

function buildHeaderLabel(headerRows, columnIndex) {
  const parts = [];
  for (const row of headerRows) {
    const value = String(row.cells[columnIndex]?.value || "").trim();
    if (!value) continue;
    if (parts[parts.length - 1] === value) continue;
    parts.push(value);
  }
  return parts.join(" / ");
}

function collectHeaderRows(rows, pickedIndex) {
  const collected = [{ row: rows[pickedIndex], index: pickedIndex }];

  for (let index = pickedIndex - 1; index >= Math.max(0, pickedIndex - 2); index--) {
    if (!isStructuralHeaderRow(rows[index])) break;
    collected.unshift({ row: rows[index], index });
  }

  for (
    let index = pickedIndex + 1;
    index < Math.min(rows.length, pickedIndex + 3);
    index++
  ) {
    if (!isFieldHeaderRow(rows[index])) break;
    collected.push({ row: rows[index], index });
  }

  return collected;
}

function getRowTextValues(row) {
  return row.cells
    .map(cell => String(cell?.value || "").trim())
    .filter(Boolean);
}

function countHeaderKeywords(values) {
  return values.filter(value => HEADER_KEYWORD_RE.test(value)).length;
}

function isMostlyNumeric(value) {
  return /^\d+([./-]\d+)*$/.test(String(value).trim());
}

function isStructuralHeaderRow(row) {
  const values = getRowTextValues(row);
  if (!values.length) return false;
  const uniqueCount = new Set(values).size;
  const keywordHits = countHeaderKeywords(values);
  return keywordHits >= 1 || uniqueCount <= Math.max(3, Math.floor(values.length / 3));
}

function isFieldHeaderRow(row) {
  const values = getRowTextValues(row);
  if (!values.length) return false;
  const keywordHits = countHeaderKeywords(values);
  const numericRatio = values.filter(isMostlyNumeric).length / values.length;
  const longTextCount = values.filter(value => value.length >= 20).length;
  return keywordHits >= 2 && numericRatio < 0.4 && longTextCount <= Math.ceil(values.length / 2);
}

function uniqueHeaders(seedHeaders, customers) {
  const seen = new Set();
  const headers = [];
  for (const header of seedHeaders) {
    const clean = String(header || "").trim();
    if (clean && !seen.has(clean)) {
      seen.add(clean);
      headers.push(clean);
    }
  }
  for (const customer of customers) {
    for (const header of Object.keys(customer.data)) {
      if (header && !seen.has(header)) {
        seen.add(header);
        headers.push(header);
      }
    }
  }
  return headers;
}

function inferGrade(row) {
  const fills = row.cells.map(cell => normalizeColor(cell.fillColor)).filter(Boolean);
  if (fills.some(color => color.family === "red")) return "closed";
  if (fills.some(color => color.family === "yellow")) return "valuable";
  if (fills.some(color => color.family === "pink")) return "unclosed";
  return "unknown";
}

function normalizeColor(color) {
  if (!color) return null;
  let hex = color.replace(/^#/, "").toUpperCase();
  if (hex.length === 8) hex = hex.slice(2);
  if (hex.length !== 6) return null;
  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const family =
    r > 220 && g < 130 && b < 130
      ? "red"
      : r > 220 && g > 180 && b < 180
        ? "yellow"
        : r > 220 && g >= 130 && b >= 130 && max - min < 120
          ? "pink"
          : null;
  return family ? { hex, family } : null;
}

function seedTrackingRows() {
  for (const customer of appState.closedCustomers) {
    if (!appState.trackingById.has(customer.id)) {
      const caseName = pickFirst(customer.data, ["案件", "案由", "委托事项", "咨询内容"]) || "";
      appState.trackingById.set(customer.id, {
        caseName,
        filingDate: "",
        caseType: "",
        arbitrationHearingDate: "",
        firstTrialDate: "",
        secondTrialDate: "",
        civilResult: "",
        detentionDate: "",
        procuratorateTransferDate: "",
        criminalCourtDate: "",
        criminalResult: "",
        caseClosed: false,
        followUpNotes: ""
      });
    }
  }
}

function pickFirst(data, keywords) {
  for (const [key, value] of Object.entries(data)) {
    if (keywords.some(keyword => key.includes(keyword)) && String(value).trim()) {
      return value;
    }
  }
  return "";
}

function syncClosedCustomers() {
  appState.closedCustomers = appState.customers.filter(customer => customer.gradeKey === "closed");
}

function render() {
  els.downloadBtn.disabled = !appState.customers.length;
  els.emptyState.classList.toggle("hidden", appState.customers.length > 0);
  els.caseTableToolbar?.classList.remove("hidden");
  els.caseTableWrap.classList.remove("hidden");
  els.allCustomersPanel.classList.toggle("hidden", appState.customers.length === 0);
  els.allCustomersPage.classList.toggle("hidden", appState.activeView !== "all");
  els.closedCustomersPage.classList.toggle("hidden", appState.activeView !== "closed");
  els.allCustomersTabBtn.classList.toggle("active", appState.activeView === "all");
  els.closedCustomersTabBtn.classList.toggle("active", appState.activeView === "closed");
  els.sourceHint.className = "case-inline-stats";
  els.allCustomersHint.textContent = appState.fileName
    ? `来源：${appState.fileName}，已将 2026年每月客户情况统计 精准导入到全部客户页面。`
    : "将 2026年每月客户情况统计 导入到此页面，字段与原表一一对应。";

  renderSummary();
  renderCaseTable();
  renderCustomerTable();
  syncOverdueNotice();
  scheduleExportLengthRefresh();
  syncUploadProgress();
  syncTableMetrics();
  scheduleWorkspacePersist();
}

function restoreStoredWorkspace() {
  const saved = loadStoredWorkspace();
  if (!saved) return false;
  appState.fileName = saved.fileName || "";
  appState.customers = Array.isArray(saved.customers) ? saved.customers : [];
  appState.sourceHeaders = Array.isArray(saved.sourceHeaders) ? saved.sourceHeaders : [];
  appState.columnWidths = Array.isArray(saved.columnWidths) ? saved.columnWidths : [];
  appState.rowHeights = Array.isArray(saved.rowHeights) ? saved.rowHeights : [];
  appState.caseColumnWidths = Array.isArray(saved.caseColumnWidths) ? saved.caseColumnWidths : [];
  appState.caseRowHeights = Array.isArray(saved.caseRowHeights) ? saved.caseRowHeights : [];
  appState.rowSpacing = Number(saved.rowSpacing) || appState.rowSpacing;
  appState.columnWidth = Number(saved.columnWidth) || appState.columnWidth;
  appState.activeView = saved.activeView === "closed" ? "closed" : "all";
  if (saved.caseTypeHeaderLabel) {
    CASE_TYPE_HEADER.label = saved.caseTypeHeaderLabel;
  }
  restoreTrackingFieldLabels(saved.trackingFieldLabels);
  appState.trackingById = new Map(Object.entries(saved.trackingById || {}));
  ensureCustomerFields(appState.customers);
  syncClosedCustomers();
  seedTrackingRows();
  ensureCustomerMetrics();
  ensureCaseMetrics();
  render();
  return true;
}

function loadStoredWorkspace() {
  try {
    const raw = localStorage.getItem(WORKSPACE_STORAGE_KEY);
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
}

function persistWorkspace() {
  try {
    localStorage.setItem(WORKSPACE_STORAGE_KEY, JSON.stringify(serializeWorkspace()));
  } catch {
    // Ignore storage failures; the current in-memory table still works.
  }
}

function serializeWorkspace() {
  return {
    fileName: appState.fileName,
    customers: appState.customers,
    sourceHeaders: appState.sourceHeaders,
    columnWidths: appState.columnWidths,
    rowHeights: appState.rowHeights,
    caseColumnWidths: appState.caseColumnWidths,
    caseRowHeights: appState.caseRowHeights,
    trackingById: Object.fromEntries(appState.trackingById),
    trackingFieldLabels: Object.fromEntries(TRACKING_FIELDS.map(field => [field.key, field.label])),
    caseTypeHeaderLabel: CASE_TYPE_HEADER.label,
    activeView: appState.activeView,
    rowSpacing: appState.rowSpacing,
    columnWidth: appState.columnWidth
  };
}

function restoreTrackingFieldLabels(labels) {
  if (!labels || typeof labels !== "object") return;
  for (const field of TRACKING_FIELDS) {
    if (field.key === "caseClosed" && labels[field.key] === "案件结案") {
      continue;
    }
    if (typeof labels[field.key] === "string" && labels[field.key].trim()) {
      field.label = labels[field.key];
    }
  }
}

function renderSummary() {
  const openCases = appState.closedCustomers.filter(customer => {
    return !appState.trackingById.get(customer.id)?.caseClosed;
  }).length;
  const completedCases = appState.closedCustomers.length - openCases;
  const valuableProspects = appState.customers.filter(customer => customer.gradeKey === "valuable").length;
  if (els.totalCustomerCount) {
    els.totalCustomerCount.textContent = String(appState.customers.length);
  }
  els.closedCustomerCount.textContent = String(appState.closedCustomers.length);
  els.openCaseCount.textContent = String(openCases);
  els.completedCaseCount.textContent = String(completedCases);
  if (els.allCustomersInlineStats) {
    els.allCustomersInlineStats.innerHTML = [
      `目前客户 <strong>${appState.customers.length}</strong>`,
      `成交客户 <strong>${appState.closedCustomers.length}</strong>`,
      `潜在客户 <strong>${valuableProspects}</strong>`
    ]
      .map(item => `<span>${item}</span>`)
      .join("");
  }
  renderCaseInlineStats({ closedCustomers: appState.closedCustomers.length, completedCases, openCases });
}

function renderCaseInlineStats({ closedCustomers, completedCases, openCases }) {
  if (!els.sourceHint) return;
  els.sourceHint.innerHTML = [
    `成交客户 <strong>${closedCustomers}</strong>`,
    `已结案 <strong>${completedCases}</strong>`,
    `未结案 <strong>${openCases}</strong>`
  ]
    .map(item => `<span>${item}</span>`)
    .join("");
}

function syncOverdueNotice() {
  const overdueCases = getOverdueOpenCases();
  if (els.noticeBadge) {
    els.noticeBadge.textContent = String(overdueCases.length);
  }
  els.noticeBtn?.classList.toggle("has-alerts", overdueCases.length > 0);
  if (appState.customers.length && overdueCases.length) {
    const signature = overdueCases.map(item => item.customer.id).join("|");
    if (signature && signature !== appState.overdueNoticeShownFor) {
      appState.overdueNoticeShownFor = signature;
      window.setTimeout(() => openNoticeModal(true), 120);
    }
  } else {
    appState.overdueNoticeShownFor = "";
  }
}

function getOverdueOpenCases() {
  const threshold = new Date();
  threshold.setHours(0, 0, 0, 0);
  threshold.setMonth(threshold.getMonth() - OVERDUE_CASE_MONTHS);
  return appState.closedCustomers
    .map(customer => {
      const tracking = appState.trackingById.get(customer.id);
      const filingDate = parseDateInput(tracking?.filingDate);
      return { customer, tracking, filingDate };
    })
    .filter(item => item.tracking && !item.tracking.caseClosed && item.filingDate && item.filingDate <= threshold)
    .sort((a, b) => a.filingDate - b.filingDate);
}

function openNoticeModal(autoOpened) {
  const overdueCases = getOverdueOpenCases();
  if (overdueCases.length) {
    appState.overdueNoticeShownFor = overdueCases.map(item => item.customer.id).join("|");
  }
  if (els.noticeModalMeta) {
    els.noticeModalMeta.textContent = overdueCases.length
      ? `发现 ${overdueCases.length} 个超过 ${OVERDUE_CASE_MONTHS} 个月未结案的案件。`
      : `当前没有超过 ${OVERDUE_CASE_MONTHS} 个月未结案的案件。`;
  }
  if (els.noticeModalBody) {
    els.noticeModalBody.innerHTML = overdueCases.length
      ? overdueCases.map(renderOverdueCaseItem).join("")
      : `<div class="notice-empty">暂无逾期未结案案件。</div>`;
  }
  els.noticeModal?.classList.remove("hidden");
  els.noticeModal?.setAttribute("aria-hidden", "false");
}

function closeNoticeModal() {
  els.noticeModal?.classList.add("hidden");
  els.noticeModal?.setAttribute("aria-hidden", "true");
}

function renderOverdueCaseItem({ customer, tracking, filingDate }) {
  const customerName = getCustomerDisplayName(customer);
  const caseName = tracking.caseName || pickFirst(customer.data, ["案件", "案由", "委托事项", "咨询内容"]) || "未填写案件名称";
  const days = Math.max(0, Math.floor((Date.now() - filingDate.getTime()) / 86400000));
  return `<article class="notice-case-item">
    <div>
      <strong>${escapeHtml(customerName)}</strong>
      <span>${escapeHtml(caseName)}</span>
    </div>
    <dl>
      <dt>立案时间</dt>
      <dd>${escapeHtml(formatDateInput(filingDate))}</dd>
      <dt>未结天数</dt>
      <dd>${days} 天</dd>
    </dl>
  </article>`;
}

function getCustomerDisplayName(customer) {
  return (
    pickFirst(customer.data, ["客户姓名", "姓名", "客户", "当事人", "联系人"]) ||
    customer.data[appState.sourceHeaders.find(header => /客户|姓名/.test(header))] ||
    "未命名客户"
  );
}

function parseDateInput(value) {
  const text = String(value ?? "").trim();
  if (!text) return null;
  const normalized = text.replace(/[./]/g, "-");
  const match = normalized.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (!match) return null;
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const date = new Date(year, month - 1, day);
  if (date.getFullYear() !== year || date.getMonth() !== month - 1 || date.getDate() !== day) return null;
  date.setHours(0, 0, 0, 0);
  return date;
}

function formatDateInput(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function renderCaseTable() {
  ensureCaseMetrics();
  const columns = getCaseTableColumns();
  els.caseHead.innerHTML = `<tr>${columns
    .map((column, index) => {
      const headerAttrs =
        column.type === "case-type"
          ? `data-case-header-type="caseType"`
          : column.type === "source"
          ? `data-case-header-source-index="${column.index}"`
          : column.type === "tracking"
            ? `data-case-header-tracking-index="${column.index}"`
            : "";
      return `<th data-case-col-index="${index}" style="${getCaseColumnCellStyle(index)}">
        <div class="header-cell-shell">
          <input class="header-input" ${headerAttrs} value="${escapeAttr(column.label)}" aria-label="编辑列名" />
          ${column.type === "source" ? `<button class="icon-action delete-column-btn" type="button" data-delete-case-column="${column.index}" title="删除这一列">×</button>` : ""}
          <div class="col-resize-handle" data-case-col-resize="${index}" title="拖动调整列宽"></div>
        </div>
      </th>`;
    })
    .join("")}</tr>`;
  els.caseBody.innerHTML = "";

  for (const [rowIndex, customer] of appState.closedCustomers.entries()) {
    const tracking = appState.trackingById.get(customer.id);
    const row = document.createElement("tr");
    row.className = `grade-closed case-row ${tracking.caseClosed ? "case-row-completed" : "open-case case-row-open"}`;
    row.dataset.customerId = customer.id;
    row.dataset.caseRowIndex = String(rowIndex);
    applyCaseRowHeight(row, rowIndex);

    const cells = [];
    for (const [columnIndex, column] of columns.entries()) {
      const rowHandle =
        columnIndex === 0
          ? `<div class="row-resize-handle" data-case-row-resize="${rowIndex}" title="拖动调整行高"></div>`
          : "";
      const cellClass = columnIndex === 0 ? ` class="case-row-resize-anchor"` : "";
      cells.push(
        `<td${cellClass} data-case-col-index="${columnIndex}" style="${getCaseColumnCellStyle(columnIndex)}">
          ${renderCaseCellEditor(customer, tracking, column, rowIndex)}
          ${rowHandle}
        </td>`
      );
    }
    row.innerHTML = cells.join("");
    els.caseBody.appendChild(row);
  }

  els.caseBody.querySelectorAll("[data-field]").forEach(input => {
    input.addEventListener("input", handleTrackingInput);
    input.addEventListener("change", handleTrackingInput);
    bindCasePaste(input);
    bindExpandableEditor(input, {
      section: "成交客户",
      customerId: input.dataset.customerId,
      field: getCaseFieldLabel(input.dataset.field),
      kind: "case-tracking"
    });
  });
  els.caseBody.querySelectorAll("[data-case-customer-field]").forEach(input => {
    input.addEventListener("input", handleCaseCustomerInput);
    bindCasePaste(input);
    bindExpandableEditor(input, {
      section: "成交客户",
      customerId: input.dataset.customerId,
      field: input.dataset.caseCustomerField,
      kind: "case-customer"
    });
  });
  els.caseHead.querySelectorAll("[data-case-header-source-index]").forEach(input => {
    input.addEventListener("change", handleCaseSourceHeaderChange);
  });
  els.caseHead.querySelectorAll("[data-case-header-tracking-index]").forEach(input => {
    input.addEventListener("change", handleCaseTrackingHeaderChange);
  });
  els.caseHead.querySelectorAll("[data-case-header-type]").forEach(input => {
    input.addEventListener("change", handleCaseTypeHeaderChange);
  });
  els.caseHead.querySelectorAll("[data-delete-case-column]").forEach(button => {
    button.addEventListener("click", handleDeleteCaseColumn);
  });
  els.caseBody.querySelectorAll("[data-delete-case-row]").forEach(button => {
    button.addEventListener("click", handleDeleteCaseRow);
  });
  bindCaseResizeHandles();
  applyCaseTableSizing();
}

function getCaseTableColumns() {
  const columns = [{ type: "case-type", label: CASE_TYPE_HEADER.label }];
  const sourceInsertIndex = findCustomerSourceColumnIndex();
  for (const [sourceIndex, header] of appState.sourceHeaders.entries()) {
    if (sourceIndex === sourceInsertIndex) {
      columns.push(getTrackingColumn("caseClosed"));
    }
    columns.push({ type: "source", label: header, index: sourceIndex, header });
  }
  if (sourceInsertIndex < 0) {
    columns.push(getTrackingColumn("caseClosed"));
  }
  TRACKING_FIELDS.forEach((field, index) => {
    if (field.key !== "caseClosed") {
      columns.push({ type: "tracking", label: field.label, index, field });
    }
  });
  return columns;
}

function findCustomerSourceColumnIndex() {
  return appState.sourceHeaders.findIndex(header => /客户来源|账户来源|账号来源|来源/.test(header));
}

function getTrackingColumn(key) {
  const index = TRACKING_FIELDS.findIndex(field => field.key === key);
  const field = TRACKING_FIELDS[index];
  return { type: "tracking", label: field?.label || key, index, field };
}

function renderCaseCellEditor(customer, tracking, column, rowIndex) {
  if (column.type === "case-type") {
    return `<div class="row-action-shell">
      ${renderCaseTypeEditor(customer.id, tracking.caseType)}
      <button class="icon-action delete-row-btn" type="button" data-delete-case-row="${rowIndex}" title="删除这一行">×</button>
    </div>`;
  }
  if (column.type === "source") {
    return renderCaseCustomerEditor(customer.id, column.header, customer.data[column.header] ?? "", column.index);
  }
  return renderEditor(customer.id, column.field, tracking[column.field.key], column.index);
}

function renderEditor(customerId, field, value, trackingIndex = -1) {
  if (field.type === "checkbox") {
    return renderCaseCheckboxEditor(customerId, field.key, value, field.label, trackingIndex);
  }
  return renderCaseTextEditor(customerId, field.key, value || "", field.label, trackingIndex);
}

function renderCaseTypeEditor(customerId, value) {
  const normalized = normalizeCaseTypePasteValue(value);
  return `<select class="case-type-select" data-table-copy="case-field" data-customer-id="${escapeAttr(customerId)}" data-field="caseType" aria-label="案件类型">
    <option value="" ${normalized === "" ? "selected" : ""}></option>
    <option value="民事" ${normalized === "民事" ? "selected" : ""}>民事</option>
    <option value="刑事" ${normalized === "刑事" ? "selected" : ""}>刑事</option>
  </select>`;
}

function renderCaseTextEditor(customerId, fieldKey, value, label, trackingIndex = -1) {
  return `<textarea class="table-input case-tracking-input" data-table-copy="case-field" data-customer-id="${escapeAttr(customerId)}" data-field="${escapeAttr(fieldKey)}" data-case-tracking-index="${trackingIndex}" aria-label="${escapeAttr(label)}">${escapeHtml(value || "")}</textarea>`;
}

function renderCaseCheckboxEditor(customerId, fieldKey, checked, label, trackingIndex = -1) {
  return `<label class="case-checkbox-cell">
    <input data-table-copy="case-field" data-customer-id="${escapeAttr(customerId)}" data-field="${escapeAttr(fieldKey)}" data-case-tracking-index="${trackingIndex}" type="checkbox" ${checked ? "checked" : ""} aria-label="${escapeAttr(label)}" />
    <span>${checked ? "是" : "否"}</span>
  </label>`;
}

function getCaseFieldLabel(fieldKey) {
  if (fieldKey === "caseType") return CASE_TYPE_HEADER.label;
  return TRACKING_FIELDS.find(field => field.key === fieldKey)?.label || fieldKey;
}

function renderCaseCustomerEditor(customerId, header, value, sourceIndex = -1) {
  return `<textarea class="table-input case-customer-input" data-table-copy="case-customer" data-customer-id="${escapeAttr(customerId)}" data-case-customer-field="${escapeAttr(header)}" data-case-source-index="${sourceIndex}">${escapeHtml(value || "")}</textarea>`;
}

function handleTrackingInput(event) {
  const input = event.currentTarget;
  const tracking = appState.trackingById.get(input.dataset.customerId);
  if (!tracking) return;
  const field = TRACKING_FIELDS.find(item => item.key === input.dataset.field);
  tracking[input.dataset.field] =
    input.dataset.field === "caseType"
      ? normalizeCaseTypePasteValue(input.value)
      : field?.type === "checkbox"
      ? input.checked
      : input.value;
  renderSummary();
  syncOverdueNotice();
  scheduleWorkspacePersist();
  const row = input.closest("tr");
  if (row) {
    row.classList.toggle("open-case", !tracking.caseClosed);
    row.classList.toggle("case-row-open", !tracking.caseClosed);
    row.classList.toggle("case-row-completed", tracking.caseClosed);
  }
  if (field?.type === "checkbox") {
    input.closest(".case-checkbox-cell")?.querySelector("span")?.replaceChildren(input.checked ? "是" : "否");
  }
}

function handleCaseCustomerInput(event) {
  const input = event.currentTarget;
  const customer = findCustomer(input.dataset.customerId);
  if (!customer) return;
  customer.data[input.dataset.caseCustomerField] = input.value;
  scheduleExportLengthRefresh();
  scheduleWorkspacePersist();
}

function bindExpandableEditor(input, meta) {
  if (!(input instanceof HTMLTextAreaElement)) return;
  input.addEventListener("dblclick", () => {
    openEditorModal(input, meta);
  });
}

function openEditorModal(input, meta) {
  appState.activeModalEditor = { input, meta };
  if (els.editorModalMeta) {
    els.editorModalMeta.textContent = `${meta.section} / ${meta.field || "内容"}`;
  }
  if (els.editorModalTextarea) {
    els.editorModalTextarea.value = input.value || "";
  }
  els.editorModal?.classList.remove("hidden");
  els.editorModal?.setAttribute("aria-hidden", "false");
  window.setTimeout(() => {
    els.editorModalTextarea?.focus();
    els.editorModalTextarea?.select();
  }, 0);
}

function closeEditorModal() {
  appState.activeModalEditor = null;
  els.editorModal?.classList.add("hidden");
  els.editorModal?.setAttribute("aria-hidden", "true");
}

function saveEditorModal() {
  if (!appState.activeModalEditor || !els.editorModalTextarea) return;
  const { input, meta } = appState.activeModalEditor;
  input.value = els.editorModalTextarea.value;
  input.dispatchEvent(new Event("input", { bubbles: true }));
  if (meta.kind === "case-customer") {
    input.dispatchEvent(new Event("change", { bubbles: true }));
  }
  closeEditorModal();
}

function renderCustomerTable() {
  ensureCustomerMetrics();
  const headers = ["客户是否成交", ...appState.sourceHeaders];
  els.customerHead.innerHTML = `<tr>${headers
    .map(
      (label, index) =>
        `<th data-col-index="${index}" style="${getColumnCellStyle(index)}">
          <div class="header-cell-shell">
            ${index === 0 ? `<span class="static-header-label">${escapeHtml(label)}</span>` : renderCustomerHeaderEditor(label)}
            ${index === 0 ? "" : `<button class="icon-action delete-column-btn" type="button" data-delete-column="${index}" title="删除这一列">×</button>`}
            <div class="col-resize-handle" data-col-resize="${index}" title="drag to resize column"></div>
          </div>
        </th>`
    )
    .join("")}</tr>`;
  els.customerBody.innerHTML = "";

  for (const [rowIndex, customer] of appState.customers.entries()) {
    const row = document.createElement("tr");
    row.className = "customer-row";
    row.classList.toggle("customer-row-closed", customer.gradeKey === "closed");
    row.classList.toggle("customer-row-valuable", customer.gradeKey === "valuable");
    row.dataset.customerId = customer.id;
    row.dataset.rowIndex = String(rowIndex);
    applyCustomerRowHeight(row, rowIndex);
    const cells = [];
    cells.push(
      `<td class="status-cell" data-col-index="0" style="${getColumnCellStyle(0)}">
        <div class="row-action-shell">
          ${renderDealStatusEditor(customer.id, customer.gradeKey)}
          <button class="icon-action delete-row-btn" type="button" data-delete-row="${rowIndex}" title="删除这一行">×</button>
        </div>
        <div class="row-resize-handle" data-row-resize="${rowIndex}" title="drag to resize row"></div>
      </td>`
    );
    for (const [sourceIndex, header] of appState.sourceHeaders.entries()) {
      const columnIndex = sourceIndex + 1;
      const rowHandle =
        columnIndex === 1
          ? `<div class="row-resize-handle" data-row-resize="${rowIndex}" title="drag to resize row"></div>`
          : "";
      const cellClass = columnIndex === 1 ? "row-resize-anchor" : "customer-cell";
      cells.push(
        `<td class="${cellClass}" data-col-index="${columnIndex}" style="${getColumnCellStyle(columnIndex)}">
          ${renderCustomerDataEditor(customer.id, header, customer.data[header] ?? "", sourceIndex)}
          ${rowHandle}
        </td>`
      );
    }
    row.innerHTML = cells.join("");
    els.customerBody.appendChild(row);
  }

  els.customerBody.querySelectorAll("[data-customer-field]").forEach(input => {
    input.addEventListener("input", handleCustomerInput);
    bindCustomerPaste(input);
    bindExpandableEditor(input, {
      section: "全部客户",
      customerId: input.dataset.customerId,
      field: input.dataset.customerField,
      kind: "customer"
    });
  });
  els.customerBody.querySelectorAll("[data-customer-deal-status]").forEach(input => {
    input.addEventListener("change", handleCustomerDealStatusChange);
    bindCustomerPaste(input);
  });
  els.customerBody.querySelectorAll("[data-delete-row]").forEach(button => {
    button.addEventListener("click", handleDeleteRow);
  });
  els.customerHead.querySelectorAll("[data-delete-column]").forEach(button => {
    button.addEventListener("click", handleDeleteColumn);
  });
  els.customerHead.querySelectorAll("[data-header-name]").forEach(input => {
    input.addEventListener("input", handleCustomerHeaderInput);
  });
  bindCustomerResizeHandles();
  applyCustomerTableSizing();
}

function renderCustomerDataEditor(customerId, header, value, sourceIndex = -1) {
  return `<textarea class="table-input" data-table-copy="customer" data-customer-id="${escapeAttr(customerId)}" data-customer-field="${escapeAttr(header)}" data-customer-source-index="${sourceIndex}">${escapeHtml(value || "")}</textarea>`;
}

function renderDealStatusEditor(customerId, gradeKey) {
  return `<select class="deal-status-select" data-table-copy="customer-status" data-customer-id="${escapeAttr(customerId)}" data-customer-deal-status="true">
    <option value="unclosed" ${gradeKey === "unclosed" || gradeKey === "unknown" ? "selected" : ""}>未成交</option>
    <option value="valuable" ${gradeKey === "valuable" ? "selected" : ""}>潜在客户</option>
    <option value="closed" ${gradeKey === "closed" ? "selected" : ""}>成交</option>
  </select>`;
}

function renderCustomerHeaderEditor(header) {
  return `<input class="header-input" data-header-name="${escapeAttr(header)}" value="${escapeAttr(header)}" />`;
}

function handleCustomerInput(event) {
  const input = event.currentTarget;
  const customer = findCustomer(input.dataset.customerId);
  if (!customer) return;
  customer.data[input.dataset.customerField] = input.value;
  scheduleExportLengthRefresh();
  scheduleWorkspacePersist();
}

function bindCustomerPaste(input) {
  if (!(input instanceof HTMLTextAreaElement || input instanceof HTMLSelectElement)) return;
  input.addEventListener("paste", handleCustomerPaste);
}

function bindCasePaste(input) {
  if (
    !(
      input instanceof HTMLTextAreaElement ||
      input instanceof HTMLInputElement ||
      input instanceof HTMLSelectElement
    )
  ) {
    return;
  }
  input.addEventListener("paste", handleCasePaste);
}

function handleCustomerPaste(event) {
  const input = event.currentTarget;
  const grid = parseClipboardGrid(event.clipboardData?.getData("text/plain") || "");
  if (!grid.length) return;
  if (input.dataset.customerDealStatus) {
    event.preventDefault();
    applySingleInputPaste(input, grid[0]?.[0] || "");
    return;
  }
  const rowIndex = appState.customers.findIndex(customer => customer.id === input.dataset.customerId);
  const columnIndex = Number(input.dataset.customerSourceIndex);
  if (rowIndex < 0 || !Number.isFinite(columnIndex)) return;
  event.preventDefault();
  if (!isMultiCellPaste(grid)) {
    applySingleInputPaste(input, grid[0]?.[0] || "");
    return;
  }
  applyCustomerGridPaste(rowIndex, columnIndex, grid);
}

function handleCasePaste(event) {
  const input = event.currentTarget;
  const rowIndex = appState.closedCustomers.findIndex(customer => customer.id === input.dataset.customerId);
  if (rowIndex < 0) return;
  const grid = parseClipboardGrid(event.clipboardData?.getData("text/plain") || "");
  if (!grid.length) return;
  event.preventDefault();

  if (input.dataset.caseCustomerField) {
    const columnIndex = Number(input.dataset.caseSourceIndex);
    if (!Number.isFinite(columnIndex)) return;
    if (!isMultiCellPaste(grid)) {
      applySingleInputPaste(input, grid[0]?.[0] || "");
      return;
    }
    applyCaseCustomerGridPaste(rowIndex, columnIndex, grid);
    return;
  }

  if (input.dataset.field) {
    const columnIndex = findCaseInputColumnIndex(input);
    if (!Number.isFinite(columnIndex)) {
      applySingleInputPaste(input, grid[0]?.[0] || "");
      return;
    }
    if (!isMultiCellPaste(grid)) {
      applySingleInputPaste(input, grid[0]?.[0] || "");
      return;
    }
    applyCaseTableGridPaste(rowIndex, columnIndex, grid);
  }
}

function applySingleInputPaste(input, value) {
  if (input instanceof HTMLInputElement && input.type === "checkbox") {
    input.checked = normalizeTrackingPasteValue({ type: "checkbox" }, value);
    input.dispatchEvent(new Event("change", { bubbles: true }));
    return;
  }
  input.value =
    input.dataset.customerDealStatus
      ? normalizeDealStatusPasteValue(value)
      : input.dataset.field === "caseType"
        ? normalizeCaseTypePasteValue(value)
        : value;
  input.dispatchEvent(new Event("input", { bubbles: true }));
  input.dispatchEvent(new Event("change", { bubbles: true }));
}

function normalizeCaseTypePasteValue(value) {
  const text = String(value || "").trim();
  const normalized = text.toLowerCase();
  if (text === "民事" || normalized === "civil") return "民事";
  if (text === "刑事" || normalized === "criminal") return "刑事";
  return "";
}

function normalizeDealStatusPasteValue(value) {
  const text = String(value || "").trim();
  if (["closed", "成交", "成交客户"].includes(text)) return "closed";
  if (["valuable", "有价值潜在客户", "潜在客户"].includes(text)) return "valuable";
  if (["unclosed", "未成交", "潜在客户未成交", "未识别"].includes(text)) return "unclosed";
  return text;
}

function handleTableCopy(event) {
  const selection = window.getSelection();
  if (selection && !selection.isCollapsed && String(selection).trim()) return;

  const active = document.activeElement;
  if (!isCopyableTableInput(active)) return;

  const text = getCopyableInputValue(active);
  if (event.clipboardData) {
    event.clipboardData.setData("text/plain", text);
    event.preventDefault();
    return;
  }
  navigator.clipboard?.writeText(text).catch(() => {});
}

function isCopyableTableInput(element) {
  return (
    element instanceof HTMLInputElement ||
    element instanceof HTMLTextAreaElement ||
    element instanceof HTMLSelectElement
  ) && Boolean(element.dataset.tableCopy);
}

function getCopyableInputValue(input) {
  if (input instanceof HTMLInputElement && input.type === "checkbox") {
    return input.checked ? "是" : "否";
  }
  if (input instanceof HTMLSelectElement) {
    return input.selectedOptions[0]?.textContent || input.value || "";
  }
  return input.value || "";
}

function parseClipboardGrid(text) {
  return String(text || "")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .split("\n")
    .filter((row, index, rows) => row !== "" || index < rows.length - 1)
    .map(row => row.split("\t"));
}

function isMultiCellPaste(grid) {
  return grid.length > 1 || (grid[0] && grid[0].length > 1);
}

function applyCustomerGridPaste(startRowIndex, startColumnIndex, grid) {
  grid.forEach((rowValues, rowOffset) => {
    const customer = appState.customers[startRowIndex + rowOffset];
    if (!customer) return;
    rowValues.forEach((value, columnOffset) => {
      const header = appState.sourceHeaders[startColumnIndex + columnOffset];
      if (!header) return;
      customer.data[header] = value;
    });
  });
  renderCustomerTable();
  scheduleExportLengthRefresh();
  scheduleWorkspacePersist();
}

function applyCaseCustomerGridPaste(startRowIndex, startColumnIndex, grid) {
  grid.forEach((rowValues, rowOffset) => {
    const customer = appState.closedCustomers[startRowIndex + rowOffset];
    if (!customer) return;
    rowValues.forEach((value, columnOffset) => {
      const header = appState.sourceHeaders[startColumnIndex + columnOffset];
      if (!header) return;
      customer.data[header] = value;
    });
  });
  renderCaseTable();
  renderCustomerTable();
  scheduleExportLengthRefresh();
  scheduleWorkspacePersist();
}

function applyCaseTrackingGridPaste(startRowIndex, startColumnIndex, grid) {
  applyCaseTableGridPaste(startRowIndex, getCaseTrackingColumnIndex(startColumnIndex), grid);
}

function applyCaseTableGridPaste(startRowIndex, startColumnIndex, grid) {
  const columns = getCaseTableColumns();
  grid.forEach((rowValues, rowOffset) => {
    const customer = appState.closedCustomers[startRowIndex + rowOffset];
    if (!customer) return;
    const tracking = appState.trackingById.get(customer.id);
    if (!tracking) return;
    rowValues.forEach((value, columnOffset) => {
      const column = columns[startColumnIndex + columnOffset];
      if (!column) return;
      applyCaseColumnPasteValue(customer, tracking, column, value);
    });
  });
  renderSummary();
  renderCaseTable();
  renderCustomerTable();
  scheduleExportLengthRefresh();
  scheduleWorkspacePersist();
}

function applyCaseColumnPasteValue(customer, tracking, column, value) {
  if (column.type === "case-type") {
    tracking.caseType = value;
    return;
  }
  if (column.type === "source") {
    customer.data[column.header] = value;
    return;
  }
  if (column.type === "tracking") {
    tracking[column.field.key] = normalizeTrackingPasteValue(column.field, value);
  }
}

function findCaseInputColumnIndex(input) {
  const cell = input.closest("[data-case-col-index]");
  return Number(cell?.dataset.caseColIndex);
}

function getCaseTrackingColumnIndex(trackingIndex) {
  return getCaseTableColumns().findIndex(column => column.type === "tracking" && column.index === trackingIndex);
}

function normalizeTrackingPasteValue(field, value) {
  if (field.type === "checkbox") {
    return ["1", "true", "yes", "是", "已结案"].includes(String(value).trim().toLowerCase());
  }
  return value;
}

function handleCustomerDealStatusChange(event) {
  const input = event.currentTarget;
  const customer = findCustomer(input.dataset.customerId);
  if (!customer) return;
  customer.gradeKey =
    input.value === "closed"
      ? "closed"
      : input.value === "valuable"
        ? "valuable"
        : "unclosed";
  customer.grade = CUSTOMER_GRADE[customer.gradeKey];
  syncClosedCustomers();
  seedTrackingRows();
  renderSummary();
  renderCaseTable();
  scheduleExportLengthRefresh();
  scheduleWorkspacePersist();

  const row = input.closest("tr");
  if (row) {
    row.classList.toggle("customer-row-closed", customer.gradeKey === "closed");
    row.classList.toggle("customer-row-valuable", customer.gradeKey === "valuable");
  }
}

function handleDeleteRow(event) {
  const rowIndex = Number(event.currentTarget.dataset.deleteRow);
  const customer = appState.customers[rowIndex];
  if (!customer) return;
  if (!confirmDelete("确认删除这一行客户数据吗？删除后该行的案件跟踪信息也会一起移除。")) return;
  deleteCustomerById(customer.id);
}

function handleDeleteColumn(event) {
  const columnIndex = Number(event.currentTarget.dataset.deleteColumn);
  if (!Number.isFinite(columnIndex) || columnIndex <= 0) return;
  const sourceIndex = columnIndex - 1;
  const header = appState.sourceHeaders[sourceIndex];
  if (!header) return;
  if (!confirmDelete(`确认删除「${header}」这一列吗？所有客户这一列的数据都会被移除。`)) return;
  deleteSourceColumn(sourceIndex);
}

function handleDeleteCaseRow(event) {
  const rowIndex = Number(event.currentTarget.dataset.deleteCaseRow);
  const customer = appState.closedCustomers[rowIndex];
  if (!customer) return;
  if (!confirmDelete("确认删除这一行成交客户案件数据吗？删除后对应客户行也会一起移除。")) return;
  deleteCustomerById(customer.id);
}

function handleDeleteCaseColumn(event) {
  const sourceIndex = Number(event.currentTarget.dataset.deleteCaseColumn);
  const header = appState.sourceHeaders[sourceIndex];
  if (!header) return;
  if (!confirmDelete(`确认删除「${header}」这一列吗？所有客户这一列的数据都会被移除。`)) return;
  deleteSourceColumn(sourceIndex);
}

function confirmDelete(message) {
  return window.confirm(message);
}

function handleCustomerHeaderInput(event) {
  const input = event.currentTarget;
  const oldHeader = input.dataset.headerName;
  const newHeader = String(input.value || "").trim() || oldHeader;
  if (oldHeader === DEAL_AMOUNT_HEADER) {
    input.value = DEAL_AMOUNT_HEADER;
    return;
  }
  if (oldHeader === newHeader) return;
  if (appState.sourceHeaders.includes(newHeader)) return;

  const headerIndex = appState.sourceHeaders.indexOf(oldHeader);
  if (headerIndex < 0) return;
  appState.sourceHeaders[headerIndex] = newHeader;

  for (const customer of appState.customers) {
    customer.data[newHeader] = customer.data[oldHeader] ?? "";
    delete customer.data[oldHeader];
  }

  render();
}

function handleCaseSourceHeaderChange(event) {
  const input = event.currentTarget;
  const sourceIndex = Number(input.dataset.caseHeaderSourceIndex);
  const oldHeader = appState.sourceHeaders[sourceIndex];
  const newHeader = String(input.value || "").trim() || oldHeader;
  if (!oldHeader) return;
  if (oldHeader === DEAL_AMOUNT_HEADER) {
    input.value = DEAL_AMOUNT_HEADER;
    return;
  }
  if (oldHeader === newHeader) return;
  if (appState.sourceHeaders.includes(newHeader)) {
    input.value = oldHeader;
    return;
  }

  appState.sourceHeaders[sourceIndex] = newHeader;
  for (const customer of appState.customers) {
    customer.data[newHeader] = customer.data[oldHeader] ?? "";
    delete customer.data[oldHeader];
  }
  render();
}

function handleCaseTrackingHeaderChange(event) {
  const input = event.currentTarget;
  const trackingIndex = Number(input.dataset.caseHeaderTrackingIndex);
  const field = TRACKING_FIELDS[trackingIndex];
  if (!field) return;
  const newLabel = String(input.value || "").trim() || field.label;
  field.label = newLabel;
  renderCaseTable();
  scheduleExportLengthRefresh();
  scheduleWorkspacePersist();
}

function handleCaseTypeHeaderChange(event) {
  const input = event.currentTarget;
  CASE_TYPE_HEADER.label = String(input.value || "").trim() || CASE_TYPE_HEADER.label;
  renderCaseTable();
  scheduleExportLengthRefresh();
  scheduleWorkspacePersist();
}

function findCustomer(customerId) {
  return appState.customers.find(customer => customer.id === customerId);
}

function addCustomerRow() {
  const emptyData = Object.fromEntries(appState.sourceHeaders.map(header => [header, ""]));
  appState.customers.push({
    id: `manual-row-${Date.now()}-${appState.customers.length}`,
    sheetName: "2026年每月客户情况统计",
    rowNumber: appState.customers.length + 1,
    gradeKey: "unknown",
    grade: CUSTOMER_GRADE.unknown,
    data: emptyData,
    values: appState.sourceHeaders.map(() => "")
  });
  appState.rowHeights.push(appState.rowSpacing);
  render();
}

function addCustomerColumn() {
  addSourceColumn();
}

function addCaseRow() {
  if (!appState.sourceHeaders.length) {
    appState.sourceHeaders = ["客户姓名", "客户电话"];
    resetCustomerMetrics();
  }
  const emptyData = Object.fromEntries(appState.sourceHeaders.map(header => [header, ""]));
  const customer = {
    id: `manual-case-row-${Date.now()}-${appState.customers.length}`,
    sheetName: "成交客户案件跟踪",
    rowNumber: appState.customers.length + 1,
    gradeKey: "closed",
    grade: CUSTOMER_GRADE.closed,
    data: emptyData,
    values: appState.sourceHeaders.map(() => "")
  };
  appState.customers.push(customer);
  appState.rowHeights.push(appState.rowSpacing);
  syncClosedCustomers();
  seedTrackingRows();
  appState.caseRowHeights.push(appState.rowSpacing);
  render();
}

function addCaseColumn() {
  addSourceColumn();
}

function addSourceColumn() {
  const baseName = "新增列";
  let header = baseName;
  let suffix = 1;
  while (appState.sourceHeaders.includes(header)) {
    suffix += 1;
    header = `${baseName}${suffix}`;
  }
  const dealAmountIndex = appState.sourceHeaders.indexOf(DEAL_AMOUNT_HEADER);
  const insertAt = dealAmountIndex >= 0 ? dealAmountIndex : appState.sourceHeaders.length;
  appState.sourceHeaders.splice(insertAt, 0, header);
  appState.columnWidths.splice(insertAt + 1, 0, appState.columnWidth);
  appState.caseColumnWidths.splice(insertAt + 1, 0, appState.columnWidth);
  for (const customer of appState.customers) {
    customer.data[header] = "";
  }
  render();
}

function deleteCustomerById(customerId) {
  const customerIndex = appState.customers.findIndex(customer => customer.id === customerId);
  if (customerIndex < 0) return;
  const closedIndex = appState.closedCustomers.findIndex(customer => customer.id === customerId);
  appState.customers.splice(customerIndex, 1);
  appState.rowHeights.splice(customerIndex, 1);
  if (closedIndex >= 0) {
    appState.caseRowHeights.splice(closedIndex, 1);
  }
  appState.trackingById.delete(customerId);
  syncClosedCustomers();
  render();
}

function deleteSourceColumn(sourceIndex) {
  if (!Number.isFinite(sourceIndex) || sourceIndex < 0) return;
  const header = appState.sourceHeaders[sourceIndex];
  if (!header || header === DEAL_AMOUNT_HEADER) return;
  appState.sourceHeaders.splice(sourceIndex, 1);
  appState.columnWidths.splice(sourceIndex + 1, 1);
  appState.caseColumnWidths.splice(sourceIndex + 1, 1);
  for (const customer of appState.customers) {
    delete customer.data[header];
  }
  render();
}

function syncTableMetrics() {
  document.documentElement.style.setProperty("--table-row-height", `${appState.rowSpacing}px`);
  document.documentElement.style.setProperty("--table-column-width", `${appState.columnWidth}px`);
  if (els.rowSpacingRange) els.rowSpacingRange.value = String(appState.rowSpacing);
  if (els.columnWidthRange) els.columnWidthRange.value = String(appState.columnWidth);
  applyCustomerTableSizing();
  applyCaseTableSizing();
}

function resetCustomerMetrics() {
  appState.columnWidths = Array.from(
    { length: appState.sourceHeaders.length + 1 },
    () => appState.columnWidth
  );
  appState.rowHeights = appState.customers.map(() => appState.rowSpacing);
  resetCaseMetrics();
}

function ensureCustomerMetrics() {
  const desiredColumnCount = appState.sourceHeaders.length + 1;
  while (appState.columnWidths.length < desiredColumnCount) {
    appState.columnWidths.push(appState.columnWidth);
  }
  if (appState.columnWidths.length > desiredColumnCount) {
    appState.columnWidths = appState.columnWidths.slice(0, desiredColumnCount);
  }
  while (appState.rowHeights.length < appState.customers.length) {
    appState.rowHeights.push(appState.rowSpacing);
  }
  if (appState.rowHeights.length > appState.customers.length) {
    appState.rowHeights = appState.rowHeights.slice(0, appState.customers.length);
  }
}

function resetCaseMetrics() {
  appState.caseColumnWidths = Array.from(
    { length: getCaseColumnCount() },
    () => appState.columnWidth
  );
  appState.caseRowHeights = appState.closedCustomers.map(() => appState.rowSpacing);
}

function ensureCaseMetrics() {
  const desiredColumnCount = getCaseColumnCount();
  while (appState.caseColumnWidths.length < desiredColumnCount) {
    appState.caseColumnWidths.push(appState.columnWidth);
  }
  if (appState.caseColumnWidths.length > desiredColumnCount) {
    appState.caseColumnWidths = appState.caseColumnWidths.slice(0, desiredColumnCount);
  }
  while (appState.caseRowHeights.length < appState.closedCustomers.length) {
    appState.caseRowHeights.push(appState.rowSpacing);
  }
  if (appState.caseRowHeights.length > appState.closedCustomers.length) {
    appState.caseRowHeights = appState.caseRowHeights.slice(0, appState.closedCustomers.length);
  }
}

function getCaseColumnCount() {
  return getCaseTableColumns().length;
}

function getColumnCellStyle(index) {
  const width = appState.columnWidths[index] || appState.columnWidth;
  return `width:${width}px;min-width:${width}px;`;
}

function getCaseColumnCellStyle(index) {
  const width = appState.caseColumnWidths[index] || appState.columnWidth;
  return `width:${width}px;min-width:${width}px;`;
}

function applyCustomerRowHeight(row, rowIndex) {
  const height = appState.rowHeights[rowIndex] || appState.rowSpacing;
  row.style.setProperty("--row-size", `${height}px`);
}

function applyCaseRowHeight(row, rowIndex) {
  const height = appState.caseRowHeights[rowIndex] || appState.rowSpacing;
  row.style.setProperty("--row-size", `${height}px`);
}

function applyCustomerTableSizing() {
  if (!els.customerTable) return;
  ensureCustomerMetrics();

  appState.columnWidths.forEach((width, index) => {
    els.customerTable.querySelectorAll(`[data-col-index="${index}"]`).forEach(cell => {
      cell.style.width = `${width}px`;
      cell.style.minWidth = `${width}px`;
    });
  });

  Array.from(els.customerBody.children).forEach((row, rowIndex) => {
    applyCustomerRowHeight(row, rowIndex);
  });
}

function applyCaseTableSizing() {
  if (!els.caseHead || !els.caseBody) return;
  ensureCaseMetrics();

  appState.caseColumnWidths.forEach((width, index) => {
    els.caseTableWrap.querySelectorAll(`[data-case-col-index="${index}"]`).forEach(cell => {
      cell.style.width = `${width}px`;
      cell.style.minWidth = `${width}px`;
    });
  });

  Array.from(els.caseBody.children).forEach((row, rowIndex) => {
    applyCaseRowHeight(row, rowIndex);
  });
}

function bindCustomerResizeHandles() {
  els.customerHead.querySelectorAll("[data-col-resize]").forEach(handle => {
    handle.addEventListener("pointerdown", startColumnResize);
  });
  els.customerBody.querySelectorAll("[data-row-resize]").forEach(handle => {
    handle.addEventListener("pointerdown", startRowResize);
  });
}

function bindCaseResizeHandles() {
  els.caseHead.querySelectorAll("[data-case-col-resize]").forEach(handle => {
    handle.addEventListener("pointerdown", startCaseColumnResize);
  });
  els.caseBody.querySelectorAll("[data-case-row-resize]").forEach(handle => {
    handle.addEventListener("pointerdown", startCaseRowResize);
  });
}

function startColumnResize(event) {
  if (event.button !== 0) return;
  event.preventDefault();
  event.stopPropagation();
  const columnIndex = Number(event.currentTarget.dataset.colResize);
  beginTableResize({
    type: "column",
    index: columnIndex,
    startX: event.clientX,
    startY: event.clientY,
    startSize: appState.columnWidths[columnIndex] || appState.columnWidth
  });
}

function startCaseColumnResize(event) {
  if (event.button !== 0) return;
  event.preventDefault();
  event.stopPropagation();
  const columnIndex = Number(event.currentTarget.dataset.caseColResize);
  beginTableResize({
    type: "case-column",
    index: columnIndex,
    startX: event.clientX,
    startY: event.clientY,
    startSize: appState.caseColumnWidths[columnIndex] || appState.columnWidth
  });
}

function startRowResize(event) {
  if (event.button !== 0) return;
  event.preventDefault();
  event.stopPropagation();
  const rowIndex = Number(event.currentTarget.dataset.rowResize);
  beginTableResize({
    type: "row",
    index: rowIndex,
    startX: event.clientX,
    startY: event.clientY,
    startSize: appState.rowHeights[rowIndex] || appState.rowSpacing
  });
}

function startCaseRowResize(event) {
  if (event.button !== 0) return;
  event.preventDefault();
  event.stopPropagation();
  const rowIndex = Number(event.currentTarget.dataset.caseRowResize);
  beginTableResize({
    type: "case-row",
    index: rowIndex,
    startX: event.clientX,
    startY: event.clientY,
    startSize: appState.caseRowHeights[rowIndex] || appState.rowSpacing
  });
}

function beginTableResize(session) {
  stopTableResize();
  appState.activeResize = session;
  document.body.classList.add("is-resizing");
  document.body.classList.add(session.type.includes("column") ? "is-resizing-column" : "is-resizing-row");
  window.addEventListener("pointermove", handleTableResizeMove);
  window.addEventListener("pointerup", stopTableResize);
  window.addEventListener("pointercancel", stopTableResize);
}

function handleTableResizeMove(event) {
  if (!appState.activeResize) return;
  if (appState.activeResize.type === "column") {
    const nextWidth = clamp(
      appState.activeResize.startSize + (event.clientX - appState.activeResize.startX),
      COLUMN_WIDTH_MIN,
      COLUMN_WIDTH_MAX
    );
    appState.columnWidths[appState.activeResize.index] = nextWidth;
    applyCustomerTableSizing();
    return;
  }
  if (appState.activeResize.type === "case-column") {
    const nextWidth = clamp(
      appState.activeResize.startSize + (event.clientX - appState.activeResize.startX),
      COLUMN_WIDTH_MIN,
      COLUMN_WIDTH_MAX
    );
    appState.caseColumnWidths[appState.activeResize.index] = nextWidth;
    applyCaseTableSizing();
    return;
  }
  if (appState.activeResize.type === "case-row") {
    const nextHeight = clamp(
      appState.activeResize.startSize + (event.clientY - appState.activeResize.startY),
      ROW_HEIGHT_MIN,
      ROW_HEIGHT_MAX
    );
    appState.caseRowHeights[appState.activeResize.index] = nextHeight;
    applyCaseTableSizing();
    return;
  }
  if (appState.activeResize.type === "row") {
    const nextHeight = clamp(
      appState.activeResize.startSize + (event.clientY - appState.activeResize.startY),
      ROW_HEIGHT_MIN,
      ROW_HEIGHT_MAX
    );
    appState.rowHeights[appState.activeResize.index] = nextHeight;
    applyCustomerTableSizing();
  }
}

function stopTableResize() {
  if (!appState.activeResize) return;
  document.body.classList.remove("is-resizing", "is-resizing-column", "is-resizing-row");
  window.removeEventListener("pointermove", handleTableResizeMove);
  window.removeEventListener("pointerup", stopTableResize);
  window.removeEventListener("pointercancel", stopTableResize);
  appState.activeResize = null;
  scheduleWorkspacePersist();
}

function clamp(value, min, max) {
  return Math.min(max, Math.max(min, value));
}

function updateUploadProgress(progress, message) {
  appState.uploadProgress = progress;
  appState.uploadMessage = message;
  syncUploadProgress();
}

function syncUploadProgress() {
  els.uploadStatus.classList.add("hidden");
  els.uploadProgressBar.style.width = `${Math.max(0, Math.min(100, appState.uploadProgress))}%`;
  els.uploadProgressText.textContent = appState.uploadMessage;
}

function loadStoredTheme() {
  try {
    return localStorage.getItem("law-scrm-theme") || DEFAULT_THEME;
  } catch {
    return DEFAULT_THEME;
  }
}

function persistTheme(theme) {
  try {
    localStorage.setItem("law-scrm-theme", theme);
  } catch {
    // Ignore storage failures.
  }
}

function applyTheme() {
  document.documentElement.dataset.theme = appState.theme;
  els.themeToggleBtn.textContent = appState.theme === "dark" ? "浅色" : "深色";
}

function buildManagedWorkbook() {
  const allHeaders = appState.sourceHeaders;
  const caseColumns = getCaseTableColumns();
  const caseHeaders = ["案件状态", "客户等级", "来源表", ...caseColumns.map(column => column.label)];
  const caseRows = appState.closedCustomers.map(customer => {
    const tracking = appState.trackingById.get(customer.id);
    return [
      tracking.caseClosed ? "已结案" : "未结案",
      customer.grade,
      customer.sheetName,
      ...caseColumns.map(column => getCaseColumnExportValue(customer, tracking, column))
    ];
  });
  const customerRows = getRawCustomerRows();
  return writeXlsx([
    {
      name: "成交客户案件跟踪",
      rows: [
        [`未结案件数量`, String(caseRows.filter(row => row[0] === "未结案").length)],
        [],
        caseHeaders,
        ...caseRows
      ]
    },
    {
      name: "全部客户等级",
      rows: [allHeaders, ...customerRows]
    }
  ]);
}

function getCaseColumnExportValue(customer, tracking, column) {
  if (column.type === "case-type") return tracking.caseType || "";
  if (column.type === "source") return customer.data[column.header] ?? "";
  if (column.type === "tracking") {
    return column.field.type === "checkbox"
      ? (tracking[column.field.key] ? "是" : "否")
      : tracking[column.field.key] || "";
  }
  return "";
}

async function parseXlsx(buffer) {
  const zip = await ZipReader.fromArrayBuffer(buffer);
  const workbookXml = parseXml(await zip.text("xl/workbook.xml"));
  const workbookRels = parseXml(await zip.text("xl/_rels/workbook.xml.rels"));
  const sharedStrings = zip.has("xl/sharedStrings.xml")
    ? parseSharedStrings(parseXml(await zip.text("xl/sharedStrings.xml")))
    : [];
  const styles = parseStyles(parseXml(await zip.text("xl/styles.xml")));
  const relMap = new Map(
    Array.from(workbookRels.getElementsByTagName("Relationship")).map(rel => [
      rel.getAttribute("Id"),
      normalizePath("xl/" + rel.getAttribute("Target"))
    ])
  );
  const sheets = Array.from(workbookXml.getElementsByTagNameNS(SPREADSHEET_NS, "sheet")).map(sheet => ({
    name: sheet.getAttribute("name"),
    path: relMap.get(sheet.getAttributeNS(REL_NS, "id"))
  }));

  for (const sheet of sheets) {
    sheet.rows = parseWorksheet(parseXml(await zip.text(sheet.path)), sharedStrings, styles);
  }
  return { sheets };
}

const SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

function parseXml(xml) {
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  if (doc.querySelector("parsererror")) throw new Error("Excel XML 解析失败");
  return doc;
}

function parseSharedStrings(doc) {
  return Array.from(doc.getElementsByTagNameNS(SPREADSHEET_NS, "si")).map(si =>
    Array.from(si.getElementsByTagNameNS(SPREADSHEET_NS, "t"))
      .map(t => t.textContent || "")
      .join("")
  );
}

function parseStyles(doc) {
  const fills = Array.from(doc.getElementsByTagNameNS(SPREADSHEET_NS, "fills")[0]?.children || []).map(fill => {
    const fg = fill.getElementsByTagNameNS(SPREADSHEET_NS, "fgColor")[0];
    return fg?.getAttribute("rgb") || fg?.getAttribute("indexed") || "";
  });
  const cellXfs = Array.from(doc.getElementsByTagNameNS(SPREADSHEET_NS, "cellXfs")[0]?.children || []).map(xf => ({
    fillId: Number(xf.getAttribute("fillId") || 0),
    numFmtId: Number(xf.getAttribute("numFmtId") || 0)
  }));
  return { fills, cellXfs };
}

function parseWorksheet(doc, sharedStrings, styles) {
  const mergeMap = buildMergeMap(doc);
  const rows = [];
  for (const rowNode of doc.getElementsByTagNameNS(SPREADSHEET_NS, "row")) {
    const row = { number: Number(rowNode.getAttribute("r") || rows.length + 1), cells: [] };
    for (const cellNode of rowNode.getElementsByTagNameNS(SPREADSHEET_NS, "c")) {
      const ref = cellNode.getAttribute("r") || "";
      const index = columnNameToIndex(ref.replace(/\d+/g, ""));
      const styleIndex = Number(cellNode.getAttribute("s") || 0);
      const cellStyle = styles.cellXfs[styleIndex] || { fillId: 0, numFmtId: 0 };
      row.cells[index] = {
        index,
        value: parseCellValue(cellNode, sharedStrings, cellStyle),
        fillColor: styles.fills[cellStyle.fillId] || ""
      };
    }
    rows.push(row);
  }
  applyMergedCellValues(rows, mergeMap);
  return rows;
}

function buildMergeMap(doc) {
  const map = new Map();
  for (const mergeCell of doc.getElementsByTagNameNS(SPREADSHEET_NS, "mergeCell")) {
    const ref = mergeCell.getAttribute("ref");
    if (!ref || !ref.includes(":")) continue;
    const [startRef, endRef] = ref.split(":");
    const start = parseCellRef(startRef);
    const end = parseCellRef(endRef);
    map.set(startRef, { start, end });
  }
  return map;
}

function applyMergedCellValues(rows, mergeMap) {
  const rowByNumber = new Map(rows.map(row => [row.number, row]));
  for (const row of rows) {
    for (const cell of row.cells) {
      if (!cell) continue;
      const key = `${indexToColumnName(cell.index)}${row.number}`;
      const merge = mergeMap.get(key);
      if (!merge) continue;
      for (let rowNumber = merge.start.row; rowNumber <= merge.end.row; rowNumber++) {
        const targetRow = rowByNumber.get(rowNumber);
        if (!targetRow) continue;
        for (let columnIndex = merge.start.column; columnIndex <= merge.end.column; columnIndex++) {
          if (!targetRow.cells[columnIndex]) {
            targetRow.cells[columnIndex] = {
              index: columnIndex,
              value: cell.value,
              fillColor: cell.fillColor
            };
          } else if (!String(targetRow.cells[columnIndex].value || "").trim()) {
            targetRow.cells[columnIndex].value = cell.value;
            targetRow.cells[columnIndex].fillColor ||= cell.fillColor;
          }
        }
      }
    }
  }
}

function parseCellRef(ref) {
  const match = /^([A-Z]+)(\d+)$/.exec(ref);
  if (!match) return { column: 0, row: 0 };
  return {
    column: columnNameToIndex(match[1]),
    row: Number(match[2])
  };
}

function parseCellValue(cellNode, sharedStrings, cellStyle) {
  const type = cellNode.getAttribute("t");
  if (type === "inlineStr") {
    return Array.from(cellNode.getElementsByTagNameNS(SPREADSHEET_NS, "t"))
      .map(t => t.textContent || "")
      .join("");
  }
  const value = cellNode.getElementsByTagNameNS(SPREADSHEET_NS, "v")[0]?.textContent || "";
  if (type === "s") return sharedStrings[Number(value)] || "";
  if (type === "b") return value === "1" ? "TRUE" : "FALSE";
  if (looksLikeDateFormat(cellStyle?.numFmtId) && value) {
    return excelSerialToDateString(Number(value));
  }
  return value;
}

function looksLikeDateFormat(numFmtId) {
  return [14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47].includes(numFmtId);
}

function excelSerialToDateString(serial) {
  if (!Number.isFinite(serial)) return String(serial ?? "");
  const utcDays = Math.floor(serial - 25569);
  const utcValue = utcDays * 86400;
  const dateInfo = new Date(utcValue * 1000);
  const year = dateInfo.getUTCFullYear();
  const month = String(dateInfo.getUTCMonth() + 1).padStart(2, "0");
  const day = String(dateInfo.getUTCDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function columnNameToIndex(columnName) {
  let index = 0;
  for (const char of columnName) {
    index = index * 26 + char.charCodeAt(0) - 64;
  }
  return Math.max(0, index - 1);
}

function indexToColumnName(index) {
  let name = "";
  let n = index + 1;
  while (n > 0) {
    const rem = (n - 1) % 26;
    name = String.fromCharCode(65 + rem) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
}

class ZipReader {
  constructor(bytes, entries) {
    this.bytes = bytes;
    this.entries = entries;
  }

  static async fromArrayBuffer(buffer) {
    const bytes = new Uint8Array(buffer);
    const entries = new Map();
    let eocd = -1;
    for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 66000); i--) {
      if (readU32(bytes, i) === 0x06054b50) {
        eocd = i;
        break;
      }
    }
    if (eocd < 0) throw new Error("不是有效的 xlsx 文件");
    const totalEntries = readU16(bytes, eocd + 10);
    let centralOffset = readU32(bytes, eocd + 16);
    for (let i = 0; i < totalEntries; i++) {
      if (readU32(bytes, centralOffset) !== 0x02014b50) throw new Error("xlsx 中央目录损坏");
      const compression = readU16(bytes, centralOffset + 10);
      const compressedSize = readU32(bytes, centralOffset + 20);
      const uncompressedSize = readU32(bytes, centralOffset + 24);
      const fileNameLength = readU16(bytes, centralOffset + 28);
      const extraLength = readU16(bytes, centralOffset + 30);
      const commentLength = readU16(bytes, centralOffset + 32);
      const localOffset = readU32(bytes, centralOffset + 42);
      const name = decodeUtf8(bytes.slice(centralOffset + 46, centralOffset + 46 + fileNameLength));
      entries.set(name, { compression, compressedSize, uncompressedSize, localOffset });
      centralOffset += 46 + fileNameLength + extraLength + commentLength;
    }
    return new ZipReader(bytes, entries);
  }

  has(path) {
    return this.entries.has(path);
  }

  async text(path) {
    return decodeUtf8(await this.file(path));
  }

  async file(path) {
    const entry = this.entries.get(path);
    if (!entry) throw new Error(`xlsx 缺少文件：${path}`);
    const offset = entry.localOffset;
    if (readU32(this.bytes, offset) !== 0x04034b50) throw new Error("xlsx 本地文件头损坏");
    const fileNameLength = readU16(this.bytes, offset + 26);
    const extraLength = readU16(this.bytes, offset + 28);
    const start = offset + 30 + fileNameLength + extraLength;
    const compressed = this.bytes.slice(start, start + entry.compressedSize);
    if (entry.compression === 0) return compressed;
    if (entry.compression === 8) {
      const stream = new Blob([compressed]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
      return new Uint8Array(await new Response(stream).arrayBuffer());
    }
    throw new Error(`不支持的 xlsx 压缩方式：${entry.compression}`);
  }
}

function writeXlsx(sheets) {
  const files = new Map();
  files.set("[Content_Types].xml", contentTypesXml(sheets.length));
  files.set("_rels/.rels", rootRelsXml());
  files.set("docProps/app.xml", appXml());
  files.set("docProps/core.xml", coreXml());
  files.set("xl/workbook.xml", workbookXml(sheets));
  files.set("xl/_rels/workbook.xml.rels", workbookRelsXml(sheets.length));
  files.set("xl/styles.xml", stylesXml());
  sheets.forEach((sheet, index) => {
    files.set(`xl/worksheets/sheet${index + 1}.xml`, worksheetXml(sheet.rows, index === 0));
  });
  return ZipWriter.write(files);
}

class ZipWriter {
  static write(files) {
    const encoder = new TextEncoder();
    const localParts = [];
    const centralParts = [];
    let offset = 0;
    for (const [name, content] of files) {
      const nameBytes = encoder.encode(name);
      const data = typeof content === "string" ? encoder.encode(content) : content;
      const crc = crc32(data);
      const local = new Uint8Array(30 + nameBytes.length);
      writeU32(local, 0, 0x04034b50);
      writeU16(local, 4, 20);
      writeU16(local, 6, 0);
      writeU16(local, 8, 0);
      writeU32(local, 14, crc);
      writeU32(local, 18, data.length);
      writeU32(local, 22, data.length);
      writeU16(local, 26, nameBytes.length);
      local.set(nameBytes, 30);
      localParts.push(local, data);

      const central = new Uint8Array(46 + nameBytes.length);
      writeU32(central, 0, 0x02014b50);
      writeU16(central, 4, 20);
      writeU16(central, 6, 20);
      writeU16(central, 10, 0);
      writeU32(central, 16, crc);
      writeU32(central, 20, data.length);
      writeU32(central, 24, data.length);
      writeU16(central, 28, nameBytes.length);
      writeU32(central, 42, offset);
      central.set(nameBytes, 46);
      centralParts.push(central);
      offset += local.length + data.length;
    }
    const centralOffset = offset;
    const centralSize = centralParts.reduce((sum, part) => sum + part.length, 0);
    const eocd = new Uint8Array(22);
    writeU32(eocd, 0, 0x06054b50);
    writeU16(eocd, 8, files.size);
    writeU16(eocd, 10, files.size);
    writeU32(eocd, 12, centralSize);
    writeU32(eocd, 16, centralOffset);
    return concatBytes([...localParts, ...centralParts, eocd]);
  }
}

function worksheetXml(rows, markOpenCases) {
  const maxColumns = rows.reduce((max, row) => Math.max(max, row.length), 1);
  const xmlRows = rows.map((row, rIndex) => {
    const style = rIndex === 0 || (rIndex === 2 && markOpenCases) ? 1 : 0;
    const rowStyle = markOpenCases && rIndex > 2 && row[0] === "未结案" ? 2 : style;
    const cells = Array.from({ length: maxColumns }, (_, cIndex) => {
      const value = row[cIndex] ?? "";
      const ref = `${indexToColumnName(cIndex)}${rIndex + 1}`;
      const cellStyle = rowStyle ? ` s="${rowStyle}"` : "";
      return `<c r="${ref}" t="inlineStr"${cellStyle}><is><t>${escapeXml(value)}</t></is></c>`;
    }).join("");
    return `<row r="${rIndex + 1}">${cells}</row>`;
  }).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="${SPREADSHEET_NS}" xmlns:r="${REL_NS}"><sheetData>${xmlRows}</sheetData></worksheet>`;
}

function workbookXml(sheets) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="${SPREADSHEET_NS}" xmlns:r="${REL_NS}"><sheets>${sheets
    .map((sheet, index) => `<sheet name="${escapeXml(sheet.name)}" sheetId="${index + 1}" r:id="rId${index + 1}"/>`)
    .join("")}</sheets></workbook>`;
}

function workbookRelsXml(count) {
  const sheetRels = Array.from({ length: count }, (_, index) =>
    `<Relationship Id="rId${index + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${index + 1}.xml"/>`
  ).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${sheetRels}<Relationship Id="rId${count + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`;
}

function contentTypesXml(sheetCount) {
  const sheets = Array.from({ length: sheetCount }, (_, index) =>
    `<Override PartName="/xl/worksheets/sheet${index + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`
  ).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${sheets}<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`;
}

function rootRelsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`;
}

function stylesXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="${SPREADSHEET_NS}"><fonts count="2"><font><sz val="11"/><name val="Microsoft YaHei"/></font><font><b/><sz val="11"/><name val="Microsoft YaHei"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFEFF2EF"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFF3D4"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="1" borderId="0" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`;
}

function appXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Law SCRM</Application></Properties>`;
}

function coreXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:creator>Law SCRM</dc:creator></cp:coreProperties>`;
}

function normalizePath(path) {
  const parts = [];
  for (const part of path.split("/")) {
    if (!part || part === ".") continue;
    if (part === "..") parts.pop();
    else parts.push(part);
  }
  return parts.join("/");
}

function readU16(bytes, offset) {
  return bytes[offset] | (bytes[offset + 1] << 8);
}

function readU32(bytes, offset) {
  return (bytes[offset] | (bytes[offset + 1] << 8) | (bytes[offset + 2] << 16) | (bytes[offset + 3] << 24)) >>> 0;
}

function writeU16(bytes, offset, value) {
  bytes[offset] = value & 255;
  bytes[offset + 1] = (value >>> 8) & 255;
}

function writeU32(bytes, offset, value) {
  bytes[offset] = value & 255;
  bytes[offset + 1] = (value >>> 8) & 255;
  bytes[offset + 2] = (value >>> 16) & 255;
  bytes[offset + 3] = (value >>> 24) & 255;
}

function concatBytes(parts) {
  const total = parts.reduce((sum, part) => sum + part.length, 0);
  const out = new Uint8Array(total);
  let offset = 0;
  for (const part of parts) {
    out.set(part, offset);
    offset += part.length;
  }
  return out;
}

function crc32(data) {
  let crc = -1;
  for (const byte of data) {
    crc = (crc >>> 8) ^ CRC_TABLE[(crc ^ byte) & 255];
  }
  return (crc ^ -1) >>> 0;
}

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let k = 0; k < 8; k++) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    table[i] = c >>> 0;
  }
  return table;
})();

function decodeUtf8(bytes) {
  return new TextDecoder("utf-8").decode(bytes);
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function escapeAttr(value) {
  return escapeHtml(value).replaceAll("'", "&#39;");
}

function escapeXml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

window.lawScrmDebug = {
  state: appState,
  buildManagedWorkbook
};
