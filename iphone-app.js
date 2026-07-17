import {
  buildCorrelation,
  buildLotAssessment,
  buildOriginExport,
  buildOriginExportWide,
  buildParameterTrend,
  buildTrendDateLookup,
  collectParameterRecords,
  findZonedHeaderRow,
  gaussianFitWithOptions,
  getLotValues,
  getRegionalParameters,
  getTrendParameters,
  headerIndex,
  isNumeric,
  meanAndSigma,
  number,
  recommendGaussianSettings,
  text,
  zoneColumns
} from "./analysis.js?v=12";
import {
  buildCleanDataFromAuswertung,
  buildFullSummary,
  classificationMaps,
  correctCleanData
} from "./clean-data.js?v=1";

const state = {
  workbook: null,
  originalData: null,
  workbookName: "",
  sheets: [],
  generatedSources: new Map(),
  source: "",
  headers: [],
  rows: [],
  parameters: [],
  trendParameters: [],
  lots: [],
  types: [],
  lotClassifications: new Map(),
  filterSelections: { lots: new Set(), classification: new Set() },
  filterInitialized: { lots: false, classification: false },
  generatedClean: false,
  lastBuild: null,
  lastGaussian: null,
  gaussianSnapshots: [],
  trendDateLookup: null,
  lastTrend: null,
  lastCorrelation: null,
  lastAssessment: null
};

const ALL = "__all__";
const UNCLASSIFIED = "__unclassified__";
const GENERATED_CLEAN = "Generated Clean_Data";
const GENERATED_CORRECTED = "Generated Clean_Data_Cor";
const ZONES = [1, 2, 3, 4, 5, 6];
const TREND_COLORS = ["#195d8d", "#2e7d4f", "#a86a00", "#b3261e", "#7656a1", "#00838f"];
const PREVIEW_ROWS = 18;
const PREVIEW_COLUMNS = 12;

document.addEventListener("DOMContentLoaded", () => {
  bindEvents();
  configureLocalWorkbookButton();
  registerOfflineApp();
  renderZoneChoices("gaussian-zones", "gaussian-zone");
  renderZoneChoices("export-zones", "export-zone");
  drawEmptyState();
  document.documentElement.dataset.xlsxReady = window.XLSX ? "true" : "false";
  if (!window.XLSX) setStatus("Workbook parser did not load. Reload this page.", true);
});

function configureLocalWorkbookButton() {
  const host = window.location.hostname;
  const privateHost = host === "localhost" || host === "127.0.0.1" || host === "::1" ||
    /^10\./.test(host) || /^192\.168\./.test(host) || /^172\.(?:1[6-9]|2\d|3[01])\./.test(host);
  byId("load-local-workbook").closest(".local-workbook-action").hidden = !privateHost;
}

function registerOfflineApp() {
  if (!("serviceWorker" in navigator) || !window.isSecureContext) return;
  navigator.serviceWorker.register("./service-worker.js", { scope: "./" })
    .then(() => {
      document.documentElement.dataset.offlineReady = "true";
    })
    .catch((error) => {
      console.warn("Offline cache registration failed.", error);
    });
}

function bindEvents() {
  byId("workbook-file").addEventListener("change", (event) => runAction(() => openWorkbook(event.target.files[0])));
  byId("load-local-workbook").addEventListener("click", () => runAction(loadLocalWorkbook));
  byId("source-sheet").addEventListener("change", () => runAction(() => selectSource(byId("source-sheet").value)));
  byId("filter-mode").addEventListener("change", () => {
    invalidateAnalyses();
    renderFilterOptions();
    renderCurrentData();
  });
  byId("type-filter").addEventListener("change", () => {
    invalidateAnalyses();
    renderCurrentData();
  });
  byId("filter-options").addEventListener("change", handleFilterOptionChange);
  byId("select-all-filters").addEventListener("click", () => setAllFilterOptions(true));
  byId("clear-all-filters").addEventListener("click", () => setAllFilterOptions(false));
  byId("clean-output").addEventListener("change", syncCorrectionInputs);
  byId("build-clean-data").addEventListener("click", () => runAction(buildCleanData));
  byId("classification-lot").addEventListener("change", syncClassificationInput);
  byId("apply-classification").addEventListener("click", () => runAction(applyLotClassification));
  byId("classification-value").addEventListener("keydown", (event) => {
    if (event.key === "Enter") runAction(applyLotClassification);
  });
  ["gaussian-data-scope", "trend-data-scope", "correlation-data-scope"].forEach((id) => {
    byId(id).addEventListener("change", invalidateAnalyses);
  });
  byId("assessment-data-scope").addEventListener("change", () => {
    state.lastAssessment = null;
    syncAssessmentLots();
    clearResult("assessment-result");
  });
  byId("summary-parameter").addEventListener("change", renderCurrentData);
  byId("generated-summary-parameter").addEventListener("change", renderGeneratedSummaryTable);
  byId("gaussian-parameter").addEventListener("change", () => {
    invalidateGaussian();
    syncGaussianMethod();
    syncZoneChoices("gaussian-zones", byId("gaussian-parameter").value);
  });
  byId("gaussian-method").addEventListener("change", invalidateGaussian);
  ["gaussian-bin-width", "gaussian-start", "gaussian-end"].forEach((id) => {
    byId(id).addEventListener("input", invalidateGaussian);
  });
  ["trend-parameter", "trend-start", "trend-end", "trend-window"].forEach((id) => {
    byId(id).addEventListener(id === "trend-parameter" ? "change" : "input", invalidateTrend);
  });
  byId("correlation-y").addEventListener("change", syncCorrelationZones);
  byId("correlation-x").addEventListener("change", syncCorrelationZones);
  byId("export-parameter").addEventListener("change", () => {
    syncZoneChoices("export-zones", byId("export-parameter").value);
    renderExportNote();
  });
  byId("recommend-gaussian").addEventListener("click", () => runAction(recommendGaussian));
  byId("run-gaussian").addEventListener("click", () => runAction(createGaussian));
  byId("save-gaussian-snapshot").addEventListener("click", () => runAction(saveGaussianSnapshot));
  byId("run-trend").addEventListener("click", () => runAction(createTrend));
  byId("run-assessment").addEventListener("click", () => runAction(createAssessment));
  byId("run-correlation").addEventListener("click", () => runAction(createCorrelation));
  byId("download-summary").addEventListener("click", () => runAction(downloadSummaryCsv));
  byId("download-origin").addEventListener("click", () => runAction(downloadOriginCsv));
  byId("download-workbook").addEventListener("click", () => runAction(downloadAnalysisWorkbook));
  document.querySelectorAll(".tab").forEach((tab) => {
    tab.addEventListener("click", () => openPanel(tab.dataset.panel));
  });
  window.addEventListener("resize", debounce(redrawCharts, 120));
}

async function openWorkbook(file) {
  if (!file) return;
  setStatus(`Checking ${file.name}...`);
  const data = await file.arrayBuffer();
  await parseWorkbook(data, file.name);
}

async function loadLocalWorkbook() {
  setStatus("Loading Test.xlsx from this Mac...");
  const response = await fetch("./data/Test.xlsx", { cache: "no-store" });
  if (!response.ok) throw new Error("Test.xlsx is not available from this Mac.");
  await parseWorkbook(await response.arrayBuffer(), "Test.xlsx");
}

async function parseWorkbook(data, fileName) {
  const XLSX = getXlsx();
  state.workbookName = fileName;
  state.originalData = data.slice(0);
  state.generatedSources = new Map();
  state.lotClassifications = new Map();
  state.filterSelections = { lots: new Set(), classification: new Set() };
  state.filterInitialized = { lots: false, classification: false };
  state.lastBuild = null;
  state.gaussianSnapshots = [];
  state.trendDateLookup = null;
  state.lastTrend = null;
  setStatus("Reading Clean_Data only...");
  await yieldToBrowser();
  state.workbook = XLSX.read(data, {
    type: "array",
    cellDates: true,
    cellFormula: true,
    cellHTML: false,
    cellText: true,
    dense: false,
    sheets: ["Clean_Data", "Clean_Data_Cor"]
  });
  state.sheets = ["Clean_Data", "Clean_Data_Cor"].filter((name) => state.workbook.Sheets[name]);
  const sourceName = chooseCleanSource(state.sheets);
  refreshSourceSelect(sourceName);
  await selectSource(sourceName);
}

async function selectSource(sheetName) {
  if (!state.workbook) return;
  setStatus(`Reading ${sheetName}...`);
  const source = prepareSource(sheetName);
  state.source = sheetName;
  state.headers = source.headers;
  state.rows = source.rows;
  state.generatedClean = source.generatedClean;
  ensureClassificationColumn(state.headers, state.rows);
  seedLotClassifications(state.headers, state.rows);
  applyClassificationOverrides(state.headers, state.rows);
  state.parameters = getRegionalParameters(state.headers);
  state.trendParameters = getTrendParameters(state.headers);
  if (!state.parameters.length) throw new Error("Clean_Data has no named parameter and Zone headers.");
  const lotColumn = headerIndex(state.headers, "Lot");
  state.lots = lotColumn >= 0 ? getLotValues(dataRows(), lotColumn) : [];
  state.types = getTypeValues();
  populateWorkbookControls();
  state.lastGaussian = null;
  state.lastTrend = null;
  state.lastCorrelation = null;
  state.lastAssessment = null;
  renderCurrentData();
  setStatus(`loaded: ${formatInteger(dataRows().length)} data rows, ${formatInteger(state.parameters.length)} parameters.`, false, true);
}

function prepareSource(sheetName) {
  const table = state.generatedSources.get(sheetName) || readSheet(sheetName);
  if (!table.length) throw new Error(`${sheetName} has no readable rows.`);
  const headerRow = findZonedHeaderRow(table);
  if (!headerRow) throw new Error(`${state.workbookName}: ${sheetName} has no named headers such as Thick_1 through Thick_6.`);
  const headers = headerRow.headers;
  const rows = normalizeRows(table.slice(headerRow.index + 1), headers.length);
  return { headers, rows, generatedClean: state.generatedSources.has(sheetName) };
}

function readSheet(sheetName) {
  return readWorkbookSheet(state.workbook, sheetName);
}

function readWorkbookSheet(workbook, sheetName) {
  const XLSX = getXlsx();
  const sheet = workbook?.Sheets?.[sheetName];
  if (!sheet) return [];
  const table = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "", blankrows: false });
  // The wrapper prevents Array.map from passing the row index as normalizeRow's width.
  const trimmed = table.map((row) => normalizeRow(row)).filter((row) => !isEmptyRow(row));
  const width = trimmed.reduce((max, row) => Math.max(max, row.length), 0);
  return trimmed.map((row) => normalizeRow(row, width));
}

async function buildCleanData() {
  if (!state.originalData) throw new Error("Open a workbook first.");
  setStatus("Reading Auswertung and building clean data...");
  await yieldToBrowser();
  const XLSX = getXlsx();
  const rawWorkbook = XLSX.read(state.originalData.slice(0), {
    type: "array",
    cellDates: true,
    cellFormula: true,
    cellHTML: false,
    cellText: true,
    dense: false,
    sheets: ["Auswertung"]
  });
  const rawTable = readWorkbookSheet(rawWorkbook, "Auswertung");
  if (!rawTable.length) throw new Error("This workbook does not contain a readable Auswertung sheet.");

  const preserved = classificationMaps(state.headers, dataRows());
  const clean = buildCleanDataFromAuswertung(rawTable, {
    classificationByLot: state.lotClassifications,
    classificationByKey: preserved.byKey
  });
  state.generatedSources.set(GENERATED_CLEAN, [clean.headers, ...clean.rows]);

  const correctedRequested = byId("clean-output").value === "corrected";
  let output = clean;
  let sourceName = GENERATED_CLEAN;
  if (correctedRequested) {
    output = correctCleanData(
      clean.headers,
      clean.rows,
      requiredNumber("reference-temperature"),
      requiredNumber("reference-humidity")
    );
    sourceName = GENERATED_CORRECTED;
    state.generatedSources.set(sourceName, [output.headers, ...output.rows]);
  }

  const summary = buildFullSummary(output.headers, output.rows);
  const lotIndex = headerIndex(output.headers, "Lot");
  state.lastBuild = {
    sourceName,
    rows: output.rows.length,
    lots: lotIndex >= 0 ? getLotValues(output.rows, lotIndex).length : 0,
    parameters: getRegionalParameters(output.headers).length,
    summaryRows: summary.rows.length,
    removedProbeRows: clean.removedProbeRows,
    removedVeRows: clean.removedVeRows,
    removedVeLots: clean.removedVeLots,
    corrected: correctedRequested
  };
  refreshSourceSelect(sourceName);
  await selectSource(sourceName);
  openPanel("clean-panel");
  renderBuildResult();
  setStatus(
    `${sourceName} and Summary created locally. Removed ${formatInteger(clean.removedProbeRows)} Probe rows and ${formatInteger(clean.removedVeLots)} VE lots.`,
    false,
    true
  );
}

function refreshSourceSelect(preferredValue = state.source) {
  const sources = [...state.sheets, ...state.generatedSources.keys()];
  fillSelect(byId("source-sheet"), sources, preferredValue);
  byId("source-sheet").disabled = sources.length < 2;
}

function ensureClassificationColumn(headers, rows) {
  if (headerIndex(headers, "Classification") >= 0) return;
  const nIndex = headerIndex(headers, "N");
  const insertAt = nIndex >= 0 ? nIndex : 0;
  headers.splice(insertAt, 0, "Classification");
  rows.forEach((row) => row.splice(insertAt, 0, ""));
}

function seedLotClassifications(headers, rows) {
  const maps = classificationMaps(headers, rows);
  for (const [lot, classification] of maps.byLot) {
    if (!state.lotClassifications.has(lot)) state.lotClassifications.set(lot, classification);
  }
}

function applyClassificationOverrides(headers, rows) {
  const lotIndex = headerIndex(headers, "Lot");
  const classificationIndex = headerIndex(headers, "Classification");
  if (lotIndex < 0 || classificationIndex < 0) return;
  rows.forEach((row) => {
    const lot = text(row[lotIndex]);
    if (state.lotClassifications.has(lot)) row[classificationIndex] = state.lotClassifications.get(lot);
  });
}

function applyLotClassification() {
  const lot = byId("classification-lot").value;
  if (!lot) throw new Error("Select a lot to classify.");
  const classification = byId("classification-value").value.trim();
  state.lotClassifications.set(lot, classification);
  applyClassificationOverrides(state.headers, state.rows);
  for (const table of state.generatedSources.values()) {
    applyClassificationOverrides(table[0], table.slice(1));
  }
  state.filterInitialized.classification = false;
  syncFilterSelections();
  invalidateAnalyses();
  renderFilterOptions();
  renderCurrentData();
  renderClassificationTable();
  setStatus(`Classification ${classification || "cleared"} for lot ${lot}.`, false, true);
}

function syncClassificationInput() {
  const lot = byId("classification-lot").value;
  if (!lot) {
    byId("classification-value").value = "";
    return;
  }
  const classificationIndex = headerIndex(state.headers, "Classification");
  const lotIndex = headerIndex(state.headers, "Lot");
  const sourceValue = state.rows.find((row) => text(row[lotIndex]) === lot)?.[classificationIndex];
  byId("classification-value").value = state.lotClassifications.has(lot)
    ? state.lotClassifications.get(lot)
    : text(sourceValue);
}

function syncCorrectionInputs() {
  const enabled = byId("clean-output").value === "corrected" && !byId("clean-output").disabled;
  byId("reference-temperature").disabled = !enabled;
  byId("reference-humidity").disabled = !enabled;
}

function populateWorkbookControls() {
  const parameter = state.parameters[0];
  const secondParameter = state.parameters[1] || state.parameters[0];
  fillSelect(byId("type-filter"), [ALL, ...state.types], ALL, (value) => value === ALL ? "All types" : value);
  fillSelect(byId("classification-lot"), state.lots, byId("classification-lot").value || state.lots[0]);
  [
    "summary-parameter",
    "generated-summary-parameter",
    "gaussian-parameter",
    "trend-parameter",
    "assessment-parameter",
    "correlation-y",
    "correlation-x",
    "export-parameter"
  ].forEach((id) => fillSelect(byId(id), state.parameters, byId(id).value || parameter));
  fillSelect(byId("trend-parameter"), state.trendParameters, byId("trend-parameter").value || state.trendParameters[0]);
  byId("generated-summary-parameter").disabled = !state.lastBuild;
  if (byId("correlation-x").options.length > 1) byId("correlation-x").value = secondParameter;
  enableControls([
    "filter-mode",
    "type-filter",
    "clean-output",
    "reference-temperature",
    "reference-humidity",
    "build-clean-data",
    "classification-lot",
    "classification-value",
    "apply-classification",
    "summary-parameter",
    "gaussian-data-scope",
    "gaussian-parameter",
    "gaussian-method",
    "gaussian-bin-width",
    "gaussian-start",
    "gaussian-end",
    "recommend-gaussian",
    "run-gaussian",
    "trend-data-scope",
    "trend-parameter",
    "trend-start",
    "trend-end",
    "trend-window",
    "run-trend",
    "assessment-data-scope",
    "assessment-parameter",
    "assessment-lot",
    "assessment-reference",
    "assessment-monitor",
    "assessment-outlier",
    "assessment-mu",
    "assessment-sigma",
    "run-assessment",
    "correlation-data-scope",
    "correlation-y",
    "correlation-x",
    "correlation-scope",
    "correlation-outliers",
    "correlation-removal",
    "run-correlation",
    "export-parameter",
    "download-summary",
    "download-origin",
    "download-workbook"
  ]);
  byId("trend-parameter").disabled = !state.trendParameters.length;
  byId("run-trend").disabled = !state.trendParameters.length;
  byId("save-gaussian-snapshot").disabled = !state.lastGaussian;
  syncGaussianMethod();
  syncZoneChoices("gaussian-zones", byId("gaussian-parameter").value);
  syncZoneChoices("export-zones", byId("export-parameter").value);
  syncCorrelationZones();
  syncCorrectionInputs();
  syncClassificationInput();
  syncFilterSelections();
  renderFilterOptions();
  syncAssessmentLots();
}

function renderCurrentData() {
  if (!state.headers.length) return;
  syncAssessmentLots();
  renderDataMetrics();
  renderSummaryTable();
  renderPreviewTable();
  renderClassificationTable();
  renderBuildResult();
  renderGeneratedSummaryTable();
  clearResult("gaussian-result");
  clearResult("trend-result");
  clearResult("correlation-result");
  clearResult("assessment-result");
  renderExportNote();
  renderGaussianSnapshots();
}

function invalidateAnalyses() {
  invalidateGaussian();
  invalidateTrend();
  state.lastCorrelation = null;
  state.lastAssessment = null;
}

function invalidateGaussian() {
  state.lastGaussian = null;
  byId("save-gaussian-snapshot").disabled = true;
  clearResult("gaussian-result");
}

function invalidateTrend() {
  state.lastTrend = null;
  clearResult("trend-result");
}

function handleFilterOptionChange(event) {
  if (!event.target.matches('input[type="checkbox"]')) return;
  const mode = byId("filter-mode").value;
  if (mode !== "lots" && mode !== "classification") return;
  const selection = state.filterSelections[mode];
  if (event.target.checked) selection.add(event.target.value);
  else selection.delete(event.target.value);
  state.filterInitialized[mode] = true;
  invalidateAnalyses();
  renderCurrentData();
}

function setAllFilterOptions(checked) {
  const mode = byId("filter-mode").value;
  if (mode !== "lots" && mode !== "classification") return;
  state.filterSelections[mode] = checked ? new Set(filterValues(mode)) : new Set();
  state.filterInitialized[mode] = true;
  renderFilterOptions();
  invalidateAnalyses();
  renderCurrentData();
}

function syncFilterSelections() {
  for (const mode of ["lots", "classification"]) {
    const available = filterValues(mode);
    if (!state.filterInitialized[mode]) {
      state.filterSelections[mode] = new Set(available);
      state.filterInitialized[mode] = true;
      continue;
    }
    state.filterSelections[mode] = new Set(
      [...state.filterSelections[mode]].filter((value) => available.includes(value))
    );
  }
}

function filterValues(mode) {
  if (mode === "lots") return [...state.lots];
  if (mode !== "classification") return [];
  const classificationIndex = headerIndex(state.headers, "Classification");
  if (classificationIndex < 0) return [UNCLASSIFIED];
  const values = new Set();
  let hasUnclassified = false;
  for (const row of dataRows()) {
    const value = text(row[classificationIndex]);
    if (value) values.add(value);
    else hasUnclassified = true;
  }
  return [
    ...(hasUnclassified ? [UNCLASSIFIED] : []),
    ...[...values].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }))
  ];
}

function renderFilterOptions() {
  const mode = byId("filter-mode").value;
  const panel = byId("filter-options-panel");
  const container = byId("filter-options");
  if (mode !== "lots" && mode !== "classification") {
    panel.hidden = true;
    container.replaceChildren();
    return;
  }
  syncFilterSelections();
  panel.hidden = false;
  byId("filter-options-title").textContent = mode === "lots" ? "Lots" : "Classifications";
  container.replaceChildren();
  for (const value of filterValues(mode)) {
    const label = document.createElement("label");
    const input = document.createElement("input");
    const caption = document.createElement("span");
    input.type = "checkbox";
    input.value = value;
    input.checked = state.filterSelections[mode].has(value);
    caption.textContent = value === UNCLASSIFIED ? "Unclassified" : value;
    label.append(input, caption);
    container.append(label);
  }
}

function renderClassificationTable() {
  if (!state.lots.length) {
    byId("classification-table").innerHTML = '<p class="empty-state">No lots.</p>';
    return;
  }
  const lotIndex = headerIndex(state.headers, "Lot");
  const classificationIndex = headerIndex(state.headers, "Classification");
  const rows = state.lots.map((lot) => {
    const sourceValue = state.rows.find((row) => text(row[lotIndex]) === lot)?.[classificationIndex];
    const classification = state.lotClassifications.has(lot) ? state.lotClassifications.get(lot) : text(sourceValue);
    return [lot, classification || "Unclassified"];
  });
  byId("classification-table").innerHTML = renderTable([["Lot", "Classification"], ...rows]);
}

function renderBuildResult() {
  if (!state.lastBuild) {
    byId("build-result").innerHTML = '<p class="empty-state">No generated data.</p>';
    return;
  }
  const result = state.lastBuild;
  byId("build-result").innerHTML = `
    <div class="metric-grid">
      ${metric("Rows", formatInteger(result.rows))}
      ${metric("Lots", formatInteger(result.lots))}
      ${metric("Parameters", formatInteger(result.parameters))}
      ${metric("Summary lots", formatInteger(result.summaryRows))}
      ${metric("Probe rows removed", formatInteger(result.removedProbeRows))}
      ${metric("VE lots removed", formatInteger(result.removedVeLots))}
    </div>
  `;
}

function renderGeneratedSummaryTable() {
  const container = byId("generated-summary-table");
  const source = state.lastBuild ? state.generatedSources.get(state.lastBuild.sourceName) : null;
  if (!source) {
    byId("generated-summary-parameter").disabled = true;
    container.innerHTML = '<p class="empty-state">No generated summary.</p>';
    return;
  }
  const headers = source[0];
  const rows = source.slice(1).filter((row) => isDataRow(headers, row));
  const parameters = getRegionalParameters(headers);
  const select = byId("generated-summary-parameter");
  if (!parameters.includes(select.value)) fillSelect(select, parameters, parameters[0]);
  select.disabled = false;
  container.innerHTML = renderTable(
    buildParameterSummaryRows(headers, rows, select.value || parameters[0]),
    { empty: "No generated summary rows." }
  );
}

function renderDataMetrics() {
  const rows = filteredRows();
  const lotColumn = headerIndex(state.headers, "Lot");
  const lotCount = lotColumn >= 0 ? new Set(rows.map((row) => text(row[lotColumn])).filter(Boolean)).size : 0;
  byId("data-metrics").innerHTML = [
    metric("Workbook", state.workbookName || "-"),
    metric("Rows", formatInteger(rows.length)),
    metric("Lots", formatInteger(lotCount)),
    metric("Parameters", formatInteger(state.parameters.length))
  ].join("");
}

function renderSummaryTable() {
  const parameter = byId("summary-parameter").value || state.parameters[0];
  const rows = buildSummaryRows(parameter);
  byId("summary-table").innerHTML = renderTable(rows, { empty: "No summary rows." });
}

function renderPreviewTable() {
  const rows = filteredRows().slice(0, PREVIEW_ROWS);
  const headers = state.headers.slice(0, PREVIEW_COLUMNS);
  const body = rows.map((row) => row.slice(0, PREVIEW_COLUMNS));
  byId("preview-table").innerHTML = renderTable([headers, ...body], { empty: "No data rows." });
}

function buildSummaryRows(parameter) {
  return buildParameterSummaryRows(state.headers, filteredRows(), parameter);
}

function buildParameterSummaryRows(headers, rows, parameter) {
  const lotColumn = headerIndex(headers, "Lot");
  const batchColumn = headerIndex(headers, "N");
  const typeColumn = headerIndex(headers, "Type");
  const classificationColumn = headerIndex(headers, "Classification");
  const zoneIndexes = ZONES.map((zone) => headerIndex(headers, `${parameter}_${zone}`)).filter((index) => index >= 0);
  const outputHeaders = ["Lot", "Classification", "Type", "Rows", "Batches", "Mean", "Sigma", "Min", "Max"];
  if (lotColumn < 0 || !zoneIndexes.length) return [outputHeaders];
  const grouped = new Map();
  for (const row of rows) {
    const lot = text(row[lotColumn]);
    if (!lot) continue;
    if (!grouped.has(lot)) grouped.set(lot, { classification: "", type: "", rows: 0, batches: new Set(), values: [] });
    const group = grouped.get(lot);
    group.rows += 1;
    if (!group.classification && classificationColumn >= 0) group.classification = text(row[classificationColumn]);
    if (!group.type && typeColumn >= 0) group.type = text(row[typeColumn]);
    if (batchColumn >= 0 && text(row[batchColumn])) group.batches.add(text(row[batchColumn]));
    for (const index of zoneIndexes) {
      if (isNumeric(row[index])) group.values.push(number(row[index]));
    }
  }
  const output = [outputHeaders];
  for (const [lot, group] of [...grouped.entries()].sort((a, b) => a[0].localeCompare(b[0], undefined, { numeric: true }))) {
    const stats = meanAndSigma(group.values);
    output.push([
      lot,
      group.classification || "Unclassified",
      group.type,
      group.rows,
      group.batches.size || "",
      stats.n ? stats.mean : "",
      stats.n ? stats.sigma : "",
      stats.n ? stats.min : "",
      stats.n ? stats.max : ""
    ]);
  }
  return output;
}

function recommendGaussian() {
  const parameter = byId("gaussian-parameter").value;
  const includedZones = selectedZones("gaussian-zones");
  if (!includedZones.length) throw new Error("Select at least one Zone.");
  const method = byId("gaussian-method").value;
  if (method === "nacl-truncated" && parameter.toLowerCase() !== "nacl") {
    throw new Error("NaCl truncated fitting is available only for the NaCl parameter.");
  }
  const records = collectParameterRecords(state.headers, rowsForAnalysis("gaussian-data-scope"), parameter, includedZones);
  const settings = recommendGaussianSettings(records.map((record) => record.value), { method, lowerLimit: 1 });
  byId("gaussian-bin-width").value = inputNumber(settings.binWidth);
  byId("gaussian-start").value = inputNumber(settings.start);
  byId("gaussian-end").value = inputNumber(settings.end);
  invalidateGaussian();
  setStatus(
    `Recommended from ${formatInteger(settings.n)} points: width ${formatNumber(settings.binWidth, 4)}, range ${formatNumber(settings.start, 4)} to ${formatNumber(settings.end, 4)}.`,
    false,
    true
  );
}

function createGaussian() {
  const parameter = byId("gaussian-parameter").value;
  const includedZones = selectedZones("gaussian-zones");
  if (!includedZones.length) throw new Error("Select at least one Zone.");
  const method = byId("gaussian-method").value;
  if (method === "nacl-truncated" && parameter.toLowerCase() !== "nacl") {
    throw new Error("NaCl truncated fitting is available only for the NaCl parameter.");
  }
  const scope = byId("gaussian-data-scope").value;
  const records = collectParameterRecords(state.headers, rowsForAnalysis("gaussian-data-scope"), parameter, includedZones);
  const fit = gaussianFitWithOptions(
    records.map((record) => record.value),
    optionalNumber("gaussian-bin-width"),
    optionalNumber("gaussian-start"),
    optionalNumber("gaussian-end"),
    { method, lowerLimit: 1 }
  );
  const fittedRecords = records.filter((record) => record.value >= fit.start && record.value <= fit.end &&
    (fit.method !== "nacl-truncated" || record.value >= fit.lowerLimit));
  const lots = new Set(fittedRecords.map((record) => record.lot).filter(Boolean));
  const batches = new Set(fittedRecords.map((record) => `${record.lot}|${record.batch}`).filter((key) => key !== "|"));
  state.lastGaussian = { parameter, zones: includedZones, scope, fit: { ...fit, lotCount: lots.size, batchCount: batches.size } };
  renderGaussianResult();
  byId("save-gaussian-snapshot").disabled = false;
  setStatus(
    `Gaussian fit: ${formatInteger(fit.n)} points used, ${formatInteger(fit.excluded)} excluded, ${formatInteger(lots.size)} lots.`,
    false,
    true
  );
}

function renderGaussianResult() {
  const result = state.lastGaussian;
  if (!result) return;
  const fit = result.fit;
  byId("gaussian-result").innerHTML = `
    <div class="metric-grid">
      ${metric("Points used", formatInteger(fit.n))}
      ${metric("Points excluded", formatInteger(fit.excluded))}
      ${metric("Lots used", formatInteger(fit.lotCount))}
      ${metric("Batches used", formatInteger(fit.batchCount))}
      ${metric("Mean", formatNumber(fit.mean, 4))}
      ${metric("Sigma", formatNumber(fit.sigma, 4))}
      ${metric("SSE", formatNumber(fit.sse, 2))}
      ${metric("Fit range", `${formatNumber(fit.start, 3)} to ${formatNumber(fit.end, 3)}`)}
    </div>
    <div class="chart-card"><canvas id="gaussian-chart" aria-label="Observed histogram with Gaussian curve"></canvas></div>
    <div class="table-wrap mini-table cutoff-table">${renderTable([
      ["Cutoff", "Low", "High"],
      ["2.5%", fit.low25, fit.high25],
      ["15%", fit.low15, fit.high15]
    ])}</div>
    <div class="table-wrap mini-table">${renderTable([
      ["Bin", "Observed", "Gaussian"],
      ...fit.bins.slice(0, 40).map((bin) => [bin.center, bin.observed, bin.gaussian])
    ])}</div>
  `;
  requestAnimationFrame(() => drawGaussian(byId("gaussian-chart"), fit));
}

function saveGaussianSnapshot() {
  const current = state.lastGaussian;
  const canvas = byId("gaussian-chart");
  if (!current || !canvas) throw new Error("Run a Gaussian fit before saving a snapshot.");
  const stamp = new Date();
  const fileName = `${baseFileName()}_${safeFilePart(current.parameter)}_Gaussian_${fileDateStamp(stamp)}.png`;
  const link = document.createElement("a");
  link.href = canvas.toDataURL("image/png");
  link.download = fileName;
  document.body.append(link);
  link.click();
  link.remove();
  state.gaussianSnapshots.push({
    fileName,
    savedAt: stamp.toISOString(),
    parameter: current.parameter,
    zones: current.zones.join(", "),
    scope: current.scope,
    n: current.fit.n,
    mean: current.fit.mean,
    sigma: current.fit.sigma
  });
  renderGaussianSnapshots();
  setStatus(`Saved ${fileName}.`, false, true);
}

function renderGaussianSnapshots() {
  const container = byId("gaussian-snapshots");
  if (!container) return;
  if (!state.gaussianSnapshots.length) {
    container.innerHTML = "";
    return;
  }
  container.innerHTML = `<div class="section-heading"><h2>Saved Snapshots</h2></div>
    <div class="table-wrap snapshot-table">${renderTable([
      ["Saved", "Parameter", "Zones", "File"],
      ...state.gaussianSnapshots.map((item) => [formatDateTime(item.savedAt), item.parameter, item.zones, item.fileName])
    ])}</div>`;
}

async function createTrend() {
  const parameter = byId("trend-parameter").value;
  if (!parameter) throw new Error("Select a trend parameter.");
  const dateLookup = await loadTrendDateLookup();
  const scope = byId("trend-data-scope").value;
  const result = buildParameterTrend(
    state.headers,
    rowsForAnalysis("trend-data-scope"),
    parameter,
    dateLookup,
    {
      startDate: byId("trend-start").value,
      endDate: byId("trend-end").value,
      rollingWindow: requiredNumber("trend-window")
    }
  );
  state.lastTrend = { scope, result };
  renderTrendResult();
  setStatus(
    `Parameter trend: ${formatInteger(result.batches.length)} batches, ${formatInteger(result.lots.length)} lots, ${formatInteger(result.fallbackCount)} fallback dates.`,
    false,
    true
  );
}

async function loadTrendDateLookup() {
  if (state.trendDateLookup instanceof Map) return state.trendDateLookup;
  if (!state.originalData) throw new Error("Open a workbook first.");
  setStatus("Reading Auswertung production dates...");
  await yieldToBrowser();
  const XLSX = getXlsx();
  const rawWorkbook = XLSX.read(state.originalData.slice(0), {
    type: "array",
    cellDates: true,
    cellFormula: true,
    cellHTML: false,
    cellText: true,
    dense: false,
    sheets: ["Auswertung"]
  });
  const rawTable = readWorkbookSheet(rawWorkbook, "Auswertung");
  state.trendDateLookup = buildTrendDateLookup(rawTable);
  return state.trendDateLookup;
}

function renderTrendResult() {
  const current = state.lastTrend;
  if (!current) return;
  const trend = current.result;
  const dateSource = trend.lookupCount
    ? trend.fallbackCount ? "Auswertung + fallback" : "Auswertung"
    : "Lot-code fallback";
  const regionalCharts = trend.regional ? `
    <article class="chart-card">
      <h3>Zone Values</h3>
      ${chartLegend(ZONES.map((zone, index) => [`Zone ${zone}`, TREND_COLORS[index]]))}
      <canvas id="trend-zone-chart" aria-label="Zone values over production time"></canvas>
    </article>
    <article class="chart-card">
      <h3>Persistent Zone Bias</h3>
      <canvas id="trend-bias-chart" aria-label="Zone value minus batch mean"></canvas>
    </article>` : "";
  const biasTable = trend.regional ? `
    <div class="section-heading"><h2>Zone Bias</h2></div>
    <div class="table-wrap bias-table">${renderTable([
      ["Zone", "Batches", "Mean value", "Mean bias", "Bias SD", "95% CI low", "95% CI high", "Signal"],
      ...trend.zoneBias.map((item) => [
        `Zone ${item.zone}`, item.n, item.meanValue, item.meanBias, item.biasSigma, item.ciLow, item.ciHigh, item.signal
      ])
    ])}</div>` : "";
  byId("trend-result").innerHTML = `
    <div class="metric-grid">
      ${metric("Batches", formatInteger(trend.batches.length))}
      ${metric("Lots", formatInteger(trend.lots.length))}
      ${metric("Fallback dates", formatInteger(trend.fallbackCount))}
      ${metric("Date source", dateSource)}
      ${metric("Lot center", formatNumber(trend.lotCenter, 4))}
      ${metric("Lot sigma", formatNumber(trend.lotSigma, 4))}
      ${metric("Batch center", formatNumber(trend.batchCenter, 4))}
      ${metric("Batch sigma", formatNumber(trend.batchSigma, 4))}
    </div>
    <div class="trend-chart-grid">
      <article class="chart-card">
        <h3>Lot Mean</h3>
        ${chartLegend([["Lot mean", TREND_COLORS[0]], ["Center", TREND_COLORS[1]], ["+/-2 SD", TREND_COLORS[3]]])}
        <canvas id="trend-lot-chart" aria-label="Lot mean over production time"></canvas>
      </article>
      <article class="chart-card">
        <h3>Batch Mean and Rolling Trend</h3>
        ${chartLegend([["Batch", TREND_COLORS[0]], [`Rolling ${trend.rollingWindow}`, TREND_COLORS[2]], ["+/-2 SD", TREND_COLORS[3]]])}
        <canvas id="trend-batch-chart" aria-label="Batch mean and rolling trend over production time"></canvas>
      </article>
      ${regionalCharts}
    </div>
    <div class="section-heading"><h2>Lot Trend Values</h2></div>
    <div class="table-wrap">${renderTable([
      ["Date", "Lot", "Mean", "Sigma", "N values", "Center", "+2 SD", "-2 SD"],
      ...trend.lots.map((item) => [
        formatTrendDate(item.date), item.lot, item.mean, item.sigma, item.n,
        trend.lotCenter, trend.lotCenter + 2 * trend.lotSigma, trend.lotCenter - 2 * trend.lotSigma
      ])
    ])}</div>
    ${biasTable}
  `;
  requestAnimationFrame(drawTrendCharts);
}

function chartLegend(items) {
  return `<div class="chart-key">${items.map(([label, color]) =>
    `<span style="--key-color:${color}">${escapeHtml(label)}</span>`
  ).join("")}</div>`;
}

function createAssessment() {
  const parameter = byId("assessment-parameter").value;
  const lot = byId("assessment-lot").value;
  if (!lot) throw new Error("No lot is available in the selected data scope.");
  const mode = byId("assessment-reference").value;
  const scope = byId("assessment-data-scope").value;
  const assessment = buildLotAssessment(
    state.headers,
    rowsForAnalysis("assessment-data-scope"),
    parameter,
    lot,
    mode,
    optionalNumber("assessment-mu"),
    optionalNumber("assessment-sigma"),
    requiredNumber("assessment-monitor"),
    requiredNumber("assessment-outlier")
  );
  state.lastAssessment = { parameter, lot, mode, scope, assessment };
  renderAssessmentResult();
  setStatus(`Lot assessment: ${assessment.overall}.`, assessment.overall === "OUT OF RANGE", assessment.overall === "OK");
}

function renderAssessmentResult() {
  const result = state.lastAssessment;
  if (!result) return;
  const { parameter, assessment } = result;
  const availableZones = assessment.availableZones;
  const rows = [
    ["Batch N", ...availableZones.map((zone) => `${parameter}_${zone}`)],
    ["Mu", ...availableZones.map((zone) => assessment.references[zone - 1].mean)],
    ["Sigma", ...availableZones.map((zone) => assessment.references[zone - 1].sigma)],
    ...assessment.grid.map((row) => [row.batch, ...availableZones.map((zone) => row.values[zone - 1])])
  ];
  byId("assessment-result").innerHTML = `
    <div class="metric-grid">
      ${metricHtml("Overall", statusBadge(assessment.overall))}
      ${metric("Out of range", formatInteger(assessment.outOfRangeCount))}
      ${metric("Check", formatInteger(assessment.monitorCount))}
      ${metric("No history", formatInteger(assessment.noHistoryCount))}
    </div>
    <div class="table-wrap">${renderAssessmentTable(rows, assessment.grid, availableZones)}</div>
  `;
}

function createCorrelation() {
  const yParameter = byId("correlation-y").value;
  const xParameter = byId("correlation-x").value;
  if (xParameter === yParameter) throw new Error("Select two different parameters.");
  const scope = byId("correlation-data-scope").value;
  const result = buildCorrelation(
    state.headers,
    rowsForAnalysis("correlation-data-scope"),
    xParameter,
    yParameter,
    byId("correlation-scope").value,
    byId("correlation-outliers").value,
    requiredNumber("correlation-removal")
  );
  state.lastCorrelation = { xParameter, yParameter, scope, result };
  renderCorrelationResult();
  setStatus(`Correlation: ${formatInteger(result.rawN)} included pairs.`, false, true);
}

function renderCorrelationResult() {
  const current = state.lastCorrelation;
  if (!current) return;
  const { result } = current;
  const excluded = result.pairs.length - result.included.length;
  byId("correlation-result").innerHTML = `
    <div class="metric-grid">
      ${metric("Pairs", formatInteger(result.rawN))}
      ${metric("Excluded", formatInteger(excluded))}
      ${metric("Pearson r", formatNumber(result.r, 4))}
      ${metric("Slope", formatNumber(result.slope, 4))}
    </div>
    <div class="chart-card"><canvas id="correlation-chart" aria-label="Matched-zone correlation scatter plot"></canvas></div>
    <div class="table-wrap mini-table">${renderTable([
      ["Lot", "Batch N", "Zone", "X", "Y", "Status"],
      ...result.pairs.slice(0, 80).map((pair) => [pair.lot, pair.batch, `Zone ${pair.zone}`, pair.x, pair.y, pair.included ? "Included" : "Excluded"])
    ])}</div>
  `;
  requestAnimationFrame(() => drawScatter(byId("correlation-chart"), current));
}

function renderExportNote() {
  if (!state.parameters.length) return;
  const parameter = byId("export-parameter").value || state.parameters[0];
  const includedZones = selectedZones("export-zones");
  let count = 0;
  try {
    count = Math.max(0, buildOriginExport(state.headers, filteredRows(), parameter, includedZones).length - 1);
  } catch {
    count = 0;
  }
  byId("export-result").innerHTML = `
    <div class="metric-grid">
      ${metric("Filtered rows", formatInteger(filteredRows().length))}
      ${metric("Summary lots", formatInteger(buildFullSummary(state.headers, filteredRows()).rows.length))}
      ${metric("Origin values", formatInteger(count))}
    </div>
  `;
}

function downloadSummaryCsv() {
  const summary = buildFullSummary(state.headers, filteredRows());
  downloadText(`${baseFileName()}_summary.csv`, toCsv([summary.headers, ...summary.rows]), "text/csv;charset=utf-8");
}

function downloadOriginCsv() {
  const rows = buildOriginExport(state.headers, filteredRows(), byId("export-parameter").value, selectedZones("export-zones"));
  downloadText(`${baseFileName()}_origin.csv`, toCsv(rows), "text/csv;charset=utf-8");
}

function downloadAnalysisWorkbook() {
  const XLSX = getXlsx();
  const workbook = XLSX.utils.book_new();
  const exportRows = filteredRows();
  const cleanSheetName = /Clean_Data_Cor/i.test(state.source) ? "Clean_Data_Cor" : "Clean_Data";
  const summary = buildFullSummary(state.headers, exportRows);
  appendSheet(workbook, cleanSheetName, [state.headers, ...exportRows]);
  appendSheet(workbook, "Summary", [summary.headers, ...summary.rows]);
  const parameter = byId("export-parameter").value || state.parameters[0];
  const zones = selectedZones("export-zones");
  appendSheet(workbook, "Origin_Long_App", buildOriginExport(state.headers, filteredRows(), parameter, zones));
  appendSheet(workbook, "Origin_Wide_App", buildOriginExportWide(state.headers, filteredRows(), parameter, zones));
  if (state.lastGaussian) {
    const { parameter: gaussianParameter, zones: gaussianZones, fit } = state.lastGaussian;
    appendSheet(workbook, "GaussianFit_App", [
      ["Parameter", gaussianParameter],
      ["Data scope", state.lastGaussian.scope],
      ["Zones", gaussianZones.map((zone) => `Zone ${zone}`).join(", ")],
      ["Fit N", fit.n],
      ["Visible N", fit.visibleN],
      ["Points excluded", fit.excluded],
      ["Excluded by range", fit.excludedByRange],
      ["Recorded below limit", fit.belowLimit],
      ["Lots used", fit.lotCount],
      ["Batches used", fit.batchCount],
      ["Bin width", fit.binWidth],
      ["Histogram start", fit.start],
      ["Histogram end", fit.end],
      ["Mean", fit.mean],
      ["Sigma", fit.sigma],
      ["2.5% low cutoff", fit.low25],
      ["2.5% high cutoff", fit.high25],
      ["15% low cutoff", fit.low15],
      ["15% high cutoff", fit.high15],
      ["SSE", fit.sse],
      [],
      ["Bin", "Observed", "Gaussian"],
      ...fit.bins.map((bin) => [bin.center, bin.observed, bin.gaussian])
    ]);
  }
  if (state.gaussianSnapshots.length) {
    appendSheet(workbook, "Gaussian_Snapshots", [
      ["Saved", "File", "Parameter", "Zones", "Data scope", "N", "Mean", "Sigma"],
      ...state.gaussianSnapshots.map((item) => [
        item.savedAt, item.fileName, item.parameter, item.zones, item.scope, item.n, item.mean, item.sigma
      ])
    ]);
  }
  if (state.lastTrend) {
    const { scope, result: trend } = state.lastTrend;
    appendSheet(workbook, "ParameterTrend_App", [
      ["Parameter", trend.parameter],
      ["Data scope", scope],
      ["Rolling batches", trend.rollingWindow],
      ["Visible batches", trend.batches.length],
      ["Visible lots", trend.lots.length],
      ["Auswertung date keys", trend.lookupCount],
      ["Lot-code fallback", trend.fallbackCount],
      ["Lot center", trend.lotCenter],
      ["Lot sigma", trend.lotSigma],
      ["Batch center", trend.batchCenter],
      ["Batch sigma", trend.batchSigma],
      [],
      ["Lot date", "Lot", "Lot mean", "Lot SD", "N values", "Center", "+2 SD", "-2 SD"],
      ...trend.lots.map((item) => [
        new Date(item.date), item.lot, item.mean, item.sigma, item.n,
        trend.lotCenter, trend.lotCenter + 2 * trend.lotSigma, trend.lotCenter - 2 * trend.lotSigma
      ]),
      [],
      ["DateTime", "Lot", "Batch N", "Batch mean", "Batch SD", "Rolling mean", "Center", "+2 SD", "-2 SD", ...ZONES.map((zone) => `Zone ${zone}`)],
      ...trend.batches.map((item) => [
        new Date(item.date), item.lot, item.batch, item.mean, item.sigma, item.rollingMean,
        trend.batchCenter, trend.batchCenter + 2 * trend.batchSigma, trend.batchCenter - 2 * trend.batchSigma,
        ...item.zones
      ]),
      ...(trend.regional ? [
        [],
        ["Zone", "N Batches", "Mean value", "Mean bias", "Bias SD", "95% CI low", "95% CI high", "Signal"],
        ...trend.zoneBias.map((item) => [
          `Zone ${item.zone}`, item.n, item.meanValue, item.meanBias, item.biasSigma, item.ciLow, item.ciHigh, item.signal
        ])
      ] : [])
    ]);
  }
  if (state.lastCorrelation) {
    const { xParameter, yParameter, result } = state.lastCorrelation;
    appendSheet(workbook, "Correlation_App", [
      ["X parameter", xParameter],
      ["Y parameter", yParameter],
      ["Data scope", state.lastCorrelation.scope],
      ["Included pairs", result.rawN],
      ["Pearson r", result.r],
      ["Slope", result.slope],
      ["Intercept", result.intercept],
      [],
      ["Lot", "Batch N", "Zone", "X value", "Y value", "Status"],
      ...result.pairs.map((pair) => [pair.lot, pair.batch, `Zone ${pair.zone}`, pair.x, pair.y, pair.included ? "Included" : "Excluded ratio extreme"])
    ]);
  }
  if (state.lastAssessment) {
    const { parameter: assessmentParameter, lot, mode, assessment } = state.lastAssessment;
    appendSheet(workbook, "NewLot_Assessment_App", [
      ["Parameter", assessmentParameter],
      ["Lot", lot],
      ["Reference", mode],
      ["Data scope", state.lastAssessment.scope],
      ["Overall", assessment.overall],
      ["Out of range", assessment.outOfRangeCount],
      ["Check", assessment.monitorCount],
      ["No history", assessment.noHistoryCount],
      [],
      ["Batch N", ...assessment.availableZones.map((zone) => `${assessmentParameter}_${zone}`)],
      ["Mu", ...assessment.availableZones.map((zone) => assessment.references[zone - 1].mean)],
      ["Sigma", ...assessment.availableZones.map((zone) => assessment.references[zone - 1].sigma)],
      ...assessment.grid.map((row) => [row.batch, ...assessment.availableZones.map((zone) => row.values[zone - 1])])
    ]);
  }
  XLSX.writeFile(workbook, `${baseFileName()}_analysis.xlsx`);
}

function appendSheet(workbook, name, rows) {
  const XLSX = getXlsx();
  const sheet = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(workbook, sheet, safeSheetName(name));
}

function dataRows() {
  return state.rows.filter((row) => isDataRow(state.headers, row));
}

function filteredRows() {
  const filterMode = byId("filter-mode").value || "all";
  const typeValue = byId("type-filter").value || ALL;
  const lotColumn = headerIndex(state.headers, "Lot");
  const typeColumn = headerIndex(state.headers, "Type");
  const classificationColumn = headerIndex(state.headers, "Classification");
  return dataRows().filter((row) => {
    if (typeValue !== ALL && typeColumn >= 0 && text(row[typeColumn]) !== typeValue) return false;
    if (filterMode === "lots" && lotColumn >= 0 && !state.filterSelections.lots.has(text(row[lotColumn]))) return false;
    if (filterMode === "classification") {
      const value = classificationColumn >= 0 && text(row[classificationColumn])
        ? text(row[classificationColumn])
        : UNCLASSIFIED;
      if (!state.filterSelections.classification.has(value)) return false;
    }
    return true;
  });
}

function rowsForAnalysis(scopeId) {
  return byId(scopeId).value === "all" ? dataRows() : filteredRows();
}

function syncAssessmentLots() {
  if (!state.headers.length) return;
  const lotColumn = headerIndex(state.headers, "Lot");
  const lots = lotColumn >= 0 ? getLotValues(rowsForAnalysis("assessment-data-scope"), lotColumn) : [];
  const previous = byId("assessment-lot").value;
  fillSelect(byId("assessment-lot"), lots, lots.includes(previous) ? previous : lots[0]);
  byId("assessment-lot").disabled = !lots.length;
  byId("run-assessment").disabled = !lots.length;
}

function isDataRow(headers, row) {
  if (isEmptyRow(row)) return false;
  const lotColumn = headerIndex(headers, "Lot");
  const nColumn = headerIndex(headers, "N");
  const typeColumn = headerIndex(headers, "Type");
  if (lotColumn >= 0 && !text(row[lotColumn])) return false;
  if (nColumn >= 0 && /^(AVG|SD)\b/i.test(text(row[nColumn]))) return false;
  if (typeColumn >= 0 && /summary|probe/i.test(text(row[typeColumn]))) return false;
  return true;
}

function getTypeValues() {
  const typeColumn = headerIndex(state.headers, "Type");
  if (typeColumn < 0) return [];
  const values = new Set();
  for (const row of dataRows()) {
    const value = text(row[typeColumn]);
    if (value) values.add(value);
  }
  return [...values].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
}

function chooseCleanSource(sheets) {
  if (sheets.includes("Clean_Data")) return "Clean_Data";
  if (sheets.includes("Clean_Data_Cor")) return "Clean_Data_Cor";
  throw new Error("This workbook does not contain a Clean_Data sheet.");
}

function renderZoneChoices(containerId, name) {
  const container = byId(containerId);
  container.replaceChildren();
  for (const zone of ZONES) {
    const label = document.createElement("label");
    const input = document.createElement("input");
    input.type = "checkbox";
    input.name = name;
    input.value = String(zone);
    input.checked = true;
    label.append(input, document.createTextNode(`Zone ${zone}`));
    container.append(label);
  }
  container.addEventListener("change", () => {
    if (containerId === "export-zones") renderExportNote();
    if (containerId === "gaussian-zones") invalidateGaussian();
  });
}

function selectedZones(containerId) {
  return [...byId(containerId).querySelectorAll("input:checked")].map((input) => Number(input.value));
}

function fillSelect(select, values, preferredValue, labelFor = (value) => value) {
  const previous = preferredValue || select.value;
  select.replaceChildren();
  values.forEach((value) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = labelFor(value);
    option.selected = value === previous;
    select.append(option);
  });
  if (!select.value && select.options.length) select.selectedIndex = 0;
}

function syncGaussianMethod() {
  const parameter = byId("gaussian-parameter").value.toLowerCase();
  const methodSelect = byId("gaussian-method");
  const truncatedOption = methodSelect.querySelector('option[value="nacl-truncated"]');
  truncatedOption.disabled = parameter !== "nacl";
  if (truncatedOption.disabled && methodSelect.value === "nacl-truncated") methodSelect.value = "standard";
}

function syncZoneChoices(containerId, parameter) {
  const columns = zoneColumns(state.headers, parameter) || [];
  byId(containerId).querySelectorAll("input").forEach((input) => {
    const available = (columns[Number(input.value) - 1] ?? -1) >= 0;
    input.disabled = !available;
    input.checked = available;
  });
}

function syncCorrelationZones() {
  const xColumns = zoneColumns(state.headers, byId("correlation-x").value) || [];
  const yColumns = zoneColumns(state.headers, byId("correlation-y").value) || [];
  [...byId("correlation-scope").options].forEach((option) => {
    if (option.value === "all") return;
    const zone = Number(option.value) - 1;
    option.disabled = !((xColumns[zone] ?? -1) >= 0 && (yColumns[zone] ?? -1) >= 0);
  });
  if (byId("correlation-scope").selectedOptions[0]?.disabled) byId("correlation-scope").value = "all";
}

function renderTable(rows, options = {}) {
  if (!rows.length || rows.length === 1 && rows[0].length === 0) {
    return `<p class="empty-state">${options.empty || "No rows."}</p>`;
  }
  const [headers, ...body] = rows;
  return `
    <table>
      <thead><tr>${headers.map((cell) => `<th>${escapeHtml(cell)}</th>`).join("")}</tr></thead>
      <tbody>${body.map((row) => `<tr>${headers.map((_, index) => `<td>${formatCell(row[index])}</td>`).join("")}</tr>`).join("")}</tbody>
    </table>
  `;
}

function renderAssessmentTable(rows, gridRows, availableZones) {
  const [headers, ...body] = rows;
  return `
    <table>
      <thead><tr>${headers.map((cell) => `<th>${escapeHtml(cell)}</th>`).join("")}</tr></thead>
      <tbody>${body.map((row, rowIndex) => {
        const statusRow = rowIndex >= 2 ? gridRows[rowIndex - 2]?.states : null;
        return `<tr>${headers.map((_, columnIndex) => {
          const zone = availableZones[columnIndex - 1];
          const status = statusRow?.[zone - 1];
          const className = status ? ` class="${statusClass(status)}"` : "";
          return `<td${className}>${formatCell(row[columnIndex])}</td>`;
        }).join("")}</tr>`;
      }).join("")}</tbody>
    </table>
  `;
}

function drawGaussian(canvas, fit) {
  if (!canvas || !fit?.bins?.length) return;
  const { ctx, width, height, colors } = setupCanvas(canvas);
  const pad = { left: 42, right: 12, top: 16, bottom: 34 };
  const plotWidth = width - pad.left - pad.right;
  const plotHeight = height - pad.top - pad.bottom;
  const maxY = Math.max(1, ...fit.bins.flatMap((bin) => [bin.observed, bin.gaussian]));
  ctx.clearRect(0, 0, width, height);
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, width, height);
  drawFrame(ctx, pad, width, height, colors);
  fit.bins.forEach((bin) => {
    const x0 = pad.left + (bin.lower - fit.start) / (fit.end - fit.start) * plotWidth;
    const x1 = pad.left + (bin.upper - fit.start) / (fit.end - fit.start) * plotWidth;
    const barHeight = bin.observed / maxY * plotHeight;
    ctx.fillStyle = colors.blueSoft;
    ctx.fillRect(x0 + 1, pad.top + plotHeight - barHeight, Math.max(1, x1 - x0 - 2), barHeight);
  });
  ctx.beginPath();
  fit.bins.forEach((bin, index) => {
    const x = pad.left + (bin.center - fit.start) / (fit.end - fit.start) * plotWidth;
    const y = pad.top + plotHeight - bin.gaussian / maxY * plotHeight;
    if (index === 0) ctx.moveTo(x, y);
    else ctx.lineTo(x, y);
  });
  ctx.strokeStyle = colors.green;
  ctx.lineWidth = 2;
  ctx.stroke();
  drawAxisLabels(ctx, pad, width, height, colors, fit.start, fit.end, maxY);
}

function drawScatter(canvas, current) {
  if (!canvas || !current?.result?.pairs?.length) return;
  const { ctx, width, height, colors } = setupCanvas(canvas);
  const pad = { left: 44, right: 16, top: 16, bottom: 34 };
  const plotWidth = width - pad.left - pad.right;
  const plotHeight = height - pad.top - pad.bottom;
  const valuesX = current.result.pairs.map((pair) => pair.x);
  const valuesY = current.result.pairs.map((pair) => pair.y);
  const extentX = paddedExtent(valuesX);
  const extentY = paddedExtent(valuesY);
  const xScale = (value) => pad.left + (value - extentX.min) / (extentX.max - extentX.min) * plotWidth;
  const yScale = (value) => pad.top + plotHeight - (value - extentY.min) / (extentY.max - extentY.min) * plotHeight;
  ctx.clearRect(0, 0, width, height);
  drawFrame(ctx, pad, width, height, colors);
  for (const pair of current.result.pairs) {
    ctx.beginPath();
    ctx.arc(xScale(pair.x), yScale(pair.y), pair.included ? 3 : 2.5, 0, Math.PI * 2);
    ctx.fillStyle = pair.included ? colors.blue : colors.muted;
    ctx.fill();
  }
  if (Number.isFinite(current.result.slope) && Number.isFinite(current.result.intercept)) {
    const y1 = current.result.slope * extentX.min + current.result.intercept;
    const y2 = current.result.slope * extentX.max + current.result.intercept;
    ctx.beginPath();
    ctx.moveTo(xScale(extentX.min), yScale(y1));
    ctx.lineTo(xScale(extentX.max), yScale(y2));
    ctx.strokeStyle = colors.green;
    ctx.lineWidth = 2;
    ctx.stroke();
  }
  drawAxisLabels(ctx, pad, width, height, colors, extentX.min, extentX.max, extentY.max);
}

function drawTrendCharts() {
  const trend = state.lastTrend?.result;
  if (!trend) return;
  const lotCenter = trend.lots.map((item) => ({ x: item.date, y: trend.lotCenter }));
  const lotUpper = trend.lots.map((item) => ({ x: item.date, y: trend.lotCenter + 2 * trend.lotSigma }));
  const lotLower = trend.lots.map((item) => ({ x: item.date, y: trend.lotCenter - 2 * trend.lotSigma }));
  drawTrendLineChart(byId("trend-lot-chart"), [
    { color: TREND_COLORS[0], width: 2.5, points: true, data: trend.lots.map((item) => ({ x: item.date, y: item.mean })) },
    { color: TREND_COLORS[1], width: 1.5, data: lotCenter },
    { color: TREND_COLORS[3], width: 1.2, dash: [5, 4], data: lotUpper },
    { color: TREND_COLORS[3], width: 1.2, dash: [5, 4], data: lotLower }
  ]);

  const batchPoints = (valueFor) => trend.batches.map((item) => ({
    x: item.date,
    y: valueFor(item),
    group: trendGroup(item.lot)
  }));
  drawTrendLineChart(byId("trend-batch-chart"), [
    { color: TREND_COLORS[0], width: 1.5, points: true, breakGroups: true, data: batchPoints((item) => item.mean) },
    { color: TREND_COLORS[2], width: 2.5, breakGroups: true, data: batchPoints((item) => item.rollingMean) },
    { color: TREND_COLORS[1], width: 1.2, breakGroups: true, data: batchPoints(() => trend.batchCenter) },
    { color: TREND_COLORS[3], width: 1.1, dash: [5, 4], breakGroups: true, data: batchPoints(() => trend.batchCenter + 2 * trend.batchSigma) },
    { color: TREND_COLORS[3], width: 1.1, dash: [5, 4], breakGroups: true, data: batchPoints(() => trend.batchCenter - 2 * trend.batchSigma) }
  ]);

  if (trend.regional) {
    drawTrendLineChart(byId("trend-zone-chart"), ZONES.map((zone, index) => ({
      color: TREND_COLORS[index],
      width: 1.6,
      points: true,
      breakGroups: true,
      data: trend.batches.map((item) => ({
        x: item.date,
        y: item.zones[index],
        group: trendGroup(item.lot)
      }))
    })));
    drawTrendBiasChart(byId("trend-bias-chart"), trend.zoneBias);
  }
}

function drawTrendLineChart(canvas, series) {
  if (!canvas) return;
  const finitePoints = series.flatMap((item) => item.data).filter((point) => Number.isFinite(point.x) && Number.isFinite(point.y));
  if (!finitePoints.length) return;
  const { ctx, width, height, colors } = setupCanvas(canvas);
  const pad = { left: 48, right: 14, top: 14, bottom: 38 };
  const plotWidth = width - pad.left - pad.right;
  const plotHeight = height - pad.top - pad.bottom;
  const xExtent = trendExtent(finitePoints.map((point) => point.x), 86400000);
  const yExtent = paddedExtent(finitePoints.map((point) => point.y));
  const xScale = (value) => pad.left + (value - xExtent.min) / (xExtent.max - xExtent.min) * plotWidth;
  const yScale = (value) => pad.top + plotHeight - (value - yExtent.min) / (yExtent.max - yExtent.min) * plotHeight;
  ctx.clearRect(0, 0, width, height);
  drawTrendGrid(ctx, pad, width, height, colors);
  for (const item of series) {
    ctx.beginPath();
    let previous = null;
    for (const point of item.data) {
      if (!Number.isFinite(point.x) || !Number.isFinite(point.y)) {
        previous = null;
        continue;
      }
      const x = xScale(point.x);
      const y = yScale(point.y);
      if (!previous || item.breakGroups && point.group !== previous.group) ctx.moveTo(x, y);
      else ctx.lineTo(x, y);
      previous = point;
    }
    ctx.strokeStyle = item.color;
    ctx.lineWidth = item.width || 1.5;
    ctx.setLineDash(item.dash || []);
    ctx.stroke();
    ctx.setLineDash([]);
    if (item.points) {
      ctx.fillStyle = item.color;
      for (const point of item.data) {
        if (!Number.isFinite(point.x) || !Number.isFinite(point.y)) continue;
        ctx.beginPath();
        ctx.arc(xScale(point.x), yScale(point.y), 2.2, 0, Math.PI * 2);
        ctx.fill();
      }
    }
  }
  drawTrendAxisLabels(ctx, pad, width, height, colors, xExtent, yExtent);
}

function drawTrendBiasChart(canvas, values) {
  if (!canvas || !values.length) return;
  const { ctx, width, height, colors } = setupCanvas(canvas);
  const pad = { left: 48, right: 12, top: 14, bottom: 38 };
  const plotWidth = width - pad.left - pad.right;
  const plotHeight = height - pad.top - pad.bottom;
  const absoluteMax = Math.max(1e-9, ...values.flatMap((item) => [Math.abs(item.meanBias || 0), Math.abs(item.ciLow || 0), Math.abs(item.ciHigh || 0)]));
  const minY = -absoluteMax * 1.2;
  const maxY = absoluteMax * 1.2;
  const yScale = (value) => pad.top + plotHeight - (value - minY) / (maxY - minY) * plotHeight;
  ctx.clearRect(0, 0, width, height);
  drawTrendGrid(ctx, pad, width, height, colors);
  const slot = plotWidth / values.length;
  values.forEach((item, index) => {
    const x = pad.left + index * slot + slot * 0.22;
    const barWidth = slot * 0.56;
    const zeroY = yScale(0);
    const valueY = yScale(item.meanBias || 0);
    ctx.fillStyle = item.signal === "No clear bias" ? TREND_COLORS[1]
      : item.signal === "Insufficient" ? colors.muted
        : TREND_COLORS[3];
    ctx.fillRect(x, Math.min(zeroY, valueY), barWidth, Math.max(1, Math.abs(zeroY - valueY)));
    if (item.n > 1) {
      const centerX = x + barWidth / 2;
      ctx.beginPath();
      ctx.moveTo(centerX, yScale(item.ciLow));
      ctx.lineTo(centerX, yScale(item.ciHigh));
      ctx.moveTo(centerX - 4, yScale(item.ciLow));
      ctx.lineTo(centerX + 4, yScale(item.ciLow));
      ctx.moveTo(centerX - 4, yScale(item.ciHigh));
      ctx.lineTo(centerX + 4, yScale(item.ciHigh));
      ctx.strokeStyle = colors.ink;
      ctx.lineWidth = 1;
      ctx.stroke();
    }
    ctx.fillStyle = colors.muted;
    ctx.font = "11px Aptos, Calibri, Arial, sans-serif";
    ctx.textAlign = "center";
    ctx.fillText(`Z${item.zone}`, x + barWidth / 2, height - 12);
  });
  ctx.beginPath();
  ctx.moveTo(pad.left, yScale(0));
  ctx.lineTo(width - pad.right, yScale(0));
  ctx.strokeStyle = colors.ink;
  ctx.lineWidth = 1;
  ctx.stroke();
  ctx.fillStyle = colors.muted;
  ctx.font = "11px Aptos, Calibri, Arial, sans-serif";
  ctx.textAlign = "left";
  ctx.fillText(formatNumber(maxY, 3), 4, pad.top + 8);
  ctx.fillText(formatNumber(minY, 3), 4, height - pad.bottom);
}

function drawTrendGrid(ctx, pad, width, height, colors) {
  ctx.strokeStyle = colors.line;
  ctx.lineWidth = 1;
  for (let index = 0; index <= 4; index += 1) {
    const y = pad.top + (height - pad.top - pad.bottom) * index / 4;
    ctx.beginPath();
    ctx.moveTo(pad.left, y);
    ctx.lineTo(width - pad.right, y);
    ctx.stroke();
  }
  drawFrame(ctx, pad, width, height, colors);
}

function drawTrendAxisLabels(ctx, pad, width, height, colors, xExtent, yExtent) {
  ctx.fillStyle = colors.muted;
  ctx.font = "11px Aptos, Calibri, Arial, sans-serif";
  ctx.textAlign = "left";
  ctx.fillText(formatTrendDate(xExtent.min), pad.left, height - 12);
  ctx.textAlign = "right";
  ctx.fillText(formatTrendDate(xExtent.max), width - pad.right, height - 12);
  ctx.textAlign = "left";
  ctx.fillText(formatNumber(yExtent.max, 3), 4, pad.top + 8);
  ctx.fillText(formatNumber(yExtent.min, 3), 4, height - pad.bottom);
}

function trendExtent(values, minimumRange) {
  let min = Math.min(...values);
  let max = Math.max(...values);
  if (min === max) {
    min -= minimumRange / 2;
    max += minimumRange / 2;
  }
  return { min, max };
}

function trendGroup(value) {
  const rendered = text(value);
  return Number.isFinite(Number(rendered)) ? String(Number(rendered)) : rendered.toUpperCase();
}

function setupCanvas(canvas) {
  const ratio = window.devicePixelRatio || 1;
  const rect = canvas.getBoundingClientRect();
  const width = Math.max(300, rect.width || 300);
  const height = Math.max(220, rect.height || 240);
  canvas.width = Math.floor(width * ratio);
  canvas.height = Math.floor(height * ratio);
  const ctx = canvas.getContext("2d");
  ctx.setTransform(ratio, 0, 0, ratio, 0, 0);
  const styles = getComputedStyle(document.documentElement);
  return {
    ctx,
    width,
    height,
    colors: {
      ink: styles.getPropertyValue("--ink").trim(),
      muted: styles.getPropertyValue("--muted").trim(),
      line: styles.getPropertyValue("--line").trim(),
      blue: styles.getPropertyValue("--blue").trim(),
      blueSoft: "rgba(25, 93, 141, 0.28)",
      green: styles.getPropertyValue("--green").trim()
    }
  };
}

function drawFrame(ctx, pad, width, height, colors) {
  ctx.strokeStyle = colors.line;
  ctx.lineWidth = 1;
  ctx.beginPath();
  ctx.moveTo(pad.left, pad.top);
  ctx.lineTo(pad.left, height - pad.bottom);
  ctx.lineTo(width - pad.right, height - pad.bottom);
  ctx.stroke();
}

function drawAxisLabels(ctx, pad, width, height, colors, minX, maxX, maxY) {
  ctx.fillStyle = colors.muted;
  ctx.font = "12px Aptos, Calibri, Arial, sans-serif";
  ctx.textAlign = "left";
  ctx.fillText(formatNumber(minX, 2), pad.left, height - 10);
  ctx.textAlign = "right";
  ctx.fillText(formatNumber(maxX, 2), width - pad.right, height - 10);
  ctx.textAlign = "left";
  ctx.fillText(formatNumber(maxY, 2), 6, pad.top + 8);
}

function paddedExtent(values) {
  const finite = values.filter(Number.isFinite);
  let min = Math.min(...finite);
  let max = Math.max(...finite);
  if (min === max) {
    min -= 1;
    max += 1;
  }
  const pad = (max - min) * 0.06;
  return { min: min - pad, max: max + pad };
}

function redrawCharts() {
  if (state.lastGaussian && byId("gaussian-chart")) drawGaussian(byId("gaussian-chart"), state.lastGaussian.fit);
  if (state.lastTrend && byId("trend-lot-chart")) drawTrendCharts();
  if (state.lastCorrelation && byId("correlation-chart")) drawScatter(byId("correlation-chart"), state.lastCorrelation);
}

function drawEmptyState() {
  byId("data-metrics").innerHTML = [
    metric("Workbook", "-"),
    metric("Rows", "-"),
    metric("Lots", "-"),
    metric("Parameters", "-")
  ].join("");
  byId("summary-table").innerHTML = `<p class="empty-state">No workbook loaded.</p>`;
  byId("preview-table").innerHTML = `<p class="empty-state">No workbook loaded.</p>`;
  byId("build-result").innerHTML = `<p class="empty-state">No generated data.</p>`;
  byId("generated-summary-table").innerHTML = `<p class="empty-state">No generated summary.</p>`;
  byId("classification-table").innerHTML = `<p class="empty-state">No lots.</p>`;
}

function openPanel(panelId) {
  document.querySelectorAll(".tab").forEach((tab) => tab.classList.toggle("is-active", tab.dataset.panel === panelId));
  document.querySelectorAll(".panel").forEach((panel) => panel.classList.toggle("is-active", panel.id === panelId));
  requestAnimationFrame(redrawCharts);
}

function getXlsx() {
  if (!window.XLSX) throw new Error("The workbook reader did not load. Check the internet connection and reload this page.");
  return window.XLSX;
}

function normalizeRows(rows, width) {
  return rows.map((row) => normalizeRow(row, width)).filter((row) => !isEmptyRow(row));
}

function normalizeRow(row, width = row.length) {
  return Array.from({ length: width }, (_, index) => normalizeCell(row[index]));
}

function normalizeCell(value) {
  if (value instanceof Date) return value;
  if (typeof value === "string") return value.trim();
  return value ?? "";
}

function isEmptyRow(row) {
  return row.every((cell) => text(cell) === "");
}

function metric(label, value) {
  return `<div class="metric"><span>${escapeHtml(label)}</span><strong>${escapeHtml(value)}</strong></div>`;
}

function metricHtml(label, value) {
  return `<div class="metric"><span>${escapeHtml(label)}</span><strong>${value}</strong></div>`;
}

function statusBadge(status) {
  return `<span class="badge ${statusClass(status)}">${escapeHtml(status)}</span>`;
}

function statusClass(status) {
  if (status === "OK") return "status-ok";
  if (status === "CHECK") return "status-check";
  if (status === "OUT OF RANGE") return "status-bad";
  return "status-missing";
}

function formatCell(value) {
  if (value instanceof Date) return escapeHtml(value.toISOString().slice(0, 10));
  if (typeof value === "number") return escapeHtml(formatNumber(value, 4));
  return escapeHtml(value ?? "");
}

function formatNumber(value, digits = 2) {
  if (value === null || value === undefined || value === "" || !Number.isFinite(Number(value))) return "";
  return Number(value).toLocaleString(undefined, { maximumFractionDigits: digits });
}

function formatInteger(value) {
  if (value === null || value === undefined || value === "" || !Number.isFinite(Number(value))) return "";
  return Number(value).toLocaleString(undefined, { maximumFractionDigits: 0 });
}

function inputNumber(value) {
  return Number.isFinite(Number(value)) ? String(Number(Number(value).toPrecision(12))) : "";
}

function formatTrendDate(value) {
  const date = new Date(value);
  if (!Number.isFinite(date.getTime())) return "";
  return date.toISOString().slice(0, 10);
}

function formatDateTime(value) {
  const date = new Date(value);
  if (!Number.isFinite(date.getTime())) return "";
  return date.toLocaleString(undefined, { dateStyle: "short", timeStyle: "short" });
}

function fileDateStamp(date) {
  const parts = [
    date.getFullYear(),
    String(date.getMonth() + 1).padStart(2, "0"),
    String(date.getDate()).padStart(2, "0"),
    String(date.getHours()).padStart(2, "0"),
    String(date.getMinutes()).padStart(2, "0"),
    String(date.getSeconds()).padStart(2, "0")
  ];
  return `${parts.slice(0, 3).join("")}_${parts.slice(3).join("")}`;
}

function optionalNumber(id) {
  const value = byId(id).value.trim();
  return value === "" ? undefined : Number(value);
}

function requiredNumber(id) {
  const value = Number(byId(id).value);
  if (!Number.isFinite(value)) throw new Error("Enter a numeric value for every required setting.");
  return value;
}

function toCsv(rows) {
  return rows.map((row) => row.map(csvCell).join(",")).join("\r\n");
}

function csvCell(value) {
  const rendered = value instanceof Date ? value.toISOString() : String(value ?? "");
  return /[",\r\n]/.test(rendered) ? `"${rendered.replaceAll('"', '""')}"` : rendered;
}

function downloadText(fileName, content, type) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.append(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

function safeSheetName(name) {
  return name.replace(/[\\/?*[\]:]/g, "_").slice(0, 31) || "Sheet";
}

function baseFileName() {
  return (state.workbookName || "ipw")
    .replace(/\.[^.]+$/, "")
    .replace(/[^a-z0-9_-]+/gi, "_")
    .replace(/^_+|_+$/g, "") || "ipw";
}

function safeFilePart(value) {
  return text(value).replace(/[^a-z0-9_-]+/gi, "_").replace(/^_+|_+$/g, "") || "parameter";
}

function enableControls(ids) {
  ids.forEach((id) => {
    byId(id).disabled = false;
  });
}

function yieldToBrowser() {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

function clearResult(id) {
  byId(id).innerHTML = "";
}

function byId(id) {
  return document.getElementById(id);
}

function setStatus(message, isError = false, isSuccess = false) {
  const status = byId("status");
  status.textContent = message;
  status.classList.toggle("is-error", isError);
  status.classList.toggle("is-success", isSuccess);
}

async function runAction(action) {
  const buttons = [...document.querySelectorAll("button")];
  buttons.forEach((button) => {
    if (!button.disabled) button.dataset.wasEnabled = "true";
    button.disabled = true;
  });
  try {
    await action();
  } catch (error) {
    console.error(error);
    setStatus(error.message || "The analysis could not be completed.", true);
  } finally {
    buttons.forEach((button) => {
      if (button.dataset.wasEnabled) {
        button.disabled = false;
        delete button.dataset.wasEnabled;
      }
    });
    byId("save-gaussian-snapshot").disabled = !state.lastGaussian;
  }
}

function debounce(callback, wait) {
  let timer = 0;
  return () => {
    window.clearTimeout(timer);
    timer = window.setTimeout(callback, wait);
  };
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
