import { readPreference, writePreference, clearPreferences, setRememberSettings, rememberSettingsEnabled, datasetPreferenceKey } from "./local-preferences.js?v=1";
import {
  buildCorrelation,
  buildLotAssessment,
  buildOriginExport,
  buildOriginExportWide,
  buildParameterTrend,
  buildTrendDateLookup,
  findZonedHeaderRow,
  filterLotsByPeriod,
  gaussianFitWithOptions,
  gaussianExtremeSnapshot,
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
} from "./analysis.js?v=16";
import {
  buildCleanDataFromAuswertung,
  buildFullSummary,
  classificationMaps,
  correctCleanData
} from "./clean-data.js?v=2";
import {
  buildDataLabel,
  buildDataLabelsTable,
  buildDataLabelValuesTable,
  dataLabelCellCoordinates,
  findDataLabel,
  findRawHeaderRow,
  getMasterRollWidths,
  getMasterRollZone,
  mergeAuswertungTables,
  parseDataLabels,
  removeDataLabel,
  upsertDataLabel
} from "./data-management.js?v=4";
import {
  assessmentBatchIndexes,
  buildLotReleaseSummary,
  buildPeriodComparison,
  buildV90LotAssessment,
  getV90Parameters,
  integerChartAxis,
  getZmPlanSpecification
} from "./v90-analysis.js?v=6";

import { classificationIncludesKeyword, updateWorkbookClassifications } from "./lot-classification.js?v=2";
import { buildStructuredCorrelation, structuredParameters, trimStructuredPlot, structuredConclusions } from "./structured-correlation.js?v=5";

const state = {
  workbook: null,
  originalData: null,
  workbookName: "",
  sheets: [],
  generatedSources: new Map(),
  newLotImport: null,
  dataLabels: [],
  source: "",
  headers: [],
  rows: [],
  sourceLocations: new Map(),
  sourceColumns: [],
  parameters: [],
  structuredParameters: [],
  trendParameters: [],
  v90Parameters: [],
  lots: [],
  types: [],
  lotClassifications: new Map(),
  classificationEdits: new Map(),
  classificationsUnsaved: false,
  filterSelections: { lots: new Set(), classification: new Set() },
  filterInitialized: { lots: false, classification: false },
  generatedClean: false,
  lastBuild: null,
  lastGaussian: null,
  gaussianSnapshots: [],
  gaussianExtremeExports: [],
  trendDateLookup: null,
  periodFilterLoading: false,
  periodFilterError: "",
  lastTrend: null,
  lastPeriod: null,
  lastCorrelation: null,
  lastStructuredCorrelation: null,
  lastAssessment: null,
  assessmentBatchQuery: "",
  equalReferenceLots: new Set(),
  lastRelease: null,
  lastZmPlan: null
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
  bindRememberedSettings();
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
  let hadController = Boolean(navigator.serviceWorker.controller);
  navigator.serviceWorker.addEventListener("controllerchange", () => {
    document.documentElement.dataset.offlineReady = "true";
    if (hadController) byId("app-update").hidden = false;
    hadController = true;
  });
  byId("app-update").addEventListener("click", () => {
    if (state.workbook && !window.confirm("Updating reloads the app. Save any workbook changes first. Update now?")) return;
    window.location.reload();
  });
  navigator.serviceWorker.register("./service-worker.js", { scope: "./", updateViaCache: "none" })
    .then(async (registration) => {
      if (registration.active) document.documentElement.dataset.offlineReady = "true";
      await registration.update();
      registration.waiting?.postMessage({ type: "SKIP_WAITING" });
    })
    .catch((error) => {
      console.warn("Offline cache registration failed.", error);
    });
}

const REMEMBERED_CONTROLS = [
  "correlation-analysis", "correlation-method", "correlation-compare", "correlation-min-n", "correlation-coverage", "correlation-min-regions", "correlation-expected-regions",
  "reference-temperature", "reference-humidity",
  "summary-parameter", "gaussian-parameter", "gaussian-method", "gaussian-bin-width", "gaussian-start", "gaussian-end", "gaussian-extreme-sigma", "gaussian-extreme-side",
  "trend-parameter", "period-parameter", "period-mode", "period-plot", "period-lot", "period-lot-b",
  "period-a-start", "period-a-end", "period-b-start", "period-b-end", "period-c-start", "period-c-end",
  "assessment-lot", "assessment-parameter", "assessment-reference", "assessment-reference-lot",
  "assessment-granularity", "assessment-monitor", "assessment-outlier", "assessment-mu", "assessment-sigma",
  "release-lot", "release-reference", "release-monitor", "release-not-ok", "zm-layout",
  "correlation-x", "correlation-y", "correlation-scope", "correlation-outliers", "correlation-removal", "export-parameter"
];

function bindRememberedSettings() {
  byId("remember-settings").checked = rememberSettingsEnabled();
  byId("remember-settings").addEventListener("change", () => {
    setRememberSettings(byId("remember-settings").checked);
    byId("remember-settings").checked = rememberSettingsEnabled();
    if (rememberSettingsEnabled()) saveAnalysisSettings();
  });
  byId("clear-saved-settings").addEventListener("click", () => {
    clearPreferences();
    setStatus("Saved device settings cleared. Current workbook and analysis are unchanged.", false, true);
  });
  const save = (event) => {
    if (REMEMBERED_CONTROLS.includes(event.target.id) || event.target.closest("#assessment-equal-options, #gaussian-zones, #export-zones")) {
      saveAnalysisSettings();
    }
  };
  document.addEventListener("change", save);
  document.addEventListener("input", save);
  ["assessment-equal-all", "assessment-equal-clear"].forEach((id) => byId(id).addEventListener("click", saveAnalysisSettings));
}

function saveAnalysisSettings() {
  if (!state.workbook || !state.headers.length) return;
  const controls = {};
  REMEMBERED_CONTROLS.forEach((id) => {
    const input = byId(id);
    if (input) controls[id] = input.value;
  });
  const zones = {};
  ["gaussian-zones", "export-zones"].forEach((id) => {
    zones[id] = [...byId(id).querySelectorAll("input:checked")].map((input) => input.value);
  });
  writePreference(datasetPreferenceKey(state.workbookName, state.source, state.headers), {
    controls, zones, referenceLots: [...state.equalReferenceLots]
  });
}

function restoreAnalysisSettings() {
  const saved = readPreference(datasetPreferenceKey(state.workbookName, state.source, state.headers));
  if (!saved || typeof saved !== "object" || !saved.controls || typeof saved.controls !== "object") return;
  const applyControls = () => REMEMBERED_CONTROLS.forEach((id) => {
    const input = byId(id);
    const value = saved.controls[id];
    if (!input || typeof value !== "string") return;
    if (input.tagName === "SELECT" && ![...input.options].some((option) => option.value === value)) return;
    input.value = value;
  });
  applyControls();
  syncStructuredCorrelationControls();
  syncAssessmentReferenceLots();
  syncPeriodLots();
  applyControls();
  syncStructuredCorrelationControls(false);
  const available = new Set(availableEqualReferenceLots().map(assessmentLotKey));
  state.equalReferenceLots = new Set((Array.isArray(saved.referenceLots) ? saved.referenceLots : [])
    .filter((lot) => typeof lot === "string" && available.has(lot)));
  renderEqualReferenceChoices();
  syncAssessmentReferenceMode();
  syncPeriodMode();
  syncGaussianMethod();
  syncZoneChoices("gaussian-zones", byId("gaussian-parameter").value);
  syncZoneChoices("export-zones", byId("export-parameter").value);
  syncCorrelationZones();
  syncCorrelationRemovalInput();
  if (typeof saved.controls["correlation-removal"] === "string") byId("correlation-removal").value = saved.controls["correlation-removal"];
  ["gaussian-zones", "export-zones"].forEach((id) => {
    if (!Array.isArray(saved.zones?.[id])) return;
    byId(id).querySelectorAll("input").forEach((input) => { input.checked = saved.zones[id].includes(input.value); });
  });
}

function bindEvents() {
  byId("workbook-file").addEventListener("change", (event) => runAction(() => openWorkbook(event.target.files[0])));
  byId("load-local-workbook").addEventListener("click", () => runAction(loadLocalWorkbook));
  byId("new-lot-file").addEventListener("change", (event) => runAction(() => openNewLotWorkbook(event.target.files[0])));
  byId("merge-new-lots").addEventListener("click", () => runAction(mergeNewLots));
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
  byId("classification-keyword").addEventListener("input", () => {
    invalidateAnalyses();
    renderCurrentData();
  });
  ["limit-lot-period", "lot-period-start", "lot-period-end"].forEach(id => {
    byId(id).addEventListener("change", () => runAction(updateLotPeriodFilter));
  });
  byId("filter-options").addEventListener("change", handleFilterOptionChange);
  byId("select-all-filters").addEventListener("click", () => setAllFilterOptions(true));
  byId("clear-all-filters").addEventListener("click", () => setAllFilterOptions(false));
  byId("clean-output").addEventListener("change", syncCorrectionInputs);
  byId("build-clean-data").addEventListener("click", () => runAction(buildCleanData));
  byId("classification-lot").addEventListener("change", syncClassificationInput);
  byId("classification-search").addEventListener("input", renderClassificationTable);
  byId("save-classifications").addEventListener("click", () => runAction(saveClassifications));
  byId("classification-table").addEventListener("change", (event) => {
    const input = event.target.closest("input[data-lot]");
    if (input) runAction(() => applyLotClassification(input.dataset.lot, input.value));
  });
  byId("apply-classification").addEventListener("click", () => runAction(applyLotClassification));
  byId("classification-value").addEventListener("keydown", (event) => {
    if (event.key === "Enter") runAction(applyLotClassification);
  });
  byId("label-lot").addEventListener("change", () => {
    syncLabelBatches();
    syncLabelEditor();
  });
  byId("label-batch").addEventListener("change", syncLabelEditor);
  byId("label-zone").addEventListener("change", syncLabelEditor);
  byId("label-machine").addEventListener("change", syncMasterRollControls);
  byId("label-roll-width").addEventListener("change", updateMasterRollZone);
  byId("label-roll-number").addEventListener("input", updateMasterRollZone);
  byId("label-parameters").addEventListener("change", renderLabelSelectionPreview);
  byId("save-data-label").addEventListener("click", () => runAction(saveDataLabel));
  byId("remove-data-label").addEventListener("click", () => runAction(removeSelectedDataLabel));
  ["gaussian-data-scope", "trend-data-scope", "correlation-data-scope"].forEach((id) => {
    byId(id).addEventListener("change", invalidateAnalyses);
  });
  ["period-data-scope", "period-parameter", "period-plot"].forEach((id) => byId(id).addEventListener("change", invalidatePeriod));
  ["period-a-start", "period-a-end", "period-b-start", "period-b-end", "period-c-start", "period-c-end"]
    .forEach((id) => byId(id).addEventListener("input", invalidatePeriod));
  byId("period-lot").addEventListener("change", () => {
    syncPeriodLots();
    invalidatePeriod();
  });
  byId("period-lot-b").addEventListener("change", invalidatePeriod);
  byId("period-mode").addEventListener("change", () => {
    syncPeriodMode();
    invalidatePeriod();
  });
  byId("release-lot").addEventListener("change", invalidateRelease);
  byId("release-reference").addEventListener("change", invalidateRelease);
  byId("release-monitor").addEventListener("input", invalidateRelease);
  byId("release-not-ok").addEventListener("input", invalidateRelease);
  byId("assessment-lot").addEventListener("change", () => {
    state.assessmentBatchQuery = "";
    syncAssessmentReferenceLots();
    invalidateAssessment();
  });
  byId("assessment-parameter").addEventListener("change", () => {
    syncAssessmentReferenceMode();
    invalidateAssessment();
  });
  byId("assessment-reference").addEventListener("change", () => {
    syncAssessmentReferenceMode();
    invalidateAssessment();
  });
  byId("assessment-reference-lot").addEventListener("change", invalidateAssessment);
  byId("assessment-result").addEventListener("click", (event) => {
    const button = event.target.closest(".assessment-value-button");
    if (button) showEqualLotDetails(Number(button.dataset.row), Number(button.dataset.column));
  });
  byId("assessment-result").addEventListener("input", (event) => {
    if (event.target.id !== "assessment-batch-filter") return;
    state.assessmentBatchQuery = event.target.value;
    renderAssessmentBatchTables();
  });
  byId("assessment-equal-search").addEventListener("input", renderEqualReferenceChoices);
  byId("assessment-equal-options").addEventListener("change", (event) => {
    const input = event.target.closest("input[type=checkbox]");
    if (!input) return;
    if (input.checked) state.equalReferenceLots.add(input.value);
    else state.equalReferenceLots.delete(input.value);
    updateEqualReferenceCount();
    invalidateAssessment();
  });
  byId("assessment-equal-all").addEventListener("click", () => {
    state.equalReferenceLots = new Set(availableEqualReferenceLots().map(assessmentLotKey));
    renderEqualReferenceChoices();
    invalidateAssessment();
  });
  byId("assessment-equal-clear").addEventListener("click", () => {
    state.equalReferenceLots.clear();
    renderEqualReferenceChoices();
    invalidateAssessment();
  });
  byId("assessment-granularity").addEventListener("change", invalidateAssessment);
  ["assessment-monitor", "assessment-outlier", "assessment-mu", "assessment-sigma"].forEach((id) => {
    byId(id).addEventListener("input", invalidateAssessment);
  });
  byId("summary-parameter").addEventListener("change", renderCurrentData);
  byId("generated-summary-parameter").addEventListener("change", renderGeneratedSummaryTable);
  byId("gaussian-parameter").addEventListener("change", () => {
    invalidateGaussian();
    syncGaussianMethod();
    syncZoneChoices("gaussian-zones", byId("gaussian-parameter").value);
  });
  byId("gaussian-method").addEventListener("change", invalidateGaussian);
  byId("gaussian-visible-rows").addEventListener("change", invalidateGaussian);
  byId("gaussian-full-range").addEventListener("click", () => {
    ["gaussian-start", "gaussian-end", "gaussian-bin-width"].forEach(id => { byId(id).value = ""; });
    invalidateGaussian();
    saveAnalysisSettings();
    setStatus("Gaussian range reset to all numeric values in the selected data scope.", false, true);
  });
  ["gaussian-bin-width", "gaussian-start", "gaussian-end", "gaussian-extreme-sigma"].forEach((id) => {
    byId(id).addEventListener("input", invalidateGaussian);
  });
  ["trend-parameter", "trend-start", "trend-end", "trend-window"].forEach((id) => {
    byId(id).addEventListener(id === "trend-parameter" ? "change" : "input", invalidateTrend);
  });
  byId("correlation-y").addEventListener("change", () => {
    syncCorrelationZones();
    invalidateCorrelation();
  });
  byId("correlation-x").addEventListener("change", () => {
    syncCorrelationZones();
    invalidateCorrelation();
  });
  byId("correlation-scope").addEventListener("change", invalidateCorrelation);
  byId("correlation-outliers").addEventListener("change", () => {
    syncCorrelationRemovalInput();
    invalidateCorrelation();
  });
  byId("correlation-removal").addEventListener("input", invalidateCorrelation);
  byId("correlation-analysis").addEventListener("change", () => {
    syncStructuredCorrelationControls();
    invalidateCorrelation();
  });
  ["correlation-method", "correlation-compare", "correlation-min-n", "correlation-coverage", "correlation-min-regions", "correlation-expected-regions"].forEach(id => {
    byId(id).addEventListener("input", () => {
      syncStructuredCorrelationControls(false);
      invalidateCorrelation();
    });
  });
  byId("export-parameter").addEventListener("change", () => {
    syncZoneChoices("export-zones", byId("export-parameter").value);
    renderExportNote();
  });
  byId("recommend-gaussian").addEventListener("click", () => runAction(recommendGaussian));
  byId("run-gaussian").addEventListener("click", () => runAction(createGaussian));
  byId("save-gaussian-snapshot").addEventListener("click", () => runAction(saveGaussianSnapshot));
  byId("export-gaussian-extremes").addEventListener("click", () => runAction(exportGaussianExtremes));
  byId("gaussian-extreme-side").addEventListener("change", renderGaussianExtremes);
  byId("run-trend").addEventListener("click", () => runAction(createTrend));
  byId("run-period").addEventListener("click", () => runAction(createPeriodComparison));
  byId("save-period-plot").addEventListener("click", () => runAction(savePeriodPlot));
  byId("run-assessment").addEventListener("click", () => runAction(createAssessment));
  byId("show-zm-plan").addEventListener("click", () => runAction(showZmPlan));
  byId("save-zm-plan-png").addEventListener("click", () => runAction(saveFullZmPlanPng));
  byId("apply-zm-label").addEventListener("click", () => runAction(() => updateZmPlanLabel(false)));
  byId("remove-zm-label").addEventListener("click", () => runAction(() => updateZmPlanLabel(true)));
  byId("zm-layout").addEventListener("change", invalidateZmPlan);
  byId("zm-coordinates").addEventListener("input", invalidateZmPlan);
  byId("run-release").addEventListener("click", () => runAction(createReleaseSummary));
  byId("download-release").addEventListener("click", () => runAction(downloadReleaseSummary));
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
  invalidateGaussian();
  state.sourceLocations.clear();
  byId("gaussian-visible-rows").checked = false;
  state.workbookName = fileName;
  state.originalData = data.slice(0);
  state.generatedSources = new Map();
  state.newLotImport = null;
  byId("new-lot-file").value = "";
  state.dataLabels = [];
  state.lotClassifications = new Map();
  state.classificationEdits = new Map();
  state.classificationsUnsaved = false;
  byId("classification-search").value = "";
  byId("classification-keyword").value = "";
  state.filterSelections = { lots: new Set(), classification: new Set() };
  state.filterInitialized = { lots: false, classification: false };
  state.equalReferenceLots.clear();
  state.assessmentBatchQuery = "";
  state.lastBuild = null;
  state.gaussianSnapshots = [];
  state.gaussianExtremeExports = [];
  state.trendDateLookup = null;
  state.periodFilterLoading = false;
  state.periodFilterError = "";
  byId("limit-lot-period").checked = false;
  byId("lot-period-start").value = "";
  byId("lot-period-end").value = "";
  syncLotPeriodControls();
  state.lastTrend = null;
  state.lastPeriod = null;
  state.lastRelease = null;
  state.lastZmPlan = null;
  setStatus("Reading workbook data...");
  await yieldToBrowser();
  state.workbook = XLSX.read(data, {
    type: "array",
    cellDates: true,
    cellFormula: true,
    cellStyles: true,
    cellHTML: false,
    cellText: true,
    dense: false,
    sheets: ["Clean_Data", "Clean_Data_Cor", "Data_Labels"]
  });
  state.sheets = ["Clean_Data", "Clean_Data_Cor"].filter((name) => state.workbook.Sheets[name]);
  state.dataLabels = parseDataLabels(readWorkbookSheet(state.workbook, "Data_Labels"));
  let sourceName;
  let generatedFromAuswertung = null;
  if (state.sheets.length) {
    sourceName = chooseCleanSource(state.sheets);
  } else {
    setStatus("No Clean_Data sheet found. Building it from Auswertung...");
    await yieldToBrowser();
    const rawTable = readAuswertungTable(data);
    const clean = buildCleanDataFromAuswertung(rawTable);
    const summary = buildFullSummary(clean.headers, clean.rows);
    state.generatedSources.set(GENERATED_CLEAN, [clean.headers, ...clean.rows]);
    state.lastBuild = {
      sourceName: GENERATED_CLEAN,
      rows: clean.rows.length,
      lots: getLotValues(clean.rows, headerIndex(clean.headers, "Lot")).length,
      parameters: getRegionalParameters(clean.headers).length,
      summaryRows: summary.rows.length,
      removedProbeRows: clean.removedProbeRows,
      removedVeRows: clean.removedVeRows,
      removedVeLots: clean.removedVeLots,
      corrected: false
    };
    sourceName = GENERATED_CLEAN;
    generatedFromAuswertung = clean;
  }
  refreshSourceSelect(sourceName);
  await selectSource(sourceName);
  if (generatedFromAuswertung) {
    setStatus(
      `No Clean_Data sheet was present. Generated ${formatInteger(generatedFromAuswertung.rows.length)} rows from Auswertung in the browser.`,
      false,
      true
    );
  }
}

async function selectSource(sheetName) {
  if (!state.workbook) return;
  setStatus(`Reading ${sheetName}...`);
  const source = prepareSource(sheetName);
  if (state.source !== sheetName) {
    state.assessmentBatchQuery = "";
    state.equalReferenceLots.clear();
    byId("assessment-equal-search").value = "";
  }
  state.source = sheetName;
  state.headers = source.headers;
  state.rows = source.rows;
  state.sourceLocations = new Map(state.rows.map((row, index) => [row, { row: source.rowNumbers[index], column: source.columnOffset }]));
  state.sourceColumns = state.headers.map((_, index) => source.columnOffset + index);
  state.generatedClean = source.generatedClean;
  if (headerIndex(state.headers, "Classification") < 0) {
    const nIndex = headerIndex(state.headers, "N");
    state.sourceColumns.splice(nIndex >= 0 ? nIndex : 0, 0, null);
  }
  ensureClassificationColumn(state.headers, state.rows);
  seedLotClassifications(state.headers, state.rows);
  applyClassificationOverrides(state.headers, state.rows);
  state.parameters = getRegionalParameters(state.headers);
  state.structuredParameters = structuredParameters(state.headers, state.rows);
  state.trendParameters = getTrendParameters(state.headers);
  state.v90Parameters = getV90Parameters(state.headers);
  if (!state.parameters.length && !state.structuredParameters.length) throw new Error("Clean_Data has no named numeric parameters.");
  if (!state.parameters.length) byId("correlation-analysis").value = "structured";
  const lotColumn = headerIndex(state.headers, "Lot");
  state.lots = lotColumn >= 0 ? getLotValues(dataRows(), lotColumn) : [];
  state.types = getTypeValues();
  populateWorkbookControls();
  restoreAnalysisSettings();
  state.lastGaussian = null;
  state.lastTrend = null;
  state.lastPeriod = null;
  state.lastCorrelation = null;
  state.lastStructuredCorrelation = null;
  state.lastAssessment = null;
  state.lastRelease = null;
  state.lastZmPlan = null;
  byId("show-zm-plan").disabled = true;
  state.newLotImport = null;
  renderCurrentData();
  setStatus(`loaded: ${formatInteger(dataRows().length)} data rows, ${formatInteger(state.parameters.length)} parameters.`, false, true);
}

function prepareSource(sheetName) {
  const layout = {};
  const generated = state.generatedSources.get(sheetName);
  const table = generated || readWorkbookSheet(state.workbook, sheetName, layout);
  if (!table.length) throw new Error(`${sheetName} has no readable rows.`);
  const headerRow = findZonedHeaderRow(table) || (() => {
    const index = table.slice(0, 30).findIndex(row => headerIndex(row, "Lot") >= 0 && headerIndex(row, "N") >= 0);
    return index >= 0 ? { index, headers: table[index].map(text) } : null;
  })();
  if (!headerRow) throw new Error(`${state.workbookName}: ${sheetName} has no named headers such as Thick_1 through Thick_6.`);
  const headers = headerRow.headers;
  const rows = normalizeRows(table.slice(headerRow.index + 1), headers.length);
  const rowNumbers = generated ? table.map((_, index) => index + 1) : layout.rowNumbers;
  return { headers, rows, rowNumbers: rowNumbers.slice(headerRow.index + 1), columnOffset: layout.columnOffset || 0, generatedClean: state.generatedSources.has(sheetName) };
}

function readSheet(sheetName) {
  return readWorkbookSheet(state.workbook, sheetName);
}

function readWorkbookSheet(workbook, sheetName, layout = null) {
  const XLSX = getXlsx();
  const sheet = workbook?.Sheets?.[sheetName];
  if (!sheet) return [];
  let table;
  if (layout) {
    const range = XLSX.utils.decode_range(sheet["!ref"] || "A1");
    const columns = Array.from({ length: range.e.c - range.s.c + 1 }, (_, index) => String(index));
    // SheetJS retains physical row numbers on object rows without allocating blank rows.
    const sourceRows = XLSX.utils.sheet_to_json(sheet, { header: columns, raw: true, defval: "", blankrows: false });
    table = sourceRows.map(row => columns.map(column => row[column]));
    layout.columnOffset = range.s.c;
    layout.rowNumbers = table.flatMap((row, index) => isEmptyRow(normalizeRow(row)) ? [] : [sourceRows[index].__rowNum__ + 1]);
  } else table = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "", blankrows: false });
  // The wrapper prevents Array.map from passing the row index as normalizeRow's width.
  const trimmed = table.map((row) => normalizeRow(row)).filter((row) => !isEmptyRow(row));
  const width = trimmed.reduce((max, row) => Math.max(max, row.length), 0);
  return trimmed.map((row) => normalizeRow(row, width));
}

async function openNewLotWorkbook(file) {
  if (!file) return;
  if (!state.originalData) throw new Error("Open the workbook that contains Auswertung first.");
  setStatus(`Checking ${file.name}...`);
  await yieldToBrowser();
  const XLSX = getXlsx();
  const importData = await file.arrayBuffer();
  const importWorkbook = XLSX.read(importData, {
    type: "array",
    cellDates: true,
    cellFormula: true,
    cellHTML: false,
    cellText: true,
    dense: false
  });
  const sourceSheet = ["Auswertung", ...importWorkbook.SheetNames.filter((name) => name !== "Auswertung")]
    .find((name) => findRawHeaderRow(readWorkbookSheet(importWorkbook, name)));
  if (!sourceSheet) throw new Error("The new-lot workbook has no sheet with ChargenNr and Nummer headers.");
  const importTable = readWorkbookSheet(importWorkbook, sourceSheet);
  const existingTable = readAuswertungTable(state.originalData);
  const preview = mergeAuswertungTables(existingTable, importTable);
  state.newLotImport = { fileName: file.name, sourceSheet, importTable, preview };
  renderMergePreview();
  byId("merge-new-lots").disabled = Boolean(preview.missingHeaders.length || !preview.addedRows);
  if (preview.missingHeaders.length) {
    setStatus(`New-lot workbook is missing ${preview.missingHeaders.length} Auswertung headers.`, true);
  } else {
    setStatus(
      `${file.name}: ${formatInteger(preview.addedRows)} new rows ready, ${formatInteger(preview.duplicateRows)} duplicates skipped.`,
      false,
      true
    );
  }
}

function readAuswertungTable(data) {
  const XLSX = getXlsx();
  const workbook = XLSX.read(data.slice(0), {
    type: "array",
    cellDates: true,
    cellFormula: true,
    cellHTML: false,
    cellText: true,
    dense: false,
    sheets: ["Auswertung"]
  });
  const table = readWorkbookSheet(workbook, "Auswertung");
  if (!table.length) throw new Error("The current workbook does not contain a readable Auswertung sheet.");
  return table;
}

function renderMergePreview() {
  const container = byId("merge-preview");
  if (!state.newLotImport) {
    container.innerHTML = '<p class="empty-state">No new-lot workbook selected.</p>';
    byId("merge-new-lots").disabled = true;
    return;
  }
  const current = state.newLotImport;
  const result = current.preview;
  const warning = result.missingHeaders.length
    ? `<p class="merge-warning">Missing headers: ${escapeHtml(result.missingHeaders.join(", "))}</p>`
    : "";
  container.innerHTML = `
    <div class="metric-grid">
      ${metric("File", current.fileName)}
      ${metric("Source sheet", current.sourceSheet)}
      ${metric("Rows to add", formatInteger(result.addedRows))}
      ${metric("New lots", formatInteger(result.lots.length))}
      ${metric("Duplicates", formatInteger(result.duplicateRows))}
      ${metric("Incomplete", formatInteger(result.incompleteRows))}
    </div>
    ${warning}
    ${result.lots.length ? `<div class="table-wrap compact-table">${renderTable([
      ["New lots"],
      ...result.lots.map((lot) => [lot])
    ])}</div>` : ""}
  `;
}

async function mergeNewLots() {
  if (!state.newLotImport) throw new Error("Choose a new-lot workbook first.");
  if (state.newLotImport.preview.missingHeaders.length) throw new Error("The new-lot workbook is missing required Auswertung headers.");
  const fileName = writableWorkbookName();
  const saveHandle = await requestWorkbookSaveHandle(fileName);
  setStatus("Merging Auswertung and rebuilding Clean_Data and Summary...");
  await yieldToBrowser();
  const XLSX = getXlsx();
  const workbook = XLSX.read(state.originalData.slice(0), {
    type: "array",
    cellDates: true,
    cellFormula: true,
    cellHTML: false,
    cellText: true,
    cellStyles: true,
    bookVBA: true,
    dense: false
  });
  const existingTable = readWorkbookSheet(workbook, "Auswertung");
  const merged = mergeAuswertungTables(existingTable, state.newLotImport.importTable);
  if (!merged.addedRows) throw new Error("No new rows remain to merge.");
  appendWorkbookRows(workbook.Sheets.Auswertung, merged.appendedRows);

  const preserved = classificationMaps(state.headers, dataRows());
  const clean = buildCleanDataFromAuswertung(merged.table, {
    classificationByLot: state.lotClassifications,
    classificationByKey: preserved.byKey
  });
  replaceWorkbookSheet(workbook, "Clean_Data", [clean.headers, ...clean.rows]);
  const summary = buildFullSummary(clean.headers, clean.rows);
  replaceWorkbookSheet(workbook, "Summary", [summary.headers, ...summary.rows]);
  writeDataLabelsToWorkbook(workbook, clean.headers, clean.rows);

  const bytes = XLSX.write(workbook, {
    type: "array",
    bookType: workbookBookType(fileName),
    bookVBA: /\.xlsm$/i.test(fileName),
    cellDates: true,
    cellStyles: true,
    compression: true
  });
  const saveMode = await saveWorkbookBytes(bytes, fileName, saveHandle);
  byId("new-lot-file").value = "";
  await parseWorkbook(bytes, fileName);
  openPanel("merge-panel");
  setStatus(
    `${formatInteger(merged.addedRows)} rows merged into ${fileName}. ${saveMode === "direct" ? "Workbook replaced." : "Confirm Replace in Files."}`,
    false,
    true
  );
}

function replaceWorkbookSheet(workbook, name, rows) {
  const XLSX = getXlsx();
  workbook.Sheets[name] = XLSX.utils.aoa_to_sheet(rows, { cellDates: true });
  if (!workbook.SheetNames.includes(name)) workbook.SheetNames.push(name);
}

function appendWorkbookRows(sheet, rows) {
  if (!sheet || !rows.length) return;
  const XLSX = getXlsx();
  const range = sheet["!ref"] ? XLSX.utils.decode_range(sheet["!ref"]) : { s: { r: 0, c: 0 }, e: { r: -1, c: 0 } };
  let lastValueRow = range.e.r;
  while (lastValueRow >= range.s.r) {
    const hasValue = Array.from({ length: range.e.c - range.s.c + 1 }, (_, offset) => range.s.c + offset)
      .some((column) => text(sheet[XLSX.utils.encode_cell({ r: lastValueRow, c: column })]?.v) !== "");
    if (hasValue) break;
    lastValueRow -= 1;
  }
  XLSX.utils.sheet_add_aoa(sheet, rows, { origin: { r: lastValueRow + 1, c: 0 }, cellDates: true });
}

async function requestWorkbookSaveHandle(fileName) {
  if (typeof window.showSaveFilePicker !== "function") return null;
  try {
    return await window.showSaveFilePicker({
      suggestedName: fileName,
      types: [{
        description: "Excel workbook",
        accept: { [workbookMime(fileName)]: [fileExtension(fileName)] }
      }]
    });
  } catch (error) {
    if (error?.name === "AbortError") throw new Error("Workbook save was cancelled.");
    throw error;
  }
}

async function saveWorkbookBytes(bytes, fileName, handle) {
  if (handle) {
    const writable = await handle.createWritable();
    await writable.write(bytes);
    await writable.close();
    return "direct";
  }
  const file = new File([bytes], fileName, { type: workbookMime(fileName) });
  if (navigator.share && navigator.canShare?.({ files: [file] })) {
    try {
      await navigator.share({ files: [file], title: fileName });
      return "share";
    } catch (error) {
      if (error?.name === "AbortError") throw new Error("Workbook save was cancelled.");
    }
  }
  downloadBlob(fileName, file);
  return "download";
}

function writableWorkbookName() {
  const name = state.workbookName || "IPW_Analysis.xlsx";
  return /\.(?:xlsx|xlsm)$/i.test(name) ? name : name.replace(/\.[^.]+$/, "") + ".xlsx";
}

function workbookBookType(fileName) {
  return /\.xlsm$/i.test(fileName) ? "xlsm" : "xlsx";
}

function workbookMime(fileName) {
  return /\.xlsm$/i.test(fileName)
    ? "application/vnd.ms-excel.sheet.macroEnabled.12"
    : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
}

function fileExtension(fileName) {
  return /\.xlsm$/i.test(fileName) ? ".xlsm" : ".xlsx";
}

async function buildCleanData() {
  if (!state.originalData) throw new Error("Open a workbook first.");
  if (byId("clean-output").value === "corrected") return buildCorrectedData();
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

  const output = clean;
  const sourceName = GENERATED_CLEAN;

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
    corrected: false
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

async function buildCorrectedData() {
  const sourceName = ["Clean_Data", GENERATED_CLEAN].includes(state.source) ? state.source
    : state.generatedSources.has(GENERATED_CLEAN) ? GENERATED_CLEAN : "Clean_Data";
  if (!state.generatedSources.has(sourceName) && !state.workbook?.Sheets?.[sourceName]) {
    throw new Error("An uncorrected Clean_Data source is required. Build Clean_Data first; corrected data cannot be corrected again.");
  }
  setStatus("Correcting Wicking and Wicking_Q from Clean_Data...");
  await yieldToBrowser();
  const source = sourceName === state.source
    ? { headers: [...state.headers], rows: state.rows.map(row => [...row]) }
    : prepareSource(sourceName);
  ensureClassificationColumn(source.headers, source.rows);
  applyClassificationOverrides(source.headers, source.rows);
  const previousName = state.generatedSources.has(GENERATED_CORRECTED) ? GENERATED_CORRECTED
    : state.workbook?.Sheets?.Clean_Data_Cor ? "Clean_Data_Cor" : null;
  const previous = previousName ? prepareSource(previousName) : null;
  const output = correctCleanData(source.headers, source.rows,
    requiredNumber("reference-temperature"), requiredNumber("reference-humidity"), previous);
  applyClassificationOverrides(output.headers, output.rows);
  const summary = buildFullSummary(output.headers, output.rows);
  state.generatedSources.set(GENERATED_CORRECTED, [output.headers, ...output.rows]);
  state.lastBuild = {
    sourceName: GENERATED_CORRECTED, rows: output.rows.length,
    lots: getLotValues(output.rows, headerIndex(output.headers, "Lot")).length,
    parameters: getRegionalParameters(output.headers).length, summaryRows: summary.rows.length,
    corrected: true, correction: output.correction
  };
  refreshSourceSelect(GENERATED_CORRECTED);
  await selectSource(GENERATED_CORRECTED);
  byId("reference-temperature").value = output.correction.referenceTemperature;
  byId("reference-humidity").value = output.correction.referenceHumidity;
  saveAnalysisSettings();
  openPanel("clean-panel");
  setStatus(`Clean_Data_Cor and Summary created from ${sourceName}. ${formatInteger(output.correction.correctedRows)} rows corrected; ${formatInteger(output.correction.unchangedRows)} rows unchanged.`, false, true);
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
    if (state.classificationEdits.has(lot)) row[classificationIndex] = state.classificationEdits.get(lot);
  });
}

function applyLotClassification(selectedLot, selectedValue) {
  const lot = selectedLot ?? byId("classification-lot").value;
  if (!lot) throw new Error("Select a lot to classify.");
  const classification = (selectedValue ?? byId("classification-value").value).trim();
  state.lotClassifications.set(lot, classification);
  state.classificationEdits.set(lot, classification);
  state.classificationsUnsaved = true;
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
  syncClassificationInput();
  setStatus(`Classification ${classification || "cleared"} for lot ${lot}.`, false, true);
}

async function saveClassifications() {
  if (!state.originalData) throw new Error("Open a workbook first.");
  const fileName = writableWorkbookName();
  const saveHandle = await requestWorkbookSaveHandle(fileName);
  setStatus("Saving lot classifications...");
  await yieldToBrowser();
  const XLSX = getXlsx();
  const workbook = XLSX.read(state.originalData.slice(0), {
    type: "array", cellDates: true, cellFormula: true, cellStyles: true, bookVBA: true, dense: false
  });
  for (const [source, name] of [[GENERATED_CLEAN, "Clean_Data"], [GENERATED_CORRECTED, "Clean_Data_Cor"]]) {
    const table = state.generatedSources.get(source);
    if (!table) continue;
    replaceWorkbookSheet(workbook, name, table);
    if (name === "Clean_Data") {
      const summary = buildFullSummary(table[0], table.slice(1));
      replaceWorkbookSheet(workbook, "Summary", [summary.headers, ...summary.rows]);
    }
  }
  const savedEdits = new Map(state.classificationEdits);
  updateWorkbookClassifications(workbook, savedEdits, XLSX);
  const bytes = XLSX.write(workbook, {
    type: "array", bookType: workbookBookType(fileName), bookVBA: /\.xlsm$/i.test(fileName),
    cellDates: true, cellStyles: true, compression: true
  });
  const mode = await saveWorkbookBytes(bytes, fileName, saveHandle);
  state.originalData = bytes.slice(0);
  state.classificationsUnsaved = [...state.classificationEdits].some(([lot, value]) => savedEdits.get(lot) !== value);
  renderClassificationTable();
  setStatus(mode === "direct" ? "Classifications saved in your workbook." : "Save the updated workbook in Files and reopen that copy next time.", false, true);
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
  byId("build-clean-data").textContent = byId("clean-output").value === "corrected" ? "Create Corrected Data and Summary" : "Build Data and Summary";
}

function renderLabelParameterChoices() {
  const container = byId("label-parameters");
  container.replaceChildren();
  for (const parameter of state.parameters) {
    const label = document.createElement("label");
    const input = document.createElement("input");
    const caption = document.createElement("span");
    input.type = "checkbox";
    input.value = parameter;
    input.disabled = !state.headers.length;
    caption.textContent = parameter;
    label.append(input, caption);
    container.append(label);
  }
}

function syncLabelBatches() {
  if (!state.headers.length) return;
  const lotColumn = headerIndex(state.headers, "Lot");
  const batchColumn = headerIndex(state.headers, "N");
  const lot = byId("label-lot").value;
  const batches = new Set();
  for (const row of dataRows()) {
    if (sameDataValue(row[lotColumn], lot) && text(row[batchColumn])) batches.add(text(row[batchColumn]));
  }
  const values = [...batches].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
  const previous = byId("label-batch").value;
  fillSelect(byId("label-batch"), values, values.includes(previous) ? previous : values[0]);
  byId("label-batch").disabled = !values.length;
  updateLabelActionState();
}

function syncMasterRollControls() {
  const machine = byId("label-machine").value;
  const widthSelect = byId("label-roll-width");
  const rollInput = byId("label-roll-number");
  const automatic = Boolean(machine);
  if (!automatic) {
    widthSelect.replaceChildren();
    widthSelect.disabled = true;
    rollInput.disabled = true;
    byId("label-zone").disabled = !state.headers.length;
    byId("master-roll-result").hidden = true;
    updateLabelActionState();
    renderLabelSelectionPreview();
    return;
  }
  const widths = getMasterRollWidths(machine).map(String);
  const previous = widthSelect.value;
  fillSelect(widthSelect, widths, widths.includes(previous) ? previous : widths[0]);
  widthSelect.disabled = false;
  rollInput.disabled = false;
  byId("label-zone").disabled = true;
  updateMasterRollZone();
}

function currentMasterRollMapping() {
  const machine = byId("label-machine").value;
  if (!machine) {
    return { automatic: false, valid: true, zone: Number(byId("label-zone").value), result: "" };
  }
  const width = byId("label-roll-width").value;
  const roll = byId("label-roll-number").value.trim();
  const result = roll ? getMasterRollZone(machine, width, roll) : "";
  const match = result.match(/^Zone ([1-6])$/);
  return {
    automatic: true,
    valid: Boolean(match),
    machine,
    width,
    roll,
    zone: match ? Number(match[1]) : null,
    result
  };
}

function updateMasterRollZone() {
  const mapping = currentMasterRollMapping();
  const output = byId("master-roll-result");
  if (!mapping.automatic || !mapping.roll) {
    output.hidden = true;
    updateLabelActionState();
    return;
  }
  output.hidden = false;
  output.classList.toggle("is-error", !mapping.valid);
  if (!mapping.valid) {
    output.textContent = mapping.result === "Invalid Width"
      ? "This roll width is not available for the selected machine."
      : "This roll number is outside the valid range for the selected machine and width.";
    updateLabelActionState();
    return;
  }
  const previousZone = Number(byId("label-zone").value);
  byId("label-zone").value = String(mapping.zone);
  output.textContent = `Machine ${mapping.machine}, width ${mapping.width}, roll ${mapping.roll}: Zone ${mapping.zone}.`;
  updateLabelActionState();
  if (mapping.zone !== previousZone) syncLabelEditor();
  else renderLabelSelectionPreview();
}

function updateLabelActionState() {
  const mapping = currentMasterRollMapping();
  const ready = Boolean(state.headers.length && byId("label-batch").value && mapping.valid);
  byId("save-data-label").disabled = !ready;
  byId("remove-data-label").disabled = !ready || !selectedDataLabel();
}

function syncLabelEditor() {
  if (!state.headers.length) return;
  const existing = selectedDataLabel();
  const selected = new Set(existing?.parameters || []);
  byId("label-parameters").querySelectorAll("input").forEach((input) => {
    input.checked = selected.has(input.value);
  });
  byId("label-text").value = existing?.label || "";
  byId("label-comment").value = existing?.comment || "";
  byId("label-notes").value = existing?.notes || "";
  updateLabelActionState();
  renderLabelSelectionPreview();
}

function selectedDataLabel() {
  return findDataLabel(
    state.dataLabels,
    byId("label-lot").value,
    byId("label-batch").value,
    Number(byId("label-zone").value)
  );
}

function selectedLabelParameters() {
  return [...byId("label-parameters").querySelectorAll("input:checked")].map((input) => input.value);
}

function renderLabelSelectionPreview() {
  const container = byId("label-selection-preview");
  if (!state.headers.length || !byId("label-batch").value) {
    container.innerHTML = '<p class="empty-state">No batch selected.</p>';
    return;
  }
  const lot = byId("label-lot").value;
  const batch = byId("label-batch").value;
  const zone = Number(byId("label-zone").value);
  const lotColumn = headerIndex(state.headers, "Lot");
  const batchColumn = headerIndex(state.headers, "N");
  const row = dataRows().find((item) => sameDataValue(item[lotColumn], lot) && sameDataValue(item[batchColumn], batch));
  const parameters = selectedLabelParameters();
  const mapping = currentMasterRollMapping();
  const values = parameters.map((parameter) => {
    const column = zoneColumns(state.headers, parameter)?.[zone - 1] ?? -1;
    return [parameter, column >= 0 ? row?.[column] : ""];
  });
  container.innerHTML = `
    <div class="metric-grid">
      ${metric("Lot", lot)}
      ${metric("Batch N", batch)}
      ${metric("Zone", zone)}
      ${metric("Saved", selectedDataLabel() ? "Yes" : "No")}
      ${mapping.automatic ? metric("Master roll", `${mapping.machine} / ${mapping.width} / ${mapping.roll || "-"}`) : ""}
    </div>
    ${values.length ? `<div class="table-wrap compact-table">${renderTable([
      ["Parameter", `Zone ${zone} value`],
      ...values
    ])}</div>` : ""}
  `;
}

function renderDataLabelsTable() {
  const container = byId("data-labels-table");
  if (!state.dataLabels.length) {
    container.innerHTML = '<p class="empty-state">No saved labels.</p>';
    return;
  }
  container.innerHTML = renderTable([
    ["Lot", "Batch N", "Zone", "Parameters", "Values", "Label", "Comment", "Notes", "Updated"],
    ...state.dataLabels.map((item) => [
      item.lot,
      item.batch,
      `Zone ${item.zone}`,
      item.parameters.join(", "),
      item.valuesText,
      item.label,
      item.comment,
      item.notes,
      formatDateTime(item.updated)
    ])
  ]);
}

async function saveDataLabel() {
  const mapping = currentMasterRollMapping();
  if (mapping.automatic && !mapping.valid) {
    throw new Error("Enter a valid master roll number before saving the label.");
  }
  const previous = state.dataLabels;
  const label = buildDataLabel(state.headers, dataRows(), {
    lot: byId("label-lot").value,
    batch: byId("label-batch").value,
    zone: byId("label-zone").value,
    parameters: selectedLabelParameters(),
    label: byId("label-text").value,
    comment: byId("label-comment").value,
    notes: byId("label-notes").value
  });
  state.dataLabels = upsertDataLabel(state.dataLabels, label);
  try {
    const saveMode = await saveLabelsToCurrentWorkbook();
    renderDataLabelsTable();
    renderPreviewTable();
    syncLabelEditor();
    setStatus(
      `Label saved in ${writableWorkbookName()}. ${saveMode === "direct" ? "Workbook replaced." : "Confirm Replace in Files."}`,
      false,
      true
    );
  } catch (error) {
    state.dataLabels = previous;
    syncLabelEditor();
    renderDataLabelsTable();
    throw error;
  }
}

async function removeSelectedDataLabel() {
  const existing = selectedDataLabel();
  if (!existing) throw new Error("No saved label matches that Lot, Batch N, and Zone.");
  const previous = state.dataLabels;
  state.dataLabels = removeDataLabel(state.dataLabels, existing.lot, existing.batch, existing.zone);
  try {
    const saveMode = await saveLabelsToCurrentWorkbook();
    renderDataLabelsTable();
    renderPreviewTable();
    syncLabelEditor();
    setStatus(
      `Label removed from ${writableWorkbookName()}. ${saveMode === "direct" ? "Workbook replaced." : "Confirm Replace in Files."}`,
      false,
      true
    );
  } catch (error) {
    state.dataLabels = previous;
    syncLabelEditor();
    renderDataLabelsTable();
    throw error;
  }
}

async function saveLabelsToCurrentWorkbook() {
  const fileName = writableWorkbookName();
  const saveHandle = await requestWorkbookSaveHandle(fileName);
  setStatus("Saving labels and highlighted values...");
  await yieldToBrowser();
  const XLSX = getXlsx();
  const workbook = XLSX.read(state.originalData.slice(0), {
    type: "array",
    cellDates: true,
    cellFormula: true,
    cellHTML: false,
    cellText: true,
    cellStyles: true,
    bookVBA: true,
    dense: false
  });
  if (!workbook.Sheets.Clean_Data) {
    const generated = state.generatedSources.get(GENERATED_CLEAN);
    if (generated) {
      replaceWorkbookSheet(workbook, "Clean_Data", generated);
      const summary = buildFullSummary(generated[0], generated.slice(1));
      replaceWorkbookSheet(workbook, "Summary", [summary.headers, ...summary.rows]);
    }
  }
  writeDataLabelsToWorkbook(workbook, state.headers, dataRows());
  const bytes = XLSX.write(workbook, {
    type: "array",
    bookType: workbookBookType(fileName),
    bookVBA: /\.xlsm$/i.test(fileName),
    cellDates: true,
    cellStyles: true,
    compression: true
  });
  const saveMode = await saveWorkbookBytes(bytes, fileName, saveHandle);
  state.originalData = bytes.slice(0);
  return saveMode;
}

function writeDataLabelsToWorkbook(workbook, fallbackHeaders, fallbackRows) {
  const cleanTable = readWorkbookSheet(workbook, "Clean_Data");
  const cleanHeader = findZonedHeaderRow(cleanTable);
  const headers = cleanHeader?.headers || fallbackHeaders;
  const rows = cleanHeader
    ? normalizeRows(cleanTable.slice(cleanHeader.index + 1), headers.length).filter((row) => isDataRow(headers, row))
    : fallbackRows;
  replaceWorkbookSheet(workbook, "Data_Labels", buildDataLabelsTable(state.dataLabels));
  replaceWorkbookSheet(workbook, "Data_Label_Values", buildDataLabelValuesTable(headers, rows, state.dataLabels));
  styleDataLabelsInSheet(workbook, "Clean_Data");
  styleDataLabelsInSheet(workbook, "Clean_Data_Cor");
}

function styleDataLabelsInSheet(workbook, sheetName) {
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return;
  clearAppCellComments(sheet);
  if (!state.dataLabels.length) return;
  const table = readWorkbookSheet(workbook, sheetName);
  const headerRow = findZonedHeaderRow(table);
  if (!headerRow) return;
  const headers = headerRow.headers;
  const rows = normalizeRows(table.slice(headerRow.index + 1), headers.length);
  for (const item of state.dataLabels) {
    for (const coordinate of dataLabelCellCoordinates(headers, rows, [item])) {
      styleWorkbookCell(sheet, headerRow.index + 1 + coordinate.row, coordinate.column, item);
    }
  }
}

function clearAppCellComments(sheet) {
  for (const [address, cell] of Object.entries(sheet)) {
    if (address.startsWith("!") || !cell?.c) continue;
    cell.c = cell.c.filter((comment) => comment.a !== "IPW Analysis");
    if (!cell.c.length) delete cell.c;
  }
}

function styleWorkbookCell(sheet, row, column, item) {
  if (row < 0 || column < 0) return;
  const XLSX = getXlsx();
  const address = XLSX.utils.encode_cell({ r: row, c: column });
  const cell = sheet[address];
  if (!cell) return;
  cell.s = {
    ...(cell.s || {}),
    fill: { patternType: "solid", fgColor: { rgb: "FFFF6666" } },
    font: { ...(cell.s?.font || {}), bold: true }
  };
  const details = [
    "[IPW LABEL]",
    `Lot: ${text(item.lot)}`,
    `Batch N: ${text(item.batch)}`,
    `Zone: ${item.zone}`,
    `Parameters: ${item.parameters.join(", ")}`,
    ...(item.label ? [`Label: ${item.label}`] : []),
    ...(item.comment ? [`Comment: ${item.comment}`] : []),
    ...(item.notes ? [`Notes: ${item.notes}`] : []),
    ...(item.updated ? [`Updated: ${formatDateTime(item.updated)}`] : [])
  ].join("\n");
  const existing = cell.c?.find((comment) => comment.a === "IPW Analysis");
  if (existing) existing.t += `\n\n${details}`;
  else cell.c = [...(cell.c || []), { a: "IPW Analysis", t: details }];
}

function sameDataValue(first, second) {
  if (isNumeric(first) && isNumeric(second)) return number(first) === number(second);
  return text(first).toUpperCase() === text(second).toUpperCase();
}

function populateWorkbookControls() {
  const parameter = state.parameters[0];
  const secondParameter = state.parameters[1] || state.parameters[0];
  fillSelect(byId("type-filter"), [ALL, ...state.types], ALL, (value) => value === ALL ? "All types" : value);
  fillSelect(byId("classification-lot"), state.lots, byId("classification-lot").value || state.lots[0]);
  fillSelect(byId("label-lot"), state.lots, byId("label-lot").value || state.lots[0]);
  [
    "summary-parameter",
    "generated-summary-parameter",
    "gaussian-parameter",
    "trend-parameter",
    "correlation-y",
    "correlation-x",
    "export-parameter"
  ].forEach((id) => fillSelect(byId(id), state.parameters, byId(id).value || parameter));
  fillSelect(byId("trend-parameter"), state.trendParameters, byId("trend-parameter").value || state.trendParameters[0]);
  fillSelect(byId("assessment-parameter"), ["All parameters", ...state.v90Parameters], byId("assessment-parameter").value || "All parameters");
  fillSelect(byId("period-parameter"), ["All parameters", ...state.v90Parameters], byId("period-parameter").value || "All parameters");
  syncPeriodLots();
  fillSelect(byId("release-lot"), state.lots, byId("release-lot").value || state.lots[0]);
  byId("generated-summary-parameter").disabled = !state.lastBuild;
  if (byId("correlation-x").options.length > 1) byId("correlation-x").value = secondParameter;
  enableControls([
    "correlation-analysis", "correlation-method", "correlation-compare", "correlation-min-n", "correlation-coverage", "correlation-min-regions", "correlation-expected-regions",
    "filter-mode",
    "classification-keyword",
    "limit-lot-period",
    "type-filter",
    "clean-output",
    "reference-temperature",
    "reference-humidity",
    "build-clean-data",
    "new-lot-file",
    "classification-lot",
    "classification-value",
    "apply-classification",
    "label-lot",
    "label-batch",
    "label-machine",
    "label-roll-width",
    "label-roll-number",
    "label-zone",
    "label-text",
    "label-comment",
    "label-notes",
    "save-data-label",
    "summary-parameter",
    "gaussian-data-scope",
    "gaussian-parameter",
    "gaussian-method",
    "gaussian-bin-width",
    "gaussian-start",
    "gaussian-end",
    "gaussian-extreme-sigma",
    "gaussian-extreme-side",
    "gaussian-visible-rows",
    "gaussian-full-range",
    "recommend-gaussian",
    "run-gaussian",
    "trend-data-scope",
    "trend-parameter",
    "trend-start",
    "trend-end",
    "trend-window",
    "run-trend",
    "period-data-scope",
    "period-mode",
    "period-parameter",
    "period-plot",
    "period-lot",
    "period-lot-b",
    "period-a-start",
    "period-a-end",
    "period-b-start",
    "period-b-end",
    "period-c-start",
    "period-c-end",
    "run-period",
    "release-lot",
    "release-reference",
    "release-monitor",
    "release-not-ok",
    "run-release",
    "assessment-parameter",
    "assessment-lot",
    "assessment-reference",
    "assessment-reference-lot",
    "assessment-granularity",
    "assessment-monitor",
    "assessment-outlier",
    "assessment-mu",
    "assessment-sigma",
    "run-assessment",
    "zm-layout",
    "zm-mark-label",
    "zm-coordinates",
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
  byId("period-parameter").disabled = !state.v90Parameters.length;
  byId("run-period").disabled = !state.v90Parameters.length;
  byId("release-lot").disabled = !state.lots.length;
  byId("run-release").disabled = !state.lots.length;
  byId("assessment-parameter").disabled = !state.v90Parameters.length;
  byId("save-period-plot").disabled = !state.lastPeriod;
  byId("download-release").disabled = !state.lastRelease;
  byId("show-zm-plan").disabled = !state.lastAssessment || state.lastAssessment.parameter === "All parameters";
  byId("save-gaussian-snapshot").disabled = !state.lastGaussian;
  byId("export-gaussian-extremes").disabled = !state.lastGaussian;
  syncGaussianMethod();
  syncZoneChoices("gaussian-zones", byId("gaussian-parameter").value);
  syncZoneChoices("export-zones", byId("export-parameter").value);
  syncCorrelationZones();
  syncCorrelationRemovalInput();
  syncStructuredCorrelationControls();
  byId("run-gaussian").disabled = !state.parameters.length;
  byId("recommend-gaussian").disabled = !state.parameters.length;
  byId("run-trend").disabled = !state.trendParameters.length;
  syncCorrectionInputs();
  syncPeriodMode();
  syncClassificationInput();
  syncFilterSelections();
  renderFilterOptions();
  syncAssessmentLots();
  syncAssessmentReferenceMode();
  renderLabelParameterChoices();
  syncLabelBatches();
  syncMasterRollControls();
  syncLabelEditor();
  renderDataLabelsTable();
}

function renderCurrentData() {
  if (!state.headers.length) return;
  syncAssessmentLots();
  renderDataMetrics();
  renderSummaryTable();
  renderPreviewTable();
  renderClassificationTable();
  renderMergePreview();
  renderDataLabelsTable();
  renderLabelSelectionPreview();
  renderBuildResult();
  renderGeneratedSummaryTable();
  clearResult("gaussian-result");
  clearResult("gaussian-extremes");
  clearResult("trend-result");
  clearResult("period-result");
  clearResult("correlation-result");
  clearResult("assessment-result");
  clearResult("zm-plan-result");
  clearResult("release-result");
  renderExportNote();
  renderGaussianSnapshots();
}

function invalidateAnalyses() {
  invalidateGaussian();
  invalidateTrend();
  invalidatePeriod();
  invalidateCorrelation();
  invalidateAssessment();
  invalidateRelease();
}

function invalidateCorrelation() {
  state.lastCorrelation = null;
  state.lastStructuredCorrelation = null;
  clearResult("correlation-result");
}

function invalidateGaussian() {
  state.lastGaussian = null;
  byId("save-gaussian-snapshot").disabled = true;
  byId("export-gaussian-extremes").disabled = true;
  clearResult("gaussian-result");
  clearResult("gaussian-extremes");
}

function invalidateTrend() {
  state.lastTrend = null;
  clearResult("trend-result");
}

function invalidatePeriod() {
  state.lastPeriod = null;
  byId("save-period-plot").disabled = true;
  clearResult("period-result");
}

function invalidateRelease() {
  state.lastRelease = null;
  byId("download-release").disabled = true;
  clearResult("release-result");
}

function invalidateAssessment() {
  byId("equal-lot-dialog").close();
  state.lastAssessment = null;
  byId("show-zm-plan").disabled = true;
  clearResult("assessment-result");
  invalidateZmPlan();
}

function invalidateZmPlan() {
  state.lastZmPlan = null;
  byId("zm-label-measurement").disabled = true;
  byId("zm-label-roll").disabled = true;
  byId("apply-zm-label").disabled = true;
  byId("remove-zm-label").disabled = true;
  byId("save-zm-plan-png").disabled = true;
  clearResult("zm-plan-result");
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
  byId("classification-keyword-panel").hidden = mode !== "classification-keyword";
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
  byId("save-classifications").disabled = !state.originalData || !state.lots.length;
  byId("classification-save-status").textContent = state.classificationsUnsaved ? "Unsaved workbook changes" : "No pending classification changes";
  if (!state.lots.length) {
    byId("classification-count").textContent = "0 lots";
    byId("classification-table").innerHTML = '<p class="empty-state">No lots.</p>';
    return;
  }
  const lotIndex = headerIndex(state.headers, "Lot");
  const classificationIndex = headerIndex(state.headers, "Classification");
  const valuesByLot = new Map();
  for (const row of state.rows) {
    const lot = text(row[lotIndex]);
    if (!valuesByLot.has(lot)) valuesByLot.set(lot, new Set());
    valuesByLot.get(lot).add(text(row[classificationIndex]));
  }
  const query = byId("classification-search").value.trim().toLowerCase();
  const rows = state.lots.flatMap((lot) => {
    const values = [...(valuesByLot.get(lot) || [])];
    const mixed = values.length > 1;
    const classification = mixed ? "" : (values[0] || "");
    if (query && !`${lot} ${values.join(" ")}`.toLowerCase().includes(query)) return [];
    return [`<tr><th scope="row">${escapeHtml(lot)}</th><td><input data-lot="${escapeHtml(lot)}" aria-label="Classification for lot ${escapeHtml(lot)}" value="${escapeHtml(classification)}" placeholder="${mixed ? "Mixed classifications" : "Unclassified"}">${mixed ? `<small>${escapeHtml(values.map(v => v || "Unclassified").join("; "))}</small>` : ""}</td></tr>`];
  });
  byId("classification-count").textContent = `${rows.length} of ${state.lots.length} lots`;
  byId("classification-table").innerHTML = `<table class="classification-list"><thead><tr><th>Lot</th><th>Classification</th></tr></thead><tbody>${rows.join("")}</tbody></table>`;
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
      ${result.correction ? `
        ${metric("Rows corrected", formatInteger(result.correction.correctedRows))}
        ${metric("Rows unchanged", formatInteger(result.correction.unchangedRows))}
        ${metric("Reference temperature (C)", formatNumber(result.correction.referenceTemperature))}
        ${metric("Reference humidity (%)", formatNumber(result.correction.referenceHumidity))}
        ${metric("Missing Temp. / Humidity", formatInteger(result.correction.missingEnvironmentRows))}
        ${metric("Invalid correction factor", formatInteger(result.correction.invalidFactorRows))}
      ` : `${metric("Probe rows removed", formatInteger(result.removedProbeRows))}
      ${metric("VE lots removed", formatInteger(result.removedVeLots))}`}
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
  const period = lotPeriodSelection();
  const rows = period.rows;
  const lotColumn = headerIndex(state.headers, "Lot");
  const lotCount = lotColumn >= 0 ? new Set(rows.map((row) => text(row[lotColumn])).filter(Boolean)).size : 0;
  byId("classification-keyword-count").textContent = `${formatInteger(lotCount)} matching lots, ${formatInteger(rows.length)} rows`;
  byId("lot-period-status").textContent = period.error || (byId("limit-lot-period").checked
    ? `${formatInteger(lotCount)} matching lots, ${formatInteger(rows.length)} rows. ${formatInteger(period.undatedLotCount)} selected lots without dates.`
    : "All selected lots; no date limit.");
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
  if (!rows.length) {
    byId("preview-table").innerHTML = '<p class="empty-state">No data rows.</p>';
    return;
  }
  byId("preview-table").innerHTML = `
    <table>
      <thead><tr>${headers.map((cell) => `<th>${escapeHtml(cell)}</th>`).join("")}</tr></thead>
      <tbody>${rows.map((row) => `<tr>${headers.map((_, index) => {
        const className = previewCellIsLabeled(row, index) ? ' class="labeled-cell"' : "";
        return `<td${className}>${formatCell(row[index])}</td>`;
      }).join("")}</tr>`).join("")}</tbody>
    </table>
  `;
}

function previewCellIsLabeled(row, column) {
  if (!state.dataLabels.length) return false;
  const lotColumn = headerIndex(state.headers, "Lot");
  const batchColumn = headerIndex(state.headers, "N");
  const labels = state.dataLabels.filter((item) =>
    sameDataValue(row[lotColumn], item.lot) && sameDataValue(row[batchColumn], item.batch)
  );
  if (!labels.length) return false;
  if (column === lotColumn || column === batchColumn) return true;
  const match = text(state.headers[column]).match(/^(.+)_([1-6])$/);
  if (!match) return false;
  const parameter = match[1].toUpperCase();
  const zone = Number(match[2]);
  return labels.some((item) => item.zone === zone &&
    item.parameters.some((value) => text(value).toUpperCase() === parameter));
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
  const records = collectGaussianRecords(parameter, includedZones);
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

function collectGaussianRecords(parameter, includedZones) {
  const columns = zoneColumns(state.headers, parameter), records = [], XLSX = getXlsx();
  const lotColumn = headerIndex(state.headers, "Lot"), batchColumn = headerIndex(state.headers, "N");
  const sheet = state.workbook.Sheets[state.source];
  for (const row of rowsForAnalysis("gaussian-data-scope")) {
    const location = state.sourceLocations.get(row);
    if (!location) throw new Error("Source row mapping is unavailable. Reopen the workbook before fitting.");
    if (byId("gaussian-visible-rows").checked && sheet?.["!rows"]?.[location.row - 1]?.hidden) continue;
    for (const zone of includedZones) {
      const column = columns?.[zone - 1];
      if (column === undefined || column < 0) continue;
      const value = row[column], sourceCell = XLSX.utils.encode_cell({ r: location.row - 1, c: state.sourceColumns[column] });
      if (sheet?.[sourceCell]?.t === "e" || !["number", "string"].includes(typeof value) || text(value) === "" || !Number.isFinite(Number(value))) continue;
      records.push({ value: Number(value), lot: lotColumn >= 0 ? text(row[lotColumn]) : "", batch: batchColumn >= 0 ? text(row[batchColumn]) : "", zone,
        source: state.source, sourceRow: location.row, sourceCell, parameter });
    }
  }
  return records;
}

function createGaussian() {
  invalidateGaussian();
  const parameter = byId("gaussian-parameter").value;
  const includedZones = selectedZones("gaussian-zones");
  if (!includedZones.length) throw new Error("Select at least one Zone.");
  const method = byId("gaussian-method").value;
  if (method === "nacl-truncated" && parameter.toLowerCase() !== "nacl") {
    throw new Error("NaCl truncated fitting is available only for the NaCl parameter.");
  }
  const scope = byId("gaussian-data-scope").value;
  const records = collectGaussianRecords(parameter, includedZones);
  const fit = gaussianFitWithOptions(
    records.map((record) => record.value),
    optionalNumber("gaussian-bin-width"),
    optionalNumber("gaussian-start"),
    optionalNumber("gaussian-end"),
    { method, lowerLimit: 1 }
  );
  const fittedRecords = records.filter((record) => fit.method === "robust-huber" || record.value >= fit.start && record.value <= fit.end &&
    (fit.method !== "nacl-truncated" || record.value >= fit.lowerLimit));
  const lots = new Set(fittedRecords.map((record) => record.lot).filter(Boolean));
  const batches = new Set(fittedRecords.map((record) => `${record.lot}|${record.batch}`).filter((key) => key !== "|"));
  const extremes = gaussianExtremeSnapshot(fittedRecords, fit, requiredNumber("gaussian-extreme-sigma"));
  state.lastGaussian = { parameter, zones: includedZones, scope, visibleExcelRowsOnly: byId("gaussian-visible-rows").checked, source: state.source, savedAt: new Date().toISOString(), extremes, fit: { ...fit, lotCount: lots.size, batchCount: batches.size } };
  renderGaussianResult();
  renderGaussianExtremes();
  byId("save-gaussian-snapshot").disabled = false;
  byId("export-gaussian-extremes").disabled = false;
  setStatus(
    fit.method === "robust-huber"
      ? `Robust Gaussian: ${formatInteger(fit.n)} points fitted, ${formatInteger(fit.outsideHistogram)} outside the histogram, ${formatInteger(lots.size)} lots.`
      : `Gaussian fit: ${formatInteger(fit.n)} points used, ${formatInteger(fit.excluded)} excluded, ${formatInteger(lots.size)} lots.`,
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
      ${metric("Mean", fit.mean.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }))}
      ${metric("Sigma", fit.sigma.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 }))}
      ${metric("SSE", formatNumber(fit.sse, 2))}
      ${metric(fit.method === "robust-huber" ? "Histogram range" : "Fit range", `${formatNumber(fit.start, 3)} to ${formatNumber(fit.end, 3)}`)}
      ${fit.method === "robust-huber" ? metric("Outside histogram (still fitted)", formatInteger(fit.outsideHistogram)) + metric("Huber iterations", fit.iterations) : ""}
    </div>
    <details class="gaussian-chart-section" open><summary>Observed Histogram + Fitted Gaussian</summary>
      <div class="chart-card"><canvas id="gaussian-chart" aria-label="Observed histogram with fitted Gaussian curve and percentile cutoffs"></canvas></div></details>
    <div class="table-wrap mini-table compact-table cutoff-table">${renderTable([
      ["Percentile cutoff", "Value"],
      ["2.5%", fit.low25],
      ["15%", fit.low15],
      ["85%", fit.high15],
      ["97.5%", fit.high25]
    ])}</div>
    <div class="table-wrap mini-table">${renderTable([
      ["Bin", "Observed", "Gaussian"],
      ...fit.bins.slice(0, 40).map((bin) => [bin.center, bin.observed, bin.gaussian])
    ])}</div>
  `;
  requestAnimationFrame(drawGaussianCharts);
  byId("gaussian-result").querySelectorAll(".gaussian-chart-section").forEach(section => section.addEventListener("toggle", () => { if (section.open) requestAnimationFrame(drawGaussianCharts); }));
}

function drawGaussianCharts() {
  if (!state.lastGaussian) return;
  drawGaussian(byId("gaussian-chart"), state.lastGaussian.fit, "combined");
}

function selectedGaussianExtremes() {
  const side = byId("gaussian-extreme-side").value;
  return state.lastGaussian?.extremes.records.filter(record => side === "Both" || record.side === side) || [];
}

function gaussianExtremesTable(records) {
  return [["Source sheet", "Source row", "Source cell", "Lot", "N / Batch", "Zone", "Parameter", "Value", "Z-score (fit)", "Classification", "Mu", "Sigma", "Sigma multiplier"],
    ...records.map(record => [record.source, record.sourceRow, record.sourceCell, record.lot, record.batch, record.zone, record.parameter, record.value, record.zScore, `${record.side} extreme`, record.mean, record.sigma, record.multiplier])];
}

function renderGaussianExtremes() {
  const current = state.lastGaussian;
  if (!current) return;
  const { extremes } = current, records = selectedGaussianExtremes();
  byId("gaussian-extremes").innerHTML = `<div class="metric-grid">${metric("Low extremes", extremes.lowCount)}${metric("High extremes", extremes.highCount)}
    ${metric("Lower boundary", formatNumber(extremes.lower, 2))}${metric("Upper boundary", formatNumber(extremes.upper, 2))}</div>
    <h3>Extreme Cases</h3>${records.length ? `<div class="table-wrap">${renderTable(gaussianExtremesTable(records))}</div>` : '<p class="empty-state">No extreme cases for the selected side.</p>'}`;
}

function exportGaussianExtremes() {
  const current = state.lastGaussian;
  if (!current) throw new Error("Refresh the Gaussian fit before exporting extremes.");
  const records = selectedGaussianExtremes(), side = byId("gaussian-extreme-side").value;
  const table = [["Gaussian extreme cases"], ["Fit method", current.fit.method], ["Source", current.source], ["Data scope", current.scope], ["Fit snapshot", current.savedAt],
    ["Parameter", current.parameter], ["Zones", current.zones.join(", ")], ["Side", side], ["Fit N", current.fit.n], ["Mu", current.fit.mean], ["Sigma", current.fit.sigma],
    ["Sigma multiplier", current.extremes.multiplier], ["Lower boundary", current.extremes.lower], ["Upper boundary", current.extremes.upper], ["Exported measurements", records.length], [],
    ...gaussianExtremesTable(records), ...(records.length ? [] : [["No extreme cases for the selected side."]])];
  const index = state.gaussianExtremeExports.length + 1;
  const name = `Gaussian_Extremes${index > 1 ? `_${index}` : ""}`;
  const workbook = getXlsx().utils.book_new();
  appendSheet(workbook, name, table);
  getXlsx().writeFile(workbook, `${baseFileName()}_${name}.xlsx`);
  state.gaussianExtremeExports.push({ name, table });
  setStatus(`Exported ${records.length} extreme measurements (${side}).`, false, true);
}

function saveGaussianSnapshot() {
  const current = state.lastGaussian;
  const histogram = byId("gaussian-chart");
  if (!current || !histogram) throw new Error("Run a Gaussian fit before saving a snapshot.");
  drawGaussianCharts();
  const canvas = document.createElement("canvas");
  canvas.width = histogram.width;
  canvas.height = histogram.height + 40;
  const ctx = canvas.getContext("2d");
  ctx.fillStyle = "#ffffff"; ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.fillStyle = "#172231"; ctx.font = "18px sans-serif";
  ctx.fillText(`${current.parameter} | ${current.fit.method} | Mu ${current.fit.mean.toFixed(2)} | Sigma ${current.fit.sigma.toFixed(2)}`, 12, 24, canvas.width - 24);
  ctx.drawImage(histogram, 0, 40);
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
      <div class="chart-key">${ZONES.map((zone, index) => `<label><input type="checkbox" data-trend-zone="${zone}" checked> <span style="--key-color:${TREND_COLORS[index]}">Zone ${zone}</span></label>`).join("")}</div>
      <canvas id="trend-zone-chart" aria-label="Zone values over production time"></canvas>
    </article>
    <article class="chart-card">
      <h3>Persistent Zone Bias</h3>
      <p>Average difference between each zone and its batch mean. Positive means higher; negative means lower. Error bars show the 95% confidence interval. An interval excluding zero indicates a consistent offset, not necessarily an out-of-spec result.</p>
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
        <div class="chart-key">${[["batch", "Batch mean", TREND_COLORS[0]], ["rolling", `Rolling ${trend.rollingWindow}`, TREND_COLORS[2]], ["center", "Center", TREND_COLORS[1]], ["limits", "+/-2 SD", TREND_COLORS[3]]].map(([key, label, color]) => `<label><input type="checkbox" data-trend-series="${key}" checked> <span style="--key-color:${color}">${label}</span></label>`).join("")}</div>
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
  byId("trend-result").querySelectorAll("[data-trend-series], [data-trend-zone]").forEach(input => input.addEventListener("change", drawTrendCharts));
}

function chartLegend(items) {
  return `<div class="chart-key">${items.map(([label, color]) =>
    `<span style="--key-color:${color}">${escapeHtml(label)}</span>`
  ).join("")}</div>`;
}

function syncPeriodMode() {
  const mode = byId("period-mode").value;
  const lotDatasetA = mode === "lot-vs-period" || mode === "lot-vs-lot";
  const lotDatasetB = mode === "lot-vs-lot";
  byId("period-a-period").hidden = lotDatasetA;
  byId("period-a-lot").hidden = !lotDatasetA;
  byId("period-b-period").hidden = lotDatasetB;
  byId("period-b-lot").hidden = !lotDatasetB;
  byId("period-c-period").hidden = mode !== "three-periods";
  byId("period-lot").disabled = !lotDatasetA || !state.lots.length;
  byId("period-lot-b").disabled = !lotDatasetB || byId("period-lot-b").options.length === 0;
  ["period-a-start", "period-a-end"].forEach((id) => { byId(id).disabled = lotDatasetA; });
  ["period-b-start", "period-b-end"].forEach((id) => { byId(id).disabled = lotDatasetB; });
  ["period-c-start", "period-c-end"].forEach((id) => { byId(id).disabled = mode !== "three-periods"; });
}

function syncPeriodLots() {
  const previousA = byId("period-lot").value;
  const lotA = state.lots.includes(previousA) ? previousA : state.lots[0];
  fillSelect(byId("period-lot"), state.lots, lotA);
  const lotBOptions = state.lots.filter((lot) => text(lot).toUpperCase() !== text(lotA).toUpperCase());
  const previousB = byId("period-lot-b").value;
  fillSelect(byId("period-lot-b"), lotBOptions, lotBOptions.includes(previousB) ? previousB : lotBOptions[0]);
  syncPeriodMode();
}

async function createPeriodComparison() {
  const selectedParameter = byId("period-parameter").value;
  const parameters = selectedParameter === "All parameters" ? state.v90Parameters : [selectedParameter];
  if (!parameters.length) throw new Error("No V90 parameters are available.");
  const comparisonMode = byId("period-mode").value;
  setStatus(comparisonMode === "lot-vs-lot" ? "Creating the Lot comparison..." : "Reading production dates and creating the comparison...");
  const dateLookup = comparisonMode === "lot-vs-lot" ? new Map() : await loadTrendDateLookup();
  const options = {
    mode: comparisonMode,
    lotA: byId("period-lot").value,
    lotB: byId("period-lot-b").value,
    aStart: byId("period-a-start").value,
    aEnd: byId("period-a-end").value,
    bStart: byId("period-b-start").value,
    bEnd: byId("period-b-end").value,
    cStart: byId("period-c-start").value,
    cEnd: byId("period-c-end").value
  };
  const rows = rowsForAnalysis("period-data-scope");
  const results = [];
  const skipped = [];
  for (const parameter of parameters) {
    try {
      results.push(buildPeriodComparison(state.headers, rows, parameter, dateLookup, options));
    } catch (error) {
      if (parameters.length === 1) throw error;
      skipped.push(`${parameter}: ${error.message}`);
    }
    if (parameters.length > 4) await yieldToBrowser();
  }
  if (!results.length) throw new Error(skipped[0] || "No parameter comparison could be created.");
  state.lastPeriod = {
    results,
    skipped,
    plot: byId("period-plot").value,
    scope: byId("period-data-scope").value,
    mode: options.mode
  };
  renderPeriodResult();
  byId("save-period-plot").disabled = false;
  setStatus(
    `${formatInteger(results.length)} period comparison plot(s) created${skipped.length ? `; ${formatInteger(skipped.length)} parameter(s) had no values` : ""}.`,
    false,
    true
  );
}

function renderPeriodResult() {
  const current = state.lastPeriod;
  if (!current) return;
  const normalized = current.plot === "normalized";
  const cards = current.results.map((result, index) => {
    const rows = periodComparisonRows(result);
    return `<article class="chart-card period-chart-card">
      <h3>${escapeHtml(result.parameter)}</h3>
      ${chartLegend(result.datasets.map((dataset, colorIndex) => [dataset.key, TREND_COLORS[colorIndex]]))}
      <p class="chart-caption">${result.datasets.map((dataset) => `${dataset.key}: ${dataset.label}`).map(escapeHtml).join(" | ")}</p>
      <canvas id="period-chart-${index}" aria-label="${escapeHtml(result.parameter)} period comparison"></canvas>
      <details${current.results.length === 1 ? " open" : ""}>
        <summary>Comparison values</summary>
        <div class="table-wrap mini-table">${renderTable(rows)}</div>
      </details>
    </article>`;
  }).join("");
  byId("period-result").innerHTML = `
    <div class="metric-grid">
      ${metric("Plots", formatInteger(current.results.length))}
      ${metric("Plot type", normalized ? "Normalized" : "Mean +/- SD")}
      ${metric("Data scope", current.scope === "all" ? "All lots" : "Filtered")}
      ${metric("Skipped", formatInteger(current.skipped.length))}
    </div>
    ${current.skipped.length ? `<details class="skipped-list"><summary>Parameters without comparison data</summary><p>${current.skipped.map(escapeHtml).join("<br>")}</p></details>` : ""}
    <div class="period-chart-grid">${cards}</div>
  `;
  requestAnimationFrame(drawPeriodCharts);
}

function periodComparisonRows(result) {
  const header = ["Zone"];
  result.datasets.forEach((dataset) => header.push(`${dataset.key} N`, `${dataset.key} Mean`, `${dataset.key} SD`, `${dataset.key} Normalized`));
  const zoneIndexes = result.regional ? [0, 1, 2, 3, 4, 5, 6] : [6];
  return [header, ...zoneIndexes.map((zoneIndex) => {
    const row = [zoneIndex < 6 ? `Zone ${zoneIndex + 1}` : "All"];
    result.datasets.forEach((dataset) => {
      const stats = dataset.stats[zoneIndex];
      row.push(stats.n, stats.mean, stats.sigma, stats.normalized);
    });
    return row;
  })];
}

function drawPeriodCharts() {
  const current = state.lastPeriod;
  if (!current) return;
  current.results.forEach((result, index) => drawPeriodChart(byId(`period-chart-${index}`), result, current.plot));
}

function drawPeriodChart(canvas, result, plotType) {
  if (!canvas) return;
  const zoneIndexes = result.regional ? [0, 1, 2, 3, 4, 5, 6] : [6];
  const normalized = plotType === "normalized";
  const values = result.datasets.flatMap((dataset) => zoneIndexes.flatMap((zoneIndex) => {
    const item = dataset.stats[zoneIndex];
    return normalized ? [item.normalized] : [item.mean - item.sigma, item.mean + item.sigma];
  })).filter(Number.isFinite);
  if (!values.length) return;
  const { ctx, width, height, colors } = setupCanvas(canvas);
  const yExtent = normalized ? paddedExtent([...values, 1]) : integerChartAxis(values);
  const pad = { left: 50, right: 12, top: 14, bottom: 42 };
  ctx.font = "11px Aptos, Calibri, Arial, sans-serif";
  if (!normalized) pad.left = Math.max(pad.left, ...yExtent.ticks.map((value) => ctx.measureText(formatInteger(value)).width + 12));
  const plotWidth = width - pad.left - pad.right;
  const plotHeight = height - pad.top - pad.bottom;
  const yScale = (value) => pad.top + plotHeight - (value - yExtent.min) / (yExtent.max - yExtent.min) * plotHeight;
  ctx.clearRect(0, 0, width, height);
  if (normalized) drawTrendGrid(ctx, pad, width, height, colors);
  else {
    ctx.strokeStyle = colors.line;
    ctx.lineWidth = 1;
    yExtent.ticks.forEach((value) => {
      const y = yScale(value);
      ctx.beginPath();
      ctx.moveTo(pad.left, y);
      ctx.lineTo(width - pad.right, y);
      ctx.stroke();
    });
    drawFrame(ctx, pad, width, height, colors);
  }
  const slot = plotWidth / zoneIndexes.length;
  if (normalized) {
    result.datasets.forEach((dataset, datasetIndex) => {
      ctx.beginPath();
      zoneIndexes.forEach((zoneIndex, index) => {
        const value = dataset.stats[zoneIndex].normalized;
        if (!Number.isFinite(value)) return;
        const x = pad.left + slot * (index + 0.5);
        const y = yScale(value);
        if (index === 0) ctx.moveTo(x, y);
        else ctx.lineTo(x, y);
      });
      ctx.strokeStyle = TREND_COLORS[datasetIndex];
      ctx.lineWidth = 2;
      ctx.stroke();
      zoneIndexes.forEach((zoneIndex, index) => {
        const value = dataset.stats[zoneIndex].normalized;
        if (!Number.isFinite(value)) return;
        ctx.beginPath();
        ctx.arc(pad.left + slot * (index + 0.5), yScale(value), 3, 0, Math.PI * 2);
        ctx.fillStyle = TREND_COLORS[datasetIndex];
        ctx.fill();
      });
    });
  } else {
    const groupWidth = slot * 0.72;
    const barWidth = groupWidth / result.datasets.length;
    result.datasets.forEach((dataset, datasetIndex) => zoneIndexes.forEach((zoneIndex, index) => {
      const item = dataset.stats[zoneIndex];
      if (!Number.isFinite(item.mean)) return;
      const x = pad.left + slot * index + (slot - groupWidth) / 2 + datasetIndex * barWidth;
      const zeroY = yScale(Math.min(yExtent.max, Math.max(0, yExtent.min)));
      const meanY = yScale(item.mean);
      ctx.fillStyle = TREND_COLORS[datasetIndex];
      ctx.globalAlpha = 0.78;
      ctx.fillRect(x + 1, Math.min(zeroY, meanY), Math.max(2, barWidth - 2), Math.abs(zeroY - meanY));
      ctx.globalAlpha = 1;
      if (Number.isFinite(item.sigma)) {
        const centerX = x + barWidth / 2;
        const top = yScale(item.mean + item.sigma);
        const bottom = yScale(item.mean - item.sigma);
        ctx.strokeStyle = TREND_COLORS[datasetIndex];
        ctx.lineWidth = 1.3;
        ctx.beginPath();
        ctx.moveTo(centerX, top);
        ctx.lineTo(centerX, bottom);
        ctx.moveTo(centerX - 3, top);
        ctx.lineTo(centerX + 3, top);
        ctx.moveTo(centerX - 3, bottom);
        ctx.lineTo(centerX + 3, bottom);
        ctx.stroke();
      }
    }));
  }
  ctx.fillStyle = colors.muted;
  ctx.font = "11px Aptos, Calibri, Arial, sans-serif";
  ctx.textAlign = "center";
  zoneIndexes.forEach((zoneIndex, index) => ctx.fillText(zoneIndex < 6 ? `Z${zoneIndex + 1}` : "All", pad.left + slot * (index + 0.5), height - 15));
  if (normalized) {
    ctx.textAlign = "left";
    ctx.fillText(formatNumber(yExtent.max, 3), 4, pad.top + 8);
    ctx.fillText(formatNumber(yExtent.min, 3), 4, height - pad.bottom);
  } else {
    ctx.textAlign = "right";
    yExtent.ticks.forEach((value) => ctx.fillText(formatInteger(value), pad.left - 7, yScale(value) + 4));
  }
}

function savePeriodPlot() {
  const current = state.lastPeriod;
  const canvases = current ? current.results.map((_, index) => byId(`period-chart-${index}`)).filter(Boolean) : [];
  if (!canvases.length) throw new Error("Create a period comparison before saving the plot.");
  let canvas = canvases[0];
  if (canvases.length > 1) {
    const columns = 2;
    const cellWidth = 600;
    const cellHeight = 390;
    canvas = document.createElement("canvas");
    canvas.width = cellWidth * columns;
    canvas.height = cellHeight * Math.ceil(canvases.length / columns);
    const ctx = canvas.getContext("2d");
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    canvases.forEach((source, index) => {
      const x = index % columns * cellWidth;
      const y = Math.floor(index / columns) * cellHeight;
      ctx.fillStyle = "#162433";
      ctx.font = "700 18px Aptos, Calibri, Arial, sans-serif";
      ctx.fillText(current.results[index].parameter, x + 12, y + 24);
      ctx.drawImage(source, x + 8, y + 32, cellWidth - 16, cellHeight - 40);
    });
  }
  const parameter = current.results.length === 1 ? safeFilePart(current.results[0].parameter) : "All_Parameters";
  const link = document.createElement("a");
  link.href = canvas.toDataURL("image/png");
  link.download = `${baseFileName()}_Period_Comparison_${parameter}_${fileDateStamp(new Date())}.png`;
  document.body.append(link);
  link.click();
  link.remove();
}

function createReleaseSummary() {
  const selectedLot = byId("release-lot").value;
  if (!selectedLot) throw new Error("Select the Lot to release.");
  const useFiltered = byId("release-reference").value === "filtered";
  const result = buildLotReleaseSummary(
    state.headers,
    dataRows(),
    useFiltered ? filteredRows() : dataRows(),
    selectedLot,
    requiredNumber("release-monitor"),
    requiredNumber("release-not-ok")
  );
  state.lastRelease = { result, reference: useFiltered ? "Filtered rows" : "All historical rows" };
  renderReleaseResult();
  byId("download-release").disabled = false;
  setStatus(`Lot ${selectedLot} release status: ${result.overall}.`, result.overall === "NOT OK", result.overall === "OK");
}

function renderReleaseResult() {
  const current = state.lastRelease;
  if (!current) return;
  const result = current.result;
  byId("release-result").innerHTML = `
    <div class="metric-grid">
      ${metricHtml("Overall", statusBadge(result.overall))}
      ${metric("Not OK", formatInteger(result.counts.notOk))}
      ${metric("Monitor", formatInteger(result.counts.monitor))}
      ${metric("No history", formatInteger(result.counts.noHistory))}
    </div>
    <p class="result-note">Reference: ${escapeHtml(current.reference)}; selected Lot excluded.</p>
    <div class="table-wrap release-table">${renderReleaseTable(result.results)}</div>
  `;
}

function renderReleaseTable(results) {
  const headers = ["Parameter", "Status", "Lot N", "Ref N", "Lot Mean", "Ref Mean", "Delta", "Mean Z", "Lot CV", "Ref CV", "CV Ratio", "Max Zone Bias", "Monitor Points", "Not OK Points"];
  return `<table><thead><tr>${headers.map((header) => `<th>${escapeHtml(header)}</th>`).join("")}</tr></thead><tbody>${results.map((item) => {
    const values = [item.parameter, item.status, item.lotN, item.referenceN, item.lotMean, item.referenceMean, item.meanDelta, item.meanZ, item.lotCv, item.referenceCv, item.cvRatio, item.maxZoneBiasDelta, item.monitorPoints, item.notOkPoints];
    return `<tr>${values.map((value, index) => index === 1 ? `<td class="${statusClass(value)}">${escapeHtml(value)}</td>` : `<td>${formatCell(value)}</td>`).join("")}</tr>`;
  }).join("")}</tbody></table>`;
}

function downloadReleaseSummary() {
  const current = state.lastRelease;
  if (!current) throw new Error("Create a release summary before downloading it.");
  const result = current.result;
  const rows = [
    ["Lot Release Summary", result.selectedLot],
    ["Overall status", result.overall],
    ["Reference", current.reference],
    ["Monitor limit (SD)", result.monitorLimit],
    ["Not OK limit (SD)", result.notOkLimit],
    [],
    ["Parameter", "Status", "Lot N", "Ref N", "Lot Mean", "Ref Mean", "Mean Delta", "Mean Z", "Lot CV", "Ref CV", "CV Ratio", "Max Zone Bias Delta", "Monitor Points", "Not OK Points"],
    ...result.results.map((item) => [item.parameter, item.status, item.lotN, item.referenceN, item.lotMean, item.referenceMean, item.meanDelta, item.meanZ, item.lotCv, item.referenceCv, item.cvRatio, item.maxZoneBiasDelta, item.monitorPoints, item.notOkPoints])
  ];
  downloadText(`${baseFileName()}_Release_${safeFilePart(result.selectedLot)}.csv`, toCsv(rows), "text/csv;charset=utf-8");
}

function showZmPlan() {
  const current = state.lastAssessment;
  if (!current) throw new Error("Run a New Lot Assessment before showing a ZM plan.");
  const parameter = current.parameter;
  const columns = zoneColumns(state.headers, parameter);
  if (parameter === "All parameters" || !columns?.every((column) => column >= 0)) {
    throw new Error("Select and assess one complete Zone 1-6 parameter before showing a ZM plan.");
  }
  const layoutName = byId("zm-layout").value;
  const specification = getZmPlanSpecification(layoutName);
  const lotColumn = headerIndex(state.headers, "Lot");
  const batchColumn = headerIndex(state.headers, "N");
  const lotRows = dataRows().filter((row) => sameDataValue(row[lotColumn], current.lot));
  const totalRolls = specification.zoneRollCounts.reduce((sum, count) => sum + count, 0);
  const assessmentRows = new Map(current.assessment.grid.map((row) => [planBatchIndex(row.batch), row]));
  const batchValues = new Map();
  for (const row of lotRows) {
    const batch = planBatchIndex(row[batchColumn]);
    if (batch < 1 || batch > 50 || batchValues.has(batch)) continue;
    const assessmentRow = assessmentRows.get(batch);
    batchValues.set(batch, ZONES.map((zone) => {
      const value = row[columns[zone - 1]];
      const assessmentColumn = current.assessment.columns.findIndex((column) => column.header === `${parameter}_${zone}`);
      return {
        value: isNumeric(value) ? number(value) : null,
        status: assessmentColumn >= 0 ? assessmentRow?.states?.[assessmentColumn] || "NO HISTORY" : "NO HISTORY",
        signedScore: assessmentColumn >= 0 ? assessmentRow?.signedScores?.[assessmentColumn] ?? null : null
      };
    }));
  }
  if (!batchValues.size) throw new Error(`No Batch N values from Lot ${current.lot} match M1 through M50.`);

  const coordinateResult = parseZmCoordinates(byId("zm-coordinates").value, totalRolls);
  const bulkLabel = text(byId("zm-mark-label").value) || "X";
  state.lastZmPlan = {
    layoutName,
    specification,
    totalRolls,
    lot: current.lot,
    parameter,
    batchValues,
    labels: new Map([...coordinateResult.marked].map((coordinate) => [coordinate, bulkLabel])),
    invalidCoordinates: coordinateResult.invalid,
    monitorLimit: current.assessment.monitorLimit,
    outlierLimit: current.assessment.outlierLimit,
    viewZone: 1
  };
  byId("zm-label-measurement").disabled = false;
  byId("zm-label-roll").max = String(totalRolls);
  byId("zm-label-roll").disabled = false;
  byId("apply-zm-label").disabled = false;
  byId("remove-zm-label").disabled = false;
  byId("save-zm-plan-png").disabled = false;
  renderZmPlan();
  setStatus(
    `${layoutName}: ${batchValues.size} Batch row(s) shown, ${coordinateResult.marked.size} M/R area(s) labeled${coordinateResult.invalid.length ? `; invalid: ${coordinateResult.invalid.join(", ")}` : ""}.`,
    Boolean(coordinateResult.invalid.length),
    !coordinateResult.invalid.length
  );
}

function parseZmCoordinates(input, totalRolls) {
  const marked = new Set();
  const invalid = [];
  const tokens = text(input).split(",").map((item) => item.trim()).filter(Boolean);
  for (const token of tokens) {
    const match = token.match(/^M?(\d+)\s*\/\s*(\d+)(?:\s*-\s*(\d+))?$/i);
    const measurement = match ? Number(match[1]) : 0;
    const rollStart = match ? Number(match[2]) : 0;
    const rollEnd = match && match[3] ? Number(match[3]) : rollStart;
    if (!match || measurement < 1 || measurement > 50 || rollStart < 1 || rollEnd < rollStart || rollEnd > totalRolls) {
      invalid.push(token);
      continue;
    }
    for (let roll = rollStart; roll <= rollEnd; roll += 1) {
      marked.add(`${measurement}/${roll}`);
    }
  }
  return { marked, invalid };
}

function renderZmPlan() {
  const plan = state.lastZmPlan;
  if (!plan) return;
  const { specification } = plan;
  const viewZone = Number.isInteger(plan.viewZone) && plan.viewZone >= 0 && plan.viewZone <= 6 ? plan.viewZone : 1;
  const allZones = viewZone === 0;
  const zoneStarts = specification.zoneRollCounts.map((_, zoneIndex) => (
    specification.zoneRollCounts.slice(0, zoneIndex).reduce((sum, count) => sum + count, 0) + 1
  ));
  const visibleZoneIndexes = allZones ? ZONES.map((zone) => zone - 1) : [viewZone - 1];
  const visibleRolls = visibleZoneIndexes.flatMap((zoneIndex) => Array.from(
    { length: specification.zoneRollCounts[zoneIndex] },
    (_, rollIndex) => zoneStarts[zoneIndex] + rollIndex
  ));
  const segments = [];
  for (let zoneIndex = 0; zoneIndex < 6; zoneIndex += specification.zonesPerSegment) {
    const counts = specification.zoneRollCounts.slice(zoneIndex, zoneIndex + specification.zonesPerSegment);
    segments.push(counts.reduce((sum, count) => sum + count, 0));
  }
  const tableWidth = 70 + visibleRolls.length * 28;
  const measurementRows = Array.from({ length: 50 }, (_, index) => {
    const measurement = index + 1;
    const zones = plan.batchValues.get(measurement) || Array.from({ length: 6 }, () => ({ value: null, status: "NO HISTORY" }));
    return `<tr class="zm-coordinate-row">
      <th rowspan="2">M${measurement}</th>
      ${visibleRolls.map((roll) => {
        const coordinate = `${measurement}/${roll}`;
        const label = plan.labels.get(coordinate) || "";
        return `<td class="${label ? "is-marked" : ""}"><button type="button" class="zm-mark-button" data-measurement="${measurement}" data-roll="${roll}" aria-label="M${measurement} / R${roll}">${escapeHtml(label)}</button></td>`;
      }).join("")}
    </tr>
    <tr class="zm-value-row">
      ${visibleZoneIndexes.map((zoneIndex) => {
        const item = zones[zoneIndex];
        const style = Number.isFinite(item.signedScore)
          ? directionalAssessmentStyle(item.signedScore, plan.monitorLimit, plan.outlierLimit)
          : "";
        return `<td colspan="${specification.zoneRollCounts[zoneIndex]}" class="${statusClass(item.status)}"${style ? ` style="${style}"` : ""}>${formatCell(item.value)}</td>`;
      }).join("")}
    </tr>`;
  }).join("");
  const zoneHeaderRows = allZones ? `
    <tr><th></th>${segments.map((count) => `<th colspan="${count}">${formatInteger(specification.segmentWidth)} mm</th>`).join("")}</tr>
    <tr><th></th>${specification.zoneRollCounts.map((count, zoneIndex) => `<th colspan="${count}" class="zm-zone-${zoneIndex + 1}">ZONE ${zoneIndex + 1}</th>`).join("")}</tr>` : `
    <tr><th></th><th colspan="${visibleRolls.length}" class="zm-plan-width">Rolls ${visibleRolls[0]}-${visibleRolls[visibleRolls.length - 1]}</th></tr>
    <tr><th></th><th colspan="${visibleRolls.length}" class="zm-zone-${viewZone}">ZONE ${viewZone}</th></tr>`;
  byId("zm-plan-result").innerHTML = `
    <div class="section-heading"><h2>ZM Plan</h2></div>
    <div class="metric-grid">
      ${metric("Layout", plan.layoutName)}
      ${metric("Lot", plan.lot)}
      ${metric("Parameter", plan.parameter)}
      ${metric("Batches shown", formatInteger(plan.batchValues.size))}
      ${metric("M/R labels", formatInteger(plan.labels.size))}
    </div>
    ${plan.invalidCoordinates.length ? `<p class="mapping-result is-error">Invalid coordinates: ${plan.invalidCoordinates.map(escapeHtml).join(", ")}</p>` : ""}
    <div class="zm-view-tabs" role="tablist" aria-label="ZM plan Zone view">
      ${ZONES.map((zone) => `<button type="button" class="zm-view-tab ${viewZone === zone ? "is-active" : ""}" data-zm-view-zone="${zone}" role="tab" aria-selected="${viewZone === zone}">Zone ${zone}</button>`).join("")}
      <button type="button" class="zm-view-tab ${allZones ? "is-active" : ""}" data-zm-view-zone="0" role="tab" aria-selected="${allZones}">All Zones</button>
    </div>
    <div class="table-wrap zm-plan-wrap ${allZones ? "" : "is-zone-view"}">
      <table class="zm-plan-table ${allZones ? "is-all-zones" : "is-zone-view"}" style="${allZones ? `min-width:${tableWidth}px` : "min-width:100%;width:100%"}">
        <colgroup><col class="zm-m-col">${visibleRolls.map(() => '<col class="zm-roll-col">').join("")}</colgroup>
        <thead>
          <tr><th></th><th colspan="${visibleRolls.length}" class="zm-plan-title">${escapeHtml(plan.layoutName)} - Lot ${escapeHtml(plan.lot)} - ${escapeHtml(plan.parameter)}</th></tr>
          ${allZones ? `<tr><th></th><th colspan="${visibleRolls.length}" class="zm-plan-width">${formatInteger(specification.totalWidth)} mm</th></tr>` : ""}
          ${zoneHeaderRows}
          <tr><th>M</th>${visibleRolls.map((roll) => `<th>${roll}</th>`).join("")}</tr>
        </thead>
        <tbody>${measurementRows}</tbody>
      </table>
    </div>
  `;
  byId("zm-plan-result").querySelectorAll(".zm-mark-button").forEach((button) => {
    button.addEventListener("click", handleZmPlanLabelClick);
  });
  byId("zm-plan-result").querySelectorAll(".zm-view-tab").forEach((button) => {
    button.addEventListener("click", handleZmViewClick);
  });
}

function handleZmViewClick(event) {
  const plan = state.lastZmPlan;
  const zone = Number(event.currentTarget.dataset.zmViewZone);
  if (!plan || !Number.isInteger(zone) || zone < 0 || zone > 6) return;
  plan.viewZone = zone;
  renderZmPlan();
  setStatus(zone ? `ZM plan fitted to Zone ${zone}.` : "ZM plan showing all Zones with horizontal scrolling.", false, true);
}

async function saveFullZmPlanPng() {
  const plan = state.lastZmPlan;
  if (!plan) throw new Error("Show the ZM plan before saving it as PNG.");
  setStatus("Rendering the complete ZM plan PNG...");
  const canvas = renderFullZmPlanCanvas(plan);
  const blob = await new Promise((resolve) => canvas.toBlob(resolve, "image/png"));
  if (!blob) throw new Error("The browser could not create the ZM plan PNG.");
  const fileName = `${baseFileName()}_ZM_${safeFilePart(plan.layoutName)}_${safeFilePart(plan.lot)}_${safeFilePart(plan.parameter)}_${fileDateStamp(new Date())}.png`;
  downloadBlob(fileName, blob);
  setStatus(`Full ZM plan PNG saved (${formatInteger(canvas.width)} x ${formatInteger(canvas.height)} px) with ${formatInteger(plan.labels.size)} M/R label(s).`, false, true);
}

function renderFullZmPlanCanvas(plan) {
  const rollWidth = 30;
  const measurementWidth = 72;
  const rowHeight = 25;
  const headerRows = 5;
  const width = measurementWidth + plan.totalRolls * rollWidth + 1;
  const height = headerRows * rowHeight + 50 * rowHeight * 2 + 1;
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const ctx = canvas.getContext("2d");
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, width, height);

  const zoneColors = ["#f4ccb8", "#ffe5cc", "#ccecf7", "#ddebf7", "#e2efda", "#ebf1de"];
  const { specification } = plan;
  let y = 0;
  drawZmCanvasCell(ctx, 0, y, width, rowHeight, `${plan.layoutName} | Lot ${plan.lot} | ${plan.parameter}`, "#123f61", "#ffffff", true, 13);
  y += rowHeight;
  drawZmCanvasCell(ctx, 0, y, measurementWidth, rowHeight, "", "#edf4f8");
  drawZmCanvasCell(ctx, measurementWidth, y, plan.totalRolls * rollWidth, rowHeight, `${formatInteger(specification.totalWidth)} mm`, "#ffffff", "#172231", true);
  y += rowHeight;

  let zoneOffset = 0;
  for (let zoneIndex = 0; zoneIndex < 6; zoneIndex += specification.zonesPerSegment) {
    const rollCount = specification.zoneRollCounts
      .slice(zoneIndex, zoneIndex + specification.zonesPerSegment)
      .reduce((sum, count) => sum + count, 0);
    drawZmCanvasCell(ctx, measurementWidth + zoneOffset * rollWidth, y, rollCount * rollWidth, rowHeight, `${formatInteger(specification.segmentWidth)} mm`, "#edf4f8", "#172231", true);
    zoneOffset += rollCount;
  }
  drawZmCanvasCell(ctx, 0, y, measurementWidth, rowHeight, "", "#edf4f8");
  y += rowHeight;

  let rollOffset = 0;
  specification.zoneRollCounts.forEach((rollCount, zoneIndex) => {
    drawZmCanvasCell(ctx, measurementWidth + rollOffset * rollWidth, y, rollCount * rollWidth, rowHeight, `ZONE ${zoneIndex + 1}`, zoneColors[zoneIndex], "#172231", true);
    rollOffset += rollCount;
  });
  drawZmCanvasCell(ctx, 0, y, measurementWidth, rowHeight, "", "#edf4f8");
  y += rowHeight;

  drawZmCanvasCell(ctx, 0, y, measurementWidth, rowHeight, "M", "#edf4f8", "#123f61", true);
  for (let roll = 1; roll <= plan.totalRolls; roll += 1) {
    drawZmCanvasCell(ctx, measurementWidth + (roll - 1) * rollWidth, y, rollWidth, rowHeight, roll, "#edf4f8", "#123f61", true, 9);
  }
  y += rowHeight;

  for (let measurement = 1; measurement <= 50; measurement += 1) {
    const zones = plan.batchValues.get(measurement) || Array.from({ length: 6 }, () => ({ value: null, status: "NO HISTORY", signedScore: null }));
    drawZmCanvasCell(ctx, 0, y, measurementWidth, rowHeight * 2, `M${measurement}`, "#edf4f8", "#123f61", true);
    for (let roll = 1; roll <= plan.totalRolls; roll += 1) {
      const label = plan.labels.get(`${measurement}/${roll}`) || "";
      drawZmCanvasCell(
        ctx,
        measurementWidth + (roll - 1) * rollWidth,
        y,
        rollWidth,
        rowHeight,
        label,
        label ? "#fff36b" : "#ffffff",
        label ? "#b3261e" : "#172231",
        Boolean(label),
        8
      );
    }
    let valueOffset = 0;
    zones.forEach((item, zoneIndex) => {
      const rollCount = specification.zoneRollCounts[zoneIndex];
      const colors = zmPlanValueColors(item, plan);
      drawZmCanvasCell(
        ctx,
        measurementWidth + valueOffset * rollWidth,
        y + rowHeight,
        rollCount * rollWidth,
        rowHeight,
        Number.isFinite(item.value) ? formatNumber(item.value, 4) : "",
        colors.background,
        colors.text,
        true
      );
      valueOffset += rollCount;
    });
    y += rowHeight * 2;
  }
  return canvas;
}

function drawZmCanvasCell(ctx, x, y, width, height, value, background = "#ffffff", color = "#172231", bold = false, fontSize = 10) {
  ctx.fillStyle = background;
  ctx.fillRect(x, y, width, height);
  ctx.strokeStyle = "#bac8d2";
  ctx.lineWidth = 1;
  ctx.strokeRect(x + 0.5, y + 0.5, width, height);
  const rendered = String(value ?? "");
  if (!rendered) return;
  let size = fontSize;
  do {
    ctx.font = `${bold ? "700 " : ""}${size}px Aptos, Calibri, Arial, sans-serif`;
    if (ctx.measureText(rendered).width <= width - 4 || size <= 6) break;
    size -= 1;
  } while (size >= 6);
  ctx.save();
  ctx.beginPath();
  ctx.rect(x + 1, y + 1, Math.max(0, width - 2), Math.max(0, height - 2));
  ctx.clip();
  ctx.fillStyle = color;
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.fillText(rendered, x + width / 2, y + height / 2);
  ctx.restore();
}

function zmPlanValueColors(item, plan) {
  if (Number.isFinite(item.signedScore)) {
    return directionalAssessmentColors(item.signedScore, plan.monitorLimit, plan.outlierLimit);
  }
  if (item.status === "OK" || item.status === "IN RANGE") return { background: "#d9ead3", text: "#172231" };
  if (item.status === "CHECK" || item.status === "MONITOR") return { background: "#fff2cc", text: "#172231" };
  if (item.status === "OUT OF RANGE" || item.status === "NOT OK" || item.status === "INVESTIGATE") return { background: "#b3261e", text: "#ffffff" };
  return { background: "#f2f2f2", text: "#7f7f7f" };
}

function handleZmPlanLabelClick(event) {
  const button = event.target.closest(".zm-mark-button");
  const plan = state.lastZmPlan;
  if (!button || !plan) return;
  const coordinate = `${button.dataset.measurement}/${button.dataset.roll}`;
  if (plan.labels.has(coordinate)) {
    plan.labels.delete(coordinate);
  } else {
    const label = text(byId("zm-mark-label").value);
    if (!label) {
      setStatus("Enter an M/R label before tapping the plan.", true);
      byId("zm-mark-label").focus();
      return;
    }
    plan.labels.set(coordinate, label);
  }
  renderZmPlan();
  setStatus(`M${button.dataset.measurement} / R${button.dataset.roll} updated. ${plan.labels.size} M/R area(s) labeled.`, false, true);
}

function updateZmPlanLabel(remove) {
  const plan = state.lastZmPlan;
  if (!plan) throw new Error("Show the ZM plan before labeling an M/R area.");
  const measurement = Number(byId("zm-label-measurement").value);
  const roll = Number(byId("zm-label-roll").value);
  if (!Number.isInteger(measurement) || measurement < 1 || measurement > 50) {
    throw new Error("Measurement M must be a whole number from 1 to 50.");
  }
  if (!Number.isInteger(roll) || roll < 1 || roll > plan.totalRolls) {
    throw new Error(`Roll R must be a whole number from 1 to ${plan.totalRolls}.`);
  }
  const coordinate = `${measurement}/${roll}`;
  if (remove) {
    plan.labels.delete(coordinate);
  } else {
    const label = text(byId("zm-mark-label").value);
    if (!label) throw new Error("Enter an M/R label before applying it.");
    plan.labels.set(coordinate, label);
  }
  renderZmPlan();
  setStatus(`M${measurement} / R${roll} ${remove ? "label removed" : "labeled"}. ${plan.labels.size} M/R area(s) labeled.`, false, true);
}

function planBatchIndex(value) {
  const rendered = text(value).replace(/^M\s*/i, "");
  const numeric = Number(rendered);
  return Number.isFinite(numeric) && numeric >= 1 ? Math.round(numeric) : 0;
}

function createAssessment() {
  const parameter = byId("assessment-parameter").value;
  const lot = byId("assessment-lot").value;
  if (!lot) throw new Error("No Lot is available.");
  const mode = byId("assessment-reference").value;
  const assessment = buildV90LotAssessment(
    state.headers,
    dataRows(),
    mode === "filtered" ? filteredRows() : dataRows(),
    {
      parameter,
      selectedLot: lot,
      referenceMode: mode,
      referenceGranularity: byId("assessment-granularity").value,
      referenceLot: byId("assessment-reference-lot").value,
      referenceLots: availableEqualReferenceLots().filter((lot) => state.equalReferenceLots.has(assessmentLotKey(lot))),
      manualMu: optionalNumber("assessment-mu"),
      manualSigma: optionalNumber("assessment-sigma"),
      monitorLimit: requiredNumber("assessment-monitor"),
      outlierLimit: requiredNumber("assessment-outlier")
    }
  );
  state.lastAssessment = {
    parameter,
    lot,
    mode,
    granularity: assessment.referenceGranularity,
    assessment,
    referenceViewZone: 1
  };
  renderAssessmentResult();
  byId("show-zm-plan").disabled = parameter === "All parameters" || !zoneColumns(state.headers, parameter)?.every((column) => column >= 0);
  setStatus(`Lot assessment: ${assessment.overall}.`, assessment.overall === "NOT OK", assessment.overall === "OK");
}

function renderAssessmentResult() {
  const result = state.lastAssessment;
  if (!result) return;
  const { assessment } = result;
  const equalLotMode = assessment.referenceMode === "equal-lots";
  const batchZoneMode = assessment.referenceGranularity === "batch-zone";
  const appliedReferenceTable = batchZoneMode
    ? renderAppliedReferenceViewer(assessment, result.referenceViewZone)
    : "";
  byId("assessment-result").innerHTML = `
    <div class="metric-grid">
      ${metricHtml("Overall", statusBadge(assessment.overall))}
      ${metric("Matching", batchZoneMode ? "Batch + Zone" : "Zone only")}
      ${metric("Compared", formatInteger(assessment.comparedCount))}
      ${metric("Monitor", formatInteger(assessment.monitorCount))}
      ${metric("Out of range", formatInteger(assessment.outlierCount))}
      ${metric(equalLotMode ? "Below 2 ref. lots" : "No history", formatInteger(assessment.noHistoryCount))}
      ${metric(equalLotMode ? "Selected ref. lots" : "History excluded", formatInteger(equalLotMode ? assessment.referenceLots.length : assessment.historyExcluded))}
    </div>
    ${equalLotMode ? `<p class="result-note">Equal weight; complete selected lots; no MAD trimming. Sigma includes within-lot and between-lot variation.</p>
      <p class="result-note">Reference lots: ${assessment.referenceLots.map(escapeHtml).join(", ")}</p>` : ""}
    <div class="section-heading"><h2>Whole-lot Parameter Summary</h2></div>
    <div class="table-wrap compact-table">${renderV90AssessmentSummary(assessment.summaries)}</div>
    <div class="section-heading"><h2>Batch x Parameter x Zone</h2></div>
    <label>Batch N<input id="assessment-batch-filter" type="search" autocomplete="off" placeholder="All batches" value="${escapeHtml(state.assessmentBatchQuery)}" aria-describedby="assessment-batch-status"></label>
    <p id="assessment-batch-status" class="result-note" role="status"></p>
    <p class="result-note">Higher-than-Mu values shade toward red; lower-than-Mu values shade toward green.</p>
    <div id="assessment-batch-table" class="table-wrap"></div>
    ${appliedReferenceTable}
    ${equalLotMode && !batchZoneMode ? `<details class="foldable-section" open><summary>Applied Equal Lot References</summary>
      <div class="table-wrap" data-fold-managed="true">${renderTable([
        ["Parameter / Zone", "Lots", "Values", "Mu", "Sigma"],
        ...assessment.columns.map((column) => [column.header, column.reference.lotCount, column.reference.valueCount, column.reference.mean, column.reference.sigma])
      ])}</div></details>` : ""}
  `;
  byId("assessment-result").querySelectorAll(".assessment-reference-tab").forEach((button) => {
    button.addEventListener("click", handleAssessmentReferenceView);
  });
  renderAssessmentBatchTables();
}

function renderAssessmentBatchTables() {
  const current = state.lastAssessment;
  if (!current || !byId("assessment-batch-table")) return;
  const { assessment } = current;
  let indexes = [];
  let error = "";
  try {
    indexes = assessmentBatchIndexes(assessment.grid, state.assessmentBatchQuery);
  } catch (failure) {
    error = failure.message;
  }
  byId("assessment-batch-filter").setAttribute("aria-invalid", String(Boolean(error)));
  byId("assessment-batch-status").textContent = error || `${indexes.length} of ${assessment.grid.length} batch rows shown`;
  const batchZoneMode = assessment.referenceGranularity === "batch-zone";
  const rows = [
    ["Batch N", ...assessment.columns.map(column => column.header)],
    ["Mu", ...assessment.columns.map(column => batchZoneMode ? "Per Batch" : column.reference.mean)],
    ["Sigma", ...assessment.columns.map(column => batchZoneMode ? "Per Batch" : column.reference.sigma)],
    ...indexes.map(index => [assessment.grid[index].batch, ...assessment.grid[index].values])
  ];
  byId("assessment-batch-table").innerHTML = indexes.length
    ? renderV90AssessmentTable(rows, assessment, indexes)
    : '<p class="empty-state">No matching batches.</p>';
  const referenceTable = byId("assessment-result").querySelector(".applied-reference-wrap");
  if (referenceTable) {
    const batches = new Set(indexes.map(index => text(assessment.grid[index].batch)));
    const references = assessment.appliedReferences.filter(item => batches.has(text(item.batch)));
    referenceTable.innerHTML = references.length
      ? renderAppliedReferenceTable(references, current.referenceViewZone, assessment.referenceMode === "equal-lots")
      : '<p class="empty-state">No matching batches.</p>';
  }
}

function renderAppliedReferenceViewer(assessment, requestedZone) {
  const viewZone = Number.isInteger(requestedZone) && requestedZone >= 0 && requestedZone <= 6 ? requestedZone : 1;
  const allZones = viewZone === 0;
  return `
    <details class="foldable-section applied-reference-module" open>
      <summary>Applied Batch + Zone References</summary>
      <div class="foldable-section-body">
        <div class="zm-view-tabs" role="tablist" aria-label="Applied reference Zone view">
          ${ZONES.map((zone) => `<button type="button" class="zm-view-tab assessment-reference-tab ${viewZone === zone ? "is-active" : ""}" data-reference-view-zone="${zone}" role="tab" aria-selected="${viewZone === zone}">Zone ${zone}</button>`).join("")}
          <button type="button" class="zm-view-tab assessment-reference-tab ${allZones ? "is-active" : ""}" data-reference-view-zone="0" role="tab" aria-selected="${allZones}">All Zones</button>
        </div>
        <div class="table-wrap applied-reference-wrap ${allZones ? "" : "is-zone-view"}" data-fold-managed="true">
          ${renderAppliedReferenceTable(assessment.appliedReferences, viewZone, assessment.referenceMode === "equal-lots")}
        </div>
      </div>
    </details>`;
}

function renderAppliedReferenceTable(references, viewZone, equalLotMode = false) {
  if (viewZone > 0) {
    const zoneReferences = references.filter((item) => item.zone === viewZone);
    return `<table class="applied-reference-table is-zone-view">
      <thead>
        <tr><th colspan="5" class="zm-zone-${viewZone}">ZONE ${viewZone}</th></tr>
        <tr><th>Batch N</th><th>${equalLotMode ? "Lots" : "N"}</th><th>Mu</th><th>Sigma</th><th>${equalLotMode ? "Values" : "Excluded"}</th></tr>
      </thead>
      <tbody>${zoneReferences.map((item) => `<tr><td>${formatCell(item.batch)}</td><td>${formatCell(item.n)}</td><td>${formatCell(item.mean)}</td><td>${formatCell(item.sigma)}</td><td>${formatCell(equalLotMode ? item.valueCount : item.excluded)}</td></tr>`).join("")}</tbody>
    </table>`;
  }

  const batches = new Map();
  references.forEach((item) => {
    const key = text(item.batch);
    if (!batches.has(key)) batches.set(key, { batch: item.batch, zones: new Map() });
    batches.get(key).zones.set(item.zone, item);
  });
  return `<table class="applied-reference-table is-all-zones">
    <thead>
      <tr><th rowspan="2">Batch N</th>${ZONES.map((zone) => `<th colspan="4" class="zm-zone-${zone}">ZONE ${zone}</th>`).join("")}</tr>
      <tr>${ZONES.map(() => `<th>${equalLotMode ? "Lots" : "N"}</th><th>Mu</th><th>Sigma</th><th>${equalLotMode ? "Values" : "Excluded"}</th>`).join("")}</tr>
    </thead>
    <tbody>${[...batches.values()].map((batch) => `<tr><td>${formatCell(batch.batch)}</td>${ZONES.map((zone) => {
      const item = batch.zones.get(zone);
      return item
        ? `<td>${formatCell(item.n)}</td><td>${formatCell(item.mean)}</td><td>${formatCell(item.sigma)}</td><td>${formatCell(equalLotMode ? item.valueCount : item.excluded)}</td>`
        : "<td></td><td></td><td></td><td></td>";
    }).join("")}</tr>`).join("")}</tbody>
  </table>`;
}

function handleAssessmentReferenceView(event) {
  const result = state.lastAssessment;
  const zone = Number(event.currentTarget.dataset.referenceViewZone);
  if (!result || !Number.isInteger(zone) || zone < 0 || zone > 6) return;
  result.referenceViewZone = zone;
  renderAssessmentResult();
  setStatus(zone
    ? `Applied references fitted to Zone ${zone}.`
    : "Applied references showing all Zones with horizontal scrolling.", false, true);
}

function renderV90AssessmentSummary(summaries) {
  const headers = ["Parameter", "N", "Mean Z", "Z SD", "Monitor %", "Out %", "Signal"];
  return `<table><thead><tr>${headers.map((header) => `<th>${escapeHtml(header)}</th>`).join("")}</tr></thead><tbody>${summaries.map((item) => `<tr>
    <td>${escapeHtml(item.parameter)}</td><td>${formatCell(item.n)}</td><td>${formatCell(item.meanZ)}</td><td>${formatCell(item.zSigma)}</td>
    <td>${escapeHtml(formatPercent(item.monitorPct))}</td><td>${escapeHtml(formatPercent(item.outlierPct))}</td>
    <td class="${statusClass(item.signal)}">${escapeHtml(item.signal)}</td>
  </tr>`).join("")}</tbody></table>`;
}

function renderV90AssessmentTable(rows, assessment, gridIndexes) {
  const [headers, ...body] = rows;
  return `<table><thead><tr>${headers.map((header) => `<th>${escapeHtml(header)}</th>`).join("")}</tr></thead><tbody>${body.map((row, rowIndex) => {
    const gridIndex = rowIndex >= 2 ? gridIndexes[rowIndex - 2] : null;
    const gridRow = gridIndex !== null ? assessment.grid[gridIndex] : null;
    const states = gridRow?.states;
    return `<tr>${headers.map((_, columnIndex) => {
      const status = columnIndex > 0 ? states?.[columnIndex - 1] : null;
      const signedScore = columnIndex > 0 ? gridRow?.signedScores?.[columnIndex - 1] : null;
      const style = Number.isFinite(signedScore)
        ? directionalAssessmentStyle(signedScore, assessment.monitorLimit, assessment.outlierLimit)
        : "";
      const direction = Number.isFinite(signedScore)
        ? `${signedScore > 0 ? "Above" : signedScore < 0 ? "Below" : "At"} Mu: ${formatNumber(signedScore, 2)} SD`
        : "";
      const inspectable = assessment.referenceMode === "equal-lots" && gridRow && columnIndex > 0;
      const content = inspectable
        ? `<button type="button" class="assessment-value-button" data-row="${gridIndex}" data-column="${columnIndex - 1}" aria-label="Details for Batch ${escapeHtml(gridRow.batch)}, ${escapeHtml(headers[columnIndex])}">${formatCell(row[columnIndex]) || "-"}</button>`
        : formatCell(row[columnIndex]);
      return `<td${status ? ` class="${statusClass(status)}"` : ""}${style ? ` style="${style}"` : ""}${direction ? ` title="${escapeHtml(direction)}"` : ""}>${content}</td>`;
    }).join("")}</tr>`;
  }).join("")}</tbody></table>`;
}

function showEqualLotDetails(rowIndex, columnIndex) {
  const assessment = state.lastAssessment?.assessment;
  if (assessment?.referenceMode !== "equal-lots") return;
  const row = assessment.grid[rowIndex];
  const column = assessment.columns[columnIndex];
  const reference = row?.references[columnIndex];
  if (!row || !column || !reference) return;
  byId("equal-lot-title").textContent = `${column.header} - Batch ${text(row.batch)}`;
  byId("equal-lot-details").innerHTML = `
    <p class="result-note">Target Lot ${escapeHtml(assessment.selectedLot)} | ${assessment.referenceGranularity === "batch-zone" ? "Matching Batch + Zone" : "All batches within Zone"}</p>
    <div class="metric-grid">
      ${metric("Target value", formatCell(row.values[columnIndex]))}
      ${metric("Reference Mu", formatCell(reference.mean))}
      ${metric("Reference Sigma", formatCell(reference.sigma))}
      ${metric("Contributing lots", `${reference.lotCount} / ${assessment.referenceLots.length}`)}
    </div>
    <p class="result-note">${reference.lotCount < 2 ? "No history: fewer than two contributing lots." : escapeHtml(row.states[columnIndex])}</p>
    <table class="equal-lot-contributors">
      <thead><tr><th>Lot</th><th>Values</th><th>Mean</th><th title="Population standard deviation">SD</th><th>Weight</th></tr></thead>
      <tbody>${reference.contributors.map((lot) => `<tr${lot.n ? "" : ' class="status-missing"'}>
        <td>${escapeHtml(lot.lot)}</td><td>${formatInteger(lot.n)}</td><td>${formatCell(lot.mean)}</td><td>${formatCell(lot.sigma)}</td><td>${escapeHtml(formatPercent(lot.weight))}</td>
      </tr>`).join("")}</tbody>
    </table>
    <p class="result-note">Lot SD uses population variance. Lots without numeric values have zero weight.</p>
    <dl class="equal-lot-variance"><dt>Within-lot variance</dt><dd>${formatCell(reference.withinVariance)}</dd>
      <dt>Between-lot variance</dt><dd>${formatCell(reference.betweenVariance)}</dd></dl>
  `;
  byId("equal-lot-dialog").showModal();
}

function directionalAssessmentStyle(signedScore, monitorLimit, outlierLimit) {
  const colors = directionalAssessmentColors(signedScore, monitorLimit, outlierLimit);
  return `background-color:${colors.background};color:${colors.text}`;
}

function directionalAssessmentColors(signedScore, monitorLimit, outlierLimit) {
  const score = Math.abs(Number(signedScore));
  if (!Number.isFinite(score)) return { background: "#f2f2f2", text: "#7f7f7f" };
  const monitor = Number(monitorLimit) > 0 ? Number(monitorLimit) : 2;
  const outlier = Number(outlierLimit) > monitor ? Number(outlierLimit) : monitor + 1;
  const intensity = Math.min(1, score <= monitor
    ? 0.5 * score / monitor
    : 0.5 + 0.5 * (score - monitor) / (outlier - monitor));
  const neutral = [255, 255, 204];
  const target = signedScore > 0 ? [192, 0, 0] : signedScore < 0 ? [0, 128, 0] : neutral;
  const color = neutral.map((value, index) => Math.round(value + (target[index] - value) * intensity));
  return {
    background: `rgb(${color.join(",")})`,
    text: score > outlier ? "#ffffff" : "#172231"
  };
}

async function createCorrelation() {
  if (byId("correlation-analysis").value === "structured") {
    invalidateCorrelation();
    setStatus("Calculating structured correlation methods...");
    await yieldToBrowser();
    const sourceRows = rowsForAnalysis("correlation-data-scope");
    const result = buildStructuredCorrelation(state.headers, sourceRows, byId("correlation-y").value, {
      parameters: state.structuredParameters,
      xParameter: byId("correlation-x").value,
      compareAll: byId("correlation-compare").value === "all",
      method: byId("correlation-method").value,
      minN: requiredNumber("correlation-min-n"), coverage: byId("correlation-coverage").value,
      minRegions: requiredNumber("correlation-min-regions"), expectedRegions: byId("correlation-expected-regions").value
    });
    state.lastStructuredCorrelation = { ...result, source: state.source, scope: byId("correlation-data-scope").value,
      plotHeaders: state.headers.slice(), plotRows: sourceRows.map(row => row.slice()), plotParameters: state.structuredParameters.slice(), plot: null };
    byId("correlation-result").innerHTML = `
      <div class="metric-grid">${metric("Results", result.results.length)}${metric("Numeric correlations", result.results.filter(item => item.r !== null).length)}${metric("Source rows", sourceRows.length)}${metric("Minimum N", result.minN)}</div>
      <div class="section-heading"><h2>Structured Correlation</h2></div>
      <div class="table-wrap structured-correlation-table">${renderStructuredCorrelationTable(result)}</div>
      <button type="button" class="command" id="show-structured-conclusion">Show conclusion</button>
      <details id="structured-conclusion" hidden><summary>Correlation conclusion</summary><div id="structured-conclusion-text"></div></details>
      <div id="structured-plot-result"></div>`;
    byId("show-structured-conclusion").addEventListener("click", showStructuredConclusion);
    byId("correlation-result").querySelectorAll("[data-structured-plot]").forEach(button => button.addEventListener("click", () => runAction(() => plotStructuredCorrelation(Number(button.dataset.structuredPlot)))));
    setStatus(`${result.results.length} structured correlation results.`, false, true);
    return;
  }
  const yParameter = byId("correlation-y").value;
  const xParameter = byId("correlation-x").value;
  if (xParameter === yParameter) throw new Error("Select two different parameters.");
  const scope = byId("correlation-data-scope").value;
  const removalMethod = byId("correlation-outliers").value;
  const removalValue = requiredNumber("correlation-removal");
  const result = buildCorrelation(
    state.headers,
    rowsForAnalysis("correlation-data-scope"),
    xParameter,
    yParameter,
    byId("correlation-scope").value,
    removalMethod,
    removalValue
  );
  state.lastCorrelation = { xParameter, yParameter, scope, removalMethod, removalValue, result };
  renderCorrelationResult();
  setStatus(`Correlation: ${formatInteger(result.rawN)} included, ${formatInteger(result.excludedN)} excluded.`, false, true);
}

function structuredCorrelationTable(result) {
  return [["Target (Y)", "Compare (X)", "Method", "Pearson r", "N used", "Observation unit", "Batches", "Lots", "Regions", "Informative Lots", "Informative Lot-Region groups", "Status"],
    ...result.results.map(item => [item.yParameter, item.xParameter, item.method, item.r === null ? "n.a." : item.r, item.n, item.unit, item.batches, item.lots, item.regions, item.informativeLots, item.informativeGroups, item.status])];
}

function renderStructuredCorrelationTable(result) {
  const [headers, ...rows] = structuredCorrelationTable(result).map(row => [row[2], row[3], row[4], row[0], row[1], ...row.slice(5)]);
  return `<table><thead><tr><th>Plot</th>${headers.map(header => `<th>${escapeHtml(header)}</th>`).join("")}</tr></thead>
    <tbody>${rows.map((row, resultIndex) => `<tr><td><button type="button" class="command" data-structured-plot="${resultIndex}" ${(result.results[resultIndex].originalResult || result.results[resultIndex]).r === null ? "disabled" : ""}>Plot</button></td>${row.map((cell, index) => index === row.length - 1
      ? `<td><details><summary>${row[1] === "n.a." ? "Not available" : "Details"}</summary>${escapeHtml(cell)}</details></td>`
      : `<td>${formatCell(cell)}</td>`).join("")}</tr>`).join("")}</tbody></table>`;
}

function plotStructuredCorrelation(index) {
  const current = state.lastStructuredCorrelation, tableRow = current?.results[index], selected = tableRow?.originalResult || tableRow;
  if (!selected || selected.r === null) throw new Error("Select an available correlation result.");
  const result = buildStructuredCorrelation(current.plotHeaders, current.plotRows, selected.yParameter, {
    parameters: current.plotParameters, xParameter: selected.xParameter, method: current.method,
    coverage: current.coverage, minN: current.minN, minRegions: current.minRegions,
    expectedRegions: current.expectedRegions.join(","), plotMethod: selected.method
  }).results.find(item => item.method === selected.method);
  if (!result?.points?.length || result.n !== selected.n) throw new Error("Rebuild structured correlation before plotting.");
  current.plot = result;
  current.originalPlot = result;
  current.plotIndex = index;
  if (tableRow.trimOptions) current.plot = { ...result, ...trimStructuredPlot(result.points, tableRow.trimOptions), trimOptions: tableRow.trimOptions };
  byId("structured-plot-result").innerHTML = `<div class="section-heading"><h2>${escapeHtml(result.method)}</h2><button type="button" class="command" id="save-structured-plot">Save PNG</button></div>
    <p>X: ${escapeHtml(result.xParameter)} | Y: ${escapeHtml(result.yParameter)} | ${escapeHtml(result.unit)}</p>
    <label>Exclusion method<select id="structured-trim-method"><option value="count">Largest / smallest pairs</option><option value="ratio">Y/X ratio versus original slope (%)</option><option value="line">Y versus original fitted line (%)</option></select></label>
    <label id="structured-percent-control" hidden>Tolerance (+/- %)<input id="structured-trim-percent" type="number" min="0" step="any" value="90"></label>
    <div class="control-grid" id="structured-count-controls"><label>Extreme axis<select id="structured-trim-axis"><option value="x">X</option><option value="y">Y</option></select></label>
    <label>Remove<select id="structured-trim-side"><option value="largest">Largest</option><option value="smallest">Smallest</option></select></label>
    <label>Pairs to exclude<input id="structured-trim-count" type="number" min="0" step="1" value="0"></label></div>
    <div class="button-row"><button class="command" id="structured-trim-apply" type="button">Apply exclusion</button><button class="command" id="structured-trim-reset" type="button">Reset</button><button class="command" id="structured-update-table" type="button">Update table</button></div>
    <p id="structured-plot-stats"></p>
    <div class="chart-card"><canvas id="structured-correlation-chart" aria-label="${escapeHtml(result.method)} correlation plot"></canvas></div>`;
  byId("save-structured-plot").addEventListener("click", () => runAction(saveStructuredPlot));
  byId("structured-trim-apply").addEventListener("click", () => runAction(() => applyStructuredTrim(false)));
  byId("structured-trim-reset").addEventListener("click", () => runAction(() => applyStructuredTrim(true)));
  byId("structured-update-table").addEventListener("click", () => runAction(updateStructuredTable));
  byId("structured-trim-method").addEventListener("change", () => {
    const count = byId("structured-trim-method").value === "count";
    byId("structured-count-controls").hidden = !count;
    byId("structured-percent-control").hidden = count;
  });
  if (tableRow.trimOptions) {
    const options = tableRow.trimOptions;
    byId("structured-trim-method").value = options.method;
    byId("structured-trim-percent").value = options.percent;
    byId("structured-trim-axis").value = options.axis;
    byId("structured-trim-side").value = options.side;
    byId("structured-trim-count").value = options.count;
    byId("structured-trim-method").dispatchEvent(new Event("change"));
  }
  requestAnimationFrame(() => { drawStructuredPlot(); byId("structured-plot-result").scrollIntoView({ block: "nearest" }); });
}

function applyStructuredTrim(reset) {
  const current = state.lastStructuredCorrelation;
  if (!current?.originalPlot) return;
  const trimOptions = {
    method: reset ? "count" : byId("structured-trim-method").value,
    percent: reset ? 90 : requiredNumber("structured-trim-percent"),
    axis: byId("structured-trim-axis").value, side: byId("structured-trim-side").value,
    count: reset || byId("structured-trim-method").value !== "count" ? 0 : requiredNumber("structured-trim-count"), minN: current.minN
  };
  const trimmed = trimStructuredPlot(current.originalPlot.points, trimOptions);
  current.plot = { ...current.originalPlot, ...trimmed, trimOptions };
  if (reset) byId("structured-trim-count").value = "0";
  drawStructuredPlot();
  setStatus("Plot statistics updated. The original correlation table and source data are unchanged.", false, true);
}

function updateStructuredTable() {
  const current = state.lastStructuredCorrelation;
  if (!current?.plot) return;
  const original = current.results[current.plotIndex].originalResult || current.results[current.plotIndex];
  const { points, ...plot } = current.plot;
  current.results[current.plotIndex] = { ...original, ...plot, originalResult: original,
    informativeLots: original.informativeLots ? plot.lots : 0,
    status: `${original.status}; Plot exclusions: ${plot.excluded || 0}; ${plot.removal || "None"}; original N=${original.n}, original r=${original.r}` };
  const table = byId("correlation-result").querySelector(".structured-correlation-table");
  table.innerHTML = renderStructuredCorrelationTable(current);
  byId("correlation-result").querySelector(".metric-grid").innerHTML = `${metric("Results", current.results.length)}${metric("Numeric correlations", current.results.filter(item => item.r !== null).length)}${metric("Source rows", current.plotRows.length)}${metric("Minimum N", current.minN)}`;
  table.querySelectorAll("[data-structured-plot]").forEach(button => button.addEventListener("click", () => runAction(() => plotStructuredCorrelation(Number(button.dataset.structuredPlot)))));
  if (!byId("structured-conclusion").hidden) showStructuredConclusion(false);
  setStatus("Selected table result updated. Workbook export includes the updated result and exclusion details. Source data is unchanged.", false, true);
}

function showStructuredConclusion(scroll = true) {
  const current = state.lastStructuredCorrelation;
  if (!current) return;
  byId("structured-conclusion-text").innerHTML = structuredConclusions(current.results).map(item =>
    `<section><h3>${escapeHtml(item.xParameter)} vs ${escapeHtml(item.yParameter)}</h3>${item.lines.map(line => `<p>${escapeHtml(line)}</p>`).join("")}</section>`).join("");
  const section = byId("structured-conclusion");
  section.hidden = false;
  if (scroll) { section.open = true; section.scrollIntoView({ block: "nearest" }); }
}

function drawStructuredPlot() {
  const plot = state.lastStructuredCorrelation?.plot;
  if (!plot) return;
  byId("structured-plot-stats").textContent = `Original N: ${state.lastStructuredCorrelation.originalPlot.n} | Included: ${plot.n} | Excluded: ${plot.excluded || 0} | r: ${plot.r === null ? "n.a." : formatNumber(plot.r, 4)} | Exclusion: ${plot.removal || "None"}`;
  const canvas = byId("structured-correlation-chart");
  const residual = plot.unit.includes("residual") ? " (residual)" : "";
  drawScatter(canvas, { xParameter: plot.xParameter + residual, yParameter: plot.yParameter + residual,
    result: { pairs: plot.points, included: plot.points, slope: plot.slope, intercept: plot.intercept } });
}

function saveStructuredPlot() {
  const plot = state.lastStructuredCorrelation?.plot, source = byId("structured-correlation-chart");
  if (!plot || !source) throw new Error("Plot a structured correlation result first.");
  drawStructuredPlot();
  const canvas = document.createElement("canvas");
  canvas.width = Math.max(700, source.width); canvas.height = source.height + 124;
  const ctx = canvas.getContext("2d");
  ctx.fillStyle = "#ffffff"; ctx.fillRect(0, 0, canvas.width, canvas.height);
  ctx.fillStyle = "#172231"; ctx.font = "16px sans-serif";
  ctx.fillText(plot.method, 12, 24, canvas.width - 24);
  ctx.fillText(`X: ${plot.xParameter} | Y: ${plot.yParameter}`, 12, 48, canvas.width - 24);
  ctx.fillText(`${plot.unit} | N: ${plot.n} | r: ${plot.r === null ? "n.a." : formatNumber(plot.r, 4)}`, 12, 72, canvas.width - 24);
  ctx.fillText(`Excluded: ${plot.excluded || 0} | ${plot.removal || "None"}`, 12, 96, canvas.width - 24);
  ctx.drawImage(source, (canvas.width - source.width) / 2, 124);
  const link = document.createElement("a");
  link.href = canvas.toDataURL("image/png");
  link.download = `${safeFilePart(plot.xParameter)}_${safeFilePart(plot.yParameter)}_${safeFilePart(plot.method)}.png`;
  document.body.append(link); link.click(); link.remove();
}

function renderCorrelationResult() {
  const current = state.lastCorrelation;
  if (!current) return;
  const { result } = current;
  const displayPairs = [
    ...result.pairs.filter((pair) => !pair.included),
    ...result.pairs.filter((pair) => pair.included)
  ];
  byId("correlation-result").innerHTML = `
    <div class="metric-grid">
      ${metric("Included", formatInteger(result.rawN))}
      ${metric("Excluded", formatInteger(result.excludedN))}
      ${metric("Pearson r", formatNumber(result.r, 4))}
      ${metric("Slope", formatNumber(result.slope, 4))}
    </div>
    <div class="chart-card"><canvas id="correlation-chart" aria-label="Matched-zone correlation scatter plot"></canvas></div>
    <div class="table-wrap mini-table">${renderTable([
      ["Lot", "Batch N", "Zone", "X", "Y", "Status"],
      ...displayPairs.slice(0, 80).map((pair) => [pair.lot, pair.batch, `Zone ${pair.zone}`, pair.x, pair.y, pair.included ? "Included" : `Excluded: ${pair.exclusionReason || "Extreme"}`])
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
  if (state.parameters.length) {
    appendSheet(workbook, "Origin_Long_App", buildOriginExport(state.headers, filteredRows(), parameter, zones));
    appendSheet(workbook, "Origin_Wide_App", buildOriginExportWide(state.headers, filteredRows(), parameter, zones));
  }
  if (state.lastStructuredCorrelation) {
    appendSheet(workbook, "Correlation_Structured", structuredCorrelationTable(state.lastStructuredCorrelation));
    appendSheet(workbook, "Correlation_Settings", [["Source", state.lastStructuredCorrelation.source], ["Data scope", state.lastStructuredCorrelation.scope],
      ["Method", state.lastStructuredCorrelation.method], ["Coverage", state.lastStructuredCorrelation.coverage], ["Minimum N", state.lastStructuredCorrelation.minN],
      ["Minimum paired Regions", state.lastStructuredCorrelation.minRegions], ["Expected Regions", state.lastStructuredCorrelation.expectedRegions.join(", ")]]);
  }
  if (state.lastGaussian) {
    const { parameter: gaussianParameter, zones: gaussianZones, fit } = state.lastGaussian;
    appendSheet(workbook, "GaussianFit_App", [
      ["Parameter", gaussianParameter],
      ["Fit method", fit.method],
      ["Huber iterations", fit.iterations || 0],
      ["Data scope", state.lastGaussian.scope],
      ["Excel rows", state.lastGaussian.visibleExcelRowsOnly ? "Visible only" : "All rows in selected lots"],
      ["Zones", gaussianZones.map((zone) => `Zone ${zone}`).join(", ")],
      ["Fit N", fit.n],
      ["Visible N", fit.visibleN],
      ["Points excluded", fit.excluded],
      ["Excluded by range", fit.excludedByRange],
      ["Outside histogram", fit.outsideHistogram],
      ["Histogram N", fit.histogramN],
      ["Recorded below limit", fit.belowLimit],
      ["Lots used", fit.lotCount],
      ["Batches used", fit.batchCount],
      ["Bin width", fit.binWidth],
      ["Histogram start", fit.start],
      ["Histogram end", fit.end],
      ["Mean", fit.mean],
      ["Sigma", fit.sigma],
      ["2.5% low cutoff", fit.low25],
      ["97.5% cutoff", fit.high25],
      ["15% low cutoff", fit.low15],
      ["85% cutoff", fit.high15],
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
  for (const exported of state.gaussianExtremeExports) appendSheet(workbook, exported.name, exported.table);
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
  if (state.lastPeriod) {
    const rows = [
      ["Period / Lot Comparison"],
      ["Data scope", state.lastPeriod.scope],
      ["Mode", state.lastPeriod.mode],
      ["Plot", state.lastPeriod.plot],
      []
    ];
    state.lastPeriod.results.forEach((result) => {
      rows.push([result.parameter], ...periodComparisonRows(result), []);
    });
    appendSheet(workbook, "Period_Comparison_App", rows);
  }
  if (state.lastCorrelation) {
    const { xParameter, yParameter, result } = state.lastCorrelation;
    appendSheet(workbook, "Correlation_App", [
      ["X parameter", xParameter],
      ["Y parameter", yParameter],
      ["Data scope", state.lastCorrelation.scope],
      ["Pair removal", state.lastCorrelation.removalMethod],
      ["Removal value", state.lastCorrelation.removalValue],
      ["Included pairs", result.rawN],
      ["Excluded pairs", result.excludedN],
      ["Pearson r", result.r],
      ["Slope", result.slope],
      ["Intercept", result.intercept],
      [],
      ["Lot", "Batch N", "Zone", "X value", "Y value", "Status"],
      ...result.pairs.map((pair) => [pair.lot, pair.batch, `Zone ${pair.zone}`, pair.x, pair.y, pair.included ? "Included" : `Excluded: ${pair.exclusionReason || "Extreme"}`])
    ]);
  }
  if (state.lastAssessment) {
    const { parameter: assessmentParameter, lot, mode, assessment } = state.lastAssessment;
    const batchZoneMode = assessment.referenceGranularity === "batch-zone";
    const equalLotMode = mode === "equal-lots";
    const assessmentRows = [
      ["Parameter", assessmentParameter],
      ["Lot", lot],
      ["Reference", mode],
      ...(equalLotMode ? [["Reference Lots", ...assessment.referenceLots], ["Sigma", "Equal-Lot population SD (within + between)"]] : []),
      ["Historical matching", batchZoneMode ? "Batch + Zone" : "Zone only"],
      ["Overall", assessment.overall],
      ["Compared", assessment.comparedCount],
      ["Monitor", assessment.monitorCount],
      ["Out of range", assessment.outlierCount],
      ["No history", assessment.noHistoryCount],
      ["History excluded", assessment.historyExcluded],
      [],
      ["Parameter", "N", "Mean Z", "Z SD", "Monitor %", "Out %", "Signal"],
      ...assessment.summaries.map((item) => [item.parameter, item.n, item.meanZ, item.zSigma, item.monitorPct, item.outlierPct, item.signal]),
      [],
      ["Batch N", ...assessment.columns.map((column) => column.header)],
      ["Mu", ...assessment.columns.map((column) => batchZoneMode ? "Per Batch" : column.reference.mean)],
      ["Sigma", ...assessment.columns.map((column) => batchZoneMode ? "Per Batch" : column.reference.sigma)],
      ...assessment.grid.map((row) => [row.batch, ...row.values])
    ];
    if (batchZoneMode) assessmentRows.push(
      [],
      ["Applied Batch + Zone References"],
      ["Batch N", "Parameter", "Zone", equalLotMode ? "Lots" : "N", "Mu", "Sigma", equalLotMode ? "Values" : "Excluded"],
      ...assessment.appliedReferences.map((item) => [item.batch, item.parameter, item.zone, item.n, item.mean, item.sigma, equalLotMode ? item.valueCount : item.excluded])
    );
    if (equalLotMode && !batchZoneMode) assessmentRows.push(
      [], ["Applied Equal Lot References"], ["Parameter / Zone", "Lots", "Values", "Mu", "Sigma"],
      ...assessment.columns.map((column) => [column.header, column.reference.lotCount, column.reference.valueCount, column.reference.mean, column.reference.sigma])
    );
    appendSheet(workbook, "NewLot_Assessment_App", assessmentRows);
  }
  if (state.lastRelease) {
    const release = state.lastRelease.result;
    appendSheet(workbook, "Lot_Release_App", [
      ["Lot Release Summary", release.selectedLot],
      ["Overall status", release.overall],
      ["Reference", state.lastRelease.reference],
      ["Monitor limit (SD)", release.monitorLimit],
      ["Not OK limit (SD)", release.notOkLimit],
      [],
      ["Parameter", "Status", "Lot N", "Ref N", "Lot Mean", "Ref Mean", "Mean Delta", "Mean Z", "Lot CV", "Ref CV", "CV Ratio", "Max Zone Bias Delta", "Monitor Points", "Not OK Points"],
      ...release.results.map((item) => [item.parameter, item.status, item.lotN, item.referenceN, item.lotMean, item.referenceMean, item.meanDelta, item.meanZ, item.lotCv, item.referenceCv, item.cvRatio, item.maxZoneBiasDelta, item.monitorPoints, item.notOkPoints])
    ]);
  }
  writeDataLabelsToWorkbook(workbook, state.headers, exportRows);
  XLSX.writeFile(workbook, `${baseFileName()}_analysis.xlsx`, { cellStyles: true });
}

function appendSheet(workbook, name, rows) {
  const XLSX = getXlsx();
  const sheet = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(workbook, sheet, safeSheetName(name));
}

function dataRows() {
  return state.rows.filter((row) => isDataRow(state.headers, row));
}

function selectedRows() {
  const filterMode = byId("filter-mode").value || "all";
  const typeValue = byId("type-filter").value || ALL;
  const lotColumn = headerIndex(state.headers, "Lot");
  const typeColumn = headerIndex(state.headers, "Type");
  const classificationColumn = headerIndex(state.headers, "Classification");
  const keyword = byId("classification-keyword").value;
  return dataRows().filter((row) => {
    if (typeValue !== ALL && typeColumn >= 0 && text(row[typeColumn]) !== typeValue) return false;
    if (filterMode === "lots" && lotColumn >= 0 && !state.filterSelections.lots.has(text(row[lotColumn]))) return false;
    if (filterMode === "classification-keyword" && !classificationIncludesKeyword(row[classificationColumn], keyword)) return false;
    if (filterMode === "classification") {
      const value = classificationColumn >= 0 && text(row[classificationColumn])
        ? text(row[classificationColumn])
        : UNCLASSIFIED;
      if (!state.filterSelections.classification.has(value)) return false;
    }
    return true;
  });
}

function syncLotPeriodControls() {
  const enabled = byId("limit-lot-period").checked;
  byId("lot-period-fields").hidden = !enabled;
  byId("lot-period-start").disabled = !enabled;
  byId("lot-period-end").disabled = !enabled;
}

async function updateLotPeriodFilter() {
  syncLotPeriodControls();
  invalidateAnalyses();
  if (state.periodFilterLoading) {
    renderCurrentData();
    return;
  }
  state.periodFilterError = "";
  if (byId("limit-lot-period").checked && !(state.trendDateLookup instanceof Map)) {
    state.periodFilterLoading = true;
    renderCurrentData();
    try {
      await loadTrendDateLookup();
    } catch (error) {
      state.periodFilterError = error.message || "Production dates could not be loaded.";
    } finally {
      state.periodFilterLoading = false;
    }
  }
  renderCurrentData();
  const result = lotPeriodSelection();
  setStatus(result.error || "Lot period filter updated.", Boolean(result.error), !result.error);
}

function lotPeriodSelection() {
  const rows = selectedRows();
  if (!byId("limit-lot-period").checked) return { rows, undatedLotCount: 0 };
  try {
    if (state.periodFilterLoading) throw new Error("Reading production dates...");
    if (state.periodFilterError) throw new Error(state.periodFilterError);
    return filterLotsByPeriod(state.headers, rows, state.trendDateLookup || new Map(), {
      startDate: byId("lot-period-start").value,
      endDate: byId("lot-period-end").value
    });
  } catch (error) {
    return { rows: [], undatedLotCount: 0, error: error.message };
  }
}

function filteredRows() {
  return lotPeriodSelection().rows;
}

function rowsForAnalysis(scopeId) {
  return byId(scopeId).value === "all" ? dataRows() : filteredRows();
}

function syncAssessmentLots() {
  if (!state.headers.length) return;
  const lots = state.lots;
  const previous = byId("assessment-lot").value;
  fillSelect(byId("assessment-lot"), lots, lots.includes(previous) ? previous : lots[0]);
  byId("assessment-lot").disabled = !lots.length;
  byId("run-assessment").disabled = !lots.length;
  syncAssessmentReferenceLots();
}

function syncAssessmentReferenceLots() {
  const selectedLot = byId("assessment-lot").value;
  const lots = state.lots.filter((lot) => text(lot).toUpperCase() !== text(selectedLot).toUpperCase());
  const previous = byId("assessment-reference-lot").value;
  fillSelect(byId("assessment-reference-lot"), lots, lots.includes(previous) ? previous : lots[0]);
  byId("assessment-reference-lot").disabled = !lots.length || byId("assessment-reference").value !== "reference-lot";
  const available = new Set(availableEqualReferenceLots().map(assessmentLotKey));
  state.equalReferenceLots = new Set([...state.equalReferenceLots].filter((lot) => available.has(lot)));
  renderEqualReferenceChoices();
}

function assessmentLotKey(value) {
  const rendered = text(value);
  return rendered && Number.isFinite(Number(rendered)) ? String(Number(rendered)) : rendered.toUpperCase();
}

function availableEqualReferenceLots() {
  const target = assessmentLotKey(byId("assessment-lot").value);
  return [...new Map(state.lots.filter((lot) => assessmentLotKey(lot) !== target)
    .map((lot) => [assessmentLotKey(lot), lot])).values()];
}

function updateEqualReferenceCount() {
  byId("assessment-equal-count").textContent = `${state.equalReferenceLots.size} selected (minimum 2)`;
}

function renderEqualReferenceChoices() {
  const search = text(byId("assessment-equal-search").value).toUpperCase();
  const lots = availableEqualReferenceLots().filter((lot) => text(lot).toUpperCase().includes(search));
  byId("assessment-equal-options").innerHTML = lots.map((lot) => {
    const key = assessmentLotKey(lot);
    return `<label><input type="checkbox" value="${escapeHtml(key)}" ${state.equalReferenceLots.has(key) ? "checked" : ""}><span>${escapeHtml(lot)}</span></label>`;
  }).join("");
  updateEqualReferenceCount();
}

function syncAssessmentReferenceMode() {
  const mode = byId("assessment-reference").value;
  byId("assessment-equal-lots-field").hidden = mode !== "equal-lots";
  byId("assessment-reference-lot-field").hidden = mode !== "reference-lot";
  byId("assessment-reference-lot").disabled = mode !== "reference-lot" || !byId("assessment-reference-lot").options.length;
  const manual = mode === "manual";
  byId("assessment-mu").disabled = !manual;
  byId("assessment-sigma").disabled = !manual;
  if (manual && byId("assessment-parameter").value === "All parameters") {
    byId("assessment-parameter").value = state.v90Parameters[0] || "";
  }
  const parameter = byId("assessment-parameter").value;
  const completeRegional = parameter !== "All parameters" && zoneColumns(state.headers, parameter)?.every((column) => column >= 0);
  const batchZoneAvailable = !manual && completeRegional;
  if (!batchZoneAvailable) byId("assessment-granularity").value = "zone";
  byId("assessment-granularity").disabled = !batchZoneAvailable;
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

function syncStructuredCorrelationControls(refreshParameters = true) {
  const structured = byId("correlation-analysis").value === "structured";
  byId("correlation-structured-settings").hidden = !structured;
  ["correlation-scope", "correlation-outliers", "correlation-removal"].forEach(id => { byId(id).closest("label").hidden = structured; });
  if (refreshParameters) {
    const choices = structured ? state.structuredParameters : state.parameters;
    fillSelect(byId("correlation-y"), choices, byId("correlation-y").value || choices[0]);
    fillSelect(byId("correlation-x"), choices, byId("correlation-x").value || choices[1]);
    if (byId("correlation-x").value === byId("correlation-y").value && choices.length > 1) byId("correlation-x").value = choices.find(value => value !== byId("correlation-y").value);
  }
  byId("correlation-x").disabled = structured && byId("correlation-compare").value === "all";
  ["correlation-coverage", "correlation-min-regions"].forEach(id => { byId(id).disabled = !structured; });
  byId("correlation-expected-regions").disabled = !structured || byId("correlation-coverage").value !== "strict";
  byId("run-correlation").disabled = (structured ? state.structuredParameters : state.parameters).length < 2;
  syncCorrelationZones();
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

function syncCorrelationRemovalInput() {
  const method = byId("correlation-outliers").value;
  const input = byId("correlation-removal");
  const label = byId("correlation-removal-label");
  const previousKind = input.dataset.removalKind;
  if (previousKind === "percent") input.dataset.percentValue = input.value;
  if (previousKind === "count") input.dataset.countValue = input.value;
  const kind = method === "ratio" ? "percent" : method === "keep" ? "none" : "count";
  if (kind === "percent") {
    label.textContent = "Maximum removal %";
    input.min = "0";
    input.max = "50";
    input.step = "0.1";
    input.inputMode = "decimal";
    input.value = input.dataset.percentValue || "0.5";
  } else {
    label.textContent = kind === "count" ? "Points to remove (N)" : "Points to remove";
    input.min = "0";
    input.removeAttribute("max");
    input.step = "1";
    input.inputMode = "numeric";
    if (kind === "count") input.value = input.dataset.countValue || "1";
  }
  input.disabled = kind === "none";
  input.dataset.removalKind = kind;
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

function drawGaussian(canvas, fit, mode = "combined") {
  if (!canvas || !fit?.bins?.length) return;
  const { ctx, width, height, colors } = setupCanvas(canvas);
  const pad = { left: 42, right: 12, top: 100, bottom: 34 };
  const plotWidth = width - pad.left - pad.right;
  const plotHeight = height - pad.top - pad.bottom;
  const maxY = Math.max(1, ...fit.bins.flatMap((bin) => [bin.observed, bin.gaussian]));
  ctx.clearRect(0, 0, width, height);
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, width, height);
  drawFrame(ctx, pad, width, height, colors);
  if (mode !== "curve") fit.bins.forEach((bin) => {
    const x0 = pad.left + (bin.lower - fit.start) / (fit.end - fit.start) * plotWidth;
    const x1 = pad.left + (bin.upper - fit.start) / (fit.end - fit.start) * plotWidth;
    const barHeight = bin.observed / maxY * plotHeight;
    ctx.fillStyle = colors.blueSoft;
    ctx.fillRect(x0 + 1, pad.top + plotHeight - barHeight, Math.max(1, x1 - x0 - 2), barHeight);
  });
  if (mode !== "histogram") {
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
  }
  const cutoffs = [
    { label: "2.5%", value: fit.low25, color: "#9f1239" },
    { label: "15%", value: fit.low15, color: "#92400e" },
    { label: "85%", value: fit.high15, color: "#0f766e" },
    { label: "97.5%", value: fit.high25, color: "#6d28d9" }
  ];
  ctx.save();
  ctx.font = "11px sans-serif";
  ctx.textAlign = "left";
  ctx.fillStyle = "#172231";
  ctx.fillText(`Mean: ${fit.mean.toFixed(2)}   Sigma: ${fit.sigma.toFixed(2)}`, 8, 16, width - 16);
  const labelLanes = [[], [], [], []];
  const cutoffLabels = [];
  cutoffs.forEach((cutoff) => {
    if (!Number.isFinite(cutoff.value) || cutoff.value < fit.start || cutoff.value > fit.end) return;
    const x = pad.left + (cutoff.value - fit.start) / (fit.end - fit.start) * plotWidth;
    const label = `${cutoff.label}: ${formatNumber(cutoff.value, 2)}`;
    const labelWidth = Math.min(ctx.measureText(label).width, width - 16);
    const labelX = Math.max(8, Math.min(width - 8 - labelWidth, x - labelWidth / 2));
    // Keep close or identical percentile labels separate without moving their cutoff lines.
    const lane = labelLanes.findIndex(items => items.every(item => labelX > item.end + 8 || labelX + labelWidth + 8 < item.start));
    labelLanes[lane].push({ start: labelX, end: labelX + labelWidth });
    const labelY = 34 + lane * 18;
    cutoffLabels.push({ label, labelX, labelY, labelWidth, color: cutoff.color });
    ctx.beginPath();
    ctx.setLineDash([4, 3]);
    ctx.strokeStyle = cutoff.color;
    ctx.lineWidth = 1.5;
    ctx.moveTo(x, labelY + 5);
    ctx.lineTo(x, height - pad.bottom);
    ctx.stroke();
  });
  cutoffLabels.forEach(({ label, labelX, labelY, labelWidth, color }) => {
    ctx.fillStyle = "#ffffff";
    ctx.fillRect(labelX - 2, labelY - 12, labelWidth + 4, 16);
    ctx.fillStyle = color;
    ctx.fillText(label, labelX, labelY, labelWidth);
  });
  ctx.restore();
  drawAxisLabels(ctx, pad, width, height, colors, fit.start, fit.end, maxY);
}

function drawScatter(canvas, current) {
  if (!canvas || !current?.result?.pairs?.length) return;
  const { ctx, width, height, colors } = setupCanvas(canvas);
  const pad = { left: 54, right: 24, top: 16, bottom: 58 };
  const valuesX = current.result.included.map((pair) => pair.x);
  const valuesY = current.result.included.map((pair) => pair.y);
  const boundsX = paddedExtent(valuesX), boundsY = paddedExtent(valuesY);
  const extentX = integerChartAxis([boundsX.min, boundsX.max]);
  const extentY = integerChartAxis([boundsY.min, boundsY.max]);
  ctx.font = "11px Aptos, Calibri, Arial, sans-serif";
  pad.left = Math.max(pad.left, ...extentY.ticks.map(value => ctx.measureText(formatInteger(value)).width + 30));
  pad.right = Math.max(pad.right, ctx.measureText(formatInteger(extentX.max)).width / 2 + 8);
  const plotWidth = width - pad.left - pad.right;
  const plotHeight = height - pad.top - pad.bottom;
  const xScale = (value) => pad.left + (value - extentX.min) / (extentX.max - extentX.min) * plotWidth;
  const yScale = (value) => pad.top + plotHeight - (value - extentY.min) / (extentY.max - extentY.min) * plotHeight;
  ctx.clearRect(0, 0, width, height);
  ctx.fillStyle = "#ffffff";
  ctx.fillRect(0, 0, width, height);
  ctx.strokeStyle = colors.line;
  ctx.lineWidth = 1;
  ctx.fillStyle = colors.muted;
  ctx.textAlign = "right";
  for (const value of extentY.ticks) {
    const y = yScale(value);
    ctx.beginPath(); ctx.moveTo(pad.left, y); ctx.lineTo(width - pad.right, y); ctx.stroke();
    ctx.fillText(formatInteger(value), pad.left - 7, y + 4);
  }
  ctx.textAlign = "center";
  const labelWidth = Math.max(...extentX.ticks.map(value => ctx.measureText(formatInteger(value)).width)) + 10;
  const stride = Math.max(1, Math.ceil(labelWidth / (plotWidth / (extentX.ticks.length - 1))));
  extentX.ticks.forEach((value, index) => {
    const x = xScale(value);
    ctx.beginPath(); ctx.moveTo(x, pad.top); ctx.lineTo(x, height - pad.bottom); ctx.stroke();
    if (index % stride === 0) ctx.fillText(formatInteger(value), x, height - pad.bottom + 18);
  });
  drawFrame(ctx, pad, width, height, colors);
  ctx.save();
  ctx.beginPath(); ctx.rect(pad.left, pad.top, plotWidth, plotHeight); ctx.clip();
  for (const pair of current.result.included) {
    ctx.beginPath();
    ctx.arc(xScale(pair.x), yScale(pair.y), 3, 0, Math.PI * 2);
    ctx.fillStyle = colors.blue;
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
  ctx.restore();
  ctx.fillStyle = colors.ink;
  ctx.font = "12px Aptos, Calibri, Arial, sans-serif";
  ctx.textAlign = "center";
  ctx.fillText(current.xParameter || "X", pad.left + plotWidth / 2, height - 8, plotWidth);
  ctx.save();
  ctx.translate(13, pad.top + plotHeight / 2); ctx.rotate(-Math.PI / 2);
  ctx.fillText(current.yParameter || "Y", 0, 0, plotHeight);
  ctx.restore();
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
  ].filter((_, index) => byId("trend-result").querySelector(`[data-trend-series="${["batch", "rolling", "center", "limits", "limits"][index]}"]`)?.checked));

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
    })).filter((_, index) => byId("trend-result").querySelector(`[data-trend-zone="${ZONES[index]}"]`)?.checked));
    drawTrendBiasChart(byId("trend-bias-chart"), trend.zoneBias);
  }
}

function drawTrendLineChart(canvas, series) {
  if (!canvas) return;
  const finitePoints = series.flatMap((item) => item.data).filter((point) => Number.isFinite(point.x) && Number.isFinite(point.y));
  const { ctx, width, height, colors } = setupCanvas(canvas);
  if (!finitePoints.length) {
    ctx.clearRect(0, 0, width, height);
    ctx.fillStyle = colors.muted;
    ctx.font = "12px sans-serif";
    ctx.fillText("No series selected or no values available.", 12, 24, width - 24);
    return;
  }
  const pad = { left: 48, right: 14, top: 14, bottom: 38 };
  const yExtent = integerChartAxis(finitePoints.map((point) => point.y));
  ctx.font = "11px Aptos, Calibri, Arial, sans-serif";
  pad.left = Math.max(pad.left, ...yExtent.ticks.map(value => ctx.measureText(formatInteger(value)).width + 12));
  const plotWidth = width - pad.left - pad.right;
  const plotHeight = height - pad.top - pad.bottom;
  const xExtent = trendExtent(finitePoints.map((point) => point.x), 86400000);
  const xScale = (value) => pad.left + (value - xExtent.min) / (xExtent.max - xExtent.min) * plotWidth;
  const yScale = (value) => pad.top + plotHeight - (value - yExtent.min) / (yExtent.max - yExtent.min) * plotHeight;
  ctx.clearRect(0, 0, width, height);
  drawTrendGrid(ctx, pad, width, height, colors, yExtent);
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
  const absoluteMax = Math.max(1e-9, ...values.flatMap((item) => [Math.abs(item.meanBias || 0), Math.abs(item.ciLow || 0), Math.abs(item.ciHigh || 0)]));
  const yExtent = integerChartAxis([-absoluteMax * 1.2, absoluteMax * 1.2]);
  ctx.font = "11px Aptos, Calibri, Arial, sans-serif";
  pad.left = Math.max(pad.left, ...yExtent.ticks.map(value => ctx.measureText(formatInteger(value)).width + 12));
  const plotWidth = width - pad.left - pad.right;
  const plotHeight = height - pad.top - pad.bottom;
  const minY = yExtent.min;
  const maxY = yExtent.max;
  const yScale = (value) => pad.top + plotHeight - (value - minY) / (maxY - minY) * plotHeight;
  ctx.clearRect(0, 0, width, height);
  drawTrendGrid(ctx, pad, width, height, colors, yExtent);
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
}

function drawTrendGrid(ctx, pad, width, height, colors, yExtent) {
  ctx.strokeStyle = colors.line;
  ctx.lineWidth = 1;
  ctx.font = "11px Aptos, Calibri, Arial, sans-serif";
  ctx.fillStyle = colors.muted;
  ctx.textAlign = "right";
  for (const value of yExtent.ticks) {
    const y = pad.top + (height - pad.top - pad.bottom) * (yExtent.max - value) / (yExtent.max - yExtent.min);
    ctx.beginPath();
    ctx.moveTo(pad.left, y);
    ctx.lineTo(width - pad.right, y);
    ctx.stroke();
    ctx.fillText(formatInteger(value), pad.left - 7, y + 4);
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
  let min = Infinity, max = -Infinity;
  for (const value of values) {
    if (!Number.isFinite(value)) continue;
    min = Math.min(min, value);
    max = Math.max(max, value);
  }
  if (min === max) {
    min -= 1;
    max += 1;
  }
  const pad = (max - min) * 0.06;
  return { min: min - pad, max: max + pad };
}

function redrawCharts() {
  drawStructuredPlot();
  if (state.lastGaussian && byId("gaussian-chart")) drawGaussianCharts();
  if (state.lastTrend && byId("trend-lot-chart")) drawTrendCharts();
  if (state.lastPeriod && byId("period-chart-0")) drawPeriodCharts();
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
  byId("merge-preview").innerHTML = `<p class="empty-state">No new-lot workbook selected.</p>`;
  byId("data-labels-table").innerHTML = `<p class="empty-state">No saved labels.</p>`;
  byId("label-selection-preview").innerHTML = `<p class="empty-state">No workbook loaded.</p>`;
  byId("master-roll-result").hidden = true;
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
  if (status === "OK" || status === "IN RANGE") return "status-ok";
  if (status === "CHECK" || status === "MONITOR") return "status-check";
  if (status === "OUT OF RANGE" || status === "NOT OK" || status === "INVESTIGATE") return "status-bad";
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

function formatPercent(value) {
  if (!Number.isFinite(Number(value))) return "";
  return Number(value).toLocaleString(undefined, { style: "percent", maximumFractionDigits: 1 });
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
  downloadBlob(fileName, new Blob([content], { type }));
}

function downloadBlob(fileName, blob) {
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
    byId("export-gaussian-extremes").disabled = !state.lastGaussian;
    const merge = state.newLotImport?.preview;
    byId("merge-new-lots").disabled = !merge || Boolean(merge.missingHeaders.length || !merge.addedRows);
    updateLabelActionState();
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
