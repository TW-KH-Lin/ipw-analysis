import { readPreference, writePreference } from "./local-preferences.js?v=1";

const resultTitles = new Map([
  ["build-result", "Build result"],
  ["merge-preview", "Merge preview"],
  ["label-selection-preview", "Selected values"],
  ["gaussian-result", "Gaussian result"],
  ["gaussian-extremes", "Extreme cases"],
  ["gaussian-snapshots", "Saved snapshots"],
  ["trend-result", "Trend result"],
  ["period-result", "Comparison result"],
  ["assessment-result", "Assessment result"],
  ["zm-plan-result", "ZM plan result"],
  ["release-result", "Release result"],
  ["correlation-result", "Correlation result"],
  ["export-result", "Export overview"]
]);

const tableTitles = new Map([
  ["summary-table", "Summary table"],
  ["preview-table", "Data preview"],
  ["generated-summary-table", "Generated summary table"],
  ["classification-table", "Lot classifications"],
  ["data-labels-table", "Saved labels"]
]);

const settingsPanels = new Map([
  ["gaussian-panel", "gaussian-result"],
  ["trend-panel", "trend-result"],
  ["period-panel", "period-result"],
  ["lot-panel", "assessment-result"],
  ["release-panel", "release-result"],
  ["correlation-panel", "correlation-result"],
  ["export-panel", "export-result"]
]);

document.addEventListener("DOMContentLoaded", () => {
  buildSettingsDrawers();
  buildSectionDrawers();
  buildStandaloneDrawers();
  wrapResultBlocks();
  enhanceTables(document);
  rememberFolds(document);
  new MutationObserver((records) => {
    for (const record of records) {
      record.addedNodes.forEach((node) => {
        if (node.nodeType === Node.ELEMENT_NODE) {
          enhanceTables(node);
          rememberFolds(node);
        }
      });
      updateResultVisibility(record.target.closest?.(".result-block") || record.target);
    }
  }).observe(document.body, { childList: true, subtree: true });
});

function foldKey(details) {
  const panel = details.closest(".panel")?.id || "page";
  const result = details.closest(".result-block")?.id || "";
  const identity = details.id || [...details.children].find((child) => child.id)?.id ||
    `${details.className}:${details.querySelector(":scope > summary")?.textContent.trim()}`;
  return `fold:${location.pathname}:${panel}:${result}:${identity}`;
}

function rememberFolds(root) {
  const details = [...(root.matches?.("details") ? [root] : []), ...root.querySelectorAll("details")];
  details.forEach((drawer) => {
    if (drawer.dataset.rememberFold) return;
    drawer.dataset.rememberFold = "true";
    const saved = readPreference(foldKey(drawer));
    if (typeof saved === "boolean") {
      drawer.open = saved;
      drawer.dataset.foldRestored = "true";
    }
    drawer.addEventListener("toggle", () => {
      if (drawer.isConnected && !drawer.hidden) writePreference(foldKey(drawer), drawer.open);
    });
  });
}

function buildSettingsDrawers() {
  for (const [panelId, resultId] of settingsPanels) {
    const panel = document.getElementById(panelId);
    const result = document.getElementById(resultId);
    const hasDrawer = panel && [...panel.children].some((child) => child.classList.contains("analysis-drawer"));
    if (!panel || !result || hasDrawer) continue;
    const drawer = createDetails("Settings", "analysis-drawer", true);
    const body = document.createElement("div");
    body.className = "analysis-drawer-body";
    while (panel.firstChild && panel.firstChild !== result) body.append(panel.firstChild);
    drawer.append(body);
    panel.insertBefore(drawer, result);
  }
}

function buildStandaloneDrawers() {
  const metrics = document.getElementById("data-metrics");
  if (metrics && !metrics.parentElement?.classList.contains("foldable-section-body")) {
    const drawer = createDetails("Workbook overview", "foldable-section", true);
    const body = document.createElement("div");
    body.className = "foldable-section-body";
    metrics.parentElement.insertBefore(drawer, metrics);
    body.append(metrics);
    drawer.append(body);
  }

  const labelsPanel = document.getElementById("labels-panel");
  const savedLabels = labelsPanel && [...labelsPanel.children].find((child) => child.classList.contains("foldable-section"));
  if (!labelsPanel || !savedLabels) return;
  const leadingNodes = [];
  for (const node of [...labelsPanel.childNodes]) {
    if (node === savedLabels) break;
    leadingNodes.push(node);
  }
  if (!leadingNodes.some((node) => node.nodeType === Node.ELEMENT_NODE)) return;
  const drawer = createDetails("Label editor", "foldable-section", true);
  const body = document.createElement("div");
  body.className = "foldable-section-body";
  labelsPanel.insertBefore(drawer, leadingNodes[0]);
  leadingNodes.forEach((node) => body.append(node));
  drawer.append(body);
}

function buildSectionDrawers() {
  document.querySelectorAll(".panel").forEach((panel) => {
    const headings = [...panel.children].filter((child) => child.classList.contains("section-heading"));
    headings.forEach((heading) => {
      if (heading.parentElement !== panel) return;
      const title = heading.querySelector("h2")?.textContent?.trim() || "Section";
      const drawer = createDetails(title, "foldable-section", true);
      const body = document.createElement("div");
      body.className = "foldable-section-body";
      panel.insertBefore(drawer, heading);
      let current = heading;
      while (current && !(current !== heading && current.classList?.contains("section-heading"))) {
        const next = current.nextSibling;
        body.append(current);
        current = next;
      }
      drawer.append(body);
    });
  });
}

function wrapResultBlocks() {
  for (const [id, title] of resultTitles) {
    const result = document.getElementById(id);
    if (!result || result.parentElement?.classList.contains("foldable-result")) continue;
    const drawer = createDetails(title, "foldable-result", true);
    result.parentElement.insertBefore(drawer, result);
    drawer.append(result);
    updateResultVisibility(result);
  }
}

function enhanceTables(root) {
  const tables = [];
  if (root.matches?.(".table-wrap")) tables.push(root);
  root.querySelectorAll?.(".table-wrap").forEach((table) => tables.push(table));
  tables.forEach((table) => {
    if (table.dataset.foldManaged === "true") return;
    if (table.parentElement?.classList.contains("foldable-table")) return;
    const heading = table.previousElementSibling?.classList.contains("section-heading")
      ? table.previousElementSibling.querySelector("h2")?.textContent?.trim()
      : "";
    const title = tableTitles.get(table.id) || heading || tableTitleFromContext(table);
    const drawer = createDetails(title, "foldable-table", true);
    table.parentElement.insertBefore(drawer, table);
    drawer.append(table);
  });
}

function tableTitleFromContext(table) {
  if (table.classList.contains("zm-plan-wrap")) return "ZM plan table";
  if (table.classList.contains("release-table")) return "Release table";
  if (table.classList.contains("snapshot-table")) return "Saved snapshots table";
  if (table.classList.contains("bias-table")) return "Zone bias table";
  const result = table.closest(".result-block");
  if (result?.id === "assessment-result") {
    return table.classList.contains("compact-table") ? "Parameter summary table" : "MR by parameter and Zone";
  }
  if (result?.id === "gaussian-result") return "Histogram values";
  if (result?.id === "correlation-result") return "Correlation pairs";
  if (result?.id === "period-result") return "Comparison values";
  return "Table";
}

function createDetails(title, className, open) {
  const details = document.createElement("details");
  details.className = className;
  details.open = open;
  const summary = document.createElement("summary");
  summary.textContent = title;
  details.append(summary);
  details.addEventListener("toggle", () => {
    if (details.open) requestAnimationFrame(() => window.dispatchEvent(new Event("resize")));
  });
  return details;
}

function updateResultVisibility(result) {
  if (!result?.classList?.contains("result-block")) return;
  const drawer = result.parentElement;
  if (!drawer?.classList.contains("foldable-result")) return;
  const visible = Boolean(result.children.length || result.textContent.trim());
  const wasHidden = drawer.hidden;
  drawer.hidden = !visible;
  if (visible && wasHidden) drawer.open = readPreference(foldKey(drawer), true) !== false;
}
