const screenNames = new Map([
  ["data-panel", "Data overview"],
  ["clean-panel", "Clean data"],
  ["merge-panel", "Merge new lots"],
  ["labels-panel", "Data labels"],
  ["gaussian-panel", "Gaussian fit"],
  ["trend-panel", "Parameter trend"],
  ["period-panel", "Period comparison"],
  ["lot-panel", "New lot assessment"],
  ["release-panel", "Lot release"],
  ["correlation-panel", "Correlation"],
  ["export-panel", "Export"]
]);

const workflows = new Map([
  ["prepare", ["data-panel", "clean-panel", "merge-panel", "labels-panel"]],
  ["analyze", ["gaussian-panel", "trend-panel", "period-panel", "correlation-panel"]],
  ["quality", ["lot-panel", "release-panel"]],
  ["output", ["export-panel"]]
]);

document.addEventListener("DOMContentLoaded", () => {
  bindWorkspaceNavigation();
  bindDatasetDrawer();
});

function bindWorkspaceNavigation() {
  const title = document.getElementById("workspace-screen-title");
  const screenTabs = document.getElementById("workspace-screen-tabs");
  const workflowTabs = [...document.querySelectorAll(".workflow-tab")];
  if (!title || !screenTabs || !workflowTabs.length) return;

  workflowTabs.forEach((tab) => {
    tab.addEventListener("click", () => {
      const firstPanel = workflows.get(tab.dataset.workflow)?.[0];
      document.querySelector(`.tab[data-panel="${firstPanel}"]`)?.click();
      syncWorkspaceNavigation(title, screenTabs, workflowTabs);
    });
  });

  screenTabs.addEventListener("click", (event) => {
    const tab = event.target.closest(".screen-tab");
    if (!tab) return;
    document.querySelector(`.tab[data-panel="${tab.dataset.panel}"]`)?.click();
    syncWorkspaceNavigation(title, screenTabs, workflowTabs);
    document.querySelector(".workspace-switcher")?.scrollIntoView({ block: "start" });
  });

  const observer = new MutationObserver(() => syncWorkspaceNavigation(title, screenTabs, workflowTabs));
  document.querySelectorAll(".panel").forEach((panel) => {
    observer.observe(panel, { attributes: true, attributeFilter: ["class"] });
  });
  syncWorkspaceNavigation(title, screenTabs, workflowTabs);
}

function syncWorkspaceNavigation(title, screenTabs, workflowTabs) {
  const activePanel = document.querySelector(".panel.is-active");
  if (!activePanel || !screenNames.has(activePanel.id)) return;
  const activeWorkflow = [...workflows].find(([, panels]) => panels.includes(activePanel.id))?.[0] || "prepare";
  workflowTabs.forEach((tab) => {
    const active = tab.dataset.workflow === activeWorkflow;
    tab.classList.toggle("is-active", active);
    tab.setAttribute("aria-selected", String(active));
  });
  const panels = workflows.get(activeWorkflow) || [];
  screenTabs.innerHTML = panels.map((panelId) => {
    const active = panelId === activePanel.id;
    return `<button class="screen-tab ${active ? "is-active" : ""}" type="button" data-panel="${panelId}" role="tab" aria-selected="${active}">${screenNames.get(panelId)}</button>`;
  }).join("");
  title.textContent = screenNames.get(activePanel.id);
  document.title = `${screenNames.get(activePanel.id)} - IPW Workspace`;
}

function bindDatasetDrawer() {
  const drawer = document.getElementById("dataset-drawer");
  const summary = document.getElementById("dataset-summary");
  const source = document.getElementById("source-sheet");
  const status = document.getElementById("status");
  if (!drawer || !summary || !source || !status) return;
  let collapsedAfterLoad = false;

  const update = () => {
    const sourceName = source.selectedOptions[0]?.textContent?.trim();
    if (sourceName) summary.textContent = sourceName;
    if (!collapsedAfterLoad && status.classList.contains("is-success") && /loaded:|generated/i.test(status.textContent)) {
      drawer.open = false;
      collapsedAfterLoad = true;
    }
  };
  source.addEventListener("change", update);
  new MutationObserver(update).observe(status, { attributes: true, childList: true, subtree: true });
  update();
}
