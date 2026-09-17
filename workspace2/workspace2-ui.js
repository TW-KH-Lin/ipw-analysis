document.addEventListener("DOMContentLoaded", () => {
  const shell = document.querySelector(".app-shell");
  const sidebar = document.createElement("nav");
  sidebar.className = "workspace2-sidebar";
  sidebar.setAttribute("aria-label", "Functions");
  const groups = [
    ["Data", [["data", "Data overview"], ["clean", "Clean data"], ["classification", "Lots & classification"], ["merge", "Merge new lots"], ["labels", "Data labels"]]],
    ["Analysis", [["gaussian", "Gaussian fit"], ["trend", "Parameter trend"], ["period", "Period comparison"], ["correlation", "Correlation"]]],
    ["Quality", [["lot", "New lot assessment"], ["release", "Lot release"]]],
    ["Save", [["export", "Save & export"]]]
  ];
  for (const [title, items] of groups) {
    const heading = document.createElement("h3"); heading.textContent = title; sidebar.append(heading);
    for (const [id, label] of items) {
      const button = document.createElement("button"); button.type = "button"; button.textContent = label; button.dataset.destination = `${id}-panel`;
      button.addEventListener("click", () => { document.querySelector(`.tab[data-panel="${id}-panel"]`).click(); window.dispatchEvent(new Event("resize")); });
      sidebar.append(button);
    }
  }
  shell.prepend(sidebar);
  const scope = document.createElement("div"); scope.className = "workspace2-scope"; scope.setAttribute("role", "status");
  document.querySelector(".workspace-switcher").after(scope);
  const sync = () => {
    const active = document.querySelector(".panel.is-active");
    sidebar.querySelectorAll("button").forEach(button => {
      if (button.dataset.destination === active?.id) button.setAttribute("aria-current", "page"); else button.removeAttribute("aria-current");
    });
    const source = document.getElementById("source-sheet");
    const filter = document.getElementById("filter-mode");
    const description = `${source.selectedOptions[0]?.textContent || "No workbook loaded"} · Dataset filter: ${filter.selectedOptions[0]?.textContent || "All data"}`;
    if (scope.textContent !== description) scope.textContent = description;
    for (const panel of document.querySelectorAll(".panel")) {
      if (!panel.querySelector(":scope > .analysis-drawer") || panel.querySelector(".workspace2-view-tabs")) continue;
      const guide = panel.querySelector(".correlation-method-guide");
      if (guide) panel.append(guide);
      const tabs = document.createElement("div"); tabs.className = "workspace2-view-tabs"; tabs.setAttribute("role", "tablist"); tabs.setAttribute("aria-label", "Workspace view");
      const choices = [["settings", "Settings"], ["results", "Results"]];
      if (panel.querySelector(".correlation-method-guide")) choices.push(["guide", "Method guide"]);
      const show = view => {
        panel.dataset.workspaceView = view;
        tabs.querySelectorAll("button").forEach(button => button.setAttribute("aria-selected", String(button.dataset.view === view)));
        if (view === "settings") panel.querySelector(".analysis-drawer").open = true;
        if (view === "guide") panel.querySelector(".correlation-method-guide").open = true;
        if (view === "results") panel.querySelectorAll(":scope > .foldable-result").forEach(drawer => drawer.open = true);
        requestAnimationFrame(() => window.dispatchEvent(new Event("resize")));
      };
      for (const [view, label] of choices) {
        const button = document.createElement("button"); button.type = "button"; button.textContent = label; button.dataset.view = view; button.setAttribute("role", "tab"); button.addEventListener("click", () => show(view)); tabs.append(button);
      }
      panel.prepend(tabs); show("settings");
      const result = panel.querySelector(".result-block");
      if (result) new MutationObserver(() => { if (result.querySelector("canvas, .metric-grid")) show("results"); }).observe(result, { childList: true });
    }
  };
  new MutationObserver(sync).observe(shell, { childList: true, subtree: true, attributes: true, attributeFilter: ["class"] });
  document.getElementById("source-sheet").addEventListener("change", sync);
  document.getElementById("filter-mode").addEventListener("change", sync);
  sync();
});
