import { applyWorkspace2SourceOutliers, resetWorkspace2SourceOutliers, workspace2SourceOutlierState } from '../workspace2-app.js?v=1';
const methods = new Set(['huber', 'percentile', 'iqr']);
let lastChoice = null, applying = false;
function enhancePlot() {
  const method = document.getElementById('structured-trim-method');
  if (!method) return;
  if (method.dataset.sourceOptions) {
    const state = workspace2SourceOutlierState(), undo = document.getElementById('structured-trim-undo');
    if (undo && state.active && !state.hasPlotSteps) undo.disabled = false;
    return;
  }
  method.dataset.sourceOptions = 'true';
  method.append(new Option('Huber robust screening', 'huber'), new Option('Percentile tail removal', 'percentile'), new Option('IQR screening', 'iqr'));
  const percentControl = document.getElementById('structured-percent-control'), percent = document.getElementById('structured-trim-percent');
  const controls = document.createElement('div'); controls.className = 'control-grid'; controls.id = 'source-outlier-controls';
  controls.innerHTML = `<label>Extreme axis<select id="source-outlier-axis"><option value="x">X</option><option value="y">Y</option></select></label><label>Remove<select id="source-outlier-side"><option value="high">High tail</option><option value="low">Low tail</option><option value="both" selected>Both tails</option></select></label><label>Maximum removal %<input id="source-outlier-budget" type="number" min="0" max="20" step="0.1" value="2"></label><p id="source-outlier-note" role="status"></p>`;
  document.getElementById('structured-count-controls').after(controls);
  const note = document.getElementById('source-outlier-note');
  function change() {
    const source = methods.has(method.value); controls.hidden = !source;
    percentControl.hidden = !source && method.value === 'count';
    percentControl.firstChild.textContent = source ? (method.value === 'huber' ? 'Sigma multiplier' : method.value === 'iqr' ? 'IQR multiplier' : 'Percent per tail') : 'Tolerance (+/- %)';
    percent.value = source ? (method.value === 'percentile' ? '1' : '3') : '90';
    note.textContent = source ? 'Applied to original paired measurements; table and plot are recalculated using the same source mask.' : '';
  }
  method.addEventListener('change', change);
  const state = workspace2SourceOutlierState();
  if (!state.active) lastChoice = null;
  method.value = lastChoice?.method ?? 'huber'; method.dispatchEvent(new Event('change'));
  if (lastChoice) {
    percent.value = lastChoice.value;
    document.getElementById('source-outlier-axis').value = lastChoice.axis;
    document.getElementById('source-outlier-side').value = lastChoice.side;
    document.getElementById('source-outlier-budget').value = lastChoice.maxRemovalPercent;
  }
  if (state.active) note.textContent = `${state.label}: ${state.excludedPairs}/${state.initialPairs} source pairs removed (${state.excludedPercent.toFixed(3)}%).`;
  document.getElementById('structured-trim-apply').addEventListener('click', async event => {
    if (!methods.has(method.value)) return;
    event.preventDefault(); event.stopImmediatePropagation(); if (applying) return;
    const choice = { method: method.value, value: Number(percent.value), axis: document.getElementById('source-outlier-axis').value, side: document.getElementById('source-outlier-side').value, maxRemovalPercent: Number(document.getElementById('source-outlier-budget').value) };
    applying = true;
    try { lastChoice = choice; await applyWorkspace2SourceOutliers(choice); enhancePlot(); } catch (error) { note.textContent = error.message; } finally { applying = false; }
  }, true);
  const undo = document.getElementById('structured-trim-undo'), reset = document.getElementById('structured-trim-reset');
  if (state.active && !state.hasPlotSteps) undo.disabled = false;
  for (const [button, isUndo] of [[reset, false], [undo, true]]) button.addEventListener('click', async event => {
    const now = workspace2SourceOutlierState();
    if (!now.active || isUndo && now.hasPlotSteps) return;
    event.preventDefault(); event.stopImmediatePropagation(); if (applying) return;
    applying = true;
    try { lastChoice = null; await resetWorkspace2SourceOutliers(); enhancePlot(); } catch (error) { note.textContent = error.message; } finally { applying = false; }
  }, true);
}
new MutationObserver(enhancePlot).observe(document.getElementById('correlation-result'), { childList: true, subtree: true });
enhancePlot();
