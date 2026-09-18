import { readRollPlan, resolveRollZone, referenceRollPlans } from "./complaint-roll-plan.js?v=1";
import { VERSION, selectedCase, matchLots } from './complaint-contract.js';
import { comparisonParameters, compareComplaintLots, comparisonSummary } from './complaint-comparison.js?v=2';
import { lookupComplaintRegion, productRollWidth, reportedRollPairs, machineForRegion } from './complaint-region.js?v=4';

const origin = location.origin;
const token = new URLSearchParams(location.hash.slice(1)).get('ipw-session');
let current = null;
let dataset = { workbook: '', source: '', lots: [], parameters: [] };
let peer = window.opener;
let accepted = false;
let analysis = null;
let regionResult = null;
let rollPlans = referenceRollPlans();
const section = document.createElement('section');
section.className = 'complaint-context';
section.hidden = true;
const heading = document.createElement('h2'); heading.textContent = 'Complaint context';
const description = document.createElement('p');
const matching = document.createElement('p'); matching.setAttribute('role', 'status');
const copy = document.createElement('button'); copy.type = 'button'; copy.className = 'command'; copy.textContent = 'Prepare Copilot summary';
const clear = document.createElement('button'); clear.type = 'button'; clear.className = 'command'; clear.textContent = 'Clear complaint context';
const preview = document.createElement('textarea'); preview.rows = 12; preview.readOnly = true; preview.hidden = true; preview.setAttribute('aria-label', 'Copilot summary preview');
const clipboard = document.createElement('button'); clipboard.type = 'button'; clipboard.textContent = 'Copy for Copilot'; clipboard.hidden = true;
const note = document.createElement('p'); note.textContent = 'Local investigation context. No automatic upload. Review the summary before sharing it with company Copilot. MR/FR and membrane mapping require verification.';
const settings = document.createElement('details'); settings.open = true;
const settingsTitle = document.createElement('summary'); settingsTitle.textContent = 'Complaint Lot comparison';
settings.append(settingsTitle);
function selectField(text, multiple = false) {
  const label = document.createElement('label'); label.textContent = text;
  const select = document.createElement('select'); select.multiple = multiple;
  select.setAttribute('aria-label', text); label.append(select); settings.append(label);
  return select;
}
const parameter = selectField('Comparison parameter');
const type = selectField('Comparison Type');
const searchLabel = document.createElement('label'); searchLabel.textContent = 'Find reference Lot';
const search = document.createElement('input'); search.type = 'search'; search.setAttribute('aria-label', 'Find reference Lot'); searchLabel.append(search); settings.append(searchLabel);
const references = document.createElement('div'); references.className = 'complaint-reference-lots';
references.setAttribute('aria-label', 'Reference Lots'); settings.append(references);
let referenceSelection = new Set();
const referenceCount = document.createElement('p'); referenceCount.setAttribute('role','status'); settings.append(referenceCount);
const zoneFields = document.createElement('fieldset'); const legend = document.createElement('legend'); legend.textContent = 'Comparison Zones'; zoneFields.append(legend);
for (const z of [1,2,3,4,5,6]) { const label = document.createElement('label'); const box = document.createElement('input'); box.type = 'checkbox'; box.value = z; box.checked = true; label.append(box, `Zone ${z}`); zoneFields.append(label); }
settings.append(zoneFields);
const scope = document.createElement('p'); scope.textContent = 'Uses complete selected Lots from the active source, independently of dataset filters and Excel row visibility. Each Lot contributes one mean. Choose comparable reference Lots manually; unrecorded complaints do not establish normality.'; settings.append(scope);
const run = document.createElement('button'); run.type = 'button'; run.className = 'command'; run.textContent = 'Compare selected Lots'; settings.append(run);
const results = document.createElement('pre'); results.className = 'complaint-comparison-result'; results.setAttribute('role','status');
const region = document.createElement('details'); region.open = true;
const regionHeading = document.createElement('summary'); regionHeading.textContent = '1. Locate reported Lot / MR / FR'; region.append(regionHeading);
function regionField(title, tag) {
  const label = document.createElement('label'); label.textContent = title;
  const control = document.createElement(tag); control.setAttribute('aria-label', title);
  label.append(control); region.append(label); return control;
}
const regionLot = regionField('Reported Lot', 'select');
const reportedPair = regionField('Reported MR / FR pair', 'select');
const mr = regionField('Master Roll N', 'input'); mr.inputMode = 'numeric';
const fr = regionField('Final Roll number', 'input'); fr.inputMode = 'numeric';
const product = regionField("Product number (width suffix)", "input");
const planFile = regionField("Optional local Master–Final Roll Plan override", "input"); planFile.type="file"; planFile.accept=".xlsm,.xlsx";
const planStatus=document.createElement("p"); planStatus.setAttribute("role","status"); region.append(planStatus); planStatus.textContent="Using seven verified Roll Plan layouts shared with Report Extraction.";
planFile.addEventListener("change", async () => {
  rollPlans=referenceRollPlans();
  try {
    if (!planFile.files[0]) { planStatus.textContent="Using verified built-in Roll Plan."; return; }
    const xlsx=window.XLSX;
    rollPlans=readRollPlan(xlsx.read(await planFile.files[0].arrayBuffer(), {type:"array",bookVBA:false}), xlsx);
    planStatus.textContent=`Roll Plan loaded locally: ${rollPlans.size} validated layouts. Macros are not executed.`;
  } catch(error) { planStatus.textContent=`${error.message} Using verified built-in Roll Plan instead.`; }
});
const regionZone = regionField('Verified IPW Zone', 'select'); regionZone.append(new Option("Automatic — use verified Roll Plan", "auto"), new Option('Unverified — show all six Zones', ''), ...[1,2,3,4,5,6].map(z=>new Option(`Zone ${z}`,z)));
const regionNote = document.createElement('p'); regionNote.textContent = 'Enter the customer-reported MR and FR. Automatic lookup uses the same verified Roll Plan as Report Extraction. A local updated Plan is optional. Manual Zone selection and all-Zone inspection are optional. IPW measurements describe sampled MR/Zone positions; they are not individual FR measurements.'; region.append(regionNote);
const lookup = document.createElement('button'); lookup.type='button'; lookup.className='command'; lookup.textContent='Find region values'; region.append(lookup);
const regionStatus = document.createElement('p'); regionStatus.setAttribute('role','status'); region.append(regionStatus);
const regionTable = document.createElement('div'); regionTable.className='table-wrap'; region.append(regionTable);
settingsTitle.textContent = '2. Compare complaint Lots with reference Lots';
section.append(heading, description, matching, region, settings, results, copy, clear, preview, clipboard, note);
document.getElementById('dataset-drawer').before(section);
function invalidate() {
  analysis = null; results.textContent = ''; preview.hidden = clipboard.hidden = true; preview.value = ''; clipboard.textContent = 'Copy for Copilot';
}
function renderReferences() {
  references.replaceChildren();
  const targets = current ? matchLots(current.lot, dataset.lots).matched : [];
  for (const lot of dataset.lots.filter(lot => !targets.includes(lot) && lot.toLowerCase().includes(search.value.trim().toLowerCase()))) {
    const label = document.createElement('label'), box = document.createElement('input');
    box.type = 'checkbox'; box.checked = referenceSelection.has(lot);
    box.addEventListener('change', () => { box.checked ? referenceSelection.add(lot) : referenceSelection.delete(lot); invalidate(); referenceCount.textContent = `${referenceSelection.size} reference Lots selected`; });
    label.append(box, lot); references.append(label);
  }
  referenceCount.textContent = `${referenceSelection.size} reference Lots selected`;
}
function render() {
  section.hidden = !current;
  invalidate(); referenceSelection.clear(); search.value = ''; renderReferences();
  if (!current) return;
  region.hidden = !dataset.workbook;
  regionResult = null; regionStatus.textContent=''; regionTable.replaceChildren();
  mr.value = ''; fr.value = ''; regionZone.value = 'auto'; product.value=current.materialNo || '';
  regionLot.replaceChildren(...matchLots(current.lot,dataset.lots).matched.map(lot=>new Option(lot,lot)));
  reportedPair.replaceChildren(new Option('Enter MR / FR manually', ''), ...reportedRollPairs(current.mrfrAreas || current.mrfrCombined).map(pair=>new Option(`MR ${pair.masterRoll} / FR ${pair.finalRoll}`, `${pair.masterRoll}:${pair.finalRoll}`)));
  description.textContent = `${current.complaintNo} · Lot ${current.lot} · ${current.membraneType || 'Membrane unspecified'} · ${current.standardizedSymptoms || current.customerReportedFailure || 'Symptom unspecified'} · MR/FR: ${current.mrfrAreas || current.mrfrCombined || 'Unspecified'}`;
  settings.hidden = !dataset.workbook;
  if (!dataset.workbook) { matching.textContent = 'Open a local IPW workbook to match Lots.'; return; }
  parameter.replaceChildren(...comparisonParameters(dataset.headers || []).map(p => new Option(p, p)));
  const types = [...new Set((dataset.rows || []).map(row => String(row[(dataset.headers || []).indexOf('Type')] ?? '').trim()).filter(Boolean))].sort();
  type.replaceChildren(new Option('All types (explicit selection)', '__all__'), ...types.map(t => new Option(t,t)));
  type.value = types.includes('P') ? 'P' : '__all__';
  const result = matchLots(current.lot, dataset.lots);
  matching.textContent = `${dataset.workbook} · ${dataset.source} · Matched Lots: ${result.matched.join(', ') || 'None'}${result.unmatched.length ? ` · Unmatched: ${result.unmatched.join(', ')}` : ''}. Analysis filters remain under your control.`;
}
window.addEventListener('ipw-investigation-dataset', event => { dataset = event.detail; render(); });
search.addEventListener('input', renderReferences);
settings.addEventListener('change', event => { if (event.target === parameter || event.target === type || zoneFields.contains(event.target)) invalidate(); });
run.addEventListener('click', () => {
  invalidate();
  try {
    const matched = matchLots(current.lot, dataset.lots);
    if (matched.unmatched.length) throw new Error('Resolve unmatched reported Lots before comparing this case.');
    analysis = compareComplaintLots({ headers: dataset.headers, rows: dataset.rows, targetLots: matched.matched, referenceLots: [...referenceSelection], parameter: parameter.value, type: type.value, zones: [...zoneFields.querySelectorAll('input:checked')].map(box => Number(box.value)), isError: dataset.isError });
    results.textContent = comparisonSummary(analysis);
  } catch (error) { results.textContent = error.message; }
});
region.addEventListener('input', () => { regionResult=null; regionStatus.textContent=''; regionTable.replaceChildren(); preview.hidden=clipboard.hidden=true; preview.value=''; });
reportedPair.addEventListener('change', () => {
  const pair=reportedPair.value.split(':'); mr.value=pair[0] || ''; fr.value=pair[1] || '';
  regionResult=null; regionStatus.textContent=''; regionTable.replaceChildren(); preview.hidden=clipboard.hidden=true;
});
lookup.addEventListener('click', () => {
  regionResult=null; regionTable.replaceChildren(); preview.hidden=clipboard.hidden=true;
  try {
    const width=productRollWidth(product.value);
    const machines=machineForRegion(dataset.rawMachineTable || [],regionLot.value,mr.value);
    const mapping=regionZone.value==='auto' ? resolveRollZone(rollPlans,machines,width,fr.value) : null;
    regionResult=lookupComplaintRegion({ headers: dataset.headers, rows: dataset.rows, lot:regionLot.value, masterRoll:mr.value, finalRoll:fr.value, zone:mapping ? mapping.zone : regionZone.value ? Number(regionZone.value) : null, locate:dataset.locate, isError:dataset.isError });
    regionResult.width=width; regionResult.mapping=mapping;
    regionResult.machines=machines;
    regionStatus.textContent=`Product ${product.value || 'unspecified'}: ${regionResult.width ? `${regionResult.width} mm (product suffix)` : 'width unverified'}. Machine from Auswertung: ${regionResult.machines.length===1 ? regionResult.machines[0] : regionResult.machines.length ? `Ambiguous: ${regionResult.machines.join(', ')}` : 'unavailable'}. ${dataset.source}: ${regionResult.matchingRows} matching Lot/MR row(s). ${regionResult.zone ? mapping ? `Zone ${mapping.zone} from Plan ${mapping.layout}!${mapping.cell} / ${mapping.zoneCell}` : `Verified Zone ${regionResult.zone} (manual)` : 'FR mapping unverified; all Zones shown'}. Type is shown separately for each record. No records were averaged or removed.`;
    const table=document.createElement('table'), head=document.createElement('thead'), tr=document.createElement('tr');
    for (const title of ['Source row','Cell','Type','Parameter','Zone','Value']) { const cell=document.createElement('th'); cell.textContent=title; tr.append(cell); }
    head.append(tr); table.append(head); const body=document.createElement('tbody');
    for (const item of regionResult.values) {
      const row=document.createElement('tr');
      const value=item.error ? 'Excel error' : item.value===null || item.value===undefined || String(item.value).trim()==='' ? 'Missing' : String(item.value);
      for (const val of [item.sourceRow ?? 'Unavailable',item.sourceCell || 'Unavailable',item.type,item.parameter,item.zone ?? 'MR-level',value]) { const cell=document.createElement('td'); cell.textContent=val; row.append(cell); } body.append(row);
    }
    table.append(body); regionTable.append(table);
  } catch(error) { regionStatus.textContent=error.message; }
});
window.addEventListener('message', event => {
  if (!token || !peer || event.origin !== origin || event.source !== peer || event.data?.session !== token || event.data?.version !== VERSION || event.data?.type !== 'ipw-selected-case' || accepted) return;
  try {
    current = selectedCase(event.data.case); accepted = true; render();
    peer.postMessage({ type: 'ipw-case-received', version: VERSION, session: token }, origin);
    history.replaceState(null, '', location.pathname + location.search);
  } catch (error) {
    peer.postMessage({ type: 'ipw-case-rejected', version: VERSION, session: token, reason: error.message }, origin);
  }
});
clear.addEventListener('click', () => { current = null; render(); });
copy.addEventListener('click', () => {
  const match = matchLots(current.lot, dataset.lots);
  preview.value = [
    'Investigate whether historical process measurements distinguish this complaint.',
    `Complaint: ${current.complaintNo}`, `Reported Lot(s): ${current.lot}`, `Membrane: ${current.membraneType || 'Unknown'}`,
    `Symptoms: ${current.standardizedSymptoms || current.customerReportedFailure || 'Unknown'}`,
    `MR/FR (unverified): ${current.mrfrAreas || current.mrfrCombined || 'Unknown'}`,
    `Source sheet: ${dataset.source || 'No workbook loaded'}`, `Matched Lots: ${match.matched.join(', ') || 'None'}`,
    `Available parameters: ${dataset.parameters.join(', ')}`,
    regionResult ? `Local region lookup: Lot ${regionResult.lot}; MR ${regionResult.masterRoll}; reported FR ${regionResult.finalRoll || 'Unknown'}; ${regionResult.zone ? regionResult.mapping ? `Zone ${regionResult.zone} from verified Roll Plan ${regionResult.mapping.layout}` : `Zone ${regionResult.zone} manually verified` : 'FR-to-Zone unverified; all Zones inspected'}; matching source rows=${regionResult.matchingRows}. Individual measurements are retained locally and excluded from this summary.` : 'No region lookup has been performed.',
    analysis ? comparisonSummary(analysis) : 'No statistical comparison has been included in this context summary. Do not infer an abnormal measurement or root cause.',
    'Propose comparisons with comparable membrane, production period and process conditions. Unrecorded complaints are not confirmed normal controls. Distinguish hypotheses from evidence. Future prediction requires validation on later Lots; keep all MR/Zone observations of a Lot together.'
  ].join('\n');
  preview.hidden = clipboard.hidden = false;
});
clipboard.addEventListener('click', async () => {
  try { await navigator.clipboard.writeText(preview.value); clipboard.textContent = 'Copied'; }
  catch { preview.focus(); preview.select(); clipboard.textContent = 'Select text and copy manually'; }
});
if (token && /^[a-f0-9]{32}$/.test(token) && peer) {
  peer.postMessage({ type: 'ipw-case-ready', version: VERSION, session: token }, origin);
}
