import { extractMany } from './report-import/report-parser.js?v=3';
import { importCandidates, candidateLabel, retainZoneEvidence, checkRetainZones } from './complaint-import-core.js?v=7';
const host=document.getElementById('complaint-label-import');
let retainEvidence=[];
const importedFiles=new Set();
let undoSnapshot=null;
function snapshot() {return {candidates:structuredClone(candidates),retainEvidence:structuredClone(retainEvidence),files:[...importedFiles]};}

let dataset=null, candidates=[], busy=false, generation=0;
const title=document.createElement('h2'); title.textContent='Mark problems from complaint files';
const note=document.createElement('p'); note.textContent='Open final reports, registration PDFs or Outlook .msg emails. Clear customer-reported locations are marked automatically in Data labels. Review unresolved items below. Files stay on this computer; scanned PDFs require selectable text. Email attachments are not read automatically.';
const files=document.createElement('input'); files.type='file'; files.multiple=true; files.accept='.pdf,.msg'; files.setAttribute('aria-label','Open complaint reports or emails');
const folder=document.createElement('input'); folder.type='file'; folder.multiple=true; folder.webkitdirectory=true; folder.setAttribute('aria-label','Open complaint folder');
const fileLabel=document.createElement('label'); fileLabel.textContent='Open PDF or MSG files'; fileLabel.append(files);
const folderLabel=document.createElement('label'); folderLabel.textContent='Open a folder of reports and emails'; folderLabel.append(folder);
const status=document.createElement('p'); status.setAttribute('role','status');
const search=document.createElement('input'); search.type='search'; search.placeholder='Find complaint, file, Lot or status'; search.setAttribute('aria-label','Find imported complaint'); search.addEventListener('input',render);
const list=document.createElement('div'); list.className='complaint-import-review-list';
const retainView=document.createElement('div');
function retainChecks() {return checkRetainZones(retainEvidence,dataset?.getComplaintLabels?.() || []);}
function renderRetain() {
  retainView.replaceChildren();
  const h=document.createElement('h3'); h.textContent='Retain sample cross-check'; retainView.append(h);
  const checks=retainChecks();
  for(const item of checks){const p=document.createElement('p'); p.textContent=`${item.complaintNo}: Lot ${item.lot}, MR ${item.masterRoll}. Report retain Zones ${item.zones.join('/')}; marked Zones ${item.markedZones.join('/') || 'none'}. ${item.result}. Evidence: ${item.evidence}`; retainView.append(p);}
  if(!checks.length) {const p=document.createElement('p');p.textContent='No explicit retain Lot / MR / Zone evidence extracted. Retain tests are not automatically treated as customer problem locations.';retainView.append(p);}
}
const apply=document.createElement('button'); apply.type='button'; apply.textContent='Mark reviewed items';
const save=document.createElement('button'); save.type='button'; save.textContent='Download labeled IPW copy';
const undo=document.createElement('button'); undo.type='button'; undo.textContent='Undo last import batch';
const exportResults=document.createElement('button'); exportResults.type='button'; exportResults.textContent='Export marking results';
exportResults.addEventListener('click',async()=>{busy=true;controls();try {await dataset.exportComplaintResults(candidates,retainChecks());status.textContent='Marking results downloaded: grouped colors and filterable parameter details.';} catch(error) {status.textContent=error.message;} finally {busy=false;controls();}});
const family=document.createElement('select'); family.setAttribute('aria-label','Complaint membrane filter'); family.append(new Option('All membrane types','all'),new Option('CN95 only','CN95'));
host.append(title,note,family,fileLabel,folderLabel,status,apply,save,undo,exportResults,retainView,search,list);
function controls() { folder.disabled=busy || !dataset?.workbook; family.disabled=busy; exportResults.disabled=busy || !dataset?.workbook; files.disabled=busy || !dataset?.workbook; apply.disabled=busy || !candidates.length; save.disabled=busy || !dataset?.workbook; undo.disabled=busy || !dataset?.workbook; }
function render() {
  renderRetain(); list.replaceChildren();
  for(const item of candidates.filter(item=>[item.sourceFile,item.complaintNo,item.lot,item.result].join(' ').toLowerCase().includes(search.value.trim().toLowerCase()))) {
    const card=document.createElement('details'); card.open=false;
    const heading=document.createElement('summary'); heading.textContent=`${item.sourceFile} · ${item.complaintNo || 'Unidentified case'} · ${item.result || 'Needs review'}`; card.append(heading);
    for(const [key,label] of [['complaintNo','Complaint number'],['lot','Reported Lot'],['materialNo','Product number'],['masterRoll','Master Roll N'],['finalRoll','Final Roll number'],['reason','Complaint reason'],['problem','Reported problem']]) {
      const wrap=document.createElement('label'); wrap.textContent=label;
      const input=document.createElement('input'); input.value=item[key] || ''; input.disabled=item.done; input.setAttribute('aria-label',label);
      input.addEventListener('input',()=>{item[key]=input.value; item.requiresReview=true; item.reviewed=false; checked.checked=false;}); wrap.append(input); card.append(wrap);
    }
    const evidence=document.createElement('p'); evidence.textContent=`Extracted customer statement: ${item.evidence || 'Not found'}. ${item.warnings || ''}`; card.append(evidence);
    const confirm=document.createElement('label'), checked=document.createElement('input'); checked.type='checkbox'; checked.checked=item.reviewed; checked.disabled=item.done;
    checked.addEventListener('change',()=>item.reviewed=checked.checked); confirm.append(checked,'I verified this problem and its Lot / MR / FR against the source.'); card.append(confirm); list.append(card);
  }
  controls();
}
function mark() {
  let added=0;
  const ready=[], items=[];
  for(const item of candidates.filter(item=>!item.done)) {
    try {ready.push(candidateLabel(item,dataset)); items.push(item);} catch(error) {item.result=error.message;}
  }
  if (ready.length) {
    const count=dataset.applyComplaintLabels(ready); added=count;
    items.forEach(item=>{item.done=true; item.result='Marked or already present in Data labels';});
    status.textContent=`${count} problem mark(s) added. Download a labeled IPW copy to keep them. ${candidates.filter(item=>!item.done).length} item(s) need review. No analysis rows were excluded.`;
  } else status.textContent='No new locations could be marked. Review missing or ambiguous information below.';
  render(); return added;
}
async function openFiles(fileList) {
  busy=true; controls(); const started=generation; const prior=snapshot(); let loaded=0, skipped=0;
  try {
    for(const file of [...fileList].filter(file=>/\.(pdf|msg)$/i.test(file.name))) {
      status.textContent=`Reading ${file.name} locally…`;
      try {
        const data=await file.arrayBuffer();
        const digest=await crypto.subtle.digest('SHA-256',data.slice(0));
        const sourceId=[...new Uint8Array(digest)].map(v=>v.toString(16).padStart(2,'0')).join('');
        if(importedFiles.has(sourceId)) {skipped++; continue;}
        const records=await extractMany(file.name,data);
        if(started!==generation) throw new Error('IPW workbook changed. Reopen the complaint files for this workbook.');
        importedFiles.add(sourceId); loaded++;
        if (!records.length) candidates.push({sourceFile:file.name,requiresReview:true,warnings:'No complaint extracted. The file may be a blank template or scanned PDF.'});
        for(const record of records) {
          if(family.value==='CN95' && record.membraneType!=='CN95' && record.productFamily!=='CN95') {
            candidates.push({sourceFile:file.name,complaintNo:record.complaintNo,lot:record.lot,done:true,result:'Skipped: not identified as CN95'}); continue;
          }
          retainEvidence.push(...retainZoneEvidence(record).filter(item=>!retainEvidence.some(old=>old.complaintNo===item.complaintNo && old.evidence===item.evidence)));
          candidates.push(...importCandidates(record,sourceId));
        }
      } catch(error) {
        if(started!==generation) throw error;
        candidates.push({sourceFile:file.name,requiresReview:true,warnings:error.message});
      }
    }
    if(loaded) {if(mark()) undoSnapshot=prior;}
    else if(skipped) status.textContent=`${skipped} previously imported file(s) skipped. Existing marks and review items are unchanged.`;
    else status.textContent='No readable complaint files found. Check the review items below.';
  } catch(error) {status.textContent=error.message;}
  finally {busy=false; files.value=''; folder.value=''; render();}
}
files.addEventListener('change',()=>openFiles(files.files));
folder.addEventListener('change',()=>openFiles(folder.files));
apply.addEventListener('click',()=>{try {const prior=snapshot(); if(mark()) undoSnapshot=prior;} catch(error) {status.textContent=error.message;}});
save.addEventListener('click',async()=>{
  busy=true; controls();
  try {await dataset.saveComplaintLabels(); status.textContent='Labeled IPW copy downloaded. Keep this copy to retain the labels.';}
  catch(error) {status.textContent=error.message;}
  finally {busy=false; controls();}
});
undo.addEventListener('click',()=>{
  try {dataset.undoComplaintLabels(); candidates=undoSnapshot?.candidates || []; retainEvidence=undoSnapshot?.retainEvidence || []; importedFiles.clear(); for(const hash of undoSnapshot?.files || []) importedFiles.add(hash); undoSnapshot=null; status.textContent='Last import batch undone. Save the workbook if you previously saved that batch.'; render();}
  catch(error) {status.textContent=error.message;}
});
window.addEventListener('ipw-investigation-dataset',event=>{
  if(!event.detail.workbook) {generation++; undoSnapshot=null; candidates=[]; importedFiles.clear(); retainEvidence=[]; retainView.replaceChildren(); list.replaceChildren(); status.textContent='Open an IPW workbook first.';}
  dataset=event.detail; if(dataset.workbook && !candidates.length) status.textContent="Ready. Open complaint files to mark reported problems."; controls();
});
status.textContent='Open an IPW workbook first.'; controls();
