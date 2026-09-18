import { parseMrFr } from "./report-import/report-parser.js?v=3";
import { productRollWidth, reportedRollPairs, machineForRegion } from './complaint-region.js?v=4';
import { referenceRollPlans, resolveRollZone } from './complaint-roll-plan.js?v=1';
import { buildDataLabel, dataLabelKey } from '../data-management.js?v=5';

const text=value=>String(value ?? '').trim();
export function retainZoneEvidence(record) {
  const raw=text(record.rawText).replace(/\s+/g,' ');
  const matches=[...raw.matchAll(/retain(?:ed)?\s+samples?\b.{0,90}?\blot\s+(\d{7,9})\s+MR\s*(\d+)\s+Zone\s*([1-6](?:\s*[/,]\s*[1-6])*)/gi)];
  return matches.map(m=>({sourceFile:record.sourceFile,complaintNo:record.complaintNo,lot:m[1],masterRoll:m[2],zones:[...new Set(m[3].match(/[1-6]/g).map(Number))],evidence:m[0]}));
}
export function checkRetainZones(evidence, labels) {
  return evidence.map(item=>{
    const actual=[...new Set(labels.filter(label=>text(label.lot)===item.lot && Number(label.batch)===Number(item.masterRoll) && label.notes?.includes(`:${item.complaintNo}:${item.lot}:`)).map(label=>label.zone))].sort();
    const overlap=actual.filter(zone=>item.zones.includes(zone));
    const result=!actual.length ? 'No problem mark for this Lot/MR yet' : actual.length===item.zones.length && overlap.length===actual.length ? 'Same Zone set' : overlap.length ? 'Partial overlap — review the different sample scopes' : 'Different Zones — review required';
    return {...item,markedZones:actual,result};
  });
}
export function problemValuesTable(review, dataset) {
  const headers=dataset.headers, parameters=[...new Set(headers.flatMap(h=>/^(.+)_([1-6])$/.exec(text(h))?.[1] || []))];
  const scalars=['Visco.','Water','Humidity','Temp.'].filter(h=>headers.includes(h));
  const output=[['Complaint','Lot','MR','FR','Zone','Complaint reason',...parameters,...scalars.map(s=>`${s} (MR level)`),'Machine','Width mm','Type','IPW source row','IPW source sheet','Source file','Missing or error parameters']];
  const seen=new Set();
  for(const item of review.filter(item=>item.done && item.result==='Marked or already present in Data labels')) {
    const identity=[item.sourceId,item.complaintNo,item.lot,item.masterRoll,item.finalRoll].join('|');
    if(seen.has(identity)) continue;
    seen.add(identity);
    const label=candidateLabel(item,dataset);
    const row=dataset.rows.find(row=>text(row[headers.indexOf('Lot')])===text(item.lot) && Number(row[headers.indexOf('N')])===Number(item.masterRoll));
    const columns=[...parameters.map(p=>headers.indexOf(`${p}_${label.zone}`)),...scalars.map(s=>headers.indexOf(s))];
    const names=[...parameters,...scalars], missing=[];
    const values=columns.map((col,i)=>{
      if(col<0 || dataset.isError?.(row,col) || row[col]===null || row[col]===undefined || text(row[col])==='') {missing.push(names[i]); return '';}
      return row[col];
    });
    output.push([item.complaintNo,item.lot,Number(item.masterRoll),Number(item.finalRoll),label.zone,item.reason || item.problem,...values,machineForRegion(dataset.rawMachineTable,item.lot,item.masterRoll)[0],productRollWidth(item.materialNo),row[headers.indexOf('Type')],dataset.locate?.(row,headers.indexOf('N'))?.row ?? '',dataset.source,item.sourceFile,missing.join(', ')]);
  }
  return output;
}
export function importCandidates(record, sourceId) {
  const pairs=reportedRollPairs(parseMrFr(record.rawText || record.mrfrAreas || record.mrfrCombined).areas.join('; '));
  const section=String(record.rawText || '').match(/1\.1\.\s*Customer statement\s*([\s\S]*?)(?:1\.2\.\s*Criticality)/i)?.[1];
  const evidence=section || record.customerReportedFailure || record.formalProblem || '';
  const customerPairs=reportedRollPairs(parseMrFr(evidence).areas.join('; '));
  const raw=section || record.rawText || evidence;
  const lots=[...raw.matchAll(/\b(?:lot|batch)\s*(?:number|no\.?)?\s*[:#]?\s*(\d{7,9})\b/gi)].map(m=>m[1]);
  const materials=raw.match(/\b1UN[A-Z0-9]+\b/gi) || [];
  const ambiguous=/\b(?:is|was|are|were|i\s+s)\s+(?:qualified|acceptable|conforming)\b/i.test(evidence) || new Set(lots).size>1 || new Set(materials.map(v=>v.toUpperCase())).size>1 || /Multi-complaint|Filename says/i.test(record.warnings || '');
  return (pairs.length ? pairs : [{masterRoll:'',finalRoll:''}]).map(pair=>({
    sourceId, sourceFile:record.sourceFile, complaintNo:text(record.complaintNo), lot:text(record.lot),
    materialNo:text(record.materialNo), masterRoll:pair.masterRoll, finalRoll:pair.finalRoll,
    reason:record.standardizedSymptoms && record.standardizedSymptoms!=='Review required' ? record.standardizedSymptoms : text(record.customerReportedFailure || record.formalProblem || record.problem),
    problem:text(record.customerReportedFailure || record.formalProblem || record.problem),
    evidence, warnings:record.warnings || '', reviewed:false,
    requiresReview:ambiguous || !customerPairs.some(p=>p.masterRoll===pair.masterRoll && p.finalRoll===pair.finalRoll),
  }));
}
export function candidateLabel(candidate, dataset) {
  if (candidate.requiresReview && !candidate.reviewed) throw new Error('Review required: confirm the reported location; the document may also list reference or retain rolls.');
  if (!text(candidate.complaintNo) || !text(candidate.problem)) throw new Error('Complaint number and problem description are required.');
  if (!/^\d+$/.test(text(candidate.masterRoll)) || Number(candidate.masterRoll)<1) throw new Error('Enter the reported Master Roll N.');
  const headers=dataset.headers || [], lc=headers.indexOf('Lot'), mc=headers.indexOf('N');
  const rows=(dataset.rows || []).filter(row=>text(row[lc])===text(candidate.lot) && /^\d+$/.test(text(row[mc])) && Number(row[mc])===Number(candidate.masterRoll));
  if (rows.length!==1) throw new Error(rows.length ? 'Multiple IPW records match this Lot/MR. Resolve the source records before marking.' : 'No exact Lot/MR match in the active IPW source.');
  const width=productRollWidth(candidate.materialNo);
  const machines=machineForRegion(dataset.rawMachineTable || [],candidate.lot,candidate.masterRoll);
  const mapping=resolveRollZone(referenceRollPlans(),machines,width,text(candidate.finalRoll));
  const parameters=headers.flatMap(h=> {const m=/^(.+)_([1-6])$/.exec(text(h)); return m && Number(m[2])===mapping.zone ? [m[1]] : [];});
  const safeRow=rows[0].map((v,c)=>dataset.isError?.(rows[0],c) ? null : v);
  const marker=`[Complaint import ${candidate.sourceId}:${candidate.complaintNo}:${candidate.lot}:${Number(candidate.masterRoll)}:${Number(candidate.finalRoll)}]`;
  return buildDataLabel(headers,[safeRow],{
    lot:candidate.lot,batch:rows[0][mc],zone:mapping.zone,parameters,
    label:`Complaint ${candidate.complaintNo}: ${candidate.reason || candidate.problem}`,
    comment:'Reported problem area. Parameter highlighting does not establish abnormal values or root cause.',
    notes:[marker,`Source: ${candidate.sourceFile}`,`Product: ${candidate.materialNo}; machine: ${machines[0]}; width: ${width} mm; MR ${candidate.masterRoll}; FR ${candidate.finalRoll}; Zone ${mapping.zone}; plan: ${mapping.layout}`,`Problem: ${candidate.problem}`,`Location: ${candidate.reviewed ? 'reviewed by user' : 'explicitly reported in customer problem section'}`].join('\n')
  });
}
export function mergeProblemLabels(existing, incoming) {
  const next=existing.slice(); let added=0;
  for (const label of incoming) {
    const index=next.findIndex(item=>dataLabelKey(item.lot,item.batch,item.zone)===dataLabelKey(label.lot,label.batch,label.zone));
    const prior=next[index];
    if (prior && text(prior.lot)!==text(label.lot)) throw new Error("Existing label has a conflicting Lot identifier. Review leading zeros before importing.");
    const marker=label.notes.split('\n')[0];
    if (prior?.notes?.includes(marker)) continue;
    if (!prior) next.push(label);
    else {
      const parameters=[...new Set([...prior.parameters,...label.parameters])];
      next[index]={...prior,parameters,values:{...label.values,...prior.values},
        valuesText:[...new Set([prior.valuesText,label.valuesText].filter(Boolean).flatMap(value=>value.split('; ')))].join('; '),
        label:[...new Set([prior.label,label.label].filter(Boolean).flatMap(value=>value.split('\n')))].join('\n'),
        comment:[...new Set([prior.comment,label.comment].filter(Boolean).flatMap(value=>value.split('\n')))].join('\n'),
        notes:[prior.notes,label.notes].filter(Boolean).join('\n'),updated:label.updated};
    }
    added++;
  }
  return {labels:next,added};
}
