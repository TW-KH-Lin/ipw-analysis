import { ROLL_PLAN_REFERENCE } from "./roll-plan-reference.js?v=1";
export function referenceRollPlans() {
  return new Map(Object.entries(ROLL_PLAN_REFERENCE).map(([layout,spec])=>[layout,new Map(Array.from({length:spec.slots},(_,i)=>{
    let zone=1; spec.zoneStarts.forEach((start,z)=>{if(i>=start) zone=z+1;});
    return [i+1,{zone,cell:"Built-in reference",zoneCell:"verified against official Plan"}];
  }))]));
}
// Read only the layout headers; no macros, customer records or persistent storage.
export function readRollPlan(workbook, xlsx) {
  const plans = new Map();
  for (const name of workbook.SheetNames) {
    if (!/^ZM\d+_(18|20|25)mm$/.test(name)) continue;
    const sheet = workbook.Sheets[name], rolls = new Map();
    const merges = (sheet['!merges'] || []).filter(m => m.s.r === 3 && m.e.r === 3);
    for (const merge of merges) {
      const label = sheet[xlsx.utils.encode_cell(merge.s)];
      const zone = /^ZONE\s+([1-6])$/i.exec(String(label?.v ?? '').trim());
      if (!zone || label.t === 'e') throw new Error(`Invalid Zone header in ${name}.`);
      for (let c = merge.s.c; c <= merge.e.c; c++) {
        const cell = xlsx.utils.encode_cell({r:4,c}), value = sheet[cell];
        if (value?.t === 'e' || !Number.isSafeInteger(value?.v) || value.v < 1 || rolls.has(value.v)) throw new Error(`Invalid or duplicate FR in ${name}!${cell}.`);
        rolls.set(value.v, {zone:Number(zone[1]), cell, zoneCell:xlsx.utils.encode_cell(merge.s)});
      }
    }
    if (!rolls.size || new Set([...rolls.values()].map(v=>v.zone)).size !== 6 || [...rolls.keys()].sort((a,b)=>a-b).some((v,i)=>v!==i+1)) throw new Error(`Incomplete six-Zone layout: ${name}.`);
    plans.set(name, rolls);
  }
  if (!plans.size) throw new Error('No supported Master–Final Roll Plan layouts found.');
  return plans;
}
export function resolveRollZone(plans, machines, width, finalRoll) {
  if (machines.length !== 1) throw new Error('A unique machine from IPW is required.');
  const machine = /^(?:ZM\s*)?(9|10|17)$/i.exec(String(machines[0]).trim());
  if (!machine) throw new Error('Unsupported IPW machine identifier.');
  if (![18,20,25].includes(width)) throw new Error('Enter a product number ending in 18, 20 or 25.');
  if (!/^[1-9]\d*$/.test(String(finalRoll))) throw new Error('Enter a positive Final Roll number.');
  const layout = `ZM${machine[1]}_${width}mm`, result = plans.get(layout)?.get(Number(finalRoll));
  if (!result) throw new Error(`FR ${finalRoll} is not covered by the loaded ${layout} plan.`);
  return {layout, ...result};
}
