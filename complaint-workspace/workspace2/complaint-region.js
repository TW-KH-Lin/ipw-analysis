const text = v => String(v ?? '').trim();
export function productRollWidth(material) {
  const match = text(material).match(/^1UN[A-Z0-9]*?(18|20|25)[A-Z]*$/i);
  return match ? Number(match[1]) : null;
}
export function reportedRollPairs(value) {
  const result = [], pattern = /\bMR\s*-\s*FR\s*:?\s*(\d+(?:\s*\/\s*\d+)*)\s*-\s*(\d+)\b/gi;
  for (const match of text(value).matchAll(pattern)) {
    for (const masterRoll of match[1].split('/').map(v=>v.trim())) {
      if (!result.some(p=>p.masterRoll===masterRoll && p.finalRoll===match[2])) result.push({masterRoll,finalRoll:match[2]});
    }
  }
  return result;
}
export function machineForRegion(table, lot, masterRoll) {
  const headerRow = table.slice(0,30).findIndex(row=>row.some(v=>text(v)==='ChargenNr') && row.some(v=>text(v)==='Nummer'));
  if (headerRow < 0) return [];
  const headers = table[headerRow], lc=headers.indexOf('ChargenNr'), mc=headers.indexOf('Nummer'), machine=headers.indexOf('ZiehmaschinenName');
  if (machine < 0) return [];
  return [...new Set(table.slice(headerRow+1).filter(row=>text(row[lc])===text(lot) && /^\d+$/.test(text(row[mc])) && Number(row[mc])===Number(masterRoll)).map(row=>text(row[machine])).filter(Boolean))].sort();
}
export function excelColumn(index) {
  if (!Number.isInteger(index) || index < 0) return '';
  let result = '', n = index + 1;
  while (n) { n--; result = String.fromCharCode(65 + n % 26) + result; n = Math.floor(n / 26); }
  return result;
}
export function lookupComplaintRegion({ headers, rows, lot, masterRoll, finalRoll = '', zone = null, type = '__all__', locate = () => ({}), isError = () => false }) {
  const lotColumn = headers.findIndex(h => text(h).toLowerCase() === 'lot');
  const mrColumn = headers.findIndex(h => text(h).toLowerCase() === 'n');
  const typeColumn = headers.findIndex(h => text(h).toLowerCase() === 'type');
  if (lotColumn < 0 || mrColumn < 0) throw new Error('Lot and original MR N columns are required.');
  if (!text(lot) || !/^\d+$/.test(text(masterRoll))) throw new Error('Select a Lot and enter a numeric Master Roll N.');
  if (Number(masterRoll) < 1 || (text(finalRoll) && (!/^\d+$/.test(text(finalRoll)) || Number(finalRoll) < 1))) throw new Error('MR and FR numbers must be positive integers.');
  if (zone !== null && (!Number.isInteger(zone) || zone < 1 || zone > 6)) throw new Error('Select Zone 1–6 or show all Zones.');
  const matching = rows.filter(row => text(row[lotColumn]) === text(lot) && /^\d+$/.test(text(row[mrColumn])) && Number(row[mrColumn]) === Number(masterRoll) && (type === '__all__' || text(row[typeColumn]) === type));
  if (!matching.length) throw new Error('No matching Lot / MR / Type rows. Check original MR N and source sheet.');
  const selectedColumns = headers.flatMap((header, index) => {
    const match = text(header).match(/^(.+)_([1-6])$/);
    if (match && (zone === null || Number(match[2]) === zone)) return [{ index, parameter: match[1], zone: Number(match[2]) }];
    if (['Visco.','Water','Humidity','Temp.'].includes(text(header))) return [{ index, parameter: text(header), zone: null }];
    return [];
  });
  const values = matching.flatMap((row, rowIndex) => selectedColumns.map(c => {
    const location = locate(row, c.index);
    return { sourceRow: location.row ?? null, sourceCell: location.cell || '', rowIndex: rowIndex + 1, type: typeColumn < 0 ? '' : text(row[typeColumn]), parameter: c.parameter, zone: c.zone, value: row[c.index], error: isError(row,c.index) };
  }));
  return { lot: text(lot), masterRoll: text(masterRoll), finalRoll: text(finalRoll), zone, matchingRows: matching.length, values };
}
