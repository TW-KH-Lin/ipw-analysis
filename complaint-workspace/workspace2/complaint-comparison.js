const label = value => String(value ?? '').trim();
function numeric(value) {
  if (typeof value === 'number') return Number.isFinite(value) ? value : null;
  if (typeof value !== 'string' || !value.trim()) return null;
  const normalized = value.trim().replace(',', '.');
  if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(normalized)) return null;
  const n = Number(normalized); return Number.isFinite(n) ? n : null;
}
export function comparisonParameters(headers) {
  return [...new Set(headers.flatMap(header => {
    const match = label(header).match(/^(.+)_([1-6])$/);
    return match ? [match[1]] : [];
  }))].sort();
}
export function statistics(values) {
  if (!values.length) return { n: 0, mean: null, sd: null, min: null, max: null };
  const mean = values.reduce((sum, v) => sum + v, 0) / values.length;
  const sd = values.length > 1 ? Math.sqrt(values.reduce((sum, v) => sum + (v - mean) ** 2, 0) / (values.length - 1)) : null;
  return { n: values.length, mean, sd, min: Math.min(...values), max: Math.max(...values) };
}
export function compareComplaintLots({ headers, rows, targetLots, referenceLots, parameter, zones = [1,2,3,4,5,6], type = 'P', isError = () => false }) {
  const lotColumn = headers.findIndex(h => label(h).toLowerCase() === 'lot');
  const typeColumn = headers.findIndex(h => label(h).toLowerCase() === 'type');
  if (lotColumn < 0) throw new Error('Lot column is required.');
  if (type !== '__all__' && typeColumn < 0) throw new Error('Type column is unavailable. Select All types explicitly.');
  const targets = [...new Set(targetLots.map(label))];
  const references = [...new Set(referenceLots.map(label))];
  if (!targets.length || !references.length) throw new Error('Select at least one matched target and one reference Lot.');
  if (targets.some(lot => references.includes(lot))) throw new Error('Target Lots cannot be reference Lots.');
  const known = new Set(rows.map(row => label(row[lotColumn])));
  if ([...targets, ...references].some(lot => !known.has(lot))) throw new Error('Selected Lot is absent from this source.');
  if (!zones.length || zones.some(z => !Number.isInteger(z) || z < 1 || z > 6)) throw new Error('Select Zones 1–6.');
  const columns = [...new Set(zones)].map(zone => ({ zone, index: headers.indexOf(`${parameter}_${zone}`) }));
  if (columns.some(c => c.index < 0)) throw new Error('Selected parameter/Zone column is unavailable.');
  function group(lots) {
    const perLot = lots.map(lot => {
      const all = rows.filter(row => label(row[lotColumn]) === lot);
      const selected = all.filter(row => type === '__all__' || label(row[typeColumn]) === type);
      const values = selected.flatMap(row => columns.map(c => isError(row,c.index) ? null : numeric(row[c.index])).filter(v => v !== null));
      return { lot, rows: selected.length, excludedTypeRows: all.length - selected.length, invalidCells: selected.length * columns.length - values.length, ...statistics(values) };
    });
    const means = perLot.filter(l => l.n > 0).map(l => l.mean);
    return { perLot, selectedLots: lots.length, contributingLots: means.length, measurements: perLot.reduce((n,l) => n + l.n, 0), invalidCells: perLot.reduce((n,l) => n + l.invalidCells, 0), excludedTypeRows: perLot.reduce((n,l) => n + l.excludedTypeRows, 0), lotMeans: statistics(means) };
  }
  const target = group(targets), reference = group(references);
  if (!target.contributingLots || !reference.contributingLots) throw new Error('No numeric values in target or reference selection.');
  return { parameter, zones: columns.map(c => c.zone), type, target, reference, delta: target.lotMeans.mean - reference.lotMeans.mean };
}
export function comparisonSummary(result) {
  const fmt = n => n === null ? 'Unavailable' : Number(n.toPrecision(7)).toString();
  return [
    `Parameter: ${result.parameter}; Zones: ${result.zones.join(', ')}; Type: ${result.type === '__all__' ? 'All types' : result.type}`,
    'Scope: complete selected Lots from the active source; independent of IPW dataset filters and saved Excel row visibility. No extreme removal.',
    'Comparison unit: one mean per Lot across available numeric MR/Zone cells, equal weight for each contributing Lot. Within-Lot cells are descriptive measurements, not independent replicates. Missing cells are not imputed.',
    ...[['Target',result.target],['Reference',result.reference]].flatMap(([name, g]) => [
      `${name} Lots selected/contributing: ${g.selectedLots}/${g.contributingLots}; numeric cells N: ${g.measurements}; blank/nonnumeric cells: ${g.invalidCells}; rows excluded by Type: ${g.excludedTypeRows}`,
      `${name} mean of Lot means: ${fmt(g.lotMeans.mean)}; SD across Lot means: ${fmt(g.lotMeans.sd)}`,
      ...g.perLot.map(l => `${name} Lot ${l.lot}: rows=${l.rows}; numeric cells N=${l.n}; mean=${fmt(l.mean)}; sample SD of cells=${fmt(l.sd)}; blank/nonnumeric cells=${l.invalidCells}`)
    ]),
    `Difference in mean of Lot means (target - reference): ${fmt(result.delta)}`,
    'Descriptive comparison only: no significance test, predictive model or QC acceptance rule. Reference Lots are manually selected; absence of recorded complaints does not establish normality. Verify membrane, machine, dates and process comparability. Uneven Zone coverage can affect means.'
  ].join('\n');
}
