import { fitHuberGaussian } from '../analysis.js?v=17';
const identity = value => String(value ?? '').trim().toUpperCase();
const numeric = value => (typeof value === 'number' || typeof value === 'string') && String(value).trim() !== '' && Number.isFinite(Number(value));
function columns(headers, parameter) {
  const zones = new Map(); let scalar = -1;
  headers.forEach((header, column) => {
    if (identity(header) === identity(parameter)) scalar = column;
    const match = String(header ?? '').trim().match(/^(.*)_(\d+)$/);
    if (match && identity(match[1]) === identity(parameter) && +match[2] > 0) {
      if (zones.has(+match[2])) throw new Error('Duplicate Zone header.');
      zones.set(+match[2], column);
    }
  });
  return { zones, scalar };
}
function quantile(sorted, probability) {
  const position = (sorted.length - 1) * probability, lower = Math.floor(position);
  return sorted[lower] + (sorted[Math.min(lower + 1, sorted.length - 1)] - sorted[lower]) * (position - lower);
}
export function filterSourceOutliers(headers, sourceRows, xParameter, yParameter, options = {}) {
  const method = options.method ?? 'huber', side = options.side ?? 'both', axis = options.axis ?? 'x';
  const value = Number(options.value ?? (method === 'percentile' ? 1 : 3)), budget = Number(options.maxRemovalPercent ?? 2);
  if (!['huber', 'percentile', 'iqr'].includes(method) || !['high', 'low', 'both'].includes(side) || !['x', 'y'].includes(axis)) throw new Error('Choose a valid method, direction and axis.');
  if (!(value > 0) || !Number.isFinite(value) || method === 'percentile' && value >= 50) throw new Error('Enter a positive cutoff; percentile tails must be less than 50%.');
  if (!Number.isFinite(budget) || budget < 0 || budget > 20) throw new Error('Maximum removal must be between 0 and 20%.');
  if (identity(xParameter) === identity(yParameter)) throw new Error('X and Y must differ.');
  const xs = columns(headers, xParameter), ys = columns(headers, yParameter), scalar = !xs.zones.size && !ys.zones.size;
  if (Boolean(xs.zones.size) !== Boolean(ys.zones.size) || scalar && (xs.scalar < 0 || ys.scalar < 0)) throw new Error('X and Y must have the same measurement level.');
  const lotColumn = headers.findIndex(h => identity(h) === 'LOT'), nColumn = headers.findIndex(h => identity(h) === 'N');
  if (lotColumn < 0 || nColumn < 0) throw new Error('Lot and N columns are required.');
  const pairs = scalar ? [[0, xs.scalar, ys.scalar]] : [...xs.zones].filter(([zone]) => ys.zones.has(zone)).map(([zone, column]) => [zone, column, ys.zones.get(zone)]);
  const observations = [];
  sourceRows.forEach((row, rowIndex) => {
    if (!identity(row[lotColumn]) || !String(row[nColumn] ?? '').trim()) return;
    pairs.forEach(([zone, xColumn, yColumn]) => {
      if (numeric(row[xColumn]) && numeric(row[yColumn])) observations.push({ rowIndex, column: axis === 'x' ? xColumn : yColumn, zone: zone || null, lot: row[lotColumn], N: row[nColumn], value: Number(row[axis === 'x' ? xColumn : yColumn]) });
    });
  });
  if (observations.length < 3) throw new Error('At least three paired source observations are required.');
  const sorted = observations.map(o => o.value).sort((a, b) => a - b);
  let lower, upper, label;
  if (method === 'huber') {
    const fit = fitHuberGaussian(sorted);
    if (!(fit.sigma > 0) || !Number.isFinite(fit.sigma)) throw new Error('Huber scale is unavailable for this data.');
    lower = fit.mean - value * fit.sigma; upper = fit.mean + value * fit.sigma; label = `Huber ${value}σ`;
  } else if (method === 'percentile') {
    lower = quantile(sorted, value / 100); upper = quantile(sorted, 1 - value / 100); label = `Percentile ${value}% per tail`;
  } else {
    const q1 = quantile(sorted, .25), q3 = quantile(sorted, .75), iqr = q3 - q1;
    if (!(iqr > 0)) throw new Error('IQR is zero. Choose another method.');
    lower = q1 - value * iqr; upper = q3 + value * iqr; label = `IQR ${value}×`;
  }
  const exclusions = observations.filter(o => side !== 'low' && o.value > upper || side !== 'high' && o.value < lower);
  const excludedPercent = 100 * exclusions.length / observations.length;
  if (excludedPercent > budget + 1e-12) throw new Error(`This cutoff would remove ${excludedPercent.toFixed(3)}%, above the ${budget}% limit. Increase the cutoff or adjust the limit.`);
  if (observations.length - exclusions.length < 3) throw new Error('Keep at least three source pairs.');
  const rows = sourceRows.map(row => row.slice()); exclusions.forEach(o => { rows[o.rowIndex][o.column] = null; });
  return { rows, metadata: { method, label, side, axis, value, lower, upper, initialPairs: observations.length, excludedPairs: exclusions.length, excludedPercent, maxRemovalPercent: budget, exclusions } };
}
