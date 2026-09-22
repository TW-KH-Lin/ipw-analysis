import { getRegionalParameters, headerIndex, isNumeric, meanAndSigma, number, text, zoneColumns } from '../analysis.js?v=17';

export function commonRegionalParameters(sources) {
  if (sources.length < 2) return [];
  const available = sources.map(source => new Map(getRegionalParameters(source.headers)
    .filter(parameter => zoneColumns(source.headers, parameter).every(index => index >= 0))
    .map(parameter => [parameter.toLowerCase(), parameter])));
  return [...available[0]].filter(([key]) => available.every(parameters => parameters.has(key)))
    .map(([, parameter]) => parameter).sort((a, b) => a.localeCompare(b));
}

export function compareRegionalWorkbooks(sources, parameter, exactLot = '') {
  if (sources.length < 2 || sources.length > 3) throw new Error('Select 2 or 3 workbooks.');
  const common = commonRegionalParameters(sources);
  if (!common.some(name => name.toLowerCase() === parameter.toLowerCase())) {
    throw new Error('Choose a parameter with all six Zones in every selected workbook.');
  }
  const lotFilter = text(exactLot);
  const workbooks = sources.map(source => {
    const matchingParameter = getRegionalParameters(source.headers).find(name => name.toLowerCase() === parameter.toLowerCase());
    const columns = zoneColumns(source.headers, matchingParameter);
    const lotColumn = headerIndex(source.headers, 'Lot');
    if (lotFilter && lotColumn < 0) throw new Error(`${source.name}: Lot column is unavailable.`);
    const rows = lotFilter ? source.rows.filter(row => text(row[lotColumn]) === lotFilter) : source.rows;
    const zones = columns.map(column => meanAndSigma(rows.filter(row => !source.isError?.(row, column))
      .map(row => row[column]).filter(isNumeric).map(number)));
    const values = columns.flatMap(column => rows.filter(row => !source.isError?.(row, column))
      .map(row => row[column]).filter(isNumeric).map(number));
    const lots = lotColumn < 0 ? [] : [...new Set(rows.map(row => text(row[lotColumn])).filter(Boolean))];
    return { id: source.id, name: source.name, source: source.source, machine: source.name.match(/(?:^|[^a-z0-9])(ZM[\s_-]?\d+)(?=[^a-z0-9]|$)/i)?.[1]?.replace(/[\s_-]/g, '').toUpperCase() || '',
      rowCount: rows.length, lotCount: lots.length, zones, overall: meanAndSigma(values) };
  });
  const reference = workbooks[0].overall.mean;
  workbooks.forEach(book => { book.delta = book.overall.mean == null || reference == null ? null : book.overall.mean - reference; });
  return { parameter: common.find(name => name.toLowerCase() === parameter.toLowerCase()), exactLot: lotFilter, workbooks };
}

export function comparisonCsvRows(result) {
  const rows = [['Parameter', result.parameter], ['Lot filter', result.exactLot || 'All lots'], [],
    ['Workbook', 'Sheet', 'Machine', 'Rows', 'Lots', 'Values N', 'Mean', 'SD', 'Min', 'Max', 'Delta vs first']];
  for (const book of result.workbooks) rows.push([book.name, book.source, book.machine, book.rowCount, book.lotCount,
    book.overall.n, book.overall.mean, book.overall.sigma, book.overall.min, book.overall.max, book.delta]);
  rows.push([], ['Workbook', 'Sheet', 'Machine', 'Zone', 'N', 'Mean', 'SD', 'Min', 'Max']);
  for (const book of result.workbooks) book.zones.forEach((zone, index) => rows.push([book.name, book.source, book.machine,
    index + 1, zone.n, zone.mean, zone.sigma, zone.min, zone.max]));
  return rows;
}
