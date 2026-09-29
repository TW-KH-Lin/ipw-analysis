function numeric(value) {
  return value !== "" && value !== null && value !== undefined && Number.isFinite(Number(value));
}

function summary(values) {
  if (!values.length) return { n: 0, mean: null, median: null, sigma: null, min: null, max: null };
  const sorted = [...values].sort((a, b) => a - b);
  const mean = values.reduce((total, value) => total + value, 0) / values.length;
  const sigma = values.length > 1
    ? Math.sqrt(values.reduce((total, value) => total + (value - mean) ** 2, 0) / (values.length - 1))
    : 0;
  const middle = Math.floor(sorted.length / 2);
  const median = sorted.length % 2 ? sorted[middle] : (sorted[middle - 1] + sorted[middle]) / 2;
  return { n: values.length, mean, median, sigma, min: sorted[0], max: sorted[sorted.length - 1] };
}

function ranks(values) {
  const ordered = values.map((value, index) => ({ value, index })).sort((a, b) => a.value - b.value);
  const result = Array(values.length);
  const ties = [];
  for (let start = 0; start < ordered.length;) {
    let end = start + 1;
    while (end < ordered.length && ordered[end].value === ordered[start].value) end += 1;
    const rank = (start + 1 + end) / 2;
    for (let index = start; index < end; index += 1) result[ordered[index].index] = rank;
    if (end - start > 1) ties.push(end - start);
    start = end;
  }
  return { values: result, ties };
}

function logGamma(value) {
  const coefficients = [676.5203681218851, -1259.1392167224028, 771.3234287776531, -176.6150291621406,
    12.507343278686905, -0.13857109526572012, 9.984369578019572e-6, 1.5056327351493116e-7];
  if (value < 0.5) return Math.log(Math.PI) - Math.log(Math.sin(Math.PI * value)) - logGamma(1 - value);
  let x = 0.9999999999998099;
  const z = value - 1;
  coefficients.forEach((coefficient, index) => { x += coefficient / (z + index + 1); });
  const t = z + coefficients.length - 0.5;
  return 0.5 * Math.log(2 * Math.PI) + (z + 0.5) * Math.log(t) - t + Math.log(x);
}

function regularizedGammaQ(shape, value) {
  if (!(shape > 0) || value < 0) return NaN;
  if (value === 0) return 1;
  const epsilon = 1e-14;
  if (value < shape + 1) {
    let sum = 1 / shape, term = sum, current = shape;
    for (let iteration = 1; iteration < 1000; iteration += 1) {
      current += 1;
      term *= value / current;
      sum += term;
      if (Math.abs(term) < Math.abs(sum) * epsilon) break;
    }
    const lower = sum * Math.exp(-value + shape * Math.log(value) - logGamma(shape));
    return Math.max(0, Math.min(1, 1 - lower));
  }
  let b = value + 1 - shape, c = 1 / 1e-300, d = 1 / b, fraction = d;
  for (let iteration = 1; iteration < 1000; iteration += 1) {
    const coefficient = -iteration * (iteration - shape);
    b += 2;
    d = coefficient * d + b;
    if (Math.abs(d) < 1e-300) d = 1e-300;
    c = b + coefficient / c;
    if (Math.abs(c) < 1e-300) c = 1e-300;
    d = 1 / d;
    const delta = d * c;
    fraction *= delta;
    if (Math.abs(delta - 1) < epsilon) break;
  }
  return Math.max(0, Math.min(1, Math.exp(-value + shape * Math.log(value) - logGamma(shape)) * fraction));
}

function friedman(completeCases) {
  const n = completeCases.length;
  const k = completeCases[0]?.length || 0;
  if (n < 3 || k < 2) return { n, statistic: null, df: k ? k - 1 : null, p: null, kendallW: null };
  const rankSums = Array(k).fill(0);
  let tieTotal = 0;
  completeCases.forEach((row) => {
    const ranked = ranks(row);
    ranked.values.forEach((rank, index) => { rankSums[index] += rank; });
    ranked.ties.forEach((size) => { tieTotal += size ** 3 - size; });
  });
  const raw = 12 * rankSums.reduce((total, value) => total + value ** 2, 0) / (n * k * (k + 1)) - 3 * n * (k + 1);
  const correction = 1 - tieTotal / (n * (k ** 3 - k));
  const statistic = correction > 0 ? raw / correction : 0;
  return { n, statistic, df: k - 1, p: regularizedGammaQ((k - 1) / 2, statistic / 2), kendallW: statistic / (n * (k - 1)) };
}

function exactSignP(positive, negative) {
  const n = positive + negative;
  if (!n) return 1;
  if (n > 200) {
    const z = Math.max(0, (Math.abs(positive - negative) - 1) / Math.sqrt(n));
    const t = 1 / (1 + 0.3275911 * z / Math.sqrt(2));
    const polynomial = (((((1.061405429 * t - 1.453152027) * t) + 1.421413741) * t - 0.284496736) * t + 0.254829592) * t;
    return Math.max(0,Math.min(1,polynomial*Math.exp(-(z*z)/2)));
  }
  const limit = Math.min(positive, negative);
  let probability = 2 ** -n, cumulative = probability;
  for (let count = 1; count <= limit; count += 1) {
    probability *= (n - count + 1) / count;
    cumulative += probability;
  }
  return Math.min(1, 2 * cumulative);
}

function pairwiseComparisons(completeCases) {
  if (!completeCases.length) return [];
  const comparisons = [];
  for (let first = 0; first < 5; first += 1) {
    for (let second = first + 1; second < 6; second += 1) {
      const differences = completeCases.map((row) => row[first] - row[second]);
      const positive = differences.filter((value) => value > 0).length;
      const negative = differences.filter((value) => value < 0).length;
      const meanDifference = differences.reduce((total, value) => total + value, 0) / differences.length;
      comparisons.push({ first: first + 1, second: second + 1, n: positive + negative, meanDifference,
        direction: meanDifference > 0 ? `${first + 1} > ${second + 1}` : meanDifference < 0 ? `${first + 1} < ${second + 1}` : "Equal",
        p: exactSignP(positive, negative), adjustedP: null });
    }
  }
  const ordered = [...comparisons].sort((a, b) => a.p - b.p);
  let previous = 0;
  ordered.forEach((item, index) => {
    item.adjustedP = Math.min(1, Math.max(previous, item.p * (ordered.length - index)));
    previous = item.adjustedP;
  });
  return comparisons;
}

export function buildZoneProfile(headers, rows, parameter) {
  const columns = Array.from({ length: 6 }, (_, index) => headers.findIndex((header) => String(header ?? "").trim().toLowerCase() === `${parameter}_${index + 1}`.toLowerCase()));
  if (columns.filter((index) => index >= 0).length < 2) throw new Error(`${parameter} does not contain at least two Zone columns.`);
  const valuesByZone = columns.map((column) => column < 0 ? [] : rows.filter((row) => numeric(row[column])).map((row) => Number(row[column])));
  const allValues = valuesByZone.flat();
  const overall = summary(allValues);
  if(overall.n<2)throw new Error(`${parameter} has fewer than two numeric Zone values in the selected data.`);
  const zones = valuesByZone.map((values, index) => {
    const stats = summary(values);
    return { zone: index + 1, ...stats, delta: stats.mean === null ? null : stats.mean - overall.mean,
      standardizedDelta: stats.mean === null || !overall.sigma ? null : (stats.mean - overall.mean) / overall.sigma };
  });
  const completeCases = rows.map((row) => columns.map((column) => column >= 0 && numeric(row[column]) ? Number(row[column]) : null))
    .filter((values) => values.every((value) => value !== null));
  const means = zones.map((zone) => zone.mean);
  const validMeans = means.map((mean, index) => ({ x: index + 1, y: mean })).filter((item) => item.y !== null);
  const meanX = validMeans.reduce((total, item) => total + item.x, 0) / (validMeans.length || 1);
  const meanY = validMeans.reduce((total, item) => total + item.y, 0) / (validMeans.length || 1);
  const denominator = validMeans.reduce((total, item) => total + (item.x - meanX) ** 2, 0);
  const slope = denominator ? validMeans.reduce((total, item) => total + (item.x - meanX) * (item.y - meanY), 0) / denominator : 0;
  return { parameter, rowCount: rows.length, overall, zones, completeCases: completeCases.length,
    friedman: friedman(completeCases), pairwise: pairwiseComparisons(completeCases), slope,
    direction: slope > 0 ? "Rising from Zone 1 to 6" : slope < 0 ? "Falling from Zone 1 to 6" : "Flat" };
}
