export function text(value) {
  return String(value ?? "").trim();
}

export function isNumeric(value) {
  return value !== "" && value !== null && value !== undefined && Number.isFinite(Number(value));
}

export function number(value) {
  return Number(value);
}

export function headerIndex(headers, name) {
  const wanted = text(name).toLowerCase();
  return headers.findIndex((header) => text(header).toLowerCase() === wanted);
}

export function canonicalizeSourceHeaders(headers) {
  return headers.map((header) => canonicalizeSourceHeader(header));
}

export function combineSourceHeaderRows(parentHeaders, childHeaders) {
  const width = Math.max(parentHeaders.length, childHeaders.length);
  let parent = "";
  return Array.from({ length: width }, (_, index) => {
    const parentCell = cleanHeader(parentHeaders[index]);
    const childCell = cleanHeader(childHeaders[index]);
    if (parentCell) parent = parentCell;
    const zone = zoneToken(childCell);
    if (zone && parent) return canonicalizeSourceHeader(`${parent}_${zone}`);
    return canonicalizeSourceHeader(childCell || parentCell);
  });
}

function canonicalizeSourceHeader(value) {
  const cleaned = cleanHeader(value);
  const key = cleaned.toLowerCase().replace(/[\s._-]+/g, "");
  if (key === "lot") return "Lot";
  if (key === "n" || key === "batchn" || key === "batchnumber") return "N";
  if (key === "type") return "Type";
  const zoneHeader = parseZoneHeader(cleaned);
  if (!zoneHeader) return cleaned;
  return `${zoneHeader.parameter}_${zoneHeader.zone}`;
}

function cleanHeader(value) {
  return text(value)
    .normalize("NFKC")
    .replace(/\u00a0/g, " ")
    .replace(/[\u200b-\u200d\ufeff]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function zoneToken(value) {
  const match = cleanHeader(value).match(/^(?:zone\s*|z\s*)?([1-6])[.):]?$/i);
  return match ? Number(match[1]) : null;
}

function parseZoneHeader(value) {
  let match = value.match(/^(?:zone\s*|z\s*)([1-6])\s*[-_.:/\\ ]+\s*(.+)$/i);
  if (match) {
    const parameter = cleanParameterName(match[2]);
    return /[A-Za-z]/.test(parameter) ? { parameter, zone: Number(match[1]) } : null;
  }
  match = value.match(/^(.+?)\s*(?:[-_.:/\\ ]+(?:zone\s*|z\s*)?|(?:zone\s*|z\s*))([1-6])\s*[)\]]?$/i);
  if (!match) return null;
  const parameter = cleanParameterName(match[1]);
  return /[A-Za-z]/.test(parameter) ? { parameter, zone: Number(match[2]) } : null;
}

function cleanParameterName(value) {
  return cleanHeader(value).replace(/[\s._:/\\-]+$/g, "");
}

export function zoneColumns(headers, baseName) {
  const columns = Array.from({ length: 6 }, (_, index) => headerIndex(headers, `${baseName}_${index + 1}`));
  return columns.some((index) => index >= 0) ? columns : null;
}

export function getRegionalParameters(headers) {
  const found = [];
  for (const header of headers) {
    const match = text(header).match(/^(.*)_([1-6])$/);
    if (!match) continue;
    if (zoneColumns(headers, match[1]) && !found.includes(match[1])) found.push(match[1]);
  }
  return found.sort((a, b) => a.localeCompare(b));
}

export function hasZonedHeaderGroup(headers) {
  return getRegionalParameters(headers).some((parameter) =>
    zoneColumns(headers, parameter).filter((index) => index >= 0).length >= 2
  );
}

export function findZonedHeaderRow(rows) {
  for (let index = 0; index < rows.length; index += 1) {
    const headers = canonicalizeSourceHeaders(rows[index]);
    if (hasZonedHeaderGroup(headers)) return { index, headers };
  }
  return null;
}

export function getLotValues(rows, lotColumn) {
  const lots = new Set();
  for (const row of rows) {
    const lot = text(row[lotColumn]);
    if (lot) lots.add(lot);
  }
  return [...lots].sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
}

export function meanAndSigma(values) {
  if (!values.length) return { n: 0, mean: null, sigma: null, min: null, max: null };
  let mean = 0;
  let m2 = 0;
  let min = values[0];
  let max = values[0];
  values.forEach((value, index) => {
    const delta = value - mean;
    mean += delta / (index + 1);
    m2 += delta * (value - mean);
    min = Math.min(min, value);
    max = Math.max(max, value);
  });
  return { n: values.length, mean, sigma: values.length > 1 ? Math.sqrt(m2 / (values.length - 1)) : 0, min, max };
}

export function median(values) {
  if (!values.length) return null;
  const sorted = [...values].sort((a, b) => a - b);
  const middle = Math.floor(sorted.length / 2);
  return sorted.length % 2 ? sorted[middle] : (sorted[middle - 1] + sorted[middle]) / 2;
}

export function robustHistory(values) {
  if (values.length < 5) return { ...meanAndSigma(values), excluded: 0 };
  const center = median(values);
  const mad = median(values.map((value) => Math.abs(value - center)));
  if (!mad) return { ...meanAndSigma(values), excluded: 0 };
  const cutoff = 3.5 * 1.4826 * mad;
  const retained = values.filter((value) => Math.abs(value - center) <= cutoff);
  const population = retained.length >= 3 ? retained : values;
  return { ...meanAndSigma(population), excluded: values.length - population.length };
}

function niceBinWidth(rawWidth) {
  if (!Number.isFinite(rawWidth) || rawWidth <= 0) return 1;
  const exponent = Math.floor(Math.log10(rawWidth));
  const fraction = rawWidth / 10 ** exponent;
  const rounded = fraction <= 1 ? 1 : fraction <= 2 ? 2 : fraction <= 5 ? 5 : 10;
  return rounded * 10 ** exponent;
}

export function gaussianFit(values, userBinWidth, userStart, userEnd) {
  return gaussianFitWithOptions(values, userBinWidth, userStart, userEnd, {});
}

export function gaussianFitWithOptions(values, userBinWidth, userStart, userEnd, options = {}) {
  const method = options.method || "standard";
  const lowerLimit = method === "nacl-truncated" ? Number(options.lowerLimit ?? 1) : null;
  if (method !== "standard" && method !== "nacl-truncated") throw new Error("Select a valid Gaussian fit method.");
  if (method === "nacl-truncated" && !(lowerLimit > 0)) throw new Error("The lower detection limit must be positive.");
  if (values.length < 2) throw new Error("At least two numeric visible values are required.");
  const recordedValues = method === "nacl-truncated" ? values.filter((value) => value >= lowerLimit) : values;
  if (recordedValues.length < 2) throw new Error("At least two recorded values are required above the detection limit.");
  const baseStats = meanAndSigma(recordedValues);
  const startInput = Number(userStart);
  const endInput = Number(userEnd);
  let binWidth = Number(userBinWidth);
  if (!Number.isFinite(binWidth) || binWidth <= 0) {
    const rawWidth = (baseStats.max - baseStats.min) / (1 + Math.log2(values.length));
    binWidth = niceBinWidth(rawWidth || 1);
  }
  let start = Number.isFinite(startInput)
    ? startInput
    : method === "nacl-truncated"
      ? lowerLimit
      : Math.floor(baseStats.min / binWidth) * binWidth;
  if (method === "nacl-truncated") start = Math.max(start, lowerLimit);
  let end = Number.isFinite(endInput) ? endInput : Math.ceil(baseStats.max / binWidth) * binWidth;
  if (end <= start) end = start + binWidth;
  const included = recordedValues.filter((value) => value >= start && value <= end);
  if (included.length < 2) throw new Error("The selected fit range leaves fewer than two values.");
  const standardStats = meanAndSigma(included);
  const truncatedStats = method === "nacl-truncated" ? fitLowerTruncatedNormal(included, start) : null;
  const stats = truncatedStats
    ? { ...standardStats, mean: truncatedStats.mean, sigma: truncatedStats.sigma }
    : standardStats;
  const binCount = Math.ceil((end - start) / binWidth);
  if (binCount > 300) throw new Error("Use a larger bin width or a narrower range (maximum 300 bins).");
  const bins = Array.from({ length: binCount }, (_, index) => {
    const lower = start + index * binWidth;
    const upper = Math.min(end, lower + binWidth);
    return { lower, upper, center: (lower + upper) / 2, observed: 0, gaussian: 0 };
  });
  for (const value of included) {
    const index = Math.min(binCount - 1, Math.floor((value - start) / binWidth));
    bins[index].observed += 1;
  }
  if (stats.sigma > 0) {
    for (const bin of bins) {
      if (method === "nacl-truncated") {
        bin.gaussian = expectedTruncatedBinCount(
          bin.lower,
          bin.upper,
          included.length,
          stats.mean,
          stats.sigma,
          start,
          end
        );
      } else {
        const z = (bin.center - stats.mean) / stats.sigma;
        bin.gaussian = included.length * (bin.upper - bin.lower) * Math.exp(-0.5 * z * z) /
          (stats.sigma * Math.sqrt(2 * Math.PI));
      }
    }
  }
  const sse = bins.reduce((total, bin) => total + (bin.observed - bin.gaussian) ** 2, 0);
  const belowLimit = method === "nacl-truncated" ? values.filter((value) => value < lowerLimit).length : 0;
  const excludedByRange = recordedValues.length - included.length;
  return {
    ...stats,
    visibleN: values.length,
    excluded: values.length - included.length,
    excludedByRange,
    belowLimit,
    binWidth,
    start,
    end,
    bins,
    sse,
    method,
    lowerLimit,
    correctionApplied: Boolean(truncatedStats?.correctionApplied)
  };
}

export function collectParameterValues(rows, zoneIndexes, includedZones) {
  const values = [];
  for (const row of rows) {
    includedZones.forEach((zone) => {
      const columnIndex = zoneIndexes[zone - 1];
      if (columnIndex < 0) return;
      const value = row[columnIndex];
      if (isNumeric(value)) values.push(number(value));
    });
  }
  return values;
}

export function collectParameterRecords(headers, rows, parameter, includedZones) {
  const lotColumn = headerIndex(headers, "Lot");
  const batchColumn = headerIndex(headers, "N");
  const zones = zoneColumns(headers, parameter);
  if (!zones) throw new Error(`${parameter} has no readable Zone columns.`);
  const records = [];
  for (const row of rows) {
    for (const zone of includedZones) {
      const columnIndex = zones[zone - 1];
      if (columnIndex < 0) continue;
      const value = row[columnIndex];
      if (isNumeric(value)) {
        records.push({
          value: number(value),
          lot: lotColumn >= 0 ? text(row[lotColumn]) : "",
          batch: batchColumn >= 0 ? text(row[batchColumn]) : "",
          zone
        });
      }
    }
  }
  return records;
}

export function buildOriginExport(headers, rows, parameter, includedZones) {
  const lotColumn = headerIndex(headers, "Lot");
  const batchColumn = headerIndex(headers, "N");
  if (lotColumn < 0 || batchColumn < 0) throw new Error("The source sheet must contain Lot and N columns.");
  const zones = zoneColumns(headers, parameter);
  if (!zones) throw new Error(`${parameter} has no readable Zone columns.`);
  const output = [["Lot", "Batch N", "Parameter", "Zone", "Value"]];
  for (const row of rows) {
    for (const zone of includedZones) {
      const columnIndex = zones[zone - 1];
      if (columnIndex < 0) continue;
      const value = row[columnIndex];
      if (isNumeric(value)) output.push([row[lotColumn], row[batchColumn], parameter, `Zone ${zone}`, number(value)]);
    }
  }
  return output;
}

export function buildOriginExportWide(headers, rows, parameter, includedZones) {
  const lotColumn = headerIndex(headers, "Lot");
  const batchColumn = headerIndex(headers, "N");
  if (lotColumn < 0 || batchColumn < 0) throw new Error("The source sheet must contain Lot and N columns.");
  const zones = zoneColumns(headers, parameter);
  if (!zones) throw new Error(`${parameter} has no readable Zone columns.`);
  const output = [["Lot", "Batch N", ...includedZones.map((zone) => `${parameter}_${zone}`)]];
  for (const row of rows) {
    const values = includedZones.map((zone) => {
      const columnIndex = zones[zone - 1];
      if (columnIndex < 0) return "";
      const value = row[columnIndex];
      return isNumeric(value) ? number(value) : "";
    });
    if (values.some((value) => value !== "")) output.push([row[lotColumn], row[batchColumn], ...values]);
  }
  return output;
}

function normalPdf(z) {
  return Math.exp(-0.5 * z * z) / Math.sqrt(2 * Math.PI);
}

function standardNormalCdf(z) {
  if (z <= -8) return 0;
  if (z >= 8) return 1;
  const absoluteZ = Math.abs(z);
  const t = 1 / (1 + 0.2316419 * absoluteZ);
  const upperTail = normalPdf(absoluteZ) * (
    0.319381530 * t -
    0.356563782 * t ** 2 +
    1.781477937 * t ** 3 -
    1.821255978 * t ** 4 +
    1.330274429 * t ** 5
  );
  return z >= 0 ? 1 - upperTail : upperTail;
}

function standardNormalSurvival(z) {
  if (z > 5) {
    const tailSeries = 1 / z - 1 / z ** 3 + 3 / z ** 5 - 15 / z ** 7 + 105 / z ** 9;
    return normalPdf(z) * tailSeries;
  }
  if (z < -8) return 1;
  return 1 - standardNormalCdf(z);
}

function upperTailInverseMills(z) {
  if (z > 5) return z + 1 / z - 2 / z ** 3 + 10 / z ** 5 - 74 / z ** 7;
  return normalPdf(z) / Math.max(standardNormalSurvival(z), 1e-12);
}

function truncatedMomentRatio(alpha) {
  const inverseMills = upperTailInverseMills(alpha);
  const meanDistance = inverseMills - alpha;
  const varianceFactor = 1 + alpha * inverseMills - inverseMills ** 2;
  return varianceFactor <= 0 ? 1.000000001 : meanDistance ** 2 / varianceFactor;
}

function fitLowerTruncatedNormal(values, lowerLimit) {
  const observedMean = values.reduce((sum, value) => sum + value, 0) / values.length;
  const observedVariance = values.reduce((sum, value) => sum + (value - observedMean) ** 2, 0) / values.length;
  if (!(observedVariance > 0)) throw new Error("The recorded values have no measurable variation.");
  const sampleSigma = Math.sqrt(observedVariance * values.length / (values.length - 1));
  const distanceAboveLimit = observedMean - lowerLimit;
  if (!(distanceAboveLimit > 0)) throw new Error("The recorded mean must be above the detection limit.");
  const targetRatio = distanceAboveLimit ** 2 / observedVariance;
  let alphaLow = -12;
  let alphaHigh = 8;
  if (targetRatio >= truncatedMomentRatio(alphaLow)) {
    return { mean: observedMean, sigma: sampleSigma, correctionApplied: false };
  }
  let alpha;
  if (targetRatio <= truncatedMomentRatio(alphaHigh)) {
    alpha = alphaHigh;
  } else {
    for (let iteration = 0; iteration < 120; iteration += 1) {
      const alphaMid = 0.5 * (alphaLow + alphaHigh);
      if (truncatedMomentRatio(alphaMid) > targetRatio) alphaLow = alphaMid;
      else alphaHigh = alphaMid;
    }
    alpha = 0.5 * (alphaLow + alphaHigh);
  }
  const denominator = upperTailInverseMills(alpha) - alpha;
  if (!(denominator > 0)) throw new Error("The truncated Gaussian fit did not converge.");
  const sigma = distanceAboveLimit / denominator;
  return { mean: lowerLimit - alpha * sigma, sigma, correctionApplied: true };
}

function expectedTruncatedBinCount(lower, upper, n, mean, sigma, lowerLimit, upperRange) {
  if (!(sigma > 0)) return 0;
  const effectiveLower = Math.max(lower, lowerLimit);
  const effectiveUpper = Math.min(upper, upperRange);
  if (effectiveUpper <= effectiveLower) return 0;
  const normalizer = standardNormalSurvival((lowerLimit - mean) / sigma) -
    standardNormalSurvival((upperRange - mean) / sigma);
  if (!(normalizer > 0)) return 0;
  const binProbability = standardNormalSurvival((effectiveLower - mean) / sigma) -
    standardNormalSurvival((effectiveUpper - mean) / sigma);
  return n * binProbability / normalizer;
}

export function buildLotAssessment(headers, rows, parameter, selectedLot, referenceMode, manualMu, manualSigma, monitorLimit, outlierLimit) {
  const lotColumn = headerIndex(headers, "Lot");
  const batchColumn = headerIndex(headers, "N");
  if (lotColumn < 0 || batchColumn < 0) throw new Error("The source sheet must contain Lot and N columns.");
  const zones = zoneColumns(headers, parameter);
  if (!zones) throw new Error(`${parameter} has no readable Zone columns.`);
  if (!(monitorLimit > 0 && outlierLimit > monitorLimit)) throw new Error("The out-of-range limit must be larger than the monitor limit.");

  const references = zones.map((zoneIndex) => {
    if (zoneIndex < 0) return { n: 0, mean: null, sigma: null, excluded: 0 };
    if (referenceMode === "manual") {
      if (!(Number.isFinite(manualMu) && Number.isFinite(manualSigma) && manualSigma > 0)) {
        throw new Error("Enter a positive manual mu and sigma.");
      }
      return { n: 0, mean: manualMu, sigma: manualSigma, excluded: 0 };
    }
    const history = rows.filter((row) => text(row[lotColumn]) !== selectedLot && isNumeric(row[zoneIndex])).map((row) => number(row[zoneIndex]));
    return robustHistory(history);
  });

  const selectedRows = rows.filter((row) => text(row[lotColumn]) === selectedLot);
  if (!selectedRows.length) throw new Error("No visible rows were found for this Lot.");
  let monitorCount = 0;
  let outOfRangeCount = 0;
  let noHistoryCount = 0;
  const grid = selectedRows.map((row) => {
    const values = zones.map((zoneIndex) => (zoneIndex >= 0 && isNumeric(row[zoneIndex]) ? number(row[zoneIndex]) : null));
    const states = values.map((value, index) => {
      const reference = references[index];
      if (value === null || (referenceMode === "history" && !reference.n) || !(reference.sigma > 0)) {
        noHistoryCount += value === null ? 0 : 1;
        return "NO HISTORY";
      }
      const score = Math.abs((value - reference.mean) / reference.sigma);
      if (score > outlierLimit) {
        outOfRangeCount += 1;
        return "OUT OF RANGE";
      }
      if (score > monitorLimit) {
        monitorCount += 1;
        return "CHECK";
      }
      return "OK";
    });
    return { batch: row[batchColumn], values, states };
  });
  const overall = outOfRangeCount ? "OUT OF RANGE" : monitorCount ? "CHECK" : "OK";
  const availableZones = zones.map((zoneIndex, index) => zoneIndex >= 0 ? index + 1 : null).filter(Boolean);
  return { references, grid, availableZones, monitorCount, outOfRangeCount, noHistoryCount, overall };
}

export function buildCorrelation(headers, rows, xParameter, yParameter, scope, outlierMethod, removalPct) {
  const lotColumn = headerIndex(headers, "Lot");
  const batchColumn = headerIndex(headers, "N");
  const xZones = zoneColumns(headers, xParameter);
  const yZones = zoneColumns(headers, yParameter);
  if (!xZones || !yZones) throw new Error("Both parameters need at least one readable Zone column.");
  const selectedZones = scope === "all" ? [1, 2, 3, 4, 5, 6] : [Number(scope)];
  const pairs = [];
  for (const row of rows) {
    for (const zone of selectedZones) {
      const xColumn = xZones[zone - 1];
      const yColumn = yZones[zone - 1];
      if (xColumn < 0 || yColumn < 0) continue;
      const x = row[xColumn];
      const y = row[yColumn];
      if (isNumeric(x) && isNumeric(y)) {
        pairs.push({ lot: lotColumn >= 0 ? row[lotColumn] : "", batch: batchColumn >= 0 ? row[batchColumn] : "", zone, x: number(x), y: number(y), included: true });
      }
    }
  }
  if (pairs.length < 3) throw new Error("At least three visible numeric pairs are required.");
  if (outlierMethod === "ratio") removeRatioExtremes(pairs, removalPct);
  const included = pairs.filter((pair) => pair.included);
  if (included.length < 3) throw new Error("Too few pairs remain after extreme-pair removal.");
  return { pairs, included, ...pearsonAndLine(included) };
}

function removeRatioExtremes(pairs, removalPct) {
  const maxRemove = Math.floor(pairs.length * Math.max(0, Math.min(50, removalPct)) / 100);
  if (!maxRemove) return;
  const ratioPairs = pairs.filter((pair) => pair.x !== 0).map((pair) => ({ pair, ratio: pair.y / pair.x }));
  if (ratioPairs.length < 3) return;
  const center = median(ratioPairs.map((item) => item.ratio));
  ratioPairs.sort((a, b) => Math.abs(b.ratio - center) - Math.abs(a.ratio - center));
  ratioPairs.slice(0, Math.min(maxRemove, ratioPairs.length - 3)).forEach((item) => { item.pair.included = false; });
}

function pearsonAndLine(pairs) {
  const n = pairs.length;
  const sums = pairs.reduce((acc, pair) => ({
    x: acc.x + pair.x,
    y: acc.y + pair.y,
    xx: acc.xx + pair.x * pair.x,
    yy: acc.yy + pair.y * pair.y,
    xy: acc.xy + pair.x * pair.y
  }), { x: 0, y: 0, xx: 0, yy: 0, xy: 0 });
  const vx = sums.xx - sums.x * sums.x / n;
  const vy = sums.yy - sums.y * sums.y / n;
  const cov = sums.xy - sums.x * sums.y / n;
  const r = vx > 0 && vy > 0 ? cov / Math.sqrt(vx * vy) : null;
  const slope = vx > 0 ? cov / vx : null;
  const intercept = slope !== null ? (sums.y - slope * sums.x) / n : null;
  return { rawN: n, r, slope, intercept };
}
