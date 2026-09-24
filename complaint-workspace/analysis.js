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

export function percentile(values, probability) {
  if (!values.length) return null;
  const sorted = [...values].sort((a, b) => a - b);
  if (probability <= 0) return sorted[0];
  if (probability >= 1) return sorted[sorted.length - 1];
  const position = (sorted.length - 1) * probability;
  const lowIndex = Math.floor(position);
  const highIndex = Math.min(sorted.length - 1, lowIndex + 1);
  return sorted[lowIndex] + (position - lowIndex) * (sorted[highIndex] - sorted[lowIndex]);
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

export function recommendGaussianSettings(values, options = {}) {
  if (values.length < 2) throw new Error("At least two numeric visible values are required.");
  const method = options.method || "standard";
  const lowerLimit = method === "nacl-truncated" ? Number(options.lowerLimit ?? 1) : null;
  const recordedValues = method === "nacl-truncated" ? values.filter((value) => value >= lowerLimit) : values;
  if (recordedValues.length < 2) throw new Error("At least two recorded values are required above the detection limit.");
  const stats = meanAndSigma(recordedValues);
  const q1 = percentile(recordedValues, 0.25);
  const q3 = percentile(recordedValues, 0.75);
  const iqr = q3 - q1;
  const range = stats.max - stats.min;
  let rawWidth = iqr > 0 ? 2 * iqr / Math.cbrt(recordedValues.length) : 0;
  if (!(rawWidth > 0) && stats.sigma > 0) rawWidth = 3.5 * stats.sigma / Math.cbrt(recordedValues.length);
  if (!(rawWidth > 0) && range > 0) rawWidth = range / 12;
  if (!(rawWidth > 0)) rawWidth = 1;

  let binWidth = niceBinWidth(rawWidth);
  if (range > 0) {
    const binCount = Math.ceil(range / binWidth);
    if (binCount < 8) binWidth = niceBinWidth(range / 12);
    if (binCount > 80) binWidth = niceBinWidth(range / 40);
  }
  if (!(binWidth > 0)) binWidth = 1;

  const start = method === "nacl-truncated"
    ? lowerLimit
    : Math.floor(stats.min / binWidth) * binWidth;
  let end = method === "nacl-truncated"
    ? start + binWidth * Math.ceil((stats.max - start) / binWidth)
    : Math.ceil(stats.max / binWidth) * binWidth;
  if (end <= start) end = start + binWidth;
  return { n: recordedValues.length, q1, q3, iqr, binWidth, start, end };
}

export function gaussianFit(values, userBinWidth, userStart, userEnd) {
  return gaussianFitWithOptions(values, userBinWidth, userStart, userEnd, {});
}

export function fitHuberGaussian(values) {
  if (values.length < 5) throw new Error("Robust Gaussian requires at least five numeric observations.");
  if (values.some(value => !Number.isFinite(value))) throw new Error("Robust Gaussian requires finite numeric values.");
  const c = 1.345, tolerance = 1e-8;
  let mu = median(values), sigma = median(values.map(value => Math.abs(value - mu))) / 0.674489750196082;
  if (!(sigma > 0)) sigma = meanAndSigma(values).sigma;
  // Integral of the standard normal over [-c,c], evaluated by its convergent series.
  let term = 1, sum = 1;
  for (let k = 1; k <= 40; k++) { term *= -c * c / (2 * k); sum += term / (2 * k + 1); }
  const probability = Math.sqrt(2 / Math.PI) * c * sum;
  const phi = Math.exp(-0.5 * c * c) / Math.sqrt(2 * Math.PI);
  const gamma = probability - 2 * c * phi + c * c * (1 - probability);
  for (let iterations = 1; iterations <= 100; iterations++) {
    if (!(sigma > 0) || !Number.isFinite(sigma)) break;
    const lower = mu - c * sigma, upper = mu + c * sigma;
    const newMu = values.reduce((total, value) => total + Math.max(lower, Math.min(upper, value)), 0) / values.length;
    const numerator = values.reduce((total, value) => total + (Math.abs((value - mu) / sigma) <= c ? (value - newMu) ** 2 : (sigma * c) ** 2), 0);
    const newSigma = Math.sqrt(numerator / ((values.length - 1) * gamma));
    if (!(newSigma > 0) || !Number.isFinite(newSigma) || !Number.isFinite(newMu)) break;
    if (Math.abs(newSigma - sigma) <= newSigma * tolerance && Math.abs(newMu - mu) <= newSigma * tolerance) {
      return { mean: newMu, sigma: newSigma, iterations, huberC: c };
    }
    mu = newMu; sigma = newSigma;
  }
  throw new Error("Robust fit failed: invalid scale or no convergence within 100 iterations. No fit or extreme export was created.");
}

export function gaussianExtremeSnapshot(records, fit, multiplier = 3) {
  if (!Number.isFinite(multiplier) || multiplier <= 0) throw new Error("Extreme sigma multiplier must be greater than zero.");
  if (!Number.isFinite(fit.mean) || !Number.isFinite(fit.sigma) || fit.sigma < 0) throw new Error("Run a successful Gaussian fit first.");
  const lower = fit.mean - multiplier * fit.sigma, upper = fit.mean + multiplier * fit.sigma;
  if (!Number.isFinite(lower) || !Number.isFinite(upper)) throw new Error("Extreme boundaries exceed the supported numeric range.");
  const extremes = records.filter(record => record.value < lower || record.value > upper).map(record => ({
    ...record, side: record.value < lower ? "Low" : "High", zScore: fit.sigma > 0 ? (record.value - fit.mean) / fit.sigma : null,
    mean: fit.mean, sigma: fit.sigma, multiplier
  }));
  return { lower, upper, multiplier, records: extremes, lowCount: extremes.filter(record => record.side === "Low").length, highCount: extremes.filter(record => record.side === "High").length };
}

export function histogramBinIndex(value, start, width, count) {
  const quotient = (value - start) / width;
  const nearest = Math.round(quotient);
  const tolerance = 16 * Number.EPSILON * Math.max(1, Math.abs(quotient));
  const stable = Math.abs(quotient - nearest) <= tolerance ? nearest : quotient;
  return Math.max(0, Math.min(count - 1, Math.floor(stable)));
}

export function gaussianFitWithOptions(values, userBinWidth, userStart, userEnd, options = {}) {
  const method = options.method || "standard";
  const robust = method === "robust-huber";
  values = values.filter(value => typeof value === "number" && Number.isFinite(value));
  const lowerLimit = method === "nacl-truncated" ? Number(options.lowerLimit ?? 1) : null;
  if (method !== "standard" && method !== "nacl-truncated" && !robust) throw new Error("Select a valid Gaussian fit method.");
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
  if (robust && Number.isFinite(startInput) && Number.isFinite(endInput) && end <= start) throw new Error("Histogram end must be larger than start.");
  if (end <= start) end = start + binWidth;
  const included = recordedValues.filter((value) => value >= start && value <= end);
  if (!robust && included.length < 2) throw new Error("The selected fit range leaves fewer than two values.");
  const population = robust ? recordedValues : included;
  const standardStats = meanAndSigma(population);
  const truncatedStats = method === "nacl-truncated" ? fitLowerTruncatedNormal(included, start) : null;
  const stats = robust ? { ...standardStats, ...fitHuberGaussian(population) } : truncatedStats
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
    const index = histogramBinIndex(value, start, binWidth, binCount);
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
        bin.gaussian = population.length * (bin.upper - bin.lower) * Math.exp(-0.5 * z * z) /
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
    excluded: values.length - population.length,
    excludedByRange: robust ? 0 : excludedByRange,
    outsideHistogram: excludedByRange,
    histogramN: included.length,
    belowLimit,
    binWidth,
    start,
    end,
    bins,
    sse,
    low25: percentile(population, 0.025),
    high25: percentile(population, 0.975),
    low15: percentile(population, 0.15),
    high15: percentile(population, 0.85),
    method,
    lowerLimit,
    correctionApplied: Boolean(truncatedStats?.correctionApplied)
  };
}

export function gaussianSigmaCounts(values, mean, sigma) {
  values=values.filter(value=>typeof value==="number"&&Number.isFinite(value));
  if(!Number.isFinite(mean)||!Number.isFinite(sigma)||sigma<0)throw new Error("Gaussian mean and sigma are required for sigma counts.");
  const count=(test)=>values.reduce((total,value)=>total+(test(value)?1:0),0);
  const lower2=mean-2*sigma,lower1=mean-sigma,upper1=mean+sigma,upper2=mean+2*sigma;
  return {
    n:values.length,mu:mean,sigma,lower2,lower1,upper1,upper2,
    below2:count(value=>value<lower2),
    lower2to1:count(value=>value>=lower2&&value<lower1),
    lower1toMean:count(value=>value>=lower1&&value<mean),
    meanToUpper1:count(value=>value>=mean&&value<=upper1),
    upper1to2:count(value=>value>upper1&&value<=upper2),
    above2:count(value=>value>upper2),
    within1:count(value=>value>=lower1&&value<=upper1),
    within2:count(value=>value>=lower2&&value<=upper2)
  };
}

export function gaussianOverlapCoefficient(firstMean, firstSigma, secondMean, secondSigma) {
  if(![firstMean,firstSigma,secondMean,secondSigma].every(Number.isFinite)||firstSigma<0||secondSigma<0)throw new Error("Two Gaussian means and non-negative sigmas are required.");
  if(firstSigma===0||secondSigma===0)return firstSigma===secondSigma&&firstMean===secondMean?1:0;
  const start=Math.min(firstMean-6*firstSigma,secondMean-6*secondSigma);
  const end=Math.max(firstMean+6*firstSigma,secondMean+6*secondSigma);
  const steps=2400,width=(end-start)/steps;
  const density=(value,mean,sigma)=>Math.exp(-0.5*((value-mean)/sigma)**2)/(sigma*Math.sqrt(2*Math.PI));
  let area=0;
  for(let index=0;index<=steps;index+=1){
    const value=start+index*width;
    const overlap=Math.min(density(value,firstMean,firstSigma),density(value,secondMean,secondSigma));
    area+=(index===0||index===steps?0.5:1)*overlap*width;
  }
  return Math.max(0,Math.min(1,area));
}

export function getTrendParameters(headers) {
  const regional = getRegionalParameters(headers).filter((parameter) =>
    zoneColumns(headers, parameter)?.every((index) => index >= 0)
  );
  for (const scalarName of ["Visco.", "Water", "Temp.", "Humidity"]) {
    const index = headerIndex(headers, scalarName);
    if (index >= 0 && !regional.includes(headers[index])) regional.push(headers[index]);
  }
  return regional.sort((a, b) => a.localeCompare(b));
}

export function buildTrendDateLookup(table) {
  const lookup = new Map();
  const headerRowIndex = table.slice(0, 30).findIndex((row) =>
    rawHeaderIndex(row, "ChargenNr") >= 0 && rawHeaderIndex(row, "Nummer") >= 0 && rawHeaderIndex(row, "Probenzeit") >= 0
  );
  if (headerRowIndex < 0) return lookup;
  const headers = table[headerRowIndex];
  const lotColumn = rawHeaderIndex(headers, "ChargenNr");
  const batchColumn = rawHeaderIndex(headers, "Nummer");
  const typeColumn = rawHeaderIndex(headers, "Ziehart");
  const dateColumn = rawHeaderIndex(headers, "Probenzeit");
  const sampleTypeColumn = rawHeaderIndex(headers, "Probentyp");
  for (const row of table.slice(headerRowIndex + 1)) {
    if (sampleTypeColumn >= 0 && /^proben?$/i.test(text(row[sampleTypeColumn]))) continue;
    const lot = text(row[lotColumn]);
    const batch = text(row[batchColumn]);
    const type = typeColumn >= 0 ? text(row[typeColumn]) : "";
    const date = trendDateValue(row[dateColumn]);
    if (!lot || !batch || !Number.isFinite(date)) continue;
    const key = trendKey(lot, batch, type);
    const existing = lookup.get(key);
    if (!Number.isFinite(existing) || date < existing) lookup.set(key, date);
  }
  return lookup;
}

export function filterLotsByPeriod(headers, rows, dateLookup = new Map(), options = {}) {
  if (!options.startDate && !options.endDate) throw new Error("Choose a start or end date to limit the period.");
  const start = trendDateBound(options.startDate, false);
  const end = trendDateBound(options.endDate, true);
  if (Number.isFinite(start) && Number.isFinite(end) && end < start) throw new Error("End date must be on or after start date.");
  const lotColumn = headerIndex(headers, "Lot");
  if (lotColumn < 0) throw new Error("The source sheet needs a Lot column for period filtering.");
  const batchColumn = headerIndex(headers, "N");
  const typeColumn = headerIndex(headers, "Type");
  const dateColumns = ["Probenzeit", "Production Date", "Date", "Datum"].map(name => headerIndex(headers, name)).filter(index => index >= 0);
  const allLots = new Set();
  const datedLots = new Set();
  const matchingLots = new Set();
  for (const row of rows) {
    const lot = text(row[lotColumn]);
    if (!lot) continue;
    allLots.add(lot);
    let date = dateLookup.get(trendKey(lot, row[batchColumn], row[typeColumn]));
    if (!Number.isFinite(date)) {
      date = dateColumns.map(column => trendDateValue(row[column])).find(Number.isFinite);
    }
    if (!Number.isFinite(date)) continue;
    datedLots.add(lot);
    if ((!Number.isFinite(start) || date >= start) && (!Number.isFinite(end) || date <= end)) matchingLots.add(lot);
  }
  return {
    rows: rows.filter(row => matchingLots.has(text(row[lotColumn]))),
    lotCount: matchingLots.size,
    undatedLotCount: [...allLots].filter(lot => !datedLots.has(lot)).length
  };
}

export function buildParameterTrend(headers, rows, parameter, dateLookup = new Map(), options = {}) {
  const lotColumn = headerIndex(headers, "Lot");
  const batchColumn = headerIndex(headers, "N");
  const typeColumn = headerIndex(headers, "Type");
  if (lotColumn < 0 || batchColumn < 0) throw new Error("The source sheet must contain Lot and N columns.");
  const regionalColumns = zoneColumns(headers, parameter);
  const regional = Boolean(regionalColumns?.every((index) => index >= 0));
  const scalarColumn = regional ? -1 : headerIndex(headers, parameter);
  if (!regional && scalarColumn < 0) throw new Error(`Parameter column is missing: ${parameter}`);
  const rollingWindow = Math.trunc(Number(options.rollingWindow ?? 5));
  if (rollingWindow < 1 || rollingWindow > 200) throw new Error("Use a rolling window between 1 and 200 batches.");
  const startDate = trendDateBound(options.startDate, false);
  const endDate = trendDateBound(options.endDate, true);
  if (Number.isFinite(startDate) && Number.isFinite(endDate) && endDate < startDate) {
    throw new Error("End date must be on or after start date.");
  }

  const batches = [];
  let fallbackCount = 0;
  for (const row of rows) {
    const lot = text(row[lotColumn]);
    const batch = text(row[batchColumn]);
    const type = typeColumn >= 0 ? text(row[typeColumn]) : "";
    let date = dateLookup.get(trendKey(lot, batch, type));
    let fallback = false;
    if (!Number.isFinite(date)) {
      date = trendDateFromLot(lot);
      fallback = Number.isFinite(date);
    }
    if (!Number.isFinite(date)) continue;
    if (Number.isFinite(startDate) && date < startDate) continue;
    if (Number.isFinite(endDate) && date > endDate) continue;
    const zoneValues = regional
      ? regionalColumns.map((column) => isNumeric(row[column]) ? number(row[column]) : null)
      : [];
    const values = regional ? zoneValues.filter(Number.isFinite) : isNumeric(row[scalarColumn]) ? [number(row[scalarColumn])] : [];
    if (!values.length) continue;
    const stats = meanAndSigma(values);
    batches.push({ date, lot, batch: row[batchColumn], type, mean: stats.mean, sigma: stats.sigma, n: stats.n, zones: zoneValues, fallback });
    if (fallback) fallbackCount += 1;
  }
  if (!batches.length) throw new Error("No visible parameter values remain in the selected date range.");

  batches.sort(compareTrendRows);
  let previousLot = "";
  let lotStart = 0;
  let rollingSum = 0;
  batches.forEach((item, index) => {
    const currentLot = trendIdentity(item.lot);
    if (index === 0 || currentLot !== previousLot) {
      rollingSum = 0;
      lotStart = index;
    }
    rollingSum += item.mean;
    if (index - lotStart + 1 > rollingWindow) rollingSum -= batches[index - rollingWindow].mean;
    item.rollingMean = rollingSum / Math.min(index - lotStart + 1, rollingWindow);
    previousLot = currentLot;
  });

  const lotGroups = new Map();
  for (const item of batches) {
    const key = trendIdentity(item.lot);
    if (!lotGroups.has(key)) lotGroups.set(key, { lot: item.lot, date: item.date, values: [] });
    const group = lotGroups.get(key);
    group.date = Math.min(group.date, item.date);
    if (regional) group.values.push(...item.zones.filter(Number.isFinite));
    else group.values.push(item.mean);
  }
  const lots = [...lotGroups.values()].map((group) => ({
    lot: group.lot,
    date: group.date,
    ...meanAndSigma(group.values)
  })).sort(compareTrendRows);
  const lotStats = meanAndSigma(lots.map((lot) => lot.mean));
  const batchStats = meanAndSigma(batches.map((batch) => batch.mean));
  const zoneBias = regional ? Array.from({ length: 6 }, (_, index) => {
    const values = batches.map((batch) => batch.zones[index]).filter(Number.isFinite);
    const biases = batches
      .filter((batch) => batch.n > 1 && Number.isFinite(batch.zones[index]))
      .map((batch) => batch.zones[index] - batch.mean);
    const valueStats = meanAndSigma(values);
    const biasStats = meanAndSigma(biases);
    const standardError = biasStats.n > 1 ? biasStats.sigma / Math.sqrt(biasStats.n) : 0;
    const ciLow = biasStats.n > 1 ? biasStats.mean - 1.96 * standardError : 0;
    const ciHigh = biasStats.n > 1 ? biasStats.mean + 1.96 * standardError : 0;
    const signal = biasStats.n < 3 ? "Insufficient"
      : ciLow > 0 ? "Consistently high"
        : ciHigh < 0 ? "Consistently low"
          : "No clear bias";
    return {
      zone: index + 1,
      n: biasStats.n,
      meanValue: valueStats.mean,
      meanBias: biasStats.mean,
      biasSigma: biasStats.sigma,
      ciLow,
      ciHigh,
      signal
    };
  }) : [];

  return {
    parameter,
    regional,
    rollingWindow,
    batches,
    lots,
    lotCenter: lotStats.mean,
    lotSigma: lotStats.sigma,
    batchCenter: batchStats.mean,
    batchSigma: batchStats.sigma,
    zoneBias,
    fallbackCount,
    lookupCount: dateLookup.size
  };
}

function rawHeaderIndex(headers, name) {
  const wanted = text(name).toLowerCase().replace(/[\s._-]+/g, "");
  return headers.findIndex((header) => text(header).toLowerCase().replace(/[\s._-]+/g, "") === wanted);
}

function trendIdentity(value) {
  const rendered = text(value);
  if (rendered && Number.isFinite(Number(rendered))) return String(Number(rendered));
  return rendered.toUpperCase();
}

function trendKey(lot, batch, type) {
  return `${trendIdentity(lot)}|${trendIdentity(batch)}|${trendIdentity(type)}`;
}

function trendDateValue(value) {
  if (value instanceof Date && Number.isFinite(value.getTime())) return value.getTime();
  if (typeof value === "number" && value > 0) return Date.UTC(1899, 11, 30) + value * 86400000;
  const parsed = Date.parse(text(value));
  return Number.isFinite(parsed) ? parsed : null;
}

function trendDateFromLot(lot) {
  const rendered = text(lot);
  if (!/^\d{2}/.test(rendered)) return null;
  const prefix = Number(rendered.slice(0, 2));
  const year = prefix <= 79 ? 2000 + prefix : 1900 + prefix;
  const yearStart = Date.UTC(year, 0, 1);
  const nextYearStart = Date.UTC(year + 1, 0, 1);
  const sequence = rendered.slice(2);
  let fraction = /^\d+$/.test(sequence) && sequence ? Number(sequence) / 10 ** sequence.length : 0;
  fraction = Math.max(0, Math.min(0.999999, fraction));
  return yearStart + fraction * (nextYearStart - yearStart);
}

function trendDateBound(value, endOfDay) {
  if (value === undefined || value === null || value === "") return null;
  const date = value instanceof Date ? new Date(value) : new Date(value);
  if (!Number.isFinite(date.getTime())) throw new Error("Enter a valid trend date.");
  if (typeof value === "string" && /^\d{4}-\d{2}-\d{2}$/.test(value)) {
    const [year, month, day] = value.split("-").map(Number);
    return Date.UTC(year, month - 1, day, endOfDay ? 23 : 0, endOfDay ? 59 : 0, endOfDay ? 59 : 0, endOfDay ? 999 : 0);
  }
  if (endOfDay) date.setHours(23, 59, 59, 999);
  else date.setHours(0, 0, 0, 0);
  return date.getTime();
}

function compareTrendRows(a, b) {
  return a.date - b.date || text(a.lot).localeCompare(text(b.lot), undefined, { numeric: true }) ||
    text(a.batch).localeCompare(text(b.batch), undefined, { numeric: true });
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

export function buildCorrelation(headers, rows, xParameter, yParameter, scope, outlierMethod, removalValue) {
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
  if (outlierMethod === "ratio") removeRatioExtremes(pairs, removalValue);
  else if (["x-largest", "x-smallest", "y-largest", "y-smallest"].includes(outlierMethod)) {
    removeAxisExtremes(pairs, outlierMethod, removalValue);
  }
  const included = pairs.filter((pair) => pair.included);
  if (included.length < 3) throw new Error("Too few pairs remain after extreme-pair removal.");
  return { pairs, included, totalN: pairs.length, excludedN: pairs.length - included.length, ...pearsonAndLine(included) };
}

function removeRatioExtremes(pairs, removalPct) {
  const maxRemove = Math.floor(pairs.length * Math.max(0, Math.min(50, removalPct)) / 100);
  if (!maxRemove) return;
  const ratioPairs = pairs.filter((pair) => pair.x !== 0).map((pair) => ({ pair, ratio: pair.y / pair.x }));
  if (ratioPairs.length < 3) return;
  const center = median(ratioPairs.map((item) => item.ratio));
  ratioPairs.sort((a, b) => Math.abs(b.ratio - center) - Math.abs(a.ratio - center));
  ratioPairs.slice(0, Math.min(maxRemove, ratioPairs.length - 3)).forEach((item) => {
    item.pair.included = false;
    item.pair.exclusionReason = "Ratio extreme";
  });
}

function removeAxisExtremes(pairs, method, removalValue) {
  const count = Number(removalValue);
  if (!Number.isInteger(count) || count < 0) throw new Error("Points to remove (N) must be a whole number of zero or more.");
  if (count > pairs.length - 3) throw new Error(`Points to remove (N) must be ${pairs.length - 3} or fewer so at least three pairs remain.`);
  if (!count) return;
  const axis = method.startsWith("x-") ? "x" : "y";
  const largest = method.endsWith("largest");
  const label = `${largest ? "Largest" : "Smallest"} ${axis.toUpperCase()}`;
  pairs
    .map((pair, index) => ({ pair, index }))
    .sort((left, right) => {
      const difference = largest ? right.pair[axis] - left.pair[axis] : left.pair[axis] - right.pair[axis];
      return difference || left.index - right.index;
    })
    .slice(0, count)
    .forEach(({ pair }) => {
      pair.included = false;
      pair.exclusionReason = label;
    });
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
