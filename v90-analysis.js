import {
  getRegionalParameters,
  headerIndex,
  isNumeric,
  meanAndSigma,
  number,
  robustHistory,
  text,
  zoneColumns
} from "./analysis.js?v=13";

const ZONES = [1, 2, 3, 4, 5, 6];

const ZM_LAYOUTS = {
  ZM17_20mm: { totalWidth: 1580, segmentWidth: 780, zonesPerSegment: 3, zoneRollCounts: [13, 12, 13, 13, 12, 13] },
  ZM17_25mm: { totalWidth: 1580, segmentWidth: 780, zonesPerSegment: 3, zoneRollCounts: [10, 10, 10, 10, 10, 10] },
  ZM9_18mm: { totalWidth: 1200, segmentWidth: 360, zonesPerSegment: 2, zoneRollCounts: [10, 9, 10, 9, 10, 9] },
  ZM9_20mm: { totalWidth: 1200, segmentWidth: 360, zonesPerSegment: 2, zoneRollCounts: [9, 8, 9, 8, 9, 8] },
  ZM9_25mm: { totalWidth: 1200, segmentWidth: 360, zonesPerSegment: 2, zoneRollCounts: [7, 7, 7, 7, 7, 7] },
  ZM10_20mm: { totalWidth: 1200, segmentWidth: 390, zonesPerSegment: 2, zoneRollCounts: [10, 9, 10, 9, 10, 9] },
  ZM10_25mm: { totalWidth: 1200, segmentWidth: 360, zonesPerSegment: 2, zoneRollCounts: [8, 7, 8, 7, 8, 7] }
};

export function getZmPlanSpecification(layoutName) {
  const specification = ZM_LAYOUTS[layoutName];
  if (!specification) throw new Error("Select a supported ZM layout.");
  return { ...specification, zoneRollCounts: [...specification.zoneRollCounts] };
}

export function getV90Parameters(headers) {
  const regional = getRegionalParameters(headers).filter((parameter) =>
    zoneColumns(headers, parameter)?.every((column) => column >= 0)
  );
  for (const scalar of ["Water", "Visco.", "Temp.", "Humidity"]) {
    const column = headerIndex(headers, scalar);
    if (column >= 0 && !regional.includes(headers[column])) regional.push(headers[column]);
  }
  return regional.sort((a, b) => a.localeCompare(b));
}

export function buildPeriodComparison(headers, rows, parameter, dateLookup = new Map(), options = {}) {
  const lotColumn = headerIndex(headers, "Lot");
  const batchColumn = headerIndex(headers, "N");
  const typeColumn = headerIndex(headers, "Type");
  if (lotColumn < 0 || batchColumn < 0) throw new Error("The source sheet must contain Lot and N columns.");

  const mode = options.mode || "two-periods";
  if (!["two-periods", "three-periods", "lot-vs-period", "lot-vs-lot"].includes(mode)) {
    throw new Error("Select a valid comparison mode.");
  }
  const regionalColumns = zoneColumns(headers, parameter);
  const regional = Boolean(regionalColumns?.every((column) => column >= 0));
  const scalarColumn = regional ? -1 : headerIndex(headers, parameter);
  if (!regional && scalarColumn < 0) throw new Error(`Parameter column is missing: ${parameter}`);

  const definitions = buildDatasetDefinitions(mode, options);
  const accumulators = definitions.map(() => Array.from({ length: 7 }, () => []));
  let fallbackCount = 0;

  for (const row of rows) {
    const lot = text(row[lotColumn]);
    const batch = text(row[batchColumn]);
    const type = typeColumn >= 0 ? text(row[typeColumn]) : "";
    let date = dateLookup.get(trendKey(lot, batch, type));
    let fallback = false;
    if (!Number.isFinite(date)) {
      date = dateFromLot(lot);
      fallback = Number.isFinite(date);
    }
    const includedIndexes = [];
    definitions.forEach((definition, index) => {
      if (definition.kind === "lot") {
        if (identity(lot) === identity(definition.lot)) includedIndexes.push(index);
      } else if (Number.isFinite(date) && date >= definition.start && date <= definition.end) {
        includedIndexes.push(index);
      }
    });
    if (!includedIndexes.length) continue;
    if (fallback) fallbackCount += 1;

    if (!regional) {
      if (!isNumeric(row[scalarColumn])) continue;
      const value = number(row[scalarColumn]);
      includedIndexes.forEach((index) => accumulators[index][6].push(value));
      continue;
    }
    ZONES.forEach((zone) => {
      const value = row[regionalColumns[zone - 1]];
      if (!isNumeric(value)) return;
      includedIndexes.forEach((index) => {
        accumulators[index][zone - 1].push(number(value));
        accumulators[index][6].push(number(value));
      });
    });
  }

  const datasets = definitions.map((definition, datasetIndex) => {
    const stats = accumulators[datasetIndex].map((values, zoneIndex) => ({
      zone: zoneIndex < 6 ? zoneIndex + 1 : "All",
      ...meanAndSigma(values)
    }));
    if (!stats[6].n) throw new Error(`${definition.label} contains no numeric values for ${parameter}.`);
    const allMean = stats[6].mean;
    stats.forEach((item) => {
      item.normalized = item.n && Number.isFinite(allMean) && allMean !== 0 ? item.mean / allMean : null;
    });
    return { key: String.fromCharCode(65 + datasetIndex), ...definition, stats };
  });

  return { parameter, mode, regional, datasets, fallbackCount, lookupCount: dateLookup.size };
}

export function buildLotReleaseSummary(headers, allRows, referenceRows, selectedLot, monitorLimit = 2, notOkLimit = 3) {
  const lotColumn = headerIndex(headers, "Lot");
  if (lotColumn < 0) throw new Error("The source sheet must contain a Lot column.");
  if (!(monitorLimit > 0 && notOkLimit > monitorLimit)) {
    throw new Error("Not OK limit must be greater than Monitor limit, and both must be positive.");
  }
  const parameters = getRegionalParameters(headers).filter((parameter) =>
    zoneColumns(headers, parameter)?.every((column) => column >= 0)
  );
  if (!parameters.length) throw new Error("No complete Zone 1-6 parameter was found.");
  const lotRows = allRows.filter((row) => identity(row[lotColumn]) === identity(selectedLot));
  if (!lotRows.length) throw new Error("No rows were found for the selected Lot.");
  const historyRows = referenceRows.filter((row) => identity(row[lotColumn]) !== identity(selectedLot));

  const results = parameters.map((parameter) => {
    const columns = zoneColumns(headers, parameter);
    const lotStats = zoneAndAllStats(lotRows, columns);
    const referenceStats = zoneAndAllStats(historyRows, columns);
    let monitorPoints = 0;
    let notOkPoints = 0;
    let status = "OK";
    if (!lotStats[6].n || referenceStats[6].n < 2 || !(referenceStats[6].sigma > 0)) {
      status = "NO HISTORY";
    } else {
      lotRows.forEach((row) => ZONES.forEach((zone) => {
        const value = row[columns[zone - 1]];
        const reference = referenceStats[zone - 1];
        if (!isNumeric(value) || reference.n < 2 || !(reference.sigma > 0)) return;
        const absoluteZ = Math.abs((number(value) - reference.mean) / reference.sigma);
        if (absoluteZ >= notOkLimit) notOkPoints += 1;
        else if (absoluteZ >= monitorLimit) monitorPoints += 1;
      }));
      status = notOkPoints ? "NOT OK" : monitorPoints ? "MONITOR" : "OK";
    }

    const lotAll = lotStats[6];
    const referenceAll = referenceStats[6];
    const lotCv = lotAll.mean ? lotAll.sigma / Math.abs(lotAll.mean) : null;
    const referenceCv = referenceAll.mean ? referenceAll.sigma / Math.abs(referenceAll.mean) : null;
    let maxZoneBiasDelta = 0;
    ZONES.forEach((zone) => {
      const lotZone = lotStats[zone - 1];
      const referenceZone = referenceStats[zone - 1];
      if (!lotZone.n || !referenceZone.n) return;
      maxZoneBiasDelta = Math.max(
        maxZoneBiasDelta,
        Math.abs((lotZone.mean - lotAll.mean) - (referenceZone.mean - referenceAll.mean))
      );
    });
    return {
      parameter,
      status,
      lotN: lotAll.n,
      referenceN: referenceAll.n,
      lotMean: lotAll.mean,
      referenceMean: referenceAll.mean,
      meanDelta: lotAll.n && referenceAll.n ? lotAll.mean - referenceAll.mean : null,
      meanZ: referenceAll.sigma > 0 ? (lotAll.mean - referenceAll.mean) / referenceAll.sigma : null,
      lotCv,
      referenceCv,
      cvRatio: Number.isFinite(lotCv) && referenceCv ? lotCv / referenceCv : null,
      maxZoneBiasDelta,
      monitorPoints,
      notOkPoints,
      lotStats,
      referenceStats
    };
  });
  const counts = {
    notOk: results.filter((item) => item.status === "NOT OK").length,
    monitor: results.filter((item) => item.status === "MONITOR").length,
    noHistory: results.filter((item) => item.status === "NO HISTORY").length
  };
  const overall = counts.notOk ? "NOT OK" : counts.monitor ? "MONITOR" : counts.noHistory ? "NO HISTORY" : "OK";
  return { selectedLot, monitorLimit, notOkLimit, overall, counts, results };
}

export function integerChartAxis(values) {
  const finite = values.filter(Number.isFinite);
  if (!finite.length) return { min: 0, max: 1, ticks: [0, 1] };
  const low = Math.min(...finite);
  const high = Math.max(...finite);
  const padding = (high - low || 1) * 0.06;
  const rawStep = Math.max(1, (high - low + 2 * padding) / 4);
  const magnitude = 10 ** Math.floor(Math.log10(rawStep));
  const step = [1, 2, 5, 10].find((factor) => factor * magnitude >= rawStep) * magnitude;
  const min = Math.floor((low >= 0 ? Math.max(0, low - padding) : low - padding) / step) * step;
  const max = Math.max(min + step, Math.ceil((high + padding) / step) * step);
  const ticks = Array.from({ length: Math.round((max - min) / step) + 1 }, (_, index) => min + index * step);
  return { min, max, ticks };
}

export function buildV90LotAssessment(headers, allRows, referenceRows, options = {}) {
  const lotColumn = headerIndex(headers, "Lot");
  const batchColumn = headerIndex(headers, "N");
  if (lotColumn < 0 || batchColumn < 0) throw new Error("The source sheet must contain Lot and N columns.");
  const selectedLot = text(options.selectedLot);
  if (!selectedLot) throw new Error("Select a Lot to assess.");
  const monitorLimit = Number(options.monitorLimit);
  const outlierLimit = Number(options.outlierLimit);
  if (!(monitorLimit > 0 && outlierLimit > monitorLimit)) {
    throw new Error("Use positive limits and make the out-of-range limit larger than the monitor limit.");
  }
  const mode = options.referenceMode || "history";
  if (!["history", "filtered", "manual", "reference-lot", "equal-lots"].includes(mode)) throw new Error("Select a valid reference method.");
  const equalLotMode = mode === "equal-lots";
  const referenceLots = [...new Map((options.referenceLots || [])
    .filter((lot) => text(lot) && identity(lot) !== identity(selectedLot))
    .map((lot) => [identity(lot), text(lot)])).values()];
  const chosenLots = new Set(referenceLots.map(identity));
  if (equalLotMode) {
    const availableLots = new Set(allRows.map((row) => identity(row[lotColumn])));
    if (referenceLots.some((lot) => !availableLots.has(identity(lot)))) {
      throw new Error("A selected reference Lot is no longer in the source. Refresh the reference Lot selection.");
    }
    if (chosenLots.size < 2) throw new Error("Choose at least two reference Lots, excluding the target. Use Reference Lot for one Lot.");
  }
  const referenceLot = text(options.referenceLot);
  const referenceGranularity = options.referenceGranularity === "batch-zone" ? "batch-zone" : "zone";
  const batchZoneMode = referenceGranularity === "batch-zone";
  if (mode === "reference-lot") {
    if (!referenceLot) throw new Error("Select the Reference Lot.");
    if (identity(referenceLot) === identity(selectedLot)) throw new Error("Selected Lot and Reference Lot must be different.");
  }
  const selectedRows = allRows.filter((row) => identity(row[lotColumn]) === identity(selectedLot));
  if (!selectedRows.length) throw new Error("The selected Lot has no data rows.");
  const availableParameters = getV90Parameters(headers);
  const requested = text(options.parameter) || "All parameters";
  const parameters = requested === "All parameters" ? availableParameters : availableParameters.filter((item) => item === requested);
  if (!parameters.length) throw new Error("No eligible numeric parameters were found.");
  if (mode === "manual" && parameters.length !== 1) throw new Error("Select one parameter when using manual mu and sigma.");
  if (batchZoneMode) {
    const regionalColumns = parameters.length === 1 ? zoneColumns(headers, parameters[0]) : null;
    if (parameters.length !== 1 || !regionalColumns?.every((column) => column >= 0)) {
      throw new Error("Batch + Zone matching requires one complete Zone 1-6 parameter.");
    }
    if (mode === "manual") throw new Error("Manual mu and sigma support Zone-only matching.");
  }

  const sourceHistory = mode === "filtered" ? referenceRows : allRows;
  // Index once per assessment, preserving source order and avoiding stale caches.
  const historyGroups = new Map();
  if (mode !== "manual") sourceHistory.forEach((row) => {
    const lot = identity(row[lotColumn]);
    if (lot === identity(selectedLot)) return;
    if (mode === "reference-lot" && lot !== identity(referenceLot)) return;
    if (equalLotMode && !chosenLots.has(lot)) return;
    const key = batchZoneMode ? assessmentBatchKey(row[batchColumn]) : "";
    if (!historyGroups.has(key)) historyGroups.set(key, []);
    historyGroups.get(key).push(row);
  });
  const referenceCache = new Map();
  const referenceFor = (sourceColumn, batchValue = null) => {
    if (mode === "manual") {
      const mean = Number(options.manualMu);
      const sigma = Number(options.manualSigma);
      if (!(Number.isFinite(mean) && Number.isFinite(sigma) && sigma > 0)) throw new Error("Enter a positive manual mu and sigma.");
      return { n: 2, mean, sigma, excluded: 0 };
    }
    const targetBatch = batchZoneMode ? assessmentBatchKey(batchValue) : "";
    if (!referenceCache.has(targetBatch)) referenceCache.set(targetBatch, new Map());
    const cache = referenceCache.get(targetBatch);
    if (cache.has(sourceColumn)) return cache.get(sourceColumn);
    const rows = historyGroups.get(targetBatch) || [];
    const values = equalLotMode ? [] : rows.filter((row) => isAssessmentNumeric(row[sourceColumn])).map((row) => number(row[sourceColumn]));
    const result = equalLotMode ? equalLotReference(rows, lotColumn, sourceColumn, referenceLots)
      : mode === "history" || mode === "filtered" ? robustHistory(values)
        : { ...meanAndSigma(values), excluded: 0 };
    cache.set(sourceColumn, result);
    return result;
  };
  const columns = [];
  const parameterGroups = [];
  parameters.forEach((parameter) => {
    const regionalColumns = zoneColumns(headers, parameter);
    const regional = Boolean(regionalColumns?.every((column) => column >= 0));
    const sourceColumns = regional ? regionalColumns : [headerIndex(headers, parameter)];
    if (sourceColumns.some((column) => column < 0)) return;
    const start = columns.length;
    sourceColumns.forEach((sourceColumn, index) => {
      const header = regional ? `${parameter}_${index + 1}` : parameter;
      const reference = referenceFor(sourceColumn);
      columns.push({ parameter, header, sourceColumn, zone: regional ? index + 1 : null, reference });
    });
    parameterGroups.push({ parameter, regional, columnIndexes: Array.from({ length: columns.length - start }, (_, index) => start + index) });
  });
  if (!columns.length) throw new Error("No eligible parameter columns were found.");

  let comparedCount = 0;
  let monitorCount = 0;
  let outlierCount = 0;
  let noHistoryCount = 0;
  const summaryState = new Map(parameterGroups.map((group) => [group.parameter, { z: [], monitor: 0, outlier: 0 }]));
  const grid = selectedRows.map((row) => {
    const values = [];
    const states = [];
    const scores = [];
    const signedScores = [];
    const references = [];
    columns.forEach((column) => {
      const reference = batchZoneMode ? referenceFor(column.sourceColumn, row[batchColumn]) : column.reference;
      references.push(reference);
      const value = isAssessmentNumeric(row[column.sourceColumn]) ? number(row[column.sourceColumn]) : null;
      values.push(value);
      if (value === null) {
        states.push("NO VALUE");
        scores.push(null);
        signedScores.push(null);
        return;
      }
      if (reference.n < 2) {
        noHistoryCount += 1;
        states.push("NO HISTORY");
        scores.push(null);
        signedScores.push(null);
        return;
      }
      let signedZ;
      if (reference.sigma > 0) signedZ = (value - reference.mean) / reference.sigma;
      else signedZ = Math.abs(value - reference.mean) < 1e-7 ? 0 : 999 * Math.sign(value - reference.mean);
      const score = Math.abs(signedZ);
      const summary = summaryState.get(column.parameter);
      summary.z.push(signedZ);
      comparedCount += 1;
      scores.push(score);
      signedScores.push(signedZ);
      if (score > outlierLimit) {
        states.push("OUT OF RANGE");
        summary.outlier += 1;
        outlierCount += 1;
      } else if (score > monitorLimit) {
        states.push("MONITOR");
        summary.monitor += 1;
        monitorCount += 1;
      } else {
        states.push("IN RANGE");
      }
    });
    return { batch: row[batchColumn], values, states, scores, signedScores, references };
  });

  const appliedReferences = batchZoneMode ? grid.flatMap((row) => columns.map((column, columnIndex) => ({
    batch: row.batch,
    parameter: column.parameter,
    header: column.header,
    zone: column.zone,
    ...row.references[columnIndex]
  }))) : [];

  const summaries = parameterGroups.map((group) => {
    const values = summaryState.get(group.parameter);
    const stats = meanAndSigma(values.z);
    const monitorPct = stats.n ? values.monitor / stats.n : 0;
    const outlierPct = stats.n ? values.outlier / stats.n : 0;
    const investigateCount = Math.max(2, Math.ceil(stats.n * 0.01));
    let signal;
    if (stats.n < 2) signal = "INSUFFICIENT";
    else if (values.outlier >= investigateCount || Math.abs(stats.mean) > 2 || stats.sigma > 2) signal = "INVESTIGATE";
    else if (values.outlier > 0 || monitorPct > 0.1 || Math.abs(stats.mean) > 1 || stats.sigma > 1.5) signal = "CHECK";
    else signal = "OK";
    return {
      parameter: group.parameter,
      n: stats.n,
      meanZ: stats.mean,
      zSigma: stats.sigma,
      monitorPct,
      outlierPct,
      signal
    };
  });
  const overall = !comparedCount ? "NO HISTORY"
    : summaries.some((item) => item.signal === "INVESTIGATE") ? "NOT OK"
      : summaries.some((item) => item.signal === "CHECK" || item.signal === "INSUFFICIENT") ? "CHECK"
        : "OK";
  return {
    selectedLot,
    referenceLot,
    referenceLots: equalLotMode ? referenceLots : [],
    referenceMode: mode,
    referenceGranularity,
    monitorLimit,
    outlierLimit,
    columns,
    grid,
    summaries,
    overall,
    comparedCount,
    monitorCount,
    outlierCount,
    noHistoryCount,
    appliedReferences,
    historyExcluded: (batchZoneMode ? appliedReferences : columns)
      .reduce((sum, item) => sum + ((item.reference || item).excluded || 0), 0)
  };
}

function isAssessmentNumeric(value) {
  return (typeof value === "number" || typeof value === "string") && Boolean(text(value)) && isNumeric(value);
}

function equalLotReference(rows, lotColumn, sourceColumn, referenceLots) {
  const lots = new Map();
  for (const row of rows) {
    const raw = row[sourceColumn];
    if (!isAssessmentNumeric(raw)) continue;
    const key = identity(row[lotColumn]);
    if (!lots.has(key)) lots.set(key, { n: 0, mean: 0, m2: 0 });
    const stats = lots.get(key);
    const value = number(raw);
    const delta = value - stats.mean;
    stats.n += 1;
    stats.mean += delta / stats.n;
    stats.m2 += delta * (value - stats.mean);
  }
  let center = 0;
  let betweenM2 = 0;
  let contributing = 0;
  let withinVariance = 0;
  let valueCount = 0;
  for (const stats of lots.values()) {
    contributing += 1;
    const delta = stats.mean - center;
    center += delta / contributing;
    betweenM2 += delta * (stats.mean - center);
    withinVariance += stats.m2 / stats.n;
    valueCount += stats.n;
  }
  // Balanced population variance = mean within-lot variance + variance of lot means.
  return {
    n: lots.size,
    lotCount: lots.size,
    valueCount,
    mean: lots.size ? center : null,
    sigma: lots.size ? Math.sqrt(Math.max(0, (withinVariance + betweenM2) / lots.size)) : null,
    withinVariance: lots.size ? withinVariance / lots.size : null,
    betweenVariance: lots.size ? betweenM2 / lots.size : null,
    contributors: referenceLots.map((lot) => {
      const stats = lots.get(identity(lot));
      return {
        lot,
        n: stats?.n || 0,
        mean: stats?.mean ?? null,
        sigma: stats ? Math.sqrt(Math.max(0, stats.m2 / stats.n)) : null,
        weight: stats ? 1 / lots.size : 0
      };
    }),
    excluded: 0
  };
}

function buildDatasetDefinitions(mode, options) {
  if (mode === "lot-vs-lot") {
    const lotA = text(options.lotA);
    const lotB = text(options.lotB);
    if (!lotA || !lotB) throw new Error("Select Dataset A Lot and Dataset B Lot.");
    if (identity(lotA) === identity(lotB)) throw new Error("Dataset A Lot and Dataset B Lot must be different.");
    return [
      { kind: "lot", lot: lotA, label: `Lot ${lotA}` },
      { kind: "lot", lot: lotB, label: `Lot ${lotB}` }
    ];
  }
  if (mode === "lot-vs-period") {
    if (!text(options.lotA)) throw new Error("Select Dataset A Lot.");
    return [
      { kind: "lot", lot: text(options.lotA), label: `Lot ${text(options.lotA)}` },
      periodDefinition(options.bStart, options.bEnd, "Dataset B")
    ];
  }
  const definitions = [
    periodDefinition(options.aStart, options.aEnd, "Dataset A"),
    periodDefinition(options.bStart, options.bEnd, "Dataset B")
  ];
  if (mode === "three-periods") definitions.push(periodDefinition(options.cStart, options.cEnd, "Dataset C"));
  return definitions;
}

function assessmentBatchKey(value) {
  let key = text(value).toUpperCase().replace(/\s+/g, "");
  if (key.startsWith("BATCH")) key = key.slice(5);
  if (key.startsWith("M") && key.length > 1) key = key.slice(1);
  const numeric = Number(key);
  return key && Number.isFinite(numeric) ? String(numeric) : key;
}

function periodDefinition(startValue, endValue, name) {
  const start = dateBound(startValue, false, name);
  const end = dateBound(endValue, true, name);
  if (end < start) throw new Error(`${name} end must be on or after its start.`);
  return { kind: "period", start, end, label: `${formatDate(start)} to ${formatDate(end)}` };
}

function zoneAndAllStats(rows, columns) {
  const values = Array.from({ length: 7 }, () => []);
  rows.forEach((row) => ZONES.forEach((zone) => {
    const value = row[columns[zone - 1]];
    if (!isNumeric(value)) return;
    values[zone - 1].push(number(value));
    values[6].push(number(value));
  }));
  return values.map((zoneValues, index) => ({ zone: index < 6 ? index + 1 : "All", ...meanAndSigma(zoneValues) }));
}

function dateBound(value, endOfDay, name) {
  if (!value) throw new Error(`Enter ${name} ${endOfDay ? "end" : "start"}.`);
  if (typeof value === "string" && /^\d{4}-\d{2}-\d{2}$/.test(value)) {
    const [year, month, day] = value.split("-").map(Number);
    const date = Date.UTC(year, month - 1, day, endOfDay ? 23 : 0, endOfDay ? 59 : 0, endOfDay ? 59 : 0, endOfDay ? 999 : 0);
    if (new Date(date).getUTCFullYear() !== year || new Date(date).getUTCMonth() !== month - 1 || new Date(date).getUTCDate() !== day) {
      throw new Error(`Enter a valid ${name} date.`);
    }
    return date;
  }
  const date = new Date(value);
  if (!Number.isFinite(date.getTime())) throw new Error(`Enter a valid ${name} date.`);
  if (endOfDay) date.setHours(23, 59, 59, 999);
  else date.setHours(0, 0, 0, 0);
  return date.getTime();
}

function identity(value) {
  const rendered = text(value);
  return rendered && Number.isFinite(Number(rendered)) ? String(Number(rendered)) : rendered.toUpperCase();
}

function trendKey(lot, batch, type) {
  return `${identity(lot)}|${identity(batch)}|${identity(type)}`;
}

function dateFromLot(lot) {
  const rendered = text(lot);
  if (!/^\d{2}/.test(rendered)) return null;
  const prefix = Number(rendered.slice(0, 2));
  const year = prefix <= 79 ? 2000 + prefix : 1900 + prefix;
  const start = Date.UTC(year, 0, 1);
  const end = Date.UTC(year + 1, 0, 1);
  const sequence = rendered.slice(2);
  const fraction = /^\d+$/.test(sequence) && sequence ? Math.min(0.999999, Number(sequence) / 10 ** sequence.length) : 0;
  return start + fraction * (end - start);
}

function formatDate(value) {
  return new Date(value).toISOString().slice(0, 10);
}
