export const CORRELATION_METHODS = [
  ["overall-group", "Overall / MR / Lot methods"], ["zone-group", "Zone-specific methods"],
  ["all", "All methods"], ["raw", "Raw pooled"], ["batch", "MR means"],
  ["within", "MR means within Lot"], ["between", "Between Lots"],
  ["region", "Each Zone raw"], ["regionWithin", "Each Zone within Lot"],
  ["overall", "Overall within Lot x Zone"]
];
const clean = value => String(value ?? "").trim();

export function structuredConclusions(results) {
  const groups = new Map();
  for (const row of results) {
    const key = JSON.stringify([row.xParameter, row.yParameter]);
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(row);
  }
  return [...groups.values()].map(rows => {
    const valid = rows.filter(row => Number.isFinite(row.r) && row.n >= Math.max(3, row.minN || 3));
    const strength = r => Math.abs(r) >= 0.7 ? "very strong" : Math.abs(r) >= 0.5 ? "strong" : Math.abs(r) >= 0.3 ? "moderate" : "weak";
    const describe = row => `${row.method}: r=${row.r.toFixed(3)}, N=${row.n} (${row.r === 0 ? "no linear association" : `${strength(row.r)} ${row.r > 0 ? "positive" : "negative"} linear association`})`;
    const coreNames = ["Raw pooled", "Between Lots", "MR means", "MR means within Lot", "Overall within Lot x Zone"];
    const core = coreNames.map(name => valid.find(row => row.method === name)).filter(Boolean);
    const lines = core.map(describe);
    const regions = valid.filter(row => /^Zone \d+/.test(row.method));
    if (regions.length) {
      const strongest = regions.reduce((best, row) => Math.abs(row.r) > Math.abs(best.r) ? row : best);
      lines.push(`Largest absolute Zone-specific r: ${describe(strongest)}. This compares associations within Zones, not differences between Zone averages.`);
      if (regions.some(row => row.r > 0) && regions.some(row => row.r < 0)) lines.push("Zone-specific results have different signs; there is no uniform direction across the displayed Zone results.");
    }
    const excluded = rows.filter(row => row.excluded > 0);
    const between = valid.find(row => row.method === "Between Lots");
    const within = valid.find(row => row.method === "MR means within Lot");
    if (between && within && !excluded.length) {
      const difference = Math.abs(between.r) - Math.abs(within.r);
      lines.push(Math.abs(difference) < 0.1
        ? "The absolute correlations between lots and between MRs within lots are similar. This does not identify a single source of the association."
        : difference > 0
          ? "The association is stronger across lot averages than across MRs within the same lot. Lot-level differences may contribute to the overall relationship."
          : "The association is stronger across MRs within the same lot than across lot averages. The relationship is not limited to differences between lots.");
    }
    lines.push("Strength uses unrounded |r|: below 0.3 weak; 0.3 to below 0.5 moderate; 0.5 to below 0.7 strong; 0.7 or above very strong. These are descriptive thresholds, not significance tests.");
    if (!valid.length) lines.push("No interpretable correlation is available in the current table. Check the minimum N and variation in both parameters.");
    if (!between || !within) lines.push("Both Between Lots and MR means within Lot are needed to compare lot-level and within-lot MR associations. Run All methods for that comparison.");
    if (excluded.length) lines.push(`Uses updated table values: ${excluded.map(row => `${row.method}: ${row.excluded} excluded (${row.removal || "user exclusion"})`).join("; ")}. Exclusions can change r and retain different populations, so these results do not establish which level drives the association.`);
    if (valid.some(row => row.n < 10)) lines.push("Some results use fewer than 10 pairs and can be sensitive to individual observations.");
    lines.push("Descriptive linear associations only, not causation or significance tests. N represents the observation unit for each method; clustered pairs are not necessarily independent.");
    return { xParameter: rows[0].xParameter, yParameter: rows[0].yParameter, lines };
  });
}
const identity = value => clean(value).toUpperCase();
const numeric = value => (typeof value === "number" || typeof value === "string") && clean(value) !== "" && Number.isFinite(Number(value));
const regionHeader = header => {
  const match = clean(header).match(/^(.*)_(\d{1,8})$/);
  return match && Number(match[2]) > 0 ? { base: match[1], id: Number(match[2]) } : null;
};

export function trimStructuredPlot(points, { axis = "x", side = "largest", count = 0, minN = 3, method = "count", percent = 90 } = {}) {
  if (!["count", "ratio", "line"].includes(method)) throw new Error("Choose a valid exclusion method.");
  if (!["x", "y"].includes(axis) || !["largest", "smallest"].includes(side)) throw new Error("Choose X or Y and a valid direction.");
  if (!Number.isInteger(count) || count < 0 || points.length - count < minN) throw new Error(`Keep at least ${minN} pairs; enter a valid whole-number removal count.`);
  const ranked = points.map((point, index) => ({ point, index })).sort((a, b) => (side === "largest" ? b.point[axis] - a.point[axis] : a.point[axis] - b.point[axis]) || a.index - b.index);
  const removed = new Set(ranked.slice(0, count).map(item => item.index));
  let undefinedRatios = 0;
  if (method !== "count") {
    if (!Number.isFinite(percent) || percent < 0) throw new Error("Enter a nonnegative tolerance percentage.");
    const baseline = accumulator();
    points.forEach(point => add(baseline, point.x, point.y, "plot"));
    if (!(baseline.xx > 0)) throw new Error("The original data must have variation in X.");
    const slope = baseline.xy / baseline.xx, intercept = baseline.y - slope * baseline.x;
    if (!Number.isFinite(slope) || !Number.isFinite(intercept)) throw new Error("The original fitted line is not finite.");
    if (method === "ratio" && slope === 0) throw new Error("Ratio percentage is undefined for a zero slope. Use fitted-line tolerance instead.");
    removed.clear();
    points.forEach((point, index) => {
      if (method === "ratio" && point.x === 0) { removed.add(index); undefinedRatios++; return; }
      const actual = method === "ratio" ? point.y / point.x : point.y;
      const reference = method === "ratio" ? slope : slope * point.x + intercept;
      const delta = Math.abs(actual - reference), limit = Math.abs(reference) * percent / 100;
      const rounding = Number.EPSILON * Math.max(1, Math.abs(actual), Math.abs(reference)) * 16;
      if (!Number.isFinite(delta) || delta > limit + rounding) removed.add(index);
    });
  }
  const retained = points.filter((_, index) => !removed.has(index));
  if (retained.length < minN) throw new Error(`This setting retains ${retained.length} pairs. Increase the tolerance to keep at least ${minN}.`);
  const stats = accumulator();
  retained.forEach(point => add(stats, point.x, point.y, point.lot ?? "plot", point.batchIds, point.regionIds, point.rawPairs, point.group));
  const slope = stats.xx > 0 ? stats.xy / stats.xx : null;
  const r = stats.xx > 0 && stats.yy > 0 ? Math.max(-1, Math.min(1, stats.xy / Math.sqrt(stats.xx) / Math.sqrt(stats.yy))) : null;
  return { points: retained, n: retained.length, excluded: removed.size, undefinedRatios, r, slope, intercept: slope === null ? null : stats.y - slope * stats.x,
    lots: stats.lots.size, batches: stats.batches.size, regions: stats.regions.size, rawPairs: stats.rawPairs, informativeGroups: stats.groups.size,
    removal: method === "count" ? (count ? `${axis.toUpperCase()} ${side} ${count}` : "None") : `${method === "ratio" ? "Y/X vs original slope" : "Y vs original fitted line"}: +/-${percent}%${undefinedRatios ? `; ${undefinedRatios} undefined X=0 ratios excluded` : ""}` };
}

export function replayStructuredExclusions(original, steps = []) {
  let result = { ...original, excluded: 0, removal: "None" };
  const trimHistory = [];
  for (const step of steps) {
    const trimmed = trimStructuredPlot(result.points, step.options);
    trimHistory.push({ options: { ...step.options }, savedAt: step.savedAt,
      removed: trimmed.excluded, remaining: trimmed.n, r: trimmed.r,
      removedPercent: result.points.length ? trimmed.excluded / result.points.length * 100 : 0,
      cumulativeRemovedPercent: original.points.length ? (original.points.length - trimmed.n) / original.points.length * 100 : 0,
      description: trimmed.removal.replace("original slope", "step baseline slope").replace("original fitted line", "step baseline fitted line") });
    result = { ...original, ...trimmed };
  }
  return { ...result, excluded: original.points.length - result.points.length, trimHistory,
    removal: trimHistory.map((step, index) => `${index + 1}. ${step.description} (${step.removed} removed, ${step.remaining} remaining)`).join("; ") || "None" };
}

export function structuredParameters(headers, rows) {
  const found = new Map();
  headers.forEach((header, column) => {
    const name = clean(header), region = regionHeader(name);
    if (region) found.set(identity(region.base), region.base);
    else if (name && !["LOT", "N", "TYPE", "CLASSIFICATION"].includes(identity(name)) && !/_SD$/i.test(name) && rows.some(row => numeric(row[column]))) {
      found.set(identity(name), name);
    }
  });
  return [...found.values()].sort((a, b) => a.localeCompare(b));
}

function parameterColumns(headers, name) {
  const regions = new Map();
  const scalar = headers.findIndex(header => identity(header) === identity(name));
  headers.forEach((header, column) => {
    const region = regionHeader(header);
    if (!region || identity(region.base) !== identity(name)) return;
    if (regions.has(region.id)) throw new Error(`Duplicate Zone header: ${header}`);
    regions.set(region.id, column);
  });
  return { scalar, regions };
}

function accumulator(capture = false) {
  return { n: 0, x: 0, y: 0, xx: 0, yy: 0, xy: 0, lots: new Set(), batches: new Set(), regions: new Set(), rawPairs: 0, groups: new Set(), points: capture ? [] : undefined };
}
function add(stats, x, y, lot, batchIds = [], regionIds = [], rawPairs = 1, group = null) {
  if (stats.points) stats.points.push({ x, y, lot, batchIds, regionIds, rawPairs, group });
  const dx = x - stats.x, dy = y - stats.y;
  stats.n++;
  stats.x += dx / stats.n; stats.y += dy / stats.n;
  stats.xx += dx * (x - stats.x); stats.yy += dy * (y - stats.y); stats.xy += dx * (y - stats.y);
  stats.lots.add(lot);
  batchIds.forEach(id => stats.batches.add(id));
  regionIds.forEach(id => stats.regions.add(id));
  stats.rawPairs += rawPairs;
  if (group !== null) stats.groups.add(group);
}
function meanPair(group, x, y) {
  group.n++; group.x += (x - group.x) / group.n; group.y += (y - group.y) / group.n;
}

export function buildStructuredCorrelation(headers, rows, yParameter, options = {}) {
  const method = options.method || "all", coverage = options.coverage || "available";
  const minN = Number(options.minN ?? 3), minRegions = Number(options.minRegions ?? 1);
  if (!CORRELATION_METHODS.some(([id]) => id === method)) throw new Error("Select a valid correlation method.");
  if (!["available", "strict"].includes(coverage)) throw new Error("Select a valid Zone coverage policy.");
  if (!Number.isInteger(minN) || minN < 3 || minN > 3000000) throw new Error("Minimum N must be an integer from 3 to 3000000.");
  if (!Number.isInteger(minRegions) || minRegions < 1 || minRegions > 16384) throw new Error("Minimum paired Zones must be an integer from 1 to 16384.");
  const expected = new Set();
  clean(options.expectedRegions).split(/[,;\s]+/).filter(Boolean).forEach(token => {
    if (!/^\d{1,8}$/.test(token) || Number(token) < 1) throw new Error("Expected Zone IDs must be positive integers, such as 1,2,4.");
    expected.add(Number(token));
  });
  const lotColumn = headers.findIndex(header => identity(header) === "LOT");
  const batchColumn = headers.findIndex(header => identity(header) === "N");
  if (lotColumn < 0 || batchColumn < 0) throw new Error("Structured correlation requires Lot and N columns.");
  // Keep the original catalog when the current Lot filter has no numeric rows.
  const parameters = options.parameters || structuredParameters(headers, rows);
  if (!parameters.some(name => identity(name) === identity(yParameter))) throw new Error("Select a valid Y parameter.");
  const comparisons = options.compareAll ? parameters.filter(name => identity(name) !== identity(yParameter)) : [options.xParameter];
  if (!comparisons.length || comparisons.some(name => !name || identity(name) === identity(yParameter) || !parameters.some(item => identity(item) === identity(name)))) {
    throw new Error("Select a different X parameter.");
  }
  const output = [];
  for (const xParameter of comparisons) {
    const xs = parameterColumns(headers, xParameter), ys = parameterColumns(headers, yParameter);
    const scalar = !xs.regions.size && !ys.regions.size;
    const ids = scalar ? [0] : [...new Set([...xs.regions.keys(), ...ys.regions.keys()])];
    const base = { xParameter, yParameter, coverage, minRegions, minN };
    if (Boolean(xs.regions.size) !== Boolean(ys.regions.size) || scalar && (xs.scalar < 0 || ys.scalar < 0)) {
      output.push({ ...base, method: "Unavailable pair", r: null, n: 0, unit: "n.a.", batches: 0, lots: 0, regions: 0, informativeLots: 0, informativeGroups: 0, rawPairs: 0, status: "Different measurement levels: select explicit MR-level columns for both parameters." });
      continue;
    }
    if (coverage === "strict" && !scalar && !expected.size) throw new Error("Enter Expected Zone IDs before using strict coverage.");
    if (rows.length * ids.length > 3000000) throw new Error("More than 3 million candidate pairs. Filter the data first.");
    const capture = name => accumulator(options.plotMethod === name);
    const raw = capture("Raw pooled"), batch = capture("MR means"), within = capture("MR means within Lot"), between = capture("Between Lots"), overall = capture("Overall within Lot x Zone");
    const regionStats = new Map(ids.map(id => [id, capture(`Zone ${id} raw`)]));
    const adjustedStats = new Map(ids.map(id => [id, capture(`Zone ${id} within Lot`)]));
    const lotMeans = new Map(), lotRegionMeans = new Map(), accepted = [];
    let missingIdentity = 0, rejected = 0;
    const signatures = new Set();
    const visitPairs = (row, callback) => {
      for (const id of ids) {
        const xc = scalar ? xs.scalar : xs.regions.get(id), yc = scalar ? ys.scalar : ys.regions.get(id);
        if (xc === undefined || yc === undefined || !numeric(row[xc]) || !numeric(row[yc])) continue;
        callback(id, Number(row[xc]), Number(row[yc]));
      }
    };
    rows.forEach((row, rowIndex) => {
      const lot = identity(row[lotColumn]);
      if (!lot || !clean(row[batchColumn])) { missingIdentity++; rejected++; return; }
      const mean = { n: 0, x: 0, y: 0 }, used = [];
      visitPairs(row, (id, x, y) => {
        add(raw, x, y, lot, [rowIndex], scalar ? [] : [id]);
        if (!scalar) {
          add(regionStats.get(id), x, y, lot, [rowIndex], [id]);
          const key = JSON.stringify([lot, id]);
          if (!lotRegionMeans.has(key)) lotRegionMeans.set(key, { n: 0, x: 0, y: 0 });
          meanPair(lotRegionMeans.get(key), x, y);
        }
        if (!scalar && coverage === "strict" && !expected.has(id)) return;
        meanPair(mean, x, y); used.push(id);
      });
      if (!mean.n || !scalar && (mean.n < minRegions || coverage === "strict" && mean.n !== expected.size)) { rejected++; return; }
      const regions = scalar ? [] : used;
      signatures.add(used.join(","));
      add(batch, mean.x, mean.y, lot, [rowIndex], regions, mean.n);
      if (!lotMeans.has(lot)) lotMeans.set(lot, { n: 0, x: 0, y: 0, batchIds: [], regions: new Set(), rawPairs: 0 });
      const group = lotMeans.get(lot);
      meanPair(group, mean.x, mean.y); group.batchIds.push(rowIndex); group.rawPairs += mean.n;
      regions.forEach(id => group.regions.add(id));
      accepted.push({ rowIndex, lot, ...mean, regions });
    });
    let singletonLots = 0, singletonGroups = 0;
    for (const [lot, group] of lotMeans) {
      add(between, group.x, group.y, lot, group.batchIds, [...group.regions], group.rawPairs);
      if (group.n === 1) singletonLots++;
    }
    for (const entry of accepted) {
      const group = lotMeans.get(entry.lot);
      if (group.n < 2) continue;
      add(within, entry.x - group.x, entry.y - group.y, entry.lot, [entry.rowIndex], entry.regions, entry.n);
    }
    if (!scalar) {
      for (const group of lotRegionMeans.values()) if (group.n === 1) singletonGroups++;
      rows.forEach((row, rowIndex) => {
        const lot = identity(row[lotColumn]);
        if (!lot || !clean(row[batchColumn])) return;
        visitPairs(row, (id, x, y) => {
          const key = JSON.stringify([lot, id]), group = lotRegionMeans.get(key);
          if (group.n < 2) return;
          add(adjustedStats.get(id), x - group.x, y - group.y, lot, [rowIndex], [id], 1, key);
          add(overall, x - group.x, y - group.y, lot, [rowIndex], [id], 1, key);
        });
      });
    }
    const methods = [
      ["raw", "Raw pooled", raw, scalar ? "MR scalar pairs" : "Zone pairs"],
      ["batch", "MR means", batch, "MR occurrences"],
      ["within", "MR means within Lot", within, "MR residual pairs"],
      ["between", "Between Lots", between, "Lots (one point each)"],
      ...(!scalar ? ids.map(id => ["region", `Zone ${id} raw`, regionStats.get(id), "Zone pairs"]) : []),
      ...(!scalar ? ids.map(id => ["regionWithin", `Zone ${id} within Lot`, adjustedStats.get(id), "Zone residual pairs"]) : []),
      ...(!scalar ? [["overall", "Overall within Lot x Zone", overall, "Zone residual pairs"]] : [])
    ].filter(([id]) => method === "all" || id === method ||
      method === "overall-group" && ["raw", "batch", "within", "between"].includes(id) ||
      method === "zone-group" && ["region", "regionWithin", "overall"].includes(id));
    if (!methods.length) {
      output.push({ ...base, method: CORRELATION_METHODS.find(([id]) => id === method)[1], r: null, n: 0, unit: "n.a.", batches: 0, lots: 0, regions: 0, informativeLots: 0, informativeGroups: 0, rawPairs: 0, status: "No Zone identity for this method." });
    }
    for (const [id, name, stats, unit] of methods) {
      let r = null;
      let status = stats.n < minN ? "Below minimum N (at least 3 required)" : stats.xx <= 0 || stats.yy <= 0 ? "No variation after processing" : "Descriptive; observations are clustered";
      if (stats.n >= minN && stats.xx > 0 && stats.yy > 0) {
        const coefficient = stats.xy / Math.sqrt(stats.xx) / Math.sqrt(stats.yy);
        if (Number.isFinite(coefficient)) r = Math.max(-1, Math.min(1, coefficient));
        else status = "Values exceed the supported numeric range";
      }
      if (["batch", "within", "between"].includes(id)) {
        if (signatures.size > 1) status += "; matched Zone subsets vary between MRs";
        if (rejected) status += `; ${rejected} MRs ineligible for MR means`;
      } else status += "; MR-mean coverage does not filter raw/Zone observations";
      if (id === "within") status += `; ${singletonLots} singleton Lots excluded`;
      if (["regionWithin", "overall"].includes(id)) status += `; ${singletonGroups} singleton Lot-Zone groups excluded across pair`;
      if (missingIdentity) status += `; ${missingIdentity} source rows excluded: missing Lot/N`;
      if (scalar) status += "; no Zone identity; coverage not applicable";
      if (id === "between") status += "; each Lot contributes one pair";
      output.push({ ...base, method: name, r, n: stats.n, unit, batches: stats.batches.size, lots: stats.lots.size, regions: stats.regions.size, rawPairs: stats.rawPairs,
        informativeLots: ["within", "regionWithin", "overall"].includes(id) ? stats.lots.size : 0,
        informativeGroups: stats.groups.size, status,
        ...(stats.points ? { points: stats.points, slope: stats.xx > 0 ? stats.xy / stats.xx : null, intercept: stats.xx > 0 ? stats.y - stats.xy / stats.xx * stats.x : null } : {}) });
    }
  }
  return { results: output, method, coverage, minRegions, minN, expectedRegions: [...expected] };
}
