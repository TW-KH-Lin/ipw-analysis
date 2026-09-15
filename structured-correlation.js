export const CORRELATION_METHODS = [
  ["all", "All methods"], ["raw", "Raw pooled"], ["batch", "Batch means"],
  ["within", "Batch means within Lot"], ["between", "Between Lots"],
  ["region", "Each Region raw"], ["regionWithin", "Each Region within Lot"],
  ["overall", "Overall within Lot x Region"]
];
const clean = value => String(value ?? "").trim();
const identity = value => clean(value).toUpperCase();
const numeric = value => (typeof value === "number" || typeof value === "string") && clean(value) !== "" && Number.isFinite(Number(value));
const regionHeader = header => {
  const match = clean(header).match(/^(.*)_(\d{1,8})$/);
  return match && Number(match[2]) > 0 ? { base: match[1], id: Number(match[2]) } : null;
};

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
    if (regions.has(region.id)) throw new Error(`Duplicate Region header: ${header}`);
    regions.set(region.id, column);
  });
  return { scalar, regions };
}

function accumulator(capture = false) {
  return { n: 0, x: 0, y: 0, xx: 0, yy: 0, xy: 0, lots: new Set(), batches: new Set(), regions: new Set(), rawPairs: 0, groups: new Set(), points: capture ? [] : undefined };
}
function add(stats, x, y, lot, batchIds = [], regionIds = [], rawPairs = 1, group = null) {
  if (stats.points) stats.points.push({ x, y });
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
  if (!["available", "strict"].includes(coverage)) throw new Error("Select a valid Region coverage policy.");
  if (!Number.isInteger(minN) || minN < 3 || minN > 3000000) throw new Error("Minimum N must be an integer from 3 to 3000000.");
  if (!Number.isInteger(minRegions) || minRegions < 1 || minRegions > 16384) throw new Error("Minimum paired Regions must be an integer from 1 to 16384.");
  const expected = new Set();
  clean(options.expectedRegions).split(/[,;\s]+/).filter(Boolean).forEach(token => {
    if (!/^\d{1,8}$/.test(token) || Number(token) < 1) throw new Error("Expected Region IDs must be positive integers, such as 1,2,4.");
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
      output.push({ ...base, method: "Unavailable pair", r: null, n: 0, unit: "n.a.", batches: 0, lots: 0, regions: 0, informativeLots: 0, informativeGroups: 0, rawPairs: 0, status: "Different measurement levels: select explicit Batch-level columns for both parameters." });
      continue;
    }
    if (coverage === "strict" && !scalar && !expected.size) throw new Error("Enter Expected Region IDs before using strict coverage.");
    if (rows.length * ids.length > 3000000) throw new Error("More than 3 million candidate pairs. Filter the data first.");
    const capture = name => accumulator(options.plotMethod === name);
    const raw = capture("Raw pooled"), batch = capture("Batch means"), within = capture("Batch means within Lot"), between = capture("Between Lots"), overall = capture("Overall within Lot x Region");
    const regionStats = new Map(ids.map(id => [id, capture(`Region ${id} raw`)]));
    const adjustedStats = new Map(ids.map(id => [id, capture(`Region ${id} within Lot`)]));
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
      ["raw", "Raw pooled", raw, scalar ? "Batch scalar pairs" : "Region pairs"],
      ["batch", "Batch means", batch, "Batch occurrences"],
      ["within", "Batch means within Lot", within, "Batch residual pairs"],
      ["between", "Between Lots", between, "Lots (one point each)"],
      ...(!scalar ? ids.map(id => ["region", `Region ${id} raw`, regionStats.get(id), "Region pairs"]) : []),
      ...(!scalar ? ids.map(id => ["regionWithin", `Region ${id} within Lot`, adjustedStats.get(id), "Region residual pairs"]) : []),
      ...(!scalar ? [["overall", "Overall within Lot x Region", overall, "Region residual pairs"]] : [])
    ].filter(([id]) => method === "all" || id === method);
    if (!methods.length) {
      output.push({ ...base, method: CORRELATION_METHODS.find(([id]) => id === method)[1], r: null, n: 0, unit: "n.a.", batches: 0, lots: 0, regions: 0, informativeLots: 0, informativeGroups: 0, rawPairs: 0, status: "No Region identity for this method." });
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
        if (signatures.size > 1) status += "; matched Region subsets vary between Batches";
        if (rejected) status += `; ${rejected} Batches ineligible for Batch means`;
      } else status += "; Batch-mean coverage does not filter raw/Region observations";
      if (id === "within") status += `; ${singletonLots} singleton Lots excluded`;
      if (["regionWithin", "overall"].includes(id)) status += `; ${singletonGroups} singleton Lot-Region groups excluded across pair`;
      if (missingIdentity) status += `; ${missingIdentity} source rows excluded: missing Lot/N`;
      if (scalar) status += "; no Region identity; coverage not applicable";
      if (id === "between") status += "; each Lot contributes one pair";
      output.push({ ...base, method: name, r, n: stats.n, unit, batches: stats.batches.size, lots: stats.lots.size, regions: stats.regions.size, rawPairs: stats.rawPairs,
        informativeLots: ["within", "regionWithin", "overall"].includes(id) ? stats.lots.size : 0,
        informativeGroups: stats.groups.size, status,
        ...(stats.points ? { points: stats.points, slope: stats.xx > 0 ? stats.xy / stats.xx : null, intercept: stats.xx > 0 ? stats.y - stats.xy / stats.xx * stats.x : null } : {}) });
    }
  }
  return { results: output, method, coverage, minRegions, minN, expectedRegions: [...expected] };
}
