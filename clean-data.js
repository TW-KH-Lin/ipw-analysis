const ZONES = [1, 2, 3, 4, 5, 6];

export function mapAuswertungHeader(value) {
  const header = String(value ?? "").trim();
  if (!header || /\s(?:Min|Max|MW|StdAbw)\b/i.test(header)) return "";

  const exact = {
    Nummer: "N",
    ChargenNr: "Lot",
    Ziehart: "Type",
    Viskositaet: "Visco.",
    Wassergehalt: "Water"
  };
  if (exact[header]) return exact[header];
  if (/Temperatur/i.test(header)) return "Temp.";
  if (/Raumfeuchte/i.test(header)) return "Humidity";

  const match = header.match(/(?:^|\s)([1-6])$/);
  if (!match) return "";
  const zone = match[1];
  if (/Dicke/i.test(header)) return `Thick_${zone}`;
  if (/Abl/i.test(header)) return `Peel_${zone}`;
  if (/Phenolrotlinienbreite\s*1/i.test(header)) return `PheRW1_${zone}`;
  if (/Phenolrotlinienbreite\s*2/i.test(header)) return `PheRW2_${zone}`;
  if (/Konturl.*Delta\s*1/i.test(header)) return `PheRK1_${zone}`;
  if (/Konturl.*Delta\s*2/i.test(header)) return `PheRK2_${zone}`;
  if (/Migrationszeit/i.test(header)) return `Wicking_${zone}`;
  if (/NaCl/i.test(header)) return `NaCl_${zone}`;
  if (/Durchfluss/i.test(header)) return `Flow_${zone}`;
  if (/Bandseite/i.test(header)) return `Ref_B_${zone}`;
  if (/Luftseite/i.test(header)) return `Ref_A_${zone}`;
  if (/Migr auto quer/i.test(header)) return `Wicking_Q_${zone}`;
  if (/Doppelfront/i.test(header)) return `DPfront_${zone}`;

  const base = header
    .replace(/\s+[1-6]$/, "")
    .replace(/\b(?:IPW|FuE)_?/gi, "")
    .trim()
    .replace(/\s+/g, "_");
  return base ? `${base}_${zone}` : "";
}

export function buildCleanDataFromAuswertung(table, options = {}) {
  const headerRowIndex = findAuswertungHeaderRow(table);
  if (headerRowIndex < 0) throw new Error("Auswertung headers Nummer and ChargenNr were not found.");
  const rawHeaders = table[headerRowIndex].map((value) => String(value ?? "").trim());
  const mapped = rawHeaders
    .map((header, index) => ({ index, name: mapAuswertungHeader(header) }))
    .filter((item) => item.name);
  if (!mapped.some((item) => item.name === "N") || !mapped.some((item) => item.name === "Lot")) {
    throw new Error("Auswertung must contain Nummer and ChargenNr.");
  }

  const probentypIndex = rawHeaders.findIndex((header) => header.toLowerCase() === "probentyp");
  const lotSourceIndex = mapped.find((item) => item.name === "Lot").index;
  const typeSourceIndex = mapped.find((item) => item.name === "Type")?.index ?? -1;
  const baseHeaders = mapped.map((item) => item.name);
  const finalHeaders = addClassificationAndStatistics(baseHeaders);
  const classificationByLot = options.classificationByLot || new Map();
  const classificationByKey = options.classificationByKey || new Map();
  const sourceRows = table.slice(headerRowIndex + 1);
  const veLots = new Set();
  let removedProbeRows = 0;
  let removedVeRows = 0;
  const rows = [];

  for (const rawRow of sourceRows) {
    const probeValue = rawRow[probentypIndex >= 0 ? probentypIndex : 3];
    if (isProbe(probeValue)) continue;
    if (typeSourceIndex >= 0 && cleanText(rawRow[typeSourceIndex]).toUpperCase() === "VE") {
      const lot = cleanText(rawRow[lotSourceIndex]);
      if (lot) veLots.add(lot);
    }
  }

  for (const rawRow of sourceRows) {
    const probeValue = rawRow[probentypIndex >= 0 ? probentypIndex : 3];
    if (isProbe(probeValue)) {
      removedProbeRows += 1;
      continue;
    }
    const sourceLot = cleanText(rawRow[lotSourceIndex]);
    if (veLots.has(sourceLot)) {
      removedVeRows += 1;
      continue;
    }
    const values = mapped.map((item) => rawRow[item.index] ?? "");
    const baseRow = Object.fromEntries(baseHeaders.map((header, index) => [header, values[index]]));
    const lot = cleanText(baseRow.Lot);
    const batch = cleanText(baseRow.N);
    if (!lot) continue;
    const byLot = classificationByLot.get(lot);
    const byKey = classificationByKey.get(`${lot}\u0000${batch}`);
    baseRow.Classification = byLot !== undefined ? byLot : (byKey ?? "");
    rows.push(finalHeaders.map((header) => statisticValue(baseRow, header)));
  }

  const lotIndex = finalHeaders.indexOf("Lot");
  const batchIndex = finalHeaders.indexOf("N");
  rows.sort((a, b) => compareNatural(a[lotIndex], b[lotIndex]) || compareNatural(a[batchIndex], b[batchIndex]));
  return {
    headers: finalHeaders,
    rows,
    removedProbeRows,
    removedVeRows,
    removedVeLots: veLots.size
  };
}

export function correctCleanData(headers, rows, referenceTemperature, referenceHumidity) {
  const tempRef = Number(referenceTemperature);
  const humidityRef = Number(referenceHumidity);
  if (!Number.isFinite(tempRef) || !Number.isFinite(humidityRef)) {
    throw new Error("Reference temperature and humidity must be numeric.");
  }
  const tempIndex = findHeader(headers, "Temp.");
  const humidityIndex = findHeader(headers, "Humidity");
  if (tempIndex < 0 || humidityIndex < 0) {
    throw new Error("Clean data must contain Temp. and Humidity.");
  }

  const corrected = rows.map((sourceRow) => {
    const row = [...sourceRow];
    const actualTemp = numeric(row[tempIndex]);
    const actualHumidity = numeric(row[humidityIndex]);
    if (actualTemp === null || actualHumidity === null) return row;
    const factor = 1 + 0.03 * (tempRef - actualTemp) + 0.005 * (humidityRef - actualHumidity);
    if (!Number.isFinite(factor) || factor === 0) return row;
    for (const base of ["Wicking", "Wicking_Q"]) {
      const zoneIndexes = ZONES.map((zone) => findHeader(headers, `${base}_${zone}`));
      for (const index of zoneIndexes) {
        if (index < 0) continue;
        const value = numeric(row[index]);
        if (value !== null) row[index] = value / factor;
      }
      recalculateRowStatistics(headers, row, base, zoneIndexes);
    }
    return row;
  });
  return { headers: [...headers], rows: corrected };
}

export function buildFullSummary(headers, rows) {
  const lotIndex = findHeader(headers, "Lot");
  const batchIndex = findHeader(headers, "N");
  const typeIndex = findHeader(headers, "Type");
  const classificationIndex = findHeader(headers, "Classification");
  if (lotIndex < 0 || batchIndex < 0) throw new Error("Clean data must contain Lot and N.");

  const groups = new Map();
  for (const row of rows) {
    const lot = cleanText(row[lotIndex]);
    if (!lot) continue;
    if (!groups.has(lot)) groups.set(lot, []);
    groups.get(lot).push(row);
  }

  const summaryRows = [];
  for (const [lot, group] of groups) {
    const summary = headers.map((header, columnIndex) => {
      if (columnIndex === lotIndex) return lot;
      if (columnIndex === batchIndex) {
        return new Set(group.map((row) => cleanText(row[batchIndex])).filter(Boolean)).size;
      }
      if (columnIndex === typeIndex) return firstText(group, typeIndex);
      if (columnIndex === classificationIndex) return firstText(group, classificationIndex);
      const statistic = header.match(/^(.*)_(MW|SD)$/i);
      if (statistic) {
        const values = collectZoneValues(headers, group, statistic[1]);
        return statistic[2].toUpperCase() === "MW" ? average(values) : sampleStandardDeviation(values);
      }
      const values = group.map((row) => numeric(row[columnIndex])).filter((value) => value !== null);
      return values.length ? average(values) : "";
    });
    summaryRows.push(summary);
  }
  return { headers: [...headers], rows: summaryRows };
}

export function classificationMaps(headers, rows) {
  const lotIndex = findHeader(headers, "Lot");
  const batchIndex = findHeader(headers, "N");
  const classificationIndex = findHeader(headers, "Classification");
  const byLot = new Map();
  const byKey = new Map();
  if (lotIndex < 0 || classificationIndex < 0) return { byLot, byKey };
  for (const row of rows) {
    const lot = cleanText(row[lotIndex]);
    const classification = cleanText(row[classificationIndex]);
    if (!lot || !classification) continue;
    if (!byLot.has(lot)) byLot.set(lot, classification);
    if (batchIndex >= 0) byKey.set(`${lot}\u0000${cleanText(row[batchIndex])}`, classification);
  }
  return { byLot, byKey };
}

function findAuswertungHeaderRow(table) {
  return table.slice(0, 30).findIndex((row) => {
    const values = row.map((value) => String(value ?? "").trim());
    return values.includes("Nummer") && values.includes("ChargenNr");
  });
}

function addClassificationAndStatistics(headers) {
  const completeBases = new Set();
  for (const header of headers) {
    const match = header.match(/^(.*)_1$/);
    if (match && ZONES.every((zone) => headers.includes(`${match[1]}_${zone}`))) completeBases.add(match[1]);
  }
  const output = [];
  for (const header of headers) {
    if (header === "N") output.push("Classification");
    output.push(header);
    const match = header.match(/^(.*)_6$/);
    if (match && completeBases.has(match[1])) output.push(`${match[1]}_MW`, `${match[1]}_SD`);
  }
  return output;
}

function statisticValue(row, header) {
  if (Object.hasOwn(row, header)) return row[header];
  const match = header.match(/^(.*)_(MW|SD)$/);
  if (!match) return "";
  const values = ZONES.map((zone) => numeric(row[`${match[1]}_${zone}`])).filter((value) => value !== null);
  return match[2] === "MW" ? average(values) : sampleStandardDeviation(values);
}

function recalculateRowStatistics(headers, row, base, zoneIndexes) {
  const values = zoneIndexes.filter((index) => index >= 0).map((index) => numeric(row[index])).filter((value) => value !== null);
  const meanIndex = findHeader(headers, `${base}_MW`);
  const sdIndex = findHeader(headers, `${base}_SD`);
  if (meanIndex >= 0) row[meanIndex] = average(values);
  if (sdIndex >= 0) row[sdIndex] = sampleStandardDeviation(values);
}

function collectZoneValues(headers, rows, base) {
  const indexes = ZONES.map((zone) => findHeader(headers, `${base}_${zone}`)).filter((index) => index >= 0);
  return rows.flatMap((row) => indexes.map((index) => numeric(row[index])).filter((value) => value !== null));
}

function average(values) {
  return values.length ? values.reduce((sum, value) => sum + value, 0) / values.length : "";
}

function sampleStandardDeviation(values) {
  if (values.length < 2) return "";
  const mean = average(values);
  return Math.sqrt(values.reduce((sum, value) => sum + (value - mean) ** 2, 0) / (values.length - 1));
}

function numeric(value) {
  if (value === null || value === undefined || value === "") return null;
  const number = Number(value);
  return Number.isFinite(number) ? number : null;
}

function findHeader(headers, name) {
  const expected = name.toLowerCase();
  return headers.findIndex((header) => String(header ?? "").trim().toLowerCase() === expected);
}

function firstText(rows, index) {
  return rows.map((row) => cleanText(row[index])).find(Boolean) || "";
}

function cleanText(value) {
  return String(value ?? "").trim();
}

function isProbe(value) {
  return /^PROBEN?$/i.test(cleanText(value));
}

function compareNatural(a, b) {
  return cleanText(a).localeCompare(cleanText(b), undefined, { numeric: true, sensitivity: "base" });
}
