import { getRegionalParameters, headerIndex, isNumeric, number, text, zoneColumns } from "./analysis.js?v=12";

const MASTER_ROLL_ZONE_ENDS = {
  9: {
    25: [7, 14, 21, 28, 35, 42],
    20: [8, 17, 25, 34, 42, 51],
    18: [9, 19, 28, 38, 47, 57]
  },
  10: {
    25: [7, 15, 22, 30, 37, 45],
    20: [9, 19, 28, 38, 47, 57]
  },
  17: {
    25: [10, 20, 30, 40, 50, 60]
  }
};

export function getMasterRollWidths(machineId) {
  const widths = MASTER_ROLL_ZONE_ENDS[Number(machineId)];
  return widths ? Object.keys(widths).map(Number).sort((a, b) => b - a) : [];
}

export function getMasterRollZone(machineId, rollWidth, rollNo) {
  const machine = MASTER_ROLL_ZONE_ENDS[Number(machineId)];
  if (!machine) return "Unknown Machine";
  const zoneEnds = machine[Number(rollWidth)];
  if (!zoneEnds) return "Invalid Width";
  const roll = Number(rollNo);
  if (!Number.isInteger(roll) || roll < 1) return "Invalid Roll";
  const zoneIndex = zoneEnds.findIndex((end) => roll <= end);
  return zoneIndex >= 0 ? `Zone ${zoneIndex + 1}` : "Invalid Roll";
}

export function findRawHeaderRow(table) {
  const index = table.slice(0, 30).findIndex((row) =>
    rawHeaderIndex(row, "ChargenNr") >= 0 && rawHeaderIndex(row, "Nummer") >= 0
  );
  return index >= 0 ? { index, headers: table[index] } : null;
}

export function mergeAuswertungTables(existingTable, importTable) {
  const existingHeader = findRawHeaderRow(existingTable);
  const importHeader = findRawHeaderRow(importTable);
  if (!existingHeader) throw new Error("The current workbook has no readable Auswertung headers.");
  if (!importHeader) throw new Error("The new-lot workbook needs ChargenNr and Nummer headers.");

  const existingHeaders = existingHeader.headers;
  const importHeaders = importHeader.headers;
  const importColumns = new Map();
  importHeaders.forEach((header, index) => {
    const key = rawHeaderKey(header);
    if (key && !importColumns.has(key)) importColumns.set(key, index);
  });
  const missingHeaders = existingHeaders
    .map((header) => text(header))
    .filter((header) => header && !importColumns.has(rawHeaderKey(header)));
  const columnMap = existingHeaders.map((header) => importColumns.get(rawHeaderKey(header)) ?? -1);

  const existingLotColumn = rawHeaderIndex(existingHeaders, "ChargenNr");
  const existingBatchColumn = rawHeaderIndex(existingHeaders, "Nummer");
  const existingTypeColumn = rawHeaderIndex(existingHeaders, "Ziehart");
  const importLotColumn = rawHeaderIndex(importHeaders, "ChargenNr");
  const importBatchColumn = rawHeaderIndex(importHeaders, "Nummer");
  const importTypeColumn = rawHeaderIndex(importHeaders, "Ziehart");
  const existingRows = existingTable.slice(existingHeader.index + 1);
  const importRows = importTable.slice(importHeader.index + 1);
  const keys = new Set();
  for (const row of existingRows) {
    const lot = text(row[existingLotColumn]);
    const batch = text(row[existingBatchColumn]);
    if (!lot || !batch) continue;
    keys.add(importKey(lot, batch, existingTypeColumn >= 0 ? row[existingTypeColumn] : ""));
  }

  const appendedRows = [];
  const lots = new Set();
  let duplicateRows = 0;
  let incompleteRows = 0;
  let emptyRows = 0;
  for (const row of importRows) {
    const lot = text(row[importLotColumn]);
    const batch = text(row[importBatchColumn]);
    const type = importTypeColumn >= 0 ? row[importTypeColumn] : "";
    if (!lot && !batch) {
      emptyRows += 1;
      continue;
    }
    if (!lot || !batch) {
      incompleteRows += 1;
      continue;
    }
    const key = importKey(lot, batch, type);
    if (keys.has(key)) {
      duplicateRows += 1;
      continue;
    }
    if (missingHeaders.length) continue;
    appendedRows.push(columnMap.map((column) => column >= 0 ? row[column] ?? "" : ""));
    keys.add(key);
    lots.add(lot);
  }

  return {
    table: [
      ...existingTable.slice(0, existingHeader.index + 1),
      ...existingRows,
      ...appendedRows
    ],
    appendedRows,
    existingHeaderIndex: existingHeader.index,
    importHeaderIndex: importHeader.index,
    addedRows: appendedRows.length,
    duplicateRows,
    incompleteRows,
    emptyRows,
    candidateRows: importRows.length - emptyRows,
    lots: [...lots].sort((a, b) => a.localeCompare(b, undefined, { numeric: true })),
    missingHeaders
  };
}

const PARAMETER_IMPORT_IDENTIFIERS = new Set([
  "messungid", "probenzeit", "nummer", "probentyp", "probetype", "type",
  "ziehmaschinenname", "chargennr", "ziehart", "filtertypnr"
]);

function parameterImportColumns(headers) {
  const lot = rawHeaderIndex(headers, "ChargenNr");
  const batch = rawHeaderIndex(headers, "Nummer");
  const sampleType = ["Probentyp", "ProbeType", "Type"].map(name => rawHeaderIndex(headers, name)).find(index => index >= 0) ?? -1;
  if (lot < 0 || batch < 0 || sampleType < 0) throw new Error("Both workbooks need ChargenNr, Nummer and Probentyp (Probe Type) columns.");
  return { lot, batch, sampleType };
}

function parameterImportRowKey(row, columns) {
  const lot=text(row[columns.lot]),batch=text(row[columns.batch]),sampleType=text(row[columns.sampleType]);
  return lot&&batch&&sampleType?importKey(lot,batch,sampleType):"";
}

function parameterImportValuesEqual(first,second) {
  const comparableNumber=value=>{
    if(isNumeric(value))return number(value);
    const rendered=text(value);
    return /^[-+]?\d+(?:,\d+)?$/.test(rendered)?Number(rendered.replace(",",".")):null;
  };
  const a=comparableNumber(first),b=comparableNumber(second);
  if(a!==null&&b!==null){
    const scale=Math.max(1,Math.abs(a),Math.abs(b));
    return Math.abs(a-b)<=1e-9*scale;
  }
  return text(first)===text(second);
}

function uniqueImportRows(rows, columns) {
  const map=new Map(),duplicates=new Set();
  for(const row of rows){
    const key=parameterImportRowKey(row,columns);if(!key)continue;
    if(map.has(key)){duplicates.add(key);map.delete(key);}
    else if(!duplicates.has(key))map.set(key,row);
  }
  return {map,duplicates};
}

export function analyzeAuswertungParameterImport(existingTable, importTable) {
  const existingHeader=findRawHeaderRow(existingTable),importHeader=findRawHeaderRow(importTable);
  if(!existingHeader)throw new Error("The current workbook has no readable Auswertung headers.");
  if(!importHeader)throw new Error("The parameter workbook needs ChargenNr and Nummer headers.");
  const existingHeaders=existingHeader.headers,importHeaders=importHeader.headers;
  const existingColumns=parameterImportColumns(existingHeaders),importColumns=parameterImportColumns(importHeaders);
  const existingByKey=new Map();
  for(const [offset,row] of existingTable.slice(existingHeader.index+1).entries()){
    const key=parameterImportRowKey(row,existingColumns);if(!key)continue;
    if(!existingByKey.has(key))existingByKey.set(key,[]);
    existingByKey.get(key).push({row,index:existingHeader.index+1+offset});
  }
  const importRows=importTable.slice(importHeader.index+1),unique=uniqueImportRows(importRows,importColumns);
  const parameters=importHeaders.flatMap((header,sourceColumn)=>{
    const name=text(header),key=rawHeaderKey(name);
    if(!name||PARAMETER_IMPORT_IDENTIFIERS.has(key))return [];
    const targetColumn=existingHeaders.findIndex(value=>rawHeaderKey(value)===key);
    let sourceValues=0,matchedValues=0,targetValues=0,conflicts=0;
    for(const [rowKey,sourceRow] of unique.map){
      const sourceValue=sourceRow[sourceColumn];if(text(sourceValue)==="")continue;
      sourceValues++;
      const targets=existingByKey.get(rowKey)||[];
      for(const target of targets){
        matchedValues++;
        if(targetColumn<0||text(target.row[targetColumn])==="")continue;
        targetValues++;
        if(!parameterImportValuesEqual(target.row[targetColumn],sourceValue))conflicts++;
      }
    }
    return [{name,key,sourceColumn,targetColumn,sourceValues,matchedValues,targetValues,conflicts,exists:targetColumn>=0}];
  });
  const matchedKeys=[...unique.map.keys()].filter(key=>existingByKey.has(key));
  return {
    existingHeaderIndex:existingHeader.index,importHeaderIndex:importHeader.index,
    matchedRows:matchedKeys.reduce((sum,key)=>sum+existingByKey.get(key).length,0),
    matchedKeys:matchedKeys.length,sourceKeys:unique.map.size,duplicateSourceKeys:unique.duplicates.size,
    unmatchedSourceKeys:unique.map.size-matchedKeys.length,
    parameters:parameters.sort((a,b)=>a.name.localeCompare(b.name,undefined,{numeric:true})),
    existingHeaders,importHeaders
  };
}

export function applyAuswertungParameterImport(existingTable, importTable, selectedHeaders, options={}) {
  const analysis=analyzeAuswertungParameterImport(existingTable,importTable);
  const selectedKeys=new Set((selectedHeaders||[]).map(rawHeaderKey));
  const selected=analysis.parameters.filter(parameter=>selectedKeys.has(parameter.key));
  if(!selected.length)throw new Error("Select at least one parameter to import.");
  const table=existingTable.map(row=>[...row]);
  const existingHeader=findRawHeaderRow(table),importHeader=findRawHeaderRow(importTable);
  const existingColumns=parameterImportColumns(existingHeader.headers),importColumns=parameterImportColumns(importHeader.headers);
  const unique=uniqueImportRows(importTable.slice(importHeader.index+1),importColumns);
  const addedColumns=[];
  selected.forEach(parameter=>{
    if(parameter.targetColumn>=0)return;
    parameter.targetColumn=existingHeader.headers.length;
    existingHeader.headers.push(parameter.name);table[existingHeader.index]=existingHeader.headers;
    addedColumns.push(parameter.name);
  });
  let updatedCells=0,skippedExisting=0,matchedRows=0;
  const changes=[];
  table.slice(existingHeader.index+1).forEach((row,offset)=>{
    const sourceRow=unique.map.get(parameterImportRowKey(row,existingColumns));if(!sourceRow)return;
    matchedRows++;
    selected.forEach(parameter=>{
      const value=sourceRow[parameter.sourceColumn];if(text(value)==="")return;
      const current=row[parameter.targetColumn];
      if(text(current)!==""&&!options.overwrite){skippedExisting++;return;}
      if(parameterImportValuesEqual(current,value))return;
      row[parameter.targetColumn]=value;updatedCells++;
      changes.push({row:existingHeader.index+1+offset,column:parameter.targetColumn,value,newColumn:addedColumns.includes(parameter.name)});
    });
  });
  return {table,changes,updatedCells,skippedExisting,matchedRows,addedColumns,selectedHeaders:selected.map(parameter=>parameter.name),selectedParameters:selected.map(parameter=>({name:parameter.name,column:parameter.targetColumn})),existingHeaderIndex:existingHeader.index,analysis};
}

export function buildDataLabel(headers, rows, input) {
  const lotColumn = headerIndex(headers, "Lot");
  const batchColumn = headerIndex(headers, "N");
  if (lotColumn < 0 || batchColumn < 0) throw new Error("Clean_Data must contain Lot and N columns.");
  const zone = Number(input.zone);
  if (!Number.isInteger(zone) || zone < 1 || zone > 6) throw new Error("Select Zone 1 through Zone 6.");
  const parameters = [...new Set((input.parameters || []).map(text).filter(Boolean))];
  if (!parameters.length) throw new Error("Select at least one parameter.");
  const row = rows.find((item) => valuesMatch(item[lotColumn], input.lot) && valuesMatch(item[batchColumn], input.batch));
  if (!row) throw new Error("No matching Lot and original Batch N were found in Clean_Data.");
  const values = {};
  for (const parameter of parameters) {
    const columns = zoneColumns(headers, parameter);
    const column = columns?.[zone - 1] ?? -1;
    if (column < 0) throw new Error(`${parameter}_${zone} is missing.`);
    values[parameter] = isNumeric(row[column]) ? number(row[column]) : null;
  }
  return {
    lot: row[lotColumn],
    batch: row[batchColumn],
    zone,
    parameters,
    values,
    valuesText: parameters.map((parameter) => `${parameter}=${formatLabelValue(values[parameter])}`).join("; "),
    label: text(input.label),
    comment: text(input.comment),
    notes: text(input.notes),
    updated: input.updated || new Date().toISOString()
  };
}

export function upsertDataLabel(labels, label) {
  const key = dataLabelKey(label.lot, label.batch, label.zone);
  const index = labels.findIndex((item) => dataLabelKey(item.lot, item.batch, item.zone) === key);
  if (index >= 0) return labels.map((item, itemIndex) => itemIndex === index ? label : item);
  return [...labels, label];
}

export function removeDataLabel(labels, lot, batch, zone) {
  const key = dataLabelKey(lot, batch, zone);
  return labels.filter((item) => dataLabelKey(item.lot, item.batch, item.zone) !== key);
}

export function findDataLabel(labels, lot, batch, zone) {
  const key = dataLabelKey(lot, batch, zone);
  return labels.find((item) => dataLabelKey(item.lot, item.batch, item.zone) === key) || null;
}

export function parseDataLabels(table) {
  const headerIndexValue = table.findIndex((row) =>
    rawHeaderIndex(row, "Lot") >= 0 && rawHeaderIndex(row, "Batch N") >= 0 &&
    rawHeaderIndex(row, "Zone") >= 0 && rawHeaderIndex(row, "Parameters") >= 0
  );
  if (headerIndexValue < 0) return [];
  const headers = table[headerIndexValue];
  const columns = {
    lot: rawHeaderIndex(headers, "Lot"),
    batch: rawHeaderIndex(headers, "Batch N"),
    zone: rawHeaderIndex(headers, "Zone"),
    parameters: rawHeaderIndex(headers, "Parameters"),
    values: rawHeaderIndex(headers, "Values"),
    label: rawHeaderIndex(headers, "Label"),
    comment: rawHeaderIndex(headers, "Comment"),
    updated: rawHeaderIndex(headers, "Updated"),
    notes: rawHeaderIndex(headers, "Notes")
  };
  return table.slice(headerIndexValue + 1).flatMap((row) => {
    const lot = row[columns.lot];
    const batch = row[columns.batch];
    const zone = Number(row[columns.zone]);
    const parameters = text(row[columns.parameters]).split(",").map(text).filter(Boolean);
    if (!text(lot) || !text(batch) || !Number.isInteger(zone) || zone < 1 || zone > 6 || !parameters.length) return [];
    return [{
      lot,
      batch,
      zone,
      parameters,
      values: {},
      valuesText: columns.values >= 0 ? text(row[columns.values]) : "",
      label: columns.label >= 0 ? text(row[columns.label]) : "",
      comment: columns.comment >= 0 ? text(row[columns.comment]) : "",
      notes: columns.notes >= 0 ? text(row[columns.notes]) : "",
      updated: columns.updated >= 0 ? dateText(row[columns.updated]) : ""
    }];
  });
}

export function buildDataLabelsTable(labels) {
  const rows = Array.from({ length: 10 }, () => []);
  rows[0] = ["Quick Data Highlighter"];
  rows.push(["Lot", "Batch N", "Zone", "Parameters", "Values", "Label", "Comment", "Updated", "Notes"]);
  for (const item of labels) {
    rows.push([
      item.lot,
      item.batch,
      item.zone,
      item.parameters.join(", "),
      item.valuesText || item.parameters.map((parameter) => `${parameter}=${formatLabelValue(item.values?.[parameter])}`).join("; "),
      item.label,
      item.comment,
      item.updated ? new Date(item.updated) : "",
      item.notes
    ]);
  }
  return rows;
}

export function buildDataLabelValuesTable(headers, rows, labels) {
  const parameters = getRegionalParameters(headers).filter((parameter) =>
    zoneColumns(headers, parameter)?.every((column) => column >= 0)
  );
  const lotColumn = headerIndex(headers, "Lot");
  const batchColumn = headerIndex(headers, "N");
  const output = [["Lot", "Batch N", "Zone", "Label", "Comment", "Updated", "Notes", ...parameters]];
  for (const item of labels) {
    const row = rows.find((candidate) =>
      valuesMatch(candidate[lotColumn], item.lot) && valuesMatch(candidate[batchColumn], item.batch)
    );
    output.push([
      item.lot,
      item.batch,
      item.zone,
      item.label,
      item.comment,
      item.updated ? new Date(item.updated) : "",
      item.notes,
      ...parameters.map((parameter) => {
        if (!item.parameters.includes(parameter) || !row) return "";
        const column = zoneColumns(headers, parameter)?.[item.zone - 1] ?? -1;
        return column >= 0 && isNumeric(row[column]) ? number(row[column]) : "";
      })
    ]);
  }
  return output;
}

export function dataLabelCellCoordinates(headers, rows, labels) {
  const lotColumn = headerIndex(headers, "Lot");
  const batchColumn = headerIndex(headers, "N");
  if (lotColumn < 0 || batchColumn < 0) return [];
  const coordinates = [];
  for (const item of labels) {
    const row = rows.findIndex((candidate) =>
      valuesMatch(candidate[lotColumn], item.lot) && valuesMatch(candidate[batchColumn], item.batch)
    );
    if (row < 0) continue;
    const columns = new Set([lotColumn, batchColumn]);
    for (const parameter of item.parameters) {
      const column = zoneColumns(headers, parameter)?.[item.zone - 1] ?? -1;
      if (column >= 0) columns.add(column);
    }
    for (const column of columns) coordinates.push({ row, column });
  }
  return coordinates;
}

export function dataLabelKey(lot, batch, zone) {
  return `${labelIdentity(lot)}|${labelIdentity(batch)}|${Number(zone)}`;
}

function rawHeaderIndex(headers, name) {
  const wanted = rawHeaderKey(name);
  return headers.findIndex((header) => rawHeaderKey(header) === wanted);
}

function rawHeaderKey(value) {
  return text(value).normalize("NFKC").toLowerCase().replace(/[\s._-]+/g, "");
}

function importKey(lot, batch, type) {
  return `${text(lot).toUpperCase()}\u001e${text(batch).toUpperCase()}\u001e${text(type).toUpperCase()}`;
}

function labelIdentity(value) {
  const rendered = text(value);
  return rendered && Number.isFinite(Number(rendered)) ? String(Number(rendered)) : rendered.toUpperCase();
}

function valuesMatch(first, second) {
  if (isNumeric(first) && isNumeric(second)) return number(first) === number(second);
  return text(first).toUpperCase() === text(second).toUpperCase();
}

function formatLabelValue(value) {
  if (!Number.isFinite(value)) return "blank";
  const rendered = Number(value).toFixed(4).replace(/0+$/, "").replace(/\.$/, "");
  return rendered.includes(".") ? rendered : `${rendered}.0`;
}

function dateText(value) {
  if (value instanceof Date && Number.isFinite(value.getTime())) return value.toISOString();
  const parsed = Date.parse(text(value));
  return Number.isFinite(parsed) ? new Date(parsed).toISOString() : text(value);
}
