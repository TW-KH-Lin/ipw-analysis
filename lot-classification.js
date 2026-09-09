// Update only classification cells, retaining measurement cells and sheet layout.
export function updateWorkbookClassifications(workbook, changes, XLSX) {
  if (!changes.size) return 0;
  let updated = 0;
  for (const name of ["Clean_Data", "Clean_Data_Cor", "Summary", "Summary_Filtered"]) {
    const sheet = workbook.Sheets[name];
    if (!sheet?.["!ref"]) continue;
    const range = XLSX.utils.decode_range(sheet["!ref"]);
    let headerRow = -1;
    let lotColumn = -1;
    let classificationColumn = -1;
    for (let r = range.s.r; r <= Math.min(range.e.r, range.s.r + 50); r++) {
      for (let c = range.s.c; c <= range.e.c; c++) {
        const value = String(sheet[XLSX.utils.encode_cell({ r, c })]?.v ?? "").trim().toLowerCase();
        if (value === "lot") lotColumn = c;
        if (value === "classification") classificationColumn = c;
      }
      if (lotColumn >= 0) { headerRow = r; break; }
      classificationColumn = -1;
    }
    if (headerRow < 0) continue;
    let hasEditedLot = false;
    for (let r = headerRow + 1; r <= range.e.r; r++) {
      if (changes.has(String(sheet[XLSX.utils.encode_cell({ r, c: lotColumn })]?.v ?? "").trim())) hasEditedLot = true;
    }
    if (!hasEditedLot) continue;
    if (classificationColumn < 0) {
      classificationColumn = ++range.e.c;
      sheet[XLSX.utils.encode_cell({ r: headerRow, c: classificationColumn })] = { t: "s", v: "Classification" };
      sheet["!ref"] = XLSX.utils.encode_range(range);
    }
    for (let r = headerRow + 1; r <= range.e.r; r++) {
      const lot = String(sheet[XLSX.utils.encode_cell({ r, c: lotColumn })]?.v ?? "").trim();
      if (!changes.has(lot)) continue;
      const address = XLSX.utils.encode_cell({ r, c: classificationColumn });
      const cell = sheet[address];
      if (cell?.f || cell?.t === "e") throw new Error(`Classification in ${name}!${address} is a formula or error. No workbook was saved.`);
      if (sheet["!merges"]?.some((merge) => r >= merge.s.r && r <= merge.e.r && classificationColumn >= merge.s.c && classificationColumn <= merge.e.c)) {
        throw new Error(`Classification in ${name}!${address} is merged. No workbook was saved.`);
      }
      sheet[address] = { ...cell, t: "s", v: changes.get(lot) };
      delete sheet[address].w;
      delete sheet[address].h;
      delete sheet[address].r;
      updated++;
    }
  }
  return updated;
}
