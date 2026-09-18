export const VERSION = 1;
export const CASE_FIELDS = ['complaintNo', 'lot', 'membraneType', 'materialNo', 'standardizedSymptoms', 'customerReportedFailure', 'mrfrAreas', 'mrfrCombined', 'complaintRegisteredDate', 'reportDate'];
export function selectedCase(record) {
  const result = {};
  for (const key of CASE_FIELDS) {
    const value = record[key];
    if (value !== undefined && value !== null && typeof value !== 'string') throw new Error(`Invalid ${key}`);
    result[key] = value || '';
    if (result[key].length > 4000) throw new Error(`${key} is too long`);
  }
  if (!result.complaintNo.trim() || !result.lot.trim()) throw new Error('Complaint number and Lot are required.');
  return result;
}
export function matchLots(value, lots) {
  const raw = value.trim();
  if (lots.includes(raw)) return { matched: [raw], unmatched: [], ambiguous: false };
  // Only explicit separators are supported. Never strip leading zeros or infer ranges.
  const tokens = [...new Set(raw.split(/[,;\n]+/).map(v => v.trim()).filter(Boolean))];
  const matched = tokens.filter(v => lots.includes(v));
  return { matched, unmatched: tokens.filter(v => !lots.includes(v)), ambiguous: tokens.length === 1 && !matched.length };
}
