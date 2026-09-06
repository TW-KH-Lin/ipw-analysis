const prefix = `ipw-preferences-v1:${new URL(".", import.meta.url).pathname}:`;

export function rememberSettingsEnabled() {
  try { return localStorage.getItem(`${prefix}enabled`) !== "false"; }
  catch { return false; }
}

export function readPreference(key, fallback = null) {
  if (!rememberSettingsEnabled()) return fallback;
  try {
    const value = localStorage.getItem(prefix + key);
    return value === null ? fallback : JSON.parse(value);
  } catch { return fallback; }
}

export function writePreference(key, value) {
  if (!rememberSettingsEnabled()) return;
  try { localStorage.setItem(prefix + key, JSON.stringify(value)); }
  catch { /* Analysis remains usable when device storage is unavailable or full. */ }
}

export function clearPreferences() {
  try {
    for (const key of Object.keys(localStorage)) {
      if (key.startsWith(prefix) && key !== `${prefix}enabled`) localStorage.removeItem(key);
    }
  } catch { /* Device storage may be unavailable in private browsing. */ }
}

export function setRememberSettings(enabled) {
  if (!enabled) clearPreferences();
  try { localStorage.setItem(`${prefix}enabled`, String(Boolean(enabled))); }
  catch { /* The current session still works without persistence. */ }
}

export function datasetPreferenceKey(workbookName, source, headers) {
  // A stable namespace avoids storing the filename itself in device preferences.
  const value = JSON.stringify([workbookName, source, headers]);
  let hash = 2166136261;
  for (let index = 0; index < value.length; index += 1) {
    hash = Math.imul(hash ^ value.charCodeAt(index), 16777619);
  }
  return `dataset:${(hash >>> 0).toString(16)}`;
}
