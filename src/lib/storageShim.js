// DrumTracker.jsx was written against `window.storage`, the Claude Artifacts
// platform's shared/live-synced key-value store. That API doesn't exist in a
// plain browser context, and no equivalent shared-storage service is wired
// up for this standalone app, so this shim backs it with localStorage
// instead. Behavior differs from the original in one important way: writes
// are only visible in the browser that made them — there's no cross-device
// or cross-tab-user sync, so the "locked by <name>" / live-refresh logic in
// DrumTracker becomes inert (harmless, just never triggers).

const PREFIX = "mrf-drum-tracker:";

if (typeof window !== "undefined" && !window.storage) {
  window.storage = {
    async get(key) {
      try {
        const raw = window.localStorage.getItem(PREFIX + key);
        return raw === null ? null : { value: raw };
      } catch (err) {
        console.warn(`storage.get(${key}) failed`, err);
        return null;
      }
    },

    async set(key, value) {
      try {
        window.localStorage.setItem(PREFIX + key, value);
      } catch (err) {
        console.warn(`storage.set(${key}) failed`, err);
        throw err;
      }
    },
  };
}
