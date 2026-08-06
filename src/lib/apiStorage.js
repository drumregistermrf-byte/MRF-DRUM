// Backs window.storage with the MRF Drum Registry API (Node/Express on
// Render, MongoDB Atlas underneath) instead of localStorage, so drum data
// is genuinely shared across every browser/device that opens the site.
// Interface matches storageShim.js exactly: get(key) -> {value: string} | null,
// set(key, value) where value is already a JSON string.

const API_URL = import.meta.env.VITE_API_URL;
const API_KEY = import.meta.env.VITE_API_KEY;

async function apiFetch(path, options = {}) {
  const res = await fetch(`${API_URL}${path}`, {
    ...options,
    headers: {
      "Content-Type": "application/json",
      ...(API_KEY ? { "x-api-key": API_KEY } : {}),
      ...(options.headers || {}),
    },
  });
  if (!res.ok) {
    throw new Error(`storage request failed: ${res.status} ${res.statusText}`);
  }
  return res.json();
}

if (typeof window !== "undefined" && !window.storage) {
  if (!API_URL) {
    throw new Error(
      "VITE_API_URL is not set — cannot reach the MRF Drum Registry API"
    );
  }

  window.storage = {
    async get(key) {
      try {
        const data = await apiFetch(`/api/storage/${encodeURIComponent(key)}`);
        return data.value == null ? null : { value: data.value };
      } catch (err) {
        console.warn(`storage.get(${key}) failed`, err);
        return null;
      }
    },

    async set(key, value) {
      try {
        await apiFetch(`/api/storage/${encodeURIComponent(key)}`, {
          method: "PUT",
          body: JSON.stringify({ value }),
        });
      } catch (err) {
        console.warn(`storage.set(${key}) failed`, err);
        throw err;
      }
    },
  };
}
