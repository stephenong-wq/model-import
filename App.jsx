// Polyfills the window.storage API that the Model Audit Tool was originally
// built against (Claude's artifact persistent-storage feature) so it works
// unmodified in a normal browser deployment. Backed by localStorage, scoped
// under a single prefix. "shared" is accepted for interface compatibility
// but has no effect here — there's no multi-user backend in this app.
const PREFIX = "orion-tools:";

function fullKey(key) {
  return PREFIX + key;
}

if (typeof window !== "undefined" && !window.storage) {
  window.storage = {
    async get(key) {
      try {
        const value = localStorage.getItem(fullKey(key));
        return value !== null ? { key, value, shared: false } : null;
      } catch {
        return null;
      }
    },
    async set(key, value) {
      try {
        localStorage.setItem(fullKey(key), value);
        return { key, value, shared: false };
      } catch {
        return null;
      }
    },
    async delete(key) {
      try {
        const existed = localStorage.getItem(fullKey(key)) !== null;
        localStorage.removeItem(fullKey(key));
        return { key, deleted: existed, shared: false };
      } catch {
        return null;
      }
    },
    async list(prefix = "") {
      try {
        const keys = [];
        for (let i = 0; i < localStorage.length; i++) {
          const k = localStorage.key(i);
          if (k && k.startsWith(PREFIX + prefix)) keys.push(k.slice(PREFIX.length));
        }
        return { keys, prefix, shared: false };
      } catch {
        return null;
      }
    },
  };
}
