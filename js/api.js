/** Guardado local + cola opcional a Apps Script */
window.APP = window.APP || {};

APP.API = (() => {
  const KEY = "app_cosecha_records";
  const SESSION_KEY = "app_cosecha_session";

  function todayKey() {
    try {
      return new Intl.DateTimeFormat("en-CA", { timeZone: APP.CONFIG.TZ, year: "numeric", month: "2-digit", day: "2-digit" }).format(new Date());
    } catch {
      return new Date().toISOString().slice(0, 10);
    }
  }

  function uid() {
    if (crypto.randomUUID) return crypto.randomUUID();
    return "c" + Date.now() + Math.random().toString(16).slice(2);
  }

  function all() {
    try {
      const list = JSON.parse(localStorage.getItem(KEY) || "[]");
      return Array.isArray(list) ? list : [];
    } catch {
      return [];
    }
  }

  function write(list) {
    localStorage.setItem(KEY, JSON.stringify(list.slice(0, 400)));
    window.dispatchEvent(new Event("app:activity"));
  }

  function localRecords(fecha) {
    const day = fecha || todayKey();
    const byLote = {};
    all().forEach((r) => {
      if (String(r.fecha) !== day) return;
      byLote[String(r.lote)] = r;
    });
    return Object.values(byLote).sort((a, b) => String(a.lote).localeCompare(String(b.lote), "es", { numeric: true }));
  }

  async function submit(payload) {
    const data = payload.data || {};
    const rec = {
      ...data,
      clientId: payload.clientId || uid(),
      submittedAt: payload.submittedAt || new Date().toISOString(),
    };
    const list = all().filter((r) => !(r.fecha === rec.fecha && String(r.lote) === String(rec.lote)));
    list.unshift(rec);
    write(list);
    return { ok: true, mode: "local", clientId: rec.clientId };
  }

  function loadSession() {
    try {
      return JSON.parse(localStorage.getItem(SESSION_KEY) || "null");
    } catch {
      return null;
    }
  }

  function saveSession(s) {
    localStorage.setItem(SESSION_KEY, JSON.stringify(s || {}));
  }

  return { todayKey, newClientId: uid, localRecords, submit, loadSession, saveSession, pendingCount: () => 0 };
})();
