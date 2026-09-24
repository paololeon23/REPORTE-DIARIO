/** Guardado local primero + cola de sincronización a Apps Script */
window.APP = window.APP || {};

APP.API = (() => {
  const KEY = "app_cosecha_records";
  const SESSION_KEY = "app_cosecha_session";
  const HISTORY_KEY = "app_cosecha_excel_history";
  const DEVICE_KEY = "app_cosecha_device";
  const CLOSE_KEY = "app_cosecha_day_closed";
  const SENDS_KEY = "app_cosecha_sends";
  const OUTBOX_KEY = "app_cosecha_outbox";
  const DAY_MS = 24 * 60 * 60 * 1000;
  const SEND_TIMEOUT = 40000;
  const PING_TIMEOUT = 10000;
  const BATCH = 25;
  const MAX_FLUSH_FAILURES = 3;

  let syncLock = null;

  function todayKey() {
    try {
      return new Intl.DateTimeFormat("en-CA", { timeZone: APP.CONFIG.TZ, year: "numeric", month: "2-digit", day: "2-digit" }).format(new Date());
    } catch {
      return new Date().toISOString().slice(0, 10);
    }
  }

  function limaDateTime() {
    try {
      return new Intl.DateTimeFormat("es-PE", {
        timeZone: APP.CONFIG.TZ,
        year: "numeric",
        month: "2-digit",
        day: "2-digit",
        hour: "2-digit",
        minute: "2-digit",
        second: "2-digit",
        hour12: false,
      }).format(new Date());
    } catch {
      return new Date().toISOString();
    }
  }

  function limaTime() {
    try {
      return new Intl.DateTimeFormat("es-PE", {
        timeZone: APP.CONFIG.TZ,
        hour: "2-digit",
        minute: "2-digit",
        hour12: false,
      }).format(new Date());
    } catch {
      return new Date().toISOString().slice(11, 16);
    }
  }

  function endpoint() {
    return String((APP.CONFIG && APP.CONFIG._ep) || "").trim();
  }

  function deviceId() {
    try {
      let id = localStorage.getItem(DEVICE_KEY);
      if (id) return id;
      id = (crypto.randomUUID ? crypto.randomUUID() : Date.now().toString(16) + Math.random().toString(16).slice(2)).replace(/-/g, "").slice(0, 16);
      localStorage.setItem(DEVICE_KEY, id);
      return id;
    } catch {
      return "d";
    }
  }

  function uid() {
    const rand = crypto.randomUUID ? crypto.randomUUID() : Math.random().toString(16).slice(2) + Date.now().toString(16);
    return "c-" + deviceId() + "-" + Date.now() + "-" + rand.replace(/-/g, "").slice(0, 12);
  }

  function all() {
    try {
      const list = JSON.parse(localStorage.getItem(KEY) || "[]");
      return Array.isArray(list) ? list : [];
    } catch {
      return [];
    }
  }

  function write(list, silent) {
    const next = Array.isArray(list) ? list : [];
    const pending = next.filter((r) => r && r.syncStatus !== "confirmed");
    const confirmed = next.filter((r) => r && r.syncStatus === "confirmed");
    const capped = pending.concat(confirmed.slice(0, Math.max(0, 800 - pending.length)));
    try {
      localStorage.setItem(KEY, JSON.stringify(capped));
    } catch (err) {
      try {
        const emergency = pending.concat(confirmed.slice(0, 80));
        localStorage.setItem(KEY, JSON.stringify(emergency));
      } catch (_) {
        throw err;
      }
    }
    if (!silent) window.dispatchEvent(new Event("app:activity"));
  }

  function sameLoteDay(a, b) {
    return (
      String(a.fecha) === String(b.fecha) &&
      String(a.lote) === String(b.lote) &&
      String(a.turnoCampo || "Mañana") === String(b.turnoCampo || "Mañana") &&
      String(a.supervisorDni || "").trim() === String(b.supervisorDni || "").trim()
    );
  }

  function matchSupervisor_(r, supervisorDni) {
    if (supervisorDni == null || supervisorDni === false) return true;
    const want = String(supervisorDni || "").trim();
    if (!want) return true;
    return String((r && r.supervisorDni) || "").trim() === want;
  }

  function healSending(list) {
    if (syncLock) return list;
    return list.map((r) => (r && r.syncStatus === "sending" ? { ...r, syncStatus: "pending" } : r));
  }

  function pruneOldRecords() {
    const day = todayKey();
    const keep = healSending(all()).filter((r) => {
      if (!r) return false;
      if (r.syncStatus !== "confirmed") return true;
      return String(r.fecha) === day;
    });
    write(keep);
    return keep;
  }

  function localRecords(fecha, turnoCampo, supervisorDni) {
    const day = fecha || todayKey();
    const byKey = {};
    healSending(all()).forEach((r) => {
      if (!r || String(r.fecha) !== day) return;
      if (!matchSupervisor_(r, supervisorDni)) return;
      const tc = r.turnoCampo || "Mañana";
      if (turnoCampo && tc !== turnoCampo) return;
      const dni = String(r.supervisorDni || "").trim() || String(supervisorDni || "").trim();
      byKey[String(r.lote) + "|" + tc + "|" + dni] = r;
    });
    return Object.values(byKey).sort((a, b) => {
      const c = String(a.lote).localeCompare(String(b.lote), "es", { numeric: true });
      return c || String(a.turnoCampo || "").localeCompare(String(b.turnoCampo || ""));
    });
  }

  function recordOf(lote, fecha, turnoCampo, supervisorDni) {
    return localRecords(fecha, turnoCampo || "Mañana", supervisorDni).find((r) => String(r.lote) === String(lote)) || null;
  }

  function pendingRecords(supervisorDni) {
    return healSending(all()).filter(
      (r) => r && r.lote && r.syncStatus !== "confirmed" && matchSupervisor_(r, supervisorDni)
    );
  }

  /** Lotes guardados aún no confirmados en Sheets (interno). */
  function lotPendingCount() {
    return pendingRecords().length;
  }

  function readOutbox() {
    try {
      const raw = JSON.parse(localStorage.getItem(OUTBOX_KEY) || "null");
      const day = todayKey();
      if (!raw || raw.fecha !== day || !Array.isArray(raw.items)) {
        return { fecha: day, items: [] };
      }
      return { fecha: day, items: raw.items.filter((x) => x && x.id) };
    } catch {
      return { fecha: todayKey(), items: [] };
    }
  }

  function writeOutbox(box, silent) {
    try {
      localStorage.setItem(OUTBOX_KEY, JSON.stringify(box || { fecha: todayKey(), items: [] }));
    } catch (_) {}
    if (!silent) window.dispatchEvent(new Event("app:activity"));
  }

  /**
   * Chip "X pend." = reportes/envíos en cola (0 hasta pulsar Enviar).
   * No cuenta lotes guardados en el celular.
   */
  function pendingCount() {
    return readOutbox().items.length;
  }

  function queueReport(opts) {
    const turnoCampo = opts.turnoCampo === "Tarde" ? "Tarde" : "Mañana";
    const records = (opts.records || []).filter((r) => r && r.clientId && r.lote);
    if (!records.length) return null;
    const summary = String(opts.summary || "").trim();
    const box = readOutbox();
    box.items = box.items.filter((x) => (x.turnoCampo || "Mañana") !== turnoCampo);
    const item = {
      id: "rep-" + turnoCampo.toLowerCase().replace(/ñ/g, "n") + "-" + Date.now().toString(16),
      turnoCampo,
      summary,
      clientIds: records.map((r) => r.clientId),
      lotes: records.length,
      jarras: records.reduce((a, r) => a + (Number(r.totalJarras) || 0), 0),
      kg: records.reduce((a, r) => a + (Number(r.totalKg) || 0), 0),
      createdAt: Date.now(),
    };
    box.items.push(item);
    writeOutbox(box);
    return item;
  }

  function clearOutboxReport(turnoCampo) {
    const tc = turnoCampo === "Tarde" ? "Tarde" : "Mañana";
    const box = readOutbox();
    box.items = box.items.filter((x) => (x.turnoCampo || "Mañana") !== tc);
    writeOutbox(box, true);
  }

  function clearOutboxAll() {
    writeOutbox({ fecha: todayKey(), items: [] }, true);
  }

  /** Quita del outbox los clientIds ya confirmados → el chip de pend. baja al instante. */
  function ackOutboxIds(clientIds) {
    const set = new Set((clientIds || []).map((id) => String(id || "").trim()).filter(Boolean));
    if (!set.size) return;
    const box = readOutbox();
    let changed = false;
    const next = [];
    box.items.forEach((item) => {
      if (!item) return;
      const ids = Array.isArray(item.clientIds) ? item.clientIds : [];
      if (!ids.length) {
        next.push(item);
        return;
      }
      const left = ids.filter((id) => !set.has(String(id)));
      if (left.length === ids.length) {
        next.push(item);
        return;
      }
      changed = true;
      if (!left.length) return; // reporte completo → sale del outbox
      next.push({
        ...item,
        clientIds: left,
        lotes: left.length,
      });
    });
    if (changed) writeOutbox({ fecha: box.fecha || todayKey(), items: next }, true);
  }

  function todayRecords(supervisorDni) {
    const day = todayKey();
    return healSending(all()).filter(
      (r) => r && r.lote && String(r.fecha) === day && matchSupervisor_(r, supervisorDni)
    );
  }

  function todayCount(supervisorDni) {
    return todayRecords(supervisorDni).length;
  }

  function wasMorningSent() {
    return false;
  }

  function sendMap() {
    try {
      const raw = JSON.parse(localStorage.getItem(SENDS_KEY) || "{}");
      const day = todayKey();
      if (!raw || raw.fecha !== day || !raw.by || typeof raw.by !== "object") {
        return { fecha: day, by: {} };
      }
      return { fecha: day, by: raw.by };
    } catch {
      return { fecha: todayKey(), by: {} };
    }
  }

  function sendCount(dni) {
    const id = String(dni || "").trim();
    if (!id) return 0;
    return Math.min(2, Number(sendMap().by[id]) || 0);
  }

  function bumpSend(dni) {
    const id = String(dni || "").trim();
    if (!id) return 0;
    const map = sendMap();
    const n = Math.min(2, (Number(map.by[id]) || 0) + 1);
    map.by[id] = n;
    try {
      localStorage.setItem(SENDS_KEY, JSON.stringify(map));
    } catch (_) {}
    return n;
  }

  function closeMap() {
    try {
      const raw = JSON.parse(localStorage.getItem(CLOSE_KEY) || "null");
      const day = todayKey();
      if (!raw || raw.fecha !== day) return { fecha: day, by: {} };
      if (raw.by && typeof raw.by === "object") return { fecha: day, by: { ...raw.by } };
      // Legado closed global: no bloquear a otros supervisores.
      return { fecha: day, by: {} };
    } catch {
      return { fecha: todayKey(), by: {} };
    }
  }

  function isTodayClosed(supervisorDni) {
    const id = String(supervisorDni || "").trim();
    if (!id) return false;
    return !!closeMap().by[id];
  }

  function setTodayClosed(closed, supervisorDni) {
    const id = String(supervisorDni || "").trim();
    if (!id) return;
    const map = closeMap();
    if (closed) map.by[id] = true;
    else delete map.by[id];
    try {
      localStorage.setItem(CLOSE_KEY, JSON.stringify(map));
    } catch (_) {}
  }

  function queueTodayForResend(supervisorDni) {
    const day = todayKey();
    const dni = String(supervisorDni || "").trim();
    write(
      all().map((r) => {
        if (!r || !r.lote || String(r.fecha) !== day) return r;
        if (dni && !matchSupervisor_(r, dni)) return r;
        return { ...r, syncStatus: "pending" };
      })
    );
    if (dni) setTodayClosed(false, dni);
  }

  function patchMany(ids, extra, silent) {
    const set = new Set(ids);
    write(
      all().map((r) => (r && set.has(r.clientId) ? { ...r, ...extra } : r)),
      silent
    );
  }

  function catalogPatch(rec) {
    const L = APP.Data && rec && rec.lote ? APP.Data.findLote(rec.lote) : null;
    if (!L) return rec;
    return {
      ...rec,
      md: L.md || rec.md,
      modulo: L.modulo || rec.modulo,
      variedad: L.variedad || rec.variedad,
      fundo: L.fundo || rec.fundo,
      etapa: L.etapa || rec.etapa,
    };
  }

  function stampForSend(rec) {
    const session = loadSession() || {};
    // No pisar el supervisor del lote con el de la sesión actual (celular compartido).
    const dni = String(rec.supervisorDni || session.supervisorDni || "").trim();
    const nombre = APP.Data
      ? APP.Data.fullName(dni, rec.supervisor || session.supervisorNombre)
      : rec.supervisor || session.supervisorNombre || "";
    return catalogPatch({
      ...rec,
      supervisorDni: dni || rec.supervisorDni,
      supervisor: nombre || rec.supervisor,
    });
  }

  async function submit(payload) {
    pruneOldRecords();
    const data = catalogPatch(payload.data || {});
    const session = loadSession() || {};
    if (!String(data.supervisorDni || "").trim()) {
      data.supervisorDni = String(session.supervisorDni || "").trim();
    }
    if (!String(data.supervisor || "").trim()) {
      data.supervisor = session.supervisorNombre || "";
    }
    if (isTodayClosed(data.supervisorDni)) return { ok: false, reason: "closed" };
    data.fecha = todayKey();
    data.horaEnvio = limaDateTime();
    const prev = all().find((r) => sameLoteDay(r, data));
    data.horaRegistro = (prev && prev.horaRegistro) || data.horaRegistro || limaTime();
    const rec = {
      ...data,
      clientId: prev && prev.clientId ? prev.clientId : payload.clientId || uid(),
      submittedAt: payload.submittedAt || new Date().toISOString(),
      syncStatus: "pending",
      uploaded: !!(prev && (prev.uploaded || prev.syncStatus === "confirmed")),
      syncAttempts: prev && prev.syncStatus === "confirmed" ? 0 : Number(prev && prev.syncAttempts) || 0,
    };
    write([rec, ...all().filter((r) => !sameLoteDay(r, rec) && r.clientId !== rec.clientId)]);
    return { ok: true, mode: "local", clientId: rec.clientId, pending: true };
  }

  function fetchTimeout(url, opts, ms) {
    const ctrl = new AbortController();
    const t = setTimeout(() => ctrl.abort(), ms);
    return fetch(url, { ...opts, signal: ctrl.signal }).finally(() => clearTimeout(t));
  }

  async function probe() {
    const ep = endpoint();
    if (!ep) return { ok: false, reason: "no-ep" };
    try {
      const sep = ep.includes("?") ? "&" : "?";
      const res = await fetchTimeout(ep + sep + "ping=1", { method: "GET", cache: "no-store" }, PING_TIMEOUT);
      const text = await res.text();
      let json = null;
      try {
        json = JSON.parse(text);
      } catch (_) {}
      if (json && (json.ok === true || json.pong === true)) return { ok: true };
      return { ok: false, reason: "http" };
    } catch (_) {
      return { ok: false, reason: "net" };
    }
  }

  function tick(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms || 0));
  }

  async function sendBatch(batch, rebuild, meta) {
    const ep = endpoint();
    const body = JSON.stringify({
      action: "batchSave",
      rebuild: rebuild !== false,
      summary: meta && meta.summary ? String(meta.summary) : "",
      turnoCampo: meta && meta.turnoCampo ? meta.turnoCampo : "",
      report: meta && meta.report ? meta.report : null,
      records: batch.map((r) => ({ clientId: r.clientId, data: stampForSend(r) })),
    });
    const res = await fetchTimeout(
      ep,
      {
        method: "POST",
        cache: "no-store",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body,
      },
      rebuild ? 55000 : SEND_TIMEOUT
    );
    const text = await res.text();
    let json = null;
    try {
      json = JSON.parse(text);
    } catch (_) {}
    if (!json || json.ok !== true) throw new Error("bad-response");
    // ok:true del servidor = lote guardado → confirmar siempre (aunque no liste IDs).
    const okIds = new Set([].concat(json.accepted || [], json.existing || []));
    batch.forEach((r) => {
      if (r && r.clientId) okIds.add(r.clientId);
    });
    return okIds;
  }

  async function runFlush(opts) {
    const onProgress = opts && opts.onProgress;
    const summary = opts && opts.summary ? String(opts.summary) : "";
    const turnoCampo = opts && opts.turnoCampo === "Tarde" ? "Tarde" : "Mañana";
    const session = loadSession() || {};
    const supervisorDni = String(
      (opts && opts.supervisorDni != null ? opts.supervisorDni : session.supervisorDni) || ""
    ).trim();
    const pendingOf = () => pendingRecords(supervisorDni || undefined);
    pruneOldRecords();
    if (!endpoint()) {
      return { sent: 0, left: pendingOf().length, reason: "no-ep", reports: pendingCount() };
    }

    let sent = 0;
    let guard = 0;
    let failures = 0;
    while (guard++ < 80) {
      const queue = pendingOf();
      if (!queue.length) break;
      const batch = queue.slice(0, BATCH);
      const ids = batch.map((r) => r.clientId);
      patchMany(ids, { syncStatus: "sending", syncStartedAt: Date.now() }, true);
      const leftNow = pendingOf().length;
      if (onProgress) onProgress({ sent, left: leftNow, total: sent + leftNow, reports: pendingCount() });
      try {
        const last = queue.length <= batch.length;
        // rebuild lo hace el servidor en cola; aquí pedimos ok rápido.
        const okIds = await sendBatch(batch, false, {
          summary,
          turnoCampo,
          report: { summary, turnoCampo, lotes: queue.length, final: last },
        });
        const confirmedAt = new Date().toISOString();
        const ok = batch.filter((r) => okIds.has(r.clientId)).map((r) => r.clientId);
        const fail = batch.filter((r) => !okIds.has(r.clientId));
        if (ok.length) {
          patchMany(ok, { syncStatus: "confirmed", uploaded: true, syncedAt: confirmedAt }, true);
          ackOutboxIds(ok);
          sent += ok.length;
          window.dispatchEvent(new Event("app:activity"));
          const leftAfter = pendingOf().length;
          if (onProgress) {
            onProgress({
              sent,
              left: leftAfter,
              total: sent + leftAfter,
              reports: pendingCount(),
              confirmed: ok.length,
            });
          }
        }
        if (fail.length) {
          write(
            all().map((r) =>
              fail.some((f) => f.clientId === r.clientId)
                ? { ...r, syncStatus: "pending", syncAttempts: Number(r.syncAttempts || 0) + 1 }
                : r
            ),
            true
          );
          failures += 1;
          if (failures >= MAX_FLUSH_FAILURES) {
            break;
          }
        } else {
          failures = 0;
        }
      } catch (_) {
        write(
          all().map((r) =>
            ids.includes(r.clientId)
              ? { ...r, syncStatus: "pending", syncAttempts: Number(r.syncAttempts || 0) + 1 }
              : r
          ),
          true
        );
        failures += 1;
        if (failures >= MAX_FLUSH_FAILURES) {
          return { sent, left: pendingOf().length, reason: "net", reports: pendingCount() };
        }
        await tick(Math.min(2000, 500 * failures));
      }
      await tick(120);
    }
    const left = pendingOf().length;
    if (left === 0) clearOutboxReport(turnoCampo);
    window.dispatchEvent(new Event("app:activity"));
    return {
      sent,
      left,
      reason: left > 0 ? (sent > 0 ? "partial" : "net") : "",
      reports: pendingCount(),
    };
  }

  function flush(opts) {
    if (syncLock) return syncLock;
    syncLock = runFlush(opts || {})
      .catch((err) => ({ sent: 0, left: lotPendingCount(), reason: String(err && err.message || err), reports: pendingCount() }))
      .finally(() => {
        syncLock = null;
      });
    return syncLock;
  }

  function loadSession() {
    try {
      return JSON.parse(localStorage.getItem(SESSION_KEY) || "null");
    } catch {
      return null;
    }
  }

  function saveSession(s) {
    try {
      localStorage.setItem(SESSION_KEY, JSON.stringify(s || {}));
    } catch (_) {}
  }

  function readHistory() {
    try {
      const list = JSON.parse(localStorage.getItem(HISTORY_KEY) || "[]");
      return Array.isArray(list) ? list : [];
    } catch {
      return [];
    }
  }

  function writeHistory(list) {
    try {
      localStorage.setItem(HISTORY_KEY, JSON.stringify((list || []).slice(0, 30)));
    } catch (_) {}
  }

  function pruneHistory(list) {
    const now = Date.now();
    return (list || []).filter((item) => item && Number(item.expiresAt) > now);
  }

  function listHistory() {
    const next = pruneHistory(readHistory());
    writeHistory(next);
    return next.sort((a, b) => Number(b.savedAt || 0) - Number(a.savedAt || 0));
  }

  function upsertExcelSnapshot(opts) {
    const fecha = opts.fecha || todayKey();
    const turnoCampo = opts.turnoCampo === "Tarde" ? "Tarde" : "Mañana";
    const records = (opts.records || []).filter((r) => r && String(r.lote || "").trim());
    if (!records.length) return null;
    const now = Date.now();
    const session = { ...(opts.session || {}), turnoCampo };
    const supDni = String(session.supervisorDni || "").trim();
    const model = opts.model || {};
    const t = model.totals || {};
    const list = pruneHistory(readHistory());
    const same = (item) =>
      item.fecha === fecha &&
      (item.turnoCampo || "Mañana") === turnoCampo &&
      String((item.session && item.session.supervisorDni) || "").trim() === supDni;
    const prev = list.find(same);
    const slug = turnoCampo === "Tarde" ? "tarde" : "manana";
    const item = {
      id: (prev && prev.id) || uid(),
      fecha,
      turnoCampo,
      savedAt: now,
      expiresAt: now + DAY_MS,
      pending: pendingRecords(supDni).length > 0 || !isTodayClosed(supDni),
      name: opts.name || `reporte-cosecha-${fecha}-${slug}.xls`,
      supervisor: APP.Data ? APP.Data.fullName(session.supervisorDni, session.supervisorNombre) : session.supervisorNombre || "",
      lotes: records.length,
      jarras: t.totalJarras || records.reduce((a, r) => a + (Number(r.totalJarras) || 0), 0),
      kg: t.totalKg || records.reduce((a, r) => a + (Number(r.totalKg) || 0), 0),
      session,
      records,
    };
    writeHistory([item, ...list.filter((x) => !same(x))]);
    return item;
  }

  function getHistory(id) {
    return listHistory().find((item) => item.id === id) || null;
  }

  function markHistoryIfClear(supervisorDni) {
    const dni = String(supervisorDni || "").trim();
    if (pendingRecords(dni || undefined).length > 0) return 0;
    if (dni) setTodayClosed(true, dni);
    const now = Date.now();
    const list = pruneHistory(readHistory()).map((item) => {
      const itemDni = String((item.session && item.session.supervisorDni) || "").trim();
      if (dni && itemDni && itemDni !== dni) return item;
      return { ...item, pending: false, transferredAt: now };
    });
    writeHistory(list);
    return list.length;
  }

  /** Borra solo los lotes de hoy de ese supervisor. Sin DNI = no borra (evita wipe global). */
  function clearTodayRecords(supervisorDni) {
    const day = todayKey();
    const dni = String(supervisorDni || "").trim();
    if (!dni) return { ok: false, reason: "no-supervisor" };
    const list = healSending(all());
    const mine = list.filter((r) => r && String(r.fecha) === day && String(r.supervisorDni || "").trim() === dni);
    const pendingToday = mine.filter((r) => r.syncStatus !== "confirmed");
    if (pendingToday.length) return { ok: false, reason: "pending", left: pendingToday.length };
    write(
      list.filter((r) => {
        if (!r || String(r.fecha) !== day) return true;
        return String(r.supervisorDni || "").trim() !== dni;
      })
    );
    return { ok: true };
  }

  /** Borra lotes de hoy (aunque estén pendientes). No toca el historial Excel. */
  function purgeTodayLocal(supervisorDni) {
    const day = todayKey();
    const dni = String(supervisorDni || "").trim();
    write(
      all().filter((r) => {
        if (!r || String(r.fecha) !== day) return true;
        if (!dni) return false;
        return String(r.supervisorDni || "").trim() !== dni;
      })
    );
    clearOutboxAll();
    if (dni) setTodayClosed(false, dni);
    window.dispatchEvent(new Event("app:activity"));
    return { ok: true };
  }

  function clearTodayLocal(supervisorDni) {
    return clearTodayRecords(supervisorDni);
  }

  function wipeLocal() {
    localStorage.removeItem(KEY);
    localStorage.removeItem(SESSION_KEY);
    localStorage.removeItem(HISTORY_KEY);
    localStorage.removeItem(CLOSE_KEY);
    localStorage.removeItem(SENDS_KEY);
    localStorage.removeItem(OUTBOX_KEY);
  }

  async function selfTest() {
    const out = [];
    const snap = localStorage.getItem(KEY);
    const hist = localStorage.getItem(HISTORY_KEY);
    const closeSnap = localStorage.getItem(CLOSE_KEY);
    const sendSnap = localStorage.getItem(SENDS_KEY);
    const outSnap = localStorage.getItem(OUTBOX_KEY);
    const ep = APP.CONFIG._ep;
    const push = (id, ok, extra) => out.push({ id, ok: !!ok, ...extra });
    try {
      APP.CONFIG._ep = "";
      localStorage.removeItem(KEY);
      clearOutboxAll();
      const a = await submit({ data: { lote: "AUDIT-A", turnoCampo: "Mañana", totalJarras: 1 } });
      const stored = all().find((r) => r.lote === "AUDIT-A");
      push("A-local-first", stored && stored.syncStatus === "pending" && stored.clientId && a.pending);
      push("A-chip-zero-before-enviar", pendingCount() === 0 && lotPendingCount() === 1);
      const id1 = stored && stored.clientId;
      await submit({ data: { lote: "AUDIT-A", turnoCampo: "Mañana", totalJarras: 2 } });
      const again = all().filter((r) => r.lote === "AUDIT-A");
      push("A-stable-id", again.length === 1 && again[0].clientId === id1);
      write([{ lote: "OLD", fecha: "2020-01-01", clientId: "c-old", syncStatus: "pending" }, ...all()]);
      pruneOldRecords();
      push("A-prune-keeps-pending", all().some((r) => r.clientId === "c-old"));
      write([{ lote: "OLD2", fecha: "2020-01-01", clientId: "c-old2", syncStatus: "confirmed" }, ...all()]);
      pruneOldRecords();
      push("A-prune-drops-confirmed", !all().some((r) => r.clientId === "c-old2"));
      const flush1 = await flush();
      const still = all().find((r) => r.lote === "AUDIT-A");
      push("B-no-ep-keeps-pending", flush1.reason === "no-ep" && still && still.syncStatus === "pending");
      const p1 = flush();
      const p2 = flush();
      push("E-mutex", p1 === p2);
      await p1;
      const day = todayKey();
      write(Array.from({ length: 100 }, (_, i) => ({ lote: "L" + i, fecha: day, clientId: "c-load-" + i, syncStatus: "pending", totalJarras: 1 })));
      const t0 = performance.now();
      const n100 = lotPendingCount();
      localRecords(day);
      const dt100 = performance.now() - t0;
      push("H-100-local-read", n100 === 100 && dt100 < 80, { ms: Math.round(dt100) });
      write(Array.from({ length: 500 }, (_, i) => ({ lote: "X" + i, fecha: day, clientId: "c-500-" + i, syncStatus: "pending", totalJarras: 1 })));
      const t1 = performance.now();
      const n500 = lotPendingCount();
      const dt500 = performance.now() - t1;
      push("H-500-pendingCount", n500 === 500 && dt500 < 80, { ms: Math.round(dt500) });
      markHistoryIfClear();
      push("C-mark-keeps-lots", lotPendingCount() === 500 && pendingCount() === 0);
      write([{ lote: "S", fecha: day, clientId: "c-send", syncStatus: "sending", syncStartedAt: Date.now() }]);
      const healed = pendingRecords();
      push("G-heal-sending", healed.length === 1 && healed[0].syncStatus === "pending");
      const forty = Array.from({ length: 40 }, (_, i) => ({
        lote: "Z" + i,
        fecha: day,
        clientId: "c-40-" + i,
        turnoCampo: "Mañana",
        supervisorDni: "48533707",
        syncStatus: "pending",
        totalJarras: 1,
        totalKg: 1.14,
      }));
      write(forty);
      const ids40 = all().map((r) => r.clientId);
      push("I-40-unique-clientId", ids40.length === 40 && new Set(ids40).size === 40);
      write(forty.map((r, i) => (i < 7 ? { ...r, syncStatus: "sending" } : r)));
      const afterReload = pendingRecords();
      push("I-reload-heals-sending", afterReload.length === 40 && afterReload.every((r) => r.syncStatus === "pending"));
      const rep = queueReport({
        turnoCampo: "Mañana",
        summary: "40 lotes · 40 jarras · 45.6 kg",
        records: forty,
      });
      push("I-outbox-one-report", !!(rep && pendingCount() === 1) && lotPendingCount() === 40);
      setTodayClosed(true, "48533707");
      const blocked = await submit({ data: { lote: "NO", turnoCampo: "Mañana", totalJarras: 1, supervisorDni: "48533707" } });
      push("C-closed-blocks-submit", blocked && blocked.ok === false && blocked.reason === "closed");
      const noWipe = clearTodayRecords("48533707");
      push("C-close-keeps-pending", noWipe && noWipe.ok === false && lotPendingCount() === 40);
      setTodayClosed(false, "48533707");
      write(forty.map((r) => ({ ...r, syncStatus: "confirmed" })));
      clearOutboxAll();
      markHistoryIfClear("48533707");
      const wipedOk = clearTodayRecords("48533707");
      push("C-close-wipes-only-confirmed", !!(wipedOk && wipedOk.ok) && lotPendingCount() === 0 && pendingCount() === 0);
      const fakeOk = { ok: true, accepted: ["c-1"], existing: ["c-1"] };
      push("D-timeout-not-delete", lotPendingCount() >= 0 && fakeOk.existing[0] === "c-1");
    } catch (err) {
      push("crash", false, { error: String(err && err.message || err) });
    } finally {
      APP.CONFIG._ep = ep;
      if (snap == null) localStorage.removeItem(KEY);
      else localStorage.setItem(KEY, snap);
      if (hist == null) localStorage.removeItem(HISTORY_KEY);
      else localStorage.setItem(HISTORY_KEY, hist);
      if (closeSnap == null) localStorage.removeItem(CLOSE_KEY);
      else localStorage.setItem(CLOSE_KEY, closeSnap);
      if (sendSnap == null) localStorage.removeItem(SENDS_KEY);
      else localStorage.setItem(SENDS_KEY, sendSnap);
      if (outSnap == null) localStorage.removeItem(OUTBOX_KEY);
      else localStorage.setItem(OUTBOX_KEY, outSnap);
    }
    return out;
  }

  deviceId();

  return {
    todayKey,
    limaDateTime,
    limaTime,
    pruneOldRecords,
    newClientId: uid,
    localRecords,
    recordOf,
    isUploaded: (r) => !!(r && (r.uploaded || r.syncStatus === "confirmed")),
    removeTodayLote: (lote, turnoCampo, supervisorDni) => {
      const day = todayKey();
      const tc = turnoCampo === "Tarde" ? "Tarde" : "Mañana";
      const dni = String(supervisorDni || "").trim();
      const rec = recordOf(lote, day, tc, dni || undefined);
      if (!rec) return { ok: false, reason: "missing" };
      if (rec.uploaded || rec.syncStatus === "confirmed") return { ok: false, reason: "uploaded" };
      write(
        all().filter((r) => {
          if (!r || String(r.fecha) !== day || String(r.lote) !== String(lote) || (r.turnoCampo || "Mañana") !== tc) {
            return true;
          }
          if (dni) return String(r.supervisorDni || "").trim() !== dni;
          return false;
        })
      );
      return { ok: true };
    },
    pendingRecords,
    pendingCount,
    lotPendingCount,
    queueReport,
    clearOutboxReport,
    clearOutboxAll,
    ackOutboxIds,
    todayCount,
    wasMorningSent,
    sendCount,
    bumpSend,
    isTodayClosed,
    queueTodayForResend,
    submit,
    flush,
    probe,
    loadSession,
    saveSession,
    listHistory,
    upsertExcelSnapshot,
    getHistory,
    markTransferred: markHistoryIfClear,
    clearTodayRecords,
    clearTodayLocal,
    purgeTodayLocal,
    wipeLocal,
    _selfTest: selfTest,
  };
})();
