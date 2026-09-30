/** Flujo simple: supervisor → escáner → lote → avance → convencional/China */
window.APP = window.APP || {};

APP.App = (() => {
  const $ = (s) => document.querySelector(s);

  const state = {
    session: {
      scannerDni: "",
      scannerNombre: "",
      supervisorDni: "",
      supervisorNombre: "",
      jornales: 0,
      grupo: "",
      etapa: "",
      fundo: "",
      turnoCampo: "Mañana",
      pesoJarra: "1.14",
    },
    lote: null,
    avance: "",
    conv: false,
    china: false,
    jarrasConv: "",
    jarrasChina: "",
    openLots: [],
    saving: false,
    transferring: false,
  };

  function today() {
    return APP.API.todayKey();
  }

  function supDni() {
    return String(state.session.supervisorDni || "").trim();
  }

  function fundoList() {
    const list = APP.CONFIG && APP.CONFIG.FUNDOS;
    return Array.isArray(list) && list.length ? list : ["LICAPA I", "LICAPA II", "LICAPA III", "LICAPA IV"];
  }

  function canonFundo(v) {
    const s = String(v || "").trim().toUpperCase().replace(/\s+/g, " ");
    return fundoList().find((f) => String(f).toUpperCase() === s) || "";
  }

  function dayRecords(fecha, turnoCampo) {
    const day = fecha || today();
    const dni = supDni();
    if (!dni) return [];
    if (turnoCampo) return APP.API.localRecords(day, turnoCampo, dni);
    return APP.API.localRecords(day, null, dni);
  }

  function fmtDate(iso) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || "");
    return m ? `${m[3]}/${m[2]}/${m[1]}` : iso;
  }

  function feedback(title, text, opts) {
    return new Promise((resolve) => {
      const root = $("#qb-fb");
      const head = $("#qb-fb-title");
      const line = $("#qb-fb-text");
      const actions = $("#qb-fb-actions");
      if (!root || !head || !line || !actions) {
        resolve(true);
        return;
      }
      head.textContent = title || "";
      if (opts && opts.html) line.innerHTML = text || "";
      else line.textContent = text || "";
      line.classList.toggle("is-detail", !!(opts && opts.html));
      actions.innerHTML = "";
      const ok = document.createElement("button");
      ok.textContent = "Listo";
      const close = () => {
        root.hidden = true;
        resolve(true);
      };
      ok.onclick = close;
      actions.appendChild(ok);
      root.hidden = false;
      setTimeout(close, (opts && opts.ms) || 1600);
    });
  }

  function ask(title, text, okLabel) {
    return new Promise((resolve) => {
      const root = $("#qb-fb");
      const head = $("#qb-fb-title");
      const body = $("#qb-fb-text");
      const actions = $("#qb-fb-actions");
      if (!root || !head || !body || !actions) {
        resolve(false);
        return;
      }
      head.textContent = title || "";
      body.textContent = text || "";
      actions.innerHTML = "";
      const cancel = document.createElement("button");
      cancel.textContent = "Cancelar";
      cancel.style.background = "#E8EEE9";
      cancel.style.color = "#1B5E20";
      cancel.onclick = () => {
        root.hidden = true;
        resolve(false);
      };
      const ok = document.createElement("button");
      ok.textContent = okLabel || "Sí";
      ok.onclick = () => {
        root.hidden = true;
        resolve(true);
      };
      actions.appendChild(cancel);
      actions.appendChild(ok);
      root.hidden = false;
    });
  }

  function rememberExcel(pack, turnoCampo) {
    const turno = turnoCampo === "Tarde" || turnoCampo === "Mañana"
      ? turnoCampo
      : state.session.turnoCampo === "Tarde" ? "Tarde" : "Mañana";
    const fecha = (pack && pack.model && pack.model.fecha) || today();
    const records = turno === "Tarde"
      ? dayRecords(fecha)
      : dayRecords(fecha, "Mañana");
    const session = { ...state.session, turnoCampo: turno };
    const use = pack && turno === (pack.model && pack.model.turnoCampo)
      ? pack
      : APP.Excel.buildFile({ session, fecha, records });
    return APP.API.upsertExcelSnapshot({
      fecha,
      turnoCampo: turno,
      session,
      records,
      model: use.model,
      name: use.name,
    });
  }

  function showLoader(title, text) {
    const root = $("#qb-loader");
    const head = $("#qb-loader-title");
    const line = $("#qb-loader-text");
    if (head) head.textContent = title || "Cargando";
    if (line) line.textContent = text || "";
    if (root) root.hidden = false;
  }

  function hideLoader() {
    const root = $("#qb-loader");
    if (root) root.hidden = true;
  }

  function wait(ms) {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  async function fetchCloud(path) {
    const url = path + (path.includes("?") ? "&" : "?") + "cloud=" + Date.now();
    const res = await fetch(url, {
      cache: "no-store",
      headers: { "Cache-Control": "no-cache", Pragma: "no-cache" },
    });
    if (!res.ok) throw new Error("cloud " + res.status);
    return res;
  }

  function persistSession() {
    APP.API.saveSession(state.session);
  }

  function loadSession() {
    const saved = APP.API.loadSession();
    if (saved) state.session = { ...state.session, ...saved };
  }

  function paintPeople(opts) {
    const sup = $("#lbl-supervisor");
    const scan = $("#lbl-scanner");
    const jn = $("#inp-jornales");
    if (sup) {
      const name = APP.Data.fullName(state.session.supervisorDni, state.session.supervisorNombre);
      if (name) state.session.supervisorNombre = name;
      sup.textContent = name || "Seleccionar…";
    }
    if (scan) {
      const name = APP.Data.fullName(state.session.scannerDni, state.session.scannerNombre);
      if (name) state.session.scannerNombre = name;
      scan.textContent = name || "Seleccionar…";
    }
    if (jn) jn.value = state.session.jornales > 0 ? String(state.session.jornales) : "";
    if (!state.session.pesoJarra) state.session.pesoJarra = "1.14";
    paintPeso();
    const fundo = canonFundo(state.session.fundo);
    if (fundo) state.session.fundo = fundo;
    const lblFundo = $("#lbl-fundo");
    if (lblFundo) lblFundo.textContent = fundo || "Seleccionar…";
    document.querySelectorAll("#menu-fundo [data-fundo]").forEach((btn) => {
      const on = btn.dataset.fundo === fundo;
      btn.classList.toggle("on", on);
      btn.setAttribute("aria-selected", on ? "true" : "false");
    });
    const envios = APP.API.sendCount(state.session.supervisorDni);
    if (!opts || !opts.keepTurno) syncTurnoCampo();
    const turno = state.session.turnoCampo === "Tarde" ? "Tarde" : "Mañana";
    const lbl = $("#lbl-turno-campo");
    if (lbl) lbl.textContent = turno;
    const turnoBtn = $("#btn-turno-campo");
    if (turnoBtn) turnoBtn.disabled = envios >= 1;
    document.querySelectorAll("#menu-turno-campo [data-turno]").forEach((btn) => {
      const on = btn.dataset.turno === turno;
      btn.classList.toggle("on", on);
      btn.setAttribute("aria-selected", on ? "true" : "false");
    });
    const env = $("#lbl-envios");
    if (env) env.textContent = envios + " / 2 envíos hoy";
    const sub = $("#lote-sub");
    if (sub) sub.textContent = turno === "Tarde" ? "Sigue los lotes de la mañana o agrega uno nuevo." : "Lote, avance y jarras del día.";
  }

  function syncTurnoCampo() {
    const envios = APP.API.sendCount(state.session.supervisorDni);
    const turno = envios >= 1 ? "Tarde" : "Mañana";
    if (state.session.turnoCampo !== turno) {
      state.session.turnoCampo = turno;
      persistSession();
    }
    return turno;
  }

  function isTarde() {
    return state.session.turnoCampo === "Tarde";
  }

  function morningAvanceOf(lote) {
    const rec = APP.API.recordOf(lote, today(), "Mañana", supDni(), state.session.fundo);
    const n = Number(rec && rec.avance);
    return Number.isFinite(n) && n > 0 ? n : 0;
  }

  function avanceParaGuardar() {
    const typed = Number(state.avance) || 0;
    if (state.cierre) {
      const mine = Number(state.cierre.mio) || 0;
      const total = Math.round((mine + typed) * 1000) / 1000;
      return total > 0 ? String(total) : "";
    }
    return state.avance || "";
  }

  function hechoHa() {
    if (!state.cierre) return 0;
    return Math.round((Number(state.cierre.avance) || 0) * 1000) / 1000;
  }

  function faltaHa() {
    if (!state.cierre) return null;
    const area = Number(state.cierre.area) || 0;
    const typed = Number(state.avance) || 0;
    return Math.round((area - hechoHa() - typed) * 1000) / 1000;
  }

  function avanceTope() {
    if (state.cierre) {
      const left = Math.round(((Number(state.cierre.area) || 0) - hechoHa()) * 1000) / 1000;
      return left > 0 ? left : 0;
    }
    return loteAreaMax();
  }

  function avanceFloor() {
    if (state.cierre) return 0;
    if (!isTarde() || !state.lote) return 0;
    return morningAvanceOf(state.lote.lote);
  }

  function paintAvanceHint() {
    const hint = $("#avance-min");
    if (!hint) return;
    const lab = $("#lbl-avance");
    const inp = $("#inp-avance");
    if (state.cierre) {
      const hecho = hechoHa();
      const left = faltaHa();
      hint.hidden = false;
      const base = `Se trabajó ${haTxt(hecho)} ha.`;
      if (left > 0) {
        hint.classList.remove("is-gap");
        hint.textContent = `${base} Falta ${haTxt(left)} ha.`;
      } else if (left < 0) {
        hint.classList.add("is-gap");
        hint.textContent = `${base} Supera el área por ${haTxt(Math.abs(left))} ha.`;
      } else if (Number(state.avance) > 0) {
        hint.classList.remove("is-gap");
        hint.textContent = `${base} Con esto se completa el área.`;
      } else {
        hint.classList.remove("is-gap");
        hint.textContent = `${base} Falta 0 ha.`;
      }
      if (lab) lab.textContent = "Lo que avanzas ahora (ha)";
      if (inp) inp.placeholder = "Solo lo de ahora";
      return;
    }
    hint.classList.remove("is-gap");
    if (lab) lab.textContent = "Avance (ha)";
    if (inp && document.activeElement !== inp) inp.placeholder = "0";
    const floor = avanceFloor();
    if (floor > 0) {
      hint.hidden = false;
      hint.textContent = `Avance de la mañana: ${floor} ha. En la tarde solo se puede agregar más.`;
      return;
    }
    const max = loteAreaMax();
    const av = Number(state.avance);
    if (!isTarde() && max != null && Number.isFinite(av) && av > 0 && av < max) {
      const left = Math.round((max - av) * 1000) / 1000;
      hint.hidden = false;
      hint.textContent = `Te quedan ${left} ha. Puedes seguir editando y sumar.`;
      return;
    }
    hint.hidden = true;
    hint.textContent = "";
  }

  function fillLoteForm(L, rec) {
    state.lote = L;
    const floor = isTarde() ? morningAvanceOf(L.lote) : 0;
    const recAv = rec ? Number(rec.avance) : 0;
    if (rec && Number.isFinite(recAv) && recAv > 0) state.avance = String(rec.avance);
    else if (floor > 0) state.avance = String(floor);
    else state.avance = "";
    state.jarrasConv = rec ? String(rec.jarrasConv || "") : "";
    state.jarrasChina = rec ? String(rec.jarrasChina || "") : "";
    state.conv = Number(rec?.jarrasConv) > 0;
    state.china = Number(rec?.jarrasChina) > 0;
    if (L.grupo) state.session.grupo = String(L.grupo);
    if (L.etapa) state.session.etapa = String(L.etapa);
    $("#inp-avance").value = state.avance;
    $("#inp-jconv").value = state.jarrasConv;
    $("#inp-jchina").value = state.jarrasChina;
    paintLote();
    paintChecks();
    paintAvanceHint();
  }

  function paintLote() {
    const L = state.lote;
    $("#lbl-lote").textContent = L ? String(L.lote) : "Seleccionar lote…";
    $("#lbl-md").textContent = L ? L.modulo || L.md || "—" : "—";
    $("#lbl-turno").textContent = L ? L.turno : "—";
    $("#lbl-area").textContent = L && L.area !== "" && L.area != null ? `${L.area} ha` : "—";
    $("#lbl-extra").textContent = L
      ? [L.fundo, L.variedad, L.codLote, L.etapa, L.grupo ? "Grupo " + L.grupo : ""]
          .filter(Boolean)
          .join(" · ")
      : "";
  }

  function loteAreaMax() {
    const n = Number(state.lote && state.lote.area);
    return Number.isFinite(n) && n > 0 ? n : null;
  }

  function sanitizeAvance(raw, finish) {
    let s = String(raw || "").replace(",", ".");
    s = s.replace(/[^\d.]/g, "");
    const dot = s.indexOf(".");
    if (dot !== -1) s = s.slice(0, dot + 1) + s.slice(dot + 1).replace(/\./g, "");
    if (s === ".") s = "0.";
    if (!s) return { text: "", num: 0 };
    if (!finish && s.endsWith(".")) return { text: s, num: Number(s.slice(0, -1)) || 0 };
    const n = Number(s);
    if (!Number.isFinite(n) || n < 0) return { text: "", num: 0 };
    const max = avanceTope();
    if (max != null && n > max) return { text: String(max), num: max, capped: true };
    const floor = avanceFloor();
    if (floor > 0 && n < floor) return { text: String(floor), num: floor, floored: true };
    return { text: finish ? String(n) : s, num: n };
  }

  function applyAvance(raw, finish) {
    const next = sanitizeAvance(raw, finish);
    state.avance = next.text;
    const el = $("#inp-avance");
    if (el && el.value !== next.text) el.value = next.text;
    paintAvanceHint();
    paintTotals();
    return next;
  }

  function currentData() {
    const L = state.lote || {};
    const jConv = state.conv ? Number(state.jarrasConv) || 0 : 0;
    const jChina = state.china ? Number(state.jarrasChina) || 0 : 0;
    return {
      fecha: state.cierre && state.cierre.fecha ? state.cierre.fecha : today(),
      cierre: !!(state.cierre && state.cierre.fecha),
      lote: L.lote || "",
      fundo: canonFundo(state.session.fundo) || L.fundo || "",
      variedad: L.variedad || "",
      md: L.md || "",
      turno: L.turno || "",
      area: L.area || "",
      avance: avanceParaGuardar(),
      jarrasConv: jConv,
      jarrasChina: jChina,
      kgConv: 0,
      kgChina: 0,
      grupo: L.grupo || state.session.grupo,
      etapa: L.etapa || state.session.etapa,
      jornales: Number(state.session.jornales) || 0,
      turnoCampo: state.session.turnoCampo === "Tarde" ? "Tarde" : "Mañana",
      horaRegistro: "",
      horaEnvio: APP.API.limaDateTime(),
      scanner: APP.Data.fullName(state.session.scannerDni, state.session.scannerNombre),
      scannerDni: state.session.scannerDni,
      supervisor: APP.Data.fullName(state.session.supervisorDni, state.session.supervisorNombre) || state.session.supervisorNombre || "",
      supervisorDni: state.session.supervisorDni,
      pesoJarra: pesoNum(),
    };
  }

  function sanitizePeso(raw) {
    let s = String(raw ?? "").replace(",", ".");
    s = s.replace(/[^\d.]/g, "");
    const dot = s.indexOf(".");
    if (dot !== -1) s = s.slice(0, dot + 1) + s.slice(dot + 1).replace(/\./g, "").slice(0, 3);
    return s;
  }

  function pesoNum() {
    const raw = sanitizePeso(state.session.pesoJarra ?? "1.14");
    const n = Number(raw.endsWith(".") ? raw.slice(0, -1) : raw);
    return Number.isFinite(n) && n > 0 ? n : 1.14;
  }

  function paintPeso() {
    const el = $("#inp-peso");
    if (!el || document.activeElement === el) return;
    const v = sanitizePeso(state.session.pesoJarra ?? "1.14");
    el.value = v || "1.14";
  }

  function jarrasDe(records) {
    return (records || []).reduce((sum, r) => sum + (Number(APP.Data.derive(r).totalJarras) || 0), 0);
  }

  function paintTotals() {
    const d = APP.Data.derive(currentData());
    const j = $("#lbl-jarras");
    const k = $("#lbl-kg");
    if (j) {
      const lote = state.lote && state.lote.lote;
      if (isTarde() && lote) {
        const mananaRec = APP.API.recordOf(lote, today(), "Mañana", supDni(), state.session.fundo);
        const manana = mananaRec ? Number(APP.Data.derive(mananaRec).totalJarras) || 0 : 0;
        const ahora = Number(d.totalJarras) || 0;
        j.classList.add("is-split");
        j.innerHTML =
          `<span class="jar-bit"><small>Mañana</small><b>${manana}</b></span>` +
          `<span class="jar-bit"><small>Ahora</small><b>${ahora}</b></span>` +
          `<span class="jar-bit"><small>Total</small><b>${manana + ahora}</b></span>`;
      } else {
        j.classList.remove("is-split");
        j.textContent = d.totalJarras + " jarras";
      }
    }
    if (k) k.textContent = d.totalKg + " kg";
    paintSaveBtn();
  }

  function hasJarrasInput() {
    const conv = state.conv ? Number(String(state.jarrasConv || "").replace(/\D/g, "")) || 0 : 0;
    const china = state.china ? Number(String(state.jarrasChina || "").replace(/\D/g, "")) || 0 : 0;
    return conv > 0 || china > 0;
  }

  function paintSaveBtn() {
    const btn = $("#btn-save");
    if (!btn) return;
    const on = hasJarrasInput();
    btn.disabled = !on;
    btn.setAttribute("aria-disabled", on ? "false" : "true");
  }

  function paintChecks() {
    $("#chk-conv").checked = state.conv;
    $("#chk-china").checked = state.china;
    $("#box-conv").hidden = !state.conv;
    $("#box-china").hidden = !state.china;
    paintTotals();
  }

  function dayModel() {
    const records = dayRecords();
    return APP.Excel.buildModel({ session: state.session, fecha: today(), records });
  }

  function paintDay() {
    const records = dayRecords();
    const model = dayModel();
    const t = model.totals;
    const tarde = isTarde();
    const mananaJ = tarde ? jarrasDe(dayRecords(today(), "Mañana")) : 0;
    const tardeJ = tarde ? jarrasDe(dayRecords(today(), "Tarde")) : 0;
    const jarrasKpi = tarde
      ? `<div class="kpi-split"><small>Jarras</small><b><i>Mañana ${mananaJ}</i><i>Ahora ${tardeJ}</i><i>Total ${mananaJ + tardeJ}</i></b></div>`
      : `<div><small>Jarras</small><strong>${t.totalJarras}</strong></div>`;
    $("#day-kpis").innerHTML = `
      <div><small>Lotes</small><strong>${model.filledCount}</strong></div>
      <div><small>Área</small><strong>${t.area} ha</strong></div>
      ${jarrasKpi}
      <div><small>Kg</small><strong>${t.totalKg}</strong></div>`;
    $("#lot-list").innerHTML = records.length
      ? records
          .map((r) => {
            const d = APP.Data.derive(r);
            const L = APP.Data.findLote(r.lote, r.fundo) || {};
            const tc = r.turnoCampo || "Mañana";
            const area = L.area !== "" && L.area != null ? L.area : (r.area !== "" && r.area != null ? r.area : "—");
            const avance = r.avance !== "" && r.avance != null ? r.avance : "—";
            const areaN = Number(area);
            const avN = Number(avance);
            const falta = Number.isFinite(areaN) && areaN > 0 && !(Number.isFinite(avN) && avN >= areaN);
            const subido = APP.API.isUploaded(r);
            return `<div class="lot-row">
              <article class="lot-item${falta ? " is-short" : " is-ok"}" data-lote="${esc(r.lote)}" data-tc="${tc}">
                <div class="lot-top">
                  <strong class="lot-kg-line">${d.totalJarras} jarras - ${d.totalKg} kg</strong>
                  ${subido ? `<span class="lot-badge">Subido</span>` : `<span class="lot-badge is-local">Guardado</span>`}
                </div>
                <div class="lot-grid">
                  <div class="lot-cell">
                    <em>Lote</em>
                    <b>${esc(r.lote)}</b>
                  </div>
                  <div class="lot-cell">
                    <em>Área</em>
                    <b>${esc(area)}</b>
                  </div>
                  <div class="lot-cell${falta ? " is-warn" : ""}">
                    <em>Avance</em>
                    <b>${esc(avance)}</b>
                  </div>
                </div>
              </article>
              <div class="lot-acts">
                <button type="button" class="lot-act edit" data-edit="${esc(r.lote)}" data-tc="${tc}" data-fundo="${esc(r.fundo || "")}" aria-label="Editar">
                  <svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2"><path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4 12.5-12.5z"/></svg>
                </button>
                <button type="button" class="lot-act del" data-del="${esc(r.lote)}" data-tc="${tc}" data-fundo="${esc(r.fundo || "")}" aria-label="Eliminar"${subido ? " disabled" : ""}>
                  <svg viewBox="0 0 24 24" width="16" height="16" fill="none" stroke="currentColor" stroke-width="2"><path d="M3 6h18"/><path d="M8 6V4h8v2"/><path d="M19 6l-1 14H6L5 6"/></svg>
                </button>
              </div>
            </div>`;
          })
          .join("")
      : `<p class="empty">${supDni() ? "Aún no hay lotes hoy." : "Elige un supervisor para ver su día."}</p>`;
    $("#lot-list").querySelectorAll("[data-edit]").forEach((b) => {
      b.onclick = (e) => {
        e.stopPropagation();
        loadRecord(b.dataset.edit, b.dataset.tc, b.dataset.fundo);
      };
    });
    $("#lot-list").querySelectorAll("[data-del]").forEach((b) => {
      b.onclick = (e) => {
        e.stopPropagation();
        removeLote(b.dataset.del, b.dataset.tc, b.dataset.fundo);
      };
    });
    paintOpenLots();
    refreshOpenLots();
  }

  function haTxt(n) {
    const x = Math.round(Number(n) * 1000) / 1000;
    return Number.isFinite(x) ? String(x) : "—";
  }

  function paintOpenLots() {
    const box = $("#open-lots");
    if (!box) return;
    const list = state.openLots || [];
    if (!list.length) {
      box.hidden = true;
      box.innerHTML = "";
      return;
    }
    box.hidden = false;
    box.innerHTML = `<p class="open-lots-title">Sin terminar · todos</p>` + list.map((r, i) => {
      const tag = (r.fecha ? String(r.fecha).slice(8, 10) + "/" + String(r.fecha).slice(5, 7) : "") + " · toca para terminar";
      const body = `
        <b>Lote ${esc(r.lote)}</b>
        <span>avance ${esc(haTxt(r.avance))}</span>
        <span>área ${esc(haTxt(r.area))}</span>
        <small>${esc([tag, r.supervisor, r.fundo].filter(Boolean).join(" · "))}</small>`;
      return r.puede
        ? `<button type="button" class="open-lot" data-open="${i}">${body}</button>`
        : `<article class="open-lot">${body}</article>`;
    }).join("");
    box.querySelectorAll("[data-open]").forEach((btn) => {
      btn.onclick = () => startCierre(list[Number(btn.dataset.open)]);
    });
  }

  function refreshOpenLots(force) {
    if (!navigator.onLine || !APP.API.openLots) return;
    APP.API.openLots(!!force).then((items) => {
      state.openLots = Array.isArray(items) ? items : [];
      paintOpenLots();
    }).catch(() => {});
  }

  function startCierre(item) {
    if (!item || !item.puede || !item.fecha) return;
    const L = APP.Data.findLote(item.lote, item.fundo) || {
      lote: item.lote,
      fundo: item.fundo,
      modulo: item.modulo || "",
      md: item.modulo || "",
      turno: item.turno || "",
      variedad: "",
      etapa: "",
    };
    L.area = item.area;
    const dni = supDni();
    const mine = (item.partes || []).reduce((sum, p) => sum + (String(p.dni || "") === dni ? Number(p.avance) || 0 : 0), 0);
    state.cierre = {
      fecha: item.fecha,
      avance: Number(item.avance) || 0,
      area: Number(item.area) || 0,
      mio: mine,
    };
    if (item.fundo) state.session.fundo = canonFundo(item.fundo) || state.session.fundo;
    persistSession();
    paintPeople({ keepTurno: true });
    fillLoteForm(L, null);
    state.avance = "";
    const inp = $("#inp-avance");
    if (inp) inp.value = "";
    paintLote();
    paintAvanceHint();
    const card = $("#lote-card");
    if (card) card.scrollIntoView({ block: "start", behavior: "smooth" });
  }

  async function removeLote(loteId, turnoCampo, fundo) {
    const rec = APP.API.recordOf(loteId, today(), turnoCampo, supDni(), fundo);
    if (APP.API.isUploaded(rec)) {
      await feedback("No se puede eliminar", "Este lote ya está subido. Si hay un error, edítalo.");
      return;
    }
    const ok = await ask("Eliminar lote", `Se quita el lote ${loteId} de hoy. No está subido. ¿Seguro?`, "Eliminar");
    if (!ok) return;
    const res = APP.API.removeTodayLote(loteId, turnoCampo, supDni(), fundo);
    if (!res.ok) {
      await feedback("No se eliminó", res.reason === "uploaded" ? "Este lote ya está subido." : "No se encontró el lote.");
      return;
    }
    if (state.lote && String(state.lote.lote) === String(loteId)) resetLoteInputs();
    rememberExcel();
    paintDay();
    paintStatus();
  }

  function loadRecord(loteId, turnoCampo, fundo) {
    const tc = turnoCampo === "Tarde" ? "Tarde" : "Mañana";
    state.session.turnoCampo = tc;
    persistSession();
    paintPeople({ keepTurno: true });
    const rec = APP.API.recordOf(loteId, today(), tc, supDni(), fundo || state.session.fundo);
    const L = APP.Data.findLote(loteId, (rec && rec.fundo) || fundo || state.session.fundo) || (rec && {
      lote: rec.lote,
      md: rec.md,
      modulo: rec.md ? "M" + rec.md : "",
      turno: rec.turno,
      area: rec.area,
      fundo: rec.fundo,
      variedad: rec.variedad,
      etapa: rec.etapa,
      grupo: rec.grupo,
      codLote: rec.codLote,
    });
    if (!L || !rec) return;
    fillLoteForm(L, rec);
  }

  function resetLoteInputs() {
    state.lote = null;
    state.avance = "";
    state.conv = false;
    state.china = false;
    state.jarrasConv = "";
    state.jarrasChina = "";
    state.cierre = null;
    $("#inp-avance").value = "";
    $("#inp-jconv").value = "";
    $("#inp-jchina").value = "";
    syncTurnoCampo();
    paintLote();
    paintChecks();
    paintAvanceHint();
    paintSaveBtn();
  }

  function emptyWorkspace() {
    resetLoteInputs();
    state.session.scannerDni = "";
    state.session.scannerNombre = "";
    state.session.supervisorDni = "";
    state.session.supervisorNombre = "";
    state.session.jornales = 0;
    state.session.grupo = "";
    state.session.etapa = "";
    state.session.turnoCampo = "Mañana";
    persistSession();
    paintPeople();
    paintDay();
    paintStatus();
  }

  async function saveLote() {
    if (state.saving) return;
    state.saving = true;
    try {
      if (!String(state.session.supervisorDni || "").trim()) {
        await feedback("Falta el supervisor", "Selecciona el supervisor.");
        return;
      }
      if (!canonFundo(state.session.fundo)) {
        await feedback("Falta el fundo", "Selecciona LICAPA I, II, III o IV.");
        return;
      }
      if (!state.lote) {
        await feedback("Falta el lote", "Selecciona un lote.");
        return;
      }
      if (!state.conv && !state.china) {
        await feedback("Marca el tipo", "Convencional y/o China.");
        return;
      }
      const av = applyAvance(state.avance, true);
      const maxHa = avanceTope();
      if (!(av.num > 0)) {
        await feedback("Avance inválido", "Escribe solo lo que avanzas ahora.");
        return;
      }
      if (maxHa != null && av.num > maxHa) {
        await feedback("Avance máximo", state.cierre ? `Solo falta ${haTxt(maxHa)} ha.` : `No puede superar ${maxHa} ha.`);
        return;
      }
      const floor = avanceFloor();
      if (floor > 0 && av.num < floor) {
        await feedback("Avance mínimo", `En la tarde no puede bajar de ${floor} ha. Solo se sube.`);
        return;
      }
      const data = currentData();
      const d = APP.Data.derive(data);
      if (!(d.totalJarras > 0)) {
        await feedback("Faltan jarras", "Ingresa la cantidad.");
        return;
      }
      data.kgConv = d.kgConvEff;
      data.kgChina = d.kgChinaEff;
      data.totalJarras = d.totalJarras;
      data.totalKg = d.totalKg;
      state.session.grupo = String(data.grupo);
      state.session.etapa = String(data.etapa);
      persistSession();
      showLoader("Cargando", "Guardando lote…");
      const saved = await APP.API.submit({ data });
      if (!saved || !saved.ok) {
        hideLoader();
        if (saved && saved.reason === "closed") {
          await feedback("Día cerrado", "Ya se envió la tarde. No se agregan lotes.");
        } else {
          await feedback("No se guardó en el celular", "Intenta otra vez.");
        }
        return;
      }
      rememberExcel();
      resetLoteInputs();
      paintDay();
      refreshOpenLots(true);
      paintStatus();
      hideLoader();
      feedback(
        "Lote guardado",
        `<span class="qb-fb-line"><em>Lote</em><b>${esc(data.lote)}</b></span>
         <span class="qb-fb-line"><em>Jarras</em><b>${esc(d.totalJarras)}</b></span>
         <span class="qb-fb-line"><em>Peso aproximado</em><b>${esc(d.totalKg)} kg</b></span>`,
        { html: true, ms: 2800 }
      );
    } catch (e) {
      hideLoader();
      await feedback("No se guardó en el celular", String(e.message || e));
    } finally {
      state.saving = false;
    }
  }

  function bind() {
    const on = (id, ev, fn) => {
      const el = $(id);
      if (el) el.addEventListener(ev, fn);
    };
    const click = (id, fn) => {
      const el = $(id);
      if (el) el.onclick = fn;
    };
    click("#trig-supervisor", async () => {
      await APP.Data.load();
      APP.PreciseSelect.open({
        title: "Supervisor",
        placeholder: "Buscar DNI o nombre…",
        getOptions: (q) => APP.Data.supervisorOptions(q),
        onSelect: async (opt) => {
          const nextDni = String(opt.dni || opt.id || "");
          const nextName = opt.nombre || opt.label;
          const prevDni = String(state.session.supervisorDni || "").trim();
          if (nextDni && nextDni === prevDni) return;
          // Primera vez: elegir sin preguntar. Solo confirmar si ya había supervisor.
          if (prevDni) {
            const n = APP.API.todayCount(prevDni);
            const ok = await ask(
              "Cambiar supervisor",
              n
                ? "Se elimina el día de este supervisor en el celular (lotes de hoy). El historial no se toca. ¿Seguro?"
                : "Vas a cambiar de supervisor. El historial no se toca. ¿Seguro?",
              "Cambiar"
            );
            if (!ok) return;
            APP.API.purgeTodayLocal(prevDni);
          }
          state.session.supervisorDni = nextDni;
          state.session.supervisorNombre = nextName;
          state.session.turnoCampo = APP.API.sendCount(nextDni) >= 1 ? "Tarde" : "Mañana";
          state.session.scannerDni = "";
          state.session.scannerNombre = "";
          state.session.jornales = 0;
          state.session.grupo = "";
          state.session.etapa = "";
          resetLoteInputs();
          persistSession();
          paintPeople();
          paintDay();
          paintHistorial();
          paintStatus();
        },
      });
    });
    click("#trig-scanner", async () => {
      await APP.Data.load();
      APP.PreciseSelect.open({
        title: "Escáner",
        placeholder: "Buscar DNI o nombre…",
        empty: "Sin resultados. Escribe el DNI de 8 dígitos para ingresarlo.",
        allowManual: true,
        findKnown: (dni) => APP.Data.findScanner(dni),
        getOptions: (q) => APP.Data.scannerOptions(q),
        onSelect: (opt) => {
          state.session.scannerDni = opt.dni || opt.id;
          state.session.scannerNombre = APP.Data.cleanNombre(opt.nombre || opt.label);
          persistSession();
          paintPeople();
        },
      });
    });
    click("#trig-lote", async () => {
      const fundo = canonFundo(state.session.fundo);
      if (!fundo) {
        await feedback("Falta el fundo", "Selecciona LICAPA I, II, III o IV.");
        return;
      }
      await APP.Data.load();
      if (!APP.Data.lotes.length) await APP.Data.load({ force: true });
      APP.PreciseSelect.open({
        title: "Lote · " + fundo,
        placeholder: "Número de lote",
        empty: APP.Data.lotes.length ? "Sin lotes de " + fundo : "No se cargaron los lotes. Cierra y vuelve a abrir.",
        getOptions: (q) => {
          const opts = APP.Data.loteOptions(q, fundo);
          return opts
            .map((o) => {
              const tc = isTarde() ? "Tarde" : "Mañana";
              const rec = APP.API.recordOf(o.id, today(), tc, supDni(), fundo) || (isTarde() ? APP.API.recordOf(o.id, today(), "Mañana", supDni(), fundo) : null);
              if (!rec) return o;
              const av = Number(rec.avance);
              const ha = Number.isFinite(av) && av > 0 ? av : 0;
              const L = o.lote || {};
              const area = L.area !== "" && L.area != null ? L.area : "—";
              const lugar = [L.fundo, L.modulo || (L.md ? "M" + L.md : ""), L.turno ? "T" + L.turno : "", L.variedad]
                .filter(Boolean)
                .join(" · ");
              return {
                ...o,
                marked: true,
                note: `Avance ${ha} ha · Área ${area} ha`,
                meta: lugar,
              };
            })
            .sort((a, b) => Number(!!b.marked) - Number(!!a.marked));
        },
        onSelect: (opt) => {
          state.cierre = null;
          const L = opt.lote || APP.Data.findLote(opt.id, fundo);
          if (!L) return;
          const tarde = APP.API.recordOf(L.lote, today(), "Tarde", supDni(), fundo);
          const manana = APP.API.recordOf(L.lote, today(), "Mañana", supDni(), fundo);
          fillLoteForm(L, tarde || (isTarde() ? null : manana));
          persistSession();
        },
      });
    });

    const turnoBtn = $("#btn-turno-campo");
    const turnoMenu = $("#menu-turno-campo");
    const closeTurno = () => {
      if (!turnoMenu || !turnoBtn) return;
      turnoMenu.hidden = true;
      turnoBtn.setAttribute("aria-expanded", "false");
    };
    if (turnoBtn && turnoMenu) {
      turnoBtn.addEventListener("click", (e) => {
        e.stopPropagation();
        const open = turnoMenu.hidden;
        turnoMenu.hidden = !open;
        turnoBtn.setAttribute("aria-expanded", open ? "true" : "false");
      });
      turnoMenu.addEventListener("click", (e) => {
        const opt = e.target.closest("[data-turno]");
        if (!opt) return;
        if (APP.API.sendCount(state.session.supervisorDni) >= 1) {
          state.session.turnoCampo = "Tarde";
        } else {
          state.session.turnoCampo = opt.dataset.turno === "Tarde" ? "Tarde" : "Mañana";
        }
        persistSession();
        paintPeople();
        paintDay();
        closeTurno();
      });
      document.addEventListener("click", (e) => {
        if (!$("#sel-turno-campo")?.contains(e.target)) closeTurno();
      });
    }
    const fundoBtn = $("#btn-fundo");
    const fundoMenu = $("#menu-fundo");
    const closeFundo = () => {
      if (!fundoMenu || !fundoBtn) return;
      fundoMenu.hidden = true;
      fundoBtn.setAttribute("aria-expanded", "false");
    };
    if (fundoBtn && fundoMenu) {
      fundoBtn.addEventListener("click", (e) => {
        e.stopPropagation();
        const open = fundoMenu.hidden;
        fundoMenu.hidden = !open;
        fundoBtn.setAttribute("aria-expanded", open ? "true" : "false");
      });
      fundoMenu.addEventListener("click", (e) => {
        const opt = e.target.closest("[data-fundo]");
        if (!opt) return;
        state.session.fundo = canonFundo(opt.dataset.fundo);
        if (state.lote && !APP.Data.findLote(state.lote.lote, state.session.fundo)) resetLoteInputs();
        persistSession();
        paintPeople();
        closeFundo();
      });
      document.addEventListener("click", (e) => {
        if (!$("#sel-fundo")?.contains(e.target)) closeFundo();
      });
    }
    on("#inp-jornales", "input", () => {
      const el = $("#inp-jornales");
      if (!el) return;
      state.session.jornales = Math.max(0, Number(String(el.value).replace(/\D/g, "")) || 0);
      persistSession();
    });
    on("#inp-avance", "input", () => applyAvance($("#inp-avance")?.value));
    on("#inp-avance", "blur", () => applyAvance($("#inp-avance")?.value, true));
    on("#inp-peso", "input", () => {
      const el = $("#inp-peso");
      const text = sanitizePeso(el?.value);
      state.session.pesoJarra = text;
      if (el && el.value !== text) el.value = text;
      persistSession();
      paintTotals();
    });
    on("#inp-peso", "blur", () => {
      state.session.pesoJarra = String(pesoNum());
      const el = $("#inp-peso");
      if (el) el.value = state.session.pesoJarra;
      persistSession();
      paintTotals();
    });
    on("#inp-jconv", "input", () => {
      const el = $("#inp-jconv");
      state.jarrasConv = String(el?.value || "").replace(/\D/g, "");
      if (el && el.value !== state.jarrasConv) el.value = state.jarrasConv;
      if (state.jarrasConv && !state.conv) {
        state.conv = true;
        paintChecks();
        return;
      }
      paintTotals();
    });
    on("#inp-jchina", "input", () => {
      const el = $("#inp-jchina");
      state.jarrasChina = String(el?.value || "").replace(/\D/g, "");
      if (el && el.value !== state.jarrasChina) el.value = state.jarrasChina;
      if (state.jarrasChina && !state.china) {
        state.china = true;
        paintChecks();
        return;
      }
      paintTotals();
    });
    on("#chk-conv", "change", () => {
      state.conv = !!$("#chk-conv")?.checked;
      if (!state.conv) state.jarrasConv = "";
      const inp = $("#inp-jconv");
      if (inp) inp.value = state.jarrasConv;
      paintChecks();
    });
    on("#chk-china", "change", () => {
      state.china = !!$("#chk-china")?.checked;
      if (!state.china) state.jarrasChina = "";
      const inp = $("#inp-jchina");
      if (inp) inp.value = state.jarrasChina;
      paintChecks();
    });
    ensureExportSheet();
    click("#btn-save", saveLote);
    click("#btn-export", openExport);
    click("#btn-historial", openHistorial);
    click("#btn-hist-back", closeHistorial);
    click("#btn-sync", openSync);
    click("#btn-sync-close", closeSync);
    on("#qb-sync", "click", (e) => {
      if (e.target === $("#qb-sync")) closeSync();
    });
    click("#btn-wipe", wipeCache);
    click("#btn-update", updateApp);
    click("#btn-transfer", transferMode);
    click("#btn-export-close", closeExport);
    on("#qb-export", "click", (e) => {
      if (e.target === $("#qb-export")) closeExport();
    });
    click("#btn-export-share", exportShare);
    click("#btn-export-dl", exportDownload);
    click("#btn-export-send", () => {
      closeExport();
      transferMode();
    });

    const pendingBtn = $("#chip-pending");
    if (pendingBtn) {
      pendingBtn.onclick = () => {
        const n = APP.API.pendingCount();
        feedback(
          n ? (n === 1 ? "1 envío por confirmar" : `${n} envíos por confirmar`) : "Nada por enviar",
          n
            ? "Un Enviar = 1 pendiente. Baja a 0 cuando el servidor responde ok."
            : "Los lotes Guardados están en el celular. El pendiente aparece al pulsar Enviar."
        );
      };
    }

    window.addEventListener("online", paintStatus);
    window.addEventListener("offline", paintStatus);
    let activityTick = 0;
    window.addEventListener("app:activity", () => {
      if (activityTick) return;
      activityTick = requestAnimationFrame(() => {
        activityTick = 0;
        paintDay();
        paintStatus();
      });
    });
  }

  function syncJornales() {
    const el = $("#inp-jornales");
    if (!el) return;
    state.session.jornales = Math.max(0, Number(String(el.value).replace(/\D/g, "")) || 0);
  }

  function draftRecord() {
    const data = currentData();
    if (!String(data.lote || "").trim()) return null;
    const d = APP.Data.derive(data);
    if (!(d.totalJarras > 0) && !(Number(data.avance) > 0)) return null;
    return {
      ...data,
      kgConv: d.kgConvEff,
      kgChina: d.kgChinaEff,
      totalJarras: d.totalJarras,
      totalKg: d.totalKg,
    };
  }

  function liveRecords() {
    const list = dayRecords().slice();
    const draft = draftRecord();
    if (!draft) return list;
    const tc = draft.turnoCampo || "Mañana";
    const fa = String(draft.fundo || "").trim().toUpperCase();
    const i = list.findIndex((r) => {
      const fb = String(r.fundo || "").trim().toUpperCase();
      return String(r.lote) === String(draft.lote) && (r.turnoCampo || "Mañana") === tc && (!fa || !fb || fa === fb);
    });
    if (i >= 0) list[i] = { ...list[i], ...draft };
    else list.push(draft);
    return list.sort((a, b) => String(a.lote).localeCompare(String(b.lote), "es", { numeric: true }));
  }

  function exportPack() {
    syncJornales();
    persistSession();
    return APP.Excel.buildFile({
      session: state.session,
      fecha: today(),
      records: liveRecords(),
    });
  }

  function ensureExportSheet() {
    const sheet = document.querySelector("#qb-export .qb-export-sheet");
    if (!sheet) return;
    let preview = $("#export-preview");
    if (!preview) {
      preview = document.createElement("div");
      preview.className = "export-card";
      preview.id = "export-preview";
      const oldCard = $("#export-card");
      if (oldCard) oldCard.replaceWith(preview);
      else {
        const head = sheet.querySelector(".qb-export-head");
        if (head) head.insertAdjacentElement("afterend", preview);
        else sheet.insertBefore(preview, sheet.firstChild);
      }
    }
    let send = $("#btn-export-send");
    if (!send) {
      send = document.createElement("button");
      send.type = "button";
      send.className = "export-send";
      send.id = "btn-export-send";
      send.textContent = "Enviar";
      preview.insertAdjacentElement("afterend", send);
    }
    let share = $("#btn-export-share");
    let dl = $("#btn-export-dl");
    if (!share) {
      share = document.createElement("button");
      share.type = "button";
      share.id = "btn-export-share";
    }
    if (!dl) {
      dl = document.createElement("button");
      dl.type = "button";
      dl.id = "btn-export-dl";
    }
    share.className = "export-choice";
    share.textContent = "Exportar";
    dl.className = "export-choice ghost";
    dl.textContent = "Descargar";
    let row = share.closest(".export-row");
    if (!row) {
      row = document.createElement("div");
      row.className = "export-row";
      send.insertAdjacentElement("afterend", row);
    }
    if (share.parentElement !== row) row.appendChild(share);
    if (dl.parentElement !== row) row.appendChild(dl);
    send.onclick = () => {
      closeExport();
      transferMode();
    };
    share.onclick = exportShare;
    dl.onclick = exportDownload;
  }

  function openExport() {
    ensureExportSheet();
    const pack = exportPack();
    paintExportPreview(pack.model);
    const root = $("#qb-export");
    if (root) root.hidden = false;
  }

  function esc(s) {
    return String(s ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function paintExportPreview(model) {
    const box = $("#export-preview");
    if (!box) return;
    const t = model.totals || {};
    const rows = model.rows || [];
    const fundo = rows[0]?.fundo || "LICAPA";
    const etapa = model.etapa ? (String(model.etapa).toUpperCase().includes("LICAPA") ? model.etapa : "Licapa " + model.etapa) : "Licapa I";
    const body = rows.length
      ? rows
          .map((r) => {
            const area = r.area !== "" && r.area != null ? r.area : "—";
            const loc = `M${esc(r.md || "—")} · T${esc(r.turno || "—")} · ${esc(area)} ha`;
            return `<tr>
              <td>${esc(r.variedad || "—")}</td>
              <td>${esc(r.lote)}</td>
              <td>${loc}</td>
              <td class="num">${esc(r.totalJarras)}</td>
              <td class="num">${esc(r.totalKg)}</td>
            </tr>`;
          })
          .join("")
      : `<tr><td colspan="5" class="empty-row">Aún no hay lotes hoy.</td></tr>`;
    box.innerHTML = `
      <article class="guide-card">
        <header class="guide-head">
          <img src="assets/logo-qberries.png" alt="" width="36" height="36" />
          <div>
            <p>REGISTRO DE COSECHA</p>
            <strong>Q Berries · ${esc(etapa)}</strong>
          </div>
        </header>
        <div class="guide-meta">
          <div class="guide-meta-row">
            <div><small>Fecha</small><b>${esc(fmtDate(model.fecha))}</b></div>
            <div><small>Turno campo</small><b>${esc(model.turnoCampo || "Mañana")}</b></div>
          </div>
          <div class="guide-meta-sup">
            <small>Supervisor</small>
            <b>${esc(model.supervisor || "—")}</b>
          </div>
        </div>
        <table class="guide-table">
          <thead>
            <tr>
              <th>Variedad</th>
              <th>Lote</th>
              <th>M · T · Área</th>
              <th>Jarras</th>
              <th>Kg</th>
            </tr>
          </thead>
          <tbody>${body}</tbody>
        </table>
        <div class="guide-rates">
          <div><small>Kg / ha</small><b>${esc(t.kgHa || 0)}</b></div>
          <div><small>Kg / jn</small><b>${esc(t.kgJn || 0)}</b></div>
          <div><small>Jornales</small><b>${esc(t.jornales || 0)}</b></div>
        </div>
        <footer class="guide-foot">
          <span>Total</span>
          <strong>${rows.length} lote${rows.length === 1 ? "" : "s"} · ${esc(t.totalJarras || 0)} jarras · ${esc(t.totalKg || 0)} kg</strong>
        </footer>
      </article>`;
  }

  function closeExport() {
    $("#qb-export").hidden = true;
  }

  function hoursLeft(expiresAt) {
    const ms = Number(expiresAt) - Date.now();
    if (ms <= 0) return "Venció";
    const h = Math.max(1, Math.ceil(ms / 3600000));
    return h === 1 ? "Queda 1 h" : `Quedan ${h} h`;
  }

  function paintHistorial() {
    const box = $("#hist-list");
    if (!box) return;
    const items = APP.API.listHistory();
    box.innerHTML = items.length
      ? items
          .map((item) => {
            const pending = item.pending !== false;
            return `<article class="hist-card">
              <div class="hist-card-top">
                <div>
                  <strong>${esc(fmtDate(item.fecha))} · ${esc(item.turnoCampo || "Mañana")}</strong>
                  <small>${esc(item.supervisor || "—")}</small>
                </div>
                <span class="hist-badge${pending ? "" : " ok"}">${pending ? "Pendiente" : "Enviado"}</span>
              </div>
              <div class="hist-meta">${item.lotes || 0} lote${item.lotes === 1 ? "" : "s"} · ${esc(item.jarras || 0)} jarras · ${esc(item.kg || 0)} kg</div>
              <div class="hist-ttl">${hoursLeft(item.expiresAt)} para descargar el Excel</div>
              <button type="button" class="hist-dl" data-hist="${esc(item.id)}">Descargar Excel</button>
            </article>`;
          })
          .join("")
      : `<p class="empty">No hay reportes de las últimas 24 horas.</p>`;
    box.querySelectorAll("[data-hist]").forEach((btn) => {
      btn.onclick = () => downloadHistory(btn.dataset.hist);
    });
  }

  function openHistorial() {
    rememberExcel();
    paintHistorial();
    paintStatus();
    const view = $("#view-historial");
    if (view) view.hidden = false;
  }

  function closeHistorial() {
    const view = $("#view-historial");
    if (view) view.hidden = true;
  }

  function downloadHistory(id) {
    const item = APP.API.getHistory(id);
    if (!item || !item.records || !item.records.length) {
      feedback("Sin Excel", "Este reporte ya no está en el celular.");
      return;
    }
    APP.Excel.download({ session: item.session || state.session, fecha: item.fecha, records: item.records });
    feedback("Excel descargado", item.name || "reporte-cosecha.xls");
  }

  function openSync() {
    const ver = $("#sync-version");
    if (ver) ver.textContent = "v" + (APP.CONFIG.VERSION || "2.0.52");
    const root = $("#qb-sync");
    if (root) root.hidden = false;
  }

  function closeSync() {
    const root = $("#qb-sync");
    if (root) root.hidden = true;
  }

  async function wipeCache() {
    const ok = await ask(
      "Eliminar caché",
      "¿Seguro que quieres eliminar todo el caché? Se borra lo guardado en el celular y la app se carga limpia.",
      "Sí, eliminar"
    );
    if (!ok) return;
    closeSync();
    APP.API.wipeLocal();
    try {
      if ("serviceWorker" in navigator) {
        const regs = await navigator.serviceWorker.getRegistrations();
        await Promise.all(regs.map((r) => r.unregister()));
      }
      if (window.caches) {
        const keys = await caches.keys();
        await Promise.all(keys.map((k) => caches.delete(k)));
      }
    } catch (_) {}
    location.reload();
  }

  async function updateApp() {
    closeSync();
    const started = Date.now();
    showLoader("Cargando", "Actualizando…");
    try {
      await wait(500);
      showLoader("Cargando", "Buscando en la nube…");
      const cloudFiles = [
        "./js/config.js",
        "./index.html",
        "./css/styles.css",
        "./js/app.js",
        "./js/data.js",
        "./js/api.js",
        "./js/catalog-lotes.js",
        "./sw.js",
      ];
      for (const file of cloudFiles) {
        await fetchCloud(file);
        await wait(180);
      }
      showLoader("Actualizando", "Descargando la última versión…");
      if ("serviceWorker" in navigator) {
        const regs = await navigator.serviceWorker.getRegistrations();
        await Promise.all(regs.map((r) => r.unregister()));
      }
      if (window.caches) {
        const keys = await caches.keys();
        await Promise.all(keys.map((k) => caches.delete(k)));
      }
      if ("serviceWorker" in navigator && /^https?:$/i.test(location.protocol)) {
        await navigator.serviceWorker.register("./sw.js?v=" + Date.now());
        if (navigator.serviceWorker.ready) await navigator.serviceWorker.ready;
      }
      const left = 3200 - (Date.now() - started);
      if (left > 0) await wait(left);
      showLoader("Actualizando", "Listo. Reiniciando…");
      await wait(700);
      location.reload();
    } catch (_) {
      hideLoader();
      await feedback("Sin red", "No se pudo buscar la actualización en la nube. Intenta otra vez.");
    }
  }

  function sendTotalsLabel(records) {
    const list = (records || []).filter((r) => r && String(r.lote || "").trim());
    const jarras = list.reduce((a, r) => a + (Number(r.totalJarras) || 0), 0);
    const kg = APP.Data.round2(list.reduce((a, r) => a + (Number(r.totalKg) || 0), 0));
    const n = list.length;
    return `${n} lote${n === 1 ? "" : "s"} · ${jarras} jarras · ${kg} kg`;
  }

  function unsavedDraftWarning() {
    const draft = draftRecord();
    if (!draft) return "";
    const tc = draft.turnoCampo || "Mañana";
    const saved = APP.API.recordOf(draft.lote, today(), tc, supDni(), draft.fundo);
    if (!saved) {
      return `Hay un lote ${draft.lote} en pantalla sin Guardar. No irá en el envío. Pulsa Guardar lote primero.`;
    }
    const same =
      Number(saved.totalJarras) === Number(draft.totalJarras) &&
      Number(saved.totalKg) === Number(draft.totalKg) &&
      String(saved.avance) === String(draft.avance);
    if (same) return "";
    return `El lote ${draft.lote} tiene cambios sin Guardar. El envío usa lo guardado, no lo de la pantalla.`;
  }

  async function transferMode(opts) {
    const auto = !!(opts && opts.auto);
    if (state.transferring) return;
    state.transferring = true;
    closeSync();
    try {
      const warn = unsavedDraftWarning();
      if (warn) {
        if (!auto) await feedback("Guarda antes de enviar", warn);
        return;
      }
      const dni = String(state.session.supervisorDni || "").trim();
      if (!dni) {
        if (!auto) await feedback("Falta el supervisor", "Selecciona el supervisor para enviar.");
        return;
      }
      const sent = APP.API.sendCount(dni);
      if (sent >= 2) {
        if (!auto) await feedback("Ya enviaste 2 veces", "Hoy ya se envió mañana y tarde con este supervisor.");
        return;
      }
      const isSecond = sent >= 1;
      const todayN = APP.API.todayCount(dni);
      if (!todayN) {
        if (!auto) await feedback("Sin lotes", "Primero guarda los lotes del día.");
        return;
      }
      const turnoEnvio = isSecond ? "Tarde" : "Mañana";
      state.session.turnoCampo = turnoEnvio;
      persistSession();
      paintPeople();
      if (isSecond) APP.API.queueTodayForResend(dni);
      const queue = APP.API.pendingRecords(dni).filter((r) => String(r.fecha) === today());
      if (!queue.length) {
        if (!auto) await feedback("Ya está enviado", "No hay nada nuevo por enviar.");
        return;
      }
      const totalTxt = sendTotalsLabel(queue);
      if (!auto) {
        const ok = isSecond
          ? await ask(
              "Enviar turno tarde",
              `Se enviará el reporte de tarde: ${totalTxt}. Al terminar se limpia el día en el celular. ¿Continuar?`,
              "Enviar"
            )
          : await ask(
              "Enviar turno mañana",
              `Se enviará el reporte de mañana: ${totalTxt}. ¿Continuar?`,
              "Enviar"
            );
        if (!ok) return;
      }

      APP.API.queueReport({
        turnoCampo: turnoEnvio,
        summary: totalTxt,
        records: queue,
      });
      paintStatus();

      if (typeof navigator !== "undefined" && navigator.onLine === false) {
        if (!auto) {
          await feedback(
            "Sin señal",
            "Quedó guardado en el celular. Se enviará solo cuando vuelva la red.",
            { ms: 2200 }
          );
        }
        return;
      }

      const result = await APP.API.flush({
        summary: totalTxt,
        turnoCampo: turnoEnvio,
        supervisorDni: dni,
      });
      paintDay();
      paintHistorial();
      paintStatus();

      if (result.reason === "no-ep" || result.reason === "offline" || result.reason === "net" || result.reason === "http" || result.reason === "partial") {
        if (!auto) {
          await feedback(
            "Quedó pendiente",
            `${totalTxt}. Puedes seguir trabajando. Cuando haya red, pulsa Enviar otra vez.`,
            { ms: 2400 }
          );
        }
        return;
      }
      if (result.left > 0) {
        if (!auto) {
          await feedback(
            "No terminó el envío",
            `Parte del reporte quedó pendiente. Tus datos siguen en el celular. Pulsa Enviar otra vez.`
          );
        }
        return;
      }

      APP.API.clearOutboxReport(turnoEnvio);
      APP.API.bumpSend(dni);
      if (!isSecond) {
        state.session.turnoCampo = "Tarde";
        persistSession();
      }
      paintPeople();
      paintDay();
      paintStatus();

      if (isSecond) {
        rememberExcel(null, "Tarde");
        APP.API.markTransferred(dni);
        const wiped = APP.API.clearTodayRecords(dni);
        if (!wiped || !wiped.ok) {
          if (!auto) await feedback("Revisa el envío", "Algo quedó sin confirmar. No se borró el día. Pulsa Enviar otra vez.");
          return;
        }
        APP.API.clearOutboxAll();
        // Solo limpia el formulario de este supervisor. No toca lotes de otros ni el historial.
        resetLoteInputs();
        state.session.scannerDni = "";
        state.session.scannerNombre = "";
        state.session.jornales = 0;
        state.session.grupo = "";
        state.session.etapa = "";
        state.session.turnoCampo = "Mañana";
        persistSession();
        paintPeople();
        paintDay();
        paintHistorial();
        paintStatus();
      toast("Se envió.");
        return;
      }

      rememberExcel(null, "Mañana");
      paintHistorial();
      paintStatus();
      toast("Se envió.");
    } catch (e) {
      hideLoader();
      paintStatus();
      if (!auto) {
        await feedback("No se pudo enviar", "Tus lotes siguen en el celular. Si quedó 1 pendiente, reintenta con señal.");
      }
    } finally {
      hideLoader();
      state.transferring = false;
      paintStatus();
    }
  }

  let autoSyncTimer = 0;
  let autoSyncBusy = false;

  function hasInterruptedSend() {
    try {
      const dni = String(state.session.supervisorDni || "").trim();
      return APP.API.pendingCount() > 0 && APP.API.pendingRecords(dni || undefined).length > 0;
    } catch (_) {
      return false;
    }
  }

  function scheduleAutoSync(delayMs) {
    if (autoSyncTimer) clearTimeout(autoSyncTimer);
    autoSyncTimer = setTimeout(() => {
      autoSyncTimer = 0;
      tryAutoSync();
    }, Math.max(0, delayMs == null ? 2800 : delayMs));
  }

  async function tryAutoSync() {
    if (autoSyncBusy || state.transferring) {
      if (hasInterruptedSend()) scheduleAutoSync(10000);
      return;
    }
    if (!hasInterruptedSend()) return;
    if (unsavedDraftWarning()) return;
    const dni = String(state.session.supervisorDni || "").trim();
    if (!dni) return;
    if (APP.API.sendCount(dni) >= 2) return;

    autoSyncBusy = true;
    try {
      // Un solo camino: transferMode → flush (ahí hace el probe). Sin doble ping.
      await transferMode({ auto: true });
      if (hasInterruptedSend()) scheduleAutoSync(45000);
    } catch (_) {
      if (hasInterruptedSend()) scheduleAutoSync(45000);
    } finally {
      autoSyncBusy = false;
      paintStatus();
    }
  }

  function exportDownload() {
    const pack = exportPack();
    rememberExcel(pack);
    APP.Excel.download({ session: state.session, fecha: today(), records: liveRecords() });
    closeExport();
    paintStatus();
    feedback("Excel descargado", pack.name);
  }

  async function exportShare() {
    const pack = exportPack();
    try {
      if (navigator.canShare && pack.file && navigator.canShare({ files: [pack.file] })) {
        await navigator.share({
          files: [pack.file],
          title: `Cosecha ${fmtDate(pack.model.fecha)}`,
          text: `REPORTE DIARIO DE COSECHA · ${fmtDate(pack.model.fecha)}`,
        });
        rememberExcel(pack);
        closeExport();
        paintStatus();
        await feedback("Excel exportado", "Listo para enviar.");
        return;
      }
    } catch (e) {
      if (e && e.name === "AbortError") return;
    }
    exportDownload();
  }

  function toast(text, warn) {
    const el = $("#qb-toast");
    if (!el) return;
    el.textContent = text || "";
    el.classList.toggle("is-warn", !!warn);
    el.hidden = false;
    clearTimeout(toast._t);
    toast._t = setTimeout(() => { el.hidden = true; }, 2200);
  }

  function paintStatus() {
    const on = navigator.onLine;
    document.querySelectorAll("#chip-online, [data-chip-online]").forEach((el) => {
      el.classList.toggle("offline", !on);
      el.classList.toggle("online", on);
      const text = el.querySelector(".chip-text");
      if (text) text.textContent = on ? "En línea" : "Sin red";
    });
    const n = APP.API.pendingCount();
    document.querySelectorAll("#chip-pending-text, [data-chip-pending-text]").forEach((el) => {
      el.textContent = n === 1 ? "1 pend." : n + " pend.";
    });
    document.querySelectorAll("#chip-pending, [data-chip-pending]").forEach((el) => {
      el.hidden = false;
      el.classList.toggle("has-items", n > 0);
    });
  }

  let deferredInstall = null;
  const INSTALL_DISMISS_KEY = "qb_install_dismiss_until";

  function isAppInstalled() {
    try {
      if (window.matchMedia && window.matchMedia("(display-mode: standalone)").matches) return true;
      if (window.matchMedia && window.matchMedia("(display-mode: fullscreen)").matches) return true;
      if (navigator.standalone === true) return true;
      const q = new URLSearchParams(location.search || "");
      if (q.get("source") === "pwa") return true;
    } catch (_) {}
    return false;
  }

  function isIosBrowser() {
    const ua = navigator.userAgent || "";
    return /iPad|iPhone|iPod/.test(ua) || (navigator.platform === "MacIntel" && navigator.maxTouchPoints > 1);
  }

  function installDismissed() {
    try {
      return Number(localStorage.getItem(INSTALL_DISMISS_KEY) || 0) > Date.now();
    } catch (_) {
      return false;
    }
  }

  function dismissInstall(days) {
    try {
      localStorage.setItem(INSTALL_DISMISS_KEY, String(Date.now() + (days == null ? 3 : days) * 86400000));
    } catch (_) {}
    hideInstall();
  }

  function hideInstall() {
    const root = $("#qb-install");
    if (root) root.hidden = true;
  }

  function wantsForcedInstall() {
    try {
      return new URLSearchParams(location.search || "").get("instalar") === "1";
    } catch (_) {
      return false;
    }
  }

  function showInstall() {
    if (isAppInstalled()) return;
    if (installDismissed() && !wantsForcedInstall()) return;
    const root = $("#qb-install");
    if (!root) return;
    const ios = isIosBrowser();
    const iosHint = $("#install-ios-hint");
    const lead = $("#install-lead");
    const btnNow = $("#btn-install-now");
    const btnLater = $("#btn-install-later");
    const title = $("#install-title");

    if (ios) {
      if (title) title.textContent = "Cómo instalar en iPhone";
      if (lead) lead.hidden = true;
      if (iosHint) iosHint.hidden = false;
      if (btnNow) btnNow.hidden = true;
      if (btnLater) btnLater.textContent = "Entendido";
    } else {
      if (title) title.textContent = "Instala la app";
      if (lead) {
        lead.hidden = false;
        lead.textContent = "Queda en tu celular como app. Si cierras Chrome o se apaga el teléfono, tus lotes siguen hasta que envíes el turno.";
      }
      if (iosHint) iosHint.hidden = true;
      if (btnNow) {
        btnNow.hidden = false;
        btnNow.textContent = "Instalar";
        btnNow.disabled = false;
      }
      if (btnLater) btnLater.textContent = "Ahora no";
    }
    root.hidden = false;
  }

  async function runInstall() {
    if (isIosBrowser()) return;
    if (deferredInstall) {
      try {
        deferredInstall.prompt();
        const choice = await deferredInstall.userChoice;
        deferredInstall = null;
        if (choice && choice.outcome === "accepted") {
          hideInstall();
          try {
            localStorage.removeItem(INSTALL_DISMISS_KEY);
          } catch (_) {}
          return;
        }
      } catch (_) {}
    }
    await feedback("Instalar app", "En Chrome: menú ⋮ → Instalar app o Agregar a pantalla de inicio.");
  }

  function setupInstallPrompt() {
    if (isAppInstalled()) return;
    window.addEventListener("beforeinstallprompt", (e) => {
      e.preventDefault();
      deferredInstall = e;
      if (!installDismissed() || wantsForcedInstall()) showInstall();
    });
    window.addEventListener("appinstalled", () => {
      deferredInstall = null;
      hideInstall();
      try {
        localStorage.removeItem(INSTALL_DISMISS_KEY);
      } catch (_) {}
    });
    const btnNow = $("#btn-install-now");
    const btnLater = $("#btn-install-later");
    const root = $("#qb-install");
    if (btnNow) btnNow.onclick = () => { runInstall(); };
    if (btnLater) btnLater.onclick = () => dismissInstall(3);
    if (root) {
      root.addEventListener("click", (e) => {
        if (e.target === root) dismissInstall(1);
      });
    }
    setTimeout(() => {
      if (!isAppInstalled() && (!installDismissed() || wantsForcedInstall())) showInstall();
    }, 1200);
  }

  function allowMobileOrTablet() {
    const w = window.innerWidth || 0;
    const h = window.innerHeight || 0;
    const shortSide = Math.min(w, h);
    const longSide = Math.max(w, h);
    return shortSide >= 280 && shortSide <= 1024 && longSide <= 1400;
  }

  function setDesktopGate(on) {
    const gate = document.getElementById("desktop-gate");
    if (!gate) return;
    if (on) {
      gate.hidden = false;
      gate.removeAttribute("inert");
      gate.setAttribute("aria-hidden", "false");
      return;
    }
    gate.hidden = true;
    gate.setAttribute("inert", "");
    gate.setAttribute("aria-hidden", "true");
  }

  let appReady = false;

  function hideBoot() {
    const boot = document.getElementById("boot");
    if (!boot || boot.classList.contains("is-out")) return;
    boot.classList.add("is-out");
    setTimeout(() => boot.remove(), 400);
  }

  function bootApp() {
    if (appReady) return;
    appReady = true;
    loadSession();
    APP.API.pruneOldRecords();
    paintPeople();
    paintLote();
    paintChecks();
    paintDay();
    paintStatus();
    paintSaveBtn();
    bind();
    setupInstallPrompt();
    const refreshDay = () => {
      APP.API.pruneOldRecords();
      paintDay();
      paintStatus();
      paintHistorial();
    };
    document.addEventListener("visibilitychange", () => {
      if (!document.hidden) {
        refreshDay();
        if (hasInterruptedSend()) scheduleAutoSync(2000);
      }
    });
    window.addEventListener("focus", () => {
      refreshDay();
      if (hasInterruptedSend()) scheduleAutoSync(2000);
    });
    window.addEventListener("online", () => {
      refreshDay();
      if (hasInterruptedSend()) scheduleAutoSync(2800);
    });
    if (hasInterruptedSend()) scheduleAutoSync(1500);
    if (APP.Data && APP.Data.load) {
      APP.Data.load()
        .then(() => paintPeople())
        .catch(() => {});
    }
    window.addEventListener("app:catalogs", paintPeople);
    hideBoot();
    if ("serviceWorker" in navigator && /^https?:$/i.test(location.protocol)) {
      navigator.serviceWorker.register("./sw.js?v=" + (APP.CONFIG.VERSION || "")).catch(() => {});
      navigator.serviceWorker.addEventListener("controllerchange", () => {
        if (window.__qbReloading) return;
        window.__qbReloading = true;
        location.reload();
      });
    }
  }

  function applySizeGate() {
    const ok = allowMobileOrTablet();
    document.documentElement.classList.toggle("is-desktop", !ok);
    document.body.classList.toggle("is-desktop", !ok);
    setDesktopGate(!ok);
    if (ok) bootApp();
    else hideBoot();
  }

  function init() {
    applySizeGate();
    let sizeTimer = 0;
    window.addEventListener("resize", () => {
      clearTimeout(sizeTimer);
      sizeTimer = setTimeout(applySizeGate, 120);
    });
    window.addEventListener("orientationchange", () => setTimeout(applySizeGate, 180));
    setTimeout(() => {
      if (appReady && !isAppInstalled() && (!installDismissed() || wantsForcedInstall())) showInstall();
    }, 1200);
  }

  document.addEventListener("DOMContentLoaded", init);
  return { state };
})();
