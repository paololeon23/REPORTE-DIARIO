/** Flujo simple: supervisor → escáner → lote → avance → convencional/China */
window.APP = window.APP || {};

APP.App = (() => {
  const $ = (s) => document.querySelector(s);

  const state = {
    session: {
      scannerDni: "48055477",
      scannerNombre: "ZAVALETA GONZALES LUSBET ERELID",
      supervisorDni: "48533707",
      supervisorNombre: "JULCA GAMBOA DILMER ELICER",
      jornales: 36,
      grupo: "2",
      etapa: "1",
    },
    lote: null,
    avance: "",
    conv: false,
    china: false,
    jarrasConv: "",
    jarrasChina: "",
    saving: false,
  };

  function today() {
    return APP.API.todayKey();
  }

  function fmtDate(iso) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || "");
    return m ? `${m[3]}/${m[2]}/${m[1]}` : iso;
  }

  function feedback(title, text) {
    return new Promise((resolve) => {
      const root = $("#qb-fb");
      $("#qb-fb-title").textContent = title;
      $("#qb-fb-text").textContent = text || "";
      const actions = $("#qb-fb-actions");
      actions.innerHTML = "";
      const ok = document.createElement("button");
      ok.textContent = "Listo";
      ok.onclick = () => {
        root.hidden = true;
        resolve(true);
      };
      actions.appendChild(ok);
      root.hidden = false;
      setTimeout(() => {
        root.hidden = true;
        resolve(true);
      }, 1600);
    });
  }

  function persistSession() {
    APP.API.saveSession(state.session);
  }

  function loadSession() {
    const saved = APP.API.loadSession();
    if (saved && saved.scannerDni) state.session = { ...state.session, ...saved };
  }

  function paintPeople() {
    $("#lbl-supervisor").textContent = APP.Data.displayName(state.session.supervisorNombre) || "Seleccionar…";
    $("#lbl-scanner").textContent = APP.Data.displayName(state.session.scannerNombre) || "Seleccionar…";
    $("#inp-jornales").value = String(state.session.jornales || 0);
  }

  function paintLote() {
    const L = state.lote;
    $("#lbl-lote").textContent = L ? String(L.lote) : "Seleccionar lote…";
    $("#lbl-md").textContent = L ? L.md : "—";
    $("#lbl-turno").textContent = L ? L.turno : "—";
    $("#lbl-area").textContent = L ? `${L.area} ha` : "—";
    $("#lbl-extra").textContent = L ? `${L.fundo} · ${L.variedad} · Grupo ${L.grupo} · Etapa ${L.etapa}` : "";
  }

  function currentData() {
    const L = state.lote || {};
    const jConv = state.conv ? Number(state.jarrasConv) || 0 : 0;
    const jChina = state.china ? Number(state.jarrasChina) || 0 : 0;
    return {
      fecha: today(),
      lote: L.lote || "",
      fundo: L.fundo || "",
      variedad: L.variedad || "",
      md: L.md || "",
      turno: L.turno || "",
      area: L.area || "",
      avance: state.avance || L.area || "",
      jarrasConv: jConv,
      jarrasChina: jChina,
      kgConv: 0,
      kgChina: 0,
      grupo: L.grupo || state.session.grupo,
      etapa: L.etapa || state.session.etapa,
      jornales: Number(state.session.jornales) || 0,
      scanner: APP.Data.displayName(state.session.scannerNombre),
      scannerDni: state.session.scannerDni,
      supervisor: APP.Data.displayName(state.session.supervisorNombre),
      supervisorDni: state.session.supervisorDni,
    };
  }

  function paintTotals() {
    const d = APP.Data.derive(currentData());
    $("#live-totals").innerHTML = `<span>${d.totalJarras} jarras</span><span>${d.totalKg} kg</span>`;
  }

  function paintChecks() {
    $("#chk-conv").checked = state.conv;
    $("#chk-china").checked = state.china;
    $("#box-conv").hidden = !state.conv;
    $("#box-china").hidden = !state.china;
    paintTotals();
  }

  function dayModel() {
    const records = APP.API.localRecords(today());
    return APP.Excel.buildModel({ session: state.session, fecha: today(), records });
  }

  function paintDay() {
    const { records } = { records: APP.API.localRecords(today()) };
    const model = dayModel();
    const t = model.totals;
    $("#day-kpis").innerHTML = `
      <div><small>Lotes</small><strong>${model.filledCount}</strong></div>
      <div><small>Área</small><strong>${t.area} ha</strong></div>
      <div><small>Jarras</small><strong>${t.totalJarras}</strong></div>
      <div><small>Kg</small><strong>${t.totalKg}</strong></div>`;
    $("#lot-list").innerHTML = records.length
      ? records
          .map((r) => {
            const d = APP.Data.derive(r);
            return `<button type="button" class="lot-item" data-lote="${r.lote}">
              <div><strong>Lote ${r.lote}</strong><small>${r.md ? "MD " + r.md : ""} · T${r.turno || "—"} · ${r.avance || r.area} ha</small></div>
              <div><strong>${d.totalJarras} j</strong><small>${d.totalKg} kg</small></div>
            </button>`;
          })
          .join("")
      : `<p class="empty">Aún no hay lotes hoy.</p>`;
    $("#lot-list").querySelectorAll("[data-lote]").forEach((b) => {
      b.onclick = () => loadRecord(b.dataset.lote);
    });
  }

  function loadRecord(loteId) {
    const rec = APP.API.localRecords(today()).find((r) => String(r.lote) === String(loteId));
    const L = APP.Data.findLote(loteId);
    if (!L) return;
    state.lote = L;
    state.avance = String(rec?.avance || L.area);
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
  }

  function resetLoteInputs() {
    state.lote = null;
    state.avance = "";
    state.conv = false;
    state.china = false;
    state.jarrasConv = "";
    state.jarrasChina = "";
    $("#inp-avance").value = "";
    $("#inp-jconv").value = "";
    $("#inp-jchina").value = "";
    paintLote();
    paintChecks();
  }

  async function saveLote() {
    if (state.saving) return;
    if (!state.lote) {
      await feedback("Falta el lote", "Selecciona un lote.");
      return;
    }
    if (!state.conv && !state.china) {
      await feedback("Marca el tipo", "Convencional y/o China.");
      return;
    }
    const data = currentData();
    const d = APP.Data.derive(data);
    if (!(d.totalJarras > 0)) {
      await feedback("Faltan jarras", "Ingresa la cantidad.");
      return;
    }
    data.totalJarras = d.totalJarras;
    data.totalKg = d.totalKg;
    state.session.grupo = String(data.grupo);
    state.session.etapa = String(data.etapa);
    persistSession();
    state.saving = true;
    $("#qb-loader").hidden = false;
    try {
      await APP.API.submit({ type: "produccion", clientId: APP.API.newClientId(), data });
      resetLoteInputs();
      paintDay();
      await feedback("Lote guardado", `${data.lote} · ${d.totalJarras} jarras · ${d.totalKg} kg`);
    } catch (e) {
      await feedback("No se guardó", String(e.message || e));
    } finally {
      $("#qb-loader").hidden = true;
      state.saving = false;
    }
  }

  function bind() {
    $("#chip-fecha").textContent = fmtDate(today());

    $("#trig-supervisor").onclick = () => {
      APP.PreciseSelect.open({
        title: "Supervisor",
        getOptions: (q) => APP.Data.supervisorOptions(q),
        onSelect: (opt) => {
          state.session.supervisorDni = opt.dni || opt.id;
          state.session.supervisorNombre = opt.nombre || opt.label;
          persistSession();
          paintPeople();
        },
      });
    };
    $("#trig-scanner").onclick = () => {
      APP.PreciseSelect.open({
        title: "Escáner",
        getOptions: (q) => APP.Data.scannerOptions(q),
        onSelect: (opt) => {
          state.session.scannerDni = opt.dni || opt.id;
          state.session.scannerNombre = opt.nombre || opt.label;
          persistSession();
          paintPeople();
        },
      });
    };
    $("#trig-lote").onclick = () => {
      APP.PreciseSelect.open({
        title: "Lote",
        getOptions: (q) => APP.Data.loteOptions(q),
        onSelect: (opt) => {
          const L = opt.lote || APP.Data.findLote(opt.id);
          if (!L) return;
          state.lote = L;
          state.avance = String(L.area);
          state.session.grupo = String(L.grupo);
          state.session.etapa = String(L.etapa);
          persistSession();
          $("#inp-avance").value = state.avance;
          paintLote();
          paintTotals();
        },
      });
    };

    $("#inp-jornales").addEventListener("input", () => {
      state.session.jornales = Math.max(0, Number(String($("#inp-jornales").value).replace(/\D/g, "")) || 0);
      persistSession();
    });
    $("#inp-avance").addEventListener("input", () => {
      state.avance = String($("#inp-avance").value).replace(",", ".");
      paintTotals();
    });
    $("#inp-jconv").addEventListener("input", () => {
      state.jarrasConv = String($("#inp-jconv").value).replace(/\D/g, "");
      paintTotals();
    });
    $("#inp-jchina").addEventListener("input", () => {
      state.jarrasChina = String($("#inp-jchina").value).replace(/\D/g, "");
      paintTotals();
    });
    $("#chk-conv").addEventListener("change", () => {
      state.conv = $("#chk-conv").checked;
      if (!state.conv) state.jarrasConv = "";
      $("#inp-jconv").value = state.jarrasConv;
      paintChecks();
    });
    $("#chk-china").addEventListener("change", () => {
      state.china = $("#chk-china").checked;
      if (!state.china) state.jarrasChina = "";
      $("#inp-jchina").value = state.jarrasChina;
      paintChecks();
    });
    $("#btn-save").onclick = saveLote;
    $("#btn-export").onclick = () => {
      const records = APP.API.localRecords(today());
      APP.Excel.download({ session: state.session, fecha: today(), records });
      feedback("Excel listo", "Se descargó el reporte del día.");
    };

    window.addEventListener("online", paintOnline);
    window.addEventListener("offline", paintOnline);
    window.addEventListener("app:activity", paintDay);
  }

  function paintOnline() {
    const on = navigator.onLine;
    const el = $("#chip-online");
    el.classList.toggle("offline", !on);
    el.querySelector("span:last-child").textContent = on ? "En línea" : "Sin red";
  }

  function init() {
    loadSession();
    paintPeople();
    paintLote();
    paintChecks();
    paintDay();
    paintOnline();
    bind();
    if ("serviceWorker" in navigator) {
      try {
        navigator.serviceWorker.register("./sw.js");
      } catch (_) {}
    }
  }

  document.addEventListener("DOMContentLoaded", init);
  return { state };
})();
