/** Catálogos: lotes y personas desde /data (offline en SW) */
window.APP = window.APP || {};

APP.Data = (() => {
  let scanners = {
    "48055477": { nombre: "ZAVALETA GONZALES LUSBET ERELID", cargo: "SCANER" },
    "18078464": { nombre: "LOYOLA DOMINGUEZ OSWALDO ELMER", cargo: "SCANER" },
    "41520670": { nombre: "VASQUEZ SANCHEZ JOSE ISIDRO", cargo: "SCANER" },
    "41809717": { nombre: "DIAZ SALAZAR OLGA ELIZABETH", cargo: "SCANER" },
    "44932896": { nombre: "PAREDES MANTILLA GINA IVONNE", cargo: "SCANER" },
    "77799828": { nombre: "ZAVALETA IGLESIAS KAHORY MARIANELA", cargo: "SCANER" },
    "40354659": { nombre: "REYES JAVE LOURDES LIZETH", cargo: "SCANER" },
    "44262821": { nombre: "CHAVEZ CABRERA MARIBEL VICENTA", cargo: "SCANER" },
    "45123552": { nombre: "LINARES CERNA OSCAR PAUL", cargo: "SCANER" },
    "46819781": { nombre: "PEREZ LEON SALLY ELIZABETH", cargo: "SCANER" },
    "47188311": { nombre: "PEREIRA VASQUEZ LUCELIA LEONORA", cargo: "SCANER" },
    "47407697": { nombre: "ROLDAN PEREZ JULIO ANTONIO", cargo: "SCANER" },
  };

  let supervisors = {
    "48268173": { nombre: "NORIEGA PONTE MICELY" },
    "77534125": { nombre: "HERRERA ALBERCA PAMELA" },
    "70192702": { nombre: "PURIZAGA SAAVEDRA PIERRE OSNAR" },
    "75141739": { nombre: "ROJAS AREDO YERSI YEN" },
    "48446147": { nombre: "HILARIO ÁVALOS EVELYN" },
    "48533707": { nombre: "JULCA GAMBOA DILMER ELICER" },
    "70507014": { nombre: "RODRÍGUEZ CABRERA KENYI JENNY" },
    "77146080": { nombre: "CHACÓN BERMÚDEZ NADIA SARAHÍ" },
    "45372928": { nombre: "PEÑA ROJAS LAURA PATRICIA" },
    "71806261": { nombre: "LEÓN TRIGOSO JHONY ANDRÓNICO" },
    "70656198": { nombre: "VARGAS DÍAZ ALDRIN" },
    "18851808": { nombre: "AGUIRRE NORIEGA MARCO ANTONIO" },
    "41501158": { nombre: "CHILON SANCHEZ DANNY ROBERTH" },
    "42493820": { nombre: "VASQUEZ DELGADO ROBERTO CARLOS" },
    "42992833": { nombre: "PONCE RUIZ ISIDRO" },
    "43558894": { nombre: "DIAZ VARAS FANNY DEL MILAGRO" },
    "43583858": { nombre: "PLASENCIA CORREA NADIA YVONNE" },
    "44141396": { nombre: "VASQUEZ URBINA EDIN CLAY" },
  };

  let lotes = [];
  let lotesById = {};
  let ready = false;
  let loading = null;
  const EXTRA_KEY = "qb-scanners-extra";

  function displayName(nombre) {
    return String(nombre || "").trim();
  }

  let scannerIndex = null;
  let supervisorIndex = null;

  function fullName(dni, fallback) {
    const id = String(dni || "").trim();
    const hit = (id && supervisors[id]) || (id && scanners[id]);
    return displayName((hit && hit.nombre) || fallback);
  }

  function rebuildPersonIndex(map) {
    const list = [];
    const keys = Object.keys(map || {});
    for (let i = 0; i < keys.length; i++) {
      const dni = keys[i];
      const row = map[dni] || {};
      const nombre = String(row.nombre || "");
      const cargo = String(row.cargo || row.puesto || "").toUpperCase();
      list.push({
        dni,
        nombre,
        cargo,
        label: displayName(nombre),
        nombreLower: nombre.toLowerCase(),
        isScaner: cargo.indexOf("SCAN") === 0,
      });
    }
    list.sort((a, b) => {
      if (a.isScaner !== b.isScaner) return a.isScaner ? -1 : 1;
      return a.label.localeCompare(b.label, "es");
    });
    return list;
  }

  function ensureScannerIndex() {
    if (!scannerIndex) scannerIndex = rebuildPersonIndex(scanners);
    return scannerIndex;
  }

  function ensureSupervisorIndex() {
    if (!supervisorIndex) supervisorIndex = rebuildPersonIndex(supervisors);
    return supervisorIndex;
  }

  function personOptionsFromIndex(index, query, limits) {
    const minList = limits && limits.minList != null ? limits.minList : 40;
    const maxList = limits && limits.limit != null ? limits.limit : 60;
    const preferScaner = !!(limits && limits.preferScaner);
    const raw = String(query || "").trim();
    const q = raw.toLowerCase();
    const digits = raw.replace(/\D/g, "");
    const out = [];

    if (!q) {
      for (let i = 0; i < index.length && out.length < minList; i++) {
        const row = index[i];
        if (preferScaner && !row.isScaner && out.length >= Math.min(minList, 28)) break;
        out.push({ id: row.dni, label: row.label, meta: row.dni, dni: row.dni, nombre: row.nombre });
      }
      if (preferScaner && out.length < minList) {
        for (let i = 0; i < index.length && out.length < minList; i++) {
          const row = index[i];
          if (row.isScaner) continue;
          out.push({ id: row.dni, label: row.label, meta: row.dni, dni: row.dni, nombre: row.nombre });
        }
      }
      return out;
    }

    // DNI exacto primero (rápido)
    if (digits.length === 8) {
      for (let i = 0; i < index.length; i++) {
        if (index[i].dni === digits) {
          const row = index[i];
          return [{ id: row.dni, label: row.label, meta: row.dni, dni: row.dni, nombre: row.nombre }];
        }
      }
    }

    const onlyDigits = digits.length > 0 && digits === raw.replace(/\s/g, "");
    for (let i = 0; i < index.length; i++) {
      const row = index[i];
      let ok = false;
      if (digits.length >= 1 && row.dni.indexOf(digits) !== -1) ok = true;
      else if (!onlyDigits && (row.nombreLower.indexOf(q) !== -1 || row.label.toLowerCase().indexOf(q) !== -1)) ok = true;
      if (!ok) continue;
      out.push({ id: row.dni, label: row.label, meta: row.dni, dni: row.dni, nombre: row.nombre });
      if (out.length >= maxList) break;
    }
    return out;
  }

  function personOptions(map, query, limits) {
    const preferScaner = !!(limits && limits.preferScaner);
    const index = preferScaner || map === scanners ? ensureScannerIndex() : map === supervisors ? ensureSupervisorIndex() : rebuildPersonIndex(map);
    return personOptionsFromIndex(index, query, limits);
  }

  function findPerson(map, raw) {
    const v = String(raw || "").trim();
    if (map[v]) return { dni: v, ...map[v] };
    const dni = v.replace(/\D/g, "").slice(0, 8);
    if (dni.length === 8 && map[dni]) return { dni, ...map[dni] };
    return null;
  }

  function mdFrom(modulo) {
    const m = String(modulo || "").trim();
    const n = m.replace(/^M/i, "");
    return n || m;
  }

  function normalizeLote(raw) {
    const lote = String(raw.lote ?? "").trim();
    const modulo = String(raw.modulo || "").trim();
    const areaRaw = raw.area != null ? raw.area : raw.ha;
    const areaNum = Number(areaRaw);
    return {
      lote,
      md: mdFrom(modulo),
      modulo,
      turno: String(raw.turno || "").trim(),
      etapa: String(raw.etapa || "").trim(),
      variedad: String(raw.variedad || "").trim(),
      fundo: String(raw.fundo || "LICAPA").trim(),
      codLote: String(raw.codLote || "").trim(),
      area: Number.isFinite(areaNum) && areaNum > 0 ? areaNum : "",
      grupo: raw.grupo != null && raw.grupo !== "" ? String(raw.grupo) : "",
    };
  }

  function etapaKind(etapa) {
    const e = String(etapa || "").trim().toUpperCase();
    if (e.includes("II")) return "II";
    if (e.includes("I") || e.includes("LICAPA")) return "I";
    return "";
  }

  function buildLoteCatalog(raw) {
    const all = Array.isArray(raw) ? raw : [];
    const groups = new Map();
    all.forEach((l) => {
      const n = String(l.lote ?? "").trim();
      if (!n) return;
      if (!groups.has(n)) groups.set(n, []);
      groups.get(n).push(l);
    });
    lotes = [];
    lotesById = {};
    groups.forEach((group, num) => {
      const licapaI = group.find((l) => etapaKind(l.etapa) === "I");
      const licapaII = group.find((l) => etapaKind(l.etapa) === "II");
      let pick;
      let etapaLabel;
      if (licapaI && licapaII) {
        pick = licapaI;
        etapaLabel = "Licapa I/II";
      } else if (licapaI) {
        pick = licapaI;
        etapaLabel = "Licapa I";
      } else if (licapaII) {
        pick = licapaII;
        etapaLabel = "Licapa II";
      } else {
        pick = group[0];
        etapaLabel = String(pick.etapa || "").trim();
      }
      const entry = normalizeLote({ ...pick, lote: num, etapa: etapaLabel });
      lotes.push(entry);
      lotesById[num] = entry;
      lotesById[`Q${num}`] = entry;
      if (entry.codLote) lotesById[entry.codLote] = entry;
    });
    lotes.sort((a, b) => {
      const na = parseInt(String(a.lote), 10);
      const nb = parseInt(String(b.lote), 10);
      const fa = Number.isFinite(na) ? na : 99999;
      const fb = Number.isFinite(nb) ? nb : 99999;
      return fa !== fb ? fa - fb : String(a.lote).localeCompare(String(b.lote), "es", { numeric: true });
    });
  }

  function findLote(id) {
    const key = String(id || "").trim();
    return lotesById[key] || lotesById[key.replace(/^Q/i, "")] || null;
  }

  function loteOptions(query) {
    const s = String(query || "").trim().toLowerCase();
    const out = [];
    const limit = s ? 80 : 50;
    for (let i = 0; i < lotes.length; i++) {
      const l = lotes[i];
      const loteStr = String(l.lote).toLowerCase();
      if (s) {
        const hay = `${l.lote} ${l.md} ${l.modulo} ${l.turno} ${l.variedad} ${l.etapa} ${l.codLote}`.toLowerCase();
        const prefix = loteStr === s || loteStr.startsWith(s);
        if (!prefix && !hay.includes(s)) continue;
      }
      out.push({
        id: String(l.lote),
        label: String(l.lote),
        meta: [l.fundo, l.modulo || (l.md ? "MD " + l.md : ""), l.turno ? "T" + l.turno : "", l.variedad, l.area !== "" ? l.area + " ha" : ""]
          .filter(Boolean)
          .join(" · "),
        lote: l,
        _rank: s && loteStr === s ? 0 : s && loteStr.startsWith(s) ? 1 : 2,
      });
    }
    if (s) out.sort((a, b) => a._rank - b._rank || parseInt(a.id, 10) - parseInt(b.id, 10));
    return out.slice(0, limit).map(({ _rank, ...opt }) => opt);
  }

  function tunelToTurno(tunel) {
    return String(tunel || "").replace(/^T/i, "").replace(/M$/i, "").trim();
  }

  function applyMapData(map) {
    if (!map || !Array.isArray(map.lots) || !map.lots.length) return;
    const modules = {};
    (map.modules || []).forEach((m) => {
      modules[String(m.id)] = m;
    });
    const byLote = {};
    map.lots.forEach((lot) => {
      const key = String(lot.lote ?? "").trim();
      if (!key) return;
      const prev = byLote[key];
      const ha = Number(lot.ha);
      if (!prev || (Number.isFinite(ha) && ha > (Number(prev.ha) || 0))) byLote[key] = lot;
    });
    lotes.forEach((l) => {
      const hit = byLote[String(l.lote)];
      if (!hit) return;
      const ha = Number(hit.ha);
      if (Number.isFinite(ha) && ha > 0) l.area = ha;
      const turno = tunelToTurno(hit.tunel);
      if (turno) l.turno = turno;
      if (hit.modulo != null && hit.modulo !== "") {
        const mid = String(hit.modulo);
        const mod = modules[mid];
        l.md = mid;
        l.modulo = (mod && mod.label) || "M" + mid;
      }
    });
    Object.keys(byLote).forEach((key) => {
      if (lotesById[key]) return;
      const hit = byLote[key];
      const mid = hit.modulo != null ? String(hit.modulo) : "";
      const mod = modules[mid];
      const entry = normalizeLote({
        lote: key,
        modulo: (mod && mod.label) || (mid ? "M" + mid : ""),
        turno: tunelToTurno(hit.tunel),
        etapa: "Licapa I",
        variedad: "",
        fundo: "LICAPA",
        ha: hit.ha,
      });
      lotes.push(entry);
      lotesById[key] = entry;
    });
    lotes.sort((a, b) => {
      const na = parseInt(String(a.lote), 10);
      const nb = parseInt(String(b.lote), 10);
      const fa = Number.isFinite(na) ? na : 99999;
      const fb = Number.isFinite(nb) ? nb : 99999;
      return fa !== fb ? fa - fb : String(a.lote).localeCompare(String(b.lote), "es", { numeric: true });
    });
  }

  function applyEmbeddedCatalogs() {
    if (Array.isArray(window.APP_LOTES) && window.APP_LOTES.length) {
      if (!lotes.length) buildLoteCatalog(window.APP_LOTES);
    }
    const sup = window.APP_SUPERVISORES;
    if (sup && sup.byDni && Object.keys(sup.byDni).length) {
      supervisors = { ...supervisors, ...sup.byDni };
      supervisorIndex = null;
    }
    if (window.APP_PLANO && Array.isArray(window.APP_PLANO.lots)) {
      applyMapData(window.APP_PLANO);
    }
    if (window.APP_TRABAJADORES && window.APP_TRABAJADORES.byDni) {
      applyTrabajadores(window.APP_TRABAJADORES);
    }
    if (lotes.length) ready = true;
    return lotes.length > 0;
  }

  function applyTrabajadores(json) {
    const byDni = json && json.byDni;
    if (!byDni || !Object.keys(byDni).length) return;
    scanners = { ...scanners, ...byDni };
    scannerIndex = null;
    loadExtras();
  }

  function cleanDni(raw) {
    return String(raw || "").replace(/\D/g, "").slice(0, 8);
  }

  function cleanNombre(raw) {
    return String(raw || "")
      .replace(/[^A-Za-zÁÉÍÓÚÜÑáéíóúüñ\s]/g, "")
      .toUpperCase()
      .replace(/\s+/g, " ")
      .trim();
  }

  function loadExtras() {
    try {
      const raw = JSON.parse(localStorage.getItem(EXTRA_KEY) || "{}");
      if (!raw || typeof raw !== "object") return;
      Object.keys(raw).forEach((dni) => {
        const id = cleanDni(dni);
        const nombre = cleanNombre((raw[dni] && raw[dni].nombre) || raw[dni]);
        if (id.length === 8 && nombre) scanners[id] = { nombre, cargo: "SCANER" };
      });
      scannerIndex = null;
    } catch (e) {}
  }

  function addScanner(dni, nombre) {
    const id = cleanDni(dni);
    const name = cleanNombre(nombre);
    if (id.length !== 8 || name.length < 3) return null;
    scanners[id] = { nombre: name, cargo: "SCANER" };
    scannerIndex = null;
    try {
      const raw = JSON.parse(localStorage.getItem(EXTRA_KEY) || "{}");
      raw[id] = { nombre: name };
      localStorage.setItem(EXTRA_KEY, JSON.stringify(raw));
    } catch (e) {}
    return { dni: id, nombre: name };
  }

  function fetchJson(url) {
    return fetch(url, { cache: "no-cache" })
      .then((r) => {
        if (!r.ok) throw new Error(String(r.status));
        const ct = (r.headers.get("content-type") || "").toLowerCase();
        if (ct.includes("text/html")) throw new Error("html");
        return r.json();
      })
      .catch(() => null);
  }

  function load(opts) {
    const force = !!(opts && opts.force);
    applyEmbeddedCatalogs();
    if (!force && ready && Object.keys(supervisors).length > 20 && lotes.length > 0 && Object.keys(scanners).length > 20) {
      return Promise.resolve(true);
    }
    if (!force && loading) return loading;
    loading = Promise.all([
      fetchJson("./data/lotes-licapa.json"),
      fetchJson("./data/supervisores-cosecha.json"),
      fetchJson("./data/plano-cosecha-etapa-i.json"),
    ])
      .then(([lotesJson, supJson, mapJson]) => {
        if (Array.isArray(lotesJson) && lotesJson.length) buildLoteCatalog(lotesJson);
        else applyEmbeddedCatalogs();
        if (supJson && supJson.byDni && Object.keys(supJson.byDni).length) {
          supervisors = { ...supervisors, ...supJson.byDni };
        } else {
          applyEmbeddedCatalogs();
        }
        const map = mapJson && Array.isArray(mapJson.lots) && mapJson.lots.length
          ? mapJson
          : window.APP_PLANO;
        if (map && Array.isArray(map.lots)) applyMapData(map);
        if (!lotes.length) applyEmbeddedCatalogs();
        ready = lotes.length > 0;
        window.dispatchEvent(new Event("app:catalogs"));
        return fetchJson("./data/trabajadores.json").then((trabJson) => {
          if (trabJson) applyTrabajadores(trabJson);
          if (Object.keys(scanners).length < 20) applyEmbeddedCatalogs();
          if (!lotes.length) applyEmbeddedCatalogs();
          ready = lotes.length > 0;
          loading = null;
          window.dispatchEvent(new Event("app:catalogs"));
          return lotes.length > 0;
        });
      })
      .catch(() => {
        applyEmbeddedCatalogs();
        ready = lotes.length > 0;
        loading = null;
        window.dispatchEvent(new Event("app:catalogs"));
        return lotes.length > 0;
      });
    return loading;
  }

  function jarraKg() {
    const n = Number(APP.CONFIG && APP.CONFIG.JARRA_KG);
    return Number.isFinite(n) && n > 0 ? n : 1.14;
  }

  function round2(n) {
    return Math.round((Number(n) || 0) * 100) / 100;
  }

  function kgEffective(jarras) {
    return round2((Number(jarras) || 0) * jarraKg());
  }

  function derive(data) {
    const jConv = Number(data.jarrasConv) || 0;
    const jChina = Number(data.jarrasChina) || 0;
    const area = Number(data.avance || data.area) || 0;
    const jornales = Number(data.jornales) || 0;
    const kgConvEff = kgEffective(jConv);
    const kgChinaEff = kgEffective(jChina);
    const totalJarras = jConv + jChina;
    const totalKg = round2(kgConvEff + kgChinaEff);
    return {
      jarrasConv: jConv,
      jarrasChina: jChina,
      kgConv: kgConvEff,
      kgChina: kgChinaEff,
      kgConvEff,
      kgChinaEff,
      totalJarras,
      totalKg,
      area,
      jornales,
      kgHa: area > 0 ? Math.round(totalKg / area) : 0,
      kgJn: jornales > 0 ? Math.round(totalKg / jornales) : 0,
    };
  }

  function isoWeek(iso) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(iso || ""));
    if (!m) return "";
    const d = new Date(Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3])));
    const day = d.getUTCDay() || 7;
    d.setUTCDate(d.getUTCDate() + 4 - day);
    const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
    return String(Math.ceil(((d - yearStart) / 86400000 + 1) / 7));
  }

  applyEmbeddedCatalogs();
  loadExtras();

  return {
    get lotes() {
      return lotes;
    },
    load,
    findLote,
    loteOptions,
    scannerOptions: (q) => personOptions(scanners, q, { minList: 40, limit: 60, preferScaner: true }),
    supervisorOptions: (q) => personOptions(supervisors, q, { minList: 40, limit: 60 }),
    findScanner: (id) => findPerson(scanners, id),
    findSupervisor: (id) => findPerson(supervisors, id),
    addScanner,
    cleanDni,
    cleanNombre,
    displayName,
    fullName,
    derive,
    kgEffective,
    round2,
    isoWeek,
    JARRA_KG: jarraKg,
  };
})();

APP.Data.load();
