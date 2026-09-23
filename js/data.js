/** Catálogos: lotes, escáneres, supervisores */
window.APP = window.APP || {};

APP.Data = (() => {
  const scanners = {
    "48055477": { nombre: "ZAVALETA GONZALES LUSBET ERELID" },
    "18078464": { nombre: "LOYOLA DOMINGUEZ OSWALDO ELMER" },
    "41520670": { nombre: "VASQUEZ SANCHEZ JOSE ISIDRO" },
    "41809717": { nombre: "DIAZ SALAZAR OLGA ELIZABETH" },
    "44932896": { nombre: "PAREDES MANTILLA GINA IVONNE" },
    "77799828": { nombre: "ZAVALETA IGLESIAS KAHORY MARIANELA" },
    "40354659": { nombre: "REYES JAVE LOURDES LIZETH" },
    "44262821": { nombre: "CHAVEZ CABRERA MARIBEL VICENTA" },
    "45123552": { nombre: "LINARES CERNA OSCAR PAUL" },
    "46819781": { nombre: "PEREZ LEON SALLY ELIZABETH" },
    "47188311": { nombre: "PEREIRA VASQUEZ LUCELIA LEONORA" },
    "47407697": { nombre: "ROLDAN PEREZ JULIO ANTONIO" },
  };

  const supervisors = {
    "48533707": { nombre: "JULCA GAMBOA DILMER ELICER" },
    "18851808": { nombre: "AGUIRRE NORIEGA MARCO ANTONIO" },
    "41501158": { nombre: "CHILON SANCHEZ DANNY ROBERTH" },
    "42493820": { nombre: "VASQUEZ DELGADO ROBERTO CARLOS" },
    "42992833": { nombre: "PONCE RUIZ ISIDRO" },
    "43558894": { nombre: "DIAZ VARAS FANNY DEL MILAGRO" },
    "43583858": { nombre: "PLASENCIA CORREA NADIA YVONNE" },
    "44141396": { nombre: "VASQUEZ URBINA EDIN CLAY" },
  };

  const lotes = [
    { lote: "198", md: "10", turno: "5", fundo: "LICAPA", variedad: "SEKOYA", etapa: "1", area: 1.5, grupo: "1" },
    { lote: "199", md: "10", turno: "5", fundo: "LICAPA", variedad: "SEKOYA", etapa: "1", area: 1.6, grupo: "1" },
    { lote: "200", md: "10", turno: "5", fundo: "LICAPA", variedad: "SEKOYA", etapa: "1", area: 1.8, grupo: "1" },
    { lote: "201", md: "10", turno: "5", fundo: "LICAPA", variedad: "SEKOYA", etapa: "1", area: 1.4, grupo: "1" },
    { lote: "202", md: "10", turno: "7", fundo: "LICAPA", variedad: "SEKOYA", etapa: "1", area: 1.7, grupo: "1" },
    { lote: "206", md: "10", turno: "7", fundo: "LICAPA", variedad: "SEKOYA", etapa: "1", area: 1.9, grupo: "1" },
    { lote: "207", md: "10", turno: "7", fundo: "LICAPA", variedad: "SEKOYA", etapa: "1", area: 2.0, grupo: "1" },
    { lote: "209", md: "10", turno: "7", fundo: "LICAPA", variedad: "SEKOYA", etapa: "1", area: 1.94, grupo: "2" },
    { lote: "210", md: "10", turno: "7", fundo: "LICAPA", variedad: "SEKOYA", etapa: "1", area: 0.7, grupo: "2" },
    { lote: "211", md: "10", turno: "7", fundo: "LICAPA", variedad: "SEKOYA", etapa: "1", area: 2.28, grupo: "2" },
    { lote: "212", md: "10", turno: "7", fundo: "LICAPA", variedad: "SEKOYA", etapa: "1", area: 1.94, grupo: "2" },
    { lote: "213", md: "10", turno: "8", fundo: "LICAPA", variedad: "MAGICA", etapa: "1", area: 1.8, grupo: "3" },
    { lote: "214", md: "10", turno: "8", fundo: "LICAPA", variedad: "MAGICA", etapa: "1", area: 1.55, grupo: "3" },
    { lote: "215", md: "10", turno: "8", fundo: "LICAPA", variedad: "MAGICA", etapa: "1", area: 1.62, grupo: "3" },
    { lote: "216", md: "10", turno: "8", fundo: "LICAPA", variedad: "MAGICA", etapa: "1", area: 1.7, grupo: "3" },
    { lote: "223", md: "10", turno: "10", fundo: "LICAPA", variedad: "MAGICA", etapa: "2", area: 1.4, grupo: "4" },
    { lote: "224", md: "10", turno: "10", fundo: "LICAPA", variedad: "MAGICA", etapa: "2", area: 1.5, grupo: "4" },
  ];

  const lotesById = {};
  lotes.forEach((l) => {
    lotesById[String(l.lote)] = l;
  });

  function shortName(nombre) {
    const parts = String(nombre || "").trim().split(/\s+/).filter(Boolean);
    if (parts.length <= 2) return parts.join(" ");
    return `${parts[2] || ""} ${parts[0] || ""}`.trim();
  }

  function displayName(nombre) {
    return shortName(nombre) || String(nombre || "").trim();
  }

  function personOptions(map, query) {
    const q = String(query || "").trim().toLowerCase();
    const digits = String(query || "").replace(/\D/g, "");
    const out = [];
    Object.keys(map).forEach((dni) => {
      const nombre = map[dni].nombre || "";
      const short = displayName(nombre);
      if (q) {
        const byDni = digits.length >= 2 && dni.includes(digits);
        const byName = nombre.toLowerCase().includes(q) || short.toLowerCase().includes(q);
        if (!byDni && !byName) return;
      }
      out.push({ id: dni, label: short, meta: dni, dni, nombre });
    });
    out.sort((a, b) => a.label.localeCompare(b.label, "es"));
    return out;
  }

  function findPerson(map, raw) {
    const v = String(raw || "").trim();
    if (map[v]) return { dni: v, ...map[v] };
    const dni = v.replace(/\D/g, "").slice(0, 8);
    if (dni.length === 8 && map[dni]) return { dni, ...map[dni] };
    return null;
  }

  function jarraKg() {
    return Number((APP.CONFIG && APP.CONFIG.JARRA_KG) || 1.16);
  }

  function round2(n) {
    return Math.round((Number(n) || 0) * 100) / 100;
  }

  function kgEffective(jarras, kgInput) {
    const kg = Number(kgInput);
    if (Number.isFinite(kg) && kg > 0) return round2(kg);
    return round2((Number(jarras) || 0) * jarraKg());
  }

  function derive(data) {
    const jConv = Number(data.jarrasConv) || 0;
    const jChina = Number(data.jarrasChina) || 0;
    const kgConv = Number(data.kgConv) || 0;
    const kgChina = Number(data.kgChina) || 0;
    const area = Number(data.avance || data.area) || 0;
    const jornales = Number(data.jornales) || 0;
    const kgConvEff = kgEffective(jConv, kgConv);
    const kgChinaEff = kgEffective(jChina, kgChina);
    const totalJarras = jConv + jChina;
    const totalKg = round2(kgConvEff + kgChinaEff);
    return {
      jarrasConv: jConv,
      jarrasChina: jChina,
      kgConv,
      kgChina,
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

  return {
    lotes,
    findLote: (id) => lotesById[String(id || "").trim()] || null,
    loteOptions: (q) => {
      const s = String(q || "").trim().toLowerCase();
      return lotes
        .filter((l) => !s || String(l.lote).includes(s) || String(l.md).includes(s) || l.variedad.toLowerCase().includes(s))
        .map((l) => ({
          id: String(l.lote),
          label: String(l.lote),
          meta: `${l.fundo} · MD ${l.md} · T${l.turno} · ${l.area} ha`,
          lote: l,
        }));
    },
    scannerOptions: (q) => personOptions(scanners, q),
    supervisorOptions: (q) => personOptions(supervisors, q),
    findScanner: (id) => findPerson(scanners, id),
    findSupervisor: (id) => findPerson(supervisors, id),
    displayName,
    derive,
    kgEffective,
    round2,
    isoWeek,
    JARRA_KG: jarraKg,
  };
})();
