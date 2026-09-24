var TZ = 'America/Lima';
var JARRA_KG = 1.14;
var LOG = '_lotes';
var RESP = 'Responsables';
var RESPONSABLE_FIJO = 'Luis Verde';
var GREEN = '#1B5E20';
var HEADERS = [
  'Fecha', 'Scanner', 'Scanner DNI', 'Supervisor', 'Supervisor DNI',
  'Grupo', 'Etapa', 'Lote', 'Fundo', 'Variedad', 'MD', 'Turno', 'Area', 'Avance',
  'Jarras Conv', 'Kg Conv', 'Jarras China', 'Kg China', 'Total Jarras', 'Total Kg', 'Jornales', 'ClientId', 'TurnoCampo', 'Hora envío', 'Hora registro'
];
var RESUMEN_HEADERS = [
  'Semana', 'Fecha', 'Tipo', 'Fundo', 'Módulo', 'Variedad', 'Turno',
  'Envase', 'Calibre', 'Responsable', 'Frecuencia',
  'Área', 'Jornales', 'Kilos', 'Kg/ha', 'Kg/Jornales', 'Hora registro'
];
var RESP_HEADERS = [
  'Fecha', 'Supervisor', 'Supervisor DNI', 'Turno campo', 'Jarras', 'Kilos', 'Hora registro'
];
var ACUM = 'Acumulado';
var ACUM_HEADERS = [
  'Fecha', 'Semana', 'Supervisor', 'Supervisor DNI', 'Turno campo',
  'Fundo', 'Variedad', 'Módulo', 'Lote', 'Turno', 'Área',
  'Jarras Conv', 'Kg Conv', 'Jarras China', 'Kg China',
  'Total Jarras', 'Total Kg', 'Jornales', 'Kg/ha', 'Kg/Jn',
  'ClientId', 'Hora registro'
];
var YEAR_REPORTE = 2026;

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  if (e && e.parameter && e.parameter.ping) return jsonOut_({ ok: true, pong: true });
  return jsonOut_({ ok: true });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var raw = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
    var body = JSON.parse(raw);
    var ss = SpreadsheetApp.getActive();
    var sh = logSheet_(ss);
    var recs = normalizeRecs_(body);
    var headers = ensureHeaders_(sh);
    var result = upsertBatch_(sh, recs, headers);
    // Suma solo este envío sobre lo que YA está en la hoja del día / Responsables.
    // No reconstruye desde _lotes (si borras la hoja, no vuelve el historial).
    if (body.rebuild !== false && result.changed && result.deltas && result.deltas.length) {
      try {
        mergeDaysFromDeltas_(ss, headers, result.deltas, sh);
        mergeResponsablesFromDeltas_(ss, headers, result.deltas);
        mergeReportesFromDeltas_(ss, headers, result.deltas);
        mergeAcumuladoFromDeltas_(ss, headers, result.deltas);
      } catch (rebuildErr) {}
    }
    return jsonOut_({
      ok: true,
      accepted: result.accepted || [],
      existing: result.existing || []
    });
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

function normalizeRecs_(body) {
  if (!body) return [];
  if (body.action === 'batchSave' && body.records) return body.records;
  if (Object.prototype.toString.call(body.records) === '[object Array]') return body.records;
  if (body.data) return [{ clientId: body.clientId || body.data.clientId, data: body.data }];
  return [body];
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Q Berries')
    .addItem('Actualizar resumen', 'rebuildResumen')
    .addItem('Preparar hojas', 'setupSheets')
    .addToUi();
}

function logSheet_(ss) {
  ss = ss || SpreadsheetApp.getActive();
  migrateProduccion_(ss);
  var sh = ss.getSheetByName(LOG);
  if (!sh) sh = ss.insertSheet(LOG);
  try { sh.hideSheet(); } catch (e) {}
  return sh;
}

function migrateProduccion_(ss) {
  var props = PropertiesService.getScriptProperties();
  if (props.getProperty('mig_prod_v1') === '1' && !ss.getSheetByName('Produccion')) return;
  var old = ss.getSheetByName('Produccion');
  if (!old) {
    props.setProperty('mig_prod_v1', '1');
    return;
  }
  var dest = ss.getSheetByName(LOG);
  if (!dest) dest = ss.insertSheet(LOG);
  if (old.getLastRow() > 0 && dest.getLastRow() < 2) {
    var rng = old.getDataRange();
    dest.getRange(1, 1, rng.getNumRows(), rng.getNumColumns()).setValues(rng.getValues());
  }
  if (ss.getSheets().length > 1) {
    try { ss.deleteSheet(old); } catch (e) {}
  }
  props.setProperty('mig_prod_v1', '1');
}

function setupSheets() {
  var ss = SpreadsheetApp.getActive();
  var sh = logSheet_(ss);
  ensureHeaders_(sh);
  ensureRespSheet_(ss);
  ensureAcumuladoSheet_(ss);
  rebuildResumen();
}

function ensureHeaders_(sh) {
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    return HEADERS.slice();
  }
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var have = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  if (lastCol >= HEADERS.length && String(have[0]) === 'Fecha') {
    var missing = false;
    var h;
    for (h = 0; h < HEADERS.length; h++) {
      if (have.indexOf(HEADERS[h]) === -1) {
        missing = true;
        break;
      }
    }
    if (!missing) return have;
  }
  HEADERS.forEach(function (name) {
    if (have.indexOf(name) === -1) {
      sh.getRange(1, have.length + 1).setValue(name);
      have.push(name);
    }
  });
  return have;
}

function col_(headers, name) {
  var i = headers.indexOf(name);
  return i < 0 ? -1 : i;
}

function cell_(row, headers, name) {
  var i = col_(headers, name);
  return i < 0 ? '' : row[i];
}

function normTurno_(v) {
  return /tarde/i.test(String(v || '')) ? 'Tarde' : 'Mañana';
}

function fallbackKey_(fecha, lote, turnoCampo, supervisorDni) {
  var f = toIsoFecha_(fecha);
  var l = String(lote || '').trim();
  var s = String(supervisorDni || '').trim();
  if (!f || !l || !s) return '';
  return f + '|' + l + '|' + normTurno_(turnoCampo) + '|' + s;
}

function fallbackKeyFromData_(d) {
  return fallbackKey_(d.fecha, d.lote, d.turnoCampo, d.supervisorDni);
}

function fallbackKeyFromRow_(row, headers) {
  return fallbackKey_(
    cell_(row, headers, 'Fecha'),
    cell_(row, headers, 'Lote'),
    cell_(row, headers, 'TurnoCampo') || 'Mañana',
    cell_(row, headers, 'Supervisor DNI')
  );
}

function buildRow_(headers, d, clientId) {
  var map = {
    'Fecha': toIsoFecha_(d.fecha) || d.fecha,
    'Scanner': d.scanner,
    'Scanner DNI': d.scannerDni,
    'Supervisor': String(d.supervisor || '').trim(),
    'Supervisor DNI': d.supervisorDni,
    'Grupo': d.grupo,
    'Etapa': d.etapa,
    'Lote': d.lote,
    'Fundo': d.fundo,
    'Variedad': d.variedad,
    'MD': d.md,
    'Turno': d.turno,
    'Area': d.area,
    'Avance': d.avance,
    'Jarras Conv': d.jarrasConv,
    'Kg Conv': d.kgConv,
    'Jarras China': d.jarrasChina,
    'Kg China': d.kgChina,
    'Total Jarras': d.totalJarras,
    'Total Kg': d.totalKg,
    'Jornales': d.jornales,
    'ClientId': clientId || '',
    'TurnoCampo': normTurno_(d.turnoCampo),
    'Hora registro': horaCorta_(d.horaRegistro || d.horaEnvio),
    'Hora envío': d.horaEnvio || ''
  };
  return headers.map(function (name) {
    return map.hasOwnProperty(name) ? map[name] : '';
  });
}

function indexLog_(data, headers) {
  var byId = {};
  var byKey = {};
  var i;
  for (i = 0; i < data.length; i++) {
    var id = String(cell_(data[i], headers, 'ClientId') || '').trim();
    if (id) byId[id] = i;
    var key = fallbackKeyFromRow_(data[i], headers);
    if (key) byKey[key] = i;
  }
  return { byId: byId, byKey: byKey };
}

function writeRanges_(sh, updates, appends) {
  if (updates.length) {
    updates.sort(function (a, b) { return a.row - b.row; });
    var i = 0;
    while (i < updates.length) {
      var start = i;
      var block = [updates[i].values];
      while (i + 1 < updates.length && updates[i + 1].row === updates[i].row + 1) {
        i++;
        block.push(updates[i].values);
      }
      sh.getRange(updates[start].row, 1, block.length, block[0].length).setValues(block);
      i++;
    }
  }
  if (appends.length) {
    sh.getRange(sh.getLastRow() + 1, 1, appends.length, appends[0].length).setValues(appends);
  }
}

/** Solo lee ClientId (+ claves cortas) para upsert rápido. No carga toda la hoja. */
function upsertBatch_(sh, recs, headers) {
  headers = headers || ensureHeaders_(sh);
  var last = sh.getLastRow();
  var idCol = col_(headers, 'ClientId');
  var byId = {};
  if (last > 1 && idCol >= 0) {
    var idVals = sh.getRange(2, idCol + 1, last - 1, 1).getValues();
    var i;
    for (i = 0; i < idVals.length; i++) {
      var cid = String(idVals[i][0] || '').trim();
      if (cid) byId[cid] = i + 2;
    }
  }

  var updates = [];
  var appends = [];
  var accepted = [];
  var existing = [];
  var fechas = {};
  var deltas = [];
  var iRec;

  for (iRec = 0; iRec < recs.length; iRec++) {
    var item = recs[iRec] || {};
    var d = item.data || {};
    var id = String(item.clientId || d.clientId || '').trim();
    var lote = String(d.lote || '').trim();
    if (!id || !lote) continue;
    d.fecha = toIsoFecha_(d.fecha) || String(d.fecha || '').trim();
    d.lote = lote;
    d.turnoCampo = normTurno_(d.turnoCampo);
    d.supervisorDni = String(d.supervisorDni || '').trim();
    var values = buildRow_(headers, d, id);
    if (byId.hasOwnProperty(id)) {
      var rowNum = byId[id];
      var oldVals = sh.getRange(rowNum, 1, 1, headers.length).getValues()[0];
      updates.push({ row: rowNum, values: values });
      deltas.push({ old: oldVals, neu: values });
      existing.push(id);
    } else {
      appends.push(values);
      deltas.push({ old: null, neu: values });
      byId[id] = -1;
      accepted.push(id);
    }
    if (d.fecha) fechas[d.fecha] = true;
  }

  writeRanges_(sh, updates, appends);

  return {
    accepted: accepted,
    existing: existing,
    changed: accepted.length + updates.length > 0,
    fechas: Object.keys(fechas),
    deltas: deltas
  };
}

/** Lee solo filas de las fechas pedidas (1 pasada). */
function readLogForFechas_(sh, headers, isoList) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  var want = {};
  (isoList || []).forEach(function (iso) {
    var f = toIsoFecha_(iso);
    if (f) want[f] = true;
  });
  if (!Object.keys(want).length) return [];
  var fechaCol = col_(headers, 'Fecha');
  var data = sh.getRange(2, 1, last - 1, headers.length).getValues();
  if (fechaCol < 0) return data;
  var out = [];
  var i;
  for (i = 0; i < data.length; i++) {
    var f = toIsoFecha_(data[i][fechaCol]);
    if (want[f]) out.push(data[i]);
  }
  return out;
}

function toIsoFecha_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) {
    return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  }
  var s = String(v || '').trim();
  var m = /^(\d{4})-(\d{2})-(\d{2})/.exec(s);
  if (m) return m[1] + '-' + m[2] + '-' + m[3];
  m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})/.exec(s);
  if (m) {
    var d = ('0' + m[1]).slice(-2);
    var mo = ('0' + m[2]).slice(-2);
    return m[3] + '-' + mo + '-' + d;
  }
  return '';
}

function isoWeek(iso) {
  iso = toIsoFecha_(iso);
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso);
  if (!m) return '';
  var d = new Date(Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3])));
  var day = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - day);
  var yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil(((d - yearStart) / 86400000 + 1) / 7);
}

function horaCorta_(v) {
  var s = String(v || '').trim();
  var m = /(\d{1,2}):(\d{2})/.exec(s);
  if (!m) return '';
  return ('0' + m[1]).slice(-2) + ':' + m[2];
}

function horaMin_(a, b) {
  a = horaCorta_(a);
  b = horaCorta_(b);
  if (!a) return b;
  if (!b) return a;
  return a < b ? a : b;
}

function horaLabel_(g) {
  if (g.horaManana && g.horaTarde) return g.horaManana + ' / ' + g.horaTarde;
  return g.horaTarde || g.horaManana || '';
}

function fmtDate(iso) {
  iso = toIsoFecha_(iso);
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso);
  return m ? m[3] + '/' + m[2] + '/' + m[1] : '';
}

function fundoLabel(fundo, etapa) {
  var e = String(etapa || '').trim();
  var f = String(fundo || 'LICAPA').trim();
  if (/licapa/i.test(e)) return e;
  if (e) return f + ' ' + e;
  return f;
}

function mdLabel_(md) {
  return String(md || '').replace(/^m/i, '').trim();
}

function turnoCampoOnly_(row, headers) {
  var raw = String(cell_(row, headers, 'TurnoCampo') || '').trim();
  return /tarde/i.test(raw) ? 'Tarde' : 'Mañana';
}

function turnoLote_(v) {
  return String(v || '').replace(/^T/i, '').trim();
}

function tiposDeFila_(row, headers) {
  var jConv = num(cell_(row, headers, 'Jarras Conv'));
  var jChina = num(cell_(row, headers, 'Jarras China'));
  var kgConv = num(cell_(row, headers, 'Kg Conv'));
  var kgChina = num(cell_(row, headers, 'Kg China'));
  var out = [];
  if (jConv > 0 || kgConv > 0) out.push({ tipo: 'CONVENCIONAL', envase: jConv, kilos: kgConv });
  if (jChina > 0 || kgChina > 0) out.push({ tipo: 'CHINA', envase: jChina, kilos: kgChina });
  if (!out.length) {
    var totJ = num(cell_(row, headers, 'Total Jarras'));
    var totK = num(cell_(row, headers, 'Total Kg'));
    if (totJ > 0 || totK > 0) out.push({ tipo: 'CONVENCIONAL', envase: totJ, kilos: totK });
  }
  return out;
}

function num(v) {
  var n = Number(v);
  return isFinite(n) ? n : 0;
}

function sheetNameDia_(iso) {
  return fmtDate(iso).replace(/\//g, '-');
}

function daySheet_(ss, iso) {
  var name = sheetNameDia_(iso);
  if (!name) return null;
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function dropOldResumen_(ss) {
  ss.getSheets().forEach(function (s) {
    if (ss.getSheets().length <= 1) return;
    var name = String(s.getName() || '').trim();
    if (/^resumen$/i.test(name) || /^produccion$/i.test(name)) ss.deleteSheet(s);
  });
}

function writeDaySheet_(sh, rows) {
  var cols = RESUMEN_HEADERS.length;
  var wasEmpty = sh.getLastRow() === 0;
  var totRow = (rows && rows.length ? rows.length : 0) + 2;
  var clearTo = Math.max(sh.getLastRow(), totRow);
  if (clearTo > 0) sh.getRange(1, 1, clearTo, cols).clearContent();

  sh.getRange(1, 1, 1, cols).setValues([RESUMEN_HEADERS]);
  var lastData = 1;
  if (rows && rows.length) {
    sh.getRange(2, 1, rows.length, cols).setValues(rows);
    lastData = 1 + rows.length;
  }
  var empty = ['', '', '', '', '', '', '', 0, '', '', '', 0, 0, 0, '', '', ''];
  if (rows && rows.length) {
    empty[7] = rows.reduce(function (a, r) { return a + num(r[7]); }, 0);
    empty[11] = Math.round(rows.reduce(function (a, r) { return a + num(r[11]); }, 0) * 100) / 100;
    empty[12] = rows.reduce(function (a, r) { return a + num(r[12]); }, 0);
    empty[13] = rows.reduce(function (a, r) { return a + num(r[13]); }, 0);
    empty[14] = empty[11] > 0 ? Math.round(empty[13] / empty[11]) : '';
    empty[15] = empty[12] > 0 ? Math.round(empty[13] / empty[12]) : '';
  }
  sh.getRange(totRow, 1, 1, cols).setValues([empty]);
  sh.getRange(totRow, 1).setValue('TOTAL');
  if (wasEmpty) styleResumen_(sh, lastData, totRow);
  else {
    sh.getRange(1, 1, 1, cols).setBackground('#F3F3F3').setFontWeight('bold');
    sh.getRange(totRow, 1, 1, cols).setBackground('#F3F3F3').setFontWeight('bold');
  }
}

function supervisorKey_(dni, nombre, nameToDni) {
  var id = String(dni || '').replace(/\D/g, '').slice(0, 8);
  var nom = String(nombre || '').replace(/\s+/g, ' ').trim().toUpperCase();
  if (id.length === 8) return id;
  if (nom && nameToDni && nameToDni[nom]) return nameToDni[nom];
  return nom;
}

function loadMananaAv_(sh, headers, isoList) {
  var map = {};
  var rows = readLogForFechas_(sh, headers, isoList);
  (rows || []).forEach(function (row) {
    if (turnoCampoOnly_(row, headers) !== 'Mañana') return;
    var f = toIsoFecha_(cell_(row, headers, 'Fecha'));
    var lote = String(cell_(row, headers, 'Lote') || '').trim();
    if (f && lote) map[f + '|' + lote] = num(cell_(row, headers, 'Avance') || cell_(row, headers, 'Area'));
  });
  return map;
}

function groupLogToDays_(headers, values, mananaAvExtra) {
  var groups = {};
  var order = [];
  var mananaAv = {};
  if (mananaAvExtra) {
    Object.keys(mananaAvExtra).forEach(function (k) { mananaAv[k] = mananaAvExtra[k]; });
  }
  (values || []).forEach(function (row) {
    if (turnoCampoOnly_(row, headers) !== 'Mañana') return;
    var f = toIsoFecha_(cell_(row, headers, 'Fecha'));
    var lote = String(cell_(row, headers, 'Lote') || '').trim();
    if (f && lote) mananaAv[f + '|' + lote] = num(cell_(row, headers, 'Avance') || cell_(row, headers, 'Area'));
  });
  (values || []).forEach(function (row) {
    var fecha = toIsoFecha_(cell_(row, headers, 'Fecha'));
    if (!fecha) return;
    var md = mdLabel_(cell_(row, headers, 'MD'));
    var turno = turnoLote_(cell_(row, headers, 'Turno'));
    var variedad = String(cell_(row, headers, 'Variedad') || '').trim();
    var turnoCampo = turnoCampoOnly_(row, headers);
    var fundo = fundoLabel(cell_(row, headers, 'Fundo'), cell_(row, headers, 'Etapa'));
    var area = num(cell_(row, headers, 'Avance') || cell_(row, headers, 'Area'));
    if (turnoCampo === 'Tarde') {
      var base = mananaAv[fecha + '|' + String(cell_(row, headers, 'Lote') || '').trim()] || 0;
      if (area >= base) area = Math.round((area - base) * 1000) / 1000;
    }
    var jornales = num(cell_(row, headers, 'Jornales'));
    var tipos = tiposDeFila_(row, headers);
    tipos.forEach(function (t, idx) {
      var key = [fecha, t.tipo, fundo, md, variedad, turno].join('|');
      if (!groups[key]) {
        groups[key] = {
          iso: fecha,
          semana: isoWeek(fecha),
          fecha: fmtDate(fecha),
          tipo: t.tipo,
          fundo: fundo,
          md: md,
          variedad: variedad,
          turno: turno,
          responsable: RESPONSABLE_FIJO,
          manana: false,
          tarde: false,
          horaManana: '',
          horaTarde: '',
          envase: 0,
          area: 0,
          kilos: 0,
          jornales: 0
        };
        order.push(key);
      }
      var g = groups[key];
      var hr = horaCorta_(cell_(row, headers, 'Hora registro') || cell_(row, headers, 'Hora envío'));
      if (turnoCampo === 'Tarde') {
        g.tarde = true;
        g.horaTarde = horaMin_(g.horaTarde, hr);
      } else {
        g.manana = true;
        g.horaManana = horaMin_(g.horaManana, hr);
      }
      g.envase += t.envase;
      g.kilos += t.kilos;
      if (idx === 0) {
        g.area += area;
        g.jornales = Math.max(num(g.jornales), jornales);
      }
    });
  });
  var byDay = {};
  order.forEach(function (key) {
    var g = groups[key];
    if (!byDay[g.iso]) byDay[g.iso] = [];
    var jornales = num(g.jornales);
    var kgHa = g.area > 0 ? Math.round(g.kilos / g.area) : '';
    var kgJn = jornales > 0 ? Math.round(g.kilos / jornales) : '';
    byDay[g.iso].push([
      g.semana, g.fecha, g.tipo, g.fundo, g.md, g.variedad, g.turno || '',
      g.envase, '', g.responsable || RESPONSABLE_FIJO, '',
      Math.round(g.area * 100) / 100, jornales,
      Math.round(g.kilos), kgHa, kgJn, horaLabel_(g)
    ]);
  });
  return byDay;
}

function writeByDay_(ss, byDay) {
  Object.keys(byDay || {}).sort().forEach(function (iso) {
    var sh = daySheet_(ss, iso);
    if (sh) writeDaySheet_(sh, byDay[iso]);
  });
}

function dayRowKey_(row) {
  return [
    String(row[2] || '').trim(),
    String(row[3] || '').trim(),
    String(row[4] || '').trim(),
    String(row[5] || '').trim(),
    String(row[6] || '').trim()
  ].join('|');
}

function readDaySheetRows_(sh) {
  var cols = RESUMEN_HEADERS.length;
  var last = sh.getLastRow();
  if (last < 3) return [];
  var data = sh.getRange(2, 1, last - 1, cols).getValues();
  return data.filter(function (r) {
    if (String(r[0] || '').toUpperCase() === 'TOTAL') return false;
    return String(r[2] || '') || String(r[4] || '') || num(r[7]) || num(r[13]);
  });
}

function mergeHoraLabel_(a, b) {
  a = String(a || '').trim();
  b = String(b || '').trim();
  if (!a) return b;
  if (!b) return a;
  if (a === b) return a;
  var partsA = a.split(/\s*\/\s*/);
  var partsB = b.split(/\s*\/\s*/);
  var man = horaMin_(partsA[0] || '', partsB[0] || '');
  var tar = horaMin_(partsA[1] || '', partsB[1] || '');
  if (partsA.length > 1 || partsB.length > 1) {
    if (man && tar) return man + ' / ' + tar;
    return tar || man || a;
  }
  return horaMin_(a, b);
}

function applyDayRows_(map, order, rows, sign) {
  (rows || []).forEach(function (r) {
    var k = dayRowKey_(r);
    if (!map[k]) {
      if (sign < 0) return;
      map[k] = r.slice(0, RESUMEN_HEADERS.length);
      order.push(k);
      return;
    }
    var cur = map[k];
    cur[7] = num(cur[7]) + sign * num(r[7]);
    cur[11] = Math.round((num(cur[11]) + sign * num(r[11])) * 1000) / 1000;
    cur[13] = Math.round(num(cur[13]) + sign * num(r[13]));
    if (sign > 0) {
      cur[12] = Math.max(num(cur[12]), num(r[12]));
      cur[16] = mergeHoraLabel_(cur[16], r[16]);
      if (!cur[9]) cur[9] = r[9] || RESPONSABLE_FIJO;
    }
    cur[14] = num(cur[11]) > 0 ? Math.round(num(cur[13]) / num(cur[11])) : '';
    cur[15] = num(cur[12]) > 0 ? Math.round(num(cur[13]) / num(cur[12])) : '';
    if (num(cur[7]) <= 0 && num(cur[13]) <= 0 && Math.abs(num(cur[11])) < 0.0005) {
      delete map[k];
    }
  });
}

/** Suma (o resta en reenvíos) solo el lote de este POST sobre la hoja del día. */
function mergeDaysFromDeltas_(ss, headers, deltas, logSh) {
  var oldRows = [];
  var newRows = [];
  var fechas = {};
  (deltas || []).forEach(function (d) {
    if (d.old) {
      oldRows.push(d.old);
      var fo = toIsoFecha_(cell_(d.old, headers, 'Fecha'));
      if (fo) fechas[fo] = true;
    }
    if (d.neu) {
      newRows.push(d.neu);
      var fn = toIsoFecha_(cell_(d.neu, headers, 'Fecha'));
      if (fn) fechas[fn] = true;
    }
  });
  var isoList = Object.keys(fechas);
  if (!isoList.length) return;
  var mananaAv = loadMananaAv_(logSh, headers, isoList);
  var bySub = groupLogToDays_(headers, oldRows, mananaAv);
  var byAdd = groupLogToDays_(headers, newRows, mananaAv);
  isoList.forEach(function (iso) {
    var sh = daySheet_(ss, iso);
    if (!sh) return;
    var map = {};
    var order = [];
    readDaySheetRows_(sh).forEach(function (r) {
      var k = dayRowKey_(r);
      if (!map.hasOwnProperty(k)) order.push(k);
      map[k] = r.slice(0, RESUMEN_HEADERS.length);
    });
    applyDayRows_(map, order, bySub[iso] || [], -1);
    applyDayRows_(map, order, byAdd[iso] || [], 1);
    var out = [];
    order.forEach(function (k) {
      if (map[k]) out.push(map[k]);
    });
    writeDaySheet_(sh, out);
  });
}

function respKeyFromLogRow_(row, headers) {
  var fecha = toIsoFecha_(cell_(row, headers, 'Fecha'));
  var nombre = String(cell_(row, headers, 'Supervisor') || '').replace(/\s+/g, ' ').trim();
  var dni = String(cell_(row, headers, 'Supervisor DNI') || '').replace(/\D/g, '').slice(0, 8);
  if (!fecha || (!nombre && !dni)) return '';
  return fecha + '|' + (dni.length === 8 ? dni : nombre.toUpperCase());
}

function contribResponsables_(headers, values) {
  var map = {};
  (values || []).forEach(function (row) {
    var key = respKeyFromLogRow_(row, headers);
    if (!key) return;
    var fecha = toIsoFecha_(cell_(row, headers, 'Fecha'));
    var nombre = String(cell_(row, headers, 'Supervisor') || '').replace(/\s+/g, ' ').trim();
    var dni = String(cell_(row, headers, 'Supervisor DNI') || '').replace(/\D/g, '').slice(0, 8);
    var turnoCampo = turnoCampoOnly_(row, headers);
    var jarras = num(cell_(row, headers, 'Total Jarras'));
    if (!(jarras > 0)) jarras = num(cell_(row, headers, 'Jarras Conv')) + num(cell_(row, headers, 'Jarras China'));
    var kilos = num(cell_(row, headers, 'Total Kg'));
    if (!(kilos > 0)) kilos = num(cell_(row, headers, 'Kg Conv')) + num(cell_(row, headers, 'Kg China'));
    var hr = horaCorta_(cell_(row, headers, 'Hora registro') || cell_(row, headers, 'Hora envío'));
    if (!map[key]) {
      map[key] = {
        iso: fecha,
        fecha: fmtDate(fecha),
        supervisor: nombre || dni,
        dni: dni,
        manana: false,
        tarde: false,
        jarras: 0,
        kilos: 0,
        horaManana: '',
        horaTarde: ''
      };
    }
    var g = map[key];
    if (nombre) g.supervisor = nombre;
    if (dni) g.dni = dni;
    g.jarras += jarras;
    g.kilos += kilos;
    if (turnoCampo === 'Tarde') {
      g.tarde = true;
      if (hr) g.horaTarde = hr;
    } else {
      g.manana = true;
      if (hr) g.horaManana = hr;
    }
  });
  return map;
}

function mergeResponsablesFromDeltas_(ss, headers, deltas) {
  var oldRows = [];
  var newRows = [];
  (deltas || []).forEach(function (d) {
    if (d.old) oldRows.push(d.old);
    if (d.neu) newRows.push(d.neu);
  });
  var sub = contribResponsables_(headers, oldRows);
  var add = contribResponsables_(headers, newRows);
  var sh = ensureRespSheet_(ss);
  var map = {};
  var order = [];
  if (sh.getLastRow() > 1) {
    var have = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
    var old = sh.getRange(2, 1, sh.getLastRow(), have.length).getValues();
    var fi = have.indexOf('Fecha');
    var si = have.indexOf('Supervisor');
    var di = have.indexOf('Supervisor DNI');
    var ti = have.indexOf('Turno campo');
    var ji = have.indexOf('Jarras');
    var ki = have.indexOf('Kilos');
    var hi = have.indexOf('Hora registro');
    old.forEach(function (row) {
      var fechaDisp = row[fi >= 0 ? fi : 0];
      var nombre = row[si >= 0 ? si : 1];
      var dni = row[di >= 0 ? di : 2];
      var iso = toIsoFecha_(fechaDisp);
      if (!iso) {
        var m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(String(fechaDisp || '').trim());
        if (m) iso = m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
      }
      var key = (iso || '') + '|' + (String(dni || '').replace(/\D/g, '').slice(0, 8).length === 8
        ? String(dni || '').replace(/\D/g, '').slice(0, 8)
        : String(nombre || '').replace(/\s+/g, ' ').trim().toUpperCase());
      if (!key || key === '|') return;
      var tc = String(row[ti >= 0 ? ti : 3] || '');
      map[key] = {
        iso: iso,
        fecha: fmtDate(iso) || String(fechaDisp || ''),
        supervisor: String(nombre || '').trim() || String(dni || ''),
        dni: String(dni || '').replace(/\D/g, '').slice(0, 8),
        manana: /mañana/i.test(tc),
        tarde: /tarde/i.test(tc),
        jarras: num(row[ji >= 0 ? ji : 4]),
        kilos: num(row[ki >= 0 ? ki : 5]),
        horaManana: '',
        horaTarde: ''
      };
      var hr = horaCorta_(row[hi >= 0 ? hi : 6]);
      if (map[key].tarde) map[key].horaTarde = hr;
      else map[key].horaManana = hr;
      order.push(key);
    });
  }
  function applyResp(contrib, sign) {
    Object.keys(contrib).forEach(function (key) {
      var c = contrib[key];
      if (!map[key]) {
        if (sign < 0) return;
        map[key] = {
          iso: c.iso,
          fecha: c.fecha,
          supervisor: c.supervisor,
          dni: c.dni,
          manana: false,
          tarde: false,
          jarras: 0,
          kilos: 0,
          horaManana: '',
          horaTarde: ''
        };
        order.push(key);
      }
      var g = map[key];
      if (c.supervisor) g.supervisor = c.supervisor;
      if (c.dni) g.dni = c.dni;
      g.jarras = num(g.jarras) + sign * num(c.jarras);
      g.kilos = num(g.kilos) + sign * num(c.kilos);
      if (sign > 0) {
        if (c.manana) {
          g.manana = true;
          if (c.horaManana) g.horaManana = c.horaManana;
        }
        if (c.tarde) {
          g.tarde = true;
          if (c.horaTarde) g.horaTarde = c.horaTarde;
        }
      }
      if (num(g.jarras) <= 0 && num(g.kilos) <= 0) delete map[key];
    });
  }
  applyResp(sub, -1);
  applyResp(add, 1);
  var rows = [];
  order.forEach(function (key) {
    var g = map[key];
    if (!g) return;
    var turnoCampo = g.manana && g.tarde ? 'Mañana / Tarde' : g.tarde ? 'Tarde' : 'Mañana';
    var hora = g.tarde && g.horaTarde ? g.horaTarde : g.horaManana || g.horaTarde || '';
    rows.push([
      g.fecha,
      g.supervisor,
      g.dni,
      turnoCampo,
      Math.round(g.jarras),
      Math.round(g.kilos * 100) / 100,
      hora
    ]);
  });
  rows.sort(function (a, b) {
    var c = String(a[0]).localeCompare(String(b[0]));
    return c || String(a[1]).localeCompare(String(b[1]), 'es');
  });
  writeResponsablesSheet_(sh, rows);
}

function rebuildDays_(ss, headers, values, isoList) {
  var want = {};
  (isoList || []).forEach(function (iso) {
    var f = toIsoFecha_(iso);
    if (f) want[f] = true;
  });
  if (!Object.keys(want).length) return;
  var filtered = [];
  (values || []).forEach(function (row) {
    var f = toIsoFecha_(cell_(row, headers, 'Fecha'));
    if (want[f]) filtered.push(row);
  });
  var byDay = groupLogToDays_(headers, filtered);
  Object.keys(want).forEach(function (iso) {
    var sh = daySheet_(ss, iso);
    if (sh) writeDaySheet_(sh, byDay[iso] || []);
  });
}

function ensureRespSheet_(ss) {
  ss = ss || SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(RESP);
  if (!sh) sh = ss.insertSheet(RESP);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, RESP_HEADERS.length).setValues([RESP_HEADERS]);
  } else {
    var lastCol = Math.max(sh.getLastColumn(), 1);
    var have = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    RESP_HEADERS.forEach(function (name) {
      if (have.indexOf(name) === -1) {
        sh.getRange(1, have.length + 1).setValue(name);
        have.push(name);
      }
    });
  }
  return sh;
}

function groupResponsables_(headers, values) {
  var groups = {};
  var order = [];
  (values || []).forEach(function (row) {
    var fecha = toIsoFecha_(cell_(row, headers, 'Fecha'));
    if (!fecha) return;
    var nombre = String(cell_(row, headers, 'Supervisor') || '').replace(/\s+/g, ' ').trim();
    var dni = String(cell_(row, headers, 'Supervisor DNI') || '').replace(/\D/g, '').slice(0, 8);
    if (!nombre && !dni) return;
    var key = fecha + '|' + (dni || nombre.toUpperCase());
    var turnoCampo = turnoCampoOnly_(row, headers);
    var jarras = num(cell_(row, headers, 'Total Jarras'));
    if (!(jarras > 0)) {
      jarras = num(cell_(row, headers, 'Jarras Conv')) + num(cell_(row, headers, 'Jarras China'));
    }
    var kilos = num(cell_(row, headers, 'Total Kg'));
    if (!(kilos > 0)) {
      kilos = num(cell_(row, headers, 'Kg Conv')) + num(cell_(row, headers, 'Kg China'));
    }
    var hr = horaCorta_(cell_(row, headers, 'Hora registro') || cell_(row, headers, 'Hora envío'));
    if (!groups[key]) {
      groups[key] = {
        iso: fecha,
        fecha: fmtDate(fecha),
        supervisor: nombre || dni,
        dni: dni,
        manana: false,
        tarde: false,
        jarras: 0,
        kilos: 0,
        horaManana: '',
        horaTarde: ''
      };
      order.push(key);
    }
    var g = groups[key];
    if (nombre) g.supervisor = nombre;
    if (dni) g.dni = dni;
    g.jarras += jarras;
    g.kilos += kilos;
    if (turnoCampo === 'Tarde') {
      g.tarde = true;
      if (hr) g.horaTarde = hr;
    } else {
      g.manana = true;
      if (hr) g.horaManana = hr;
    }
  });
  var rows = [];
  order.forEach(function (key) {
    var g = groups[key];
    var turnoCampo = g.manana && g.tarde ? 'Mañana / Tarde' : g.tarde ? 'Tarde' : 'Mañana';
    var hora = g.tarde && g.horaTarde ? g.horaTarde : g.horaManana || g.horaTarde || '';
    rows.push([
      g.fecha,
      g.supervisor,
      g.dni,
      turnoCampo,
      Math.round(g.jarras),
      Math.round(g.kilos * 100) / 100,
      hora
    ]);
  });
  rows.sort(function (a, b) {
    var c = String(a[0]).localeCompare(String(b[0]));
    return c || String(a[1]).localeCompare(String(b[1]), 'es');
  });
  return rows;
}

function writeResponsablesSheet_(sh, rows) {
  var cols = RESP_HEADERS.length;
  var wasEmpty = sh.getLastRow() === 0;
  var last = 1 + (rows ? rows.length : 0);
  var clearTo = Math.max(sh.getLastRow(), last);
  if (clearTo > 0) sh.getRange(1, 1, clearTo, cols).clearContent();
  sh.getRange(1, 1, 1, cols).setValues([RESP_HEADERS]);
  if (rows && rows.length) {
    sh.getRange(2, 1, rows.length, cols).setValues(rows);
  }
  if (wasEmpty) {
    var all = sh.getRange(1, 1, Math.max(last, 1), cols);
    all.setFontFamily('Calibri')
      .setFontSize(10)
      .setFontColor('#000000')
      .setBackground('#FFFFFF')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle')
      .setBorder(true, true, true, true, true, true, '#B7B7B7', SpreadsheetApp.BorderStyle.SOLID);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, cols).setBackground('#F3F3F3').setFontWeight('bold');
    var widths = [92, 220, 100, 110, 70, 70, 86];
    widths.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
    sh.setTabColor('#F7941D');
  } else {
    sh.getRange(1, 1, 1, cols).setBackground('#F3F3F3').setFontWeight('bold');
  }
  if (rows && rows.length) {
    sh.getRange(2, 5, rows.length, 1).setNumberFormat('#,##0');
    sh.getRange(2, 6, rows.length, 1).setNumberFormat('0.00');
  }
}

function rebuildResponsables_(ss, headers, values, isoList) {
  ss = ss || SpreadsheetApp.getActive();
  var sh = ensureRespSheet_(ss);
  var want = null;
  if (isoList && isoList.length) {
    want = {};
    isoList.forEach(function (iso) {
      var f = toIsoFecha_(iso);
      if (f) want[f] = true;
    });
  }
  var existing = [];
  if (want && sh.getLastRow() > 1) {
    var have = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
    var old = sh.getRange(2, 1, sh.getLastRow() - 1, have.length).getValues();
    var fechaIdx = have.indexOf('Fecha');
    old.forEach(function (row) {
      var f = toIsoFecha_(row[fechaIdx >= 0 ? fechaIdx : 0]);
      if (!f || want[f]) return;
      existing.push(row.slice(0, RESP_HEADERS.length));
    });
  }
  var filtered = values || [];
  if (want) {
    filtered = [];
    (values || []).forEach(function (row) {
      var f = toIsoFecha_(cell_(row, headers, 'Fecha'));
      if (want[f]) filtered.push(row);
    });
  }
  var rebuilt = groupResponsables_(headers, filtered);
  var allRows = existing.concat(rebuilt);
  allRows.sort(function (a, b) {
    var c = String(a[0]).localeCompare(String(b[0]));
    return c || String(a[1]).localeCompare(String(b[1]), 'es');
  });
  writeResponsablesSheet_(sh, allRows);
}

function rebuildResumen() {
  // Ya no reconstruye desde _lotes: si borraste la hoja del día, no se revive el historial.
  // El resumen solo crece con cada Enviar (suma en el mismo módulo).
  var ss = SpreadsheetApp.getActive();
  dropOldResumen_(ss);
  ensureRespSheet_(ss);
  ensureAcumuladoSheet_(ss);
}

function styleResumen_(sh, lastData, totRow) {
  var cols = RESUMEN_HEADERS.length;
  sh.setTabColor('#1B5E20');
  var all = sh.getRange(1, 1, totRow, cols);
  all.setFontFamily('Calibri')
    .setFontSize(10)
    .setFontColor('#000000')
    .setBackground('#FFFFFF')
    .setFontWeight('normal')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setNumberFormat('@')
    .setBorder(true, true, true, true, true, true, '#B7B7B7', SpreadsheetApp.BorderStyle.SOLID);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, cols)
    .setBackground('#F3F3F3')
    .setFontColor('#000000')
    .setFontWeight('bold')
    .setWrap(true);
  if (lastData > 1) {
    sh.getRange(2, 1, lastData - 1, 1).setNumberFormat('0');
    sh.getRange(2, 2, lastData - 1, 1).setNumberFormat('@');
    sh.getRange(2, 7, lastData - 1, 1).setNumberFormat('@');
    sh.getRange(2, 8, lastData - 1, 1).setNumberFormat('#,##0');
    sh.getRange(2, 10, lastData - 1, 1).setNumberFormat('@').setWrap(true);
    sh.getRange(2, 12, lastData - 1, 1).setNumberFormat('0.00');
    sh.getRange(2, 13, lastData - 1, 2).setNumberFormat('#,##0');
    sh.getRange(2, 15, lastData - 1, 2).setNumberFormat('#,##0');
    sh.getRange(2, 17, lastData - 1, 1).setNumberFormat('@');
  }
  sh.getRange(totRow, 1, 1, cols)
    .setBackground('#F3F3F3')
    .setFontColor('#000000')
    .setFontWeight('bold');
  sh.getRange(totRow, 8).setNumberFormat('#,##0');
  sh.getRange(totRow, 12).setNumberFormat('0.00');
  sh.getRange(totRow, 13, 1, 2).setNumberFormat('#,##0');
  sh.getRange(totRow, 15, 1, 2).setNumberFormat('#,##0');
  var widths = [70, 92, 110, 78, 70, 96, 70, 72, 70, 168, 86, 70, 90, 70, 68, 92, 86];
  widths.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  sh.setRowHeight(1, 32);
  try {
    var f = sh.getFilter();
    if (f) f.remove();
  } catch (e1) {}
  if (lastData > 1) sh.getRange(1, 1, lastData, cols).createFilter();
}

/* ─── Reporte por supervisor+fecha (formato exportación) + Acumulado ─── */

function placeAfterSheet_(ss, sh, afterName) {
  if (!ss || !sh) return;
  var after = ss.getSheetByName(afterName);
  if (!after) return;
  try {
    var target = after.getIndex() + 1;
    if (sh.getIndex() !== target) {
      ss.setActiveSheet(sh);
      ss.moveActiveSheet(target);
    }
  } catch (e) {}
}

function placeAfterResponsables_(ss, sh) {
  placeAfterSheet_(ss, sh, RESP);
}

function sanitizeSheetPart_(s, maxLen) {
  var t = String(s || '')
    .replace(/[\\\/\?\*\[\\:]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  if (maxLen && t.length > maxLen) t = t.slice(0, maxLen).trim();
  return t || 'Supervisor';
}

function reporteSheetName_(iso, supervisor, dni) {
  var fecha = sheetNameDia_(iso) || 'fecha';
  var nom = sanitizeSheetPart_(supervisor || dni, 36);
  var name = 'R ' + fecha + ' ' + nom;
  if (name.length > 90) name = name.slice(0, 90).trim();
  return name;
}

function reporteKeyFromRow_(row, headers) {
  var iso = toIsoFecha_(cell_(row, headers, 'Fecha'));
  var dni = String(cell_(row, headers, 'Supervisor DNI') || '').replace(/\D/g, '').slice(0, 8);
  var nom = String(cell_(row, headers, 'Supervisor') || '').replace(/\s+/g, ' ').trim();
  if (!iso) return '';
  return iso + '|' + (dni.length === 8 ? dni : nom.toUpperCase());
}

function loteKeyFromRow_(row, headers) {
  return String(cell_(row, headers, 'Lote') || '').trim() + '|' + turnoCampoOnly_(row, headers);
}

function areaFromLogRow_(row, headers) {
  var av = num(cell_(row, headers, 'Avance'));
  if (av > 0) return av;
  return num(cell_(row, headers, 'Area'));
}

function reporteDataRowFromLog_(row, headers) {
  var area = areaFromLogRow_(row, headers);
  var jConv = num(cell_(row, headers, 'Jarras Conv'));
  var kgConv = num(cell_(row, headers, 'Kg Conv'));
  var jChina = num(cell_(row, headers, 'Jarras China'));
  var kgChina = num(cell_(row, headers, 'Kg China'));
  var totJ = num(cell_(row, headers, 'Total Jarras'));
  var totK = num(cell_(row, headers, 'Total Kg'));
  if (!(totJ > 0)) totJ = jConv + jChina;
  if (!(totK > 0)) totK = kgConv + kgChina;
  var jornales = num(cell_(row, headers, 'Jornales'));
  var kgHa = area > 0 ? Math.round(totK / area) : '';
  var kgJn = jornales > 0 ? Math.round(totK / jornales) : '';
  return {
    key: loteKeyFromRow_(row, headers),
    jornal: jornales > 0 ? jornales : '',
    area: Math.round(area * 1000) / 1000,
    lote: String(cell_(row, headers, 'Lote') || '').trim(),
    turno: turnoLote_(cell_(row, headers, 'Turno')),
    modulo: mdLabel_(cell_(row, headers, 'MD')),
    jarrasConv: jConv,
    kgConv: Math.round(kgConv * 10) / 10,
    jarrasChina: jChina,
    kgChina: Math.round(kgChina * 10) / 10,
    totalJarras: Math.round(totJ),
    totalKg: Math.round(totK * 10) / 10,
    kgHa: kgHa,
    kgJn: kgJn,
    fundo: fundoLabel(cell_(row, headers, 'Fundo'), cell_(row, headers, 'Etapa')),
    variedad: String(cell_(row, headers, 'Variedad') || '').trim(),
    turnoCampo: turnoCampoOnly_(row, headers),
    clientId: String(cell_(row, headers, 'ClientId') || '').trim()
  };
}

function ensureReporteSheet_(ss, iso, supervisor, dni) {
  var name = reporteSheetName_(iso, supervisor, dni);
  var sh = ss.getSheetByName(name);
  if (!sh) {
    ensureAcumuladoSheet_(ss);
    sh = ss.insertSheet(name);
    placeAfterSheet_(ss, sh, ACUM);
  }
  return sh;
}

function readReporteLotes_(sh) {
  var map = {};
  var order = [];
  if (!sh || sh.getLastRow() < 5) return { map: map, order: order };
  var last = sh.getLastRow();
  if (last < 6) return { map: map, order: order }; // solo cabecera + total vacío
  // Filas de datos: 5 .. last-1 (última = TOTAL). Col 14 = Turno campo (interno).
  var data = sh.getRange(5, 1, last - 1, 14).getValues();
  var i;
  for (i = 0; i < data.length; i++) {
    var r = data[i];
    if (String(r[0] || '').toUpperCase() === 'TOTAL') continue;
    var lote = String(r[2] || '').trim();
    if (!lote) continue;
    var tc = /tarde/i.test(String(r[13] || '')) ? 'Tarde' : 'Mañana';
    var key = lote + '|' + tc;
    if (map[key]) continue;
    map[key] = {
      key: key,
      jornal: r[0],
      area: num(r[1]),
      lote: lote,
      turno: String(r[3] || '').trim(),
      modulo: String(r[4] || '').trim(),
      jarrasConv: num(r[5]),
      kgConv: num(r[6]),
      jarrasChina: num(r[7]),
      kgChina: num(r[8]),
      totalJarras: num(r[9]),
      totalKg: num(r[10]),
      kgHa: r[11],
      kgJn: r[12],
      turnoCampo: tc
    };
    order.push(key);
  }
  return { map: map, order: order };
}

function writeReporteSheet_(sh, meta, rows) {
  var NAVY = '#0B3A66';
  var PEACH = '#F6D0B0';
  var BROWN = '#C47848';
  var LGREEN = '#C8E6A0';
  var LBLUE = '#B7D4F5';
  var WHITE = '#FFFFFF';
  var cols = 14;
  var n = rows && rows.length ? rows.length : 0;
  var totRow = 5 + n;
  var clearTo = Math.max(sh.getLastRow(), totRow);
  if (clearTo > 0) sh.getRange(1, 1, clearTo, cols).clearContent().clearFormat();

  sh.getRange(1, 1, 1, 13).merge().setValue(
    'REPORTE DIARIO DE COSECHA EXPORTACIÓN - Q BERRIES ' + YEAR_REPORTE
  );
  sh.getRange(1, 1)
    .setBackground(NAVY)
    .setFontColor(WHITE)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  sh.setRowHeight(1, 28);

  sh.getRange(2, 1).setValue('SUPERVISOR');
  sh.getRange(2, 2, 1, 4).merge().setValue(meta.supervisor || '');
  sh.getRange(2, 6).setValue('SEMANA');
  sh.getRange(2, 7).setValue(meta.semana || '');
  sh.getRange(2, 9).setValue('FECHA');
  sh.getRange(2, 10, 1, 2).merge().setValue(meta.fecha || '');
  sh.getRange(2, 1, 1, 13)
    .setBackground('#1F6B3A')
    .setFontColor(WHITE)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sh.getRange(2, 7).setBackground('#F4B183');
  sh.getRange(2, 10).setBackground('#FFE699').setFontColor('#C00000');

  sh.getRange(3, 1, 1, 5).merge().setValue('');
  sh.getRange(3, 6, 1, 2).merge().setValue('CONVENCIONAL');
  sh.getRange(3, 8, 1, 2).merge().setValue('CHINA');
  sh.getRange(3, 10, 1, 2).merge().setValue('TOTAL');
  sh.getRange(3, 1, 1, 5).setBackground(NAVY).setFontColor(WHITE);
  sh.getRange(3, 6, 1, 2).setBackground(PEACH).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(3, 8, 1, 2).setBackground(LGREEN).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(3, 10, 1, 2).setBackground(LBLUE).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(3, 12, 1, 2).setBackground(NAVY);

  var heads = ['JORNAL', 'AREA', 'LOTE', 'TURNO', 'MODULO', 'JARRAS', 'KG', 'JARRAS', 'KG', 'JARRAS', 'KG', 'KG/HA', 'KG/JN', 'TC'];
  sh.getRange(4, 1, 1, cols).setValues([heads]);
  sh.getRange(4, 1, 1, 5).setBackground(NAVY).setFontColor(WHITE).setFontWeight('bold');
  sh.getRange(4, 6).setBackground(PEACH).setFontWeight('bold');
  sh.getRange(4, 7).setBackground(BROWN).setFontColor(WHITE).setFontWeight('bold');
  sh.getRange(4, 8, 1, 2).setBackground(LGREEN).setFontWeight('bold');
  sh.getRange(4, 10, 1, 2).setBackground(LBLUE).setFontWeight('bold');
  sh.getRange(4, 12, 1, 2).setBackground(NAVY).setFontColor(WHITE).setFontWeight('bold');
  sh.getRange(4, 14).setBackground('#EEEEEE').setFontColor('#888888');
  sh.getRange(4, 1, 1, cols).setHorizontalAlignment('center');

  if (n) {
    var body = rows.map(function (r) {
      return [
        r.jornal === '' || r.jornal == null ? '' : r.jornal,
        r.area,
        r.lote,
        r.turno,
        r.modulo,
        r.jarrasConv,
        r.kgConv,
        r.jarrasChina,
        r.kgChina,
        r.totalJarras,
        r.totalKg,
        r.kgHa,
        r.kgJn,
        r.turnoCampo || 'Mañana'
      ];
    });
    sh.getRange(5, 1, n, cols).setValues(body);
    sh.getRange(5, 1, n, 13)
      .setBackground(PEACH)
      .setHorizontalAlignment('center')
      .setBorder(true, true, true, true, true, true, '#8FA3B5', SpreadsheetApp.BorderStyle.SOLID);
    sh.getRange(5, 2, n, 1).setNumberFormat('0.00');
    sh.getRange(5, 6, n, 1).setNumberFormat('#,##0');
    sh.getRange(5, 7, n, 1).setNumberFormat('0.0');
    sh.getRange(5, 8, n, 1).setNumberFormat('#,##0');
    sh.getRange(5, 9, n, 1).setNumberFormat('0.0');
    sh.getRange(5, 10, n, 1).setNumberFormat('#,##0');
    sh.getRange(5, 11, n, 1).setNumberFormat('0.0');
  }

  var sumArea = 0, sumJC = 0, sumKC = 0, sumJH = 0, sumKH = 0, sumTJ = 0, sumTK = 0, sumJn = 0;
  (rows || []).forEach(function (r) {
    sumArea += num(r.area);
    sumJC += num(r.jarrasConv);
    sumKC += num(r.kgConv);
    sumJH += num(r.jarrasChina);
    sumKH += num(r.kgChina);
    sumTJ += num(r.totalJarras);
    sumTK += num(r.totalKg);
    sumJn = Math.max(sumJn, num(r.jornal));
  });
  var totKgHa = sumArea > 0 ? Math.round(sumTK / sumArea) : '';
  var totKgJn = sumJn > 0 ? Math.round(sumTK / sumJn) : '';
  sh.getRange(totRow, 1, 1, cols).setValues([[
    'TOTAL', Math.round(sumArea * 1000) / 1000, '', '', '',
    sumJC, Math.round(sumKC * 10) / 10, sumJH, Math.round(sumKH * 10) / 10,
    sumTJ, Math.round(sumTK * 10) / 10, totKgHa, totKgJn, ''
  ]]);
  sh.getRange(totRow, 1, 1, 13)
    .setBackground(NAVY)
    .setFontColor(WHITE)
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  var widths = [70, 70, 70, 60, 70, 70, 70, 70, 70, 70, 70, 70, 70, 1];
  widths.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
  try { sh.hideColumns(14); } catch (eHide) {}
  sh.setTabColor('#0B3A66');
  sh.setFrozenRows(4);
}

function applyReporteDelta_(map, order, item, sign) {
  if (!item || !item.lote) return;
  var key = String(item.lote) + '|' + (item.turnoCampo === 'Tarde' ? 'Tarde' : 'Mañana');
  if (!map[key]) {
    if (sign < 0) return;
    map[key] = {
      key: key,
      jornal: item.jornal,
      area: 0,
      lote: item.lote,
      turno: item.turno,
      modulo: item.modulo,
      jarrasConv: 0,
      kgConv: 0,
      jarrasChina: 0,
      kgChina: 0,
      totalJarras: 0,
      totalKg: 0,
      kgHa: '',
      kgJn: '',
      turnoCampo: item.turnoCampo === 'Tarde' ? 'Tarde' : 'Mañana'
    };
    order.push(key);
  }
  var cur = map[key];
  if (sign > 0) {
    if (item.turno) cur.turno = item.turno;
    if (item.modulo) cur.modulo = item.modulo;
    if (item.jornal !== '' && item.jornal != null) cur.jornal = item.jornal;
    cur.turnoCampo = item.turnoCampo === 'Tarde' ? 'Tarde' : 'Mañana';
  }
  cur.area = Math.round((num(cur.area) + sign * num(item.area)) * 1000) / 1000;
  cur.jarrasConv = num(cur.jarrasConv) + sign * num(item.jarrasConv);
  cur.kgConv = Math.round((num(cur.kgConv) + sign * num(item.kgConv)) * 10) / 10;
  cur.jarrasChina = num(cur.jarrasChina) + sign * num(item.jarrasChina);
  cur.kgChina = Math.round((num(cur.kgChina) + sign * num(item.kgChina)) * 10) / 10;
  cur.totalJarras = num(cur.totalJarras) + sign * num(item.totalJarras);
  cur.totalKg = Math.round((num(cur.totalKg) + sign * num(item.totalKg)) * 10) / 10;
  cur.kgHa = num(cur.area) > 0 ? Math.round(num(cur.totalKg) / num(cur.area)) : '';
  cur.kgJn = num(cur.jornal) > 0 ? Math.round(num(cur.totalKg) / num(cur.jornal)) : '';
  if (num(cur.totalJarras) <= 0 && num(cur.totalKg) <= 0 && Math.abs(num(cur.area)) < 0.0005) {
    delete map[key];
  }
}

function mergeReportesFromDeltas_(ss, headers, deltas) {
  var groups = {};
  (deltas || []).forEach(function (d) {
    ['old', 'neu'].forEach(function (side) {
      var row = d[side];
      if (!row) return;
      var rk = reporteKeyFromRow_(row, headers);
      if (!rk) return;
      if (!groups[rk]) {
        groups[rk] = {
          iso: toIsoFecha_(cell_(row, headers, 'Fecha')),
          supervisor: String(cell_(row, headers, 'Supervisor') || '').trim(),
          dni: String(cell_(row, headers, 'Supervisor DNI') || '').replace(/\D/g, '').slice(0, 8),
          oldItems: [],
          newItems: []
        };
      }
      var g = groups[rk];
      var nom = String(cell_(row, headers, 'Supervisor') || '').trim();
      if (nom) g.supervisor = nom;
      var item = reporteDataRowFromLog_(row, headers);
      if (side === 'old') g.oldItems.push(item);
      else g.newItems.push(item);
    });
  });

  Object.keys(groups).forEach(function (rk) {
    var g = groups[rk];
    if (!g.iso) return;
    var sh = ensureReporteSheet_(ss, g.iso, g.supervisor, g.dni);
    var cur = readReporteLotes_(sh);
    var map = cur.map;
    var order = cur.order;
    g.oldItems.forEach(function (it) { applyReporteDelta_(map, order, it, -1); });
    g.newItems.forEach(function (it) { applyReporteDelta_(map, order, it, 1); });
    var rows = [];
    order.forEach(function (k) {
      if (map[k]) rows.push(map[k]);
    });
    rows.sort(function (a, b) {
      return String(a.lote).localeCompare(String(b.lote), 'es', { numeric: true });
    });
    writeReporteSheet_(sh, {
      supervisor: g.supervisor || g.dni,
      fecha: fmtDate(g.iso),
      semana: isoWeek(g.iso)
    }, rows);
  });
}

function ensureAcumuladoSheet_(ss) {
  ss = ss || SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(ACUM);
  if (!sh) {
    sh = ss.insertSheet(ACUM);
    placeAfterResponsables_(ss, sh);
  }
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, ACUM_HEADERS.length).setValues([ACUM_HEADERS]);
  } else {
    var lastCol = Math.max(sh.getLastColumn(), 1);
    var have = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    ACUM_HEADERS.forEach(function (name) {
      if (have.indexOf(name) === -1) {
        sh.getRange(1, have.length + 1).setValue(name);
        have.push(name);
      }
    });
  }
  sh.getRange(1, 1, 1, ACUM_HEADERS.length)
    .setBackground('#0B3A66')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sh.setTabColor('#1565C0');
  sh.setFrozenRows(1);
  placeAfterResponsables_(ss, sh);
  return sh;
}

function acumuladoRowFromLog_(row, headers) {
  var iso = toIsoFecha_(cell_(row, headers, 'Fecha'));
  var area = areaFromLogRow_(row, headers);
  var jConv = num(cell_(row, headers, 'Jarras Conv'));
  var kgConv = num(cell_(row, headers, 'Kg Conv'));
  var jChina = num(cell_(row, headers, 'Jarras China'));
  var kgChina = num(cell_(row, headers, 'Kg China'));
  var totJ = num(cell_(row, headers, 'Total Jarras'));
  var totK = num(cell_(row, headers, 'Total Kg'));
  if (!(totJ > 0)) totJ = jConv + jChina;
  if (!(totK > 0)) totK = kgConv + kgChina;
  var jornales = num(cell_(row, headers, 'Jornales'));
  return [
    fmtDate(iso),
    isoWeek(iso),
    String(cell_(row, headers, 'Supervisor') || '').trim(),
    String(cell_(row, headers, 'Supervisor DNI') || '').replace(/\D/g, '').slice(0, 8),
    turnoCampoOnly_(row, headers),
    fundoLabel(cell_(row, headers, 'Fundo'), cell_(row, headers, 'Etapa')),
    String(cell_(row, headers, 'Variedad') || '').trim(),
    mdLabel_(cell_(row, headers, 'MD')),
    String(cell_(row, headers, 'Lote') || '').trim(),
    turnoLote_(cell_(row, headers, 'Turno')),
    Math.round(area * 1000) / 1000,
    jConv,
    Math.round(kgConv * 10) / 10,
    jChina,
    Math.round(kgChina * 10) / 10,
    Math.round(totJ),
    Math.round(totK * 10) / 10,
    jornales > 0 ? jornales : '',
    area > 0 ? Math.round(totK / area) : '',
    jornales > 0 ? Math.round(totK / jornales) : '',
    String(cell_(row, headers, 'ClientId') || '').trim(),
    horaCorta_(cell_(row, headers, 'Hora registro') || cell_(row, headers, 'Hora envío'))
  ];
}

function mergeAcumuladoFromDeltas_(ss, headers, deltas) {
  var sh = ensureAcumuladoSheet_(ss);
  var idCol = ACUM_HEADERS.indexOf('ClientId');
  var byId = {};
  var last = sh.getLastRow();
  if (last > 1 && idCol >= 0) {
    var ids = sh.getRange(2, idCol + 1, last, 1).getValues();
    var i;
    for (i = 0; i < ids.length; i++) {
      var cid = String(ids[i][0] || '').trim();
      if (cid) byId[cid] = i + 2;
    }
  }

  var updates = [];
  var appends = [];
  var seen = {};
  (deltas || []).forEach(function (d) {
    var neu = d.neu;
    if (!neu) return;
    var cid = String(cell_(neu, headers, 'ClientId') || '').trim();
    if (!cid || seen[cid]) return;
    seen[cid] = true;
    var values = acumuladoRowFromLog_(neu, headers);
    if (byId.hasOwnProperty(cid)) {
      updates.push({ row: byId[cid], values: values });
    } else {
      appends.push(values);
      byId[cid] = -1;
    }
  });

  if (updates.length) {
    updates.sort(function (a, b) { return a.row - b.row; });
    var u = 0;
    while (u < updates.length) {
      var start = u;
      var block = [updates[u].values];
      while (u + 1 < updates.length && updates[u + 1].row === updates[u].row + 1) {
        u++;
        block.push(updates[u].values);
      }
      sh.getRange(updates[start].row, 1, block.length, ACUM_HEADERS.length).setValues(block);
      u++;
    }
  }
  if (appends.length) {
    sh.getRange(sh.getLastRow() + 1, 1, appends.length, ACUM_HEADERS.length).setValues(appends);
  }
}
