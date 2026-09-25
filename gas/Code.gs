var TZ = 'America/Lima';
var JARRA_KG = 1.14;
var LOG = '_lotes';
var RESP = 'Responsables';
var RESPONSABLE_FIJO = 'Luis Verde';
var GREEN = '#1B5E20';
var HEADERS = [
  'Fecha', 'Scanner', 'Scanner DNI', 'Supervisor', 'Supervisor DNI',
  'Grupo', 'Etapa', 'Lote', 'Fundo', 'Variedad', 'MD', 'Turno', 'Area', 'Avance',
  'Jarras Conv', 'Kg Conv', 'Jarras China', 'Kg China', 'Total Jarras', 'Total Kg', 'Jornales',
  'TurnoCampo', 'Hora envío', 'Hora registro'
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
  'Total Jarras', 'Total Kg', 'Jornales', 'Kg/ha', 'Kg/Jn', 'Hora registro'
];
var MERGE_PROP_KEY = 'qb_merge_queue';
var MERGE_TRIGGER_FN = 'runQueuedMerge_';

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  if (e && e.parameter && e.parameter.ping) return jsonOut_({ ok: true, pong: true });
  if (e && e.parameter && e.parameter.test === '1') return jsonOut_(runSaveSelfTest_());
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

    // 1) Guardar SIEMPRE el POST (esto es lo que importa).
    var result = upsertBatch_(sh, recs, headers);
    SpreadsheetApp.flush();

    // 2) Responder ok YA → el celular confirma y no se congela.
    var response = {
      ok: true,
      accepted: result.accepted || [],
      existing: result.existing || [],
      saved: (result.accepted || []).length + (result.existing || []).length,
      flushed: true
    };

    // 3) Armar hojas visibles DESPUÉS (cola durable, no CacheService).
    if (result.changed && result.deltas && result.deltas.length) {
      queueMergeDeltas_(result.deltas);
    }

    return jsonOut_(response);
  } catch (err) {
    return jsonOut_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

/** Cola durable en PropertiesService (no es cache): el POST ya quedó en _lotes. */
function queueMergeDeltas_(deltas) {
  try {
    var props = PropertiesService.getScriptProperties();
    var arr = [];
    var prev = props.getProperty(MERGE_PROP_KEY);
    if (prev) {
      try {
        var p = JSON.parse(prev);
        if (Object.prototype.toString.call(p) === '[object Array]') arr = p;
      } catch (e0) {}
    }
    arr = arr.concat(deltas || []);
    var payload = JSON.stringify(arr);
    if (payload.length > 8500) {
      applyMergeDeltas_(arr);
      props.deleteProperty(MERGE_PROP_KEY);
      return;
    }
    props.setProperty(MERGE_PROP_KEY, payload);
    ensureMergeTrigger_();
  } catch (e1) {
    try { applyMergeDeltas_(deltas); } catch (e2) {}
  }
}

function ensureMergeTrigger_() {
  var triggers = ScriptApp.getProjectTriggers();
  var i;
  for (i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === MERGE_TRIGGER_FN) return;
  }
  ScriptApp.newTrigger(MERGE_TRIGGER_FN).timeBased().after(800).create();
}

function applyMergeDeltas_(deltas) {
  if (!deltas || !deltas.length) return;
  var ss = SpreadsheetApp.getActive();
  var sh = logSheet_(ss);
  var headers = ensureHeaders_(sh);
  try { dropReporteSheets_(ss); } catch (e0) {}
  try { mergeDaysFromDeltas_(ss, headers, deltas, sh); } catch (e1) {}
  try { mergeResponsablesFromDeltas_(ss, headers, deltas); } catch (e2) {}
  try { mergeAcumuladoFromDeltas_(ss, headers, deltas); } catch (e3) {}
  SpreadsheetApp.flush();
}

function runQueuedMerge_() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(25000)) return;
  try {
    var props = PropertiesService.getScriptProperties();
    var raw = props.getProperty(MERGE_PROP_KEY);
    if (!raw) return;
    props.deleteProperty(MERGE_PROP_KEY);
    var deltas = JSON.parse(raw);
    applyMergeDeltas_(deltas);
  } catch (err) {
  } finally {
    try { lock.releaseLock(); } catch (e4) {}
    try {
      ScriptApp.getProjectTriggers().forEach(function (t) {
        if (t.getHandlerFunction() === MERGE_TRIGGER_FN) ScriptApp.deleteTrigger(t);
      });
    } catch (e5) {}
  }
}

/**
 * Autotest de guardado (sin cache). Ejecutar desde el editor: runSaveSelfTest_
 * o GET ?test=1
 */
function runSaveSelfTest_() {
  var out = { ok: true, checks: [] };
  function push(id, pass, extra) {
    out.checks.push({ id: id, ok: !!pass, extra: extra || '' });
    if (!pass) out.ok = false;
  }
  try {
    var ss = SpreadsheetApp.getActive();
    var sh = logSheet_(ss);
    var headers = ensureHeaders_(sh);
    var stamp = Utilities.formatDate(new Date(), TZ, 'HHmmss');
    var lote = 'TEST-' + stamp;
    var dni = '99999999';
    var fecha = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
    var recs = [{
      clientId: 'test-' + stamp,
      data: {
        fecha: fecha,
        lote: lote,
        supervisor: 'TEST AUTO',
        supervisorDni: dni,
        turnoCampo: 'Mañana',
        turno: '1',
        md: '1',
        variedad: 'TEST',
        fundo: 'LICAPA',
        jarrasConv: 7,
        kgConv: 7.98,
        totalJarras: 7,
        totalKg: 7.98,
        avance: 0.1,
        jornales: 1,
        horaRegistro: '12:00'
      }
    }];
    var r1 = upsertBatch_(sh, recs, headers);
    SpreadsheetApp.flush();
    push('A-upsert-accept', r1.changed && r1.accepted.length === 1, String(r1.accepted));

    var key = fallbackKey_(fecha, lote, 'Mañana', dni);
    var found = false;
    var last = sh.getLastRow();
    if (last > 1) {
      var data = sh.getRange(2, 1, last - 1, headers.length).getValues();
      var i;
      for (i = 0; i < data.length; i++) {
        if (fallbackKeyFromRow_(data[i], headers) === key) {
          found = true;
          push('B-no-clientId', String(cell_(data[i], headers, 'ClientId') || '') === '', '');
          push('C-jarras', num(cell_(data[i], headers, 'Total Jarras')) === 7, '');
          break;
        }
      }
    }
    push('D-row-in-sheet', found, key);

    // Reenvío (update) misma clave
    recs[0].data.totalJarras = 9;
    recs[0].data.jarrasConv = 9;
    recs[0].data.kgConv = 10.26;
    recs[0].data.totalKg = 10.26;
    var r2 = upsertBatch_(sh, recs, headers);
    SpreadsheetApp.flush();
    push('E-update-existing', r2.existing.length === 1 && r2.accepted.length === 0, '');

    var jarras = 0;
    last = sh.getLastRow();
    if (last > 1) {
      var data2 = sh.getRange(2, 1, last - 1, headers.length).getValues();
      var count = 0;
      var j;
      for (j = 0; j < data2.length; j++) {
        if (fallbackKeyFromRow_(data2[j], headers) === key) {
          count++;
          jarras = num(cell_(data2[j], headers, 'Total Jarras'));
        }
      }
      push('F-no-duplicate-row', count === 1, 'count=' + count);
      push('G-updated-value', jarras === 9, 'jarras=' + jarras);
    }

    // Merge visible
    try {
      mergeDaysFromDeltas_(ss, headers, r2.deltas, sh);
      mergeResponsablesFromDeltas_(ss, headers, r2.deltas);
      mergeAcumuladoFromDeltas_(ss, headers, r2.deltas);
      SpreadsheetApp.flush();
      push('H-merge-ok', true, '');
    } catch (me) {
      push('H-merge-ok', false, String(me));
    }

    // Limpieza del lote de prueba
    last = sh.getLastRow();
    if (last > 1) {
      var keep = [];
      var all = sh.getRange(2, 1, last - 1, headers.length).getValues();
      var k;
      for (k = 0; k < all.length; k++) {
        if (fallbackKeyFromRow_(all[k], headers) !== key) keep.push(all[k]);
      }
      sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), headers.length).clearContent();
      if (keep.length) sh.getRange(2, 1, keep.length, headers.length).setValues(keep);
      SpreadsheetApp.flush();
    }
    push('I-cleanup', true, '');
  } catch (err) {
    out.ok = false;
    out.error = String(err);
  }
  return out;
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
    .addItem('Probar guardado POST', 'runSaveSelfTestMenu_')
    .addToUi();
}

function runSaveSelfTestMenu_() {
  var r = runSaveSelfTest_();
  var lines = (r.checks || []).map(function (c) {
    return (c.ok ? 'OK' : 'FAIL') + ' · ' + c.id + (c.extra ? ' · ' + c.extra : '');
  });
  SpreadsheetApp.getUi().alert(
    r.ok ? 'Test OK — cada POST se guarda' : 'Test con fallos',
    lines.join('\n') + (r.error ? '\n\n' + r.error : ''),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
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
    'TurnoCampo': normTurno_(d.turnoCampo),
    'Hora registro': horaCorta_(d.horaRegistro || d.horaEnvio),
    'Hora envío': d.horaEnvio || ''
  };
  return headers.map(function (name) {
    // No guardamos ClientId (aunque la hoja vieja aún tenga la columna).
    if (name === 'ClientId') return '';
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
  SpreadsheetApp.flush();
}

/** Upsert por fecha+lote+turnoCampo+supervisor (sin ClientId). */
function upsertBatch_(sh, recs, headers) {
  headers = headers || ensureHeaders_(sh);
  var last = sh.getLastRow();
  var byKey = {};
  if (last > 1) {
    var nData = last - 1;
    var data = sh.getRange(2, 1, nData, headers.length).getValues();
    var i;
    for (i = 0; i < data.length; i++) {
      var k0 = fallbackKeyFromRow_(data[i], headers);
      if (k0) byKey[k0] = i + 2;
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
    var echoId = String(item.clientId || d.clientId || '').trim();
    var lote = String(d.lote || '').trim();
    if (!lote) continue;
    d.fecha = toIsoFecha_(d.fecha) || String(d.fecha || '').trim();
    d.lote = lote;
    d.turnoCampo = normTurno_(d.turnoCampo);
    d.supervisorDni = String(d.supervisorDni || '').trim();
    var key = fallbackKeyFromData_(d);
    if (!key) continue;
    var values = buildRow_(headers, d, '');
    var ack = echoId || key;
    if (byKey.hasOwnProperty(key) && byKey[key] > 0) {
      var rowNum = byKey[key];
      var oldVals = sh.getRange(rowNum, 1, 1, headers.length).getValues()[0];
      updates.push({ row: rowNum, values: values });
      deltas.push({ old: oldVals, neu: values });
      existing.push(ack);
    } else {
      appends.push(values);
      deltas.push({ old: null, neu: values });
      byKey[key] = -1;
      accepted.push(ack);
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

function canonFundo_(v) {
  var s = String(v || '').trim().toUpperCase().replace(/\s+/g, ' ');
  if (s === 'LICAPA I' || s === 'LICAPA 1') return 'LICAPA I';
  if (s === 'LICAPA II' || s === 'LICAPA 2') return 'LICAPA II';
  if (s === 'LICAPA III' || s === 'LICAPA 3') return 'LICAPA III';
  if (s === 'LICAPA IV' || s === 'LICAPA 4') return 'LICAPA IV';
  return '';
}

function fundoLabel(fundo, etapa) {
  var picked = canonFundo_(fundo);
  if (picked) return picked;
  var e = String(etapa || '').trim();
  var f = String(fundo || 'LICAPA').trim();
  var fromEtapa = canonFundo_(e);
  if (fromEtapa) return fromEtapa;
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

function etapaSheetLabel_(fundo, etapa) {
  var picked = canonFundo_(fundo) || canonFundo_(etapa);
  if (picked === 'LICAPA I') return 'ETAPA I';
  if (picked === 'LICAPA II') return 'ETAPA II';
  if (picked === 'LICAPA III') return 'ETAPA III';
  if (picked === 'LICAPA IV') return 'ETAPA IV';
  var s = String(etapa || fundo || '').toUpperCase();
  var m = /\b(IV|III|II|I)\b/.exec(s);
  return m ? 'ETAPA ' + m[1] : 'ETAPA I';
}

function daySheet_(ss, iso, etapaLabel, create) {
  var base = sheetNameDia_(iso);
  if (!base) return null;
  var tag = String(etapaLabel || '').trim() || 'ETAPA I';
  var name = base + '-' + tag;
  if (name.length > 99) name = name.slice(0, 99);
  var sh = ss.getSheetByName(name);
  if (!sh && create) sh = ss.insertSheet(name);
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
  var n = rows && rows.length ? rows.length : 0;
  var totRow = n + 2;
  var clearTo = Math.max(sh.getLastRow(), totRow, 2);
  if (clearTo > 0) {
    sh.getRange(1, 1, clearTo, cols).clearContent().clearFormat();
  }

  sh.getRange(1, 1, 1, cols).setValues([RESUMEN_HEADERS]);
  var lastData = 1;
  if (n) {
    sh.getRange(2, 1, n, cols).setValues(rows);
    lastData = 1 + n;
    // Datos siempre en peso normal (evita heredar negrita del TOTAL anterior).
    sh.getRange(2, 1, n, cols)
      .setFontWeight('normal')
      .setBackground('#FFFFFF')
      .setFontColor('#000000');
  }
  var empty = ['', '', '', '', '', '', '', 0, '', '', '', 0, 0, 0, '', '', ''];
  if (n) {
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
    var fundoRaw = cell_(row, headers, 'Fundo');
    var etapaRaw = cell_(row, headers, 'Etapa');
    var etapaSheet = etapaSheetLabel_(fundoRaw, etapaRaw);
    var fundo = fundoLabel(fundoRaw, etapaRaw);
    var area = num(cell_(row, headers, 'Avance') || cell_(row, headers, 'Area'));
    if (turnoCampo === 'Tarde') {
      var base = mananaAv[fecha + '|' + String(cell_(row, headers, 'Lote') || '').trim()] || 0;
      if (area >= base) area = Math.round((area - base) * 1000) / 1000;
    }
    var jornales = num(cell_(row, headers, 'Jornales'));
    var tipos = tiposDeFila_(row, headers);
    tipos.forEach(function (t, idx) {
      var key = [fecha, etapaSheet, t.tipo, fundo, md, variedad, turno].join('|');
      if (!groups[key]) {
        groups[key] = {
          iso: fecha,
          etapaSheet: etapaSheet,
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
    var bucket = g.iso + '|' + (g.etapaSheet || 'ETAPA I');
    if (!byDay[bucket]) byDay[bucket] = [];
    var jornales = num(g.jornales);
    var kgHa = g.area > 0 ? Math.round(g.kilos / g.area) : '';
    var kgJn = jornales > 0 ? Math.round(g.kilos / jornales) : '';
    byDay[bucket].push([
      g.semana, g.fecha, g.tipo, g.fundo, g.md, g.variedad, g.turno || '',
      g.envase, '', g.responsable || RESPONSABLE_FIJO, '',
      Math.round(g.area * 100) / 100, jornales,
      Math.round(g.kilos), kgHa, kgJn, horaLabel_(g)
    ]);
  });
  return byDay;
}

function writeByDay_(ss, byDay) {
  Object.keys(byDay || {}).sort().forEach(function (bucket) {
    var rows = byDay[bucket] || [];
    if (!rows.length) return;
    var parts = String(bucket).split('|');
    var iso = parts[0];
    var etapa = parts.slice(1).join('|') || 'ETAPA I';
    var sh = daySheet_(ss, iso, etapa, true);
    if (sh) writeDaySheet_(sh, rows);
  });
}

function dayRowKey_(row) {
  return [
    String(row[2] || '').trim().toUpperCase(),
    String(row[3] || '').trim().toUpperCase().replace(/\s+/g, ' '),
    mdLabel_(row[4]),
    String(row[5] || '').trim().toUpperCase().replace(/\s+/g, ' '),
    turnoLote_(row[6])
  ].join('|');
}

function readDaySheetRows_(sh) {
  var cols = RESUMEN_HEADERS.length;
  var last = sh.getLastRow();
  // Fila 1 = encabezado, última = TOTAL → datos = last - 2 filas.
  var nData = last - 2;
  if (nData < 1) return [];
  var data = sh.getRange(2, 1, nData, cols).getValues();
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
      // Preferir textos canónicos del envío nuevo
      if (r[2]) cur[2] = r[2];
      if (r[3]) cur[3] = r[3];
      if (r[4] !== '' && r[4] != null) cur[4] = r[4];
      if (r[5]) cur[5] = r[5];
      if (r[6] !== '' && r[6] != null) cur[6] = r[6];
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
    var buckets = {};
    Object.keys(bySub || {}).forEach(function (k) {
      if (k === iso || k.indexOf(iso + '|') === 0) buckets[k] = true;
    });
    Object.keys(byAdd || {}).forEach(function (k) {
      if (k === iso || k.indexOf(iso + '|') === 0) buckets[k] = true;
    });
    if (!Object.keys(buckets).length) return;
    Object.keys(buckets).forEach(function (bucket) {
      var parts = String(bucket).split('|');
      var etapa = parts.length > 1 ? parts.slice(1).join('|') : 'ETAPA I';
      var map = {};
      var order = [];
      var existing = daySheet_(ss, iso, etapa, false);
      if (existing) {
        readDaySheetRows_(existing).forEach(function (r) {
          var k = dayRowKey_(r);
          if (!map.hasOwnProperty(k)) {
            order.push(k);
            map[k] = r.slice(0, RESUMEN_HEADERS.length);
          }
        });
      }
      applyDayRows_(map, order, bySub[bucket] || [], -1);
      applyDayRows_(map, order, byAdd[bucket] || [], 1);
      var out = [];
      var seen = {};
      order.forEach(function (k) {
        if (!map[k] || seen[k]) return;
        seen[k] = true;
        out.push(map[k]);
      });
      if (!out.length) return;
      var sh = existing || daySheet_(ss, iso, etapa, true);
      if (sh) writeDaySheet_(sh, out);
    });
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

  function respKeyFromSheetRow_(fechaDisp, dni, nombre) {
    var iso = toIsoFecha_(fechaDisp);
    if (!iso) {
      var m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(String(fechaDisp || '').trim());
      if (m) iso = m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
    }
    var id = String(dni || '').replace(/\D/g, '').slice(0, 8);
    var nom = String(nombre || '').replace(/\s+/g, ' ').trim().toUpperCase();
    if (!iso) return '';
    return iso + '|' + (id.length === 8 ? id : nom);
  }

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
      var key = respKeyFromSheetRow_(fechaDisp, dni, nombre);
      if (!key || key.charAt(key.length - 1) === '|') return;
      // Duplicado en hoja: conservar solo la primera (no sumar otra vez).
      if (map.hasOwnProperty(key)) return;
      var iso = key.split('|')[0];
      var tc = String(row[ti >= 0 ? ti : 3] || '');
      map[key] = {
        iso: iso,
        fecha: fmtDate(iso) || String(fechaDisp || ''),
        supervisor: String(nombre || '').replace(/\s+/g, ' ').trim() || String(dni || ''),
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
  var seen = {};
  order.forEach(function (key) {
    if (!map[key] || seen[key]) return;
    seen[key] = true;
    var g = map[key];
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
  writeByDay_(ss, byDay);
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
  if (clearTo > 0) sh.getRange(1, 1, clearTo, cols).clearContent().clearFormat();
  sh.getRange(1, 1, 1, cols).setValues([RESP_HEADERS]);
  sh.getRange(1, 1, 1, cols).setBackground('#F3F3F3').setFontWeight('bold').setHorizontalAlignment('center');
  if (rows && rows.length) {
    sh.getRange(2, 1, rows.length, cols).setValues(rows);
    sh.getRange(2, 1, rows.length, cols)
      .setFontWeight('normal')
      .setBackground('#FFFFFF')
      .setHorizontalAlignment('center');
    sh.getRange(2, 5, rows.length, 1).setNumberFormat('#,##0');
    sh.getRange(2, 6, rows.length, 1).setNumberFormat('0.00');
  }
  if (wasEmpty) {
    sh.setFrozenRows(1);
    var widths = [92, 220, 100, 110, 70, 70, 86];
    widths.forEach(function (w, i) { sh.setColumnWidth(i + 1, w); });
    sh.setTabColor('#F7941D');
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
  dropReporteSheets_(ss);
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

/* ─── Acumulado (todos los supervisores, orden de envío) ─── */

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

function areaFromLogRow_(row, headers) {
  var av = num(cell_(row, headers, 'Avance'));
  if (av > 0) return av;
  return num(cell_(row, headers, 'Area'));
}

/** Borra hojas viejas tipo 'R 24-09-2026 NOMBRE' (ya no se usan). */
function dropReporteSheets_(ss) {
  ss = ss || SpreadsheetApp.getActive();
  ss.getSheets().slice().forEach(function (s) {
    if (ss.getSheets().length <= 1) return;
    var name = String(s.getName() || '').trim();
    if (/^R \d{2}-\d{2}-\d{4}\b/.test(name)) {
      try { ss.deleteSheet(s); } catch (e) {}
    }
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
    horaCorta_(cell_(row, headers, 'Hora registro') || cell_(row, headers, 'Hora envío'))
  ];
}

function acumuladoKeyFromValues_(values) {
  if (!values || !values.length) return '';
  var iso = toIsoFecha_(values[0]);
  if (!iso) {
    var m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(String(values[0] || '').trim());
    if (m) iso = m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  }
  var dni = String(values[3] || '').replace(/\D/g, '').slice(0, 8);
  var lote = String(values[8] || '').trim();
  var tc = /tarde/i.test(String(values[4] || '')) ? 'Tarde' : 'Mañana';
  if (!iso || !lote || !dni) return '';
  return iso + '|' + lote + '|' + tc + '|' + dni;
}

function mergeAcumuladoFromDeltas_(ss, headers, deltas) {
  var sh = ensureAcumuladoSheet_(ss);
  var byKey = {};
  var last = sh.getLastRow();
  if (last > 1) {
    var data = sh.getRange(2, 1, last, ACUM_HEADERS.length).getValues();
    var i;
    for (i = 0; i < data.length; i++) {
      var k = acumuladoKeyFromValues_(data[i]);
      if (k) byKey[k] = i + 2;
    }
  }

  var updates = [];
  var appends = [];
  var seen = {};
  (deltas || []).forEach(function (d) {
    var neu = d.neu;
    if (!neu) return;
    var key = fallbackKeyFromRow_(neu, headers);
    if (!key || seen[key]) return;
    seen[key] = true;
    var values = acumuladoRowFromLog_(neu, headers);
    if (byKey.hasOwnProperty(key)) {
      updates.push({ row: byKey[key], values: values });
    } else {
      appends.push(values);
      byKey[key] = -1;
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
