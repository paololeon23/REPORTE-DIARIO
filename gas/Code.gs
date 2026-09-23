var TZ = 'America/Lima';
var JARRA_KG = 1.16;
var SHEET = 'Produccion';
var HEADERS = [
  'Fecha', 'Scanner', 'Scanner DNI', 'Supervisor', 'Supervisor DNI',
  'Grupo', 'Etapa', 'Lote', 'Fundo', 'Variedad', 'MD', 'Turno', 'Area', 'Avance',
  'Jarras Conv', 'Kg Conv', 'Jarras China', 'Kg China', 'Total Jarras', 'Total Kg', 'Jornales', 'ClientId'
];

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    var body = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActive();
    var sh = ss.getSheetByName(SHEET) || ss.insertSheet(SHEET);
    if (sh.getLastRow() === 0) sh.appendRow(HEADERS);
    var recs = body.action === 'batchSave' ? (body.records || []) : [body];
    recs.forEach(function (item) {
      var d = item.data || {};
      sh.appendRow([
        d.fecha, d.scanner, d.scannerDni, d.supervisor, d.supervisorDni,
        d.grupo, d.etapa, d.lote, d.fundo, d.variedad, d.md, d.turno, d.area, d.avance,
        d.jarrasConv, d.kgConv, d.jarrasChina, d.kgChina, d.totalJarras, d.totalKg, d.jornales, item.clientId
      ]);
    });
    return ContentService.createTextOutput(JSON.stringify({ ok: true })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function setupSheets() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(SHEET) || ss.insertSheet(SHEET);
  sh.clear();
  sh.appendRow(HEADERS);
}
