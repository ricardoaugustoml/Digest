// ==================================================================
// Diário de Produção — API (Google Apps Script)
// Cole este código no editor de Apps Script vinculado à sua planilha.
// ==================================================================

var SHEET_NAME = 'Lancamentos';

// Troque por uma senha só sua. Use o MESMO valor no arquivo HTML (const TOKEN).
var TOKEN = 'CLARA-RAFAEL';

function getSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['id', 'date', 'procedure', 'payment', 'convenio', 'valor', 'quantidade']);
  }
  return sheet;
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function rowsToEntries(values) {
  var entries = [];
  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    if (!r[0]) continue;
    entries.push({
      id: r[0],
      date: r[1],
      procedure: r[2],
      payment: r[3],
      convenio: r[4],
      valor: Number(r[5]),
      quantidade: Number(r[6])
    });
  }
  return entries;
}

function doGet(e) {
  var action = e.parameter.action;
  var token = e.parameter.token;
  if (token !== TOKEN) return jsonResponse({ error: 'não autorizado' });

  if (action === 'list') {
    var values = getSheet().getDataRange().getValues();
    return jsonResponse({ ok: true, entries: rowsToEntries(values) });
  }
  return jsonResponse({ error: 'ação desconhecida' });
}

function doPost(e) {
  var body = JSON.parse(e.postData.contents);
  if (body.token !== TOKEN) return jsonResponse({ error: 'não autorizado' });
  var sheet = getSheet();

  if (body.action === 'add') {
    var entry = body.entry;
    entry.id = 'e' + new Date().getTime() + Math.floor(Math.random() * 1000);
    sheet.appendRow([entry.id, entry.date, entry.procedure, entry.payment, entry.convenio, entry.valor, entry.quantidade]);
    return jsonResponse({ ok: true, entry: entry });
  }

  if (body.action === 'update') {
    var values = sheet.getDataRange().getValues();
    for (var i = 1; i < values.length; i++) {
      if (values[i][0] === body.id) {
        var e2 = body.entry;
        sheet.getRange(i + 1, 1, 1, 7).setValues([[body.id, e2.date, e2.procedure, e2.payment, e2.convenio, e2.valor, e2.quantidade]]);
        break;
      }
    }
    return jsonResponse({ ok: true });
  }

  if (body.action === 'delete') {
    var values2 = sheet.getDataRange().getValues();
    for (var j = 1; j < values2.length; j++) {
      if (values2[j][0] === body.id) {
        sheet.deleteRow(j + 1);
        break;
      }
    }
    return jsonResponse({ ok: true });
  }

  return jsonResponse({ error: 'ação desconhecida' });
}
