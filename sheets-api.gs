// Japan Trip — Google Apps Script Web App
// Deploy as: Execute as Me | Anyone (even anonymous)
const SECRET = 'Japan26';

function doGet(e) {
  const p = e.parameter;
  if (p.token !== SECRET) return json({ error: 'Unauthorized' });

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (p.action === 'list_tabs') {
    return json({ tabs: ss.getSheets().map(s => s.getName()) });
  }

  const tabName = p.tab;
  if (!tabName) return json({ error: 'Missing tab parameter' });

  const sheet = ss.getSheetByName(tabName);
  if (!sheet) return json({ error: `Tab "${tabName}" not found` });

  const data = sheet.getDataRange().getValues();
  if (data.length === 0) return json({ tab: tabName, headers: [], rows: [] });

  const headers = data[0].map(String);
  const rows = data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i] !== undefined ? String(row[i]) : ''; });
    return obj;
  });

  return json({ tab: tabName, headers, rows });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); }
  catch (_) { return json({ error: 'Invalid JSON body' }); }

  if (body.token !== SECRET) return json({ error: 'Unauthorized' });

  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const tab  = body.tab;
  const act  = body.action;
  if (!tab)  return json({ error: 'Missing tab' });
  if (!act)  return json({ error: 'Missing action' });

  // Helper: get or create sheet
  function getOrCreate(name) {
    return ss.getSheetByName(name) || ss.insertSheet(name);
  }

  const sheet = act === 'replace_sheet' ? getOrCreate(tab) : ss.getSheetByName(tab);
  if (!sheet) return json({ error: `Tab "${tab}" not found` });

  if (act === 'update_row') {
    const row   = Number(body.row);   // 1-based data row (header=1, first data=2)
    const data  = body.data;          // object { col: value }
    if (!row || !data) return json({ error: 'Missing row or data' });
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    headers.forEach((h, i) => {
      if (data[h] !== undefined) sheet.getRange(row, i + 1).setValue(data[h]);
    });
    return json({ ok: true });
  }

  if (act === 'append_row') {
    const data = body.data;
    if (!data) return json({ error: 'Missing data' });
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const rowVals = headers.map(h => data[h] !== undefined ? data[h] : '');
    sheet.appendRow(rowVals);
    return json({ ok: true, row: sheet.getLastRow() });
  }

  if (act === 'delete_row') {
    const row = Number(body.row);
    if (!row) return json({ error: 'Missing row' });
    sheet.deleteRow(row);
    return json({ ok: true });
  }

  if (act === 'update_cell') {
    const row = Number(body.row), col = Number(body.col);
    const val = body.value;
    if (!row || !col || val === undefined) return json({ error: 'Missing row/col/value' });
    sheet.getRange(row, col).setValue(val);
    return json({ ok: true });
  }

  if (act === 'replace_sheet') {
    const headers = body.headers;
    const rows    = body.rows;
    if (!headers || !rows) return json({ error: 'Missing headers or rows' });
    sheet.clearContents();
    sheet.appendRow(headers);
    rows.forEach(r => sheet.appendRow(Array.isArray(r) ? r : headers.map(h => r[h] !== undefined ? r[h] : '')));
    return json({ ok: true, rows: rows.length });
  }

  return json({ error: `Unknown action "${act}"` });
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
