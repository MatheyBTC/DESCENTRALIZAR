// ═══════════════════════════════════════════════════════════════════
// DEX-eventos — Google Apps Script Backend
// Pegar en: Extensions > Apps Script > pegar todo > Deploy > Web App
// Execute as: Me | Who has access: Anyone
// ═══════════════════════════════════════════════════════════════════

const SHEET_ID = '1wl2ClpRqJ5I4j92D0Xa3vinm0JHckCUCAu0fMfXJ07U';

// ── GET — Leer hojas / versiones ────────────────────────────────────
function doGet(e) {
  try {
    const sheet = (e.parameter.sheet || 'Principal').trim();

    // Acción especial: devolver versión actual de una hoja
    if (e.parameter.action === 'get_version') {
      const props = PropertiesService.getScriptProperties();
      const version = props.getProperty('version_' + sheet) || '0';
      return respond({ ok: true, sheet, version });
    }

    const ss  = SpreadsheetApp.openById(SHEET_ID);
    const ws  = ss.getSheetByName(sheet);
    if (!ws) return respond({ error: 'Hoja no encontrada: ' + sheet }, 404);
    const data = ws.getDataRange().getValues();

    // Devolver versión junto con los datos
    const props = PropertiesService.getScriptProperties();
    const version = props.getProperty('version_' + sheet) || '0';

    return respond({ ok: true, sheet, data, version });
  } catch(err) {
    return respond({ error: err.message }, 500);
  }
}

// ── POST — Escribir / modificar hojas ──────────────────────────────
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const { sheet, action, row, data, rowIndex, key } = payload;

    // Verificar clave de escritura
    const props = PropertiesService.getScriptProperties();
    const writeKey = props.getProperty('write_key');
    if (writeKey && key !== writeKey) {
      return respond({ error: 'Clave incorrecta', code: 401 });
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const ws = ss.getSheetByName(sheet);
    if (!ws) return respond({ error: 'Hoja no encontrada: ' + sheet }, 404);

    if (action === 'append') {
      ws.appendRow(data);
      return respond({ ok: true, action });
    }

    if (action === 'update') {
      // rowIndex es 1-based (fila 1 = headers, datos desde fila 2)
      const r = ws.getRange(rowIndex, 1, 1, data.length);
      r.setValues([data]);
      return respond({ ok: true, action });
    }

    if (action === 'delete') {
      ws.deleteRow(rowIndex);
      return respond({ ok: true, action });
    }

    if (action === 'replace_all') {
      // Reemplaza todos los datos (excepto headers) con el nuevo array
      const lastRow = ws.getLastRow();
      if (lastRow > 1) ws.deleteRows(2, lastRow - 1);
      if (data.length > 0) ws.getRange(2, 1, data.length, data[0].length).setValues(data);

      // Actualizar versión (timestamp) para detección de concurrencia
      const newVersion = Date.now().toString();
      PropertiesService.getScriptProperties().setProperty('version_' + sheet, newVersion);

      return respond({ ok: true, action, version: newVersion });
    }

    return respond({ error: 'Acción desconocida: ' + action }, 400);

  } catch(err) {
    return respond({ error: err.message }, 500);
  }
}

function respond(obj, code) {
  const out = ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
  return out;
}
