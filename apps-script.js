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

    // Acción especial: importar respuestas del Form → Speakers
    if (e.parameter.action === 'import_form_speakers') {
      const result = importarFormSpeakers(ss);
      return respond(result);
    }
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

// ── IMPORTAR RESPUESTAS DEL FORM → SPEAKERS ────────────────────────
const FORM_RESP_SHEET_ID = '1nChz2Vjur-ChW3fwnIj7aXXsGn--064Hu8Xf8DBvuDY';

function importarFormSpeakers(dexSS) {
  let respSS;
  try {
    respSS = SpreadsheetApp.openById(FORM_RESP_SHEET_ID);
  } catch(e) {
    return { ok: false, error: 'No puedo abrir el Sheet de respuestas: ' + e.message };
  }

  const respSheet = respSS.getSheets()[0];
  const allData   = respSheet.getDataRange().getValues();
  const totalRows = allData.length;

  const props        = PropertiesService.getScriptProperties();
  const lastImported = parseInt(props.getProperty('form_last_imported_row') || '1');

  if (totalRows <= lastImported) {
    return { ok: true, imported: 0, msg: 'No hay respuestas nuevas para importar.' };
  }

  const spSheet = dexSS.getSheetByName('Speakers');
  if (!spSheet) return { ok: false, error: 'No existe la pestaña "Speakers".' };

  const newRows  = allData.slice(lastImported);
  const imported = [];

  newRows.forEach(r => {
    // Columnas del form (orden del Sheet de respuestas):
    // [0] Timestamp
    // [1] Nombre completo
    // [2] Tipo
    // [3] Mail
    // [4] Móvil (WhatsApp)
    // [5] X (Twitter)
    // [6] Instagram
    // [7] LinkedIn
    // [8] Empresa/Referencia
    // [9] Ciudad(es)
    // [10] Tema(s) que vas a cubrir
    // [11] Notas / Comentarios
    // [12] Biografía
    // [13] Eventos anteriores
    const nombre   = String(r[1]  || '').trim();
    const tipo     = String(r[2]  || 'speaker').trim().toLowerCase();
    const mail     = String(r[3]  || '').trim();
    const movil    = String(r[4]  || '').trim();
    const xUser    = String(r[5]  || '').trim();
    const ig       = String(r[6]  || '').trim();
    const linkedin = String(r[7]  || '').trim();
    const empresa  = String(r[8]  || '').trim();
    const ciudRaw  = String(r[9]  || '').trim();
    // r[10]: "Tema1 — Bajada1, Tema2 — Bajada2" (checkboxes, hasta 3)
    const temasRaw = String(r[10] || '').trim();
    const temas    = temasRaw.split(',').map(t => t.split(' — ')[0].trim()).filter(Boolean).join(', ');
    const notas    = String(r[11] || '').trim();
    const bio      = String(r[12] || '').trim().slice(0, 100);
    const eventos  = String(r[13] || '').trim();

    if (!nombre) return;

    const ciudades = ciudRaw
      .replace(/🟣 San Luis \([^)]+\)/g, 'San Luis')
      .replace(/🔵 Córdoba \([^)]+\)/g, 'Córdoba')
      .replace(/🟡 Tucumán \([^)]+\)/g, 'Tucumán');

    // Estado inicial: disponible en cada ciudad seleccionada
    const slEst  = ciudades.includes('San Luis') ? 'disponible' : '';
    const sjEst  = ciudades.includes('Tucumán')  ? 'disponible' : '';
    const cbaEst = ciudades.includes('Córdoba')  ? 'disponible' : '';

    // Formato fila Speakers:
    // nombre, tipo, mail, ciudades, temas, notas, x, ig, empresa,
    // sl_estado, sj_estado, cba_estado, movil, linkedin, bio, eventos_anteriores
    spSheet.appendRow([nombre, tipo, mail, ciudades, temas, notas, xUser, ig, empresa, slEst, sjEst, cbaEst, movil, linkedin, bio, eventos]);
    imported.push(nombre);
  });

  props.setProperty('form_last_imported_row', String(totalRows));
  props.setProperty('version_Speakers', Date.now().toString());

  return {
    ok: true,
    imported: imported.length,
    msg: imported.length > 0
      ? '✅ ' + imported.length + ' speaker(s) importados: ' + imported.join(', ')
      : 'No se importó nada (filas sin nombre).'
  };
}

// ── BACKUP diario ───────────────────────────────────────────────────
function backupPrincipal() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const src = ss.getSheetByName('Principal');
  if (!src) return;

  const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm');
  const nombre = 'Backup_' + fecha;

  // Borrar backups de más de 7 días para no acumular hojas
  ss.getSheets().forEach(sh => {
    if (sh.getName().startsWith('Backup_')) {
      const partes = sh.getName().replace('Backup_', '').split('_');
      const fechaSh = new Date(partes[0]);
      const dias = (new Date() - fechaSh) / 86400000;
      if (dias > 7) ss.deleteSheet(sh);
    }
  });

  src.copyTo(ss).setName(nombre);
}

// Correr UNA vez desde el editor para instalar el trigger diario
function installBackupTrigger() {
  // Eliminar triggers anteriores del mismo tipo para no duplicar
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'backupPrincipal')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('backupPrincipal')
    .timeBased()
    .everyDays(1)
    .atHour(3)
    .create();
}

function respond(obj, code) {
  const out = ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
  return out;
}
