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

// ── HELPERS internos ────────────────────────────────────────────────
// Normaliza ciudades del form → nombres simples
function _parseCiudadesForm(raw) {
  return String(raw||'')
    .replace(/🟣 San Luis \([^)]+\)/g, 'San Luis')
    .replace(/🔵 Córdoba \([^)]+\)/g, 'Córdoba')
    .replace(/🟡 Tucumán \([^)]+\)/g, 'Tucumán');
}

// Extiende un string CSV de estados para que tenga `n` elementos
// (los nuevos llevan el valor `fill` si la ciudad aplica, '' si no)
function _extenderEstados(existingCsv, nViejo, nNuevo, fill) {
  const arr = String(existingCsv||'').split(',').map(s=>s.trim());
  while (arr.length < nViejo) arr.push('');      // completar hasta el viejo largo
  for (let i = nViejo; i < nNuevo; i++) arr.push(fill); // nuevos temas
  return arr.slice(0, nNuevo).join(',');
}

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

  // ── Cargar speakers existentes indexados por mail ────────────────
  const existingData = spSheet.getDataRange().getValues();
  // Fila 0 = encabezados (o primer dato si no hay encabezado)
  // Datos desde fila índice 1 → rowIndex en sheet = i + 2
  const byMail = {};
  existingData.slice(1).forEach((row, i) => {
    const m = String(row[2]||'').trim().toLowerCase();
    if (m) byMail[m] = { sheetRow: i + 2, row: [...row] };
  });

  const newRows      = allData.slice(lastImported);
  const imported     = [];
  const actualizados = [];

  newRows.forEach(r => {
    // Columnas del form (orden del Sheet de respuestas del Form):
    // [0] Timestamp  [1] Nombre completo  [2] Tipo  [3] Mail
    // [4] Móvil (WhatsApp)  [5] X (Twitter)  [6] Instagram  [7] LinkedIn
    // [8] Empresa/Referencia  [9] Ciudad(es)  [10] Tema(s)
    // [11] Notas / Comentarios  [12] Biografía  [13] Eventos anteriores
    const nombre   = String(r[1]  || '').trim();
    const tipo     = String(r[2]  || 'speaker').trim().toLowerCase();
    const mail     = String(r[3]  || '').trim();
    const movil    = String(r[4]  || '').trim();
    const xUser    = String(r[5]  || '').trim();
    const ig       = String(r[6]  || '').trim();
    const linkedin = String(r[7]  || '').trim();
    const empresa  = String(r[8]  || '').trim();
    const ciudRaw  = String(r[9]  || '').trim();
    // r[10]: checkboxes separados por coma, cada opción "Tema — Bajada"
    const temasRaw  = String(r[10] || '').trim();
    const temasArr  = temasRaw.split(',').map(t => t.split(' — ')[0].trim()).filter(Boolean);
    const temasCsv  = temasArr.join(', ');
    const notas     = String(r[11] || '').trim();
    const bio       = String(r[12] || '').trim().slice(0, 100);
    const eventos   = String(r[13] || '').trim();

    if (!nombre) return;

    const ciudades = _parseCiudadesForm(ciudRaw);
    const hasSL  = ciudades.includes('San Luis');
    const hasCBA = ciudades.includes('Córdoba');
    const hasTUC = ciudades.includes('Tucumán');

    const mailKey = mail.toLowerCase();

    if (mailKey && byMail[mailKey]) {
      // ── Speaker ya existe → fusionar temas / ciudades ────────────
      const entry = byMail[mailKey];
      const exRow = entry.row;

      const existTemasArr = String(exRow[4]||'').split(',').map(t=>t.trim()).filter(Boolean);
      const addedTemas = temasArr.filter(t => t && !existTemasArr.includes(t));
      const mergedTemas = [...existTemasArr, ...addedTemas];
      const nViejo = existTemasArr.length;
      const nNuevo = mergedTemas.length;

      // Ciudades: agregar las nuevas
      const existCiuds = String(exRow[3]||'');
      const ciudNuevas = ['San Luis','Córdoba','Tucumán'].filter(c => ciudades.includes(c) && !existCiuds.includes(c));
      const mergedCiuds = ciudNuevas.length
        ? (existCiuds ? existCiuds + ', ' + ciudNuevas.join(', ') : ciudNuevas.join(', '))
        : existCiuds;

      // Extender arrays de estado para los temas nuevos
      const mergedSL  = _extenderEstados(exRow[9],  nViejo, nNuevo, hasSL  ? 'disponible' : '');
      const mergedSJ  = _extenderEstados(exRow[10], nViejo, nNuevo, hasTUC ? 'disponible' : '');
      const mergedCBA = _extenderEstados(exRow[11], nViejo, nNuevo, hasCBA ? 'disponible' : '');

      const updRow = [...exRow];
      updRow[3]  = mergedCiuds;
      updRow[4]  = mergedTemas.join(', ');
      updRow[9]  = mergedSL;
      updRow[10] = mergedSJ;
      updRow[11] = mergedCBA;
      // Completar campos vacíos con los del form
      if (!updRow[5]  && notas)    updRow[5]  = notas;
      if (!updRow[12] && movil)    updRow[12] = movil;
      if (!updRow[13] && linkedin) updRow[13] = linkedin;
      if (!updRow[14] && bio)      updRow[14] = bio;
      if (!updRow[15] && eventos)  updRow[15] = eventos;

      spSheet.getRange(entry.sheetRow, 1, 1, updRow.length).setValues([updRow]);

      if (addedTemas.length) {
        actualizados.push(nombre + ' (+ ' + addedTemas.join(', ') + ')');
      } else {
        actualizados.push(nombre + ' (sin temas nuevos)');
      }

    } else {
      // ── Speaker nuevo → agregar fila ─────────────────────────────
      const n = Math.max(1, temasArr.length);
      const mkEst = (aplica) => aplica ? Array(n).fill('disponible').join(',') : Array(n).fill('').join(',');
      const slEst  = mkEst(hasSL);
      const sjEst  = mkEst(hasTUC);
      const cbaEst = mkEst(hasCBA);

      // Formato fila Speakers (16 cols):
      // nombre, tipo, mail, ciudades, temas, notas, x, ig, empresa,
      // sl_estado, sj_estado, cba_estado, movil, linkedin, bio, eventos_anteriores
      spSheet.appendRow([nombre, tipo, mail, ciudades, temasCsv, notas,
        xUser, ig, empresa, slEst, sjEst, cbaEst, movil, linkedin, bio, eventos]);
      byMail[mailKey] = { sheetRow: spSheet.getLastRow(), row: [] }; // registrar para evitar dup en mismo lote
      imported.push(nombre);
    }
  });

  props.setProperty('form_last_imported_row', String(totalRows));
  props.setProperty('version_Speakers', Date.now().toString());

  const total = imported.length + actualizados.length;
  let msg = '';
  if (total === 0) {
    msg = 'No se procesó nada (filas sin nombre o sin mail).';
  } else {
    const partes = [];
    if (imported.length)     partes.push('✅ Nuevos: ' + imported.join(', '));
    if (actualizados.length) partes.push('🔄 Actualizados: ' + actualizados.join(', '));
    msg = partes.join(' | ');
  }

  return { ok: true, imported: total, msg };
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
