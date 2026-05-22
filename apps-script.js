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
    .replace(/🟥 San Luis \([^)]+\)/g, 'San Luis')
    .replace(/🟨 Córdoba \([^)]+\)/g, 'Córdoba')
    .replace(/🟩 Tucumán \([^)]+\)/g, 'Tucumán')
    // compatibilidad con emojis anteriores
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

// ── Detectar offset de columnas en el Sheet de respuestas ───────────
// El form puede haber generado columnas duplicadas si se reinició.
// Esta función lee la fila de encabezados y detecta en qué columna
// empieza el bloque que tiene datos reales en la fila `dataRow`.
// Devuelve el índice (0-based) de la columna "Nombre completo" activa.
function _detectarOffsetNombre(headers, dataRow) {
  const nombreIndices = [];
  headers.forEach((h, i) => {
    if (String(h).trim() === 'Nombre completo') nombreIndices.push(i);
  });
  // Encontrar cuál de los bloques tiene datos
  for (const idx of nombreIndices) {
    if (dataRow[idx] && String(dataRow[idx]).trim()) return idx;
  }
  // Fallback: primer "Nombre completo" encontrado
  return nombreIndices[0] !== undefined ? nombreIndices[0] : 1;
}

function importarFormSpeakers(dexSS) {
  let respSS;
  try {
    respSS = SpreadsheetApp.openById(FORM_RESP_SHEET_ID);
  } catch(e) {
    return { ok: false, error: 'No puedo abrir el Sheet de respuestas: ' + e.message };
  }

  // Buscar la hoja de respuestas activa (la que tiene más datos entre las "Form Responses X")
  const allSheets = respSS.getSheets();
  // Buscar la pestaña de respuestas activa: preferir "Form Responses X"
  // (la que Google crea al vincular); si no hay, usar la de más filas
  const formSheets = allSheets.filter(s => {
    const n = s.getName().toLowerCase();
    return n.includes('form response') || n.includes('respuestas del formulario');
  });
  // Usar la hoja con número más alto (= la más reciente, la activa del form)
  const _numSheet = (s) => parseInt(s.getName().replace(/\D/g,'')) || 0;
  const respSheet = formSheets.length > 0
    ? formSheets.sort((a, b) => _numSheet(b) - _numSheet(a))[0]
    : allSheets.sort((a, b) => b.getLastRow() - a.getLastRow())[0];

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
  const byMail = {};
  existingData.slice(1).forEach((row, i) => {
    const m = String(row[2]||'').trim().toLowerCase();
    if (m) byMail[m] = { sheetRow: i + 2, row: [...row] };
  });

  // ── Encabezados del sheet de respuestas (para detectar offset) ───
  const headers = allData[0] || [];

  const newRows      = allData.slice(lastImported);
  const imported     = [];
  const actualizados = [];

  newRows.forEach(r => {
    // Detectar en qué bloque de columnas cayeron los datos de esta fila
    // (el form puede generar columnas duplicadas si fue reseteado)
    const nombreCol = _detectarOffsetNombre(headers, r);
    // nombreCol = índice de "Nombre completo" → los demás campos van +1, +2, etc.
    // Timestamp siempre está en col 0 (independiente del bloque)
    const c = (offset) => r[nombreCol + offset] || '';

    const nombre   = String(c(0)  || '').trim();              // Nombre completo
    const tipo     = String(c(1)  || 'speaker').trim();        // Tipo
    const mail     = String(c(2)  || '').trim();               // Mail
    const movilRaw = String(c(3)  || '').trim().replace(/\D/g,''); // Móvil — solo dígitos
    const movil    = movilRaw ? 'https://wa.me/' + movilRaw : '';  // → URL wa.me
    const telegram = String(c(4)  || '').trim();               // Telegram/Signal
    const xUser    = String(c(5)  || '').trim();               // X (Twitter)
    const ig       = String(c(6)  || '').trim();               // Instagram
    const linkedin = String(c(7)  || '').trim();               // LinkedIn
    const empresa  = String(c(8)  || '').trim();               // Empresa/Referencia
    const ciudRaw  = String(c(9)  || '').trim();               // Ciudad(es)
    const temasRaw = String(c(10) || '').trim();               // Tema(s)
    const notas    = String(c(11) || '').trim();               // Notas
    const bio      = String(c(12) || '').trim().slice(0, 150); // Biografía (150 chars)
    const eventos  = String(c(13) || '').trim();               // Eventos anteriores
    const signal   = '';                                        // Signal (campo separado — pendiente en form)

    const temasArr = temasRaw.split(',').map(t => t.split(' — ')[0].trim()).filter(Boolean);
    const temasCsv = temasArr.join(', ');

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
      if (!updRow[13] && telegram) updRow[13] = telegram;
      if (!updRow[14] && linkedin) updRow[14] = linkedin;
      if (!updRow[15] && bio)      updRow[15] = bio;
      if (!updRow[16] && eventos)  updRow[16] = eventos;
      // updRow[17] = signal (se preserva lo que ya había)

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

      // Formato fila Speakers (18 cols):
      // nombre, tipo, mail, ciudades, temas, notas, x, ig, empresa,
      // sl_estado, sj_estado, cba_estado, movil, telegram, linkedin, bio, eventos_anteriores, signal
      spSheet.appendRow([nombre, tipo, mail, ciudades, temasCsv, notas,
        xUser, ig, empresa, slEst, sjEst, cbaEst, movil, telegram, linkedin, bio, eventos, signal]);
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

// ── UTILIDADES DE MANTENIMIENTO ─────────────────────────────────────
// Correr desde el editor cuando el contador de importación queda desfasado
function resetContadorImport() {
  PropertiesService.getScriptProperties()
    .setProperty('form_last_imported_row', '1');
  Logger.log('✅ Contador reseteado a 1');
}

// ── CLAVE DE ESCRITURA ───────────────────────────────────────────────
// Correr UNA vez desde el editor para setear o cambiar la clave
function setWriteKey() {
  const CLAVE = 'DEX/2026';
  PropertiesService.getScriptProperties().setProperty('write_key', CLAVE);
  Logger.log('✅ Clave seteada: ' + CLAVE);
}

// Correr para ver cuál es la clave actual guardada
function verWriteKey() {
  const k = PropertiesService.getScriptProperties().getProperty('write_key');
  Logger.log('Clave actual: ' + (k || '(no seteada — cualquier clave pasa)'));
}

// Correr para borrar la clave (cualquier usuario podrá guardar sin clave)
function borrarWriteKey() {
  PropertiesService.getScriptProperties().deleteProperty('write_key');
  Logger.log('✅ Clave borrada');
}

// ── SINCRONIZACIÓN DE TEMAS → GOOGLE FORM ───────────────────────────
// Form de registro de speakers:
const FORM_ID = '1x5OzFZXkSv2dqt7933fCiB3zQO7pYeFhlKPfSm-BCI0';

// Tipos de bloque sin speaker (se excluyen del Form)
const TIPOS_SIN_SPEAKER = ['break','almuerzo','kahoot','apertura','cierre','sorteo','premios','concurso','ama'];

// Util: listar todos los items del Form con su tipo, título y opciones (para debug)
function listarItemsForm() {
  const form = FormApp.openById(FORM_ID);
  form.getItems().forEach(item => {
    Logger.log('[' + item.getType() + '] id:' + item.getId() + ' → "' + item.getTitle() + '"');
    try {
      const type = item.getType();
      let choices = [];
      if (type === FormApp.ItemType.CHECKBOX)        choices = item.asCheckboxItem().getChoices();
      else if (type === FormApp.ItemType.LIST)        choices = item.asListItem().getChoices();
      else if (type === FormApp.ItemType.MULTIPLE_CHOICE) choices = item.asMultipleChoiceItem().getChoices();
      choices.forEach(c => Logger.log('   · "' + c.getValue() + '"'));
    } catch(e) {}
  });
}

// Actualiza los emojis de ciudad en TODAS las preguntas del Form
// 🟣→🟥  🔵→🟨  🟡→🟩
// Correr manualmente una vez para migrar el Form
function actualizarCiudadesEnForm() {
  const REEMPLAZOS = [
    ['🟣', '🟥'],
    ['🔵', '🟨'],
    ['🟡', '🟩'],
  ];
  const form  = FormApp.openById(FORM_ID);
  const items = form.getItems();
  let cambios = 0;

  items.forEach(item => {
    const type = item.getType();
    try {
      if (type === FormApp.ItemType.CHECKBOX) {
        const cb = item.asCheckboxItem();
        const orig = cb.getChoices().map(c => c.getValue());
        const nuevo = orig.map(v => REEMPLAZOS.reduce((s,[a,b]) => s.split(a).join(b), v));
        if (orig.join('|') !== nuevo.join('|')) {
          cb.setChoices(nuevo.map(v => cb.createChoice(v)));
          Logger.log('✅ Checkbox "' + item.getTitle() + '" actualizado');
          orig.forEach((v,i) => { if (v!==nuevo[i]) Logger.log('   "'+v+'" → "'+nuevo[i]+'"'); });
          cambios++;
        }
      } else if (type === FormApp.ItemType.LIST) {
        const li = item.asListItem();
        const orig = li.getChoices().map(c => c.getValue());
        const nuevo = orig.map(v => REEMPLAZOS.reduce((s,[a,b]) => s.split(a).join(b), v));
        if (orig.join('|') !== nuevo.join('|')) {
          li.setChoices(nuevo.map(v => li.createChoice(v)));
          Logger.log('✅ Lista "' + item.getTitle() + '" actualizada');
          cambios++;
        }
      } else if (type === FormApp.ItemType.MULTIPLE_CHOICE) {
        const mc = item.asMultipleChoiceItem();
        const orig = mc.getChoices().map(c => c.getValue());
        const nuevo = orig.map(v => REEMPLAZOS.reduce((s,[a,b]) => s.split(a).join(b), v));
        if (orig.join('|') !== nuevo.join('|')) {
          mc.setChoices(nuevo.map(v => mc.createChoice(v)));
          Logger.log('✅ Opción múltiple "' + item.getTitle() + '" actualizada');
          cambios++;
        }
      }
    } catch(e) {
      Logger.log('⚠️ Error en "' + item.getTitle() + '": ' + e.message);
    }
  });

  Logger.log(cambios ? '✅ ' + cambios + ' pregunta(s) actualizada(s)' : 'ℹ️ Sin cambios — los emojis ya están actualizados');
}

// Actualiza las opciones del campo Tema(s) del Form con los temas actuales de Principal
// Se puede correr manualmente o via trigger nocturno
function actualizarTemasEnForm() {
  // 1. Leer temas de Principal (hoja de cálculo)
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const ws = ss.getSheetByName('Principal');
  if (!ws) { Logger.log('❌ Hoja Principal no encontrada'); return; }

  const rows = ws.getDataRange().getValues().slice(1); // saltar header
  // Columnas: tipo=0, dur=1, inicio=2, fin=3, tema=4, bajada=5
  const TIPO_COL = 0, TEMA_COL = 4, BAJADA_COL = 5;

  const temas = [];
  const seen  = new Set();
  rows.forEach(r => {
    const tipo  = String(r[TIPO_COL]  || '').trim().toLowerCase();
    const tema  = String(r[TEMA_COL]  || '').trim();
    const bajada = String(r[BAJADA_COL] || '').trim();
    if (!tema) return;
    if (TIPOS_SIN_SPEAKER.some(t => tipo.includes(t))) return;
    if (seen.has(tema)) return;
    seen.add(tema);
    // Formato de opción: "Tema — bajada" (el import ya sabe extraer solo el tema)
    temas.push(bajada ? tema + ' — ' + bajada : tema);
  });

  if (!temas.length) { Logger.log('⚠️ Sin temas en Principal — Form no modificado'); return; }

  // 2. Abrir el Form y buscar el item de temas
  const form  = FormApp.openById(FORM_ID);
  const items = form.getItems();

  // Buscar por título que contenga "tema" (case-insensitive)
  const temaItem = items.find(i => i.getTitle().toLowerCase().includes('tema'));
  if (!temaItem) {
    Logger.log('❌ No encontré item con "tema" en el título. Items disponibles:');
    items.forEach(i => Logger.log('  [' + i.getType() + '] "' + i.getTitle() + '"'));
    return;
  }
  Logger.log('📋 Item encontrado: "' + temaItem.getTitle() + '" (tipo: ' + temaItem.getType() + ')');

  // 3. Actualizar opciones según el tipo de item
  const type = temaItem.getType();
  try {
    if (type === FormApp.ItemType.CHECKBOX) {
      const cb = temaItem.asCheckboxItem();
      cb.setChoices(temas.map(t => cb.createChoice(t)));
    } else if (type === FormApp.ItemType.LIST) {
      const li = temaItem.asListItem();
      li.setChoices(temas.map(t => li.createChoice(t)));
    } else if (type === FormApp.ItemType.MULTIPLE_CHOICE) {
      const mc = temaItem.asMultipleChoiceItem();
      mc.setChoices(temas.map(t => mc.createChoice(t)));
    } else {
      Logger.log('❌ Tipo de item no soportado: ' + type + ' — usá Checkbox, Lista o Opción múltiple');
      return;
    }
    Logger.log('✅ Form actualizado con ' + temas.length + ' temas:');
    temas.forEach(t => Logger.log('   · ' + t));
  } catch(e) {
    Logger.log('❌ Error al actualizar opciones: ' + e.message);
  }
}

// Instalar trigger nocturno (correr UNA vez desde el editor)
// Después de instalarlo, actualizarTemasEnForm() corre solo todos los días a las 3am
function installFormUpdateTrigger() {
  // Eliminar triggers anteriores del mismo tipo para no duplicar
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'actualizarTemasEnForm')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('actualizarTemasEnForm')
    .timeBased()
    .atHour(3)       // 3am (hora del servidor de Google, UTC−3 aprox. para ARG)
    .everyDays(1)
    .create();

  Logger.log('✅ Trigger nocturno instalado — actualizarTemasEnForm() correrá diariamente a las 3am');
}
