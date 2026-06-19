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

    if (e.parameter.action === 'get_version') {
      const props = PropertiesService.getScriptProperties();
      const version = props.getProperty('version_' + sheet) || '0';
      return respond({ ok: true, sheet, version });
    }

    const ss = SpreadsheetApp.openById(SHEET_ID);

    if (e.parameter.action === 'get_meta') {
      const ss = SpreadsheetApp.openById(SHEET_ID);
      let ws = ss.getSheetByName('Meta');
      if (!ws) return respond({ ok: true, meta: {} });
      const rows = ws.getDataRange().getValues();
      const meta = {};
      rows.forEach(r => { if (r[0]) meta[String(r[0])] = String(r[1]||''); });
      return respond({ ok: true, meta });
    }

    if (e.parameter.action === 'import_form_speakers') {
      const force = e.parameter.force === '1';
      if (force) PropertiesService.getScriptProperties().setProperty('form_last_imported_row', '1');
      const result = importarFormSpeakers(ss);
      return respond(result);
    }

    const ws = ss.getSheetByName(sheet);
    if (!ws) return respond({ error: 'Hoja no encontrada: ' + sheet }, 404);
    const data = ws.getDataRange().getValues();

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
      const r = ws.getRange(rowIndex, 1, 1, data.length);
      r.setValues([data]);
      return respond({ ok: true, action });
    }

    if (action === 'delete') {
      ws.deleteRow(rowIndex);
      return respond({ ok: true, action });
    }

    if (action === 'save_meta') {
      const ss = SpreadsheetApp.openById(SHEET_ID);
      let ws = ss.getSheetByName('Meta');
      if (!ws) {
        ws = ss.insertSheet('Meta');
        ws.getRange(1,1,1,2).setValues([['clave','valor']]);
      }
      const metaData = payload.meta || {};
      const existing = ws.getDataRange().getValues();
      const keyIndex = {};
      existing.forEach((r,i) => { if (i>0 && r[0]) keyIndex[String(r[0])] = i+1; });
      Object.entries(metaData).forEach(([k,v]) => {
        if (keyIndex[k]) {
          ws.getRange(keyIndex[k], 2).setValue(v);
        } else {
          ws.appendRow([k, v]);
        }
      });
      return respond({ ok: true, action });
    }

    if (action === 'replace_all') {
      const lastRow = ws.getLastRow();
      if (lastRow > 1) ws.deleteRows(2, lastRow - 1);
      if (data.length > 0) ws.getRange(2, 1, data.length, data[0].length).setValues(data);

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

function _parseCiudadesForm(raw) {
  return String(raw||'')
    .replace(/🟥 San Luis \([^)]+\)/g, 'San Luis')
    .replace(/🟨 Córdoba \([^)]+\)/g, 'Córdoba')
    .replace(/🟩 Tucumán \([^)]+\)/g, 'Tucumán')
    .replace(/🟣 San Luis \([^)]+\)/g, 'San Luis')
    .replace(/🔵 Córdoba \([^)]+\)/g, 'Córdoba')
    .replace(/🟡 Tucumán \([^)]+\)/g, 'Tucumán');
}

function _extenderEstados(existingCsv, nViejo, nNuevo, fill) {
  const arr = String(existingCsv||'').split(',').map(s=>s.trim());
  while (arr.length < nViejo) arr.push('');
  for (let i = nViejo; i < nNuevo; i++) arr.push(fill);
  return arr.slice(0, nNuevo).join(',');
}

function importarFormSpeakers(dexSS) {
  let respSS;
  try {
    respSS = SpreadsheetApp.openById(FORM_RESP_SHEET_ID);
  } catch(e) {
    return { ok: false, error: 'No puedo abrir el Sheet de respuestas: ' + e.message };
  }

  const allSheets  = respSS.getSheets();
  const formSheets = allSheets.filter(s => {
    const n = s.getName().toLowerCase();
    return n.includes('form response') || n.includes('respuestas del formulario');
  });
  const _numSheet  = (s) => parseInt(s.getName().replace(/\D/g,'')) || 0;
  const respSheet  = formSheets.length > 0
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

  // Indexar speakers existentes por mail (col D = índice 3)
  const existingData = spSheet.getDataRange().getValues();
  // Sin encabezado — detectar si fila 0 es header o dato
  const firstRow = existingData[0] || [];
  const hasHeader = typeof firstRow[1] === 'string' && firstRow[1].toLowerCase().includes('nombre');
  const dataRows = hasHeader ? existingData.slice(1) : existingData;
  const byMail = {};
  dataRows.forEach((row, i) => {
    const m = String(row[3]||'').trim().toLowerCase(); // col D = Mail
    const sheetRow = hasHeader ? i + 2 : i + 1;
    if (m) byMail[m] = { sheetRow, row: [...row] };
  });

  const newRows      = allData.slice(lastImported);
  const imported     = [];
  const actualizados = [];

  // Estructura del form (cols 0-15):
  // A(0)=Timestamp  B(1)=Nombre completo  C(2)=Tipo  D(3)=Mail
  // E(4)=Móvil      F(5)=Telegram         G(6)=Signal H(7)=X
  // I(8)=Instagram  J(9)=LinkedIn         K(10)=Empresa
  // L(11)=Ciudad(es) M(12)=Tema(s)        N(13)=Notas
  // O(14)=Biografía  P(15)=Eventos anteriores
  //
  // Speakers sheet = copia exacta A→A ... P→P
  // + Q(16)=# temas  R(17)=sl_estado  S(18)=sj_estado  T(19)=cba_estado

  newRows.forEach(r => {
    const nombre = String(r[1]  || '').trim();  // B
    const mail   = String(r[3]  || '').trim();  // D
    if (!nombre) return;

    // Contar temas (col M=12): separados por "., " antes de mayúscula/¿
    const temasRaw = String(r[12] || '').trim();
    const temasArr = temasRaw
      .split(/\.,\s*(?=[¿¡A-ZÁÉÍÓÚÜA-z])/)
      .map(t => t.trim())
      .filter(t => t.length > 3);
    const nTemas = Math.max(1, temasArr.length);

    // Estados iniciales según ciudades (col L=11)
    const ciudRaw = String(r[11] || '').trim();
    const ciudades = _parseCiudadesForm(ciudRaw);
    const slEst  = ciudades.includes('San Luis') ? Array(nTemas).fill('disponible').join(',') : '';
    const sjEst  = ciudades.includes('Tucumán')  ? Array(nTemas).fill('disponible').join(',') : '';
    const cbaEst = ciudades.includes('Córdoba')  ? Array(nTemas).fill('disponible').join(',') : '';

    const mailKey = mail.toLowerCase();

    if (mailKey && byMail[mailKey]) {
      // Speaker ya existe → solo rellena campos vacíos, respeta ediciones manuales
      const entry  = byMail[mailKey];
      const updRow = [...entry.row];

      for (let i = 0; i < 16; i++) {
        const sheetVal = String(updRow[i] || '').trim();
        const formVal  = String(r[i]      || '').trim();
        // Solo sobreescribir si el campo en el sheet está vacío
        if (!sheetVal && formVal) updRow[i] = r[i];
      }
      updRow[16] = nTemas;
      if (!updRow[17]) updRow[17] = slEst;
      if (!updRow[18]) updRow[18] = sjEst;
      if (!updRow[19]) updRow[19] = cbaEst;

      spSheet.getRange(entry.sheetRow, 1, 1, updRow.length).setValues([updRow]);
      actualizados.push(nombre);

    } else {
      // Speaker nuevo → copia exacta A-P + columnas calculadas
      const newRow = [
        r[0],  // A Timestamp
        r[1],  // B Nombre completo
        r[2],  // C Tipo
        r[3],  // D Mail
        r[4],  // E Móvil
        r[5],  // F Telegram
        r[6],  // G Signal
        r[7],  // H X
        r[8],  // I Instagram
        r[9],  // J LinkedIn
        r[10], // K Empresa
        r[11], // L Ciudad(es)
        r[12], // M Tema(s)
        r[13], // N Notas
        r[14], // O Biografía
        r[15], // P Eventos anteriores
        nTemas, // Q # temas
        slEst,  // R sl_estado
        sjEst,  // S sj_estado
        cbaEst  // T cba_estado
      ];
      spSheet.appendRow(newRow);
      byMail[mailKey] = { sheetRow: spSheet.getLastRow(), row: newRow };
      imported.push(nombre);
    }
  });

  props.setProperty('form_last_imported_row', String(totalRows));
  props.setProperty('version_Speakers', Date.now().toString());

  const total = imported.length + actualizados.length;
  let msg = '';
  if (total === 0) {
    msg = 'No se procesó nada (filas sin nombre).';
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

function installBackupTrigger() {
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
function resetContadorImport() {
  PropertiesService.getScriptProperties()
    .setProperty('form_last_imported_row', '1');
  Logger.log('✅ Contador reseteado a 1');
}

// ── CLAVE DE ESCRITURA ───────────────────────────────────────────────
function setWriteKey() {
  const CLAVE = 'DEX/2026';
  PropertiesService.getScriptProperties().setProperty('write_key', CLAVE);
  Logger.log('✅ Clave seteada: ' + CLAVE);
}

function verWriteKey() {
  const k = PropertiesService.getScriptProperties().getProperty('write_key');
  Logger.log('Clave actual: ' + (k || '(no seteada — cualquier clave pasa)'));
}

function borrarWriteKey() {
  PropertiesService.getScriptProperties().deleteProperty('write_key');
  Logger.log('✅ Clave borrada');
}

// ── SINCRONIZACIÓN DE TEMAS → GOOGLE FORM ───────────────────────────
const FORM_ID = '1x5OzFZXkSv2dqt7933fCiB3zQO7pYeFhlKPfSm-BCI0';

const TIPOS_SIN_SPEAKER = ['break','almuerzo','kahoot','apertura','cierre','sorteo','premios','concurso','ama'];

function listarItemsForm() {
  const form = FormApp.openById(FORM_ID);
  form.getItems().forEach(item => {
    Logger.log('[' + item.getType() + '] id:' + item.getId() + ' → "' + item.getTitle() + '"');
    try {
      const type = item.getType();
      let choices = [];
      if (type === FormApp.ItemType.CHECKBOX)             choices = item.asCheckboxItem().getChoices();
      else if (type === FormApp.ItemType.LIST)             choices = item.asListItem().getChoices();
      else if (type === FormApp.ItemType.MULTIPLE_CHOICE)  choices = item.asMultipleChoiceItem().getChoices();
      choices.forEach(c => Logger.log('   · "' + c.getValue() + '"'));
    } catch(e) {}
  });
}

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

  Logger.log(cambios ? '✅ ' + cambios + ' pregunta(s) actualizada(s)' : 'ℹ️ Sin cambios');
}

function actualizarTemasEnForm() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const ws = ss.getSheetByName('Principal');
  if (!ws) { Logger.log('❌ Hoja Principal no encontrada'); return; }

  const rows = ws.getDataRange().getValues().slice(1);
  const TIPO_COL = 0, TEMA_COL = 4, BAJADA_COL = 5;

  const temas = [];
  const seen  = new Set();
  rows.forEach(r => {
    const tipo   = String(r[TIPO_COL]   || '').trim().toLowerCase();
    const tema   = String(r[TEMA_COL]   || '').trim();
    const bajada = String(r[BAJADA_COL] || '').trim();
    if (!tema) return;
    if (TIPOS_SIN_SPEAKER.some(t => tipo.includes(t))) return;
    if (seen.has(tema)) return;
    seen.add(tema);
    temas.push(bajada ? tema + ' — ' + bajada : tema);
  });

  if (!temas.length) { Logger.log('⚠️ Sin temas en Principal — Form no modificado'); return; }

  const form  = FormApp.openById(FORM_ID);
  const items = form.getItems();

  const temaItem = items.find(i => i.getTitle().toLowerCase().includes('tema'));
  if (!temaItem) {
    Logger.log('❌ No encontré item con "tema" en el título.');
    items.forEach(i => Logger.log('  [' + i.getType() + '] "' + i.getTitle() + '"'));
    return;
  }
  Logger.log('📋 Item encontrado: "' + temaItem.getTitle() + '" (tipo: ' + temaItem.getType() + ')');

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
      Logger.log('❌ Tipo no soportado: ' + type);
      return;
    }
    Logger.log('✅ Form actualizado con ' + temas.length + ' temas:');
    temas.forEach(t => Logger.log('   · ' + t));
  } catch(e) {
    Logger.log('❌ Error: ' + e.message);
  }
}

function installFormUpdateTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'actualizarTemasEnForm')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('actualizarTemasEnForm')
    .timeBased()
    .atHour(3)
    .everyDays(1)
    .create();

  Logger.log('✅ Trigger nocturno instalado — actualizarTemasEnForm() correrá diariamente a las 3am');
}
