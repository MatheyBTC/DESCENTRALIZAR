/**
 * form-creator.gs — Crea el Google Form de inscripción de Speakers DEX 2026
 *
 * INSTRUCCIONES (una sola vez):
 *   1. Abrí el Google Sheet del DEX → Extensiones → Apps Script
 *   2. Nuevo archivo → "form-creator" → pegá este código
 *   3. Ejecutá createSpeakerForm() → aceptá permisos
 *   4. El log te da la URL pública para compartir con speakers
 *
 * Las respuestas van a un Sheet INDEPENDIENTE (no toca el Sheet del DEX).
 * Para importar al DEX: botón "📥 Importar del Form" en la pestaña Speakers.
 */

const DEX_SHEET_ID = '1wl2ClpRqJ5I4j92D0Xa3vinm0JHckCUCAu0fMfXJ07U';

// Tipos que NO son charlas (se excluyen del desplegable de temas)
const TIPOS_EXCLUIDOS = ['Kahoot', 'Break', 'Almuerzo', 'Apertura', 'Cierre', 'Premios', 'Sorteo', 'Concurso'];

// ═══════════════════════════════════════════════════════════════
// CREAR EL FORM (ejecutar una sola vez)
// ═══════════════════════════════════════════════════════════════
function createSpeakerForm() {

  // ── Leer temas reales desde la hoja Principal ────────────────
  const temas = getTemasDesdeSheet();
  if (temas.length === 0) {
    Logger.log('⚠️  No se encontraron temas en el Sheet. El desplegable quedará vacío.');
    Logger.log('    Verificá que la hoja "Principal" tenga datos en la columna Tema.');
  }

  // ── Crear el Form ────────────────────────────────────────────
  const form = FormApp.create('DEX 2026 — Inscripción de Speaker');
  form.setDescription(
    'Completá este formulario para participar como speaker en el DEX 2026.\n' +
    'El equipo organizador te contactará para confirmar tu participación.'
  );
  form.setCollectEmail(false);
  form.setAllowResponseEdits(true);
  form.setConfirmationMessage('✅ ¡Gracias! Recibimos tu inscripción. Te contactamos pronto.');

  // ── Campo 1: Nombre ──────────────────────────────────────────
  form.addTextItem()
    .setTitle('Nombre completo')
    .setRequired(true);

  // ── Campo 2: Tipo ────────────────────────────────────────────
  form.addMultipleChoiceItem()
    .setTitle('Tipo')
    .setRequired(true)
    .setChoiceValues(['Speaker', 'Empresa', 'Sponsor']);

  // ── Campo 3: Contacto ────────────────────────────────────────
  form.addTextItem()
    .setTitle('Contacto (mail / móvil)')
    .setHelpText('Ej: nombre@mail.com / +54 9 11 1234-5678')
    .setRequired(true);

  // ── Campo 4: X (Twitter) ─────────────────────────────────────
  form.addTextItem()
    .setTitle('X (Twitter)')
    .setHelpText('Ej: @usuario')
    .setRequired(false);

  // ── Campo 5: Instagram ───────────────────────────────────────
  form.addTextItem()
    .setTitle('Instagram')
    .setHelpText('Ej: @usuario')
    .setRequired(false);

  // ── Campo 6: Empresa / Referencia ────────────────────────────
  form.addTextItem()
    .setTitle('Empresa / Referencia')
    .setHelpText('Proyecto, empresa o rol actual.')
    .setRequired(false);

  // ── Campo 7: Ciudad(es) ──────────────────────────────────────
  form.addCheckboxItem()
    .setTitle('Ciudad(es) en las que podés participar')
    .setRequired(true)
    .setChoiceValues([
      '🟣 San Luis (14 ago 2026)',
      '🔵 Córdoba (4 sep 2026)',
      '🟡 Tucumán (11 sep 2026)'
    ]);

  // ── Campo 8: Tema(s) — desplegable con los temas del Sheet ───
  const temaItem = form.addListItem()
    .setTitle('Tema que vas a cubrir')
    .setHelpText('Seleccioná el tema principal de tu charla.')
    .setRequired(true);

  const opciones = temas.length > 0
    ? temas
    : ['(Sin temas cargados aún — completar manualmente)'];
  temaItem.setChoiceValues(opciones);

  // ── Campo 9: Notas ───────────────────────────────────────────
  form.addParagraphTextItem()
    .setTitle('Notas / Comentarios')
    .setHelpText('Disponibilidad, restricciones horarias, necesidades técnicas, lo que quieras aclarar.')
    .setRequired(false);

  // ── Linkear respuestas a un Sheet NUEVO e independiente ──────
  // Google crea automáticamente un Sheet nuevo con las respuestas
  const respSheet = SpreadsheetApp.create('DEX2026 — Respuestas Speakers');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, respSheet.getId());

  // Guardar el ID del Sheet de respuestas para que el import lo encuentre
  PropertiesService.getScriptProperties().setProperty('FORM_RESP_SHEET_ID', respSheet.getId());
  PropertiesService.getScriptProperties().setProperty('SPEAKER_FORM_URL', form.getPublishedUrl());
  PropertiesService.getScriptProperties().setProperty('form_last_imported_row', '1');

  Logger.log('');
  Logger.log('✅ Form creado exitosamente');
  Logger.log('');
  Logger.log('🔗 URL para compartir con speakers:');
  Logger.log('   ' + form.getPublishedUrl());
  Logger.log('');
  Logger.log('📊 Sheet de respuestas (independiente):');
  Logger.log('   https://docs.google.com/spreadsheets/d/' + respSheet.getId());
  Logger.log('');
  Logger.log('Cuando quieras importar al DEX, usá el botón "📥 Importar del Form" en la app.');
}

// ═══════════════════════════════════════════════════════════════
// LEER TEMAS DESDE EL SHEET DEX
// Lee columna Tipo (col 0) y Tema (col 4), excluye tipos no-charla
// ═══════════════════════════════════════════════════════════════
function getTemasDesdeSheet() {
  try {
    const ss    = SpreadsheetApp.openById(DEX_SHEET_ID);
    const sheet = ss.getSheetByName('Principal');
    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();

    const temas = data.slice(1) // saltar header
      .filter(r => {
        const tipo  = String(r[0] || '').trim();
        const tema  = String(r[4] || '').trim();
        return tema.length > 0 && !TIPOS_EXCLUIDOS.includes(tipo);
      })
      .map(r => String(r[4]).trim());

    return [...new Set(temas)]; // únicos, sin duplicados
  } catch(e) {
    Logger.log('Error leyendo temas: ' + e.message);
    return [];
  }
}

// ═══════════════════════════════════════════════════════════════
// IMPORTAR RESPUESTAS → SPEAKERS (llamado desde la app)
// ═══════════════════════════════════════════════════════════════
function importarFormSpeakers() {
  const props       = PropertiesService.getScriptProperties();
  const respSheetId = props.getProperty('FORM_RESP_SHEET_ID');

  if (!respSheetId) {
    return { ok: false, error: 'No encontré el Sheet de respuestas. Ejecutá createSpeakerForm() primero.' };
  }

  // Abrir el Sheet de respuestas independiente
  let respSS;
  try {
    respSS = SpreadsheetApp.openById(respSheetId);
  } catch(e) {
    return { ok: false, error: 'No puedo abrir el Sheet de respuestas: ' + e.message };
  }

  const respSheet = respSS.getSheets()[0]; // primera pestaña = respuestas
  const allData   = respSheet.getDataRange().getValues();
  const totalRows = allData.length;

  const lastImported = parseInt(props.getProperty('form_last_imported_row') || '1');

  if (totalRows <= lastImported) {
    return { ok: true, imported: 0, msg: 'No hay respuestas nuevas para importar.' };
  }

  // Abrir el Sheet del DEX y la pestaña Speakers
  const dexSS   = SpreadsheetApp.openById(DEX_SHEET_ID);
  const spSheet = dexSS.getSheetByName('Speakers');
  if (!spSheet) return { ok: false, error: 'No existe la pestaña "Speakers" en el Sheet del DEX.' };

  const newRows  = allData.slice(lastImported);
  const imported = [];

  newRows.forEach(r => {
    // Columnas del form (en orden de aparición en el Sheet de respuestas):
    // [0] Timestamp
    // [1] Nombre completo
    // [2] Tipo
    // [3] Contacto (mail / móvil)
    // [4] X (Twitter)
    // [5] Instagram
    // [6] Empresa / Referencia
    // [7] Ciudad(es)
    // [8] Tema que vas a cubrir
    // [9] Notas / Comentarios

    const nombre   = String(r[1] || '').trim();
    const tipo     = String(r[2] || 'Speaker').trim().toLowerCase();
    const contacto = String(r[3] || '').trim();
    const xUser    = String(r[4] || '').trim();
    const ig       = String(r[5] || '').trim();
    const empresa  = String(r[6] || '').trim();
    const ciudRaw  = String(r[7] || '').trim();
    const temas    = String(r[8] || '').trim();
    const notas    = String(r[9] || '').trim();

    if (!nombre) return;

    // Normalizar ciudades (quitar emoji y fecha)
    const ciudades = ciudRaw
      .replace(/🟣 San Luis \([^)]+\)/g, 'San Luis')
      .replace(/🔵 Córdoba \([^)]+\)/g, 'Córdoba')
      .replace(/🟡 Tucumán \([^)]+\)/g, 'Tucumán');

    // Formato fila Speakers (igual al loadSpeakers del frontend):
    // [0]=nombre [1]=tipo [2]=contacto [3]=ciudades [4]=temas
    // [5]=notas  [6]=x   [7]=ig       [8]=empresa
    // [9]=sl_estado [10]=sj_estado [11]=cba_estado
    spSheet.appendRow([
      nombre, tipo, contacto, ciudades, temas,
      notas, xUser, ig, empresa,
      '', '', ''
    ]);

    imported.push(nombre);
  });

  // Actualizar puntero para no reimportar en el futuro
  props.setProperty('form_last_imported_row', String(totalRows));
  // Invalidar versión Speakers para que el frontend detecte el cambio
  props.setProperty('version_Speakers', Date.now().toString());

  return {
    ok: true,
    imported: imported.length,
    msg: imported.length > 0
      ? '✅ ' + imported.length + ' speaker(s) importados: ' + imported.join(', ')
      : 'No se importó nada (filas sin nombre).'
  };
}
