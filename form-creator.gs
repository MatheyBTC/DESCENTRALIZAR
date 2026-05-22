/**
 * form-creator.gs — Google Form de inscripción de Speakers DEX 2026
 *
 * Dónde pegar este código:
 *   → script.google.com → el proyecto del form (Apps Script del form)
 *   → Reemplazá todo el código → Guardar
 *
 * ═══════════════════════════════════════════════════════════════════
 * FUNCIONES DISPONIBLES — ejecutar desde el editor (▷ Run):
 *
 *   resetFormCampos()    ← USA ESTA para resetear el form al orden correcto
 *                          Borra todos los campos y los recrea en el orden
 *                          exacto que espera el backend.
 *
 *   updateTemaQuestion() ← Actualiza la pregunta de temas con bajada
 *                          (correr después de resetFormCampos, o cada vez
 *                          que cambien los temas en el Sheet Principal)
 *
 *   createSpeakerForm()  ← Solo si necesitás crear un form NUEVO desde cero
 * ═══════════════════════════════════════════════════════════════════
 *
 * ORDEN FINAL DE CAMPOS (= columnas en el Sheet de respuestas):
 *   [0] Timestamp (auto)
 *   [1] Nombre completo
 *   [2] Tipo
 *   [3] Mail
 *   [4] Móvil (WhatsApp)
 *   [5] X (Twitter)
 *   [6] Instagram
 *   [7] LinkedIn
 *   [8] Empresa/Referencia
 *   [9] Ciudad(es)
 *  [10] Tema(s)
 *  [11] Notas / Comentarios
 *  [12] Biografía
 *  [13] Eventos anteriores
 */

const DEX_SHEET_ID = '1wl2ClpRqJ5I4j92D0Xa3vinm0JHckCUCAu0fMfXJ07U';
const TIPOS_EXCLUIDOS = ['Kahoot', 'Break', 'Almuerzo', 'Apertura', 'Cierre', 'Premios', 'Sorteo', 'Concurso'];

// ═══════════════════════════════════════════════════════════════════
// RESET — Borra todos los campos y los recrea en el orden correcto
// Ejecutar UNA VEZ. No toca la configuración del form (título,
// mensaje de confirmación, destino del Sheet de respuestas).
// ═══════════════════════════════════════════════════════════════════
function resetFormCampos() {
  const form = _abrirForm();
  if (!form) return;

  // 1. Borrar TODOS los items existentes
  form.getItems().slice().forEach(i => {
    try { form.deleteItem(i); } catch(e) {}
  });
  Logger.log('🗑️  Todos los campos eliminados. Recreando en orden correcto…');

  // 2. Obtener temas desde el Sheet Principal
  const temas = getTemasDesdeSheet();
  const opcionesTema = temas.length > 0 ? temas : ['(Sin temas cargados aún)'];

  // 3. Recrear campos EN EL ORDEN EXACTO que espera el backend
  //    → cada campo agrega una columna al Sheet de respuestas

  // [1] Nombre completo
  form.addTextItem()
    .setTitle('Nombre completo')
    .setRequired(true);

  // [2] Tipo — checkbox (permite seleccionar múltiples roles)
  form.addCheckboxItem()
    .setTitle('Tipo')
    .setHelpText('Podés seleccionar más de uno. Ej: Speaker + Panelista.')
    .setChoiceValues(['Speaker', 'Moderador', 'Panelista', 'Sponsor'])
    .setRequired(false);

  // [3] Mail
  form.addTextItem()
    .setTitle('Mail')
    .setHelpText('Ej: nombre@mail.com — se usa para evitar duplicados al importar.')
    .setRequired(true);

  // [4] Móvil (WhatsApp)
  form.addTextItem()
    .setTitle('Móvil (WhatsApp)')
    .setHelpText('Solo números sin el +. Ej: 5491112345678')
    .setValidation(
      FormApp.createTextValidation()
        .requireTextMatchesPattern('[0-9]+')
        .build()
    );

  // [5] Telegram/Signal
  form.addTextItem()
    .setTitle('Telegram/Signal')
    .setHelpText('Ej: @usuario');

  // [6] X (Twitter)
  form.addTextItem()
    .setTitle('X (Twitter)')
    .setHelpText('Ej: @usuario');

  // [7] Instagram
  form.addTextItem()
    .setTitle('Instagram')
    .setHelpText('Ej: @usuario');

  // [8] LinkedIn
  form.addTextItem()
    .setTitle('LinkedIn')
    .setHelpText('Ej: linkedin.com/in/usuario');

  // [9] Empresa/Referencia
  form.addTextItem()
    .setTitle('Empresa/Referencia')
    .setHelpText('Proyecto, empresa o rol actual.');

  // [10] Ciudad(es) — checkboxes
  form.addCheckboxItem()
    .setTitle('Ciudad(es) en las que podés participar')
    .setRequired(true)
    .setChoiceValues([
      '🟣 San Luis (14 ago 2026)',
      '🔵 Córdoba (4 sep 2026)',
      '🟡 Tucumán (11 sep 2026)'
    ]);

  // [11] Tema(s) — checkboxes con bajada (cargados desde Sheet)
  form.addCheckboxItem()
    .setTitle('Tema(s) que vas a cubrir')
    .setHelpText('Seleccioná uno o más temas.')
    .setRequired(true)
    .setChoiceValues(opcionesTema);

  // [12] Notas / Comentarios
  form.addParagraphTextItem()
    .setTitle('Notas / Comentarios')
    .setHelpText('Disponibilidad, restricciones horarias, necesidades técnicas.');

  // [13] Biografía
  form.addParagraphTextItem()
    .setTitle('Biografía')
    .setHelpText('Descripción breve de tu perfil profesional (máx. 100 caracteres).');

  // [14] Eventos anteriores
  form.addTextItem()
    .setTitle('Eventos anteriores')
    .setHelpText('Eventos en los que participaste como speaker (opcional).');

  Logger.log('');
  Logger.log('✅ Form reseteado. Orden final:');
  Logger.log('   [1] Nombre · [2] Tipo · [3] Mail · [4] Móvil · [5] Telegram/Signal');
  Logger.log('   [6] X · [7] IG · [8] LinkedIn · [9] Empresa · [10] Ciudad(es) · [11] Tema(s)');
  Logger.log('   [12] Notas · [13] Biografía · [14] Eventos anteriores');
  Logger.log('');
  Logger.log('⚠️  IMPORTANTE: el Sheet de respuestas ahora tiene columnas viejas.');
  Logger.log('   → Si querés empezar limpio, borrá las filas de respuestas pasadas,');
  Logger.log('     o corré resetRespuestasSheet() para limpiar los encabezados.');
  Logger.log('');
  Logger.log('👉  Verificá el form: ' + form.getPublishedUrl());
}

// ═══════════════════════════════════════════════════════════════════
// ACTUALIZAR PREGUNTA DE TEMAS (con bajada desde el Sheet)
// Correr después de resetFormCampos(), o cada vez que cambien los
// temas en la pestaña Principal del Sheet DEX.
// ═══════════════════════════════════════════════════════════════════
function updateTemaQuestion() {
  const form = _abrirForm();
  if (!form) return;

  const temaItem = form.getItems().find(i => i.getTitle().startsWith('Tema'));
  if (!temaItem) { Logger.log('❌ No se encontró la pregunta de Tema en el form.'); return; }

  const opciones = getTemasDesdeSheet();
  if (opciones.length === 0) { Logger.log('⚠️ No hay temas en el Sheet. Cargá la hoja Principal primero.'); return; }

  const idx = temaItem.getIndex();
  form.deleteItem(temaItem);

  const nuevo = form.addCheckboxItem()
    .setTitle('Tema(s) que vas a cubrir')
    .setHelpText('Seleccioná uno o más temas.')
    .setRequired(true)
    .setChoiceValues(opciones);

  form.moveItem(nuevo.getIndex(), idx);

  Logger.log('✅ Temas actualizados: ' + opciones.length + ' opciones.');
  opciones.forEach(o => Logger.log('   • ' + o));
}

// ═══════════════════════════════════════════════════════════════════
// LIMPIAR SHEET DE RESPUESTAS
// Borra todas las filas de respuestas anteriores y recrea los
// encabezados con el orden correcto. Usar si querés empezar limpio
// después de resetFormCampos().
// ═══════════════════════════════════════════════════════════════════
function resetRespuestasSheet() {
  const RESP_SHEET_ID = '1nChz2Vjur-ChW3fwnIj7aXXsGn--064Hu8Xf8DBvuDY';

  // 1. Desvincular el form
  const form = _abrirForm();
  if (form) {
    try { form.removeDestination(); Logger.log('✅ Form desvinculado.'); }
    catch(e) { Logger.log('⚠️  removeDestination: ' + e.message); }
  }

  // 2. Abrir el spreadsheet
  let ss;
  try { ss = SpreadsheetApp.openById(RESP_SHEET_ID); }
  catch(e) { Logger.log('❌ No se pudo abrir el Sheet: ' + e.message); return; }

  // 3. Crear (o limpiar) una hoja llamada "Respuestas" como destino limpio
  //    No intentamos borrar las viejas "Form Responses X" — Google lo bloquea.
  //    El import las ignora porque busca por nombre y toma la de más datos.
  let target = ss.getSheetByName('Respuestas');
  if (!target) {
    target = ss.insertSheet('Respuestas', 0); // primera posición
    Logger.log('➕ Hoja "Respuestas" creada.');
  } else {
    target.clear();
    Logger.log('🧹 Hoja "Respuestas" limpiada.');
  }

  // 4. Escribir encabezados correctos en las 14 columnas
  const encabezados = [
    'Timestamp', 'Nombre completo', 'Tipo', 'Mail',
    'Móvil (WhatsApp)', 'X (Twitter)', 'Instagram', 'LinkedIn',
    'Empresa/Referencia', 'Ciudad(es) en las que podés participar',
    'Tema(s) que vas a cubrir', 'Notas / Comentarios',
    'Biografía', 'Eventos anteriores'
  ];
  target.getRange(1, 1, 1, encabezados.length).setValues([encabezados]);
  target.getRange(1, 1, 1, encabezados.length).setFontWeight('bold');
  Logger.log('✅ Encabezados escritos.');

  // 5. Re-vincular el form → Google crea "Form Responses 1" dentro del mismo spreadsheet
  //    Los próximos envíos van a esa nueva pestaña con columnas en orden correcto
  if (form) {
    try {
      form.setDestination(FormApp.DestinationType.SPREADSHEET, RESP_SHEET_ID);
      Logger.log('✅ Form re-vinculado al spreadsheet.');
      Logger.log('   Al primer envío Google crea "Form Responses X" con columnas en orden correcto.');
    } catch(e) {
      Logger.log('⚠️  No se pudo re-vincular: ' + e.message);
    }
  }

  // 6. Resetear contador de importación
  PropertiesService.getScriptProperties().setProperty('form_last_imported_row', '1');

  Logger.log('');
  Logger.log('✅ Todo listo. Próximos pasos:');
  Logger.log('   1. Mandá una respuesta de prueba al form');
  Logger.log('   2. Google crea "Form Responses X" con columnas correctas');
  Logger.log('   3. Importá desde la app → debería funcionar');
  Logger.log('👉 ' + 'https://docs.google.com/spreadsheets/d/' + RESP_SHEET_ID);
}

// ═══════════════════════════════════════════════════════════════════
// LEER TEMAS DESDE EL SHEET DEX
// ═══════════════════════════════════════════════════════════════════
function getTemasDesdeSheet() {
  try {
    const ss    = SpreadsheetApp.openById(DEX_SHEET_ID);
    const sheet = ss.getSheetByName('Principal');
    if (!sheet) return [];

    const data    = sheet.getDataRange().getValues();
    const seen    = new Set();
    const opciones = [];

    data.slice(1).forEach(r => {
      const tipo   = String(r[0]  || '').trim();
      const tema   = String(r[4]  || '').trim();
      const bajada = String(r[5]  || '').trim();
      if (!tema || TIPOS_EXCLUIDOS.includes(tipo) || seen.has(tema)) return;
      seen.add(tema);
      opciones.push(bajada ? tema + ' — ' + bajada : tema);
    });

    return opciones;
  } catch(e) {
    Logger.log('Error leyendo temas: ' + e.message);
    return [];
  }
}

// ═══════════════════════════════════════════════════════════════════
// CREAR FORM DESDE CERO (solo si necesitás uno nuevo)
// ═══════════════════════════════════════════════════════════════════
function createSpeakerForm() {
  const temas = getTemasDesdeSheet();

  const form = FormApp.create('DEX 2026 — Inscripción de Speaker');
  form.setDescription(
    'Completá este formulario para participar como speaker en el DEX 2026.\n' +
    'El equipo organizador te contactará para confirmar tu participación.'
  );
  form.setCollectEmail(false);
  form.setAllowResponseEdits(true);
  form.setConfirmationMessage('✅ ¡Gracias! Recibimos tu inscripción. Te contactamos pronto.');

  form.addTextItem().setTitle('Nombre completo').setRequired(true);
  form.addCheckboxItem().setTitle('Tipo').setHelpText('Podés seleccionar más de uno.').setChoiceValues(['Speaker','Moderador','Panelista','Sponsor']);
  form.addTextItem().setTitle('Mail').setHelpText('Ej: nombre@mail.com').setRequired(true);
  form.addTextItem().setTitle('Móvil (WhatsApp)').setHelpText('Ej: +54 9 11 1234-5678');
  form.addTextItem().setTitle('X (Twitter)').setHelpText('Ej: @usuario');
  form.addTextItem().setTitle('Instagram').setHelpText('Ej: @usuario');
  form.addTextItem().setTitle('LinkedIn').setHelpText('Ej: linkedin.com/in/usuario');
  form.addTextItem().setTitle('Empresa/Referencia').setHelpText('Proyecto, empresa o rol actual.');

  form.addCheckboxItem()
    .setTitle('Ciudad(es) en las que podés participar')
    .setRequired(true)
    .setChoiceValues([
      '🟣 San Luis (14 ago 2026)',
      '🔵 Córdoba (4 sep 2026)',
      '🟡 Tucumán (11 sep 2026)'
    ]);

  const opcionesTema = temas.length > 0 ? temas : ['(Sin temas cargados aún)'];
  form.addCheckboxItem()
    .setTitle('Tema(s) que vas a cubrir')
    .setHelpText('Seleccioná uno o más temas.')
    .setRequired(true)
    .setChoiceValues(opcionesTema);

  form.addParagraphTextItem()
    .setTitle('Notas / Comentarios')
    .setHelpText('Disponibilidad, restricciones horarias, necesidades técnicas.');

  form.addParagraphTextItem()
    .setTitle('Biografía')
    .setHelpText('Descripción breve de tu perfil profesional (máx. 100 caracteres).');

  form.addTextItem()
    .setTitle('Eventos anteriores')
    .setHelpText('Eventos en los que participaste como speaker (opcional).');

  const respSheet = SpreadsheetApp.create('DEX2026 — Respuestas Speakers');
  form.setDestination(FormApp.DestinationType.SPREADSHEET, respSheet.getId());

  const props = PropertiesService.getScriptProperties();
  props.setProperty('FORM_RESP_SHEET_ID', respSheet.getId());
  props.setProperty('SPEAKER_FORM_URL', form.getPublishedUrl());
  props.setProperty('form_last_imported_row', '1');

  Logger.log('✅ Form creado: ' + form.getPublishedUrl());
  Logger.log('📊 Sheet de respuestas: https://docs.google.com/spreadsheets/d/' + respSheet.getId());
}

// ── Helper: abre el form por ID hardcodeado ──────────────────────
const SPEAKER_FORM_ID = '1x5OzFZXkSv2dqt7933fCiB3zQO7pYeFhlKPfSm-BCI0';

function _abrirForm() {
  try {
    return FormApp.openById(SPEAKER_FORM_ID);
  } catch(e) {
    Logger.log('❌ No se pudo abrir el form: ' + e.message);
    return null;
  }
}
