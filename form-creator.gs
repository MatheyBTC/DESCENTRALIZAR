/**
 * form-creator.gs — Google Form de inscripción de Speakers DEX 2026
 *
 * Dónde pegar este código:
 *   → script.google.com → el proyecto donde creaste el form
 *   → Reemplazá todo el código → Guardar
 *
 * Funciones disponibles:
 *   · updateFormCampos()   — actualiza campos del form existente (correr 1 vez)
 *   · updateTemaQuestion() — convierte Tema a checkboxes con bajada (correr 1 vez)
 *   · createSpeakerForm()  — solo si necesitás crear el form desde cero
 */

const DEX_SHEET_ID = '1wl2ClpRqJ5I4j92D0Xa3vinm0JHckCUCAu0fMfXJ07U';

const TIPOS_EXCLUIDOS = ['Kahoot', 'Break', 'Almuerzo', 'Apertura', 'Cierre', 'Premios', 'Sorteo', 'Concurso'];

// ═══════════════════════════════════════════════════════════════
// PASO 1 — Actualizar campos del form existente
// Ejecutar UNA VEZ. Orden final:
// Nombre | Tipo | Mail | Móvil (WhatsApp) | X | Instagram |
// LinkedIn | Empresa/Referencia | Ciudad(es) | Tema(s) | Notas
// ═══════════════════════════════════════════════════════════════
function updateFormCampos() {
  const form = _abrirForm();
  if (!form) return;

  // ── 1. Eliminar campos que se van a recrear o ya no van ──────
  const ELIMINAR = [
    'Contacto (mail / móvil)', 'Contacto',
    'Móvil / WhatsApp', 'LinkedIn', 'Empresa / Referencia'
  ];
  // iterar sobre copia porque deleteItem modifica el array
  form.getItems().slice().forEach(i => {
    if (ELIMINAR.includes(i.getTitle())) {
      form.deleteItem(i);
      Logger.log('🗑️  Eliminado: ' + i.getTitle());
    }
  });

  // ── 2. Asegurar campo Tipo (texto libre, después de Nombre) ──
  if (!form.getItems().find(i => i.getTitle() === 'Tipo')) {
    const nombreIdx = form.getItems().findIndex(i => i.getTitle() === 'Nombre completo');
    const tipoNew   = form.addTextItem()
      .setTitle('Tipo')
      .setHelpText('Ej: Speaker, Panelista, Moderador');
    form.moveItem(tipoNew.getIndex(), nombreIdx >= 0 ? nombreIdx + 1 : 1);
    Logger.log('➕ Agregado: Tipo');
  }

  // ── 3. Renombrar "Empresa / Referencia" → "Empresa/Referencia"
  const empresaItem = form.getItems().find(i => i.getTitle().includes('Empresa'));
  if (empresaItem) {
    empresaItem.asTextItem().setTitle('Empresa/Referencia');
    Logger.log('✏️  Renombrado: Empresa/Referencia');
  }

  // ── 4. Agregar Mail (después de Tipo, requerido) ─────────────
  const tipoIdx = form.getItems().findIndex(i => i.getTitle() === 'Tipo');
  const mailItem = form.addTextItem()
    .setTitle('Mail')
    .setHelpText('Ej: nombre@mail.com')
    .setRequired(true);
  form.moveItem(mailItem.getIndex(), tipoIdx >= 0 ? tipoIdx + 1 : 2);
  Logger.log('➕ Agregado: Mail');

  // ── 5. Agregar Móvil (WhatsApp) después de Mail ──────────────
  const mailIdx  = form.getItems().findIndex(i => i.getTitle() === 'Mail');
  const movilItem = form.addTextItem()
    .setTitle('Móvil (WhatsApp)')
    .setHelpText('Ej: +54 9 11 1234-5678');
  form.moveItem(movilItem.getIndex(), mailIdx + 1);
  Logger.log('➕ Agregado: Móvil (WhatsApp)');

  // ── 6. Agregar LinkedIn antes de Empresa/Referencia ──────────
  const empresaIdx = form.getItems().findIndex(i => i.getTitle() === 'Empresa/Referencia');
  const linkedinItem = form.addTextItem()
    .setTitle('LinkedIn')
    .setHelpText('Ej: linkedin.com/in/usuario');
  form.moveItem(linkedinItem.getIndex(), empresaIdx >= 0 ? empresaIdx : form.getItems().length - 2);
  Logger.log('➕ Agregado: LinkedIn');

  // ── 7. Agregar Biografía si no existe (después de Notas) ────────
  if (!form.getItems().find(i => i.getTitle() === 'Biografía')) {
    const notasIdx = form.getItems().findIndex(i => i.getTitle() === 'Notas / Comentarios');
    const bioItem = form.addParagraphTextItem()
      .setTitle('Biografía')
      .setHelpText('Descripción breve de tu perfil profesional (máx. 100 caracteres).');
    form.moveItem(bioItem.getIndex(), notasIdx >= 0 ? notasIdx + 1 : form.getItems().length - 1);
    Logger.log('➕ Agregado: Biografía');
  }

  // ── 8. Agregar Eventos anteriores si no existe (después de Biografía) ──
  if (!form.getItems().find(i => i.getTitle() === 'Eventos anteriores')) {
    const bioIdx = form.getItems().findIndex(i => i.getTitle() === 'Biografía');
    const eventosItem = form.addTextItem()
      .setTitle('Eventos anteriores')
      .setHelpText('Eventos en los que participaste como speaker (opcional).');
    form.moveItem(eventosItem.getIndex(), bioIdx >= 0 ? bioIdx + 1 : form.getItems().length - 1);
    Logger.log('➕ Agregado: Eventos anteriores');
  }

  Logger.log('');
  Logger.log('✅ Campos actualizados. Orden final:');
  Logger.log('   Nombre → Tipo → Mail → Móvil (WhatsApp) → X → Instagram → LinkedIn → Empresa/Referencia → Ciudad(es) → Tema(s) → Notas → Biografía → Eventos anteriores');
  Logger.log('');
  Logger.log('👉 Ahora ejecutá updateTemaQuestion() para actualizar los temas con bajada.');
}

// ═══════════════════════════════════════════════════════════════
// PASO 2 — Actualizar pregunta de Tema(s) con bajada
// Convierte el campo a checkboxes (hasta 3), mostrando
// "Tema — Bajada" como etiqueta de cada opción.
// Ejecutar UNA VEZ (o cada vez que cambien los temas en el Sheet).
// ═══════════════════════════════════════════════════════════════
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

  Logger.log('✅ Tema(s) actualizado: ' + opciones.length + ' opciones (checkboxes, máx. 3).');
  opciones.slice(0, 3).forEach(o => Logger.log('   • ' + o));
}

// ═══════════════════════════════════════════════════════════════
// LEER TEMAS DESDE EL SHEET DEX
// Devuelve array de "Tema — Bajada" (o solo "Tema" si no hay bajada)
// Tema = col 4 | Bajada = col 12
// ═══════════════════════════════════════════════════════════════
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

// ═══════════════════════════════════════════════════════════════
// CREAR FORM DESDE CERO (solo si necesitás uno nuevo)
// ═══════════════════════════════════════════════════════════════
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
  form.addTextItem().setTitle('Tipo').setHelpText('Ej: Speaker, Panelista, Moderador');
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

// ── Helper: abre el form por ID hardcodeado ────────────────────
const SPEAKER_FORM_ID = '1x5OzFZXkSv2dqt7933fCiB3zQO7pYeFhlKPfSm-BCI0';

function _abrirForm() {
  try {
    return FormApp.openById(SPEAKER_FORM_ID);
  } catch(e) {
    Logger.log('❌ No se pudo abrir el form: ' + e.message);
    return null;
  }
}
