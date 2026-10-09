/**
 * ================================================================
 *  KW MARBELLA · MOTOR DE RECLUTAMIENTO DE AGENTES  ·  v1.0
 * ================================================================
 *  Módulo complementario al Dashboard Inmobiliario KW.
 *
 *  QUÉ HACE:
 *   1. CAPTA candidatos (agentes inmobiliarios) de fuentes públicas
 *      legítimas y los normaliza en una base única.
 *   2. PUNTÚA cada candidato por probabilidad de reclutamiento.
 *   3. SECUENCIA el seguimiento (Smart Plan) por el canal correcto:
 *      llamada → LinkedIn → WhatsApp → email, respetando consentimiento.
 *   4. REGISTRA todo para cumplir RGPD/LSSI (base legal, opt-out,
 *      retención, auditoría de toques).
 *
 *  INSTALACIÓN (3 pasos):
 *   A) Pega este archivo como nuevo .gs en el proyecto Apps Script.
 *   B) En tu onOpen() existente (archivo "gs", línea ~198) añade
 *      una sola línea antes del cierre:       recCrearMenu();
 *   C) Ejecuta recInicializarTodo() una vez desde el editor.
 *
 *  OJO: este archivo NO define onOpen() a propósito, para no chocar
 *  con el que ya tienes.
 * ================================================================
 */

// ============================================================
//  1. CONFIGURACIÓN
// ============================================================

const REC = {
  // --- Hojas ---
  H_CANDIDATOS:  'Rec_Candidatos',
  H_AGENCIAS:    'Rec_Agencias',
  H_TOQUES:      'Rec_Toques',
  H_PLAN:        'Rec_SmartPlan',
  H_SUPRESION:   'Rec_Supresion',
  H_ENTREVISTAS: 'Rec_Entrevistas',
  H_CONFIG:      'Rec_Config',
  H_RGPD:        'Rec_RGPD',

  // --- Identidad del Market Center ---
  MC_NOMBRE:   'Keller Williams Marbella',
  MC_CIUDAD:   'Marbella',
  // Rellena estos en la hoja Rec_Config (tienen prioridad sobre estos valores)
  TL_NOMBRE:   '[Team Leader]',
  TL_TELEFONO: '[+34 ...]',
  MC_EMAIL:    '[email del MC]',
  MC_WEB:      '[web del MC]',
  MC_DIRECCION:'[dirección oficina]',

  // --- Zonas objetivo (Costa del Sol occidental) ---
  ZONAS: [
    'Marbella', 'San Pedro de Alcántara', 'Nueva Andalucía', 'Puerto Banús',
    'Golden Mile Marbella', 'Benahavís', 'Estepona', 'Mijas', 'Fuengirola',
    'Sotogrande', 'Elviria', 'La Zagaleta', 'Guadalmina'
  ],

  // --- Términos de búsqueda para Google Places ---
  TERMINOS_PLACES: [
    'inmobiliaria', 'real estate agency', 'agencia inmobiliaria',
    'estate agent', 'luxury real estate', 'property consultant'
  ],

  // --- Modo de envío de WhatsApp ---
  //  'LINK' = genera enlaces wa.me que la TL pulsa y envía a mano.
  //           Coste 0, sin riesgo de política, sin necesidad de opt-in previo
  //           de plataforma. ES EL MODO RECOMENDADO PARA EMPEZAR.
  //  'API'  = WhatsApp Business Cloud API con plantillas aprobadas.
  //           SOLO para candidatos con Consentimiento_WhatsApp = SI.
  MODO_WHATSAPP: 'LINK',

  // --- ¿Permitir WhatsApp como primer toque si no contestan la llamada? ---
  // true  = más contactabilidad, riesgo legal/plataforma algo mayor
  // false = el primer toque escrito va por LinkedIn o SMS
  WA_EN_PRIMER_TOQUE: true,

  // --- Pesos del scoring (deben sumar 100) ---
  PESOS: {
    produccion:    30,  // nº de inmuebles activos y rango de precio
    zona:          15,  // trabaja en zona núcleo del MC
    perfil:        15,  // autónomo/independiente > agencia pequeña > gran marca
    idiomas:       10,  // multilingüe (clave en Marbella)
    experiencia:   10,  // 2-8 años = punto óptimo
    dolor:         10,  // señales de insatisfacción con su agencia
    accesibilidad:  5,  // tenemos teléfono directo
    red:            5   // referido o contacto en común
  },

  // --- Retención de datos (días) antes de purga automática ---
  RETENCION_DIAS: 365,

  // --- Límite de toques diarios para no saturar ni al equipo ni los canales ---
  MAX_TOQUES_DIA: 40,

  COLOR_CABECERA: '#b70000'
};

const REC_ESTADOS = [
  'Nuevo', 'Investigado', 'En secuencia', 'Contactado', 'Conversación activa',
  'Entrevista agendada', 'Entrevistado', 'Career Visioning', 'Oferta',
  'Incorporado', 'No ahora (nurture)', 'Descartado', 'Opt-out'
];

const REC_CANALES = ['Llamada', 'WhatsApp', 'LinkedIn', 'Email', 'SMS', 'Presencial', 'Referido'];

const REC_PERFILES = [
  'Agente en agencia independiente',
  'Agente en gran franquicia',
  'Agente autónomo / sin agencia',
  'Agente en otro MC KW',
  'Cambio de sector (hostelería/lujo/banca)',
  'Recién titulado / sin experiencia',
  'Team Leader / Broker',
  'Desconocido'
];

// ============================================================
//  2. MENÚ (llamar desde tu onOpen existente)
// ============================================================

function recCrearMenu() {
  SpreadsheetApp.getUi().createMenu('🎯 Reclutamiento')
    .addItem('🚀 Inicializar sistema de reclutamiento', 'recInicializarTodo')
    .addItem('📋 Panel diario de la Team Leader', 'recAbrirPanel')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🔍 Captar candidatos')
      .addItem('1. Mapear agencias de la zona (Google Places)', 'recImportarGooglePlaces')
      .addItem('2. Extraer agentes de las webs de agencias', 'recRastrearWebsAgencias')
      .addItem('3. Importar CSV (LinkedIn Sales Navigator)', 'recAbrirImportadorCSV')
      .addItem('4. Pegar ficha de portal (Idealista/Fotocasa)', 'recAbrirPegadoPortal')
      .addItem('5. Generar búsquedas de LinkedIn', 'recGenerarBooleanLinkedIn')
      .addItem('6. X-ray de LinkedIn vía Google (sin Sales Navigator)', 'recBuscarXRayGoogle'))
    .addSeparator()
    .addItem('⭐ Recalcular puntuaciones', 'recRecalcularScores')
    .addItem('🧹 Deduplicar base de candidatos', 'recDeduplicar')
    .addItem('📨 Generar toques de hoy', 'recGenerarToquesDelDia')
    .addSeparator()
    .addItem('📤 Exportar para CommandMC', 'recExportarCommandMC')
    .addItem('🔐 Gestionar opt-out / supresión', 'recAbrirSupresion')
    .addItem('🗑️ Purgar datos fuera de retención', 'recPurgarRetencion')
    .addSeparator()
    .addItem('⏰ Activar automatización diaria (7:00)', 'recInstalarTriggerDiario')
    .addItem('♻️ Recargar configuración', 'recRefrescarConfig')
    .addSeparator()
    .addItem('🔑 Configurar claves de API', 'recConfigurarClaves')
    .addItem('🚨 Migrar la clave de Gemini expuesta', 'recMigrarClaveGemini')
    .addToUi();
}

// ============================================================
//  3. INICIALIZACIÓN DE HOJAS
// ============================================================

function recInicializarTodo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const creadas = [];

  creadas.push(recCrearHoja_(ss, REC.H_CANDIDATOS, [
    'ID', 'Nombre', 'Apellidos', 'Teléfono', 'Email', 'LinkedIn', 'Instagram',
    'Agencia_Actual', 'Cargo', 'Zona', 'Idiomas', 'Perfil',
    'Inmuebles_Activos', 'Precio_Medio', 'Precio_Max', 'Años_Experiencia',
    'Fuente', 'URL_Fuente', 'Fecha_Captura',
    'Score', 'Temperatura', 'Estado', 'Responsable',
    'Plan_Activo', 'Paso_Plan', 'Último_Toque', 'Próximo_Toque', 'Nº_Toques',
    'Canal_Preferido', 'Idioma_Pref', 'Consentimiento_WA', 'Fecha_Consentimiento',
    'Base_Legal', 'Info_RGPD_Enviada', 'ID_CommandMC', 'Notas'
  ]));

  creadas.push(recCrearHoja_(ss, REC.H_AGENCIAS, [
    'ID', 'Agencia', 'Web', 'Teléfono', 'Email', 'Dirección', 'Zona',
    'Google_Rating', 'Google_Reviews', 'Agentes_Detectados', 'Modelo',
    'Competidor_Directo', 'Prioridad', 'Fuente', 'Fecha_Captura',
    'Web_Rastreada', 'Notas'
  ]));

  creadas.push(recCrearHoja_(ss, REC.H_TOQUES, [
    'ID_Toque', 'ID_Candidato', 'Nombre', 'Fecha', 'Hora', 'Canal', 'Paso_Plan',
    'Tipo', 'Estado_Envío', 'Resultado', 'Mensaje', 'Usuario', 'Notas'
  ]));

  creadas.push(recCrearHoja_(ss, REC.H_PLAN, [
    'Plan', 'Paso', 'Día_Offset', 'Canal', 'Tipo', 'Objetivo',
    'Asunto', 'Mensaje_ES', 'Mensaje_EN', 'Requiere_Consentimiento', 'Activo'
  ]));

  creadas.push(recCrearHoja_(ss, REC.H_SUPRESION, [
    'Identificador', 'Tipo', 'Nombre', 'Fecha', 'Motivo', 'Origen', 'Usuario'
  ]));

  creadas.push(recCrearHoja_(ss, REC.H_ENTREVISTAS, [
    'ID_Entrevista', 'ID_Candidato', 'Nombre', 'Fecha', 'Fase', 'Entrevistador',
    'Transacciones_12m', 'GCI_Estimado', 'Split_Actual', 'Dolor_Principal',
    'Objetivo_Personal', 'Motivación', 'Objeciones', 'Resultado',
    'Siguiente_Paso', 'Fecha_Siguiente', 'Notas'
  ]));

  creadas.push(recCrearHoja_(ss, REC.H_CONFIG, ['Clave', 'Valor', 'Descripción']));
  creadas.push(recCrearHoja_(ss, REC.H_RGPD, ['Campo', 'Contenido']));

  recSembrarConfig_(ss);
  recSembrarSmartPlan_(ss);
  recSembrarRGPD_(ss);
  recAplicarValidaciones_(ss);

  SpreadsheetApp.getUi().alert(
    '✅ Sistema de reclutamiento listo',
    'Hojas creadas/verificadas:\n\n' + creadas.join('\n') +
    '\n\nSIGUIENTE PASO:\n' +
    '1) Rellena Rec_Config (nombre de la TL, teléfono, zonas).\n' +
    '2) Menú 🎯 Reclutamiento → Configurar claves de API.\n' +
    '3) Menú → Captar candidatos → paso 1.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function recCrearHoja_(ss, nombre, headers) {
  let hoja = ss.getSheetByName(nombre);
  let estado = '• ' + nombre;
  if (!hoja) {
    hoja = ss.insertSheet(nombre);
    estado += ' (nueva)';
  } else {
    estado += ' (ya existía, cabeceras actualizadas)';
  }
  if (hoja.getMaxColumns() < headers.length) {
    hoja.insertColumnsAfter(hoja.getMaxColumns(), headers.length - hoja.getMaxColumns());
  }
  hoja.getRange(1, 1, 1, headers.length)
      .setValues([headers])
      .setBackground(REC.COLOR_CABECERA)
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setFontSize(10);
  hoja.setFrozenRows(1);
  if (hoja.getMaxColumns() > headers.length) {
    hoja.deleteColumns(headers.length + 1, hoja.getMaxColumns() - headers.length);
  }
  return estado;
}

function recSembrarConfig_(ss) {
  const hoja = ss.getSheetByName(REC.H_CONFIG);
  if (hoja.getLastRow() > 1) return;
  const filas = [
    ['TL_NOMBRE', '', 'Nombre de la Team Leader que firma los mensajes'],
    ['TL_TELEFONO', '', 'Teléfono de la TL en formato +34...'],
    ['MC_EMAIL', '', 'Email del Market Center'],
    ['MC_WEB', '', 'Web del Market Center'],
    ['MC_DIRECCION', '', 'Dirección de la oficina (para invitaciones)'],
    ['MODO_WHATSAPP', 'LINK', 'LINK = enlaces wa.me manuales | API = Cloud API con plantillas'],
    ['WA_EN_PRIMER_TOQUE', 'SI', 'SI/NO — permitir WhatsApp si no contestan la llamada'],
    ['MAX_TOQUES_DIA', '40', 'Tope de toques generados por día'],
    ['RETENCION_DIAS', '365', 'Días que conservamos un candidato sin avance'],
    ['ZONAS_EXTRA', '', 'Zonas adicionales separadas por coma'],
    ['RESPONSABLE_DEFECTO', '', 'Quién trabaja los candidatos nuevos por defecto']
  ];
  hoja.getRange(2, 1, filas.length, 3).setValues(filas);
  hoja.autoResizeColumns(1, 3);
}

let REC_CACHE_CONFIG = null;

function recLeerConfig_(clave, porDefecto) {
  if (REC_CACHE_CONFIG === null) {
    REC_CACHE_CONFIG = {};
    const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REC.H_CONFIG);
    if (hoja && hoja.getLastRow() >= 2) {
      hoja.getRange(2, 1, hoja.getLastRow() - 1, 2).getValues().forEach(f => {
        const k = String(f[0]).trim();
        if (k) REC_CACHE_CONFIG[k] = String(f[1]).trim();
      });
    }
  }
  const v = REC_CACHE_CONFIG[clave];
  return (v !== undefined && v !== '') ? v : porDefecto;
}

/** Invalida la caché tras editar Rec_Config a mano. */
function recRefrescarConfig() {
  REC_CACHE_CONFIG = null;
  SpreadsheetApp.getActiveSpreadsheet().toast('Configuración recargada', 'Reclutamiento', 3);
}

function recAplicarValidaciones_(ss) {
  const hc = ss.getSheetByName(REC.H_CANDIDATOS);
  const n = Math.max(hc.getMaxRows() - 1, 1);

  const regla = (lista) => SpreadsheetApp.newDataValidation()
      .requireValueInList(lista, true).setAllowInvalid(true).build();

  // Perfil = col 12, Estado = col 22, Consentimiento_WA = col 31
  hc.getRange(2, 12, n, 1).setDataValidation(regla(REC_PERFILES));
  hc.getRange(2, 22, n, 1).setDataValidation(regla(REC_ESTADOS));
  hc.getRange(2, 31, n, 1).setDataValidation(regla(['SI', 'NO', 'PENDIENTE']));
  hc.getRange(2, 29, n, 1).setDataValidation(regla(REC_CANALES));
  hc.getRange(2, 30, n, 1).setDataValidation(regla(['ES', 'EN']));

  // Formato condicional por temperatura (col 21)
  const rango = hc.getRange(2, 21, n, 1);
  const reglas = [
    ['A', '#16a34a'], ['B', '#f59e0b'], ['C', '#64748b'], ['D', '#ef4444']
  ].map(([t, c]) => SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(t).setBackground(c).setFontColor('#ffffff')
      .setRanges([rango]).build());
  hc.setConditionalFormatRules(reglas);
}

// ============================================================
//  4. CLAVES DE API (seguras, en PropertiesService)
// ============================================================

function recConfigurarClaves() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();

  const campos = [
    ['GEMINI_API_KEY',  'Clave de Google AI Studio (Gemini). Se usa para extraer agentes de webs y clasificar perfiles.'],
    ['PLACES_API_KEY',  'Clave de Google Maps Platform con Places API (New) activada.'],
    ['CSE_API_KEY',     'Clave de Google Custom Search API. Para el X-ray de LinkedIn sin Sales Navigator.'],
    ['CSE_CX',          'ID del motor de Programmable Search (cx), creado con "Buscar en toda la web" activada.'],
    ['WA_TOKEN',        'Token permanente de WhatsApp Business Cloud API. Solo si MODO_WHATSAPP = API.'],
    ['WA_PHONE_ID',     'Phone Number ID de WhatsApp Business. Solo si MODO_WHATSAPP = API.']
  ];

  for (const [clave, desc] of campos) {
    const actual = props.getProperty(clave);
    const pista = actual ? '\n\nActual: ' + actual.substring(0, 8) + '…(configurada)' : '\n\nActual: SIN CONFIGURAR';
    const r = ui.prompt('🔑 ' + clave, desc + pista + '\n\nPega el valor nuevo (vacío = dejar como está):', ui.ButtonSet.OK_CANCEL);
    if (r.getSelectedButton() !== ui.Button.OK) return;
    const v = r.getResponseText().trim();
    if (v) props.setProperty(clave, v);
  }
  ui.alert('✅ Claves guardadas', 'Están en Propiedades del Script, no en el código ni en el repositorio.', ui.ButtonSet.OK);
}

function recClave_(nombre) {
  const v = PropertiesService.getScriptProperties().getProperty(nombre);
  if (!v) throw new Error('Falta la clave ' + nombre + '. Menú 🎯 Reclutamiento → Configurar claves de API.');
  return v;
}

/**
 * MIGRACIÓN DE SEGURIDAD: mueve la clave de Gemini que está escrita
 * a pelo en el archivo "gs" (línea ~3565) a PropertiesService.
 * Ejecútala una vez y después borra la constante del código.
 */
function recMigrarClaveGemini() {
  const ui = SpreadsheetApp.getUi();
  let encontrada = null;
  try { encontrada = GEMINI_API_KEY; } catch (e) { /* ya no existe, perfecto */ }

  if (!encontrada) {
    ui.alert('Nada que migrar', 'No hay constante GEMINI_API_KEY en el código. Correcto.', ui.ButtonSet.OK);
    return;
  }
  PropertiesService.getScriptProperties().setProperty('GEMINI_API_KEY', encontrada);
  ui.alert('⚠️ Clave migrada — AHORA HAZ ESTO',
    'La clave ya está en PropertiesService.\n\n' +
    'PERO sigue expuesta en el historial de Git, así que:\n\n' +
    '1) Entra en aistudio.google.com → API Keys → BORRA esa clave.\n' +
    '2) Crea una clave nueva.\n' +
    '3) Menú 🎯 Reclutamiento → Configurar claves de API → pega la nueva.\n' +
    '4) En el archivo "gs", borra la línea:  const GEMINI_API_KEY = \'AIza...\';\n' +
    '5) Sustituye las llamadas por recLlamarGemini().',
    ui.ButtonSet.OK);
}

function recLlamarGemini(prompt, esperaJson) {
  const key = recClave_('GEMINI_API_KEY');
  const modelos = ['gemini-2.0-flash', 'gemini-flash-latest', 'gemini-2.0-flash-lite'];

  for (const modelo of modelos) {
    try {
      const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + modelo + ':generateContent?key=' + key;
      const payload = {
        contents: [{ parts: [{ text: prompt }] }],
        generationConfig: { temperature: 0.1, maxOutputTokens: 8192 }
      };
      if (esperaJson) payload.generationConfig.responseMimeType = 'application/json';

      const res = UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
      if (res.getResponseCode() !== 200) continue;
      const j = JSON.parse(res.getContentText());
      const texto = j.candidates && j.candidates[0] && j.candidates[0].content.parts[0].text;
      if (texto) return texto;
    } catch (e) { /* probamos el siguiente modelo */ }
  }
  throw new Error('Gemini no respondió con ningún modelo disponible.');
}

// ============================================================
//  5. FUENTE 1 — GOOGLE PLACES: mapa de agencias de la zona
// ============================================================

/**
 * Barre todas las zonas × términos y vuelca las agencias en Rec_Agencias.
 * Legalmente limpio: API oficial de pago, datos de empresa.
 * Rinde aprox. 250-500 agencias en la Costa del Sol occidental.
 */
function recImportarGooglePlaces() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(REC.H_AGENCIAS);
  if (!hoja) { ui.alert('Ejecuta primero recInicializarTodo()'); return; }

  let key;
  try { key = recClave_('PLACES_API_KEY'); }
  catch (e) { ui.alert('❌ ' + e.message); return; }

  const zonas = recZonas_();
  const conf = ui.alert('🗺️ Mapear agencias',
    'Voy a buscar en ' + zonas.length + ' zonas × ' + REC.TERMINOS_PLACES.length + ' términos.\n\n' +
    'Son ~' + (zonas.length * REC.TERMINOS_PLACES.length) + ' consultas a Places API (coste aprox. 0,03 €/consulta).\n\n' +
    '¿Continuar?', ui.ButtonSet.YES_NO);
  if (conf !== ui.Button.YES) return;

  const existentes = recIndiceColumna_(hoja, 1); // por place_id
  const nuevas = [];
  let consultas = 0;

  for (const zona of zonas) {
    for (const termino of REC.TERMINOS_PLACES) {
      const encontrados = recPlacesBuscar_(key, termino + ' en ' + zona + ', Málaga, España');
      consultas++;
      for (const p of encontrados) {
        const id = p.id || '';
        if (!id || existentes[id]) continue;
        existentes[id] = true;
        nuevas.push([
          id,
          (p.displayName && p.displayName.text) || '',
          p.websiteUri || '',
          p.nationalPhoneNumber || '',
          '',
          p.formattedAddress || '',
          zona,
          p.rating || '',
          p.userRatingCount || '',
          '',
          recClasificarModelo_((p.displayName && p.displayName.text) || ''),
          '',
          '',
          'Google Places',
          new Date(),
          'NO',
          ''
        ]);
      }
      Utilities.sleep(150); // cortesía con la API
    }
  }

  if (nuevas.length) {
    hoja.getRange(hoja.getLastRow() + 1, 1, nuevas.length, nuevas[0].length).setValues(nuevas);
  }
  ui.alert('✅ Mapa de agencias actualizado',
    'Consultas realizadas: ' + consultas + '\n' +
    'Agencias nuevas añadidas: ' + nuevas.length + '\n\n' +
    'SIGUIENTE: menú → Captar candidatos → "2. Extraer agentes de las webs".',
    ui.ButtonSet.OK);
}

function recPlacesBuscar_(key, consulta) {
  const resultados = [];
  let pageToken = null;
  let vueltas = 0;

  do {
    const payload = { textQuery: consulta, languageCode: 'es', maxResultCount: 20 };
    if (pageToken) payload.pageToken = pageToken;

    const res = UrlFetchApp.fetch('https://places.googleapis.com/v1/places:searchText', {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'X-Goog-Api-Key': key,
        'X-Goog-FieldMask': 'places.id,places.displayName,places.formattedAddress,places.websiteUri,places.nationalPhoneNumber,places.rating,places.userRatingCount,nextPageToken'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });

    if (res.getResponseCode() !== 200) {
      Logger.log('Places error ' + res.getResponseCode() + ': ' + res.getContentText().substring(0, 300));
      break;
    }
    const j = JSON.parse(res.getContentText());
    if (j.places) resultados.push.apply(resultados, j.places);
    pageToken = j.nextPageToken || null;
    vueltas++;
    if (pageToken) Utilities.sleep(400);
  } while (pageToken && vueltas < 3);

  return resultados;
}

function recClasificarModelo_(nombre) {
  const n = nombre.toLowerCase();
  const franquicias = ['keller williams', 'remax', 're/max', 'century 21', 'engel', 'völkers', 'volkers',
    'lucas fox', 'savills', 'sotheby', 'christie', 'berkshire', 'knight frank', 'barnes',
    'gilmar', 'tecnocasa', 'donpiso', 'look & find', 'redpiso', 'alfa inmobiliaria',
    'iad', 'exp realty', 'nexthome', 'coldwell'];
  for (const f of franquicias) if (n.indexOf(f) !== -1) return 'Gran franquicia';
  return 'Independiente';
}

function recZonas_() {
  const extra = recLeerConfig_('ZONAS_EXTRA', '');
  const lista = REC.ZONAS.slice();
  if (extra) extra.split(',').forEach(z => { const t = z.trim(); if (t) lista.push(t); });
  return lista;
}

// ============================================================
//  6. FUENTE 2 — WEBS DE AGENCIAS: extraer agentes del equipo
// ============================================================

/**
 * Para cada agencia con web, busca su página de equipo y extrae los
 * agentes individuales (nombre, cargo, email, teléfono, idiomas).
 *
 * Respeta robots.txt antes de descargar nada.
 * Usa Gemini para convertir el HTML en datos estructurados.
 *
 * Esta es la fuente MÁS RENTABLE y limpia: las agencias publican a su
 * equipo precisamente para que se les contacte profesionalmente.
 */
function recRastrearWebsAgencias() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hAg = ss.getSheetByName(REC.H_AGENCIAS);
  const hCa = ss.getSheetByName(REC.H_CANDIDATOS);
  if (!hAg || hAg.getLastRow() < 2) { ui.alert('No hay agencias. Ejecuta primero el paso 1.'); return; }

  const lote = parseInt(recLeerConfig_('LOTE_WEBS', '25'), 10);
  const conf = ui.alert('🕸️ Extraer agentes de webs',
    'Proceso por lotes de ' + lote + ' agencias (Apps Script tiene límite de 6 min por ejecución).\n\n' +
    'Se saltan las que ya estén marcadas como rastreadas y las que lo prohíban en robots.txt.\n\n' +
    '¿Continuar?', ui.ButtonSet.YES_NO);
  if (conf !== ui.Button.YES) return;

  const datos = hAg.getRange(2, 1, hAg.getLastRow() - 1, 17).getValues();
  const yaExisten = recIndiceCandidatos_(hCa);
  const nuevos = [];
  let procesadas = 0, bloqueadas = 0, sinEquipo = 0;
  const inicio = Date.now();

  for (let i = 0; i < datos.length; i++) {
    if (procesadas >= lote) break;
    if (Date.now() - inicio > 4.5 * 60 * 1000) break; // margen de seguridad

    const fila = datos[i];
    const web = String(fila[2]).trim();
    const agencia = String(fila[1]).trim();
    const zona = String(fila[6]).trim();
    const rastreada = String(fila[15]).trim().toUpperCase();

    if (!web || rastreada === 'SI' || rastreada === 'BLOQUEADA') continue;

    procesadas++;
    let marca = 'SI';

    try {
      const paginas = recBuscarPaginasEquipo_(web);
      if (paginas.bloqueado) { marca = 'BLOQUEADA'; bloqueadas++; }
      else if (!paginas.texto) { sinEquipo++; }
      else {
        const agentes = recExtraerAgentesConIA_(paginas.texto, agencia, paginas.url);
        for (const a of agentes) {
          const clave = recClaveDedupe_(a.nombre, a.telefono, a.email);
          if (!clave || yaExisten[clave]) continue;
          yaExisten[clave] = true;
          nuevos.push(recFilaCandidato_({
            nombre: a.nombre, apellidos: a.apellidos, telefono: a.telefono, email: a.email,
            linkedin: a.linkedin || '', agencia: agencia, cargo: a.cargo || '',
            zona: zona, idiomas: a.idiomas || '',
            perfil: recClasificarModelo_(agencia) === 'Gran franquicia'
                    ? 'Agente en gran franquicia' : 'Agente en agencia independiente',
            fuente: 'Web de agencia', urlFuente: paginas.url
          }));
        }
      }
    } catch (e) {
      Logger.log('Error en ' + web + ': ' + e.message);
      marca = 'ERROR';
    }
    hAg.getRange(i + 2, 16).setValue(marca);
  }

  if (nuevos.length) {
    hCa.getRange(hCa.getLastRow() + 1, 1, nuevos.length, nuevos[0].length).setValues(nuevos);
    recRecalcularScores(true);
  }

  ui.alert('✅ Lote completado',
    'Webs procesadas: ' + procesadas + '\n' +
    'Candidatos nuevos: ' + nuevos.length + '\n' +
    'Bloqueadas por robots.txt: ' + bloqueadas + '\n' +
    'Sin página de equipo: ' + sinEquipo + '\n\n' +
    (procesadas >= lote ? 'Quedan agencias pendientes: vuelve a ejecutar para el siguiente lote.' : 'No quedan agencias pendientes.'),
    ui.ButtonSet.OK);
}

/** Rutas habituales de página de equipo en webs inmobiliarias ES/EN. */
const REC_RUTAS_EQUIPO = [
  '/equipo', '/nuestro-equipo', '/team', '/our-team', '/meet-the-team',
  '/agentes', '/asesores', '/about-us', '/sobre-nosotros', '/nosotros',
  '/quienes-somos', '/staff', '/people', '/agents', '/es/equipo', '/en/team'
];

function recBuscarPaginasEquipo_(web) {
  const base = recNormalizarUrl_(web);
  if (!base) return { texto: '', url: '', bloqueado: false };

  const robots = recLeerRobots_(base);

  for (const ruta of REC_RUTAS_EQUIPO) {
    if (!recRobotsPermite_(robots, ruta)) continue;
    const url = base + ruta;
    const html = recDescargar_(url);
    if (!html) continue;
    const texto = recHtmlATexto_(html);
    // Heurística: una página de equipo real menciona varias personas y roles
    if (texto.length > 400 && /asesor|agente|agent|consultant|realtor|director|partner/i.test(texto)) {
      return { texto: texto.substring(0, 18000), url: url, bloqueado: false };
    }
  }

  // Fallback: la home, por si el equipo está ahí
  if (recRobotsPermite_(robots, '/')) {
    const html = recDescargar_(base + '/');
    if (html) {
      const texto = recHtmlATexto_(html);
      if (/equipo|our team|meet the team|asesores/i.test(texto)) {
        return { texto: texto.substring(0, 18000), url: base, bloqueado: false };
      }
    }
    return { texto: '', url: base, bloqueado: false };
  }
  return { texto: '', url: base, bloqueado: true };
}

function recNormalizarUrl_(web) {
  let w = String(web).trim();
  if (!w) return '';
  if (!/^https?:\/\//i.test(w)) w = 'https://' + w;
  const m = w.match(/^(https?:\/\/[^\/]+)/i);
  return m ? m[1] : '';
}

function recDescargar_(url) {
  try {
    const res = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true,
      validateHttpsCertificates: true,
      headers: { 'User-Agent': 'KW-Marbella-Recruiting/1.0 (contacto: ' + recLeerConfig_('MC_EMAIL', 'info') + ')' }
    });
    if (res.getResponseCode() !== 200) return '';
    return res.getContentText();
  } catch (e) { return ''; }
}

function recLeerRobots_(base) {
  const cache = CacheService.getScriptCache();
  const ck = 'robots_' + Utilities.base64EncodeWebSafe(base).substring(0, 40);
  const hit = cache.get(ck);
  if (hit !== null) return hit;
  const txt = recDescargar_(base + '/robots.txt') || '';
  cache.put(ck, txt, 21600); // 6 h
  return txt;
}

/** Parser simple de robots.txt para User-agent: * */
function recRobotsPermite_(robots, ruta) {
  if (!robots) return true;
  const lineas = robots.split(/\r?\n/);
  let aplica = false;
  const disallow = [], allow = [];

  for (let l of lineas) {
    l = l.replace(/#.*$/, '').trim();
    if (!l) continue;
    const m = l.match(/^([A-Za-z-]+)\s*:\s*(.*)$/);
    if (!m) continue;
    const campo = m[1].toLowerCase(), valor = m[2].trim();
    if (campo === 'user-agent') {
      aplica = (valor === '*');
    } else if (aplica && campo === 'disallow' && valor) {
      disallow.push(valor);
    } else if (aplica && campo === 'allow' && valor) {
      allow.push(valor);
    }
  }
  // Allow más específico gana sobre Disallow
  const coincide = (patron) => ruta.indexOf(patron.replace(/\*$/, '')) === 0;
  let bloqueoMax = -1, permisoMax = -1;
  disallow.forEach(p => { if (coincide(p)) bloqueoMax = Math.max(bloqueoMax, p.length); });
  allow.forEach(p => { if (coincide(p)) permisoMax = Math.max(permisoMax, p.length); });
  if (bloqueoMax === -1) return true;
  return permisoMax >= bloqueoMax;
}

function recHtmlATexto_(html) {
  return String(html)
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<noscript[\s\S]*?<\/noscript>/gi, ' ')
    .replace(/<!--[\s\S]*?-->/g, ' ')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/(p|div|li|h[1-6]|tr)>/gi, '\n')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&')
    .replace(/&quot;/g, '"').replace(/&#39;/g, "'")
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/[ \t]{2,}/g, ' ')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function recExtraerAgentesConIA_(texto, agencia, url) {
  const prompt =
    'Extrae las personas del equipo comercial de esta agencia inmobiliaria.\n' +
    'AGENCIA: ' + agencia + '\nURL: ' + url + '\n\n' +
    'Devuelve SOLO un array JSON. Un objeto por persona, con estas claves:\n' +
    '{"nombre":"", "apellidos":"", "cargo":"", "email":"", "telefono":"", "linkedin":"", "idiomas":""}\n\n' +
    'REGLAS ESTRICTAS:\n' +
    '- Incluye solo personas con rol comercial o directivo (asesor, agente, consultant, ' +
    'sales, director, partner, broker, Team Leader). EXCLUYE administración, marketing, ' +
    'recepción, contabilidad y jurídico.\n' +
    '- Si no encuentras un dato, deja la cadena vacía. NUNCA lo inventes.\n' +
    '- "idiomas": lista separada por comas con los idiomas que declare hablar.\n' +
    '- Teléfonos en formato internacional (+34...) si es posible.\n' +
    '- Si el texto no contiene personas identificables, devuelve [].\n\n' +
    'TEXTO DE LA PÁGINA:\n"""\n' + texto + '\n"""';

  try {
    const respuesta = recLlamarGemini(prompt, true);
    const limpio = respuesta.replace(/^```json\s*/i, '').replace(/```\s*$/, '').trim();
    const arr = JSON.parse(limpio);
    if (!Array.isArray(arr)) return [];
    return arr.filter(a => a && a.nombre && String(a.nombre).trim().length > 1);
  } catch (e) {
    Logger.log('IA no pudo extraer de ' + url + ': ' + e.message);
    return [];
  }
}

// ============================================================
//  7. ÍNDICES DE COLUMNA (1-based) — no toques el orden de la hoja
// ============================================================

const REC_COL = {
  ID: 1, NOMBRE: 2, APELLIDOS: 3, TELEFONO: 4, EMAIL: 5, LINKEDIN: 6, INSTAGRAM: 7,
  AGENCIA: 8, CARGO: 9, ZONA: 10, IDIOMAS: 11, PERFIL: 12,
  INMUEBLES: 13, PRECIO_MEDIO: 14, PRECIO_MAX: 15, EXPERIENCIA: 16,
  FUENTE: 17, URL_FUENTE: 18, FECHA_CAPTURA: 19,
  SCORE: 20, TEMPERATURA: 21, ESTADO: 22, RESPONSABLE: 23,
  PLAN: 24, PASO: 25, ULTIMO_TOQUE: 26, PROXIMO_TOQUE: 27, N_TOQUES: 28,
  CANAL_PREF: 29, IDIOMA_PREF: 30, CONSENT_WA: 31, FECHA_CONSENT: 32,
  BASE_LEGAL: 33, INFO_RGPD: 34, ID_CMC: 35, NOTAS: 36
};
const REC_N_COLS = 36;

function recFilaCandidato_(d) {
  const fila = new Array(REC_N_COLS).fill('');
  fila[REC_COL.ID - 1]            = d.id || recNuevoId_('C');
  fila[REC_COL.NOMBRE - 1]        = d.nombre || '';
  fila[REC_COL.APELLIDOS - 1]     = d.apellidos || '';
  fila[REC_COL.TELEFONO - 1]      = recNormalizarTelefono_(d.telefono);
  fila[REC_COL.EMAIL - 1]         = String(d.email || '').trim().toLowerCase();
  fila[REC_COL.LINKEDIN - 1]      = d.linkedin || '';
  fila[REC_COL.INSTAGRAM - 1]     = d.instagram || '';
  fila[REC_COL.AGENCIA - 1]       = d.agencia || '';
  fila[REC_COL.CARGO - 1]         = d.cargo || '';
  fila[REC_COL.ZONA - 1]          = d.zona || '';
  fila[REC_COL.IDIOMAS - 1]       = d.idiomas || '';
  fila[REC_COL.PERFIL - 1]        = d.perfil || 'Desconocido';
  fila[REC_COL.INMUEBLES - 1]     = d.inmuebles || '';
  fila[REC_COL.PRECIO_MEDIO - 1]  = d.precioMedio || '';
  fila[REC_COL.PRECIO_MAX - 1]    = d.precioMax || '';
  fila[REC_COL.EXPERIENCIA - 1]   = d.experiencia || '';
  fila[REC_COL.FUENTE - 1]        = d.fuente || 'Manual';
  fila[REC_COL.URL_FUENTE - 1]    = d.urlFuente || '';
  fila[REC_COL.FECHA_CAPTURA - 1] = new Date();
  fila[REC_COL.SCORE - 1]         = 0;
  fila[REC_COL.TEMPERATURA - 1]   = '';
  fila[REC_COL.ESTADO - 1]        = 'Nuevo';
  fila[REC_COL.RESPONSABLE - 1]   = d.responsable || recLeerConfig_('RESPONSABLE_DEFECTO', '');
  fila[REC_COL.N_TOQUES - 1]      = 0;
  fila[REC_COL.IDIOMA_PREF - 1]   = recIdiomaProbable_(d);
  fila[REC_COL.CONSENT_WA - 1]    = 'PENDIENTE';
  fila[REC_COL.BASE_LEGAL - 1]    = 'Interés legítimo (art. 6.1.f RGPD) — contacto profesional publicado';
  fila[REC_COL.INFO_RGPD - 1]     = 'NO';
  fila[REC_COL.NOTAS - 1]         = d.notas || '';
  return fila;
}

function recNuevoId_(prefijo) {
  return prefijo + '-' + Utilities.getUuid().substring(0, 8).toUpperCase();
}

/** Normaliza a +34XXXXXXXXX cuando se puede; respeta otros prefijos internacionales. */
function recNormalizarTelefono_(tel) {
  let t = String(tel || '').replace(/[^\d+]/g, '');
  if (!t) return '';
  if (t.indexOf('00') === 0) t = '+' + t.substring(2);
  if (t.indexOf('+') !== 0) {
    if (t.length === 9 && /^[679]/.test(t)) t = '+34' + t;   // móvil/fijo español
    else if (t.length === 11 && t.indexOf('34') === 0) t = '+' + t;
    else t = '+' + t;
  }
  return t.length >= 9 ? t : '';
}

function recIdiomaProbable_(d) {
  const idiomas = String(d.idiomas || '').toLowerCase();
  const nombre = String((d.nombre || '') + ' ' + (d.apellidos || ''));
  if (/espa|spanish|castellano/.test(idiomas)) return 'ES';
  if (/english|ingl/.test(idiomas) && !/espa|spanish/.test(idiomas)) return 'EN';
  // Sin datos: en Marbella, apellido no hispano → probablemente EN
  if (!/[áéíóúñ]/i.test(nombre) && /\b(van|von|de la|smith|jones|brown|johansson|nielsen|ivanov|petrov|dubois)\b/i.test(nombre)) return 'EN';
  return 'ES';
}

function recClaveDedupe_(nombre, telefono, email) {
  const tel = recNormalizarTelefono_(telefono);
  if (tel) return 'T:' + tel;
  const em = String(email || '').trim().toLowerCase();
  if (em) return 'E:' + em;
  const n = String(nombre || '').toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/[^a-z ]/g, '').replace(/\s+/g, ' ').trim();
  return n ? 'N:' + n : '';
}

function recIndiceColumna_(hoja, col) {
  const idx = {};
  if (hoja.getLastRow() < 2) return idx;
  hoja.getRange(2, col, hoja.getLastRow() - 1, 1).getValues()
    .forEach(f => { const v = String(f[0]).trim(); if (v) idx[v] = true; });
  return idx;
}

function recIndiceCandidatos_(hoja) {
  const idx = {};
  if (hoja.getLastRow() < 2) return idx;
  const datos = hoja.getRange(2, 1, hoja.getLastRow() - 1, REC_N_COLS).getValues();
  datos.forEach(f => {
    const c = recClaveDedupe_(
      f[REC_COL.NOMBRE - 1] + ' ' + f[REC_COL.APELLIDOS - 1],
      f[REC_COL.TELEFONO - 1], f[REC_COL.EMAIL - 1]);
    if (c) idx[c] = true;
  });
  return idx;
}

// ============================================================
//  8. FUENTE 3 — IMPORTAR CSV (LinkedIn Sales Navigator, etc.)
// ============================================================

function recAbrirImportadorCSV() {
  const html = HtmlService.createHtmlOutput(recHtmlImportador_())
    .setWidth(620).setHeight(560);
  SpreadsheetApp.getUi().showModalDialog(html, '📥 Importar candidatos desde CSV o texto');
}

/**
 * Recibe el contenido pegado y lo mete en Rec_Candidatos.
 * Acepta CSV con cabecera (detecta columnas por nombre) o texto libre
 * (en cuyo caso lo interpreta con IA).
 */
function recImportarTexto(contenido, fuente, modo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hCa = ss.getSheetByName(REC.H_CANDIDATOS);
  const yaExisten = recIndiceCandidatos_(hCa);
  let registros = [];

  if (modo === 'csv') {
    registros = recParsearCSVCandidatos_(contenido);
  } else {
    registros = recParsearTextoConIA_(contenido, fuente);
  }

  const nuevos = [], duplicados = [];
  for (const r of registros) {
    const clave = recClaveDedupe_(r.nombre + ' ' + (r.apellidos || ''), r.telefono, r.email);
    if (!clave) continue;
    if (yaExisten[clave]) { duplicados.push(r.nombre); continue; }
    yaExisten[clave] = true;
    r.fuente = r.fuente || fuente;
    nuevos.push(recFilaCandidato_(r));
  }

  if (nuevos.length) {
    hCa.getRange(hCa.getLastRow() + 1, 1, nuevos.length, REC_N_COLS).setValues(nuevos);
    recRecalcularScores(true);
  }
  return {
    detectados: registros.length,
    nuevos: nuevos.length,
    duplicados: duplicados.length
  };
}

function recParsearCSVCandidatos_(csv) {
  let filas;
  try { filas = Utilities.parseCsv(csv); }
  catch (e) { try { filas = Utilities.parseCsv(csv, ';'); } catch (e2) { return []; } }
  if (!filas || filas.length < 2) return [];

  const cab = filas[0].map(c => String(c).toLowerCase().trim());
  const buscar = (alternativas) => {
    for (const a of alternativas) {
      const i = cab.findIndex(c => c === a);
      if (i !== -1) return i;
    }
    for (const a of alternativas) {
      const i = cab.findIndex(c => c.indexOf(a) !== -1);
      if (i !== -1) return i;
    }
    return -1;
  };

  const iNombre  = buscar(['first name', 'nombre', 'firstname', 'name', 'full name']);
  const iApell   = buscar(['last name', 'apellidos', 'lastname', 'surname']);
  const iTel     = buscar(['phone', 'teléfono', 'telefono', 'mobile', 'móvil', 'movil', 'tel']);
  const iMail    = buscar(['email', 'e-mail', 'correo', 'email address']);
  const iEmpresa = buscar(['company', 'empresa', 'agencia', 'current company', 'organization']);
  const iCargo   = buscar(['title', 'cargo', 'position', 'job title', 'puesto']);
  const iUrl     = buscar(['linkedin', 'profile url', 'url', 'person linkedin url']);
  const iZona    = buscar(['location', 'ubicación', 'ubicacion', 'zona', 'city', 'geo']);
  const iIdiomas = buscar(['languages', 'idiomas']);

  const out = [];
  for (let i = 1; i < filas.length; i++) {
    const f = filas[i];
    if (!f || f.length === 0) continue;
    const g = (idx) => idx >= 0 && idx < f.length ? String(f[idx]).trim() : '';

    let nombre = g(iNombre), apellidos = g(iApell);
    if (nombre && !apellidos && nombre.indexOf(' ') !== -1) {
      const p = nombre.split(/\s+/);
      nombre = p.shift();
      apellidos = p.join(' ');
    }
    if (!nombre) continue;

    out.push({
      nombre: nombre, apellidos: apellidos,
      telefono: g(iTel), email: g(iMail),
      agencia: g(iEmpresa), cargo: g(iCargo),
      linkedin: g(iUrl), zona: g(iZona), idiomas: g(iIdiomas),
      perfil: recPerfilDesdeCargo_(g(iCargo), g(iEmpresa)),
      urlFuente: g(iUrl)
    });
  }
  return out;
}

function recPerfilDesdeCargo_(cargo, empresa) {
  const c = String(cargo).toLowerCase();
  const e = String(empresa).toLowerCase();
  if (/team leader|broker|ceo|founder|fundador|director general|managing/.test(c)) return 'Team Leader / Broker';
  if (/keller williams|kw /.test(e)) return 'Agente en otro MC KW';
  if (/autónom|autonom|freelance|independiente|self.?employed|personal brand/.test(c + ' ' + e)) return 'Agente autónomo / sin agencia';
  if (/hotel|restaur|yacht|luxury retail|private bank|banca privada|concierge|golf/.test(c + ' ' + e)) return 'Cambio de sector (hostelería/lujo/banca)';
  if (recClasificarModelo_(empresa) === 'Gran franquicia') return 'Agente en gran franquicia';
  if (/agent|asesor|realtor|consultant|comercial|sales/.test(c)) return 'Agente en agencia independiente';
  return 'Desconocido';
}

function recParsearTextoConIA_(texto, fuente) {
  const prompt =
    'Extrae los agentes inmobiliarios que aparezcan en este texto.\n' +
    'ORIGEN: ' + fuente + '\n\n' +
    'Devuelve SOLO un array JSON, un objeto por persona:\n' +
    '{"nombre":"","apellidos":"","telefono":"","email":"","agencia":"","cargo":"",' +
    '"zona":"","idiomas":"","inmuebles":"","precioMedio":"","precioMax":""}\n\n' +
    'REGLAS:\n' +
    '- "inmuebles" = número de propiedades que tiene publicadas, si aparece.\n' +
    '- "precioMedio"/"precioMax" = en euros, solo dígitos, si se puede deducir.\n' +
    '- Deja vacío lo que no encuentres. NO INVENTES NADA.\n' +
    '- Si no hay personas, devuelve [].\n\n' +
    'TEXTO:\n"""\n' + String(texto).substring(0, 25000) + '\n"""';
  try {
    const r = recLlamarGemini(prompt, true);
    const arr = JSON.parse(r.replace(/^```json\s*/i, '').replace(/```\s*$/, '').trim());
    if (!Array.isArray(arr)) return [];
    return arr.filter(a => a && a.nombre).map(a => {
      a.perfil = recPerfilDesdeCargo_(a.cargo || '', a.agencia || '');
      a.fuente = fuente;
      return a;
    });
  } catch (e) {
    throw new Error('No he podido interpretar el texto: ' + e.message);
  }
}

// ============================================================
//  9. FUENTE 4 — PEGADO DESDE PORTALES (Idealista, Fotocasa…)
// ============================================================

/**
 * Los portales prohíben el rastreo automático en sus condiciones y en
 * robots.txt. Este flujo es DELIBERADAMENTE semiautomático: la persona
 * abre la ficha de la agencia en el navegador, copia (Ctrl+A, Ctrl+C) y
 * pega aquí. La IA estructura los agentes y sus métricas de producción.
 *
 * Así conseguimos el dato de oro (nº de inmuebles y rango de precio =
 * producción real) sin rastrear el portal.
 */
function recAbrirPegadoPortal() {
  const html = HtmlService.createHtmlOutput(recHtmlPegadoPortal_())
    .setWidth(640).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '📋 Pegar ficha de portal inmobiliario');
}

// ============================================================
//  FUENTE 6 — X-RAY DE GOOGLE: LinkedIn sin Sales Navigator
// ============================================================

/**
 * Localiza perfiles de LinkedIn usando la API de búsqueda de Google
 * (Programmable Search / Custom Search JSON API) con consultas `site:`.
 *
 * POR QUÉ ASÍ Y NO CON UN SCRAPER:
 *   • No accedemos a LinkedIn: consultamos el índice público de Google con
 *     su propia API oficial. Es exactamente lo que verías buscando a mano.
 *   • Sales Navigator cuesta ~100 €/mes. Esto son 100 consultas gratis al día
 *     y 5 $ por cada 1.000 adicionales.
 *   • Los scrapers de LinkedIn incumplen sus condiciones y la AEPD ya ha
 *     sancionado el uso de datos de perfiles públicos para contacto no
 *     consentido. Esta vía no toca LinkedIn.
 *
 * QUÉ TE DA:  nombre, cargo, agencia, zona y URL del perfil.
 * QUÉ NO TE DA: teléfono ni email. Eso lo cruzas con la web de su agencia
 *               (fuente 2) o con la ficha del portal (fuente 4).
 *
 * REQUISITOS (una vez):
 *   1. console.cloud.google.com → habilita "Custom Search API" → crea clave.
 *   2. programmablesearchengine.google.com → crea un motor con la opción
 *      "Buscar en toda la web" ACTIVADA. Copia el ID del motor (cx).
 *   3. Menú → Configurar claves de API → CSE_API_KEY y CSE_CX.
 */
function recBuscarXRayGoogle() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hCa = ss.getSheetByName(REC.H_CANDIDATOS);
  if (!hCa) { ui.alert('Ejecuta primero recInicializarTodo()'); return; }

  let key, cx;
  try { key = recClave_('CSE_API_KEY'); cx = recClave_('CSE_CX'); }
  catch (e) {
    ui.alert('❌ Falta configuración',
      e.message + '\n\nCómo se consigue:\n' +
      '1. console.cloud.google.com → habilita "Custom Search API" → crea una clave.\n' +
      '2. programmablesearchengine.google.com → crea un motor con\n' +
      '   "Buscar en toda la web" ACTIVADA → copia el ID (cx).\n' +
      '3. Menú → Configurar claves de API.\n\n' +
      'Las primeras 100 consultas de cada día son gratis.',
      ui.ButtonSet.OK);
    return;
  }

  const consultas = recConsultasXRay_();
  const conf = ui.alert('🔎 X-ray de LinkedIn vía Google',
    consultas.length + ' consultas × hasta 30 resultados cada una.\n\n' +
    'Consume ' + (consultas.length * 3) + ' llamadas de tu cuota diaria (100 gratis/día).\n\n' +
    'Saca nombre, cargo, agencia y URL del perfil. NO saca teléfono:\n' +
    'eso se cruza después con la web de la agencia o la ficha del portal.\n\n' +
    '¿Continuar?', ui.ButtonSet.YES_NO);
  if (conf !== ui.Button.YES) return;

  const yaExisten = recIndiceCandidatos_(hCa);
  const urlsVistas = {};
  hCa.getLastRow() > 1 && hCa.getRange(2, REC_COL.LINKEDIN, hCa.getLastRow() - 1, 1)
    .getValues().forEach(f => { const v = recNormalizarLinkedIn_(f[0]); if (v) urlsVistas[v] = true; });

  const nuevos = [];
  let llamadas = 0, errores = 0;
  const inicio = Date.now();

  for (const c of consultas) {
    if (Date.now() - inicio > 4.5 * 60 * 1000) break;   // margen del límite de 6 min

    for (let pagina = 0; pagina < 3; pagina++) {
      const res = recCSEConsultar_(key, cx, c.query, pagina * 10 + 1);
      llamadas++;
      if (res.error) { errores++; break; }
      if (!res.items.length) break;

      for (const item of res.items) {
        const perfil = recParsearPerfilLinkedIn_(item);
        if (!perfil) continue;

        const url = recNormalizarLinkedIn_(perfil.linkedin);
        if (url && urlsVistas[url]) continue;
        const clave = recClaveDedupe_(perfil.nombre + ' ' + perfil.apellidos, '', '');
        if (clave && yaExisten[clave]) continue;
        if (url) urlsVistas[url] = true;
        if (clave) yaExisten[clave] = true;

        perfil.fuente = 'Google X-Ray (LinkedIn)';
        perfil.urlFuente = perfil.linkedin;
        perfil.zona = perfil.zona || c.zona || '';
        perfil.perfil = recPerfilDesdeCargo_(perfil.cargo, perfil.agencia);
        perfil.notas = c.segmento;
        nuevos.push(recFilaCandidato_(perfil));
      }
      if (res.items.length < 10) break;
      Utilities.sleep(200);
    }
  }

  if (nuevos.length) {
    hCa.getRange(hCa.getLastRow() + 1, 1, nuevos.length, REC_N_COLS).setValues(nuevos);
    recRecalcularScores(true);
  }

  ui.alert('✅ X-ray completado',
    'Llamadas a la API: ' + llamadas + (errores ? ' (' + errores + ' con error)' : '') + '\n' +
    'Perfiles nuevos añadidos: ' + nuevos.length + '\n\n' +
    'Ninguno tiene teléfono todavía. Para conseguirlo:\n' +
    '• Si su agencia ya está en Rec_Agencias → ejecuta "Extraer agentes de las webs".\n' +
    '• Si no → pega su ficha del portal (opción 4).\n' +
    '• O escríbele por LinkedIn, que para eso tienes la URL.',
    ui.ButtonSet.OK);
}

function recCSEConsultar_(key, cx, query, start) {
  const url = 'https://www.googleapis.com/customsearch/v1'
    + '?key=' + encodeURIComponent(key)
    + '&cx=' + encodeURIComponent(cx)
    + '&q=' + encodeURIComponent(query)
    + '&num=10&start=' + start
    + '&hl=es&gl=es';
  try {
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) {
      Logger.log('CSE ' + res.getResponseCode() + ': ' + res.getContentText().substring(0, 250));
      return { items: [], error: true };
    }
    const j = JSON.parse(res.getContentText());
    return { items: j.items || [], error: false };
  } catch (e) {
    Logger.log('CSE excepción: ' + e.message);
    return { items: [], error: true };
  }
}

function recNormalizarLinkedIn_(url) {
  const m = String(url || '').match(/linkedin\.com\/in\/([^\/?#\s]+)/i);
  return m ? ('linkedin.com/in/' + m[1].toLowerCase()) : '';
}

/**
 * LinkedIn titula sus páginas de forma muy regular:
 *   "Ana García - Asesora Inmobiliaria - Panorama Properties | LinkedIn"
 *   "Lars Nilsson - Marbella, Andalucía, España | Perfil profesional | LinkedIn"
 * Y el fragmento suele traer "Ubicación: X · Experiencia: Y".
 */
function recParsearPerfilLinkedIn_(item) {
  const link = String(item.link || '');
  if (!/linkedin\.com\/in\//i.test(link)) return null;

  let titulo = String(item.title || '')
    .replace(/\s*\|\s*LinkedIn\s*$/i, '')
    .replace(/\s*\|\s*Perfil profesional\s*$/i, '')
    .replace(/\s*\|\s*Professional Profile\s*$/i, '')
    .trim();
  if (!titulo) return null;

  const partes = titulo.split(/\s+[-–—]\s+/).map(p => p.trim()).filter(Boolean);
  const nombreCompleto = partes.shift() || '';
  if (nombreCompleto.length < 3 || nombreCompleto.length > 60) return null;
  if (/^(perfiles|profiles|\d+\+?)/i.test(nombreCompleto)) return null;   // páginas de listado

  const esUbicacion = (t) =>
    /españa|spain|andaluc|málaga|malaga|provincia/i.test(t) ||
    recZonas_().some(z => t.toLowerCase().indexOf(z.toLowerCase()) !== -1);

  let cargo = '', agencia = '', zona = '';
  for (const p of partes) {
    if (!zona && esUbicacion(p)) { zona = p; continue; }
    if (!cargo) { cargo = p; continue; }
    if (!agencia) { agencia = p; }
  }

  // El fragmento completa lo que falte
  const frag = String(item.snippet || '').replace(/\s+/g, ' ');
  if (!zona) {
    const m = frag.match(/Ubicaci[oó]n:\s*([^·•|]{3,60})/i) || frag.match(/Location:\s*([^·•|]{3,60})/i);
    if (m) zona = m[1].trim();
  }
  let experiencia = '';
  const mExp = frag.match(/Experiencia:\s*([^·•|]{2,40})/i) || frag.match(/Experience:\s*([^·•|]{2,40})/i);
  if (mExp && !agencia) agencia = mExp[1].trim();
  const mAnios = frag.match(/(\d{1,2})\s*a[ñn]os?\s+(?:de\s+)?experiencia/i);
  if (mAnios) experiencia = mAnios[1];

  const trozos = nombreCompleto.split(/\s+/);
  return {
    nombre: trozos.shift(),
    apellidos: trozos.join(' '),
    cargo: cargo.substring(0, 90),
    agencia: agencia.substring(0, 90),
    zona: zona.substring(0, 60),
    experiencia: experiencia,
    linkedin: link.split('?')[0],
    idiomas: ''
  };
}

/** Las consultas X-ray que se lanzan contra la API. */
function recConsultasXRay_() {
  const base = 'site:linkedin.com/in';
  const zonas = ['Marbella', 'Estepona', 'Benahavís', 'San Pedro de Alcántara',
                 'Nueva Andalucía', 'Mijas', 'Sotogrande'];
  const out = [];

  zonas.forEach(z => {
    out.push({ segmento: 'Agentes en activo', zona: z,
      query: base + ' ("asesor inmobiliario" OR "agente inmobiliario" OR "real estate agent" OR "property consultant") "' + z + '"' });
  });

  out.push({ segmento: 'Autónomos / marca propia', zona: '',
    query: base + ' ("asesor inmobiliario" OR "real estate agent") ("autónomo" OR "freelance" OR "independiente" OR "self-employed") ("Marbella" OR "Costa del Sol")' });

  out.push({ segmento: 'Competencia de lujo', zona: 'Marbella',
    query: base + ' ("Engel" OR "Lucas Fox" OR "Savills" OR "Sotheby" OR "Knight Frank" OR "Panorama" OR "Terra Meridiana") ("Marbella" OR "Costa del Sol")' });

  out.push({ segmento: 'Grandes franquicias', zona: '',
    query: base + ' ("RE/MAX" OR "Century 21" OR "iad" OR "eXp" OR "Tecnocasa") ("Marbella" OR "Estepona" OR "Mijas")' });

  out.push({ segmento: 'Cambio de sector (lujo/hostelería)', zona: 'Marbella',
    query: base + ' ("guest relations" OR "concierge" OR "private banker" OR "yacht broker" OR "luxury retail" OR "club manager") ("Marbella" OR "Puerto Banús" OR "Sotogrande")' });

  out.push({ segmento: 'Idioma de mercado comprador', zona: '',
    query: base + ' ("real estate" OR "property") "Marbella" ("Swedish" OR "Norwegian" OR "Dutch" OR "German" OR "Russian" OR "Arabic")' });

  return out;
}

// ============================================================
//  10. FUENTE 5 — GENERADOR DE BÚSQUEDAS DE LINKEDIN
// ============================================================

function recGenerarBooleanLinkedIn() {
  const zonas = ['Marbella', 'Estepona', 'Benahavís', 'San Pedro de Alcántara', 'Mijas', 'Sotogrande'];
  const cargosES = '"asesor inmobiliario" OR "agente inmobiliario" OR "consultor inmobiliario" OR "asesora inmobiliaria"';
  const cargosEN = '"real estate agent" OR "real estate consultant" OR "property consultant" OR "sales advisor" OR "estate agent"';

  const bloques = [];

  bloques.push({
    titulo: '1 · Agentes en activo (bilingüe, núcleo del mercado)',
    porque: 'El grueso de la base. Busca por cargo en ES y EN porque en Marbella conviven las dos nomenclaturas.',
    query: '(' + cargosES + ' OR ' + cargosEN + ') AND (' + zonas.map(z => '"' + z + '"').join(' OR ') + ')'
  });

  bloques.push({
    titulo: '2 · Autónomos y marca personal (máxima prioridad)',
    porque: 'Ya venden solos, sin estructura ni marca. Es el perfil al que la propuesta de KW le cambia la vida: se quedan con mucho más y dejan de estar solos.',
    query: '(' + cargosES + ' OR ' + cargosEN + ') AND ("autónomo" OR "freelance" OR "independiente" OR "self-employed" OR "own business" OR "personal brand") AND ("Marbella" OR "Costa del Sol" OR "Estepona")'
  });

  bloques.push({
    titulo: '3 · Agentes de competencia directa (lujo)',
    porque: 'Producen ticket alto. Van a comparar split y herramientas. Aborda por modelo económico, no por marca.',
    query: '(' + cargosES + ' OR ' + cargosEN + ') AND ("Engel & Völkers" OR "Lucas Fox" OR "Savills" OR "Christie\'s" OR "Sotheby\'s" OR "Knight Frank" OR "Barnes" OR "Panorama" OR "Diana Morales" OR "DM Properties" OR "Terra Meridiana" OR "Gilmar") AND ("Marbella" OR "Costa del Sol")'
  });

  bloques.push({
    titulo: '4 · Grandes franquicias (acostumbrados al modelo de red)',
    porque: 'Ya entienden franquicia, formación y CRM. La conversación es corta: se compara split, propiedad de la cartera y tope de comisión.',
    query: '(' + cargosES + ' OR ' + cargosEN + ') AND ("RE/MAX" OR "Century 21" OR "iad" OR "eXp" OR "Tecnocasa" OR "Alfa Inmobiliaria" OR "Look & Find" OR "Redpiso") AND ("Marbella" OR "Estepona" OR "Mijas" OR "Fuengirola")'
  });

  bloques.push({
    titulo: '5 · Cambio de sector: lujo y hostelería premium',
    porque: 'Los mejores agentes de Marbella no vienen del sector. Vienen de tratar a diario con el cliente que compra aquí. Traen agenda y trato, les falta el técnico, que se enseña.',
    query: '("sales manager" OR "director comercial" OR "guest relations" OR "concierge" OR "private banker" OR "banca privada" OR "yacht broker" OR "luxury retail" OR "club manager" OR "golf membership") AND ("Marbella" OR "Puerto Banús" OR "Sotogrande" OR "Benahavís")'
  });

  bloques.push({
    titulo: '6 · Mercados de origen del comprador (idioma = ventaja)',
    porque: 'Un agente que habla sueco, neerlandés, alemán, árabe o ruso accede a compradores que el resto no puede atender. Es el argumento de captación más fuerte que tienes.',
    query: '(' + cargosEN + ') AND ("Marbella" OR "Costa del Sol") AND ("Swedish" OR "Norwegian" OR "Danish" OR "Dutch" OR "German" OR "Russian" OR "Arabic" OR "Polish" OR "Finnish")'
  });

  bloques.push({
    titulo: '7 · Señal de movimiento reciente (ventana de oportunidad)',
    porque: 'Alguien que acaba de cambiar de agencia o lleva poco en una grande está en el momento de máxima apertura. Filtra en Sales Navigator por "changed jobs in last 90 days".',
    query: '(' + cargosES + ' OR ' + cargosEN + ') AND ("Marbella" OR "Estepona") — ' +
           'añade en Sales Navigator los filtros: Changed jobs = Last 90 days; Years in current company = Less than 1 year'
  });

  bloques.push({
    titulo: '8 · Agentes con producción visible (prueba social)',
    porque: 'Si publican sus propios cierres, son productores y les mueve el reconocimiento. Busca en el contenido, no en el perfil.',
    query: 'Búsqueda de CONTENIDO (no de personas): "vendido" OR "sold" OR "nueva exclusiva" OR "just listed" OR "cerramos" AND "Marbella" — ' +
           'filtra por publicaciones del último mes y mira quién escribe'
  });

  // Volcar a una hoja
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName('Rec_Busquedas_LinkedIn');
  if (hoja) ss.deleteSheet(hoja);
  hoja = ss.insertSheet('Rec_Busquedas_LinkedIn');

  hoja.getRange(1, 1, 1, 3).setValues([['Segmento', 'Por qué este segmento', 'Cadena de búsqueda (pegar en LinkedIn)']])
    .setBackground(REC.COLOR_CABECERA).setFontColor('#ffffff').setFontWeight('bold');
  hoja.setFrozenRows(1);

  const filas = bloques.map(b => [b.titulo, b.porque, b.query]);
  hoja.getRange(2, 1, filas.length, 3).setValues(filas);
  hoja.setColumnWidth(1, 260).setColumnWidth(2, 380).setColumnWidth(3, 620);
  hoja.getRange(2, 1, filas.length, 3).setWrap(true).setVerticalAlignment('top');

  // --- Bloque 2: X-ray de Google, gratis y sin Sales Navigator ---
  const filaX = filas.length + 3;
  hoja.getRange(filaX, 1).setValue('X-RAY DE GOOGLE · sin Sales Navigator, sin coste')
    .setFontWeight('bold').setFontSize(12).setFontColor('#b70000');
  hoja.getRange(filaX + 1, 1, 1, 3)
    .setValues([['Segmento', 'Por qué', 'Pegar en google.com (no en LinkedIn)']])
    .setBackground('#334155').setFontColor('#ffffff').setFontWeight('bold');

  const filasX = recConsultasXRay_().map(c => [
    c.segmento + (c.zona ? ' · ' + c.zona : ''),
    'Busca en el índice público de Google, no dentro de LinkedIn. Sin licencia y sin incumplir condiciones.',
    c.query
  ]);
  hoja.getRange(filaX + 2, 1, filasX.length, 3).setValues(filasX)
    .setWrap(true).setVerticalAlignment('top');

  hoja.getRange(filaX + filasX.length + 3, 1).setValue(
    'CÓMO USARLO\n' +
    '1. Pega la cadena en el buscador de LinkedIn (o en Sales Navigator, que da mejores filtros).\n' +
    '2. Filtra por Ubicación = Provincia de Málaga / Costa del Sol.\n' +
    '3. Guarda la búsqueda: Sales Navigator te avisa de los perfiles nuevos que entran. Eso convierte la captación en un flujo continuo.\n' +
    '4. Exporta o copia los resultados y entra por Menú → Captar candidatos → Importar CSV.\n\n' +
    'SIN LICENCIA DE SALES NAVIGATOR\n' +
    'Usa el bloque X-ray de arriba: se pega en Google, no en LinkedIn, y es gratis.\n' +
    'Si quieres que esas mismas consultas se ejecuten solas y vuelquen los perfiles\n' +
    'directamente en Rec_Candidatos, configura CSE_API_KEY y CSE_CX y usa la\n' +
    'opción 6 del menú (100 consultas gratis al día).\n\n' +
    'IMPORTANTE: no uses extensiones de scraping de LinkedIn (Phantombuster y similares).\n' +
    'Incumplen las condiciones de LinkedIn y la AEPD ya ha sancionado el uso de datos de\n' +
    'perfiles públicos para contacto no consentido. Exportación manual, X-ray o Sales Navigator.'
  ).setWrap(true).setFontWeight('bold');

  hoja.activate();
  SpreadsheetApp.getUi().alert('✅ Búsquedas generadas',
    'Hoja "Rec_Busquedas_LinkedIn" con 8 segmentos listos para pegar en LinkedIn.\n\n' +
    'Empieza por el segmento 2 (autónomos) y el 7 (cambio reciente de trabajo): son los de mayor tasa de respuesta.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

// ============================================================
//  11. SCORING — ¿a quién llamo primero?
// ============================================================

const REC_ZONAS_NUCLEO = ['marbella', 'nueva andalucía', 'nueva andalucia', 'puerto banús',
  'puerto banus', 'golden mile', 'milla de oro', 'san pedro', 'benahavís', 'benahavis',
  'estepona', 'guadalmina', 'elviria', 'la zagaleta'];

function recRecalcularScores(silencioso) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(REC.H_CANDIDATOS);
  if (!hoja || hoja.getLastRow() < 2) {
    if (!silencioso) SpreadsheetApp.getUi().alert('No hay candidatos todavía.');
    return;
  }

  const n = hoja.getLastRow() - 1;
  const datos = hoja.getRange(2, 1, n, REC_N_COLS).getValues();
  const salida = datos.map(f => {
    const s = recScoreCandidato_(f);
    return [s, recTemperatura_(s)];
  });

  hoja.getRange(2, REC_COL.SCORE, n, 2).setValues(salida);
  hoja.sort({ column: REC_COL.SCORE, ascending: false });

  if (!silencioso) {
    const porTemp = { A: 0, B: 0, C: 0, D: 0 };
    salida.forEach(s => { porTemp[s[1]] = (porTemp[s[1]] || 0) + 1; });
    SpreadsheetApp.getUi().alert('⭐ Puntuaciones recalculadas',
      'Candidatos: ' + n + '\n\n' +
      'A (llamar esta semana): ' + porTemp.A + '\n' +
      'B (llamar este mes): ' + porTemp.B + '\n' +
      'C (secuencia larga): ' + porTemp.C + '\n' +
      'D (solo nurture): ' + porTemp.D + '\n\n' +
      'La hoja queda ordenada de mayor a menor.',
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function recScoreCandidato_(f) {
  const P = REC.PESOS;
  let total = 0;

  // --- Producción: lo que más predice que merezca la pena la entrevista ---
  const inmuebles = parseFloat(f[REC_COL.INMUEBLES - 1]) || 0;
  const precioMedio = parseFloat(String(f[REC_COL.PRECIO_MEDIO - 1]).replace(/[^\d]/g, '')) || 0;
  let prod = 0;
  if (inmuebles >= 25) prod = 1.0;
  else if (inmuebles >= 15) prod = 0.85;
  else if (inmuebles >= 8) prod = 0.65;
  else if (inmuebles >= 4) prod = 0.45;
  else if (inmuebles >= 1) prod = 0.25;
  // El ticket alto multiplica: 3 villas de 4M valen más que 20 pisos de 200k
  if (precioMedio >= 2000000) prod = Math.min(1, prod + 0.30);
  else if (precioMedio >= 900000) prod = Math.min(1, prod + 0.20);
  else if (precioMedio >= 450000) prod = Math.min(1, prod + 0.10);
  total += prod * P.produccion;

  // --- Zona ---
  const zona = String(f[REC_COL.ZONA - 1]).toLowerCase();
  const enNucleo = REC_ZONAS_NUCLEO.some(z => zona.indexOf(z) !== -1);
  total += (enNucleo ? 1 : (zona ? 0.4 : 0.2)) * P.zona;

  // --- Perfil ---
  const perfil = String(f[REC_COL.PERFIL - 1]);
  const pesoPerfil = {
    'Agente autónomo / sin agencia': 1.0,
    'Agente en agencia independiente': 0.85,
    'Agente en gran franquicia': 0.70,
    'Cambio de sector (hostelería/lujo/banca)': 0.55,
    'Team Leader / Broker': 0.50,
    'Recién titulado / sin experiencia': 0.25,
    'Agente en otro MC KW': 0.05,   // no se recluta dentro de la propia red
    'Desconocido': 0.35
  }[perfil];
  total += (pesoPerfil === undefined ? 0.35 : pesoPerfil) * P.perfil;

  // --- Idiomas (en Marbella es dinero directo) ---
  const idiomas = String(f[REC_COL.IDIOMAS - 1]).toLowerCase();
  const nIdiomas = idiomas ? idiomas.split(/[,;\/]/).filter(x => x.trim()).length : 0;
  const idiomasValiosos = /sueco|swedish|noruego|norwegian|danés|danish|neerland|dutch|holand|alem|german|ruso|russian|árabe|arabic|polaco|polish|finlan|finnish|chino|chinese/.test(idiomas);
  let pIdiomas = 0.3;
  if (nIdiomas >= 3) pIdiomas = 0.9;
  else if (nIdiomas === 2) pIdiomas = 0.7;
  if (idiomasValiosos) pIdiomas = Math.min(1, pIdiomas + 0.25);
  total += pIdiomas * P.idiomas;

  // --- Experiencia: 2-8 años es el punto óptimo ---
  const exp = parseFloat(f[REC_COL.EXPERIENCIA - 1]) || 0;
  let pExp = 0.4;
  if (exp >= 2 && exp <= 8) pExp = 1.0;
  else if (exp > 8 && exp <= 15) pExp = 0.7;
  else if (exp > 15) pExp = 0.45;
  else if (exp >= 1) pExp = 0.6;
  total += pExp * P.experiencia;

  // --- Señales de dolor (de las notas y del cargo) ---
  const texto = (String(f[REC_COL.NOTAS - 1]) + ' ' + String(f[REC_COL.CARGO - 1])).toLowerCase();
  const señales = ['sin exclusiv', 'comisión baja', 'comision baja', 'split', 'buscando cambio',
    'open to work', 'abierto a oportunidades', 'sin formación', 'sin formacion',
    'solo', 'sin equipo', 'rotación', 'rotacion', 'descontent'];
  const nSeñales = señales.filter(s => texto.indexOf(s) !== -1).length;
  total += Math.min(1, nSeñales * 0.4) * P.dolor;

  // --- Accesibilidad ---
  total += (String(f[REC_COL.TELEFONO - 1]).trim() ? 1 : (String(f[REC_COL.EMAIL - 1]).trim() ? 0.5 : 0)) * P.accesibilidad;

  // --- Red / referido ---
  const fuente = String(f[REC_COL.FUENTE - 1]).toLowerCase();
  total += (/referido|referral|recomend/.test(fuente) ? 1 : 0.2) * P.red;

  return Math.round(total);
}

function recTemperatura_(score) {
  if (score >= 75) return 'A';
  if (score >= 55) return 'B';
  if (score >= 35) return 'C';
  return 'D';
}

// ============================================================
//  12. DEDUPLICACIÓN
// ============================================================

function recDeduplicar() {
  const ui = SpreadsheetApp.getUi();
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REC.H_CANDIDATOS);
  if (!hoja || hoja.getLastRow() < 3) { ui.alert('No hay suficientes filas.'); return; }

  const n = hoja.getLastRow() - 1;
  const datos = hoja.getRange(2, 1, n, REC_N_COLS).getValues();
  const vistos = {};
  const conservar = [];
  let eliminados = 0;

  // Mantener el registro más completo de cada duplicado
  const completitud = (f) => f.filter(c => String(c).trim() !== '').length;

  for (const f of datos) {
    const clave = recClaveDedupe_(
      f[REC_COL.NOMBRE - 1] + ' ' + f[REC_COL.APELLIDOS - 1],
      f[REC_COL.TELEFONO - 1], f[REC_COL.EMAIL - 1]);
    if (!clave) { conservar.push(f); continue; }

    if (vistos[clave] === undefined) {
      vistos[clave] = conservar.length;
      conservar.push(f);
    } else {
      eliminados++;
      const idx = vistos[clave];
      // Si el nuevo trae más información, fusionamos campo a campo
      if (completitud(f) > completitud(conservar[idx])) {
        for (let c = 0; c < REC_N_COLS; c++) {
          if (String(conservar[idx][c]).trim() === '' && String(f[c]).trim() !== '') {
            conservar[idx][c] = f[c];
          }
        }
        // Preservamos el estado más avanzado del pipeline
        const avance = REC_ESTADOS.indexOf(String(f[REC_COL.ESTADO - 1]));
        const avanceActual = REC_ESTADOS.indexOf(String(conservar[idx][REC_COL.ESTADO - 1]));
        if (avance > avanceActual) conservar[idx][REC_COL.ESTADO - 1] = f[REC_COL.ESTADO - 1];
      }
    }
  }

  if (eliminados === 0) { ui.alert('✅ Sin duplicados', 'La base ya está limpia.', ui.ButtonSet.OK); return; }

  hoja.getRange(2, 1, n, REC_N_COLS).clearContent();
  hoja.getRange(2, 1, conservar.length, REC_N_COLS).setValues(conservar);
  ui.alert('🧹 Base deduplicada',
    'Duplicados fusionados y eliminados: ' + eliminados + '\n' +
    'Candidatos únicos: ' + conservar.length, ui.ButtonSet.OK);
}

// ============================================================
//  13. SMART PLAN — la cadencia de seguimiento
// ============================================================

/**
 * SECUENCIA DE CANAL (esto no es estético, es legal y de conversión):
 *
 *  Toque 1  → LLAMADA al teléfono que el agente publica para recibir
 *             llamadas profesionales. Es el canal más defendible y el
 *             que más convierte.
 *  Toque 2  → WhatsApp SOLO si no contesta, de una línea, identificado
 *             y con salida ("dime baja y no te escribo más").
 *  Toque 3+ → WhatsApp libre una vez hay conversación.
 *  Email    → secundario, con la información del art. 14 RGPD.
 *
 *  Nunca al revés. Un primer contacto masivo por WhatsApp degrada la
 *  calificación del número, te lo bloquean, y encima es el canal con
 *  más riesgo legal.
 */
function recSembrarSmartPlan_(ss) {
  const hoja = ss.getSheetByName(REC.H_PLAN);
  if (hoja.getLastRow() > 1) return;  // no sobreescribir copys ya editados

  const filas = [];
  const add = (plan, paso, dia, canal, tipo, objetivo, asunto, es, en, consent) =>
    filas.push([plan, paso, dia, canal, tipo, objetivo, asunto || '', es, en, consent || 'NO', 'SI']);

  // ---------------------------------------------------------
  //  PLAN A — AGENTE_ACTIVO (el principal, 90 días, 15 toques)
  // ---------------------------------------------------------
  const A = 'AGENTE_ACTIVO';

  add(A, 1, -1, 'Llamada', 'Investigación',
    'Antes de llamar: mirar su cartera, LinkedIn e Instagram. Apuntar 1 dato concreto suyo.',
    '', 
    '[NO ES UN ENVÍO — ES PREPARACIÓN]\n\nAntes de marcar, ten delante:\n• Nº de inmuebles publicados y en qué zona\n• Rango de precio con el que trabaja\n• Cuánto lleva en su agencia actual\n• Idiomas\n• Un detalle personal de su perfil (deporte, origen, familia)\n\nSi no tienes un dato concreto suyo, NO llames todavía. La llamada genérica quema el contacto para siempre.',
    '[NOT A MESSAGE — PREPARATION]\n\nBefore dialling, have in front of you:\n• Number of active listings and area\n• Price range they work in\n• Time at their current agency\n• Languages\n• One personal detail from their profile\n\nIf you have no specific detail about them, do NOT call yet. A generic call burns the contact permanently.',
    'NO');

  add(A, 2, 0, 'Llamada', 'Apertura',
    'Conseguir permiso para una conversación. NO vender KW.',
    '',
    'Hola {{nombre}}, soy {{tl}} de {{mc}}.\n\nTe llamo por una razón concreta y en 30 segundos te dejo.\n\nHe estado viendo tu actividad en {{zona}} — {{inmuebles}} propiedades, y por el rango de precio se ve que trabajas producto bueno.\n\nNo te llamo para ficharte. Te llamo porque estamos cerrando el mapa de quién está produciendo de verdad en la zona y tú apareces ahí.\n\n¿Te puedo hacer dos preguntas de mercado o te pillo en mal momento?\n\n[SI DICE QUE SÍ]\n1. ¿Cuánto de tu negocio viene de cartera propia y cuánto te lo da la agencia?\n2. Si pudieras cambiar UNA cosa de cómo trabajas hoy, ¿cuál sería?\n\n[CIERRE]\nOye, me ha gustado hablar contigo. Te mando por WhatsApp el desglose de {{zona}} que te comentaba, sin compromiso. ¿Este número es el bueno?\n\n→ Ese "¿este número es el bueno?" ES el opt-in de WhatsApp. Márcalo en la ficha.',
    'Hi {{nombre}}, this is {{tl}} from {{mc}}.\n\nI am calling for one specific reason and I will be done in 30 seconds.\n\nI have been looking at your activity in {{zona}} — {{inmuebles}} listings, and the price range tells me you work good product.\n\nI am not calling to recruit you. I am calling because we are putting together the map of who is actually producing in this area, and you are on it.\n\nCan I ask you two market questions, or have I caught you at a bad time?\n\n[IF YES]\n1. How much of your business comes from your own database versus what the agency hands you?\n2. If you could change ONE thing about how you work today, what would it be?\n\n[CLOSE]\nLook, I enjoyed this. Let me WhatsApp you that {{zona}} breakdown I mentioned, no strings. Is this the right number?\n\n→ That "is this the right number?" IS your WhatsApp opt-in. Flag it on the record.',
    'NO');

  add(A, 3, 0, 'WhatsApp', 'Presentación',
    'Solo si NO ha contestado la llamada. Una línea, identificado, con salida.',
    '',
    'Hola {{nombre}}, soy {{tl}}, Team Leader de {{mc}} — te he llamado hace un rato, sin agobios.\n\nTe escribía por tu actividad en {{zona}}. No es una oferta de trabajo al uso: estamos cerrando un mapa de los agentes que más producen en la zona y me interesaba tu visión del mercado.\n\n¿Te va bien que te llame mañana sobre las 10:00, o prefieres por aquí?\n\nSi no te interesa, dime "baja" y no te vuelvo a escribir.',
    'Hi {{nombre}}, this is {{tl}}, Team Leader at {{mc}} — I rang you earlier, no pressure.\n\nI was getting in touch about your activity in {{zona}}. This is not a standard job offer: we are mapping the agents producing the most in the area and I wanted your read on the market.\n\nWould tomorrow around 10:00 work for a quick call, or would you rather do it here?\n\nIf it is not for you, just reply "stop" and I will not message again.',
    'NO');

  add(A, 4, 2, 'WhatsApp', 'Valor 1 — dato de mercado',
    'Dar algo útil sin pedir nada. Cero KW en el mensaje.',
    '',
    '{{nombre}}, te dejo el dato que te comentaba de {{zona}}:\n\n• Tiempo medio de venta: [X] días (hace un año eran [Y])\n• [Z]% de las operaciones se cierran por debajo del precio de salida\n• Lo que más se mueve: [tipología y rango]\n\nSi te interesa te paso el desglose por urbanización, lo tenemos por código postal.\n\n⚠️ RELLENA LOS CORCHETES CON DATOS REALES DE TU MC. Si no puedes respaldar un número, bórralo. Un dato inventado te cierra la puerta para siempre.',
    '{{nombre}}, here is the {{zona}} data I mentioned:\n\n• Average days on market: [X] (a year ago it was [Y])\n• [Z]% of deals close below asking\n• Moving fastest: [property type and range]\n\nIf it is useful I can send you the breakdown by urbanisation — we have it by postcode.\n\n⚠️ FILL THE BRACKETS WITH REAL MC DATA. If you cannot back a number up, delete it. One invented figure closes the door for good.',
    'NO');

  add(A, 5, 5, 'Llamada', 'Segundo intento',
    'Cambiar la franja horaria respecto al primer intento.',
    '',
    'Si la primera llamada fue por la mañana, esta por la tarde (17:00-19:00) y viceversa.\n\n"{{nombre}}, {{tl}} otra vez, de {{mc}}. Te mandé el dato de {{zona}} el otro día — ¿te sirvió de algo?"\n\n→ Pregunta abierta sobre algo que YA le diste. No sobre ti.',
    'If the first call was in the morning, make this one late afternoon (17:00-19:00) and vice versa.\n\n"{{nombre}}, {{tl}} again, from {{mc}}. I sent you the {{zona}} data the other day — was it any use?"\n\n→ Open question about something you ALREADY gave them. Not about you.',
    'NO');

  add(A, 6, 7, 'LinkedIn', 'Conexión',
    'Solicitud con nota. Te deja en su radar de forma pasiva.',
    '',
    'Hola {{nombre}}, {{tl}} de {{mc}}. Coincidimos en zona ({{zona}}) y me gusta cómo trabajas el producto. Te agrego para seguir lo que publicas.',
    'Hi {{nombre}}, {{tl}} from {{mc}}. We work the same patch ({{zona}}) and I like the product you handle. Connecting to follow what you post.',
    'NO');

  add(A, 7, 10, 'WhatsApp', 'Valor 2 — caso comparable',
    'Prueba social con alguien de su mismo perfil. Aquí ya se puede mencionar KW.',
    '',
    '{{nombre}}, una cosa que quizá te sirva.\n\n[Nombre del agente] venía de [tipo de agencia] con un volumen parecido al tuyo. El año pasado hizo [X]€ de honorarios con nosotros.\n\nLo que cambió no fue trabajar más horas. Fue quedarse con [%] en lugar de [%].\n\nSi te da curiosidad te paso su desglose — él me ha dado permiso para contarlo.\n\n⚠️ USA UN CASO REAL DE TU MC, CON PERMISO DE LA PERSONA. Si no tienes el caso todavía, salta este paso.',
    '{{nombre}}, something that might be useful.\n\n[Agent name] came from [type of agency] with a volume similar to yours. Last year they billed [X]€ in fees with us.\n\nWhat changed was not working more hours. It was keeping [%] instead of [%].\n\nIf you are curious I will send you their breakdown — they gave me permission to share it.\n\n⚠️ USE A REAL CASE FROM YOUR MC, WITH THAT PERSON PERMISSION. No case yet? Skip this step.',
    'NO');

  add(A, 8, 14, 'WhatsApp', 'Invitación a evento',
    'La vía con mayor conversión: que entre en la oficina sin compromiso.',
    '',
    '{{nombre}}, el [fecha] hacemos [formación/mesa redonda] en la oficina ({{direccion}}).\n\nEs abierto, vienen agentes de varias agencias y no hay pitch de nada. Se habla de [tema muy concreto: fiscalidad del no residente, cómo defender honorarios, captación en zona de obra nueva...].\n\n¿Te guardo sitio? Van a estar [2-3 nombres que le suenen].',
    '{{nombre}}, on [date] we are running [training/panel] at the office ({{direccion}}).\n\nIt is open, agents from several agencies come, and there is no pitch. The topic is [something very specific: non-resident tax, defending your fee, sourcing in new-build areas...].\n\nShall I save you a seat? [2-3 names they will recognise] will be there.',
    'NO');

  add(A, 9, 21, 'Llamada', 'Check-in',
    'Llamada corta apoyada en todo lo que ya le has dado.',
    '',
    '"{{nombre}}, {{tl}}. Te llamo 2 minutos. ¿Pudiste ver lo que te mandé?"\n\n[ESCUCHAR]\n\n"Te hago una pregunta directa y si me dices que no, lo dejo aquí: ¿has mirado alguna vez, con números delante, cuánto te quedaría a ti con otro modelo de reparto? No te pido que me des tus cifras. Te paso la hoja y la rellenas tú."',
    '"{{nombre}}, {{tl}}. Two minutes. Did you get a chance to look at what I sent?"\n\n[LISTEN]\n\n"Let me ask you one direct question, and if you say no I will leave it here: have you ever sat down with actual numbers and worked out what you would keep under a different split? I am not asking you for your figures. I will send you the sheet and you fill it in yourself."',
    'NO');

  add(A, 10, 28, 'WhatsApp', 'Valor 3 — calculadora',
    'Le damos el control del cálculo. No le pedimos datos.',
    '',
    '{{nombre}}, te paso la hoja de la que te hablaba.\n\nMetes tus operaciones del año pasado y tu comisión media, y te dice lo que te habrías quedado con cada modelo. La rellenas tú: no sale de tu ordenador y no me mandas ningún número.\n\nLa mayoría se sorprende con la línea del tope de aportación: a partir de ahí dejas de aportar al Market Center.\n\n[adjuntar calculadora_ingresos_agente.xlsx — rellena antes la pestaña Parametros_MC con los datos reales del MC]',
    '{{nombre}}, here is the sheet I mentioned.\n\nYou enter last year deals and your average fee, and it tells you what you would have kept under each model. You fill it in — it never leaves your computer and you send me no numbers.\n\nMost people are surprised by the cap line: past that point you stop contributing to the Market Center.\n\n[attach calculadora_ingresos_agente.xlsx — fill the Parametros_MC tab with your real MC figures first]',
    'NO');

  add(A, 11, 35, 'Email', 'Valor 4 — modelo económico',
    'Documento largo. El email aguanta lo que WhatsApp no.',
    'El modelo económico, en 2 páginas ({{nombre}})',
    'Hola {{nombre}},\n\nTe mando en PDF lo que no cabe en un WhatsApp: cómo funciona el reparto, el tope de aportación, el beneficio compartido y a quién pertenece la cartera cuando te vas.\n\nEsa última parte es la que nadie cuenta y es la importante: aquí tus clientes son tuyos.\n\nSi quieres lo vemos en 30 minutos con tus números delante, sin compromiso. Y si no es el momento, me lo dices y te dejo tranquilo.\n\n{{tl}}\n{{mc}} · {{telefono_tl}}\n\n---\nTe escribo a tu dirección profesional publicada en {{url_fuente}}, por interés legítimo en contactar con profesionales del sector (art. 6.1.f RGPD). Puedes ejercer tus derechos de acceso, rectificación, supresión y oposición escribiendo a {{email_mc}}. Si no quieres recibir más comunicaciones, responde "BAJA" y te eliminamos de inmediato.',
    'Hi {{nombre}},\n\nSending you the PDF with what does not fit in a WhatsApp: how the split works, the cap, profit share, and who owns the database when you leave.\n\nThat last part is the one nobody talks about, and it is the one that matters: here your clients are yours.\n\nIf you want we can go through it in 30 minutes with your own numbers, no commitment. And if it is not the right time, tell me and I will leave you alone.\n\n{{tl}}\n{{mc}} · {{telefono_tl}}\n\n---\nI am writing to your professional address published at {{url_fuente}}, on the basis of legitimate interest in contacting professionals in the sector (art. 6.1.f GDPR). You can exercise your rights of access, rectification, erasure and objection by writing to {{email_mc}}. To stop receiving messages, reply "STOP" and we will remove you immediately.',
    'NO');

  add(A, 12, 45, 'WhatsApp', 'Valor 5 — prueba social',
    'Alguien que acaba de entrar. Movimiento real, no promesas.',
    '',
    '{{nombre}}, [nombre] se ha incorporado este mes — venía de [agencia]. Lo cuento porque trabajaba tu misma zona y quizá os conocéis.\n\nSi quieres hablar con él directamente y que te cuente sin filtro cómo ha sido el cambio, te paso su número. A mí no me vas a creer, a él sí.',
    '{{nombre}}, [name] joined us this month — came over from [agency]. I mention it because they worked your patch and you may know each other.\n\nIf you want to speak to them directly and get the unfiltered version of how the move went, I will pass on their number. You will not believe me, but you will believe them.',
    'NO');

  add(A, 13, 60, 'Llamada', 'Reapertura',
    'Nueva razón para llamar. Nunca "¿lo has pensado?".',
    '',
    '"{{nombre}}, {{tl}}. No te llamo para insistir. Te llamo porque [novedad real: hemos abierto X, ha entrado Y, lanzamos formación en Z] y me acordé de lo que me dijiste sobre [su dolor concreto].\n\n¿Sigue siendo igual o ha cambiado algo?"\n\n→ Si sigue siendo no: "Perfecto. Te mando el informe de mercado cada trimestre y ya está, sin más. Si algún día cambia la cosa, ya sabes dónde estoy."',
    '"{{nombre}}, {{tl}}. I am not calling to push. I am calling because [real news: we opened X, Y joined, we are launching training on Z] and it reminded me of what you said about [their specific pain].\n\nIs that still the case or has something changed?"\n\n→ If it is still no: "Perfect. I will send you the quarterly market report and leave it at that. If things change, you know where I am."',
    'NO');

  add(A, 14, 75, 'WhatsApp', 'Informe trimestral',
    'Mantener presencia con valor, cero presión.',
    '',
    '{{nombre}}, informe del trimestre de {{zona}}. Te lo mando porque dijiste que te servía, no para venderte nada.\n\n[adjuntar informe]',
    '{{nombre}}, this quarter report for {{zona}}. Sending it because you said it was useful, not to sell you anything.\n\n[attach report]',
    'NO');

  add(A, 15, 90, 'Llamada', 'Cierre de ciclo',
    'Decisión: pasa a nurture mensual o se archiva. No se queda en el limbo.',
    '',
    '"{{nombre}}, llevamos tres meses hablando y quiero ser claro para no hacerte perder tiempo.\n\n¿Lo dejamos en que te mando el informe trimestral y ya, o hay algo que te gustaría ver de cerca antes de descartarlo del todo?"\n\n→ Pase lo que pase, ACTUALIZA EL ESTADO en la ficha:\n  • "No ahora (nurture)" → entra en plan NURTURE (1 toque/mes)\n  • "Descartado" → fuera de secuencia\n  • "Entrevista agendada" → plan POST_ENTREVISTA',
    '"{{nombre}}, we have been talking for three months and I want to be straight with you so neither of us wastes time.\n\nShall we leave it at me sending the quarterly report, or is there something you would want to see up close before ruling it out completely?"\n\n→ Whatever happens, UPDATE THE STATUS on the record:\n  • "No ahora (nurture)" → goes into NURTURE plan (1 touch/month)\n  • "Descartado" → out of sequence\n  • "Entrevista agendada" → POST_ENTREVISTA plan',
    'NO');

  // ---------------------------------------------------------
  //  PLAN B — AUTONOMO (el de mayor conversión)
  // ---------------------------------------------------------
  const B = 'AUTONOMO';

  add(B, 1, 0, 'Llamada', 'Apertura autónomo',
    'Su dolor no es el split: es la soledad y la falta de estructura.',
    '',
    'Hola {{nombre}}, soy {{tl}} de {{mc}}.\n\nTe llamo porque vi que trabajas por tu cuenta en {{zona}}, y de hecho eso es lo que me hizo llamarte: la gente que vende sola es la que más sabe vender, porque no tiene a nadie detrás.\n\nUna pregunta y te dejo: ¿qué es lo que más te pesa de ir solo — el marketing, los portales, la parte legal, o simplemente no tener con quién consultar una operación?\n\n[ESCUCHAR. ESE ES SU DOLOR Y ES TU ENTRADA]\n\n"Te lo pregunto porque aquí la gente llega por eso, no por la comisión. La comisión es que te quedas más. Pero lo que les cambia el día a día es tener estructura detrás y seguir siendo dueños de su cartera."',
    'Hi {{nombre}}, this is {{tl}} from {{mc}}.\n\nI am calling because I saw you work independently in {{zona}}, and honestly that is why I called: people who sell on their own are the ones who really know how to sell, because they have nobody behind them.\n\nOne question and I will let you go: what weighs on you most about working alone — the marketing, the portals, the legal side, or simply having nobody to sanity-check a deal with?\n\n[LISTEN. THAT IS THEIR PAIN AND THAT IS YOUR WAY IN]\n\n"I ask because that is why people come here, not for the commission. The commission means you keep more. But what changes their day to day is having structure behind them while still owning their database."',
    'NO');

  add(B, 2, 0, 'WhatsApp', 'Presentación autónomo',
    'Si no contesta.',
    '',
    'Hola {{nombre}}, soy {{tl}} de {{mc}} (te he llamado hace un rato).\n\nTe escribo porque trabajas por tu cuenta en {{zona}} y tenemos bastantes agentes que venían de ahí. No te quiero vender nada por WhatsApp: solo saber si te interesa que hablemos 15 minutos de cómo lo tienen montado.\n\nSi no, dime "baja" y listo.',
    'Hi {{nombre}}, this is {{tl}} from {{mc}} (I called you earlier).\n\nI am getting in touch because you work independently in {{zona}} and quite a few of our agents came from exactly that. I am not going to sell you anything over WhatsApp: I just want to know if you would spend 15 minutes hearing how they have it set up.\n\nIf not, reply "stop" and that is that.',
    'NO');

  add(B, 3, 3, 'WhatsApp', 'Valor — lo que no tiene',
    'Enseñar la estructura, no la comisión.',
    '',
    '{{nombre}}, lo concreto que tendrías y hoy no tienes:\n\n• Portales y fotografía/vídeo pagados por el MC\n• Un abogado y un fiscalista a los que preguntar sin facturar consulta\n• Formación semanal en la oficina\n• Un CRM con los seguimientos automatizados\n• Alguien con quien repartirte una operación grande\n\nY la cartera sigue siendo tuya. Eso es lo que no te va a ofrecer una agencia al uso.\n\n¿Te interesa verlo de cerca?',
    '{{nombre}}, concretely, what you would have and do not have today:\n\n• Portals and photo/video paid by the MC\n• A lawyer and a tax adviser you can ask without being billed for the consultation\n• Weekly training at the office\n• A CRM with follow-ups automated\n• Someone to split a big deal with\n\nAnd your database stays yours. That is what a standard agency will not offer you.\n\nWant to see it up close?',
    'NO');

  // ---------------------------------------------------------
  //  PLAN C — CAMBIO_SECTOR (hostelería, lujo, banca privada)
  // ---------------------------------------------------------
  const C = 'CAMBIO_SECTOR';

  add(C, 1, 0, 'LinkedIn', 'Apertura cambio de sector',
    'No saben que son candidatos. Hay que explicárselo.',
    '',
    'Hola {{nombre}}, soy {{tl}}, Team Leader de {{mc}}.\n\nTe escribo por algo que te va a sonar raro: no busco gente del sector inmobiliario.\n\nLos mejores agentes que tenemos vienen de hostelería de lujo, banca privada y retail premium. Y la razón es simple: ya saben tratar con el cliente que compra una casa de 2 millones en Marbella. Eso no se enseña. Lo técnico sí.\n\nTú llevas [X años] en [sector] aquí en Marbella. ¿Te has planteado alguna vez el salto?\n\nSi es un no rotundo, dímelo y no insisto más.',
    'Hi {{nombre}}, this is {{tl}}, Team Leader at {{mc}}.\n\nI am writing about something that will sound odd: I am not looking for people from real estate.\n\nOur best agents come from luxury hospitality, private banking and premium retail. The reason is simple: they already know how to handle the client who buys a 2 million euro house in Marbella. That cannot be taught. The technical side can.\n\nYou have spent [X years] in [sector] here in Marbella. Have you ever considered the move?\n\nIf it is a flat no, tell me and I will not push.',
    'NO');

  add(C, 2, 4, 'WhatsApp', 'Resolver el miedo real',
    'Su objeción no es el interés: es el riesgo económico.',
    '',
    '{{nombre}}, lo que frena a todo el mundo en tu situación es lo mismo: "y mientras aprendo, ¿de qué vivo?".\n\nTe cuento cómo lo hacemos, sin adornos:\n• [X] semanas de formación antes de salir a la calle\n• Acompañamiento en tus primeras visitas y negociaciones\n• [Modelo de ingresos del primer año — SÉ HONESTO AQUÍ]\n\nNo te voy a decir que es fácil. Los primeros [X] meses son duros. Pero el que aguanta, en este mercado y con tu agenda, no vuelve.\n\n¿Un café sin compromiso?',
    '{{nombre}}, everyone in your position is held back by the same thing: "and while I learn, what do I live on?".\n\nHere is how we handle it, straight:\n• [X] weeks of training before you go out\n• Someone with you on your first viewings and negotiations\n• [First-year income model — BE HONEST HERE]\n\nI am not going to tell you it is easy. The first [X] months are hard. But those who stick it out, in this market and with your contacts, do not go back.\n\nCoffee, no commitment?',
    'NO');

  // ---------------------------------------------------------
  //  PLAN D — NURTURE (el "no ahora", 1 toque/mes, 12 meses)
  // ---------------------------------------------------------
  const D = 'NURTURE';
  add(D, 1, 30, 'WhatsApp', 'Nurture mensual',
    'Un dato útil al mes. Cero presión. El 40% de las incorporaciones salen de aquí.',
    '',
    '{{nombre}}, el dato del mes de {{zona}}: [dato concreto].\n\nSin más. Que vaya bien el mes.',
    '{{nombre}}, this month data point for {{zona}}: [specific figure].\n\nThat is all. Have a good month.',
    'NO');
  add(D, 2, 90, 'Llamada', 'Nurture trimestral',
    'Llamada de temperatura cada trimestre.',
    '',
    '"{{nombre}}, {{tl}}. Llamada de las de siempre: ¿sigue todo bien por ahí o ha cambiado algo?"',
    '"{{nombre}}, {{tl}}. Usual check-in: all still good your end, or has anything changed?"',
    'NO');

  // ---------------------------------------------------------
  //  PLAN E — POST_ENTREVISTA
  // ---------------------------------------------------------
  const E = 'POST_ENTREVISTA';
  add(E, 1, 0, 'WhatsApp', 'Mismo día de la entrevista',
    'Cerrar el siguiente paso en caliente, con fecha.',
    '',
    '{{nombre}}, gracias por el rato de hoy. Me quedo con dos cosas que me dijiste: [lo que quiere conseguir] y [lo que le frena].\n\nTe mando [lo prometido] antes del viernes.\n\nY lo siguiente sería [paso concreto] el [fecha]. ¿Te cuadra?',
    '{{nombre}}, thanks for your time today. Two things stayed with me: [what they want] and [what is holding them back].\n\nI will send you [what you promised] before Friday.\n\nNext step would be [specific step] on [date]. Does that work?',
    'NO');
  add(E, 2, 2, 'Email', 'Lo prometido',
    'Cumplir en plazo. Es la primera prueba de cómo trabajas.',
    'Lo que te prometí, {{nombre}}',
    'Hola {{nombre}},\n\nTe dejo lo que quedamos:\n• [documento 1]\n• [documento 2]\n\nY la respuesta a lo que me preguntaste sobre [tema]: [respuesta clara].\n\nNos vemos el [fecha] para [paso]. Si te surge cualquier cosa antes, me llamas.\n\n{{tl}} · {{telefono_tl}}',
    'Hi {{nombre}},\n\nHere is what we agreed:\n• [document 1]\n• [document 2]\n\nAnd the answer to your question about [topic]: [clear answer].\n\nSee you on [date] for [step]. Anything comes up before then, call me.\n\n{{tl}} · {{telefono_tl}}',
    'NO');
  add(E, 3, 7, 'Llamada', 'Career Visioning',
    'La entrevista de verdad: sus objetivos de vida, no el puesto.',
    '',
    'ESTRUCTURA (no es una entrevista de trabajo, es una conversación sobre su vida):\n\n1. ¿Dónde quieres estar dentro de 5 años? (personal, no profesional)\n2. ¿Cuánto necesitas ganar para eso? → cifra concreta\n3. ¿Cuántas operaciones son, con tu comisión media?\n4. ¿Cuántas hiciste el año pasado? → el hueco entre 3 y 4 es la conversación\n5. ¿Qué te falta para cerrar ese hueco?\n6. ¿Qué pasa si dentro de 5 años sigues exactamente igual que hoy?\n\n→ Tú no vendes. Él se vende a sí mismo el cambio. Tu trabajo es hacer las preguntas y callarte.',
    'STRUCTURE (this is not a job interview, it is a conversation about their life):\n\n1. Where do you want to be in 5 years? (personal, not professional)\n2. What do you need to earn for that? → a specific number\n3. How many deals is that, at your average fee?\n4. How many did you do last year? → the gap between 3 and 4 is the conversation\n5. What is missing to close that gap?\n6. What happens if in 5 years you are exactly where you are today?\n\n→ You do not sell. They sell the change to themselves. Your job is to ask and then be quiet.',
    'NO');

  hoja.getRange(2, 1, filas.length, 11).setValues(filas);
  hoja.getRange(2, 8, filas.length, 2).setWrap(true).setVerticalAlignment('top');
  hoja.setColumnWidth(6, 260).setColumnWidth(8, 520).setColumnWidth(9, 520);
  hoja.getRange(2, 6, filas.length, 1).setWrap(true);
}

// ============================================================
//  14. MOTOR DEL SMART PLAN
// ============================================================

function recCargarPlan_(plan) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REC.H_PLAN);
  if (!hoja || hoja.getLastRow() < 2) return [];
  return hoja.getRange(2, 1, hoja.getLastRow() - 1, 11).getValues()
    .filter(f => String(f[0]).trim() === plan && String(f[10]).toUpperCase() === 'SI')
    .map(f => ({
      plan: f[0], paso: Number(f[1]), dia: Number(f[2]), canal: String(f[3]),
      tipo: String(f[4]), objetivo: String(f[5]), asunto: String(f[6]),
      es: String(f[7]), en: String(f[8]),
      requiereConsent: String(f[9]).toUpperCase() === 'SI'
    }))
    .sort((a, b) => a.paso - b.paso);
}

/** Asigna automáticamente el plan que corresponde por perfil. */
function recAutoAsignarPlanes() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REC.H_CANDIDATOS);
  if (!hoja || hoja.getLastRow() < 2) return 0;

  const n = hoja.getLastRow() - 1;
  const datos = hoja.getRange(2, 1, n, REC_N_COLS).getValues();
  let asignados = 0;

  for (let i = 0; i < datos.length; i++) {
    const f = datos[i];
    if (String(f[REC_COL.PLAN - 1]).trim()) continue;             // ya tiene plan
    const estado = String(f[REC_COL.ESTADO - 1]).trim();
    if (estado === 'Opt-out' || estado === 'Descartado' || estado === 'Incorporado') continue;
    if (String(f[REC_COL.PERFIL - 1]) === 'Agente en otro MC KW') continue;  // no se recluta en la propia red
    if (recTemperatura_(Number(f[REC_COL.SCORE - 1]) || 0) === 'D') continue; // los D solo nurture manual

    const perfil = String(f[REC_COL.PERFIL - 1]);
    let plan = 'AGENTE_ACTIVO';
    if (perfil === 'Agente autónomo / sin agencia') plan = 'AUTONOMO';
    else if (perfil === 'Cambio de sector (hostelería/lujo/banca)') plan = 'CAMBIO_SECTOR';
    else if (perfil === 'Recién titulado / sin experiencia') plan = 'CAMBIO_SECTOR';

    const pasos = recCargarPlan_(plan);
    if (!pasos.length) continue;

    hoja.getRange(i + 2, REC_COL.PLAN).setValue(plan);
    hoja.getRange(i + 2, REC_COL.PASO).setValue(0);
    const primero = pasos[0];
    const fecha = new Date();
    fecha.setDate(fecha.getDate() + Math.max(0, primero.dia));
    hoja.getRange(i + 2, REC_COL.PROXIMO_TOQUE).setValue(fecha);
    hoja.getRange(i + 2, REC_COL.ESTADO).setValue('En secuencia');
    asignados++;
  }
  return asignados;
}

/**
 * NÚCLEO DEL SISTEMA.
 * Recorre los candidatos con toque vencido, resuelve el paso que toca,
 * comprueba supresión y consentimiento, y escribe la cola del día en
 * Rec_Toques con el mensaje ya personalizado y el enlace de WhatsApp.
 */
function recGenerarToquesDelDia() {
  return recGenerarToques_(false);
}

const REC_UI_MUDA = {
  alert: function () {},
  ButtonSet: { OK: null, YES_NO: null },
  Button: { YES: null, OK: null }
};

function recGenerarToques_(silencioso) {
  const ui = silencioso ? REC_UI_MUDA : SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hCa = ss.getSheetByName(REC.H_CANDIDATOS);
  const hTo = ss.getSheetByName(REC.H_TOQUES);
  if (!hCa || !hTo) { ui.alert('Ejecuta primero recInicializarTodo()'); return; }

  recRecalcularScores(true);
  const asignados = recAutoAsignarPlanes();

  const hoy = new Date(); hoy.setHours(0, 0, 0, 0);
  const maxDia = parseInt(recLeerConfig_('MAX_TOQUES_DIA', String(REC.MAX_TOQUES_DIA)), 10);
  const supresion = recCargarSupresion_();
  const planesCache = {};

  const n = hCa.getLastRow() - 1;
  if (n < 1) { ui.alert('No hay candidatos.'); return; }
  const datos = hCa.getRange(2, 1, n, REC_N_COLS).getValues();

  const modoWA = recLeerConfig_('MODO_WHATSAPP', REC.MODO_WHATSAPP).toUpperCase();
  const waPrimerToque = recLeerConfig_('WA_EN_PRIMER_TOQUE',
    REC.WA_EN_PRIMER_TOQUE ? 'SI' : 'NO').toUpperCase() === 'SI';

  const toques = [];
  const actualizaciones = [];   // [filaHoja, paso, proximoToque, nToques, ultimoToque, estado]
  let saltadosSupresion = 0, saltadosConsent = 0, finalizados = 0;

  for (let i = 0; i < datos.length; i++) {
    if (toques.length >= maxDia) break;
    const f = datos[i];
    const plan = String(f[REC_COL.PLAN - 1]).trim();
    if (!plan) continue;

    const estado = String(f[REC_COL.ESTADO - 1]).trim();
    if (['Opt-out', 'Descartado', 'Incorporado', 'Entrevista agendada'].indexOf(estado) !== -1) continue;

    const proximo = f[REC_COL.PROXIMO_TOQUE - 1];
    if (!(proximo instanceof Date)) continue;
    const fProximo = new Date(proximo); fProximo.setHours(0, 0, 0, 0);
    if (fProximo > hoy) continue;    // aún no toca

    // --- Supresión: esto bloquea TODO, es la primera comprobación ---
    if (recEstaSuprimido_(supresion, f[REC_COL.TELEFONO - 1], f[REC_COL.EMAIL - 1], f[REC_COL.LINKEDIN - 1])) {
      actualizaciones.push([i + 2, f[REC_COL.PASO - 1], '', f[REC_COL.N_TOQUES - 1], f[REC_COL.ULTIMO_TOQUE - 1], 'Opt-out']);
      saltadosSupresion++;
      continue;
    }

    if (!planesCache[plan]) planesCache[plan] = recCargarPlan_(plan);
    const pasos = planesCache[plan];
    if (!pasos.length) continue;

    const pasoActual = Number(f[REC_COL.PASO - 1]) || 0;
    const pasoPlan = pasos.find(p => p.paso > pasoActual);
    // Copia por candidato: si no, canalEfectivo se quedaría pegado en la caché
    // del plan y contaminaría a todos los candidatos posteriores del mismo lote.
    const siguiente = pasoPlan ? Object.assign({}, pasoPlan) : null;

    if (!siguiente) {
      // Plan terminado → a nurture si no se ha descartado
      actualizaciones.push([i + 2, pasoActual, '', f[REC_COL.N_TOQUES - 1], f[REC_COL.ULTIMO_TOQUE - 1],
                            estado === 'En secuencia' ? 'No ahora (nurture)' : estado]);
      finalizados++;
      continue;
    }

    // --- WhatsApp: ¿hace falta consentimiento para este paso? ---
    const consent = String(f[REC_COL.CONSENT_WA - 1]).toUpperCase();

    if (siguiente.canal === 'WhatsApp') {
      const esPrimerContactoEscrito = (Number(f[REC_COL.N_TOQUES - 1]) || 0) <= 1;
      const bloqueaApi = (modoWA === 'API' && consent !== 'SI');
      const bloqueaPrimero = (esPrimerContactoEscrito && !waPrimerToque && consent !== 'SI');
      if ((siguiente.requiereConsent && consent !== 'SI') || bloqueaApi || bloqueaPrimero) {
        // Reencaminamos a LinkedIn si lo tenemos, si no a llamada
        siguiente.canalEfectivo = String(f[REC_COL.LINKEDIN - 1]).trim() ? 'LinkedIn' : 'Llamada';
        saltadosConsent++;
      }
    }

    const idioma = String(f[REC_COL.IDIOMA_PREF - 1]).toUpperCase() === 'EN' ? 'en' : 'es';
    const plantilla = siguiente[idioma] || siguiente.es;
    const mensaje = recRenderPlantilla_(plantilla, f);
    const canal = siguiente.canalEfectivo || siguiente.canal;
    const tel = String(f[REC_COL.TELEFONO - 1]).trim();

    let enlace = '';
    if (canal === 'WhatsApp' && tel) enlace = recEnlaceWhatsApp_(tel, mensaje);
    else if (canal === 'Llamada' && tel) enlace = 'tel:' + tel;
    else if (canal === 'LinkedIn') enlace = String(f[REC_COL.LINKEDIN - 1]).trim();
    else if (canal === 'Email') enlace = 'mailto:' + String(f[REC_COL.EMAIL - 1]).trim();

    const ahora = new Date();
    toques.push([
      recNuevoId_('T'),
      f[REC_COL.ID - 1],
      (f[REC_COL.NOMBRE - 1] + ' ' + f[REC_COL.APELLIDOS - 1]).trim(),
      ahora,
      Utilities.formatDate(ahora, Session.getScriptTimeZone(), 'HH:mm'),
      canal,
      plan + ' · paso ' + siguiente.paso,
      siguiente.tipo,
      'PENDIENTE',
      '',
      mensaje + (enlace ? '\n\n▶ ' + enlace : ''),
      Session.getActiveUser().getEmail() || '',
      siguiente.objetivo
    ]);

    // Calcular la fecha del paso siguiente (relativa, se autocorrige si vamos con retraso)
    const posterior = pasos.find(p => p.paso > siguiente.paso);
    let fechaSiguiente = '';
    if (posterior) {
      const delta = Math.max(1, posterior.dia - siguiente.dia);
      const d = new Date(hoy); d.setDate(d.getDate() + delta);
      fechaSiguiente = d;
    }
    actualizaciones.push([
      i + 2, siguiente.paso, fechaSiguiente,
      (Number(f[REC_COL.N_TOQUES - 1]) || 0) + 1,
      new Date(),
      estado === 'Nuevo' || estado === 'Investigado' ? 'En secuencia' : estado
    ]);
  }

  // Escribir toques
  if (toques.length) {
    hTo.getRange(hTo.getLastRow() + 1, 1, toques.length, 13).setValues(toques);
    hTo.getRange(2, 11, hTo.getLastRow() - 1, 1).setWrap(true).setVerticalAlignment('top');
    hTo.setColumnWidth(11, 560);
  }

  // Aplicar actualizaciones a los candidatos
  for (const u of actualizaciones) {
    const [fila, paso, proximo, nToques, ultimo, estado] = u;
    hCa.getRange(fila, REC_COL.PASO).setValue(paso);
    hCa.getRange(fila, REC_COL.PROXIMO_TOQUE).setValue(proximo);
    hCa.getRange(fila, REC_COL.N_TOQUES).setValue(nToques);
    hCa.getRange(fila, REC_COL.ULTIMO_TOQUE).setValue(ultimo);
    hCa.getRange(fila, REC_COL.ESTADO).setValue(estado);
  }

  const stats = {
    toques: toques.length, asignados: asignados,
    saltadosSupresion: saltadosSupresion, saltadosConsent: saltadosConsent,
    finalizados: finalizados
  };
  ui.alert('📨 Cola del día generada',
    'Toques pendientes: ' + toques.length + '\n' +
    (asignados ? 'Candidatos inscritos en plan: ' + asignados + '\n' : '') +
    (saltadosSupresion ? 'Bloqueados por opt-out: ' + saltadosSupresion + '\n' : '') +
    (saltadosConsent ? 'WhatsApp reencaminado por falta de consentimiento: ' + saltadosConsent + '\n' : '') +
    (finalizados ? 'Planes finalizados (pasan a nurture): ' + finalizados + '\n' : '') +
    '\nAbre el panel: menú 🎯 Reclutamiento → Panel diario.',
    ui.ButtonSet.OK);
  return stats;
}

/** Sustituye los {{marcadores}} por los datos reales del candidato. */
function recRenderPlantilla_(texto, f) {
  const nombre = String(f[REC_COL.NOMBRE - 1]).trim();
  const inm = f[REC_COL.INMUEBLES - 1];
  const vals = {
    nombre: nombre || 'hola',
    apellidos: String(f[REC_COL.APELLIDOS - 1]).trim(),
    agencia: String(f[REC_COL.AGENCIA - 1]).trim() || 'tu agencia',
    zona: String(f[REC_COL.ZONA - 1]).trim() || 'la zona',
    inmuebles: inm ? String(inm) : 'varias',
    cargo: String(f[REC_COL.CARGO - 1]).trim(),
    idiomas: String(f[REC_COL.IDIOMAS - 1]).trim(),
    url_fuente: String(f[REC_COL.URL_FUENTE - 1]).trim(),
    tl: recLeerConfig_('TL_NOMBRE', REC.TL_NOMBRE),
    telefono_tl: recLeerConfig_('TL_TELEFONO', REC.TL_TELEFONO),
    mc: REC.MC_NOMBRE,
    email_mc: recLeerConfig_('MC_EMAIL', REC.MC_EMAIL),
    web_mc: recLeerConfig_('MC_WEB', REC.MC_WEB),
    direccion: recLeerConfig_('MC_DIRECCION', REC.MC_DIRECCION)
  };
  return String(texto).replace(/\{\{(\w+)\}\}/g, (m, k) =>
    vals[k] !== undefined ? vals[k] : m);
}

function recEnlaceWhatsApp_(tel, mensaje) {
  const num = String(tel).replace(/[^\d]/g, '');
  // Quitamos del enlace las notas internas en corchetes y las advertencias
  const limpio = String(mensaje)
    .replace(/^\[NO ES UN ENVÍO[\s\S]*$/m, '')
    .replace(/⚠️[\s\S]*$/m, '')
    .replace(/→[^\n]*/g, '')
    .trim();
  return 'https://wa.me/' + num + '?text=' + encodeURIComponent(limpio);
}

/** Envío por Cloud API. Solo con consentimiento y plantilla aprobada. */
function recEnviarWhatsAppAPI_(tel, nombrePlantilla, parametros, idioma) {
  const token = recClave_('WA_TOKEN');
  const phoneId = recClave_('WA_PHONE_ID');
  const num = String(tel).replace(/[^\d]/g, '');

  const payload = {
    messaging_product: 'whatsapp',
    to: num,
    type: 'template',
    template: {
      name: nombrePlantilla,
      language: { code: idioma === 'en' ? 'en' : 'es' },
      components: [{
        type: 'body',
        parameters: (parametros || []).map(p => ({ type: 'text', text: String(p) }))
      }]
    }
  };

  const res = UrlFetchApp.fetch('https://graph.facebook.com/v21.0/' + phoneId + '/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + token },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  const ok = res.getResponseCode() === 200;
  return { ok: ok, respuesta: res.getContentText().substring(0, 500) };
}

// ============================================================
//  15. RGPD / LSSI — supresión, base legal, retención
// ============================================================

function recCargarSupresion_() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REC.H_SUPRESION);
  const set = {};
  if (!hoja || hoja.getLastRow() < 2) return set;
  hoja.getRange(2, 1, hoja.getLastRow() - 1, 1).getValues().forEach(f => {
    const v = String(f[0]).trim().toLowerCase();
    if (!v) return;
    set[v] = true;
    const tel = recNormalizarTelefono_(v);
    if (tel) set[tel.toLowerCase()] = true;
  });
  return set;
}

function recEstaSuprimido_(set, tel, email, linkedin) {
  const t = recNormalizarTelefono_(tel).toLowerCase();
  const e = String(email || '').trim().toLowerCase();
  const l = String(linkedin || '').trim().toLowerCase();
  return !!(set[t] || set[e] || set[l]);
}

function recAbrirSupresion() {
  const html = HtmlService.createHtmlOutput(recHtmlSupresion_()).setWidth(560).setHeight(460);
  SpreadsheetApp.getUi().showModalDialog(html, '🔐 Opt-out / derecho de supresión');
}

/**
 * Registra una baja y la aplica a TODOS los canales de golpe.
 * Esto es lo que hay que ejecutar cuando alguien responde "BAJA"/"STOP"
 * o ejerce su derecho de oposición o supresión.
 */
function recRegistrarOptOut(identificador, motivo, nombre) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hSu = ss.getSheetByName(REC.H_SUPRESION);
  const hCa = ss.getSheetByName(REC.H_CANDIDATOS);

  const id = String(identificador).trim();
  if (!id) throw new Error('Indica un teléfono, email o URL de LinkedIn.');

  const tipo = id.indexOf('@') !== -1 ? 'Email'
             : /linkedin/i.test(id) ? 'LinkedIn'
             : 'Teléfono';
  const normalizado = tipo === 'Teléfono' ? recNormalizarTelefono_(id) : id.toLowerCase();

  hSu.appendRow([normalizado, tipo, nombre || '', new Date(), motivo || 'Solicitud del interesado',
                 'Manual', Session.getActiveUser().getEmail() || '']);

  // Marcar y detener la secuencia de todas las fichas que coincidan
  let afectados = 0;
  if (hCa.getLastRow() > 1) {
    const n = hCa.getLastRow() - 1;
    const datos = hCa.getRange(2, 1, n, REC_N_COLS).getValues();
    for (let i = 0; i < datos.length; i++) {
      const f = datos[i];
      const coincide =
        (tipo === 'Teléfono' && recNormalizarTelefono_(f[REC_COL.TELEFONO - 1]) === normalizado) ||
        (tipo === 'Email' && String(f[REC_COL.EMAIL - 1]).trim().toLowerCase() === normalizado) ||
        (tipo === 'LinkedIn' && String(f[REC_COL.LINKEDIN - 1]).trim().toLowerCase() === normalizado);
      if (!coincide) continue;
      hCa.getRange(i + 2, REC_COL.ESTADO).setValue('Opt-out');
      hCa.getRange(i + 2, REC_COL.PLAN).setValue('');
      hCa.getRange(i + 2, REC_COL.PROXIMO_TOQUE).setValue('');
      hCa.getRange(i + 2, REC_COL.CONSENT_WA).setValue('NO');
      hCa.getRange(i + 2, REC_COL.NOTAS).setValue(
        String(f[REC_COL.NOTAS - 1]) + ' | OPT-OUT ' +
        Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'));
      afectados++;
    }
  }
  return { tipo: tipo, identificador: normalizado, fichas: afectados };
}

/** Marca el consentimiento de WhatsApp cuando el candidato lo da de palabra o por escrito. */
function recMarcarConsentimientoWA(idCandidato, comoSeObtuvo) {
  const hCa = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REC.H_CANDIDATOS);
  const n = hCa.getLastRow() - 1;
  const ids = hCa.getRange(2, REC_COL.ID, n, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === String(idCandidato).trim()) {
      hCa.getRange(i + 2, REC_COL.CONSENT_WA).setValue('SI');
      hCa.getRange(i + 2, REC_COL.FECHA_CONSENT).setValue(new Date());
      const notas = hCa.getRange(i + 2, REC_COL.NOTAS).getValue();
      hCa.getRange(i + 2, REC_COL.NOTAS).setValue(
        notas + ' | Consentimiento WA: ' + (comoSeObtuvo || 'verbal en llamada') + ' ' +
        Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd'));
      return true;
    }
  }
  return false;
}

function recPurgarRetencion() {
  const ui = SpreadsheetApp.getUi();
  const hCa = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REC.H_CANDIDATOS);
  if (!hCa || hCa.getLastRow() < 2) { ui.alert('No hay datos.'); return; }

  const dias = parseInt(recLeerConfig_('RETENCION_DIAS', String(REC.RETENCION_DIAS)), 10);
  const limite = new Date(); limite.setDate(limite.getDate() - dias);

  const n = hCa.getLastRow() - 1;
  const datos = hCa.getRange(2, 1, n, REC_N_COLS).getValues();
  const aBorrar = [];

  for (let i = 0; i < datos.length; i++) {
    const f = datos[i];
    const estado = String(f[REC_COL.ESTADO - 1]);
    if (['Incorporado', 'Entrevistado', 'Entrevista agendada', 'Career Visioning', 'Oferta'].indexOf(estado) !== -1) continue;
    const ref = f[REC_COL.ULTIMO_TOQUE - 1] instanceof Date ? f[REC_COL.ULTIMO_TOQUE - 1]
              : (f[REC_COL.FECHA_CAPTURA - 1] instanceof Date ? f[REC_COL.FECHA_CAPTURA - 1] : null);
    if (ref && ref < limite) aBorrar.push(i + 2);
  }

  if (!aBorrar.length) { ui.alert('✅ Nada que purgar', 'Ningún candidato supera los ' + dias + ' días sin actividad.', ui.ButtonSet.OK); return; }

  const conf = ui.alert('🗑️ Purgar por retención',
    'Voy a borrar ' + aBorrar.length + ' candidatos sin actividad desde hace más de ' + dias + ' días.\n\n' +
    'Es la obligación de limitación del plazo de conservación (art. 5.1.e RGPD).\n' +
    'Los que están en proceso (entrevistados, oferta, incorporados) NO se tocan.\n\n' +
    '¿Continuar? Esto no se puede deshacer.', ui.ButtonSet.YES_NO);
  if (conf !== ui.Button.YES) return;

  aBorrar.reverse().forEach(fila => hCa.deleteRow(fila));
  ui.alert('✅ Purga completada', 'Registros eliminados: ' + aBorrar.length, ui.ButtonSet.OK);
}

function recSembrarRGPD_(ss) {
  const hoja = ss.getSheetByName(REC.H_RGPD);
  if (hoja.getLastRow() > 1) return;

  const filas = [
    ['ACTIVIDAD DE TRATAMIENTO',
     'Capacitación y selección de agentes inmobiliarios para incorporación al Market Center de ' + REC.MC_NOMBRE + '.'],

    ['RESPONSABLE',
     '[Razón social del Market Center] · NIF [—] · [Dirección] · [email] — RELLENAR'],

    ['CATEGORÍAS DE DATOS',
     'Identificativos (nombre y apellidos), de contacto profesional (teléfono, email, perfil de LinkedIn), ' +
     'profesionales (agencia actual, cargo, años de experiencia, idiomas, zona de actividad) y ' +
     'métricas de actividad pública (número de inmuebles publicados, rango de precio). ' +
     'NO se tratan categorías especiales de datos (art. 9 RGPD) ni datos de menores.'],

    ['ORIGEN DE LOS DATOS',
     'Fuentes de acceso público de carácter profesional: páginas web corporativas de agencias, ' +
     'Google Places, portales inmobiliarios donde el propio agente publica sus datos de contacto ' +
     'profesional, perfiles profesionales de LinkedIn, y referencias de terceros. ' +
     'El origen concreto queda registrado por cada ficha en las columnas Fuente y URL_Fuente.'],

    ['BASE JURÍDICA',
     'Interés legítimo del responsable (art. 6.1.f RGPD) en la capacitación de profesionales del sector, ' +
     'en relación con la presunción del art. 19 LOPDGDD para datos de contacto profesional. ' +
     'La ponderación consta en el campo PONDERACIÓN de esta hoja. ' +
     'Para el canal WhatsApp y para el envío de comunicaciones electrónicas con contenido promocional ' +
     'se recaba CONSENTIMIENTO previo (art. 6.1.a RGPD y art. 21 LSSI-CE), registrado en las columnas ' +
     'Consentimiento_WA y Fecha_Consentimiento.'],

    ['PONDERACIÓN DE INTERÉS LEGÍTIMO',
     'A) Interés perseguido: contactar profesionales del sector inmobiliario para ofrecerles una ' +
     'oportunidad de desarrollo profesional. Interés lícito, real y actual.\n' +
     'B) Necesidad: no existe alternativa menos intrusiva, ya que el contacto directo es inherente ' +
     'a la captación de talento. Se limitan los datos al mínimo imprescindible para el contacto profesional.\n' +
     'C) Equilibrio: los datos se obtienen de fuentes que el propio interesado ha hecho públicas en su ' +
     'condición profesional y con la finalidad de ser contactado profesionalmente. No se tratan datos de ' +
     'su esfera privada. La expectativa razonable de un agente inmobiliario que publica su teléfono ' +
     'profesional incluye recibir contactos profesionales.\n' +
     'D) Garantías aplicadas: (1) información del art. 14 RGPD en el primer contacto escrito; ' +
     '(2) derecho de oposición de ejercicio inmediato y sin justificación en cada mensaje; ' +
     '(3) lista de supresión que bloquea todos los canales; (4) plazo de conservación limitado; ' +
     '(5) no se realiza elaboración de perfiles con efectos jurídicos; (6) no hay cesiones a terceros.\n' +
     'CONCLUSIÓN: prevalece el interés legítimo, condicionado al mantenimiento de las garantías. ' +
     'REVISAR CON ASESORÍA JURÍDICA ANTES DE LA PRIMERA CAMPAÑA.'],

    ['INFORMACIÓN ART. 14 RGPD (incluir en el primer contacto escrito)',
     'Tus datos profesionales de contacto los hemos obtenido de [FUENTE]. Los trata ' +
     '[RAZÓN SOCIAL] con la única finalidad de valorar contigo una oportunidad de desarrollo ' +
     'profesional, sobre la base de nuestro interés legítimo (art. 6.1.f RGPD). No los cedemos a ' +
     'nadie. Puedes acceder, rectificar, suprimir, oponerte al tratamiento y solicitar su limitación ' +
     'o portabilidad escribiendo a [EMAIL]. También puedes reclamar ante la AEPD (www.aepd.es). ' +
     'Los conservaremos un máximo de [X] meses desde el último contacto. ' +
     'Si no quieres recibir más comunicaciones, responde "BAJA" y te eliminaremos de inmediato.'],

    ['PLAZO DE CONSERVACIÓN',
     'Máximo ' + REC.RETENCION_DIAS + ' días desde el último contacto efectivo para candidatos que no avanzan. ' +
     'Los candidatos en proceso activo se conservan hasta la resolución del mismo. ' +
     'Purga ejecutable desde el menú (función recPurgarRetencion).'],

    ['MEDIDAS DE SEGURIDAD',
     'Datos alojados en Google Workspace con control de acceso por cuenta corporativa. ' +
     'Claves de API en PropertiesService, nunca en el código ni en el repositorio. ' +
     'Registro de auditoría de todos los contactos en la hoja Rec_Toques. ' +
     'PENDIENTE: restringir el acceso a la hoja de cálculo solo a la Team Leader y al equipo de dirección.'],

    ['DECISIONES AUTOMATIZADAS',
     'El sistema calcula una puntuación de prioridad de contacto. No produce efectos jurídicos ni ' +
     'afecta significativamente a los interesados: solo ordena el trabajo comercial interno. ' +
     'Toda decisión de contacto, entrevista o contratación la toma una persona. ' +
     'No constituye elaboración de perfiles del art. 22 RGPD.'],

    ['⚠️ ADVERTENCIA IMPORTANTE',
     'Esta hoja es un punto de partida, NO un dictamen jurídico. Antes de la primera campaña:\n' +
     '1. Que tu asesor de protección de datos valide la ponderación.\n' +
     '2. Incorpora esta actividad a tu Registro de Actividades de Tratamiento (art. 30 RGPD).\n' +
     '3. Riesgo concreto a cubrir: la AEPD ha sancionado el envío de comunicaciones electrónicas ' +
     'a profesionales sin consentimiento cuando se consideran comunicaciones comerciales (art. 21 LSSI). ' +
     'Por eso en este sistema el primer contacto es TELEFÓNICO y el WhatsApp y el email llegan después. ' +
     'No invierta ese orden sin asesoramiento.']
  ];

  hoja.getRange(2, 1, filas.length, 2).setValues(filas);
  hoja.setColumnWidth(1, 300).setColumnWidth(2, 760);
  hoja.getRange(2, 1, filas.length, 2).setWrap(true).setVerticalAlignment('top');
  hoja.getRange(2, 1, filas.length, 1).setFontWeight('bold');
}

// ============================================================
//  16. PANEL DIARIO (servidor)
// ============================================================

function recAbrirPanel() {
  const html = HtmlService.createHtmlOutputFromFile('panel_reclutamiento')
    .setWidth(1180).setHeight(720);
  SpreadsheetApp.getUi().showModalDialog(html, '🎯 Reclutamiento · Panel de la Team Leader');
}

/** Todo lo que el panel necesita en una sola llamada. */
function recDatosPanel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hCa = ss.getSheetByName(REC.H_CANDIDATOS);
  const hTo = ss.getSheetByName(REC.H_TOQUES);
  const hEn = ss.getSheetByName(REC.H_ENTREVISTAS);

  const resultado = {
    cola: [], embudo: {}, temperaturas: { A: 0, B: 0, C: 0, D: 0 },
    fuentes: {}, totales: { candidatos: 0, enSecuencia: 0, toquesHoy: 0, toquesPendientes: 0 },
    entrevistas: [], tl: recLeerConfig_('TL_NOMBRE', REC.TL_NOMBRE)
  };

  // --- Cola de toques pendientes ---
  if (hTo && hTo.getLastRow() > 1) {
    const datos = hTo.getRange(2, 1, hTo.getLastRow() - 1, 13).getValues();
    const hoy = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const infoCand = recMapaCandidatos_(hCa);

    for (let i = datos.length - 1; i >= 0 && resultado.cola.length < 60; i--) {
      const f = datos[i];
      if (String(f[8]).toUpperCase() !== 'PENDIENTE') continue;
      const c = infoCand[String(f[1])] || {};
      const msg = String(f[10]);
      const partes = msg.split('\n\n▶ ');
      resultado.cola.push({
        idToque: f[0], idCand: f[1], nombre: f[2],
        fecha: f[3] instanceof Date ? Utilities.formatDate(f[3], Session.getScriptTimeZone(), 'dd/MM') : '',
        canal: f[5], paso: f[6], tipo: f[7],
        mensaje: partes[0], enlace: partes[1] || '',
        objetivo: f[12],
        agencia: c.agencia || '', zona: c.zona || '', score: c.score || 0,
        temperatura: c.temperatura || '', telefono: c.telefono || '',
        inmuebles: c.inmuebles || '', idiomas: c.idiomas || '',
        consentWA: c.consentWA || '', linkedin: c.linkedin || ''
      });
    }
    resultado.totales.toquesPendientes = datos.filter(f => String(f[8]).toUpperCase() === 'PENDIENTE').length;
    resultado.totales.toquesHoy = datos.filter(f =>
      f[3] instanceof Date &&
      Utilities.formatDate(f[3], Session.getScriptTimeZone(), 'yyyy-MM-dd') === hoy).length;
  }

  // --- Embudo y métricas ---
  if (hCa && hCa.getLastRow() > 1) {
    const datos = hCa.getRange(2, 1, hCa.getLastRow() - 1, REC_N_COLS).getValues();
    resultado.totales.candidatos = datos.length;
    REC_ESTADOS.forEach(e => { resultado.embudo[e] = 0; });
    datos.forEach(f => {
      const e = String(f[REC_COL.ESTADO - 1]) || 'Nuevo';
      resultado.embudo[e] = (resultado.embudo[e] || 0) + 1;
      const t = String(f[REC_COL.TEMPERATURA - 1]);
      if (resultado.temperaturas[t] !== undefined) resultado.temperaturas[t]++;
      const fu = String(f[REC_COL.FUENTE - 1]) || 'Sin fuente';
      resultado.fuentes[fu] = (resultado.fuentes[fu] || 0) + 1;
      if (String(f[REC_COL.PLAN - 1]).trim()) resultado.totales.enSecuencia++;
    });
  }

  // --- Próximas entrevistas ---
  if (hEn && hEn.getLastRow() > 1) {
    const datos = hEn.getRange(2, 1, hEn.getLastRow() - 1, 17).getValues();
    const hoy = new Date(); hoy.setHours(0, 0, 0, 0);
    resultado.entrevistas = datos
      .filter(f => f[3] instanceof Date && f[3] >= hoy)
      .sort((a, b) => a[3] - b[3])
      .slice(0, 10)
      .map(f => ({
        nombre: f[2],
        fecha: Utilities.formatDate(f[3], Session.getScriptTimeZone(), 'dd/MM HH:mm'),
        fase: f[4], entrevistador: f[5], resultado: f[13], siguiente: f[14]
      }));
  }

  return resultado;
}

function recMapaCandidatos_(hCa) {
  const mapa = {};
  if (!hCa || hCa.getLastRow() < 2) return mapa;
  hCa.getRange(2, 1, hCa.getLastRow() - 1, REC_N_COLS).getValues().forEach(f => {
    mapa[String(f[REC_COL.ID - 1])] = {
      agencia: f[REC_COL.AGENCIA - 1], zona: f[REC_COL.ZONA - 1],
      score: f[REC_COL.SCORE - 1], temperatura: f[REC_COL.TEMPERATURA - 1],
      telefono: f[REC_COL.TELEFONO - 1], inmuebles: f[REC_COL.INMUEBLES - 1],
      idiomas: f[REC_COL.IDIOMAS - 1], consentWA: f[REC_COL.CONSENT_WA - 1],
      linkedin: f[REC_COL.LINKEDIN - 1], email: f[REC_COL.EMAIL - 1]
    };
  });
  return mapa;
}

/**
 * Registra el resultado de un toque y mueve el pipeline en consecuencia.
 * Es la única acción que la TL hace en el panel: marcar qué ha pasado.
 */
function recMarcarResultadoToque(idToque, resultado, notas) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hTo = ss.getSheetByName(REC.H_TOQUES);
  const hCa = ss.getSheetByName(REC.H_CANDIDATOS);

  const n = hTo.getLastRow() - 1;
  if (n < 1) throw new Error('No hay toques registrados.');
  const ids = hTo.getRange(2, 1, n, 2).getValues();
  let fila = -1, idCand = '';
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(idToque)) { fila = i + 2; idCand = String(ids[i][1]); break; }
  }
  if (fila === -1) throw new Error('No encuentro ese toque.');

  hTo.getRange(fila, 9).setValue('HECHO');
  hTo.getRange(fila, 10).setValue(resultado);
  if (notas) hTo.getRange(fila, 13).setValue(notas);

  // Efectos en la ficha del candidato
  const mapa = { nuevoEstado: null, pararPlan: false, consentWA: false, optOut: false };
  switch (resultado) {
    case 'Contactado — interesado':      mapa.nuevoEstado = 'Conversación activa'; mapa.consentWA = true; break;
    case 'Contactado — no interesado':   mapa.nuevoEstado = 'No ahora (nurture)'; break;
    case 'Entrevista agendada':          mapa.nuevoEstado = 'Entrevista agendada'; mapa.pararPlan = true; mapa.consentWA = true; break;
    case 'No contesta':                  break;  // sigue la secuencia
    case 'Número erróneo':               mapa.nuevoEstado = 'Descartado'; mapa.pararPlan = true; break;
    case 'Pide la baja':                 mapa.optOut = true; break;
    case 'Enviado':                      break;
    default: break;
  }

  if (!idCand) return { ok: true };

  const nCa = hCa.getLastRow() - 1;
  const idsCa = hCa.getRange(2, REC_COL.ID, nCa, 1).getValues();
  for (let i = 0; i < idsCa.length; i++) {
    if (String(idsCa[i][0]) !== idCand) continue;
    const r = i + 2;

    if (mapa.optOut) {
      const tel = hCa.getRange(r, REC_COL.TELEFONO).getValue();
      const email = hCa.getRange(r, REC_COL.EMAIL).getValue();
      recRegistrarOptOut(tel || email, 'Solicitada en toque ' + idToque,
        hCa.getRange(r, REC_COL.NOMBRE).getValue());
      return { ok: true, optOut: true };
    }
    if (mapa.nuevoEstado) hCa.getRange(r, REC_COL.ESTADO).setValue(mapa.nuevoEstado);
    if (mapa.pararPlan) { hCa.getRange(r, REC_COL.PLAN).setValue(''); hCa.getRange(r, REC_COL.PROXIMO_TOQUE).setValue(''); }
    if (mapa.consentWA && String(hCa.getRange(r, REC_COL.CONSENT_WA).getValue()).toUpperCase() !== 'SI') {
      hCa.getRange(r, REC_COL.CONSENT_WA).setValue('SI');
      hCa.getRange(r, REC_COL.FECHA_CONSENT).setValue(new Date());
    }
    if (notas) {
      const prev = hCa.getRange(r, REC_COL.NOTAS).getValue();
      hCa.getRange(r, REC_COL.NOTAS).setValue((prev ? prev + ' | ' : '') + notas);
    }
    break;
  }
  return { ok: true };
}

/** Crea la entrevista y saca al candidato de la secuencia de captación. */
function recAgendarEntrevista(idCandidato, fechaIso, fase, entrevistador) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hEn = ss.getSheetByName(REC.H_ENTREVISTAS);
  const hCa = ss.getSheetByName(REC.H_CANDIDATOS);

  const n = hCa.getLastRow() - 1;
  const datos = hCa.getRange(2, 1, n, REC_N_COLS).getValues();
  let nombre = '', filaCand = -1;
  for (let i = 0; i < datos.length; i++) {
    if (String(datos[i][REC_COL.ID - 1]) === String(idCandidato)) {
      nombre = (datos[i][REC_COL.NOMBRE - 1] + ' ' + datos[i][REC_COL.APELLIDOS - 1]).trim();
      filaCand = i + 2;
      break;
    }
  }
  if (filaCand === -1) throw new Error('Candidato no encontrado.');

  hEn.appendRow([recNuevoId_('E'), idCandidato, nombre, new Date(fechaIso),
    fase || '1ª entrevista', entrevistador || recLeerConfig_('TL_NOMBRE', REC.TL_NOMBRE),
    '', '', '', '', '', '', '', 'Pendiente', '', '', '']);

  hCa.getRange(filaCand, REC_COL.ESTADO).setValue('Entrevista agendada');
  hCa.getRange(filaCand, REC_COL.PLAN).setValue('');
  hCa.getRange(filaCand, REC_COL.PROXIMO_TOQUE).setValue('');
  return { ok: true, nombre: nombre };
}

// ============================================================
//  17. EXPORTAR A COMMANDMC
// ============================================================

/**
 * CommandMC (Command de liderazgo) tiene su propio módulo de Recruits
 * con SmartPlans de email/SMS/tareas. Lo que NO tiene es WhatsApp.
 *
 * Reparto recomendado:
 *   • Este sistema  → captación, scoring y toda la capa de WhatsApp y llamada.
 *   • CommandMC     → expediente oficial del recruit, email y tareas del equipo.
 *
 * Esta función genera el CSV para importar en Recruit Management.
 */
function recExportarCommandMC() {
  const ui = SpreadsheetApp.getUi();
  const hCa = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REC.H_CANDIDATOS);
  if (!hCa || hCa.getLastRow() < 2) { ui.alert('No hay candidatos.'); return; }

  const r = ui.alert('📤 Exportar a CommandMC',
    'Exporta los candidatos de temperatura A y B que ya están en conversación.\n\n' +
    'Se genera un CSV en tu Google Drive listo para subir a Recruit Management.\n\n' +
    '¿Continuar?', ui.ButtonSet.YES_NO);
  if (r !== ui.Button.YES) return;

  const datos = hCa.getRange(2, 1, hCa.getLastRow() - 1, REC_N_COLS).getValues();
  const estadosValidos = ['Conversación activa', 'Entrevista agendada', 'Entrevistado',
                          'Career Visioning', 'Oferta', 'Contactado'];

  const filas = [['First Name', 'Last Name', 'Email', 'Phone', 'Current Company',
                  'Title', 'City', 'Source', 'Notes', 'Stage']];

  datos.forEach(f => {
    const temp = String(f[REC_COL.TEMPERATURA - 1]);
    const estado = String(f[REC_COL.ESTADO - 1]);
    if (temp !== 'A' && temp !== 'B') return;
    if (estadosValidos.indexOf(estado) === -1) return;
    if (!String(f[REC_COL.EMAIL - 1]).trim() && !String(f[REC_COL.TELEFONO - 1]).trim()) return;

    filas.push([
      f[REC_COL.NOMBRE - 1], f[REC_COL.APELLIDOS - 1],
      f[REC_COL.EMAIL - 1], f[REC_COL.TELEFONO - 1],
      f[REC_COL.AGENCIA - 1], f[REC_COL.CARGO - 1], f[REC_COL.ZONA - 1],
      f[REC_COL.FUENTE - 1],
      'Score ' + f[REC_COL.SCORE - 1] + ' (' + temp + ') · ' +
        f[REC_COL.INMUEBLES - 1] + ' inmuebles · ' + f[REC_COL.IDIOMAS - 1] + ' · ' +
        String(f[REC_COL.NOTAS - 1]).substring(0, 180),
      estado
    ]);
  });

  if (filas.length === 1) { ui.alert('Nada que exportar', 'Ningún candidato cumple los criterios (A/B en conversación).', ui.ButtonSet.OK); return; }

  const csv = filas.map(fila => fila.map(c => {
    const v = String(c === null || c === undefined ? '' : c).replace(/"/g, '""');
    return '"' + v + '"';
  }).join(',')).join('\n');

  const nombre = 'KW_Marbella_Recruits_' +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HHmm') + '.csv';
  const archivo = DriveApp.createFile(nombre, csv, MimeType.CSV);

  ui.alert('✅ CSV generado',
    'Registros exportados: ' + (filas.length - 1) + '\n\n' +
    'Archivo: ' + nombre + '\n' +
    'Está en la raíz de tu Google Drive:\n' + archivo.getUrl() + '\n\n' +
    'Súbelo en CommandMC → Recruits → Recruit Management → Import.',
    ui.ButtonSet.OK);
}

// ============================================================
//  18. AUTOMATIZACIÓN DIARIA
// ============================================================

function recInstalarTriggerDiario() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'recRutinaDiaria')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('recRutinaDiaria')
    .timeBased().atHour(7).everyDays(1)
    .inTimezone(Session.getScriptTimeZone())
    .create();

  SpreadsheetApp.getUi().alert('⏰ Automatización activada',
    'Cada día a las 7:00 el sistema:\n' +
    '• Inscribe en plan los candidatos nuevos\n' +
    '• Recalcula puntuaciones\n' +
    '• Genera la cola de toques del día\n' +
    '• Avisa por email si hay toques pendientes\n\n' +
    'La Team Leader solo tiene que abrir el panel y trabajar la cola.',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

function recRutinaDiaria() {
  try {
    recRecalcularScores(true);
    recAutoAsignarPlanes();
    recGenerarToques_(true);

    const datos = recDatosPanel();
    if (datos.totales.toquesPendientes > 0) {
      const destino = recLeerConfig_('MC_EMAIL', '') || Session.getEffectiveUser().getEmail();
      MailApp.sendEmail({
        to: destino,
        subject: '🎯 Reclutamiento KW Marbella · ' + datos.totales.toquesPendientes + ' toques pendientes',
        htmlBody:
          '<h2 style="font-family:Arial">Cola de hoy</h2>' +
          '<p style="font-family:Arial">Toques pendientes: <b>' + datos.totales.toquesPendientes + '</b><br>' +
          'Candidatos en secuencia: <b>' + datos.totales.enSecuencia + '</b><br>' +
          'Temperatura A (prioritarios): <b>' + datos.temperaturas.A + '</b></p>' +
          '<h3 style="font-family:Arial">Embudo</h3><ul style="font-family:Arial">' +
          Object.keys(datos.embudo).filter(k => datos.embudo[k] > 0)
            .map(k => '<li>' + k + ': ' + datos.embudo[k] + '</li>').join('') +
          '</ul>' +
          '<p style="font-family:Arial"><a href="' +
          SpreadsheetApp.getActiveSpreadsheet().getUrl() +
          '">Abrir el panel de reclutamiento</a></p>'
      });
    }
  } catch (e) {
    Logger.log('Error en recRutinaDiaria: ' + e.message);
  }
}

// ============================================================
//  19. DIÁLOGOS HTML
// ============================================================

const REC_CSS_DIALOGO = `
  * { box-sizing: border-box; }
  body { font-family: -apple-system, "Segoe UI", Roboto, Arial, sans-serif;
         margin: 0; padding: 18px; background: #0f172a; color: #f1f5f9; font-size: 13px; }
  h2 { margin: 0 0 4px; font-size: 17px; }
  p.sub { margin: 0 0 14px; color: #94a3b8; font-size: 12px; line-height: 1.5; }
  label { display: block; margin: 12px 0 5px; font-weight: 600; font-size: 12px; }
  textarea, input, select {
    width: 100%; padding: 9px 10px; border-radius: 7px; font-size: 12px;
    border: 1px solid rgba(255,255,255,.15); background: rgba(255,255,255,.06);
    color: #f1f5f9; font-family: inherit; }
  textarea { min-height: 170px; resize: vertical; font-family: ui-monospace, Menlo, monospace; }
  button { background: linear-gradient(135deg,#dc2626,#991b1b); color: #fff; border: 0;
    padding: 11px 20px; border-radius: 7px; font-weight: 700; cursor: pointer; font-size: 13px; }
  button:disabled { opacity: .5; cursor: wait; }
  button.sec { background: rgba(255,255,255,.1); }
  .fila { display: flex; gap: 10px; align-items: center; margin-top: 16px; }
  .aviso { background: rgba(245,158,11,.12); border-left: 3px solid #f59e0b;
           padding: 10px 12px; border-radius: 0 6px 6px 0; margin: 12px 0;
           font-size: 11.5px; line-height: 1.55; color: #fcd34d; }
  .ok { background: rgba(34,197,94,.12); border-left: 3px solid #22c55e; color: #86efac;
        padding: 10px 12px; border-radius: 0 6px 6px 0; margin: 12px 0; font-size: 12px; }
  #res { margin-top: 14px; }
`;

function recHtmlImportador_() {
  return `<!DOCTYPE html><html><head><meta charset="utf-8"><style>${REC_CSS_DIALOGO}</style></head><body>
    <h2>📥 Importar candidatos</h2>
    <p class="sub">Pega un CSV con cabecera (de LinkedIn Sales Navigator, de un Excel propio…)
    o texto suelto. Detecto las columnas solas; si es texto libre, lo interpreta la IA.</p>

    <div class="aviso"><b>No uses extensiones de scraping de LinkedIn.</b>
    Incumplen sus condiciones y la AEPD ya ha sancionado el uso de datos de perfiles públicos
    para contacto no consentido. Exporta desde Sales Navigator o copia y pega a mano.</div>

    <label>Origen (queda registrado en cada ficha, es obligatorio para el RGPD)</label>
    <select id="fuente">
      <option>LinkedIn Sales Navigator</option>
      <option>LinkedIn búsqueda manual</option>
      <option>Portal inmobiliario</option>
      <option>Evento / networking</option>
      <option>Referido de agente</option>
      <option>Referido de cliente</option>
      <option>Instagram</option>
      <option>Base propia anterior</option>
      <option>Otra</option>
    </select>

    <label>Formato</label>
    <select id="modo">
      <option value="csv">CSV / tabla con cabecera</option>
      <option value="texto">Texto libre (lo interpreta la IA)</option>
    </select>

    <label>Contenido</label>
    <textarea id="contenido" placeholder="First Name,Last Name,Company,Title,Location&#10;Ana,García,Panorama Properties,Real Estate Agent,Marbella&#10;..."></textarea>

    <div class="fila">
      <button id="btn" onclick="enviar()">Importar</button>
      <button class="sec" onclick="google.script.host.close()">Cerrar</button>
    </div>
    <div id="res"></div>

    <script>
      function enviar() {
        var c = document.getElementById('contenido').value.trim();
        if (!c) { alert('Pega algo primero.'); return; }
        var b = document.getElementById('btn');
        b.disabled = true; b.textContent = 'Procesando…';
        document.getElementById('res').innerHTML = '';
        google.script.run
          .withSuccessHandler(function (r) {
            b.disabled = false; b.textContent = 'Importar';
            document.getElementById('res').innerHTML =
              '<div class="ok">Detectados: <b>' + r.detectados + '</b> · ' +
              'Nuevos: <b>' + r.nuevos + '</b> · Duplicados omitidos: <b>' + r.duplicados + '</b>' +
              '<br>Ya están puntuados y ordenados en Rec_Candidatos.</div>';
            document.getElementById('contenido').value = '';
          })
          .withFailureHandler(function (e) {
            b.disabled = false; b.textContent = 'Importar';
            document.getElementById('res').innerHTML =
              '<div class="aviso">Error: ' + e.message + '</div>';
          })
          .recImportarTexto(c,
            document.getElementById('fuente').value,
            document.getElementById('modo').value);
      }
    </script></body></html>`;
}

function recHtmlPegadoPortal_() {
  return `<!DOCTYPE html><html><head><meta charset="utf-8"><style>${REC_CSS_DIALOGO}</style></head><body>
    <h2>📋 Pegar ficha de portal</h2>
    <p class="sub">Esta es la fuente con el dato más valioso: <b>cuántos inmuebles tiene publicados
    y en qué rango de precio</b>. Eso es producción real, no lo que pone en su LinkedIn.</p>

    <div class="aviso"><b>Por qué es manual y no automático.</b>
    Idealista, Fotocasa y el resto de portales prohíben el rastreo automatizado en sus
    condiciones de uso y en su robots.txt. Rastrearlos te expone a bloqueo de IP y a una
    reclamación. Pegando tú la página no hay rastreo: es una persona consultando una web
    pública, que es exactamente para lo que está publicada.</div>

    <label>Pasos</label>
    <p class="sub">1. Abre la ficha de la agencia en el portal (la que lista a su equipo o sus inmuebles).<br>
    2. <b>Ctrl+A</b> y <b>Ctrl+C</b> en la página.<br>
    3. Pega aquí abajo y dale a Procesar.</p>

    <label>Portal</label>
    <select id="fuente">
      <option>Idealista</option>
      <option>Fotocasa</option>
      <option>Kyero</option>
      <option>ThinkSpain</option>
      <option>Resales Online</option>
      <option>James Edition</option>
      <option>Otro portal</option>
    </select>

    <label>Contenido de la página</label>
    <textarea id="contenido" placeholder="Pega aquí el contenido copiado…"></textarea>

    <div class="fila">
      <button id="btn" onclick="enviar()">Procesar con IA</button>
      <button class="sec" onclick="google.script.host.close()">Cerrar</button>
    </div>
    <div id="res"></div>

    <script>
      function enviar() {
        var c = document.getElementById('contenido').value.trim();
        if (c.length < 100) { alert('Parece muy corto. Copia la página completa.'); return; }
        var b = document.getElementById('btn');
        b.disabled = true; b.textContent = 'Analizando…';
        google.script.run
          .withSuccessHandler(function (r) {
            b.disabled = false; b.textContent = 'Procesar con IA';
            document.getElementById('res').innerHTML =
              '<div class="ok">Agentes detectados: <b>' + r.detectados + '</b> · ' +
              'Añadidos: <b>' + r.nuevos + '</b> · Ya estaban: <b>' + r.duplicados + '</b></div>';
            document.getElementById('contenido').value = '';
          })
          .withFailureHandler(function (e) {
            b.disabled = false; b.textContent = 'Procesar con IA';
            document.getElementById('res').innerHTML = '<div class="aviso">Error: ' + e.message + '</div>';
          })
          .recImportarTexto(c, document.getElementById('fuente').value, 'texto');
      }
    </script></body></html>`;
}

function recHtmlSupresion_() {
  return `<!DOCTYPE html><html><head><meta charset="utf-8"><style>${REC_CSS_DIALOGO}</style></head><body>
    <h2>🔐 Opt-out / derecho de supresión</h2>
    <p class="sub">Si alguien contesta "BAJA", "STOP", pide que no le escribas o ejerce su
    derecho de oposición o supresión, regístralo <b>aquí y en el momento</b>.
    Queda bloqueado en todos los canales y se le para la secuencia.</p>

    <div class="aviso">Atender una baja tarde o mal es la vía más rápida a una sanción.
    La AEPD ha multado precisamente por seguir enviando comunicaciones a quien ya se había dado de baja.</div>

    <label>Teléfono, email o URL de LinkedIn</label>
    <input id="id" placeholder="+34600000000 · nombre@agencia.com · linkedin.com/in/...">

    <label>Nombre (opcional, solo para el registro)</label>
    <input id="nombre" placeholder="Ana García">

    <label>Motivo</label>
    <select id="motivo">
      <option>Solicitud de baja del interesado</option>
      <option>Derecho de oposición (art. 21 RGPD)</option>
      <option>Derecho de supresión (art. 17 RGPD)</option>
      <option>No interesado — no volver a contactar</option>
      <option>Número o email erróneo</option>
    </select>

    <div class="fila">
      <button id="btn" onclick="enviar()">Registrar baja</button>
      <button class="sec" onclick="google.script.host.close()">Cerrar</button>
    </div>
    <div id="res"></div>

    <script>
      function enviar() {
        var id = document.getElementById('id').value.trim();
        if (!id) { alert('Indica el identificador.'); return; }
        var b = document.getElementById('btn');
        b.disabled = true; b.textContent = 'Registrando…';
        google.script.run
          .withSuccessHandler(function (r) {
            b.disabled = false; b.textContent = 'Registrar baja';
            document.getElementById('res').innerHTML =
              '<div class="ok">Bloqueado: <b>' + r.identificador + '</b> (' + r.tipo + ')<br>' +
              'Fichas marcadas como Opt-out: <b>' + r.fichas + '</b></div>';
            document.getElementById('id').value = '';
            document.getElementById('nombre').value = '';
          })
          .withFailureHandler(function (e) {
            b.disabled = false; b.textContent = 'Registrar baja';
            document.getElementById('res').innerHTML = '<div class="aviso">Error: ' + e.message + '</div>';
          })
          .recRegistrarOptOut(id,
            document.getElementById('motivo').value,
            document.getElementById('nombre').value);
      }
    </script></body></html>`;
}
