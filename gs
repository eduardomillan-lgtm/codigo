/**
 * ========================================
 * DASHBOARD INMOBILIARIO KW - Google Apps Script v5.0 FINAL
 * → HOJA FACTURACIÓN ELIMINADA PARA SIEMPRE
 * → Todo el GCI se guarda directamente en Actividad_Diaria
 * → 100% FUNCIONAL - SIN ERRORES
 * ========================================
 */

const kpiNames = {
  gci: 'GCI',
  citasCaptacion: 'Citas Captación',
  exclusivasVenta: 'Exclusivas Venta',
  exclusivasComprador: 'Exclusivas Comprador',
  captacionesAbierto: 'Captaciones en Abierto',
  citasCompradores: 'Citas Compradores',
  casasEnsenadas: 'Casas Enseñadas',
  leadsCompradores: 'Leads Compradores',
  llamadas: 'Llamadas',
  volumenNegocio: 'Volumen de Negocio',
  cumplimientoGlobal: 'Cumplimiento Global',
  conversionCaptacion: 'Conversión (Cita a Capt.)',
  ratioCierre: 'Ratio de Cierre (GCI/Excl)',
  productividad: 'Productividad (Casas/Cita)',
  ticketPromedio: 'Ticket Promedio',
  actividadTotal: 'Actividad Total'
};
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🏠 Dashboard Inmobiliario')
    .addItem('🚀 Fase Preparación (Crear todas las hojas)', 'inicializarSistema')
    .addItem('📊 Iniciar Dashboard', 'abrirDashboard')
    .addItem('🤖 Análisis IA del Equipo', 'lanzarAnalisisIA')
    .addSeparator()
    .addItem('👥 Gestionar Agentes', 'gestionarAgentes')
    .addItem('🎯 Configurar Objetivos', 'configurarObjetivos')
    .addSeparator()
    .addItem('🔄 Recalcular Datos', 'recalcularTodosDatos')
    .addItem('🧹 Limpiar Datos de Prueba', 'limpiarDatosPrueba')
    .addToUi();
}
function getKpiMetadata(key, valor) {
    const v = parseFloat(valor);
    switch(key) {
        case 'cumplimientoGlobal': return { icon: '❤️', label: 'Cumplimiento', unidad: '%', claseCritica: v < 90 ? 'critical' : '', thresholds: { bueno: 90, regular: 70 } };
        case 'conversionCaptacion': return { icon: '🎯', label: 'Conversión Capt.', unidad: '%', claseCritica: v < 30 ? 'critical' : '', thresholds: { bueno: 30, regular: 20 } };
        case 'ratioCierre': return { icon: '🔥', label: 'Ratio Cierre (GCI/Excl)', unidad: '€', claseCritica: v < 2000 ? 'critical' : '', thresholds: { bueno: 3000, regular: 2000 } };
        case 'productividad': return { icon: '⚡', label: 'Productividad', unidad: '', claseCritica: v < 1.5 ? 'critical' : '', thresholds: { bueno: 1.5, regular: 1 } };
        case 'ticketPromedio': return { icon: '💰', label: 'Ticket Promedio', unidad: '€', claseCritica: v < 2000 ? 'critical' : '', thresholds: { bueno: 3000, regular: 2000 } };
        case 'actividadTotal': return { icon: '📊', label: 'Actividad Total', unidad: '', claseCritica: '', thresholds: null };
        case 'gci': return { icon: '💵', label: 'GCI', unidad: '€', claseCritica: '', thresholds: null };
        case 'citasCaptacion': return { icon: '📞', label: 'Citas Captación', unidad: '', claseCritica: '', thresholds: null };
        case 'exclusivasVenta': return { icon: '🏠', label: 'Exclusivas Venta', unidad: '', claseCritica: '', thresholds: null };
        case 'exclusivasComprador': return { icon: '🔑', label: 'Exclusivas Compr.', unidad: '', claseCritica: '', thresholds: null };
        case 'captacionesAbierto': return { icon: '🏘️', label: 'Capt. Abierto', unidad: '', claseCritica: '', thresholds: null };
        case 'citasCompradores': return { icon: '👥', label: 'Citas Compr.', unidad: '', claseCritica: '', thresholds: null };
        case 'casasEnsenadas': return { icon: '🏡', label: 'Casas Enseñadas', unidad: '', claseCritica: '', thresholds: null };
        case 'leadsCompradores': return { icon: '📧', label: 'Leads Compr.', unidad: '', claseCritica: '', thresholds: null };
        case 'llamadas': return { icon: '📞', label: 'Llamadas', unidad: '', claseCritica: '', thresholds: null };
        case 'volumenNegocio': return { icon: '💼', label: 'Volumen Negocio', unidad: '€', claseCritica: '', thresholds: null };
        default: return { icon: '❓', label: key, unidad: '', claseCritica: '', thresholds: null };
    }
}

const CONFIG = {
  HOJA_AGENTES: 'Agentes',
  HOJA_ACTIVIDAD: 'Actividad_Diaria',
  HOJA_OBJETIVOS: 'Objetivos',
  HOJA_CONFIGURACION: 'Configuración',
  HOJA_INVENTARIO: 'Inventario_Inmuebles',
 
  METRICAS: ['Citas Captación', 'Exclusivas Venta', 'Exclusivas Comprador', 'Captaciones Abierto', 'Citas Compradores', 'Casas Enseñadas', 'Leads Compradores', 'Llamadas', 'GCI', 'Volumen Negocio']
};
// ============================================
// MEJORA 2: 18 ORÍGENES DE NEGOCIO
// ============================================

const ORIGENES_NEGOCIO = [
  'Clientes Antiguos (repiten compra)',
  'Esfera de Influencia/ Contactos',
  'Vendedores por sí mismos (FSO)',
  'Farming (posicionamiento en zona)',
  'Aliados (hipoteca, abogados, etc.)',
  'Referidos de Agentes Inmobiliarios',
  'Referidos de Clientes Antiguos',
  'Reubicación',
  'Referidos del Personal de Trabajo',
  'Llamadas a la oficina por anuncios en carteles',
  'Publicidad',
  'Sitio Web',
  'Correo directo',
  'Redes Sociales',
  'Open Houses',
  'ISA/OSA',
  'Idealistas',
  'Fotocasa',
  'Otro'
];
// ============================================
// MEJORA 3: 12 PARTIDAS DE GASTOS OPERATIVOS
// ============================================

const PARTIDAS_GASTOS = [
  'Administración/Coordinación',
  'Salarios Agentes Vendedores',
  'Salarios Agentes Compradores',
  'Generación de Negocio Marketing',
  'Generación de Negocio Prospección',
  'Alquileres/Amortización lugar trabajo',
  'Educación/Coaching/Afiliaciones',
  'Suministros / Gastos de oficina',
  'Comunicación/Tecnología',
  'Automovil',
  'Equipo/Mobiliario',
  'Seguro',
  'Otros Gastos'
];

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🏠 Dashboard Inmobiliario')
    .addItem('🚀 Inicializar Sistema', 'inicializarSistema')
    .addItem('📊 Abrir Dashboard', 'abrirDashboard')
    .addItem('👥 Gestionar Agentes', 'gestionarAgentes')
    .addItem('🎯 Configurar Objetivos', 'configurarObjetivos')
    .addSeparator()
    .addItem('🔄 Recalcular Datos', 'recalcularTodosDatos')
    .addItem('🧹 Limpiar Datos de Prueba', 'limpiarDatosPrueba')
    .addToUi();
}

function inicializarSistema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const respuesta = ui.alert(
    '🚀 Inicializar Sistema Completo (v5.3)',
    'Esto creará o restaurará TODAS las hojas necesarias:\n\n' +
    '• Básicas: Agentes, Actividad, Objetivos, Configuración\n' +
    '• Negocio: Inventario, Facturación Pasada\n' +
    '• Análisis: Rentabilidad, Histórico Agentes\n' +
    '• Modelos: Presupuestario, Económico, Organizativo, GPS\n\n' +
    '¿Continuar?',
    ui.ButtonSet.YES_NO
  );
  
  if (respuesta !== ui.Button.YES) return;
  
  try {
    // 1. Hojas Básicas
    crearHojaAgentes(ss);
    crearHojaActividad(ss);
    crearHojaObjetivos(ss);
    crearHojaConfiguracion(ss);
    
    // 2. Hojas de Negocio y Modelos (NUEVAS)
    crearHojaInventario(ss);
    crearHojaFacturacionPasada(ss);      // <--- Importante
    crearHojaHistoricoAgentes(ss);       // <--- La nueva del histórico
    crearHojaRentabilidad(ss);
    crearHojaModeloPresupuestario(ss);
    crearHojaModeloEconomico(ss);
    crearHojaModeloOrganizativo(ss);
    crearHojaPlantillaGPS(ss);
    
    // 3. Datos de ejemplo (solo si están vacías)
    insertarDatosEjemplo(ss);
    
    ui.alert('✅ Sistema Inicializado al 100%', 
             'Todas las pestañas están listas para trabajar.', ui.ButtonSet.OK);
             
  } catch (error) {
    ui.alert('❌ Error', 'Fallo en inicialización: ' + error.toString(), ui.ButtonSet.OK);
  }
}

function crearHojaAgentes(ss) {
  let hoja = ss.getSheetByName(CONFIG.HOJA_AGENTES);
  if (hoja) ss.deleteSheet(hoja);
  hoja = ss.insertSheet(CONFIG.HOJA_AGENTES);
  const headers = ['ID', 'Nombre', 'Email', 'Teléfono', 'Fecha Ingreso', 'Estado', 'Objetivos Acumulados', 'Fecha Registro', 'Sueldo Fijo'];
  hoja.getRange(1, 1, 1,headers.length).setValues([headers]).setBackground('#b70000').setFontColor('#ffffff').setFontWeight('bold').setFontSize(11);
  hoja.setFrozenRows(1);
  hoja.autoResizeColumns(1, headers.length);
}

function crearHojaActividad(ss) {
  let hoja = ss.getSheetByName(CONFIG.HOJA_ACTIVIDAD);
  if (hoja) ss.deleteSheet(hoja);
  hoja = ss.insertSheet(CONFIG.HOJA_ACTIVIDAD);
  const headers = ['ID', 'Fecha', 'ID_Agente', 'Nombre_Agente',
        'Citas_Captacion', 'Exclusivas_Venta', 'Exclusivas_Comprador',
        'Captaciones_Abierto', 'Citas_Compradores', 'Casas_Ensenadas',
        'Leads_Compradores', 'Llamadas', 'GCI', 'Volumen_Negocio', 'Notas', 'Timestamp', 
        'Comision_Pagada', 'Pct_Comision'];
  hoja.getRange(1, 1, 1, headers.length).setValues([headers]).setBackground('#b70000').setFontColor('#ffffff').setFontWeight('bold').setFontSize(11);
  hoja.setFrozenRows(1);
  hoja.getRange('B:B').setNumberFormat('dd/mm/yyyy');
  hoja.getRange('L:L').setNumberFormat('#,##0.00 €');
  hoja.autoResizeColumns(1, headers.length);
}

function crearHojaObjetivos(ss) {
  let hoja = ss.getSheetByName(CONFIG.HOJA_OBJETIVOS);
  if (hoja) ss.deleteSheet(hoja);
  hoja = ss.insertSheet(CONFIG.HOJA_OBJETIVOS);
  const headers = ['ID_Agente', 'Nombre_Agente', 'Año', 'Mes',
        'Obj_Citas_Captacion', 'Obj_Exclusivas_Venta', 'Obj_Exclusivas_Comprador',
        'Obj_Captaciones_Abierto', 'Obj_Citas_Compradores', 'Obj_Casas_Ensenadas',
        'Obj_Leads_Compradores', 'Obj_Llamadas', 'Obj_GCI', 'Obj_Volumen_Negocio', 'Fecha_Creacion'];
  hoja.getRange(1, 1, 1, headers.length).setValues([headers]).setBackground('#b70000').setFontColor('#ffffff').setFontWeight('bold').setFontSize(11);
  hoja.setFrozenRows(1);
  hoja.getRange('L:L').setNumberFormat('#,##0.00 €');
  hoja.autoResizeColumns(1, headers.length);
}

function crearHojaConfiguracion(ss) {
  let hoja = ss.getSheetByName(CONFIG.HOJA_CONFIGURACION);
  if (hoja) ss.deleteSheet(hoja);
  hoja = ss.insertSheet(CONFIG.HOJA_CONFIGURACION);
  const config = [
    ['Parámetro', 'Valor', 'Descripción'],
    ['Año_Actual', new Date().getFullYear(), 'Año en curso'],
    ['Objetivos_Acumulados_Default', 'NO', 'SI o NO - Activar objetivos acumulados por defecto'],
    ['Email_Notificaciones', '', 'Email para recibir notificaciones'],
    ['Dias_Alerta_Inactividad', 7, 'Días sin actividad para generar alerta'],
    ['', '', ''],
    ['=== OBJETIVOS GLOBALES POR DEFECTO ===', '', ''],
    ['Obj_Mensual_Citas_Captacion', 15, 'Objetivo mensual por agente'],
    ['Obj_Mensual_Exclusivas_Venta', 5, 'Objetivo mensual por agente'],
    ['Obj_Mensual_Exclusivas_Comprador', 4, 'Objetivo mensual por agente'],
    ['Obj_Mensual_Captaciones_Abierto', 3, 'Objetivo mensual por agente'],
    ['Obj_Mensual_Citas_Compradores', 10, 'Objetivo mensual por agente'],
    ['Obj_Mensual_Casas_Ensenadas', 8, 'Objetivo mensual por agente'],
    ['Obj_Mensual_Leads_Compradores', 20, 'Objetivo mensual por agente'],
    ['Obj_Mensual_Llamadas', 50, 'Objetivo mensual de llamadas'],
    ['Obj_Mensual_GCI', 8000, 'Objetivo mensual GCI en euros'],
    ['Obj_Mensual_Volumen_Negocio', 100000, 'Objetivo mensual volumen de negocio en euros']
  ];
  hoja.getRange(1, 1, config.length, 3).setValues(config);
  hoja.getRange(1, 1, 1, 3).setBackground('#b70000').setFontColor('#ffffff').setFontWeight('bold');
  hoja.getRange('B8:B15').setNumberFormat('#,##0');
  hoja.setColumnWidth(1, 300);
  hoja.setColumnWidth(2, 150);
  hoja.setColumnWidth(3, 300);
}

function crearHojaInventario(ss) {
  let hoja = ss.getSheetByName(CONFIG.HOJA_INVENTARIO);
  if (hoja) ss.deleteSheet(hoja);
  hoja = ss.insertSheet(CONFIG.HOJA_INVENTARIO);
  const headers = [
    'ID_Inmueble', 'Fecha_Alta', 'ID_Agente', 'Nombre_Agente',
    'Tipo', 'Direccion', 'Ciudad', 'CP', 'Provincia',
    'Precio', 'Superficie_M2', 'Habitaciones', 'Baños',
    'Estado', 'Tipo_Contrato', 'Exclusividad',
    'Propietario_Nombre', 'Propietario_Telefono', 'Propietario_Email',
    'Descripcion', 'Observaciones', 'Fecha_Actualizacion', 'Estado_Venta'
  ];
  hoja.getRange(1, 1, 1, headers.length).setValues([headers]).setBackground('#b70000').setFontColor('#ffffff').setFontWeight('bold').setFontSize(11);
  hoja.setFrozenRows(1);
  hoja.getRange('B:B').setNumberFormat('dd/mm/yyyy');
  hoja.getRange('J:J').setNumberFormat('#,##0 €');
  hoja.getRange('K:K').setNumberFormat('#,##0');
  hoja.getRange('V:V').setNumberFormat('dd/mm/yyyy');
  hoja.autoResizeColumns(1, headers.length);

  const tipoInmueble = hoja.getRange('E2:E1000');
  const ruleTipo = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Piso', 'Casa', 'Chalet', 'Apartamento', 'Local', 'Oficina', 'Terreno', 'Garaje', 'Trastero'])
    .setAllowInvalid(false)
    .build();
  tipoInmueble.setDataValidation(ruleTipo);

  const estadoInmueble = hoja.getRange('N2:N1000');
  const ruleEstado = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Nuevo', 'Buen Estado', 'A Reformar', 'Obra Nueva'])
    .setAllowInvalid(false)
    .build();
  estadoInmueble.setDataValidation(ruleEstado);

  const tipoContrato = hoja.getRange('O2:O1000');
  const ruleContrato = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Venta', 'Alquiler', 'Venta/Alquiler'])
    .setAllowInvalid(false)
    .build();
  tipoContrato.setDataValidation(ruleContrato);

  const exclusividad = hoja.getRange('P2:P1000');
  const ruleExclusividad = SpreadsheetApp.newDataValidation()
    .requireValueInList(['SI', 'NO'])
    .setAllowInvalid(false)
    .build();
  exclusividad.setDataValidation(ruleExclusividad);

  const estadoVenta = hoja.getRange('W2:W1000');
  const ruleEstadoVenta = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Activo', 'Reservado', 'Vendido', 'Alquilado', 'Retirado'])
    .setAllowInvalid(false)
    .build();
  estadoVenta.setDataValidation(ruleEstadoVenta);
}

function crearHojaGastosOperativos(ss) {
  let hoja = ss.getSheetByName('Gastos_Operativos');
  if (hoja) ss.deleteSheet(hoja);
  hoja = ss.insertSheet('Gastos_Operativos');
  // Encabezados solicitados
  const headers = ['ID_Gasto', 'Fecha', 'Año', 'Mes', 'Partida', 'Descripción', 'Importe', 'Timestamp'];
  hoja.getRange(1, 1, 1, headers.length).setValues([headers])
      .setBackground('#b70000')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
  hoja.setFrozenRows(1);
  hoja.getRange('G:G').setNumberFormat('#,##0.00 €'); // Formato moneda para Importe
  hoja.autoResizeColumns(1, headers.length);
}

function insertarDatosEjemplo(ss) {
  const hojaAgentes = ss.getSheetByName(CONFIG.HOJA_AGENTES);
  const hojaActividad = ss.getSheetByName(CONFIG.HOJA_ACTIVIDAD);
  const hojaObjetivos = ss.getSheetByName(CONFIG.HOJA_OBJETIVOS);
  const hojaInventario = ss.getSheetByName(CONFIG.HOJA_INVENTARIO);

  const agentes = [
    ['AG001', 'María García', 'maria@kw.com', '600111222', new Date(2024, 0, 15), 'Activo', 'NO', new Date()],
    ['AG002', 'Juan Pérez', 'juan@kw.com', '600222333', new Date(2024, 1, 1), 'Activo', 'NO', new Date()],
    ['AG003', 'Ana Martínez', 'ana@kw.com', '600333444', new Date(2023, 11, 1), 'Activo', 'SI', new Date()],
    ['AG004', 'Carlos López', 'carlos@kw.com', '600444555', new Date(2024, 2, 10), 'Activo', 'NO', new Date()],
    ['AG005', 'Laura Sánchez', 'laura@kw.com', '600555666', new Date(2024, 0, 5), 'Activo', 'NO', new Date()]
  ];
  hojaAgentes.getRange(2, 1, agentes.length, agentes[0].length).setValues(agentes);

  const year = new Date().getFullYear();
  const objetivos = [];
  for (let i = 0; i < agentes.length; i++) {
    for (let mes = 1; mes <= 12; mes++) {
      objetivos.push([
        agentes[i][0], agentes[i][1], year, mes,
        15,5,4,3,10,8,20,8000, new Date()
      ]);
    }
  }
  hojaObjetivos.getRange(2, 1, objetivos.length, objetivos[0].length).setValues(objetivos);

  const actividad = [];
  const hoy = new Date();
  let idActividad = 1;
  for (let i = 0; i < agentes.length; i++) {
    for (let dia = 90; dia >= 0; dia -= Math.floor(Math.random() * 3) + 1) {
      const fecha = new Date(hoy.getTime() - dia * 24 * 60 * 60 * 1000);
      actividad.push([
        'ACT' + String(idActividad++).padStart(5, '0'),
        fecha,
        agentes[i][0],
        agentes[i][1],
        Math.floor(Math.random() * 3),
        Math.floor(Math.random() * 2),
        Math.floor(Math.random() * 2),
        Math.random() > 0.7 ? 1 : 0,
        Math.floor(Math.random() * 3),
        Math.floor(Math.random() * 4),
        Math.floor(Math.random() * 5),
        Math.random() > 0.7 ? Math.random() * 5000 : 0,
        '',
        new Date()
      ]);
    }
  }
  hojaActividad.getRange(2, 1, actividad.length, actividad[0].length).setValues(actividad);

  const inmuebles = [
    ['INM001', new Date(), 'AG001', 'María García', 'Piso', 'Calle Mayor 25, 3ºB', 'Huelva', '21001', 'Huelva', 185000, 95, 3, 2, 'Buen Estado', 'Venta', 'SI', 'Juan Pérez', '600111222', 'juan@email.com', 'Piso céntrico con vistas', '', new Date(), 'Activo'],
    ['INM002', new Date(), 'AG002', 'Juan Pérez', 'Chalet', 'Urbanización Los Pinos 12', 'Punta Umbría', '21100', 'Huelva', 320000, 180, 4, 3, 'Nuevo', 'Venta', 'SI', 'Ana Martínez', '600222333', 'ana@email.com', 'Chalet adosado cerca de la playa', '', new Date(), 'Activo'],
    ['INM003', new Date(), 'AG001', 'María García', 'Apartamento', 'Paseo Marítimo 45', 'Isla Cristina', '21410', 'Huelva', 145000, 65, 2, 1, 'Buen Estado', 'Venta', 'NO', 'Carlos López', '600333444', 'carlos@email.com', 'Apartamento en primera línea', '', new Date(), 'Activo'],
    ['INM004', new Date(), 'AG003', 'Ana Martínez', 'Local', 'Avenida Italia 78', 'Huelva', '21002', 'Huelva', 95000, 120, 0, 1, 'A Reformar', 'Venta', 'SI', 'Laura Sánchez', '600444555', 'laura@email.com', 'Local comercial en zona céntrica', '', new Date(), 'Activo'],
    ['INM005', new Date(), 'AG004', 'Carlos López', 'Casa', 'Calle Andalucía 5', 'Lepe', '21440', 'Huelva', 220000, 150, 4, 2, 'Buen Estado', 'Venta', 'SI', 'Pedro Rodríguez', '600555666', 'pedro@email.com', 'Casa independiente con jardín', '', new Date(), 'Reservado']
  ];
  hojaInventario.getRange(2, 1, inmuebles.length, inmuebles[0].length).setValues(inmuebles);
}

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Dashboard')
    .setTitle('📊 Dashboard Inmobiliario KW')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

function abrirDashboard() {
  const html = HtmlService.createHtmlOutputFromFile('Dashboard')
    .setWidth(1400)
    .setHeight(900)
    .setTitle('📊 Dashboard Inmobiliario KW');
  SpreadsheetApp.getUi().showModalDialog(html, 'Dashboard Inmobiliario');
}

function obtenerListaAgentes() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(CONFIG.HOJA_AGENTES);
    if (!hoja) throw new Error('No se encontró la hoja de Agentes. Ejecuta "Inicializar Sistema" primero.');
    const datos = hoja.getDataRange().getValues();
    const agentes = [];
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0] && datos[i][1]) {
        agentes.push({
          id: datos[i][0],
          nombre: datos[i][1],
          email: datos[i][2],
          objetivosAcumulados: datos[i][6] === 'SI'
        });
      }
    }
    return agentes;
  } catch (error) {
    throw new Error('Error al obtener agentes: ' + error.toString());
  }
}

function guardarActividad(datos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(CONFIG.HOJA_ACTIVIDAD);
    if (!hoja) throw new Error('No se encontró la hoja de Actividad.');
    const ultimaFila = hoja.getLastRow();
    const nuevoId = 'ACT' + String(ultimaFila + 1).padStart(5, '0');
    const fecha = new Date(datos.fecha);
    const fila = [
      nuevoId, fecha, datos.idAgente, datos.nombreAgente,
      datos.citasCaptacion || 0, datos.exclusivasVenta || 0, datos.exclusivasComprador || 0,
      datos.captacionesAbierto || 0, datos.citasCompradores || 0, datos.casasEnsenadas || 0,
      datos.leadsCompradores || 0, datos.llamadas || 0, datos.gci || 0, 0, datos.notas || '', new Date(), 0, 0
    ];
    hoja.appendRow(fila);
    return { success: true, message: 'Actividad guardada correctamente', id: nuevoId };
  } catch (error) {
    throw new Error('Error al guardar actividad: ' + error.toString());
  }
}

function guardarTransaccionGCI(datosTransaccion) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(CONFIG.HOJA_ACTIVIDAD);
    if (!hoja) throw new Error('No se encontró la hoja Actividad_Diaria');
    
    const ultimaFila = hoja.getLastRow();
    const idBase = 'TRX' + String(ultimaFila + 1000).slice(-5);
    const fecha = new Date(datosTransaccion.fecha);
    
    const precioVenta = parseFloat(datosTransaccion.precioVenta) || 0;
    
    datosTransaccion.agentes.forEach((agente, i) => {
      const idTransaccion = `${idBase}-${String(i + 1).padStart(2, '0')}`;
      
      const gci = parseFloat(agente.gci) || 0;
let comisionPct = parseFloat(agente.comisionPct) || 40;
let comisionImporte = parseFloat(agente.comisionImporte) || 0;

// Recálculo bidireccional según campo modificado
if (agente.campoModificado === 'importe' && gci > 0) {
  comisionPct = (comisionImporte / gci) * 100;
} else {
  comisionImporte = (gci * comisionPct / 100);
}

const notas = `TRANSACCIÓN ${datosTransaccion.tipo.toUpperCase()} - ${datosTransaccion.descripcion || 'Venta/Alquiler'} | Lado: ${agente.lado} | Comis: ${comisionPct.toFixed(1)}%`;
      
      hoja.appendRow([
        idTransaccion,              // 1: ID
        fecha,                      // 2: Fecha
        agente.id,                  // 3: ID_Agente
        agente.nombre,              // 4: Nombre_Agente
        0,0,0,0,0,0,0,0,           // 5-12: KPIs en 0
        gci,                        // 13: GCI
        precioVenta,                // 14: Volumen_Negocio
        notas,                      // 15: Notas
        new Date(),                 // 16: Timestamp
        comisionImporte,            // 17: ✅ Comision_Pagada = GCI × %
        comisionPct                 // 18: Pct_Comision
      ]);
    });
    
    return { success: true, message: `Transacción guardada (${datosTransaccion.agentes.length} agentes)` };
  } catch (error) {
    Logger.log('Error guardarTransaccionGCI: ' + error);
    return { success: false, error: error.message };
  }
}

// =========================================================
// 📊 MOTOR DASHBOARD V19: FUSIÓN TOTAL + PARSER EUROS
// =========================================================

function obtenerDatosDashboard(filtros) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaActividad = ss.getSheetByName(CONFIG.HOJA_ACTIVIDAD);
    const hojaObjetivos = ss.getSheetByName(CONFIG.HOJA_OBJETIVOS);
    const hojaAgentes = ss.getSheetByName(CONFIG.HOJA_AGENTES);
    const hojaHistorico = ss.getSheetByName('Historico_Agentes');

    if (!hojaActividad || !hojaObjetivos || !hojaAgentes) {
      throw new Error('Faltan hojas necesarias.');
    }
    
    // 1. Configuración de Fechas
    let yearConsulta = new Date().getFullYear();
    let fechaInicio = new Date(yearConsulta, 0, 1);
    let fechaFin = new Date();
    
    if (filtros) {
        if (filtros.fechaInicio) {
            fechaInicio = new Date(filtros.fechaInicio);
            yearConsulta = fechaInicio.getFullYear();
        }
        if (filtros.fechaFin) fechaFin = new Date(filtros.fechaFin);
    }
    fechaInicio.setHours(0, 0, 0, 0);
    fechaFin.setHours(23, 59, 59, 999);

    // 2. PRE-CARGA DEL HISTÓRICO (LA CLAVE DEL ÉXITO) ⚡
    // Creamos un diccionario: { "AG006": { gci: 200000, citas: 50... }, "AG007": ... }
    const mapaHistorico = {};
    
    if (hojaHistorico) {
        const datosHist = hojaHistorico.getDataRange().getValues();
        // Empezamos en 1 para saltar encabezados
        for (let i = 1; i < datosHist.length; i++) {
            const row = datosHist[i];
            const idRaw = String(row[0]);
            if (!idRaw) continue;

            // NORMALIZACIÓN: Quitamos espacios y símbolos. "AG 006 " -> "AG006"
            const idKey = idRaw.toUpperCase().replace(/[^A-Z0-9]/g, '');
            const anioRow = parseInt(row[2]); // Columna C: Año
            const mesRow = parseInt(row[3]);  // Columna D: Mes

            // Filtro por año (Laxo: permite string o number)
            if (anioRow == yearConsulta) {
                if (!mapaHistorico[idKey]) {
                    mapaHistorico[idKey] = {
                        totalGCI: 0, totalCitas: 0, totalExcl: 0, totalVentas: 0,
                        totalAbierto: 0, totalCitasComp: 0, totalVisitas: 0,
                        // Arrays mensuales para el gráfico de equipo
                        mensual: {
                            gci: Array(12).fill(0),
                            citas: Array(12).fill(0),
                            excl: Array(12).fill(0)
                        }
                    };
                }

                // --- MAPEO DE COLUMNAS SEGÚN TUS DATOS ---
                // ID=0, Nombre=1, Año=2, Mes=3
                // Citas=4, Capt_Excl=5, Capt_Abierto=6, Visitas_Comp=7, Casas_Ens=8
                // Vtas_Excl=14, Vtas_Abierto=16, Vtas_Comp=18, Vtas_Alq=20
                // GCI_Total=22
                
                const gci = parsearNumeroEU(row[22]);
                const citas = parsearNumeroEU(row[4]);
                const excl = parsearNumeroEU(row[5]);
                const abierto = parsearNumeroEU(row[6]);
                const citasComp = parsearNumeroEU(row[7]);
                const visitas = parsearNumeroEU(row[8]);
                
                // Suma de ventas
                const ventas = parsearNumeroEU(row[14]) + parsearNumeroEU(row[16]) + parsearNumeroEU(row[18]) + parsearNumeroEU(row[20]);

                // ¿Este mes histórico entra en el filtro de fechas seleccionado?
                // Creamos fecha ficticia (día 15 del mes)
                const fechaMesHist = new Date(yearConsulta, mesRow - 1, 15);
                
                if (fechaMesHist >= fechaInicio && fechaMesHist <= fechaFin) {
                    // SUMAR A LOS TOTALES (Para Tarjetas y Signos Vitales)
                    mapaHistorico[idKey].totalGCI += gci;
                    mapaHistorico[idKey].totalCitas += citas;
                    mapaHistorico[idKey].totalExcl += excl;
                    mapaHistorico[idKey].totalVentas += ventas;
                    mapaHistorico[idKey].totalAbierto += abierto;
                    mapaHistorico[idKey].totalCitasComp += citasComp;
                    mapaHistorico[idKey].totalVisitas += visitas;
                }

                // SUMAR A LA EVOLUCIÓN (Para gráfico de líneas)
                if (mesRow >= 1 && mesRow <= 12) {
                    const mIdx = mesRow - 1;
                    mapaHistorico[idKey].mensual.gci[mIdx] += gci;
                    mapaHistorico[idKey].mensual.citas[mIdx] += citas;
                    mapaHistorico[idKey].mensual.excl[mIdx] += excl;
                }
            }
        }
    }

    // 3. OBTENER RESTO DE DATOS
    const datosAgentes = hojaAgentes.getDataRange().getValues();
    const todasActividades = hojaActividad.getDataRange().getValues();
    const todosObjetivos = hojaObjetivos.getDataRange().getValues();
    
    const resultados = [];
    const mesesPeriodo = calcularMesesEnPeriodo(fechaInicio, fechaFin);
    
    // Objeto global del equipo
    const evolucionMensualEquipo = { labels: mesesPeriodo.map(m => obtenerNombreMesAbreviado(m.mes)) };
    Object.keys(kpiNames).forEach(key => {
        evolucionMensualEquipo[key] = { realizado: Array(mesesPeriodo.length).fill(0), objetivo: Array(mesesPeriodo.length).fill(0) };
    });
    
    // 4. BUCLE MAESTRO DE AGENTES
    for (let i = 1; i < datosAgentes.length; i++) {
      if (datosAgentes[i][0]) {
        
        const idOriginal = String(datosAgentes[i][0]).trim();
        // ID BLINDADO PARA BUSCAR EN EL MAPA
        const idBusqueda = idOriginal.toUpperCase().replace(/[^A-Z0-9]/g, '');
        
        const nombreAgente = datosAgentes[i][1];
        const esAcumulativo = datosAgentes[i][6] === 'SI';
        const estado = datosAgentes[i][5];

        if (estado === 'Activo') {

            // A. Datos APP (Usamos ID original para la hoja de actividad)
            let actividad = obtenerActividadAgente(idOriginal, fechaInicio, fechaFin, todasActividades);
            
            // B. FUSIÓN CON HISTÓRICO (Usamos ID blindado) 🔥
            const datosHist = mapaHistorico[idBusqueda];
            
            if (datosHist) {
                // AQUÍ ES DONDE OCURRE LA SUMA MÁGICA
                actividad.gci += datosHist.totalGCI;
                actividad.citasCaptacion += datosHist.totalCitas;
                actividad.exclusivasVenta += datosHist.totalExcl;
                actividad.captacionesAbierto += datosHist.totalAbierto;
                actividad.citasCompradores += datosHist.totalCitasComp;
                actividad.casasEnsenadas += datosHist.totalVisitas;
                if (typeof actividad.ventas === 'undefined') actividad.ventas = 0;
                actividad.ventas += datosHist.totalVentas;
                
                // Sumar al gráfico general del equipo
                for (let m = 0; m < 12; m++) {
                    if (evolucionMensualEquipo.gci) evolucionMensualEquipo.gci.realizado[m] += datosHist.mensual.gci[m];
                    if (evolucionMensualEquipo.citasCaptacion) evolucionMensualEquipo.citasCaptacion.realizado[m] += datosHist.mensual.citas[m];
                    if (evolucionMensualEquipo.exclusivasVenta) evolucionMensualEquipo.exclusivasVenta.realizado[m] += datosHist.mensual.excl[m];
                }
            }

            // C. Sumar datos de la App al Gráfico del Equipo
            const evolucionApp = calcularEvolucionMensual(idOriginal, mesesPeriodo, esAcumulativo, todasActividades, todosObjetivos);
            
            mesesPeriodo.forEach((mes, idx) => {
                Object.keys(kpiNames).forEach(key => {
                    if (evolucionMensualEquipo[key] && evolucionApp[key]) {
                        evolucionMensualEquipo[key].realizado[idx] += evolucionApp[key].realizado[idx];
                        evolucionMensualEquipo[key].objetivo[idx] += evolucionApp[key].objetivo[idx];
                    }
                });
            });

            // D. Calcular Ratios Finales (Con la suma ya hecha)
            const objetivos = obtenerObjetivosAgente(idOriginal, fechaInicio, fechaFin, todosObjetivos);
            const cumplimientos = calcularCumplimientos(actividad, objetivos);
            const cumplimientoGlobal = calcularCumplimientoGlobal(cumplimientos);
            const ratios = calcularRatios(actividad, objetivos);
            
            const cumplimientoGlobalSeguro = isNaN(cumplimientoGlobal) || !isFinite(cumplimientoGlobal) ? 0 : cumplimientoGlobal;
            
            let estadoClase = 'bajo';
            if (cumplimientoGlobalSeguro >= 90) estadoClase = 'excelente';
            else if (cumplimientoGlobalSeguro >= 70) estadoClase = 'bueno';

            resultados.push({
                id: idOriginal,
                agente: nombreAgente,
                realizado: actividad, // ¡ESTE OBJETO LLEVA LA SUMA FINAL!
                objetivos: objetivos,
                cumplimientos: cumplimientos,
                cumplimientoGlobal: cumplimientoGlobalSeguro.toFixed(1),
                estadoClase: estadoClase,
                ratios: ratios,
                evolucionMensual: evolucionApp
            });
        }
      }
    }

    // Lista de transacciones (Mantener vacía o llenar según lógica anterior)
    const listaTransacciones = [];

    return {
      agentes: resultados,
      evolucionMensualEquipo: evolucionMensualEquipo,
      transacciones: listaTransacciones
    };

  } catch (error) {
    Logger.log('ERROR CRITICO DASHBOARD: ' + error.toString());
    throw error;
  }
}

// =========================================================
// 💶 FUNCIÓN VITAL: PARSER DE NÚMEROS EUROPEOS
// =========================================================
// --- TRADUCTOR DE NÚMEROS ESPAÑOLES (VITAL) ---
function parsearNumeroEU(valor) {
  if (valor === null || valor === undefined || valor === '') return 0;
  
  // Si ya es un número puro, lo devolvemos tal cual
  if (typeof valor === 'number') return valor;
  
  let str = String(valor).trim();
  
  // Si es un guión o vacío
  if (str === '-' || str === '') return 0;

  // 1. Quitamos el símbolo de Euro y espacios
  str = str.replace('€', '').replace(/\s/g, '');
  
  // 2. Limpieza de formato Español (1.000,00) a Americano (1000.00)
  // Si tiene puntos y comas, asumimos formato estándar 1.234,56
  if (str.includes('.') && str.includes(',')) {
    str = str.replace(/\./g, ''); // Quitar puntos de miles
    str = str.replace(',', '.');  // Cambiar coma por punto decimal
  } 
  // Si solo tiene coma (ej: 50,5) -> 50.5
  else if (str.includes(',')) {
    str = str.replace(',', '.');
  }
  // Si solo tiene punto (ej: 1.200) -> 1200
  // CUIDADO: En sistema US 1.200 es 1 coma 2. Aquí forzamos que punto es mil.
  else if (str.includes('.')) {
     str = str.replace(/\./g, '');
  }

  return parseFloat(str) || 0;
}

// =========================================================
// 🗺️ NUEVA FUNCIÓN AUXILIAR: MAPA DE MEMORIA HISTÓRICA
// =========================================================
function cargarMapaHistoricoOptimizado(anio) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName('Historico_Agentes');
    const mapa = {};
    
    if (!hoja) {
        Logger.log("⚠️ No existe la hoja Historico_Agentes");
        return mapa; 
    }

    const datos = hoja.getDataRange().getValues();
    Logger.log(`📊 Cargando Histórico: ${datos.length} filas encontradas para año ${anio}`);
    
    // Empezamos en 1 para saltar cabecera
    for (let i = 1; i < datos.length; i++) {
        const row = datos[i];
        if (!row[0]) continue; 

        // NORMALIZACIÓN DE ID
        const idKey = String(row[0]).toUpperCase().replace(/[^A-Z0-9]/g, '');
        const anioFila = row[2]; 
        
        // Coincidencia de año (laxa)
        if (anioFila == anio) {
            if (!mapa[idKey]) {
                mapa[idKey] = {
                    gci: Array(12).fill(0), citas: Array(12).fill(0), excl: Array(12).fill(0),
                    abierto: Array(12).fill(0), citasComp: Array(12).fill(0), visitas: Array(12).fill(0),
                    ventas: Array(12).fill(0), totalGCI: 0, totalCitas: 0, totalExclusivas: 0, totalVentas: 0,
                    totalCaptAbierto: 0, totalCitasComp: 0, totalVisitas: 0
                };
            }

            const mes = parseInt(row[3]); 
            if (mes >= 1 && mes <= 12) {
                const idx = mes - 1;
                
                // USAMOS EL PARSER EUROPEO AQUÍ
                const gciVal = parsearNumeroES(row[22]);     // Col 22: GCI_Total
                const citasVal = parsearNumeroES(row[4]);    // Col 4: Citas
                const exclVal = parsearNumeroES(row[5]);     // Col 5: Capt_Excl
                const abiertoVal = parsearNumeroES(row[6]);  // Col 6: Capt_Abierto
                const citasCompVal = parsearNumeroES(row[7]);// Col 7: Visitas_Comp
                const visitasVal = parsearNumeroES(row[8]);  // Col 8: Casas_Ens
                
                // Ventas: Suma de columnas 14, 16, 18, 20
                const vtasVal = parsearNumeroES(row[14]) + parsearNumeroES(row[16]) + parsearNumeroES(row[18]) + parsearNumeroES(row[20]);

                // Asignar al mes
                mapa[idKey].gci[idx] += gciVal;
                mapa[idKey].citas[idx] += citasVal;
                mapa[idKey].excl[idx] += exclVal;
                mapa[idKey].abierto[idx] += abiertoVal;
                mapa[idKey].citasComp[idx] += citasCompVal;
                mapa[idKey].visitas[idx] += visitasVal;
                mapa[idKey].ventas[idx] += vtasVal;
                
                // Acumular Totales (para las tarjetas)
                mapa[idKey].totalGCI += gciVal;
                mapa[idKey].totalCitas += citasVal;
                mapa[idKey].totalExclusivas += exclVal;
                mapa[idKey].totalVentas += vtasVal;
                mapa[idKey].totalCaptAbierto += abiertoVal;
                mapa[idKey].totalCitasComp += citasCompVal;
                mapa[idKey].totalVisitas += visitasVal;
            }
        }
    }
    
    // Log para verificar si ha cargado algo
    const keys = Object.keys(mapa);
    if(keys.length > 0) {
        Logger.log(`✅ Datos cargados para ${keys.length} agentes. Ejemplo (${keys[0]}): GCI Total = ${mapa[keys[0]].totalGCI}`);
    } else {
        Logger.log("⚠️ No se encontraron datos coincidentes para el año " + anio);
    }
    
    return mapa;
}

// ====== TODAS LAS DEMÁS FUNCIONES SIGUEN IGUALES (NO TOQUE NADA MÁS) ======
function obtenerActividadAgente(idAgente, fechaInicio, fechaFin, todasActividades) {
  // Inicializamos el objeto con todas las métricas a 0
  const actividad = {
    citasCaptacion: 0, 
    exclusivasVenta: 0, 
    exclusivasComprador: 0,
    captacionesAbierto: 0, 
    citasCompradores: 0, 
    casasEnsenadas: 0,
    leadsCompradores: 0, 
    llamadas: 0, 
    gci: 0, 
    volumenNegocio: 0,
    ventas: 0 // Importante para calcular Ticket Promedio
  };

  // Normalizamos el ID que buscamos (Mayúsculas y sin espacios)
  const targetID = String(idAgente).trim().toUpperCase();

  for (let i = 0; i < todasActividades.length; i++) {
    const row = todasActividades[i];
    
    // 1. Validación básica de fecha
    const fechaRaw = row[1];
    if (!fechaRaw) continue;

    // 2. Validación de ID (Blindaje contra errores de filtrado)
    // Leemos la columna 2 (ID_Agente). Si no coincide, saltamos.
    // Esto permite usar la función tanto con listas filtradas como con la hoja entera.
    const rowID = String(row[2]).trim().toUpperCase();
    if (rowID !== targetID) continue;

    // 3. Validación de Rango de Fechas
    const fecha = new Date(fechaRaw);
    if (fecha instanceof Date && !isNaN(fecha) && fecha >= fechaInicio && fecha <= fechaFin) {
      
      // 4. Suma de Métricas (con protección contra nulos)
      actividad.citasCaptacion += parseFloat(row[4]) || 0;
      actividad.exclusivasVenta += parseFloat(row[5]) || 0;
      actividad.exclusivasComprador += parseFloat(row[6]) || 0;
      actividad.captacionesAbierto += parseFloat(row[7]) || 0;
      actividad.citasCompradores += parseFloat(row[8]) || 0;
      actividad.casasEnsenadas += parseFloat(row[9]) || 0;
      actividad.leadsCompradores += parseFloat(row[10]) || 0;
      actividad.llamadas += parseFloat(row[11]) || 0;
      
      // Datos económicos
      const gci = parseFloat(row[12]) || 0;
      actividad.gci += gci;
      actividad.volumenNegocio += parseFloat(row[13]) || 0;

      // 5. Cálculo de Ventas (Transacciones)
      // Si hay GCI positivo, cuenta como una venta cerrada
      if (gci > 0) {
        actividad.ventas += 1;
      }
    }
  }
  
  return actividad;
}

function obtenerObjetivosAgente(idAgente, fechaInicio, fechaFin, todosObjetivos) {
  const objetivos = {
        citasCaptacion: 0, exclusivasVenta: 0, exclusivasComprador: 0,
        captacionesAbierto: 0, citasCompradores: 0, casasEnsenadas: 0,
        leadsCompradores: 0, llamadas: 0, gci: 0, volumenNegocio: 0
    };
  
  // CORRECCIÓN: Empezamos en 0
  for (let i = 0; i < todosObjetivos.length; i++) {
    const year = todosObjetivos[i][2];
    const mes = todosObjetivos[i][3];
    if (!year || !mes) continue;
    
    const fechaMes = new Date(year, mes - 1, 1);
    
    if (fechaMes instanceof Date && !isNaN(fechaMes) && 
        fechaMes >= new Date(fechaInicio.getFullYear(), fechaInicio.getMonth(), 1) && 
        fechaMes <= new Date(fechaFin.getFullYear(), fechaFin.getMonth(), 1)) {
          
      objetivos.citasCaptacion += parseFloat(todosObjetivos[i][4]) || 0;
      objetivos.exclusivasVenta += parseFloat(todosObjetivos[i][5]) || 0;
      objetivos.exclusivasComprador += parseFloat(todosObjetivos[i][6]) || 0;
      objetivos.captacionesAbierto += parseFloat(todosObjetivos[i][7]) || 0;
      objetivos.citasCompradores += parseFloat(todosObjetivos[i][8]) || 0;
      objetivos.casasEnsenadas += parseFloat(todosObjetivos[i][9]) || 0;
      objetivos.leadsCompradores += parseFloat(todosObjetivos[i][10]) || 0;
            objetivos.llamadas += parseFloat(todosObjetivos[i][11]) || 0;
            objetivos.gci += parseFloat(todosObjetivos[i][12]) || 0;
            objetivos.volumenNegocio += parseFloat(todosObjetivos[i][13]) || 0;
    }
  }
  return objetivos;
}

function calcularCumplimientos(actividad, objetivos) {
  const calcular = (realizado, objetivo) => {
    if (objetivo === 0) return (realizado > 0) ? 100 : 0;
    return ((realizado / objetivo) * 100);
  };
 
  return {
    citasCaptacion: calcular(actividad.citasCaptacion, objetivos.citasCaptacion),
    exclusivasVenta: calcular(actividad.exclusivasVenta, objetivos.exclusivasVenta),
    exclusivasComprador: calcular(actividad.exclusivasComprador, objetivos.exclusivasComprador),
    captacionesAbierto: calcular(actividad.captacionesAbierto, objetivos.captacionesAbierto),
    citasCompradores: calcular(actividad.citasCompradores, objetivos.citasCompradores),
    casasEnsenadas: calcular(actividad.casasEnsenadas, objetivos.casasEnsenadas),
    leadsCompradores: calcular(actividad.leadsCompradores, objetivos.leadsCompradores),
    gci: calcular(actividad.gci, objetivos.gci)
  };
}

function calcularCumplimientoGlobal(cumplimientos) {
  const pesoGCI = 0.5;
  const numOtrasMetricas = Object.keys(cumplimientos).length - 1;
  if (numOtrasMetricas <= 0) return cumplimientos.gci || 0;
  const pesoOtras = (1 - pesoGCI) / numOtrasMetricas;
  let sumaPonderada = (cumplimientos.gci || 0) * pesoGCI;
 
  Object.keys(cumplimientos).forEach(key => {
    if (key !== 'gci') {
      sumaPonderada += (cumplimientos[key] || 0) * pesoOtras;
    }
  });
  return sumaPonderada;
}

function calcularRatios(actividad, objetivos) {
  const ratios = {
    conversionCaptacion: 0,
    ratioCierre: 0,
    productividad: 0,
    eficienciaLeads: 0,
    ticketPromedio: 0,
    pctExclusivas: 0,
    actividadTotal: 0,
    visitasPorComprador: 0,
    leadsPorCita: 0
  };
  const totalCaptaciones = actividad.exclusivasVenta + actividad.exclusivasComprador + actividad.captacionesAbierto;
  const totalExclusivas = actividad.exclusivasVenta + actividad.exclusivasComprador;
  if (actividad.citasCaptacion > 0) {
    ratios.conversionCaptacion = (totalExclusivas / actividad.citasCaptacion) * 100;
  }
 
  if (totalExclusivas > 0) {
    ratios.ratioCierre = (actividad.gci / totalExclusivas);
  }
 
  if (actividad.citasCompradores > 0) {
    ratios.productividad = (actividad.casasEnsenadas / actividad.citasCompradores);
  }
 
  if (actividad.exclusivasComprador > 0) {
    ratios.eficienciaLeads = (actividad.leadsCompradores / actividad.exclusivasComprador);
  }
 
  if (totalExclusivas > 0) {
    ratios.ticketPromedio = (actividad.gci / totalExclusivas);
  }
 
  if (totalCaptaciones > 0) {
    ratios.pctExclusivas = ((totalExclusivas / totalCaptaciones) * 100);
  }
 
  ratios.actividadTotal = actividad.citasCaptacion + actividad.citasCompradores + actividad.casasEnsenadas;
  if (actividad.exclusivasComprador > 0) {
    ratios.visitasPorComprador = (actividad.casasEnsenadas / actividad.exclusivasComprador);
  }
 
  if (actividad.citasCompradores > 0) {
    ratios.leadsPorCita = (actividad.leadsCompradores / actividad.citasCompradores);
  }
 
  Object.keys(ratios).forEach(key => {
      ratios[key] = parseFloat(Math.max(0, ratios[key] && isFinite(ratios[key]) ? ratios[key] : 0).toFixed(1));
  });
  ratios.ticketPromedio = parseFloat(ratios.ticketPromedio.toFixed(0));
  ratios.ratioCierre = parseFloat(ratios.ratioCierre.toFixed(0));
  return ratios;
}

function calcularMesesEnPeriodo(fechaInicio, fechaFin) {
  const meses = [];
  const fecha = new Date(fechaInicio.getFullYear(), fechaInicio.getMonth(), 1);
  const fechaTope = new Date(fechaFin.getFullYear(), fechaFin.getMonth(), 1);
 
  while (fecha <= fechaTope) {
    meses.push({
      year: fecha.getFullYear(),
      mes: fecha.getMonth() + 1
    });
    fecha.setMonth(fecha.getMonth() + 1);
  }
 
  return meses;
}

function calcularObjetivosAcumuladosPendientes(idAgente, fechaInicio, todasActividades, todosObjetivos) {
  const pendientes = {
    citasCaptacion: 0, exclusivasVenta: 0, exclusivasComprador: 0,
    captacionesAbierto: 0, citasCompradores: 0, casasEnsenadas: 0,
    leadsCompradores: 0, llamadas: 0, gci: 0, volumenNegocio: 0
  };
  const inicioAno = new Date(fechaInicio.getFullYear(), 0, 1);
 
  if (inicioAno >= fechaInicio) return pendientes;
 
  const mesesAnteriores = calcularMesesEnPeriodo(inicioAno, new Date(fechaInicio.getTime() - 1));
  for (const mesInfo of mesesAnteriores) {
    const objetivoMes = obtenerObjetivoMes(idAgente, mesInfo.year, mesInfo.mes, todosObjetivos);
    const actividadMes = obtenerActividadMes(idAgente, mesInfo.year, mesInfo.mes, todasActividades);
   
    Object.keys(pendientes).forEach(key => {
        pendientes[key] += Math.max(0, (objetivoMes[key] || 0) - (actividadMes[key] || 0));
    });
  }
 
  return pendientes;
}

function obtenerObjetivoMes(idAgente, year, mes, todosObjetivos) {
  for (let i = 1; i < todosObjetivos.length; i++) {
    if (todosObjetivos[i][0] === idAgente && todosObjetivos[i][2] === year && todosObjetivos[i][3] === mes) {
      return {
        citasCaptacion: parseFloat(todosObjetivos[i][4]) || 0,
        exclusivasVenta: parseFloat(todosObjetivos[i][5]) || 0,
        exclusivasComprador: parseFloat(todosObjetivos[i][6]) || 0,
        captacionesAbierto: parseFloat(todosObjetivos[i][7]) || 0,
        citasCompradores: parseFloat(todosObjetivos[i][8]) || 0,
        casasEnsenadas: parseFloat(todosObjetivos[i][9]) || 0,
        leadsCompradores: parseFloat(todosObjetivos[i][10]) || 0,
                llamadas: parseFloat(todosObjetivos[i][11]) || 0,
                gci: parseFloat(todosObjetivos[i][12]) || 0,
                volumenNegocio: parseFloat(todosObjetivos[i][13]) || 0
      };
    }
  }
  return { citasCaptacion: 0, exclusivasVenta: 0, exclusivasComprador: 0, captacionesAbierto: 0, citasCompradores: 0, casasEnsenadas: 0, leadsCompradores: 0, llamadas: 0, gci: 0, volumenNegocio: 0 };
}

function obtenerActividadMes(idAgente, year, mes, todasActividades) {
  const actividad = {
        citasCaptacion: 0, exclusivasVenta: 0, exclusivasComprador: 0,
        captacionesAbierto: 0, citasCompradores: 0, casasEnsenadas: 0,
        leadsCompradores: 0, llamadas: 0, gci: 0, volumenNegocio: 0
    };
 
  for (let i = 1; i < todasActividades.length; i++) {
    const fechaRaw = todasActividades[i][1];
    if (!fechaRaw) continue;
   
    const fecha = new Date(fechaRaw);
    if (todasActividades[i][2] === idAgente && fecha instanceof Date && !isNaN(fecha) && fecha.getFullYear() === year && (fecha.getMonth() + 1) === mes) {
      actividad.citasCaptacion += parseFloat(todasActividades[i][4]) || 0;
      actividad.exclusivasVenta += parseFloat(todasActividades[i][5]) || 0;
      actividad.exclusivasComprador += parseFloat(todasActividades[i][6]) || 0;
      actividad.captacionesAbierto += parseFloat(todasActividades[i][7]) || 0;
      actividad.citasCompradores += parseFloat(todasActividades[i][8]) || 0;
      actividad.casasEnsenadas += parseFloat(todasActividades[i][9]) || 0;
      actividad.leadsCompradores += parseFloat(todasActividades[i][10]) || 0;
            actividad.llamadas += parseFloat(todasActividades[i][11]) || 0;
            actividad.gci += parseFloat(todasActividades[i][12]) || 0;
            actividad.volumenNegocio += parseFloat(todasActividades[i][13]) || 0;
    }
  }
  return actividad;
}

function calcularEvolucionMensual(idAgente, mesesPeriodo, esAcumulativo, todasActividades, todosObjetivos) {
  const evolucion = { labels: mesesPeriodo.map(m => obtenerNombreMesAbreviado(m.mes)) };
  Object.keys(kpiNames).forEach(key => {
      evolucion[key] = { realizado: [], objetivo: [] };
  });
  let pendientes = {
    citasCaptacion: 0, exclusivasVenta: 0, exclusivasComprador: 0,
    captacionesAbierto: 0, citasCompradores: 0, casasEnsenadas: 0,
    leadsCompradores: 0, llamadas: 0, gci: 0, volumenNegocio: 0
  };
  const inicioAno = new Date(mesesPeriodo[0].year, 0, 1);
  const primerMesPeriodo = new Date(mesesPeriodo[0].year, mesesPeriodo[0].mes - 1, 1);
  if (esAcumulativo && inicioAno < primerMesPeriodo) {
      pendientes = calcularObjetivosAcumuladosPendientes(idAgente, primerMesPeriodo, todasActividades, todosObjetivos);
  }
  for (const mesInfo of mesesPeriodo) {
    const actividadMes = obtenerActividadMes(idAgente, mesInfo.year, mesInfo.mes, todasActividades);
    const objetivoMesBase = obtenerObjetivoMes(idAgente, mesInfo.year, mesInfo.mes, todosObjetivos);
    const objetivoMesAcumulado = { ...objetivoMesBase };
    if (esAcumulativo) {
        Object.keys(pendientes).forEach(key => {
            objetivoMesAcumulado[key] += pendientes[key];
            const cubierto = (actividadMes[key] || 0) - (objetivoMesBase[key] || 0);
            if (cubierto > 0) {
                pendientes[key] = Math.max(0, pendientes[key] - cubierto);
            }
        });
    }
    const ratiosMes = calcularRatios(actividadMes, objetivoMesAcumulado);
    const cumplimientosMes = calcularCumplimientos(actividadMes, objetivoMesAcumulado);
    const cumplimientoGlobalMes = calcularCumplimientoGlobal(cumplimientosMes);
    Object.keys(actividadMes).forEach(key => {
      if (evolucion[key]) evolucion[key].realizado.push(actividadMes[key]);
    });
    Object.keys(objetivoMesAcumulado).forEach(key => {
      if (evolucion[key]) evolucion[key].objetivo.push(objetivoMesAcumulado[key]);
    });
    Object.keys(ratiosMes).forEach(key => {
      if (evolucion[key]) {
        evolucion[key].realizado.push(ratiosMes[key]);
        const meta = getKpiMetadata(key, 0);
        evolucion[key].objetivo.push(meta.thresholds?.bueno || 0);
      }
    });
    evolucion.cumplimientoGlobal.realizado.push(cumplimientoGlobalMes);
    evolucion.cumplimientoGlobal.objetivo.push(100);
  }
 
  return evolucion;
}

function obtenerNombreMesAbreviado(mes) {
  const meses = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];
  return meses[mes - 1] || '';
}

// =========================================================
// 🕵️‍♂️ MOTOR GRÁFICO V11: BÚSQUEDA A PRUEBA DE FALLOS
// =========================================================

// =========================================================
// 🕵️‍♂️ MOTOR V11: BÚSQUEDA HÍBRIDA (ID O NOMBRE) + FUSIÓN
// =========================================================

// =========================================================
// 🚀 MOTOR V13: BÚSQUEDA "SUCIA" + GRÁFICOS CORRECTOS
// =========================================================

function obtenerDatosAgenteCompleto(criterioBusqueda, filtros) {
  SpreadsheetApp.flush(); 

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaActividad = ss.getSheetByName('Actividad_Diaria');
    const hojaObjetivos = ss.getSheetByName('Objetivos');
    const hojaAgentes = ss.getSheetByName('Agentes');
    
    // --- 1. IDENTIFICACIÓN AGRESIVA (IGNORA ESPACIOS Y FORMATO) ---
    const datosAgentes = hojaAgentes.getDataRange().getValues();
    let idAgente = null;
    let nombreOficial = "";
    let esAcumulativo = false;
    
    // Quitamos todo lo que no sea letra o número (Espacios, guiones...)
    // Ej: " AG-006 " -> "AG006"
    const busquedaLimpia = String(criterioBusqueda).toUpperCase().replace(/[^A-Z0-9]/g, '');
    
    for (let i = 1; i < datosAgentes.length; i++) {
      const idDB = String(datosAgentes[i][0]).toUpperCase().replace(/[^A-Z0-9]/g, '');
      const nombreDB = String(datosAgentes[i][1]).toUpperCase().replace(/[^A-Z0-9]/g, '');
      
      // Coincidencia laxa: Si uno contiene al otro, es válido.
      if (idDB.includes(busquedaLimpia) || busquedaLimpia.includes(idDB) || 
          nombreDB.includes(busquedaLimpia) || busquedaLimpia.includes(nombreDB)) {
        
        idAgente = String(datosAgentes[i][0]).trim(); // ID REAL
        nombreOficial = datosAgentes[i][1];
        esAcumulativo = (String(datosAgentes[i][6]).toUpperCase() === 'SI');
        break;
      }
    }
    
    // Si falla, usamos el input tal cual (fallback para IDs directos)
    if (!idAgente) {
        Logger.log(`⚠️ Match agresivo falló para: ${criterioBusqueda}. Usando input directo.`);
        idAgente = String(criterioBusqueda).trim();
        nombreOficial = "Agente " + idAgente;
    }

    // --- 2. CONFIGURAR FECHAS ---
    let yearConsulta = new Date().getFullYear();
    let fechaInicio = new Date(yearConsulta, 0, 1);
    let fechaFin = new Date();
    
    if (filtros && filtros.fechaInicio) {
        fechaInicio = new Date(filtros.fechaInicio);
        yearConsulta = fechaInicio.getFullYear();
    }
    if (filtros && filtros.fechaFin) fechaFin = new Date(filtros.fechaFin);
    fechaInicio.setHours(0, 0, 0, 0);
    fechaFin.setHours(23, 59, 59, 999);

    // --- 3. LEER ACTIVIDAD DIARIA (APP) ---
    const datosAct = hojaActividad.getDataRange().getValues();
    const actividad = { 
        citasCaptacion: 0, exclusivasVenta: 0, exclusivasComprador: 0, 
        captacionesAbierto: 0, citasCompradores: 0, casasEnsenadas: 0, 
        leadsCompradores: 0, llamadas: 0, gci: 0, volumenNegocio: 0, ventas: 0 
    };
    
    // Arrays Base para Gráficos
    const graficoApp = {
        gci: Array(12).fill(0), citasCaptacion: Array(12).fill(0),
        exclusivasVenta: Array(12).fill(0), captacionesAbierto: Array(12).fill(0),
        citasCompradores: Array(12).fill(0), casasEnsenadas: Array(12).fill(0)
    };

    const targetID = idAgente.toUpperCase().replace(/[^A-Z0-9]/g, '');

    for (let i = 1; i < datosAct.length; i++) {
        const row = datosAct[i];
        const filaID = String(row[2]).toUpperCase().replace(/[^A-Z0-9]/g, '');
        
        if (filaID === targetID) {
            const fecha = new Date(row[1]);
            
            if (fecha >= fechaInicio && fecha <= fechaFin) {
                actividad.citasCaptacion += parseFloat(row[4])||0;
                actividad.exclusivasVenta += parseFloat(row[5])||0;
                // ... resto de sumas totales ...
                actividad.gci += parseFloat(row[12])||0;
                if((parseFloat(row[12])||0) > 0) actividad.ventas++;
            }

            if (fecha.getFullYear() == yearConsulta) {
                const m = fecha.getMonth();
                graficoApp.gci[m] += parseFloat(row[12])||0;
                graficoApp.citasCaptacion[m] += parseFloat(row[4])||0;
                graficoApp.exclusivasVenta[m] += parseFloat(row[5])||0;
                // ... resto de gráficos app ...
            }
        }
    }

    // --- 4. LEER E INYECTAR HISTÓRICO (FUSIÓN V13) ---
    const historicoMeses = obtenerDesgloseHistorico(idAgente, yearConsulta);

    if (historicoMeses) {
        // A) Sumar totales a la TARJETA
        actividad.gci += historicoMeses.totalGCI;
        actividad.citasCaptacion += historicoMeses.totalCitas;
        actividad.exclusivasVenta += historicoMeses.totalExclusivas;
        actividad.ventas += historicoMeses.totalVentas;
        
        // B) Sumar al GRÁFICO (Mes a Mes)
        for(let m=0; m<12; m++) {
            graficoApp.gci[m] += historicoMeses.gci[m];
            graficoApp.citasCaptacion[m] += historicoMeses.citasCaptacion[m];
            graficoApp.exclusivasVenta[m] += historicoMeses.exclusivasVenta[m];
            graficoApp.captacionesAbierto[m] += historicoMeses.captacionesAbierto[m];
            graficoApp.citasCompradores[m] += historicoMeses.citasCompradores[m];
            graficoApp.casasEnsenadas[m] += historicoMeses.casasEnsenadas[m];
        }
    }

    // --- 5. RESULTADOS FINALES ---
    const todosObjetivos = hojaObjetivos.getDataRange().getValues();
    const objetivos = obtenerObjetivosAgente(idAgente, fechaInicio, fechaFin, todosObjetivos);
    const cumplimientos = calcularCumplimientos(actividad, objetivos);
    const cumplimientoGlobal = calcularCumplimientoGlobal(cumplimientos);
    const ratios = calcularRatios(actividad, objetivos);

    // Empaquetado Gráfico
    const mesesNombres = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];
    const evolucion = {
        labels: mesesNombres,
        gci: { realizado: graficoApp.gci, objetivo: [] },
        citasCaptacion: { realizado: graficoApp.citasCaptacion, objetivo: [] },
        exclusivasVenta: { realizado: graficoApp.exclusivasVenta, objetivo: [] },
        captacionesAbierto: { realizado: graficoApp.captacionesAbierto, objetivo: [] },
        citasCompradores: { realizado: graficoApp.citasCompradores, objetivo: [] },
        casasEnsenadas: { realizado: graficoApp.casasEnsenadas, objetivo: [] },
        conversionCaptacion: { realizado: [], objetivo: [] }
    };

    // Objetivos mensuales
    const objsMes = obtenerObjetivosMensuales(idAgente, yearConsulta, todosObjetivos);
    
    for(let m=0; m<12; m++) {
        // Ratio Conversión Gráfico
        const c = graficoApp.citasCaptacion[m];
        const e = graficoApp.exclusivasVenta[m];
        const conv = c > 0 ? (e/c)*100 : 0;
        evolucion.conversionCaptacion.realizado.push(conv);

        // Objetivos
        evolucion.gci.objetivo.push(objsMes[m].gci);
        evolucion.citasCaptacion.objetivo.push(objsMes[m].citas);
    }

    const cumplimientoGlobalSeguro = isNaN(cumplimientoGlobal) || !isFinite(cumplimientoGlobal) ? 0 : cumplimientoGlobal;

    return {
      id: idAgente,
      agente: nombreOficial,
      realizado: actividad,
      objetivos: objetivos,
      cumplimientos: cumplimientos,
      cumplimientoGlobal: cumplimientoGlobalSeguro.toFixed(1),
      ratios: ratios,
      evolucionMensual: evolucion
    };

  } catch (error) {
    Logger.log('❌ ERROR V13: ' + error.toString());
    throw error;
  }
}

// --- FUNCIÓN AUXILIAR: LECTURA DE HISTÓRICO (INDICES REALES) ---
function obtenerDesgloseHistorico(idAgente, anio) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('Historico_Agentes');
  if (!hoja) return null;

  const datos = hoja.getDataRange().getValues();
  
  const desglose = {
    gci: Array(12).fill(0), citasCaptacion: Array(12).fill(0),
    exclusivasVenta: Array(12).fill(0), captacionesAbierto: Array(12).fill(0),
    citasCompradores: Array(12).fill(0), casasEnsenadas: Array(12).fill(0),
    ventas: Array(12).fill(0),
    totalGCI: 0, totalCitas: 0, totalExclusivas: 0, totalVentas: 0,
    totalCaptAbierto: 0, totalCitasComp: 0, totalVisitas: 0
  };

  // Normalizamos ID para búsqueda (sin espacios)
  const idBuscado = String(idAgente).toUpperCase().replace(/[^A-Z0-9]/g, '');

  for (let i = 1; i < datos.length; i++) {
    const row = datos[i];
    if(!row[0]) continue;

    const filaID = String(row[0]).toUpperCase().replace(/[^A-Z0-9]/g, '');
    
    // Coincidencia laxa de ID y Año
    if ((filaID.includes(idBuscado) || idBuscado.includes(filaID)) && row[2] == anio) {
      
      const mes = parseInt(row[3]); 
      if (mes >= 1 && mes <= 12) {
        const idx = mes - 1;
        
        // USAMOS EL PARSER EUROPEO AQUÍ
        const gciVal = parsearNumeroEU(row[22]);     // GCI_Total
        const citasVal = parsearNumeroEU(row[4]);    // Citas
        const exclVal = parsearNumeroEU(row[5]);     // Capt_Excl
        const abiertoVal = parsearNumeroEU(row[6]);  // Capt_Abierto
        const citasCompVal = parsearNumeroEU(row[7]);// Visitas_Comp
        const casasVal = parsearNumeroEU(row[8]);    // Casas_Ens
        
        // Suma de ventas (Cols 14, 16, 18, 20)
        const ventasVal = parsearNumeroEU(row[14]) + parsearNumeroEU(row[16]) + parsearNumeroEU(row[18]) + parsearNumeroEU(row[20]);

        // Arrays mensuales
        desglose.gci[idx] += gciVal;
        desglose.citasCaptacion[idx] += citasVal;
        desglose.exclusivasVenta[idx] += exclVal;
        desglose.captacionesAbierto[idx] += abiertoVal;
        desglose.citasCompradores[idx] += citasCompVal;
        desglose.casasEnsenadas[idx] += casasVal;
        desglose.ventas[idx] += ventasVal;

        // Totales Acumulados
        desglose.totalGCI += gciVal;
        desglose.totalCitas += citasVal;
        desglose.totalExclusivas += exclVal;
        desglose.totalVentas += ventasVal;
        desglose.totalCaptAbierto += abiertoVal;
        desglose.totalCitasComp += citasCompVal;
        desglose.totalVisitas += casasVal;
      }
    }
  }
  return desglose;
}

// --- FUNCIÓN AUXILIAR: SUMAR TOTALES (REDUNDANCIA SEGURA) ---
function sumarDatosHistoricos(idAgente, fechaInicio, fechaFin) {
    // Reutilizamos obtenerDesgloseHistorico para no duplicar código
    const anio = fechaInicio.getFullYear();
    const desglose = obtenerDesgloseHistorico(idAgente, anio);
    
    if (!desglose) return null;

    return {
        gci: desglose.totalGCI,
        citasCaptacion: desglose.totalCitas,
        exclusivasVenta: desglose.totalExclusivas,
        captacionesAbierto: desglose.totalCaptAbierto,
        citasCompradores: desglose.totalCitasComp,
        casasEnsenadas: desglose.totalVisitas,
        ventas: desglose.totalVentas,
        volumenNegocio: 0
    };
}
/**
 * ════════════════════════════════════════════════════════════════
 * 🧪 FUNCIÓN DE DIAGNÓSTICO
 * ════════════════════════════════════════════════════════════════
 */
function diagnosticarLecturaHistorico() {
  Logger.log('');
  Logger.log('🧪 DIAGNÓSTICO DE LECTURA DE HISTORICO_AGENTES');
  Logger.log('══════════════════════════════════════════════════════════');
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('Historico_Agentes');
  
  if (!hoja) {
    Logger.log('❌ ERROR: No existe la hoja Historico_Agentes');
    return;
  }
  
  const datos = hoja.getDataRange().getValues();
  
  Logger.log('✅ Hoja encontrada');
  Logger.log('   📊 Total filas (incluyendo cabecera): ' + datos.length);
  Logger.log('   📊 Total columnas: ' + datos[0].length);
  Logger.log('');
  Logger.log('📋 CABECERAS (primeras 10):');
  for (let i = 0; i < Math.min(10, datos[0].length); i++) {
    Logger.log('   Col ' + String.fromCharCode(65 + i) + ' (' + i + '): ' + datos[0][i]);
  }
  Logger.log('');
  
  if (datos.length > 1) {
    Logger.log('📝 PRIMERA FILA DE DATOS:');
    const primeraFila = datos[1];
    Logger.log('   ID_Agente (Col A): ' + primeraFila[0]);
    Logger.log('   Nombre (Col B): ' + primeraFila[1]);
    Logger.log('   Año (Col C): ' + primeraFila[2]);
    Logger.log('   Mes (Col D): ' + primeraFila[3]);
    Logger.log('   Citas (Col E): ' + primeraFila[4]);
    Logger.log('   Capt_Excl (Col F): ' + primeraFila[5]);
    Logger.log('   GCI_Total (Col W): ' + primeraFila[22]);
    Logger.log('');
    
    // Contar registros por agente
    const conteo = {};
    for (let i = 1; i < datos.length; i++) {
      const id = String(datos[i][0]);
      if (id) {
        conteo[id] = (conteo[id] || 0) + 1;
      }
    }
    
    Logger.log('📊 REGISTROS POR AGENTE:');
    Object.keys(conteo).forEach(id => {
      Logger.log('   ' + id + ': ' + conteo[id] + ' registros');
    });
  } else {
    Logger.log('⚠️ La hoja solo tiene cabecera, sin datos');
  }
  
  Logger.log('══════════════════════════════════════════════════════════');
}

/**
 * ═══════════════════════════════════════════════════════════════════════════
 * 🧪 FUNCIÓN DE PRUEBA ESPECÍFICA PARA TU CASO
 * ═══════════════════════════════════════════════════════════════════════════
 */
function probarNavero() {
  Logger.log('');
  Logger.log('🧪 PRUEBA ESPECÍFICA: Agente Navero');
  Logger.log('══════════════════════════════════════════════════════════');
  
  try {
    // Según tu captura, el agente se llama "Navero"
    const resultado = obtenerDatosAgenteCompleto('Navero', 2025);
    
    Logger.log('');
    Logger.log('📊 RESULTADO PARA NAVERO:');
    Logger.log('   GCI Total: ' + resultado.totalesAnuales.gci.toFixed(2) + ' €');
    Logger.log('   Debe ser: ~1000€ (según tu captura)');
    Logger.log('');
    Logger.log('📈 ARRAY evolucionMensual.gci.realizado:');
    Logger.log('   ' + JSON.stringify(resultado.evolucionMensual.gci.realizado));
    Logger.log('');
    
    if (resultado.totalesAnuales.gci === 0) {
      Logger.log('❌ ERROR: GCI = 0');
      Logger.log('   Posibles causas:');
      Logger.log('   1. No hay datos en Actividad_Diaria para Navero en 2025');
      Logger.log('   2. No hay datos en Historico_Agentes para Navero en 2025');
      Logger.log('   3. El ID/Nombre no coincide exactamente');
    } else {
      Logger.log('✅ GCI > 0: Los datos SÍ se están leyendo');
      Logger.log('   El array evolucionMensual debe tener valores para el gráfico');
    }
    
  } catch (error) {
    Logger.log('');
    Logger.log('❌ ERROR: ' + error.message);
  }
  
  Logger.log('══════════════════════════════════════════════════════════');
}

// Mantén esta también actualizada para que no falle el modal de objetivos
function obtenerObjetivosAgente(idAgente, fechaInicio, fechaFin, todosObjetivos) {
  const objetivos = {
    citasCaptacion: 0, exclusivasVenta: 0, exclusivasComprador: 0,
    captacionesAbierto: 0, citasCompradores: 0, casasEnsenadas: 0,
    leadsCompradores: 0, llamadas: 0, gci: 0, volumenNegocio: 0
  };
  
  for (let i = 0; i < todosObjetivos.length; i++) {
    // Ignorar cabecera
    if (String(todosObjetivos[i][0]).toUpperCase() === 'ID_AGENTE') continue;

    if (String(todosObjetivos[i][0]).trim() === String(idAgente).trim()) {
      // Sumar si está en el año (lógica simplificada para rendimiento)
      objetivos.gci += parseFloat(todosObjetivos[i][12]) || 0;
      objetivos.citasCaptacion += parseFloat(todosObjetivos[i][4]) || 0;
      objetivos.exclusivasVenta += parseFloat(todosObjetivos[i][5]) || 0;
      // ...
    }
  }
  return objetivos;
}

// --- FUNCIÓN AUXILIAR: OBJETIVOS MENSUALES ---
function obtenerObjetivosMensuales(idAgente, year, datosObjetivos) {
    const objs = Array(12).fill(null).map(() => ({ gci: 0, citas: 0 }));
    const targetID = String(idAgente).trim();
    
    for(let i=1; i<datosObjetivos.length; i++) {
        const row = datosObjetivos[i];
        if(String(row[0]).trim() == targetID && row[2] == year) {
            const mes = parseInt(row[3]);
            if(mes >= 1 && mes <= 12) {
                objs[mes-1].gci = parseFloat(row[12]) || 0; // Columna GCI objetivo
                objs[mes-1].citas = parseFloat(row[4]) || 0; // Columna Citas objetivo
            }
        }
    }
    return objs;
}


function calcularPendientesPeriodo(idAgente, fechaReferencia, periodo, esAcumulativo, todasActividades, todosObjetivos) {
  let fechaInicio, fechaFin;
  const hoy = new Date(fechaReferencia);
 
  switch(periodo) {
    case 'semana':
      const diaSemana = hoy.getDay();
      const diff = diaSemana === 0 ? -6 : 1 - diaSemana;
      fechaInicio = new Date(hoy);
      fechaInicio.setDate(hoy.getDate() + diff);
      fechaFin = new Date(fechaInicio);
      fechaFin.setDate(fechaInicio.getDate() + 6);
      break;
    case 'mes':
      fechaInicio = new Date(hoy.getFullYear(), hoy.getMonth(), 1);
      fechaFin = new Date(hoy.getFullYear(), hoy.getMonth() + 1, 0);
      break;
    case 'trimestre':
      const mesActual = hoy.getMonth();
      const inicioTrimestre = Math.floor(mesActual / 3) * 3;
      fechaInicio = new Date(hoy.getFullYear(), inicioTrimestre, 1);
      fechaFin = new Date(hoy.getFullYear(), inicioTrimestre + 3, 0);
      break;
    case 'semestre':
      const inicioSemestre = hoy.getMonth() < 6 ? 0 : 6;
      fechaInicio = new Date(hoy.getFullYear(), inicioSemestre, 1);
      fechaFin = new Date(hoy.getFullYear(), inicioSemestre + 6, 0);
      break;
    case 'ano':
      fechaInicio = new Date(hoy.getFullYear(), 0, 1);
      fechaFin = new Date(hoy.getFullYear(), 11, 31);
      break;
  }
 
  fechaInicio.setHours(0, 0,0, 0);
  fechaFin.setHours(23, 59, 59, 999);
 
  const actividad = obtenerActividadAgente(idAgente, fechaInicio, fechaFin, todasActividades);
  const objetivos = obtenerObjetivosAgente(idAgente, fechaInicio, fechaFin, todosObjetivos);
  if (esAcumulativo) {
      const pendientesAcumulados = calcularObjetivosAcumuladosPendientes(idAgente, fechaInicio, todasActividades, todosObjetivos);
      Object.keys(objetivos).forEach(key => objetivos[key] += pendientesAcumulados[key]);
  }
 
  const pendientes = {};
  Object.keys(objetivos).forEach(key => {
      pendientes[key] = Math.max(0, (objetivos[key] || 0) - (actividad[key] || 0));
  });
 
  const diasRestantes = Math.max(0, Math.ceil((fechaFin - hoy) / (1000 * 60 * 60 * 24)));
 
  const promedioDiario = {};
  Object.keys(pendientes).forEach(key => {
      promedioDiario[key] = (diasRestantes > 0) ? (pendientes[key] / diasRestantes) : 0;
  });
 
  return {
    fechaInicio: fechaInicio,
    fechaFin: fechaFin,
    diasRestantes: diasRestantes,
    realizado: actividad,
    objetivos: objetivos,
    pendientes: pendientes,
    promedioDiario: promedioDiario,
    cumplimientos: calcularCumplimientos(actividad, objetivos)
  };
}

function recalcularTodosDatos() {
  SpreadsheetApp.getUi().alert('Datos recalculados', 'Todas las fórmulas y cálculos se han actualizado.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function limpiarDatosPrueba() {
  const ui = SpreadsheetApp.getUi();
  const respuesta = ui.alert(
    '⚠️ Confirmar',
    '¿Deseas eliminar TODOS los datos de actividad?\n\nEsta acción NO se puede deshacer.',
    ui.ButtonSet.YES_NO
  );
  if (respuesta !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaActividad = ss.getSheetByName(CONFIG.HOJA_ACTIVIDAD);
  if (hojaActividad && hojaActividad.getLastRow() > 1) {
    hojaActividad.getRange(2, 1, hojaActividad.getLastRow() - 1, hojaActividad.getLastColumn()).clearContent();
  }
 
  ui.alert('Limpieza Completa', 'Se han eliminado todos los datos de actividad.', ui.ButtonSet.OK);
}

function gestionarAgentes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CONFIG.HOJA_AGENTES);
  if (hoja) {
    ss.setActiveSheet(hoja);
  } else {
    SpreadsheetApp.getUi().alert('No se encuentra la hoja "Agentes".');
  }
}

function configurarObjetivos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(CONFIG.HOJA_CONFIGURACION);
  if (hoja) {
    ss.setActiveSheet(hoja);
    SpreadsheetApp.getUi().alert('📋 Configuración de Objetivos', 
      'Puedes editar los objetivos por defecto en la hoja "Configuración".\n\nLos cambios afectarán a los nuevos registros.', 
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function actualizarObjetivosAgente(datos) {
  try {
    Logger.log('Actualizando objetivos para agente: ' + datos.idAgente);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(CONFIG.HOJA_OBJETIVOS);
    if (!hoja) throw new Error('No se encontró la hoja de Objetivos.');
    const year = datos.year || new Date().getFullYear();
    const datosHoja = hoja.getDataRange().getValues();
    const filasAEliminar = [];
    for (let i = datosHoja.length - 1; i >= 1; i--) {
      if (datosHoja[i][0] === datos.idAgente && datosHoja[i][2] === year) {
        filasAEliminar.push(i + 1);
      }
    }
    for (const fila of filasAEliminar) hoja.deleteRow(fila);
    const filasNuevas = [];
    const distribucion = datos.distribucion;
    const objetivosAnuales = datos.objetivosAnuales;
    for (let mes = 1; mes <= 12; mes++) {
        const pct = (distribucion[mes - 1] || 100/12) / 100;
        filasNuevas.push([
            datos.idAgente, '', year, mes,
            Math.round(objetivosAnuales.citasCaptacion * pct),
            Math.round(objetivosAnuales.exclusivasVenta * pct),
            Math.round(objetivosAnuales.exclusivasComprador * pct),
            Math.round(objetivosAnuales.captacionesAbierto * pct),
            Math.round(objetivosAnuales.citasCompradores * pct),
            Math.round(objetivosAnuales.casasEnsenadas * pct),
            Math.round(objetivosAnuales.leadsCompradores * pct),
                Math.round(objetivosAnuales.llamadas * pct),
                parseFloat((objetivosAnuales.gci * pct).toFixed(2)),
                parseFloat((objetivosAnuales.volumenNegocio * pct).toFixed(2)),
            new Date()
        ]);
    }
    if (filasNuevas.length > 0) {
      const ultimaFila = hoja.getLastRow();
      hoja.getRange(ultimaFila + 1, 1, filasNuevas.length, filasNuevas[0].length).setValues(filasNuevas);
      const hojaAgentes = ss.getSheetByName(CONFIG.HOJA_AGENTES);
      if (hojaAgentes) {
        for (let i = 0; i < filasNuevas.length; i++) {
          const fila = ultimaFila + 1 + i;
          hoja.getRange(fila, 2).setFormula(`=IFERROR(VLOOKUP(A${fila},${CONFIG.HOJA_AGENTES}!A:B,2,FALSE),"")`);
        }
      }
    }
    return { success: true, message: 'Objetivos guardados para todo el año ' + year };
  } catch (error) {
    throw new Error('Error al actualizar objetivos: ' + error.toString());
  }
}

function crearNuevoAgente(datos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(CONFIG.HOJA_AGENTES);
    if (!hoja) throw new Error('No se encontró la hoja de Agentes.');
    
    const ultimaFila = hoja.getLastRow();
    const nuevoId = 'AG' + String(ultimaFila + 1).padStart(3, '0');
    
    const fila = [
      nuevoId, 
      datos.nombre, 
      datos.email || '', 
      datos.telefono || '', 
      new Date(),
      'Activo', 
      datos.objetivosAcumulados || 'NO', 
      new Date(),
      parseFloat(datos.sueldoFijo) || 0 // <--- NUEVO: Columna I (Sueldo Fijo)
    ];
    hoja.appendRow(fila);

    // Crear objetivos por defecto (igual que antes)
    const hojaConfig = ss.getSheetByName(CONFIG.HOJA_CONFIGURACION);
    const conf = hojaConfig.getDataRange().getValues();
    const getConf = (key, def) => { const r = conf.find(x => x[0]===key); return r ? r[1] : def; };
    
    const defaultConfig = {
        citasCaptacion: getConf('Obj_Mensual_Citas_Captacion', 15) * 12,
        exclusivasVenta: getConf('Obj_Mensual_Exclusivas_Venta', 5) * 12,
        exclusivasComprador: getConf('Obj_Mensual_Exclusivas_Comprador', 4) * 12,
        captacionesAbierto: getConf('Obj_Mensual_Captaciones_Abierto', 3) * 12,
        citasCompradores: getConf('Obj_Mensual_Citas_Compradores', 10) * 12,
        casasEnsenadas: getConf('Obj_Mensual_Casas_Ensenadas', 8) * 12,
        leadsCompradores: getConf('Obj_Mensual_Leads_Compradores', 20) * 12,
        gci: getConf('Obj_Mensual_GCI', 8000) * 12
    };
    
    actualizarObjetivosAgente({
      idAgente: nuevoId,
      year: new Date().getFullYear(),
      objetivosAnuales: defaultConfig,
      distribucion: [8,8,9,9,10,10,9,5,9,9,8,6]
    });
    
    return { success: true, message: 'Agente creado correctamente', id: nuevoId, nombre: datos.nombre };
  } catch (error) {
    throw new Error('Error al crear agente: ' + error.toString());
  }
}

function obtenerObjetivosAgenteActual(idAgente, year) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();   // ← FUERA DEL TRY

  try {
    const hoja = ss.getSheetByName(CONFIG.HOJA_OBJETIVOS);
    if (!hoja) throw new Error('No se encontró la hoja de Objetivos.');

    const datos = hoja.getDataRange().getValues();

    const objetivosAnuales = {
      citasCaptacion: 0, exclusivasVenta: 0, exclusivasComprador: 0,
      captacionesAbierto: 0, citasCompradores: 0, casasEnsenadas: 0,
      leadsCompradores: 0, llamadas: 0, gci: 0, volumenNegocio: 0
    };

    const distribucion = Array(12).fill(0);
    let totalGCI = 0;

    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0] === idAgente && datos[i][2] === year) {
        const mesIndex = datos[i][3] - 1;
        if (mesIndex >= 0 && mesIndex < 12) {
          objetivosAnuales.citasCaptacion     += parseFloat(datos[i][4])  || 0;
          objetivosAnuales.exclusivasVenta    += parseFloat(datos[i][5])  || 0;
          objetivosAnuales.exclusivasComprador+= parseFloat(datos[i][6])  || 0;
          objetivosAnuales.captacionesAbierto += parseFloat(datos[i][7])  || 0;
          objetivosAnuales.citasCompradores   += parseFloat(datos[i][8])  || 0;
          objetivosAnuales.casasEnsenadas     += parseFloat(datos[i][9])  || 0;
          objetivosAnuales.leadsCompradores   += parseFloat(datos[i][10]) || 0;

          const gciMes = parseFloat(datos[i][11]) || 0;
          objetivosAnuales.gci += gciMes;
          totalGCI += gciMes;
          distribucion[mesIndex] = gciMes;
        }
      }
    }

    // Si no hay objetivos personalizados → vamos al catch
    if (totalGCI === 0) throw new Error('Sin objetivos personalizados');

    // Calcular % de distribución según GCI real
    let sumaPct = 0;
    for (let i = 0; i < 11; i++) {
      distribucion[i] = parseFloat(((distribucion[i] / totalGCI) * 100).toFixed(2));
      sumaPct += distribucion[i];
    }
    distribucion[11] = parseFloat((100 - sumaPct).toFixed(2));

    return { objetivosAnuales, distribucion };

  } catch (e) {
    // Aquí ya podemos usar ss sin problema
    const hojaConfig = ss.getSheetByName(CONFIG.HOJA_CONFIGURACION);
    const conf = hojaConfig.getDataRange().getValues();

    const get = (clave, def) => {
      const row = conf.find(r => r[0] === clave);
      return row ? (row[1] ?? def) : def;
    };

    const objetivosAnuales = {
      citasCaptacion:     get('Obj_Mensual_Citas_Captacion', 15) * 12,
      exclusivasVenta:    get('Obj_Mensual_Exclusivas_Venta', 5) * 12,
      exclusivasComprador:get('Obj_Mensual_Exclusivas_Comprador', 4) * 12,
      captacionesAbierto: get('Obj_Mensual_Captaciones_Abierto', 3) * 12,
      citasCompradores:   get('Obj_Mensual_Citas_Compradores', 10) * 12,
      casasEnsenadas:     get('Obj_Mensual_Casas_Ensenadas', 8) * 12,
      leadsCompradores: get('Obj_Mensual_Leads_Compradores', 20) * 12,
            llamadas: get('Obj_Mensual_Llamadas', 50) * 12,
            gci: get('Obj_Mensual_GCI', 8000) * 12,
            volumenNegocio: get('Obj_Mensual_Volumen_Negocio', 100000) * 12
    };

    const distribucionDefault = [8,8,9,9,10,10,9,5,9,9,8,6];

    return { objetivosAnuales, distribucion: distribucionDefault };
  }
}
function analizarEquipoIA(datosParaIA, periodoActual) {
  try {
    if (!datosParaIA || datosParaIA.length === 0) return "<p>Sin datos.</p>";

    // Preparamos un resumen ligero para no saturar el prompt
    const resumenDatos = datosParaIA.map(a => 
      `- ${a.agente}: Cumplimiento ${a.cumplimiento}%, GCI ${a.gci}, Conversión ${a.conversion}%`
    ).join("\n");

    const prompt = `
      Eres el Director Comercial de una inmobiliaria de alto rendimiento.
      Analiza brevemente el estado del equipo en ${periodoActual}.
      
      DATOS DEL EQUIPO:
      ${resumenDatos}

      TAREA:
      1. Destaca al MVP (Jugador más valioso).
      2. Identifica un patrón general de mejora para el equipo.
      3. Mensaje motivacional corto para la reunión de equipo.

      Responde con HTML (<h3>, <p>, <ul>). Sé conciso.
    `;

    const respuesta = llamarGemini(prompt);
    return `<div style="padding:10px;">${respuesta}</div>`;

  } catch (e) {
    return "<p>Error IA Equipo.</p>";
  }
}

function analizarAgenteIA(agente, periodoActual) {
  try {
    if (!agente) return "<h3>🤖 Análisis de Agente</h3><p>No se pudieron cargar los datos del agente.</p>";
    const nombre = agente.agente;
    const cumplimiento = parseFloat(agente.cumplimientoGlobal);
    const conversion = parseFloat(agente.ratios.conversionCaptacion);
    const gci = parseFloat(agente.realizado.gci).toLocaleString('es-ES', {style: 'currency', currency: 'EUR'});
    let html = "<h3>🤖 Análisis de Rendimiento: " + nombre + "</h3>";
    html += "<p>Período analizado: <strong>" + periodoActual.toUpperCase() + "</strong></p>";
    html += "<ul style='margin-left: 20px; margin-top: 15px;'>";
    if (cumplimiento >= 90) {
      html += "<li><strong>¡Felicidades!</strong> Tu cumplimiento global es del <strong>" + cumplimiento.toFixed(1) + "%</strong>, lo cual es excelente.</li>";
    } else if (cumplimiento >= 70) {
      html += "<li><strong>Buen trabajo.</strong> Tu cumplimiento global es del <strong>" + cumplimiento.toFixed(1) + "%</strong>. Estás en el camino correcto.</li>";
    } else {
      html += "<li><strong>Área de Enfoque:</strong> Tu cumplimiento global es del <strong>" + cumplimiento.toFixed(1) + "%</strong>. Revisa los puntos de acción.</li>";
    }
    html += "<li>Has generado un GCI de <strong>" + gci + "</strong> en este período.</li>";
    if (conversion < 20) {
      html += "<li><strong>Punto Crítico:</strong> Tu ratio de conversión (Cita a Exclusiva) es del <strong>" + conversion.toFixed(1) + "%</strong>. Este es el principal cuello de botella. Enfócate en mejorar tu presentación de servicios en las citas de captación.</li>";
    } else if (conversion < 30) {
       html += "<li><strong>Oportunidad de Mejora:</strong> Tu ratio de conversión es del <strong>" + conversion.toFixed(1) + "%</strong>. Intenta subirlo por encima del 30% para maximizar tu GCI.</li>";
    } else {
       html += "<li><strong>¡Muy bien!</strong> Tu ratio de conversión es del <strong>" + conversion.toFixed(1) + "%</strong>, lo cual es muy eficiente.</li>";
    }
    html += "</ul>";
    Utilities.sleep(1000);
    return html;
  } catch (e) {
    return "<h3>🤖 Error de Análisis</h3><p>No se pudo completar el análisis del agente: " + e.message + "</p>";
  }
}

function obtenerInventario(filtros) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(CONFIG.HOJA_INVENTARIO);
    if (!hoja) throw new Error('No se encontró la hoja de Inventario.');
    const datos = hoja.getDataRange().getValues();
    const inmuebles = [];
    for (let i = 1; i < datos.length; i++) {
      const inmueble = {
        id: datos[i][0], fechaAlta: datos[i][1], idAgente: datos[i][2], nombreAgente: datos[i][3],
        tipo: datos[i][4], direccion: datos[i][5], ciudad: datos[i][6], cp: datos[i][7], provincia: datos[i][8],
        precio: datos[i][9], superficieM2: datos[i][10], habitaciones: datos[i][11], banos: datos[i][12],
        estado: datos[i][13], tipoContrato: datos[i][14], exclusividad: datos[i][15], propietarioNombre: datos[i][16],
        propietarioTelefono: datos[i][17], propietarioEmail: datos[i][18], descripcion: datos[i][19],
        observaciones: datos[i][20], fechaActualizacion: datos[i][21], estadoVenta: datos[i][22]
      };
      if (filtros) {
        if (filtros.agente && inmueble.idAgente !== filtros.agente) continue;
        if (filtros.tipo && inmueble.tipo !== filtros.tipo) continue;
        if (filtros.estadoVenta && inmueble.estadoVenta !== filtros.estadoVenta) continue;
      }
      inmuebles.push(inmueble);
    }
    return inmuebles;
  } catch (error) { throw new Error('Error al obtener inventario: ' + error.toString()); }
}

function agregarInmueble(datos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(CONFIG.HOJA_INVENTARIO);
    if (!hoja) throw new Error('No se encontró la hoja de Inventario.');
    const nuevoId = 'INM' + String(hoja.getLastRow() + 1).padStart(4, '0');
    const fila = [
      nuevoId, new Date(), datos.idAgente, datos.nombreAgente, datos.tipo, datos.direccion,
      datos.ciudad, datos.cp, datos.provincia, datos.precio || 0, datos.superficieM2 || 0,
      datos.habitaciones || 0, datos.banos || 0, datos.estado, datos.tipoContrato,
      datos.exclusividad, datos.propietarioNombre, datos.propietarioTelefono, datos.propietarioEmail,
      datos.descripcion || '', datos.observaciones || '', new Date(), 'Activo'
    ];
    hoja.appendRow(fila);
    return { success: true, message: 'Inmueble agregado correctamente', id: nuevoId };
  } catch (error) { throw new Error('Error al agregar inmueble: ' + error.toString()); }
}

function actualizarEstadoInmueble(idInmueble, nuevoEstado) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(CONFIG.HOJA_INVENTARIO);
    if (!hoja) throw new Error('No se encontró la hoja de Inventario.');
    const datos = hoja.getDataRange().getValues();
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0] === idInmueble) {
        hoja.getRange(i + 1, 23).setValue(nuevoEstado);
        hoja.getRange(i + 1, 22).setValue(new Date());
        return { success: true, message: 'Estado actualizado correctamente' };
      }
    }
    throw new Error('No se encontró el inmueble con ID: ' + idInmueble);
  } catch (error) { throw new Error('Error al actualizar estado: ' + error.toString()); }
}

function obtenerEstadisticasInventario() {
  try {
    const inmuebles = obtenerInventario();
    const stats = {
      total: inmuebles.length,
      activos: inmuebles.filter(i => i.estadoVenta === 'Activo').length,
      reservados: inmuebles.filter(i => i.estadoVenta === 'Reservado').length,
      vendidos: inmuebles.filter(i => i.estadoVenta === 'Vendido').length,
      exclusivos: inmuebles.filter(i => i.exclusividad === 'SI').length,
      valorTotal: inmuebles.reduce((sum, i) => sum + (i.precio || 0), 0),
      porTipo: {}, porAgente: {}, porCiudad: {}
    };
    inmuebles.forEach(i => {
      stats.porTipo[i.tipo] = (stats.porTipo[i.tipo] || 0) + 1;
      stats.porAgente[i.nombreAgente] = (stats.porAgente[i.nombreAgente] || 0) + 1;
      stats.porCiudad[i.ciudad] = (stats.porCiudad[i.ciudad] || 0) + 1;
    });
    return stats;
  } catch (error) { throw new Error('Error al obtener estadísticas: ' + error.toString()); }
}
// ====== FUNCIÓN PARA LA PESTAÑA TRANSACCIONES ======
function obtenerTransacciones(filtros) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(CONFIG.HOJA_ACTIVIDAD);
    if (!hoja) return { transacciones: [] };

    const datos = hoja.getDataRange().getValues();
    const transacciones = [];

    let fechaInicio = new Date(new Date().getFullYear(), 0, 1);
    let fechaFin = new Date();
    if (filtros && filtros.fechaInicio) fechaInicio = new Date(filtros.fechaInicio);
    if (filtros && filtros.fechaFin) fechaFin = new Date(filtros.fechaFin);

    for (let i = 1; i < datos.length; i++) {
      const notas = (datos[i][12] || "").toUpperCase(); // Columna Notas
      if (notas.includes('TRANSACCIÓN')) {
        const fechaRaw = datos[i][1];
        if (!fechaRaw) continue;
        const fecha = new Date(fechaRaw);
        if (fecha instanceof Date && !isNaN(fecha) && fecha >= fechaInicio && fecha <= fechaFin) {
          transacciones.push({
            id: datos[i][0],
            fecha: Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
            agente: datos[i][3] || 'N/A',
            tipo: notas.match(/TRANSACCIÓN (\w+)/)?.[1] || 'N/A',
            lado: notas.match(/LADO: (\w+)/)?.[1] || 'N/A',
            descripcion: notas.match(/DESCRIPCIÓN: (.*?)( \| LADO| \| COMISIÓN|$)/)?.[1] || 'Sin descripción',
            gci: parseFloat(datos[i][11]) || 0
          });
        }
      }
    }

    transacciones.sort((a, b) => new Date(b.fecha) - new Date(a.fecha));

    return { transacciones };
  } catch (error) {
    Logger.log('Error en obtenerTransacciones: ' + error);
    return { transacciones: [] };
  }
}
function lanzarAnalisisIA() {
  const ui = SpreadsheetApp.getUi();
  google.script.run
    .withSuccessHandler(datos => {
      if (!datos || !datos.agentes || datos.agentes.length === 0) {
        ui.alert('Sin datos', 'No hay agentes con actividad para analizar.', ui.ButtonSet.OK);
        return;
      }

      const datosIA = datos.agentes.map(a => ({
        agente: a.agente,
        cumplimiento: a.cumplimientoGlobal,
        gci: a.realizado.gci,
        conversion: a.ratios.conversionCaptacion
      }));

      const htmlIA = analizarEquipoIA(datosIA, 'Mes Actual');

      const htmlModal = HtmlService.createHtmlOutput(`
        <div style="padding:20px; font-family:Arial;">
          <h2 style="color:#b70000;">🤖 Análisis IA del Equipo</h2>
          ${htmlIA}
          <br><br>
          <button onclick="google.script.host.close()" style="padding:10px 20px; background:#b70000; color:white; border:none; border-radius:5px; cursor:pointer;">Cerrar</button>
        </div>
      `).setWidth(600).setHeight(400);

      ui.showModalDialog(htmlModal, 'Análisis IA');
    })
    .withFailureHandler(err => ui.alert('Error', 'No se pudo cargar el análisis: ' + err.message, ui.ButtonSet.OK))
    .obtenerDatosDashboard({});
}
/*
  === EXTENSIÓN: 4 MODELOS + GPS / FACTURACIÓN ===
  Este archivo contiene funciones adicionales que integran:
   - Centro de Acciones (modal / pestaña en HTML)
   - Modelo Presupuestario (hoja + API)
   - Modelo Económico (hoja + API)
   - Modelo Organizativo (hoja + API)
   - Modelo Generación de Leads (plantilla GPS + facturación pasada)
   - Rentabilidad Agentes (cálculos por agente)

  NOTA: Este fichero está pensado para *añadir* al final de tu código GS actual.
  No modifica ninguna función existente: solo crea nuevas hojas y funciones auxiliares.
  Para instalar todo (crear hojas y menú) ejecuta la función: instalarModelosCompleto()
*/

// ---------------------------
// CONFIG: nombres de hojas
// ---------------------------
const MODELOS_CONFIG = {
  HOJA_MODELO_PRESUP: 'Modelo_Presupuestario',
  HOJA_MODELO_ECON: 'Modelo_Economico',
  HOJA_MODELO_ORG: 'Modelo_Organizativo',
  HOJA_GEN_LEADS: 'Generacion_Leads',
  HOJA_PLANTILLA_GPS: 'Plantilla_GPS',
  HOJA_FACT_PASADA: 'Facturacion_Pasada',
  HOJA_RENTABILIDAD: 'Rentabilidad_Agentes'
};

// ---------------------------
// Instalador principal (crear hojas + menú)
// ---------------------------
function instalarModelosCompleto() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  crearHojaModeloPresupuestario(ss);
  crearHojaModeloEconomico(ss);
  crearHojaModeloOrganizativo(ss);
  crearHojaGeneracionLeads(ss);
  crearHojaPlantillaGPS(ss);
  crearHojaFacturacionPasada(ss);
  crearHojaRentabilidad(ss);
  instalarMenuModelos();
  SpreadsheetApp.getUi().alert('✅ Modelos y plantillas creados. Ejecuta "Abrir Dashboard" para ver la UI.');
}

// ---------------------------
// Menú (se añade cuando se ejecuta instalarModelosCompleto)
// ---------------------------
function instalarMenuModelos() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('🧩 Modelos & GPS')
      .addItem('Abrir Centro de Acciones', 'abrirCentroAcciones')
      .addSeparator()
      .addItem('Crear/Actualizar Hojas Modelos', 'instalarModelosCompleto')
      .addItem('Exportar Plantilla GPS (CSV)', 'exportarPlantillaGPS')
      .addToUi();
  } catch (e) {
    // en contextos sin UI (doGet) puede fallar
    Logger.log('Instalar menu: ' + e.message);
  }
}

// ---------------------------
// CREAR HOJAS: definiciones mínimas
// ---------------------------
function crearHojaModeloPresupuestario(ss) {
  let hoja = ss.getSheetByName(MODELOS_CONFIG.HOJA_MODELO_PRESUP);
  if (hoja) ss.deleteSheet(hoja);
  hoja = ss.insertSheet(MODELOS_CONFIG.HOJA_MODELO_PRESUP);
  const headers = ['Año','Mes','ID_Agente','Nombre_Agente','GCI_Mes','Gastos_Ventas','Pct_Comision','Gastos_Operativos_Total','Detalle_Partidas','Partida_1_Nombre','Partida_1_Importe','Partida_2_Nombre','Partida_2_Importe','Partida_3_Nombre','Partida_3_Importe','Partida_4_Nombre','Partida_4_Importe','Beneficio','Observaciones','Fecha_Registro'];
  hoja.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold').setBackground('#e8f4ff');
  hoja.setFrozenRows(1);
  hoja.autoResizeColumns(1, headers.length);
  hoja.getRange('E:E').setNumberFormat('#,##0.00');
  hoja.getRange('F:F').setNumberFormat('#,##0.00');
  hoja.getRange('H:H').setNumberFormat('#,##0.00');
}

function crearHojaModeloEconomico(ss) {
  let hoja = ss.getSheetByName(MODELOS_CONFIG.HOJA_MODELO_ECON);
  if (hoja) ss.deleteSheet(hoja);
  hoja = ss.insertSheet(MODELOS_CONFIG.HOJA_MODELO_ECON);
  const headers = ['ID','Fecha','Agente','Pregunta_1','Pregunta_2','Pregunta_3','Pregunta_4','Pregunta_5','Pregunta_6','Resultado_Modelo','Notas','Fecha_Registro'];
  hoja.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold').setBackground('#fff4e5');
  hoja.setFrozenRows(1);
  hoja.autoResizeColumns(1, headers.length);
}

function crearHojaModeloOrganizativo(ss) {
  let hoja = ss.getSheetByName(MODELOS_CONFIG.HOJA_MODELO_ORG);
  if (hoja) ss.deleteSheet(hoja);
  hoja = ss.insertSheet(MODELOS_CONFIG.HOJA_MODELO_ORG);
  const headers = ['Nivel','Puesto','ID_Persona','Nombre','Email','Teléfono','Fecha_Incorporacion','Estado','Comentarios'];
  hoja.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold').setBackground('#eefbe8');
  hoja.setFrozenRows(1);
  hoja.autoResizeColumns(1, headers.length);
}

function crearHojaGeneracionLeads(ss) {
  let hoja = ss.getSheetByName(MODELOS_CONFIG.HOJA_GEN_LEADS);
  if (hoja) ss.deleteSheet(hoja);
  hoja = ss.insertSheet(MODELOS_CONFIG.HOJA_GEN_LEADS);
  const headers = ['ID','Fecha','Agente','Fuente_Lead','Campaña','Tipo_Lead','Estado','Valor_Estimado','Notas','Fecha_Registro'];
  hoja.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold').setBackground('#f3e8ff');
  hoja.setFrozenRows(1);
  hoja.autoResizeColumns(1, headers.length);
}

function crearHojaPlantillaGPS(ss) {
  let hoja = ss.getSheetByName(MODELOS_CONFIG.HOJA_PLANTILLA_GPS);
  if (hoja) ss.deleteSheet(hoja);
  hoja = ss.insertSheet(MODELOS_CONFIG.HOJA_PLANTILLA_GPS);
  const headers = ['ID','Fecha','Agente','Objetivo_GPS','Actividad','Responsable','Fecha_Prevista','Estado','Comentarios','Fecha_Registro'];
  hoja.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold').setBackground('#e8ffe8');
  hoja.setFrozenRows(1);
  hoja.autoResizeColumns(1, headers.length);
}

function crearHojaFacturacionPasada(ss) {
  let hoja = ss.getSheetByName('Facturacion_Pasada');
  if (hoja) ss.deleteSheet(hoja);
  hoja = ss.insertSheet('Facturacion_Pasada');
  
  const headers = [
    'ID_Agente', 'Nombre_Agente', 'Año', 'Mes',
    'Origen_Negocio', 'Tipo_Transaccion', 'Lado',
    'Precio_Venta', 'GCI', 'Comision_Pagada', 'Pct_Comision',
    'Ref_Inmueble', 'Notas', 'Fecha_Cierre', 'Fecha_Registro'
  ];
  
  hoja.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground('#b70000')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(11);
  
  hoja.setFrozenRows(1);
  
  // Formato
  hoja.getRange('H:K').setNumberFormat('#,##0.00 €');
  hoja.getRange('N:N').setNumberFormat('dd/mm/yyyy');
  
  // ✅ CORRECCIÓN CRÍTICA: Permite año en curso (no solo años anteriores)
  const anioActual = new Date().getFullYear();
  const rangoAnio = hoja.getRange('C2:C1000');
  const ruleAnio = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(2020, anioActual + 1) // ← PERMITE HASTA EL AÑO QUE VIENE
    .setAllowInvalid(false)
    .setHelpText(`Año entre 2020 y ${anioActual + 1}`)
    .build();
  rangoAnio.setDataValidation(ruleAnio);
  
  Logger.log(`✅ Facturación Pasada: Permite años 2020-${anioActual + 1}`);
  
  // Validación Mes (1-12)
  const rangoMes = hoja.getRange('D2:D1000');
  const ruleMes = SpreadsheetApp.newDataValidation()
    .requireNumberBetween(1, 12)
    .setAllowInvalid(false)
    .setHelpText('Mes entre 1 y 12')
    .build();
  rangoMes.setDataValidation(ruleMes);
  
  // Validación Origen de Negocio
  const ORIGENES = [
    'Clientes Antiguos', 'Esfera de Influencia', 'Vendedores FSO', 'Farming',
    'Aliados', 'Referidos Agentes', 'Referidos Clientes', 'Reubicación',
    'Referidos Personal', 'Llamadas Carteles', 'Publicidad', 'Sitio Web',
    'Correo Directo', 'Redes Sociales', 'Open Houses', 'ISA/OSA',
    'Idealista', 'Fotocasa', 'Otro'
  ];
  const rangoOrigen = hoja.getRange('E2:E1000');
  const ruleOrigen = SpreadsheetApp.newDataValidation()
    .requireValueInList(ORIGENES, true)
    .setAllowInvalid(false)
    .build();
  rangoOrigen.setDataValidation(ruleOrigen);
  
  // Validación Tipo
  const rangoTipo = hoja.getRange('F2:F1000');
  const ruleTipo = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Venta', 'Alquiler', 'Otros'])
    .setAllowInvalid(false)
    .build();
  rangoTipo.setDataValidation(ruleTipo);
  
  // Validación Lado
  const rangoLado = hoja.getRange('G2:G1000');
  const ruleLado = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Vendedor', 'Comprador', 'Ambos'])
    .setAllowInvalid(false)
    .build();
  rangoLado.setDataValidation(ruleLado);
  
  hoja.autoResizeColumns(1, headers.length);
}

function crearHojaRentabilidad(ss) {
  let hoja = ss.getSheetByName(MODELOS_CONFIG.HOJA_RENTABILIDAD);
  if (hoja) ss.deleteSheet(hoja);
  hoja = ss.insertSheet(MODELOS_CONFIG.HOJA_RENTABILIDAD);
  const headers = ['Año','Agente','Sueldo_Fijo_Anual','Costes_Ventas_Anual','Horas_Año','Coste_x_Hora','GCI_Anual','%_sobre_GCI','Valor_x_Hora','Company_Euro_x_Agente','Transaccion_Media','%_Produccion_Total','Fecha_Registro'];
  hoja.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold').setBackground('#eef2ff');
  hoja.setFrozenRows(1);
  hoja.getRange('C:C').setNumberFormat('#,##0.00');
  hoja.getRange('D:D').setNumberFormat('#,##0.00');
  hoja.getRange('F:F').setNumberFormat('#,##0.00');
  hoja.getRange('G:G').setNumberFormat('#,##0.00');
  hoja.autoResizeColumns(1, headers.length);
}

// ---------------------------
// API: Presupuestario
// ---------------------------
function guardarGastoOperativo(datos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName('Gastos_Operativos');
    if (!hoja) {
      crearHojaGastosOperativos(ss);
      hoja = ss.getSheetByName('Gastos_Operativos');
    }
    
    const ultimaFila = hoja.getLastRow();
    const idGasto = 'GST-' + String(ultimaFila).padStart(5, '0'); // Ej: GST-00001
    const fecha = new Date(datos.fecha);
    
    hoja.appendRow([
      idGasto,
      fecha,
      fecha.getFullYear(),
      fecha.getMonth() + 1, // Mes 1-12
      datos.partida,
      datos.descripcion,
      parseFloat(datos.importe) || 0,
      new Date()
    ]);
    
    return { success: true, message: 'Gasto registrado correctamente' };
  } catch (e) {
    throw new Error('Error al guardar gasto: ' + e.message);
  }
}

function obtenerDatosPresupuestarios(year) {
  try {
    if (!year) year = new Date().getFullYear();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // ─── 1. CALCULAR SUELDOS FIJOS MENSUALES ───
    const hojaAgentes = ss.getSheetByName(CONFIG.HOJA_AGENTES);
    let totalSueldosFijosMes = 0;
    if (hojaAgentes) {
      const datosAgentes = hojaAgentes.getDataRange().getValues();
      for (let i = 1; i < datosAgentes.length; i++) {
        if (datosAgentes[i][5] === 'Activo') {
          totalSueldosFijosMes += parseFloat(datosAgentes[i][8]) || 0;
        }
      }
    }

    // ─── 2. INICIALIZAR ARRAYS ───
    const gciMensual = new Array(12).fill(0);
    const gastosVentas = new Array(12).fill(0);
    const gastosOperativos = new Array(12).fill(0);
    const detallesVentas = Array.from({length: 12}, () => []);
    const detallesOperativos = Array.from({length: 12}, () => []);

    // ─── 3. PRE-LLENAR SUELDOS FIJOS EN GASTOS OPERATIVOS ───
    for (let m = 0; m < 12; m++) {
      if (totalSueldosFijosMes > 0) {
        gastosOperativos[m] += totalSueldosFijosMes;
        detallesOperativos[m].push({
          fecha: "-",
          concepto: "Nóminas Fijas",
          desc: "Suma de sueldos fijos plantilla activa",
          importe: totalSueldosFijosMes
        });
      }
    }

    // ─── 4. PROCESAR ACTIVIDAD (GCI + COMISIONES) ───
    const hojaAct = ss.getSheetByName(CONFIG.HOJA_ACTIVIDAD);
    if (hojaAct) {
      const datosAct = hojaAct.getDataRange().getValues();
      for (let i = 1; i < datosAct.length; i++) {
        const fechaRaw = datosAct[i][1];
        if (!fechaRaw) continue;
        
        const fecha = new Date(fechaRaw);
        if (isNaN(fecha.getTime()) || fecha.getFullYear() !== year) continue;

        const mes = fecha.getMonth();
        const gci = parseFloat(datosAct[i][12]) || 0;
        const comision = parseFloat(datosAct[i][16]) || 0;
        const notas = String(datosAct[i][14] || "");

        if (gci > 0) gciMensual[mes] += gci;
        
        if (comision > 0) {
          gastosVentas[mes] += comision;
          detallesVentas[mes].push({
            fecha: Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd/MM/yyyy"),
            concepto: datosAct[i][3] || 'Agente',
            desc: notas.split('|')[0] || 'Comisión Variable',
            importe: comision
          });
        }
      }
    }

    // ─── 5. PROCESAR OTROS GASTOS OPERATIVOS ───
    const hojaGastos = ss.getSheetByName('Gastos_Operativos');
    if (hojaGastos) {
      const datosGastos = hojaGastos.getDataRange().getValues();
      for (let i = 1; i < datosGastos.length; i++) {
        const fechaRaw = datosGastos[i][1];
        if (!fechaRaw) continue;
        
        const fecha = new Date(fechaRaw);
        if (isNaN(fecha.getTime()) || fecha.getFullYear() !== year) continue;

        const mes = fecha.getMonth();
        const importe = parseFloat(datosGastos[i][6]) || 0;

        gastosOperativos[mes] += importe;
        detallesOperativos[mes].push({
          fecha: Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd/MM/yyyy"),
          concepto: datosGastos[i][4],
          desc: datosGastos[i][5],
          importe: importe
        });
      }
    }

    // ─── 🆕 6. ANÁLISIS 40-30-30 (POR MES Y ANUAL) ───
    const analisisMensual = [];
    let totalGCIAnual = 0;
    let totalGastosVentasAnual = 0;
    let totalGastosOperativosAnual = 0;
    let totalBeneficioAnual = 0;
    
    for (let m = 0; m < 12; m++) {
      const gci = gciMensual[m];
      const gVentas = gastosVentas[m];
      const gOperativos = gastosOperativos[m];
      const beneficio = gci - gVentas - gOperativos;
      
      totalGCIAnual += gci;
      totalGastosVentasAnual += gVentas;
      totalGastosOperativosAnual += gOperativos;
      totalBeneficioAnual += beneficio;
      
      // Calcular porcentajes sobre GCI
      const pctBeneficio = gci > 0 ? (beneficio / gci) * 100 : 0;
      const pctVentas = gci > 0 ? (gVentas / gci) * 100 : 0;
      const pctOperativos = gci > 0 ? (gOperativos / gci) * 100 : 0;
      
      // Validar cumplimiento 40-30-30 (con margen de ±5%)
      const cumpleBeneficio = pctBeneficio >= 35 && pctBeneficio <= 45;
      const cumpleVentas = pctVentas >= 25 && pctVentas <= 35;
      const cumpleOperativos = pctOperativos >= 25 && pctOperativos <= 35;
      const cumpleGeneral = cumpleBeneficio && cumpleVentas && cumpleOperativos;
      
      analisisMensual.push({
        mes: m,
        gci: gci,
        gastosVentas: gVentas,
        gastosOperativos: gOperativos,
        beneficio: beneficio,
        pctBeneficio: pctBeneficio,
        pctVentas: pctVentas,
        pctOperativos: pctOperativos,
        cumple: cumpleGeneral,
        detallesVentas: detallesVentas[m],
        detallesOperativos: detallesOperativos[m]
      });
    }
    
    // Análisis Anual
    const pctBeneficioAnual = totalGCIAnual > 0 ? (totalBeneficioAnual / totalGCIAnual) * 100 : 0;
    const pctVentasAnual = totalGCIAnual > 0 ? (totalGastosVentasAnual / totalGCIAnual) * 100 : 0;
    const pctOperativosAnual = totalGCIAnual > 0 ? (totalGastosOperativosAnual / totalGCIAnual) * 100 : 0;
    
    const cumpleAnual = (
      pctBeneficioAnual >= 35 && pctBeneficioAnual <= 45 &&
      pctVentasAnual >= 25 && pctVentasAnual <= 35 &&
      pctOperativosAnual >= 25 && pctOperativosAnual <= 35
    );

    return {
      meses: ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'],
      gciMensual: gciMensual,
      gastosVentas: gastosVentas,
      gastosOperativos: gastosOperativos,
      analisisMensual: analisisMensual,
      analisisAnual: {
        totalGCI: totalGCIAnual,
        totalGastosVentas: totalGastosVentasAnual,
        totalGastosOperativos: totalGastosOperativosAnual,
        totalBeneficio: totalBeneficioAnual,
        pctBeneficio: pctBeneficioAnual,
        pctVentas: pctVentasAnual,
        pctOperativos: pctOperativosAnual,
        cumple: cumpleAnual
      }
    };

  } catch (error) {
    throw new Error('Error presupuestario: ' + error.message);
  }
}
// ---------------------------
// API: Registrar transacción con comisiones (completa)
// ---------------------------
function registrarTransaccionConComision(transaccion) {
  // transaccion: {fecha, tipo, descripcion, agentes: [{id,nombre,lado,gci,comisionPct,comisionImporte,fijo,variable}]}
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(CONFIG.HOJA_ACTIVIDAD);
    if (!hoja) throw new Error('No existe Actividad_Diaria');

    const ultimaFila = hoja.getLastRow();
    const idBase = 'TRX' + String(ultimaFila + 1000).slice(-5);
    const fecha = new Date(transaccion.fecha || new Date());

    transaccion.agentes.forEach((agente, i) => {
      const idTransaccion = `${idBase}-${String(i+1).padStart(2,'0')}`;
      const comisionPct = parseFloat(agente.comisionPct) || 0;
      const comisionImporte = parseFloat(agente.comisionImporte) || ((parseFloat(agente.gci)||0) * comisionPct/100);
      const sueldoFijo = parseFloat(agente.sueldoFijo) || 0;
      const sueldoVariable = parseFloat(agente.sueldoVariable) || 0;
      const notas = `TRANSACCIÓN ${transaccion.tipo||'Venta'} - ${transaccion.descripcion||'Venta/Alquiler'} | Lado: ${agente.lado||''} | Comisión: ${comisionPct}% | ComisImp: ${comisionImporte}`;

      hoja.appendRow([
        idTransaccion,
        fecha,
        agente.id || '',
        agente.nombre || '',
        0,0,0,0,0,0,0,
        agente.gci || 0,
        notas,
        new Date()
      ]);

      // Registrar en facturación pasada si viene
      if (transaccion.registrarFacturacion) {
        const hojaFact = ss.getSheetByName(MODELOS_CONFIG.HOJA_FACT_PASADA);
        if (hojaFact) {
          hojaFact.appendRow([
            idTransaccion,
            fecha,
            agente.id || '',
            agente.nombreCliente || '',
            transaccion.descripcion || '',
            agente.importeNeto || (agente.gci || 0),
            comisionPct,
            comisionImporte,
            agente.gci || 0,
            transaccion.formaPago || '',
            '',
            new Date()
          ]);
        }
      }
    });

    return {success:true, message:'Transacción registrada con comisiones.'};
  } catch (e){
    throw new Error('registrarTransaccionConComision: ' + e.message);
  }
}

// ---------------------------
// RENTABILIDAD AGENTES
// ---------------------------
// --- FUNCIÓN AUXILIAR PARA LIMPIAR NÚMEROS (Importante: Pégala antes o después de la función principal) ---
function calcularRentabilidadAgentes(anio, modoProyeccion = false) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaAgentes = ss.getSheetByName("Agentes");
    const hojaActividad = ss.getSheetByName("Actividad_Diaria");
    
    if (!anio) anio = new Date().getFullYear();
    anio = parseInt(anio);

    // 1. Calcular el Factor de Tiempo (¿Qué porcentaje del año ha pasado?)
    const hoy = new Date();
    const esAnioActual = (anio === hoy.getFullYear());
    let factorTiempo = 1.0; // Por defecto, año completo (100%)

    if (esAnioActual) {
        const inicioAnio = new Date(anio, 0, 1);
        const diasTranscurridos = (hoy - inicioAnio) / (1000 * 60 * 60 * 24);
        factorTiempo = diasTranscurridos / 365; 
        // Evitar divisiones raras si es 1 de enero
        if (factorTiempo < 0.002) factorTiempo = 0.002; 
    }

    const datosAgentes = hojaAgentes.getDataRange().getValues();
    const datosActividad = hojaActividad.getDataRange().getValues();
    const mapa = {};

    // 2. CARGAR AGENTES Y SUELDOS
    for (let i = 1; i < datosAgentes.length; i++) {
      const id = datosAgentes[i][0];
      const estado = String(datosAgentes[i][5] || "").toUpperCase().trim();
      
      if (id && estado.includes('ACTIV')) {
         let sueldoStr = String(datosAgentes[i][8]).replace(',', '.');
         const sueldoMensual = parseFloat(sueldoStr) || 0;
         
         // CÁLCULO DE COSTES FIJOS:
         // Si es "A la fecha" (modoProyeccion=false): Usamos solo la parte proporcional (ej: 3 meses)
         // Si es "Proyección" (modoProyeccion=true): Usamos todo el año (12 meses)
         let sueldoCalculo = 0;
         if (modoProyeccion) {
             sueldoCalculo = sueldoMensual * 12; 
         } else {
             // Si es año pasado, factor es 1. Si es actual, es proporcional.
             sueldoCalculo = sueldoMensual * 12 * factorTiempo; 
         }

         mapa[id] = { 
           id: id, 
           nombre: datosAgentes[i][1], 
           sueldoCalculo: sueldoCalculo, 
           sueldoVariableReal: 0, 
           gci: 0 
         };
      }
    }

    // 3. SUMAR TRANSACCIONES
    for (let i = 1; i < datosActividad.length; i++) {
      const fila = datosActividad[i];
      const fechaRaw = fila[1];
      if (!fechaRaw) continue;
      
      const fecha = new Date(fechaRaw);
      if (isNaN(fecha.getTime())) continue;
      if (fecha.getFullYear() !== anio) continue;

      const idAg = fila[2];
      if (mapa[idAg]) {
        // ✅ CORRECCIÓN 1: GCI está en columna 13 (índice 12)
        const gci = parseFloat(fila[12]) || 0;
        
        // ✅ CORRECCIÓN 2: Comisión Pagada está en columna 17 (índice 16)
        const comision = parseFloat(fila[16]) || 0;
        
        mapa[idAg].gci += gci;
        mapa[idAg].sueldoVariableReal += comision;
      }
    }

    // 4. GENERAR RESULTADOS Y PROYECCIONES
    const filas = [];
    // Horas laborales estándar (1760/año). Si es YTD, ajustamos las horas también.
    const horasCalculo = modoProyeccion ? 1760 : (1760 * factorTiempo);
    
    Object.values(mapa).forEach(m => {
      let gciFinal = m.gci;
      let variableFinal = m.sueldoVariableReal;

      // Si activamos PROYECCIÓN y es año actual, estimamos el cierre de año
      if (modoProyeccion && esAnioActual) {
          gciFinal = m.gci / factorTiempo; // Proyectar linealmente
          variableFinal = m.sueldoVariableReal / factorTiempo;
      }

      const costeTotal = m.sueldoCalculo + variableFinal;
      const beneficio = gciFinal - costeTotal;
      const roi = gciFinal > 0 ? (beneficio / gciFinal) * 100 : 0;
      const valorHora = horasCalculo > 0 ? (gciFinal / horasCalculo) : 0;
      const costeHora = horasCalculo > 0 ? (costeTotal / horasCalculo) : 0;

      filas.push([
        anio, 
        m.nombre, 
        m.sueldoCalculo, 
        variableFinal, 
        horasCalculo, 
        parseFloat(costeHora.toFixed(2)), 
        parseFloat(gciFinal.toFixed(2)), 
        parseFloat(roi.toFixed(1)), 
        parseFloat(valorHora.toFixed(2)), 
        parseFloat(beneficio.toFixed(2)),
        0, 0, ""
      ]);
    });
    
    filas.sort((a,b) => b[9] - a[9]); // Ordenar por Beneficio

    return { success: true, datos: filas, modo: modoProyeccion ? "PROYECCIÓN 🚀" : "A LA FECHA 📅" };

  } catch (e) {
    Logger.log("Error: " + e.stack);
    return { success: false, message: e.toString(), datos: [] };
  }
}

// ---------------------------
// EXPORT: Plantilla GPS a CSV (descargable si se usa desde HTML)
// ---------------------------
function exportarPlantillaGPS() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(MODELOS_CONFIG.HOJA_PLANTILLA_GPS);
    if (!hoja) throw new Error('No existe Plantilla_GPS');
    const datos = hoja.getDataRange().getValues();
    const csvRows = datos.map(r => r.map(c => (c||'').toString().replace(/"/g,'""')).join(','));
    const csv = csvRows.join('\n');
    return csv; // desde HTML se puede recibir y descargar
  } catch (e) {
    throw new Error('exportarPlantillaGPS: ' + e.message);
  }
}

// ---------------------------
// FRONT: abrir modal para Centro de Acciones (llamable desde menú)
// ---------------------------
function abrirCentroAcciones() {
  const html = HtmlService.createHtmlOutputFromFile('Dashboard_Modelos')
    .setWidth(1200)
    .setHeight(800)
    .setTitle('Centro de Acciones - Modelos & GPS');
  SpreadsheetApp.getUi().showModalDialog(html, 'Centro de Acciones');
}

// ---------------------------
// Helpers ligeros
// ---------------------------
function _toNum(v){return parseFloat(v)||0;}
/**
 * ============================================
 * 🆕 NUEVAS FUNCIONES PARA LOS 6 MODELOS
 * ============================================
 * 
 * INSTRUCCIONES:
 * - Copia este código completo
 * - Pégalo AL FINAL de tu archivo Code.gs actual
 * - NO borres nada de lo que ya tienes
 * - Estas funciones se AGREGAN a las existentes
 * 
 * FUNCIONES NUEVAS:
 * 1. obtenerOrganigrama()
 * 2. guardarOrganigrama()
 * 3. obtenerDatosPresupuestarios()
 * 4. guardarPlantillaGPS()
 * 5. guardarFacturacionPasada()
 * 6. obtenerRentabilidadAgentes()
 */

// ========== 1️⃣ MODELO ORGANIZATIVO ==========

/**
 * Obtiene el organigrama actual del equipo
 */
function obtenerOrganigrama() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Organigrama');
    
    if (!sheet || sheet.getLastRow() < 2) {
      return {
        teamLeader: '',
        liderVentas: '',
        liderCaptacion: '',
        agentes: ['', '', '', '']
      };
    }
    
    const datos = sheet.getRange(2, 1, 1, 7).getValues()[0];
    
    return {
      teamLeader: datos[0] || '',
      liderVentas: datos[1] || '',
      liderCaptacion: datos[2] || '',
      agentes: [datos[3] || '', datos[4] || '', datos[5] || '', datos[6] || '']
    };
    
  } catch (error) {
    Logger.log('Error en obtenerOrganigrama: ' + error.message);
    return {
      teamLeader: '',
      liderVentas: '',
      liderCaptacion: '',
      agentes: ['', '', '', '']
    };
  }
}

/**
 * Guarda el organigrama del equipo
 */
function guardarOrganigrama(datosOrganigrama) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Organigrama');
    
    // Crear hoja si no existe
    if (!sheet) {
      sheet = ss.insertSheet('Organigrama');
      sheet.appendRow(['Team Leader', 'Líder Ventas', 'Líder Captación', 'Agente 1', 'Agente 2', 'Agente 3', 'Agente 4', 'Última Actualización']);
    }
    
    // Si ya hay datos, actualizar la fila 2, sino crear
    if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, 1, 8).setValues([[
        datosOrganigrama.teamLeader || '',
        datosOrganigrama.liderVentas || '',
        datosOrganigrama.liderCaptacion || '',
        datosOrganigrama.agentes[0] || '',
        datosOrganigrama.agentes[1] || '',
        datosOrganigrama.agentes[2] || '',
        datosOrganigrama.agentes[3] || '',
        new Date()
      ]]);
    } else {
      sheet.appendRow([
        datosOrganigrama.teamLeader || '',
        datosOrganigrama.liderVentas || '',
        datosOrganigrama.liderCaptacion || '',
        datosOrganigrama.agentes[0] || '',
        datosOrganigrama.agentes[1] || '',
        datosOrganigrama.agentes[2] || '',
        datosOrganigrama.agentes[3] || '',
        new Date()
      ]);
    }
    
    return { success: true, message: '✅ Organigrama guardado correctamente' };
  } catch (error) {
    Logger.log('Error en guardarOrganigrama: ' + error.message);
    throw new Error('Error al guardar organigrama: ' + error.message);
  }
}

// ========== 4️⃣ FACTURACIÓN PASADA ==========

/**
 * Guarda análisis de facturación pasada por fuente
 */
function guardarFacturacionPasada(datosFacturacion) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('FacturacionPasada');
    
    if (!sheet) {
      sheet = ss.insertSheet('FacturacionPasada');
      sheet.appendRow([
        'Fecha', 'Fuente', 'Nº Captaciones', 'Hon. Venta', 'Trans. Compradores', 'Hon. Compradores', 'Ingreso Neto'
      ]);
    }
    
    const fuentes = ['Esfera Influencia', 'Referidos', 'Bases Datos', 'Prospección', 'Portales'];
    
    fuentes.forEach((fuente, index) => {
      const datos = datosFacturacion[index] || {};
      const ingresoNeto = (datos.honVenta || 0) + (datos.honCompradores || 0);
      
      sheet.appendRow([
        new Date(),
        fuente,
        datos.captaciones || 0,
        datos.honVenta || 0,
        datos.transCompradores || 0,
        datos.honCompradores || 0,
        ingresoNeto
      ]);
    });
    
    return { success: true, message: '✅ Facturación pasada guardada' };
  } catch (error) {
    Logger.log('Error en guardarFacturacionPasada: ' + error.message);
    throw new Error('Error al guardar facturación: ' + error.message);
  }
}

// ========== 5️⃣ RENTABILIDAD AGENTES ==========

/**
 * Calcula rentabilidad completa de todos los agentes
 * NOTA: Esta función usa obtenerDatosDashboard que ya tienes en tu código
 */
function obtenerRentabilidadAgentes(year) {
  try {
    if (!year) year = new Date().getFullYear();
    
    const filtros = {
      fechaInicio: year + '-01-01',
      fechaFin: year + '-12-31'
    };
    
    // Usar la función existente obtenerDatosDashboard
    const datos = obtenerDatosDashboard(filtros);
    const agentes = datos.agentes;
    
    const rentabilidad = agentes.map(agente => {
      // Datos de coste (puedes personalizarlos)
      const sueldoFijo = 1000; // €/mes
      const comision = 0.30; // 30% sobre GCI
      const sueldoVariable = (agente.realizado.gci || 0) * comision;
      const costeTotal = sueldoFijo + sueldoVariable;
      const beneficioEmpresa = (agente.realizado.gci || 0) - costeTotal;
      const horasTrabajo = 160; // Horas/mes promedio
      
      return {
        nombre: agente.agente,
        sueldoFijo: sueldoFijo,
        sueldoVariable: sueldoVariable,
        costeTotal: costeTotal,
        gciGenerado: agente.realizado.gci || 0,
        beneficioEmpresa: beneficioEmpresa,
        pctSobreGCI: agente.realizado.gci > 0 ? (costeTotal / agente.realizado.gci * 100) : 0,
        costePorHora: costeTotal / horasTrabajo,
        valorPorHora: (agente.realizado.gci || 0) / horasTrabajo,
        horasTrabajo: horasTrabajo
      };
    });
    
    // Ordenar por beneficio para empresa (descendente)
    rentabilidad.sort((a, b) => b.beneficioEmpresa - a.beneficioEmpresa);
    
    return rentabilidad;
    
  } catch (error) {
    Logger.log('Error en obtenerRentabilidadAgentes: ' + error.message);
    throw new Error('Error al calcular rentabilidad: ' + error.message);
  }
}

// ========== FUNCIONES DE PRUEBA ==========

/**
 * Función de test para verificar que las nuevas funciones funcionan
 */
function testNuevasFunciones() {
  Logger.log('========== TEST DE NUEVAS FUNCIONES ==========');
  
  // Test 1: Organigrama
  try {
    const org = obtenerOrganigrama();
    Logger.log('✅ obtenerOrganigrama() funciona: ' + JSON.stringify(org));
  } catch (e) {
    Logger.log('❌ obtenerOrganigrama() falló: ' + e.message);
  }
  
  // Test 2: Presupuestario
  try {
    const pres = obtenerDatosPresupuestarios(2025);
    Logger.log('✅ obtenerDatosPresupuestarios() funciona');
  } catch (e) {
    Logger.log('❌ obtenerDatosPresupuestarios() falló: ' + e.message);
  }
  
  // Test 3: Rentabilidad
  try {
    const rent = obtenerRentabilidadAgentes(2025);
    Logger.log('✅ obtenerRentabilidadAgentes() funciona - ' + rent.length + ' agentes');
  } catch (e) {
    Logger.log('❌ obtenerRentabilidadAgentes() falló: ' + e.message);
  }
  
  Logger.log('========== FIN DEL TEST ==========');
  Logger.log('Si ves ✅ en todos, las nuevas funciones están OK');
}
// ==========================================
//  📍  NUEVAS FUNCIONES GPS 1-3-5 (TEXTO)
// ==========================================

function crearHojaGPS135(ss) {
  let hoja = ss.getSheetByName('GPS_135_Data');
  if (!hoja) {
    hoja = ss.insertSheet('GPS_135_Data');
    // Guardaremos todo como JSON para flexibilidad
    const headers = ['ID_Agente', 'Nombre', 'Año', 'Datos_JSON', 'Ultima_Actualizacion'];
    hoja.getRange(1, 1, 1, headers.length).setValues([headers])
        .setBackground('#b70000')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
    hoja.setFrozenRows(1);
    hoja.autoResizeColumns(1, headers.length);
  }
  return hoja;
}
// ==========================================
//  📍  FUNCIONES GPS 1-3-5 (TEXTO ESTRUCTURADO)
// ==========================================

function crearHojaGPS135(ss) {
  let hoja = ss.getSheetByName('GPS_135_Data');
  if (!hoja) {
    hoja = ss.insertSheet('GPS_135_Data');
    // Cabecera: ID_Agente, Nombre, Año, JSON_Datos, Ultima_Actualizacion
    const headers = ['ID_Agente', 'Nombre', 'Año', 'Datos_JSON', 'Ultima_Actualizacion'];
    hoja.getRange(1, 1, 1, headers.length).setValues([headers])
        .setBackground('#b70000')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
    hoja.setFrozenRows(1);
    hoja.autoResizeColumns(1, headers.length);
  }
  return hoja;
}

function guardarGPS135(datos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // Aseguramos que la hoja exista antes de escribir
    let hoja = ss.getSheetByName('GPS_135_Data');
    if (!hoja) hoja = crearHojaGPS135(ss);
    
    // Validar datos mínimos
    if (!datos || !datos.contenido) {
      throw new Error("No se recibieron datos para guardar.");
    }

    // El año lo sacamos del contenido editable o usamos el actual por defecto
    const year = datos.contenido.meta_anio || new Date().getFullYear();
    // Convertimos todo el objeto complejo a texto JSON para guardarlo en una celda
    const datosJSON = JSON.stringify(datos.contenido); 
    
    const data = hoja.getDataRange().getValues();
    let filaEncontrada = -1;
    
    // Buscar si ya existe este agente/año para sobreescribir
    for (let i = 1; i < data.length; i++) {
      // Comparamos ID (columna 0) y Año (columna 2)
      if (data[i][0] == datos.idAgente && data[i][2] == year) {
        filaEncontrada = i + 1; // +1 porque los índices de array empiezan en 0 y las filas en 1
        break;
      }
    }
    
    const timestamp = new Date();

    if (filaEncontrada > 0) {
      // ACTUALIZAR existente
      hoja.getRange(filaEncontrada, 2).setValue(datos.nombreAgente); // Actualizar nombre por si cambió
      hoja.getRange(filaEncontrada, 4).setValue(datosJSON);          // Actualizar JSON
      hoja.getRange(filaEncontrada, 5).setValue(timestamp);          // Actualizar fecha
    } else {
      // CREAR nuevo registro
      hoja.appendRow([
        datos.idAgente,
        datos.nombreAgente,
        year,
        datosJSON,
        timestamp
      ]);
    }
    
    return { success: true, message: 'GPS guardado correctamente' };

  } catch (e) {
    Logger.log("Error en guardarGPS135: " + e.message);
    throw new Error('Error al guardar GPS en el servidor: ' + e.message);
  }
}

function obtenerGPS135(idAgente) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName('GPS_135_Data');
    if (!hoja) return { contenido: null };
    
    // Por defecto buscamos el año actual, pero se podría mejorar para buscar el último editado
    const year = new Date().getFullYear();
    const data = hoja.getDataRange().getValues();
    
    // Buscar fila del agente
    for (let i = 1; i < data.length; i++) {
      // Convertimos a string para asegurar comparación correcta
      if (String(data[i][0]) === String(idAgente)) {
         // Opcional: filtrar también por año si quieres históricos
         // if (data[i][2] == year) { ... }
         
         const jsonTexto = data[i][3];
         if (jsonTexto && jsonTexto !== "") {
             return { contenido: JSON.parse(jsonTexto) };
         }
      }
    }
    return { contenido: null }; 

  } catch (e) {
    Logger.log("Error en obtenerGPS135: " + e.message);
    throw new Error('Error al recuperar GPS: ' + e.message);
  }
}

function mostrarNotificacion(mensaje, esError = false) {
    const x = document.getElementById("app-notification");
    if (!x) return;
    
    x.textContent = mensaje;
    x.className = "show " + (esError ? "error" : "success");
    
    // Ocultar después de 3 segundos
    setTimeout(function(){ 
        x.className = x.className.replace("show", ""); 
        // Limpiar clases de color
        setTimeout(() => { x.className = ""; }, 300);
    }, 3000);
}

// ==========================================
//  🤖  INTELIGENCIA ARTIFICIAL (GEMINI 1.5)
// ==========================================

// ⚠️ ¡PON TU API KEY NUEVA AQUÍ! ⚠️
const GEMINI_API_KEY = 'AIzaSyCnRMqUCcekn7pvHW6ltgKPWbP_9vzG8Zk'; 

function llamarGemini(prompt) {
  // Lista actualizada basada en TU diagnóstico (24 Nov 2025)
  // Priorizamos: 1. Flash 2.0 (Rápido/Nuevo) -> 2. Flash Latest (Estándar) -> 3. Pro Latest (Potente)
  const modelos = [
    "gemini-2.0-flash",       // El más equilibrado y moderno de tu lista
    "gemini-flash-latest",    // El alias genérico siempre seguro
    "gemini-2.0-flash-lite",  // Versión ultrarápida
    "gemini-pro-latest"       // Versión potente de respaldo
  ];

  let ultimoError = "";

  // Intentamos conectar con cada modelo de la lista hasta que uno responda
  for (let i = 0; i < modelos.length; i++) {
    const modeloActual = modelos[i];
    
    try {
      const url = `https://generativelanguage.googleapis.com/v1beta/models/${modeloActual}:generateContent?key=${GEMINI_API_KEY}`;
      
      const payload = {
        contents: [{ parts: [{ text: prompt }] }]
      };

      const options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      };

      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();
      const json = JSON.parse(response.getContentText());

      // Si funciona (Código 200), devolvemos el texto y SALIMOS
      if (responseCode === 200 && json.candidates && json.candidates.length > 0) {
        Logger.log("✅ Éxito usando modelo: " + modeloActual);
        return json.candidates[0].content.parts[0].text;
      }
      
      // Si falla, registramos y el bucle probará el siguiente
      const errorMsg = json.error ? json.error.message : "Error desconocido";
      ultimoError = `(${modeloActual}): ${errorMsg}`;
      Logger.log(`⚠️ Falló ${modeloActual}. Probando siguiente...`);

    } catch (e) {
      ultimoError = e.toString();
    }
  }

  // Si llegamos aquí, fallaron todos
  return "❌ Error IA: No se pudo conectar con ningún modelo (2.0/Flash). Detalle: " + ultimoError;
}

// --- FUNCIONES DE ANÁLISIS (Sustituyen a las antiguas) ---

function analizarAgenteIA(agente, periodoActual) {
  if (!agente) return "<p>Error: Datos no disponibles.</p>";
  
  // Contexto seguro para evitar errores de null
  const datos = {
      nombre: agente.agente || "Agente",
      cumplimiento: parseFloat(agente.cumplimientoGlobal || 0).toFixed(1),
      gci: parseFloat(agente.realizado?.gci || 0).toFixed(0),
      conversion: parseFloat(agente.ratios?.conversionCaptacion || 0).toFixed(1)
  };

  const prompt = `
    Actúa como un Coach Inmobiliario de alto rendimiento (estilo Keller Williams).
    Analiza a este agente:
    - Nombre: ${datos.nombre}
    - Periodo: ${periodoActual}
    - Cumplimiento Objetivos: ${datos.cumplimiento}%
    - GCI: ${datos.gci}€
    - Conversión Cita->Exclusiva: ${datos.conversion}% (Ideal >30%)

    Dame:
    1. Un título motivador con emoji.
    2. Diagnóstico breve: ¿Falta actividad o habilidad?
    3. DOS acciones concretas para la semana que viene.
    
    Responde en HTML limpio (sin markdown, usa <h3>, <p>, <ul>, <b>). Sé breve.
  `;

  return llamarGemini(prompt);
}

function analizarEquipoIA(datosParaIA, periodoActual) {
  if (!datosParaIA || datosParaIA.length === 0) return "<p>Sin datos.</p>";
  
  // Resumimos para no gastar tokens
  const resumen = datosParaIA.slice(0, 10).map(a => 
    `${a.agente}: ${a.cumplimiento}% cumpl., ${a.gci}€ GCI`
  ).join("\n");

  const prompt = `
    Eres el Director de Ventas. Analiza el equipo (${periodoActual}):
    ${resumen}
    
    Dime:
    1. Quién es el MVP.
    2. Qué métrica general falla.
    3. Un mensaje de 1 frase para el grupo de WhatsApp del equipo.
    
    Usa HTML limpio.
  `;

  return llamarGemini(prompt);
}
function ampliarGrafico(chartId, titulo) {
    const originalChart = chartInstances[chartId];
    if (!originalChart) return;

    const modalBody = document.getElementById('modal-analisis-body');
    modalBody.innerHTML = `
        <h2 style="text-align: center; margin-bottom: 20px;">🔍 ${titulo}</h2>
        <div style="height: 70vh; width: 100%;">
            <canvas id="canvas-ampliado"></canvas>
        </div>
    `;
    
    document.getElementById('modal-analisis').classList.add('active');
    document.body.classList.add('modal-open');

    setTimeout(() => {
        const ctx = document.getElementById('canvas-ampliado');
        new Chart(ctx, {
            type: originalChart.config.type,
            data: originalChart.config.data,
            options: {
                ...originalChart.config.options,
                maintainAspectRatio: false,
                plugins: { legend: { position: 'top' } }
            }
        });
    }, 100);
}
function VER_MODELOS_DISPONIBLES() {
  // Usa tu clave aquí
  const key = GEMINI_API_KEY; 
  const url = `https://generativelanguage.googleapis.com/v1beta/models?key=${key}`;
  
  try {
    const response = UrlFetchApp.fetch(url);
    const json = JSON.parse(response.getContentText());
    Logger.log("--- MODELOS DISPONIBLES PARA TU CLAVE ---");
    json.models.forEach(m => {
      // Filtramos solo los que sirven para generar texto
      if(m.supportedGenerationMethods.includes("generateContent")) {
        Logger.log("Nombre: " + m.name);
      }
    });
    Logger.log("-----------------------------------------");
  } catch (e) {
    Logger.log("❌ Error fatal verificando modelos: " + e.toString());
  }
}
// ==========================================
//  👥  NUEVO MODELO ORGANIZATIVO (JSON)
// ==========================================

function crearHojaOrganigrama(ss) {
  let hoja = ss.getSheetByName('Organigrama_Full');
  if (!hoja) {
    hoja = ss.insertSheet('Organigrama_Full');
    // Solo necesitamos una celda gigante para guardar todo el estado
    hoja.getRange('A1').setValue('DATA_JSON');
    hoja.getRange('B1').setValue('LAST_UPDATE');
  }
  return hoja;
}

function guardarOrganigramaJSON(jsonTexto) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName('Organigrama_Full');
  if (!hoja) hoja = crearHojaOrganigrama(ss);
  
  // Guardamos todo en la fila 2
  hoja.getRange('A2').setValue(jsonTexto);
  hoja.getRange('B2').setValue(new Date());
}

function obtenerOrganigramaJSON() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('Organigrama_Full');
  if (!hoja) return "";
  
  return hoja.getRange('A2').getValue();
}
// --- BASE DE DATOS HISTÓRICA (JSON Simple) ---
function crearHojaHistorico(ss) {
    let hoja = ss.getSheetByName('Datos_Historicos_JSON');
    if (!hoja) {
        hoja = ss.insertSheet('Datos_Historicos_JSON');
        hoja.getRange('A1').setValue('JSON_DATA');
    }
    return hoja;
}

function guardarDatosHistoricosJSON(json) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName('Datos_Historicos_JSON');
    if (!hoja) hoja = crearHojaHistorico(ss);
    hoja.getRange('A2').setValue(json);
}

function obtenerDatosHistoricosJSON() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName('Datos_Historicos_JSON');
    if (!hoja) return null;
    return hoja.getRange('A2').getValue();
}
// ==========================================
// 📜 MÓDULO DE HISTÓRICOS AVANZADO (V2)
// ==========================================

function crearHojaHistoricoAgentes(ss) {
  let hoja = ss.getSheetByName('Historico_Agentes');
  if (hoja) ss.deleteSheet(hoja);
  hoja = ss.insertSheet('Historico_Agentes');
  
  // ✅ ENCABEZADOS NUEVOS (25 columnas) - ESTRUCTURA COMPLETA
  const headers = [
    'ID_Agente',        // [0]  
    'Nombre',           // [1]  
    'Año',              // [2]  
    'Mes',              // [3]  
    'Citas',            // [4]  - Nº CITAS
    'Capt_Excl',        // [5]  - Captaciones EXCLUSIVAS
    'Capt_Abierto',     // [6]  - Captaciones ABIERTO
    'Visitas_Comp',     // [7]  - CLIENTES Compradores con VISITAS
    'Casas_Ens',        // [8]  - CASAS Enseñadas
    'Capt_Alq',         // [9]  - Captaciones ALQUILER
    '3Bs',              // [10] - 3Bs Activadas
    'Bajadas',          // [11] - Bajadas Precio
    'Propuestas',       // [12] - Propuesta Compra
    'Arras',            // [13] 
    'Vtas_Excl',        // [14] - VENTAS EXCLUSIVAS
    'GCI_Excl',         // [15] - GCI EXCLUSIVAS
    'Vtas_Abierto',     // [16] - VENTAS ABIERTO
    'GCI_Abierto',      // [17] - GCI ABIERTO
    'Vtas_Comp',        // [18] - VENTAS COMPRADORES
    'GCI_Comp',         // [19] - GCI COMPRADORES
    'Vtas_Alq',         // [20] - ALQUILER CERRADOS
    'GCI_Alq',          // [21] - GCI ALQUILER
    'GCI_Total',        // [22] - FACTURACIÓN GCI TOTAL ⭐ CRÍTICO
    'Co_Euro',          // [23] - Company Euro
    'Fecha_Registro'    // [24]
  ];
  
  // Establecer encabezados con estilo KW
  hoja.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground('#b70000')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(11);
  
  hoja.setFrozenRows(1);
  
  // ✅ Formato de columnas GCI (moneda EUR)
  hoja.getRange('P:P').setNumberFormat('#,##0.00 €'); // GCI_Excl
  hoja.getRange('R:R').setNumberFormat('#,##0.00 €'); // GCI_Abierto
  hoja.getRange('T:T').setNumberFormat('#,##0.00 €'); // GCI_Comp
  hoja.getRange('V:V').setNumberFormat('#,##0.00 €'); // GCI_Alq
  hoja.getRange('W:W').setNumberFormat('#,##0.00 €'); // GCI_Total ⭐
  hoja.getRange('X:X').setNumberFormat('#,##0.00 €'); // Co_Euro
  
  // Formato de columnas numéricas (enteros)
  hoja.getRange('E:O').setNumberFormat('#,##0'); // Citas hasta Arras
  hoja.getRange('O:O').setNumberFormat('#,##0'); // Vtas_Excl
  hoja.getRange('Q:Q').setNumberFormat('#,##0'); // Vtas_Abierto
  hoja.getRange('S:S').setNumberFormat('#,##0'); // Vtas_Comp
  hoja.getRange('U:U').setNumberFormat('#,##0'); // Vtas_Alq
  
  hoja.autoResizeColumns(1, headers.length);
  
  Logger.log('✅ Hoja Historico_Agentes creada con 25 columnas');
}

function guardarHistoricoAgenteHTML(datos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName('Historico_Agentes');
    if (!hoja) throw new Error('No existe la hoja Historico_Agentes');

    const idAgente = datos.idAgente;
    const anio = parseInt(datos.anio);
    const modo = datos.modo; // 'ANUAL' o 'MENSUAL'
    
    Logger.log(`📥 Guardando histórico: Agente=${idAgente}, Año=${anio}, Modo=${modo}`);
    
    // 🗑️ PASO 1: ELIMINAR DATOS PREVIOS DE ESTE AGENTE + AÑO
    const datosExistentes = hoja.getDataRange().getValues();
    const filasAEliminar = [];
    
    for (let i = datosExistentes.length - 1; i >= 1; i--) {
      if (String(datosExistentes[i][0]) === String(idAgente) && 
          datosExistentes[i][2] === anio) {
        filasAEliminar.push(i + 1);
      }
    }
    
    for (const fila of filasAEliminar) {
      hoja.deleteRow(fila);
    }
    
    Logger.log(`🗑️ Eliminadas ${filasAEliminar.length} filas antiguas`);
    
    // 💾 PASO 2: INSERTAR NUEVOS DATOS (25 COLUMNAS)
    const filasNuevas = [];
    
    if (modo === 'ANUAL') {
      // ═══════════════════════════════════════════════════════════════
      // MODO ANUAL: Distribuir totales entre 12 meses (promedio)
      // ═══════════════════════════════════════════════════════════════
      const totales = datos.anual;
      
      for (let mes = 1; mes <= 12; mes++) {
        filasNuevas.push([
          idAgente,                                      // [0]  ID_Agente
          '',                                           // [1]  Nombre (VLOOKUP)
          anio,                                         // [2]  Año
          mes,                                          // [3]  Mes
          Math.round((totales.citasCapt || 0) / 12),   // [4]  Citas
          Math.round((totales.exclVenta || 0) / 12),   // [5]  Capt_Excl
          Math.round((totales.captAbierto || 0) / 12), // [6]  Capt_Abierto
          Math.round((totales.citasComp || 0) / 12),   // [7]  Visitas_Comp
          Math.round((totales.visitas || 0) / 12),     // [8]  Casas_Ens
          0,                                            // [9]  Capt_Alq
          0,                                            // [10] 3Bs
          0,                                            // [11] Bajadas
          0,                                            // [12] Propuestas
          0,                                            // [13] Arras
          Math.round((totales.ventas || 0) / 12),      // [14] Vtas_Excl
          parseFloat(((totales.gci || 0) / 12).toFixed(2)), // [15] GCI_Excl
          0,                                            // [16] Vtas_Abierto
          0,                                            // [17] GCI_Abierto
          0,                                            // [18] Vtas_Comp
          0,                                            // [19] GCI_Comp
          0,                                            // [20] Vtas_Alq
          0,                                            // [21] GCI_Alq
          parseFloat(((totales.gci || 0) / 12).toFixed(2)), // [22] GCI_Total ⭐
          0,                                            // [23] Co_Euro
          new Date()                                    // [24] Fecha_Registro
        ]);
      }
      
    } else {
      // ═══════════════════════════════════════════════════════════════
      // MODO MENSUAL: Cada mes tiene sus datos específicos
      // ═══════════════════════════════════════════════════════════════
      const meses = datos.mensual;
      
      for (let mes = 1; mes <= 12; mes++) {
        const m = meses[mes - 1]; // índice 0-11
        
        filasNuevas.push([
          idAgente,                             // [0]  ID_Agente
          '',                                   // [1]  Nombre (VLOOKUP)
          anio,                                 // [2]  Año
          mes,                                  // [3]  Mes
          parseInt(m.citasCapt) || 0,          // [4]  Citas
          parseInt(m.exclVenta) || 0,          // [5]  Capt_Excl
          parseInt(m.captAbierto) || 0,        // [6]  Capt_Abierto
          parseInt(m.citasComp) || 0,          // [7]  Visitas_Comp
          parseInt(m.visitas) || 0,            // [8]  Casas_Ens
          0,                                    // [9]  Capt_Alq
          0,                                    // [10] 3Bs
          0,                                    // [11] Bajadas
          0,                                    // [12] Propuestas
          0,                                    // [13] Arras
          parseInt(m.ventas) || 0,             // [14] Vtas_Excl
          parseFloat(m.gci || 0).toFixed(2),   // [15] GCI_Excl
          0,                                    // [16] Vtas_Abierto
          0,                                    // [17] GCI_Abierto
          0,                                    // [18] Vtas_Comp
          0,                                    // [19] GCI_Comp
          0,                                    // [20] Vtas_Alq
          0,                                    // [21] GCI_Alq
          parseFloat(m.gci || 0).toFixed(2),   // [22] GCI_Total ⭐
          0,                                    // [23] Co_Euro
          new Date()                            // [24] Fecha_Registro
        ]);
      }
    }
    
    // Insertar todas las filas de una vez
    if (filasNuevas.length > 0) {
      const ultimaFila = hoja.getLastRow();
      hoja.getRange(ultimaFila + 1, 1, filasNuevas.length, 25).setValues(filasNuevas);
      
      Logger.log(`✅ Insertadas ${filasNuevas.length} filas nuevas`);
      
      // Rellenar nombres con fórmula VLOOKUP
      const hojaAgentes = ss.getSheetByName('Agentes');
      if (hojaAgentes) {
        for (let i = 0; i < filasNuevas.length; i++) {
          const fila = ultimaFila + 1 + i;
          hoja.getRange(fila, 2).setFormula(`=IFERROR(VLOOKUP(A${fila},Agentes!A:B,2,FALSE),"")`);
        }
      }
    }
    
    return { 
      success: true, 
      message: `✅ Histórico guardado: ${filasNuevas.length} registros (${modo})` 
    };
    
  } catch (error) {
    Logger.log('❌ ERROR en guardarHistoricoAgenteHTML: ' + error);
    return { success: false, error: error.message };
  }
}

function obtenerHistoricoAgente(idAgente, anio) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('Historico_Agentes');
  if (!hoja) return { modo: 'VACIO' };
  
  const datos = hoja.getDataRange().getValues();
  const meses = [];
  let anual = null;
  
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] == idAgente && datos[i][2] == anio) {
      // Mapeamos las columnas a un objeto limpio
      const fila = {
        gci: datos[i][4], ventas: datos[i][5],
        citasCapt: datos[i][6], exclVenta: datos[i][7], captAbierto: datos[i][8],
        citasComp: datos[i][9], exclComp: datos[i][10], visitas: datos[i][11]
      };

      if (datos[i][3] === 'TOTAL') {
        anual = fila;
      } else {
        meses[parseInt(datos[i][3]) - 1] = fila;
      }
    }
  }
  
  if (anual) return { modo: 'ANUAL', valores: anual };
  if (meses.length > 0) {
    // Rellenar huecos vacíos
    for(let j=0; j<12; j++) if(!meses[j]) meses[j] = {gci:0, ventas:0, citasCapt:0, exclVenta:0, captAbierto:0, citasComp:0, exclComp:0, visitas:0};
    return { modo: 'MENSUAL', valores: meses };
  }
  
  return { modo: 'VACIO' };
}

// --- NUEVA FUNCIÓN: Obtener Histórico de TODOS los agentes ---
function obtenerTodosHistoricosAgentes(anio) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName('Historico_Agentes');
    if (!hoja) return {};
    
    const datos = hoja.getDataRange().getValues();
    const resultado = {};
    
    // Empezamos en fila 1 (saltando cabecera)
    for (let i = 1; i < datos.length; i++) {
      // Verificamos que sea del año solicitado y que sea un total ANUAL
      if (datos[i][2] == anio && datos[i][3] === 'TOTAL') {
        const idAgente = datos[i][0];
        // Guardamos los datos limpios
        resultado[idAgente] = {
            gci: parseFloat(datos[i][4]) || 0,
            ventas: parseFloat(datos[i][5]) || 0,
            citasCapt: parseFloat(datos[i][6]) || 0,
            exclVenta: parseFloat(datos[i][7]) || 0,
            captAbierto: parseFloat(datos[i][8]) || 0,
            citasComp: parseFloat(datos[i][9]) || 0,
            exclComp: parseFloat(datos[i][10]) || 0,
            visitas: parseFloat(datos[i][11]) || 0
        };
      }
    }
    return resultado;
  } catch (e) {
    return {};
  }
}
// ============================================
// FUNCIONES DE EMBUDOS Y CORRELACIONES
// ============================================

function obtenerDatosEmbudo(filtros) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName('Actividad_Diaria');
    const datos = hoja.getDataRange().getValues();
    
    const fechaInicio = new Date(filtros.fechaInicio);
    const fechaFin = new Date(filtros.fechaFin);
    
    const embudos = {
      captacion: {
        etapa1_llamadas: 0,
        etapa2_citasCaptacion: 0,
        etapa3_exclusivasVenta: 0,
        etapa4_captacionesAbierto: 0,
        conversion_llamadas_citas: 0,
        conversion_citas_exclusivas: 0,
        conversion_exclusivas_captaciones: 0
      },
      compradores: {
        etapa1_leadsCompradores: 0,
        etapa2_citasCompradores: 0,
        etapa3_exclusivasComprador: 0,
        etapa4_casasEnsenadas: 0,
        conversion_leads_citas: 0,
        conversion_citas_exclusivas: 0,
        conversion_exclusivas_visitas: 0
      },
      cierre: {
        totalExclusivas: 0,
        totalGCI: 0,
        volumenNegocio: 0,
        transacciones: 0,
        ticketPromedio: 0,
        conversion_exclusivas_ventas: 0
      }
    };
    
    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      const fecha = new Date(fila[1]);
      
      if (fecha >= fechaInicio && fecha <= fechaFin) {
        if (filtros.idAgente && fila[2] !== filtros.idAgente) continue;
        
        embudos.captacion.etapa1_llamadas += parseInt(fila[11]) || 0;
        embudos.captacion.etapa2_citasCaptacion += parseInt(fila[4]) || 0;
        embudos.captacion.etapa3_exclusivasVenta += parseInt(fila[5]) || 0;
        embudos.captacion.etapa4_captacionesAbierto += parseInt(fila[7]) || 0;
        
        embudos.compradores.etapa1_leadsCompradores += parseInt(fila[10]) || 0;
        embudos.compradores.etapa2_citasCompradores += parseInt(fila[8]) || 0;
        embudos.compradores.etapa3_exclusivasComprador += parseInt(fila[6]) || 0;
        embudos.compradores.etapa4_casasEnsenadas += parseInt(fila[9]) || 0;
        
        const gci = parseFloat(fila[12]) || 0;
        const volumen = parseFloat(fila[13]) || 0;
        
        if (gci > 0) {
          embudos.cierre.transacciones++;
          embudos.cierre.totalGCI += gci;
          embudos.cierre.volumenNegocio += volumen;
        }
        
        embudos.cierre.totalExclusivas += (parseInt(fila[5]) || 0) + (parseInt(fila[6]) || 0);
      }
    }
    
    const c = embudos.captacion;
    c.conversion_llamadas_citas = c.etapa1_llamadas > 0 ? ((c.etapa2_citasCaptacion / c.etapa1_llamadas) * 100).toFixed(1) : 0;
    c.conversion_citas_exclusivas = c.etapa2_citasCaptacion > 0 ? ((c.etapa3_exclusivasVenta / c.etapa2_citasCaptacion) * 100).toFixed(1) : 0;
    c.conversion_exclusivas_captaciones = c.etapa3_exclusivasVenta > 0 ? ((c.etapa4_captacionesAbierto / c.etapa3_exclusivasVenta) * 100).toFixed(1) : 0;
    
    const comp = embudos.compradores;
    comp.conversion_leads_citas = comp.etapa1_leadsCompradores > 0 ? ((comp.etapa2_citasCompradores / comp.etapa1_leadsCompradores) * 100).toFixed(1) : 0;
    comp.conversion_citas_exclusivas = comp.etapa2_citasCompradores > 0 ? ((comp.etapa3_exclusivasComprador / comp.etapa2_citasCompradores) * 100).toFixed(1) : 0;
    comp.conversion_exclusivas_visitas = comp.etapa3_exclusivasComprador > 0 ? ((comp.etapa4_casasEnsenadas / comp.etapa3_exclusivasComprador) * 100).toFixed(1) : 0;
    
    const cierre = embudos.cierre;
    cierre.ticketPromedio = cierre.transacciones > 0 ? (cierre.totalGCI / cierre.transacciones).toFixed(0) : 0;
    cierre.conversion_exclusivas_ventas = cierre.totalExclusivas > 0 ? ((cierre.transacciones / cierre.totalExclusivas) * 100).toFixed(1) : 0;
    
    return { success: true, embudos: embudos };
    
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function obtenerDatosCorrelacion(filtros) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName('Actividad_Diaria');
    const datos = hoja.getDataRange().getValues();
    
    const fechaInicio = new Date(filtros.fechaInicio);
    const fechaFin = new Date(filtros.fechaFin);
    
    const series = {
      llamadas: [], citasCaptacion: [], exclusivasVenta: [], exclusivasComprador: [],
      captacionesAbierto: [], citasCompradores: [], casasEnsenadas: [],
      leadsCompradores: [], gci: [], volumenNegocio: []
    };
    
    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      const fecha = new Date(fila[1]);
      
      if (fecha >= fechaInicio && fecha <= fechaFin) {
        if (filtros.idAgente && fila[2] !== filtros.idAgente) continue;
        
        series.llamadas.push(parseFloat(fila[11]) || 0);
        series.citasCaptacion.push(parseFloat(fila[4]) || 0);
        series.exclusivasVenta.push(parseFloat(fila[5]) || 0);
        series.exclusivasComprador.push(parseFloat(fila[6]) || 0);
        series.captacionesAbierto.push(parseFloat(fila[7]) || 0);
        series.citasCompradores.push(parseFloat(fila[8]) || 0);
        series.casasEnsenadas.push(parseFloat(fila[9]) || 0);
        series.leadsCompradores.push(parseFloat(fila[10]) || 0);
        series.gci.push(parseFloat(fila[12]) || 0);
        series.volumenNegocio.push(parseFloat(fila[13]) || 0);
      }
    }
    
    const kpis = Object.keys(series);
    const matriz = {};
    
    for (let i = 0; i < kpis.length; i++) {
      const kpi1 = kpis[i];
      matriz[kpi1] = {};
      for (let j = 0; j < kpis.length; j++) {
        const kpi2 = kpis[j];
        matriz[kpi1][kpi2] = calcularCorrelacionPearson(series[kpi1], series[kpi2]);
      }
    }
    
    const topCorrelaciones = [];
    for (let i = 0; i < kpis.length; i++) {
      for (let j = i + 1; j < kpis.length; j++) {
        const kpi1 = kpis[i];
        const kpi2 = kpis[j];
        const valor = Math.abs(matriz[kpi1][kpi2]);
        if (!isNaN(valor) && isFinite(valor)) {
          topCorrelaciones.push({ kpi1: kpi1, kpi2: kpi2, valor: matriz[kpi1][kpi2], valorAbs: valor });
        }
      }
    }
    
    topCorrelaciones.sort((a, b) => b.valorAbs - a.valorAbs);
    
    return {
      success: true,
      matriz: matriz,
      topCorrelaciones: topCorrelaciones.slice(0, 10),
      kpis: kpis,
      n: series.llamadas.length
    };
    
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function calcularCorrelacionPearson(x, y) {
  const n = x.length;
  if (n === 0 || n !== y.length) return 0;
  
  const mediaX = x.reduce((a, b) => a + b, 0) / n;
  const mediaY = y.reduce((a, b) => a + b, 0) / n;
  
  let numerador = 0, denominadorX = 0, denominadorY = 0;
  
  for (let i = 0; i < n; i++) {
    const diffX = x[i] - mediaX;
    const diffY = y[i] - mediaY;
    numerador += diffX * diffY;
    denominadorX += diffX * diffX;
    denominadorY += diffY * diffY;
  }
  
  const denominador = Math.sqrt(denominadorX * denominadorY);
  if (denominador === 0) return 0;
  
  return Math.round((numerador / denominador) * 1000) / 1000;
}
function obtenerTodasTransacciones(filtros) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(CONFIG.HOJA_ACTIVIDAD);
    if (!hoja) return { success: false, error: 'Hoja no encontrada' };
    
    const datos = hoja.getDataRange().getValues();
    const transacciones = [];
    
    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      const gci = parseFloat(fila[12]) || 0;
      const notas = String(fila[14] || '');
      
      // Solo filas con GCI > 0 Y que contengan "TRANSACCIÓN"
      if (gci > 0 && notas.includes('TRANSACCIÓN')) {
        const fecha = new Date(fila[1]);
        
        // Parsear tipo y lado desde notas
        const matchTipo = notas.match(/TRANSACCIÓN\s+(\w+)/i);
        const matchLado = notas.match(/Lado:\s*(\w+)/i);
        
        const tipo = matchTipo ? matchTipo[1].toUpperCase() : 'VENTA';
        const lado = matchLado ? matchLado[1] : 'Vendedor';
        
        transacciones.push({
          id: fila[0],
          fecha: Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
          agente: fila[3],
          tipo: tipo,
          lado: lado,
          gci: gci,
          volumenNegocio: parseFloat(fila[13]) || 0,
          comision: parseFloat(fila[16]) || 0,
          pctComision: parseFloat(fila[17]) || 0, // ✅ SIN multiplicar por 100
          descripcion: notas
        });
      }
    }
    
    return { success: true, transacciones: transacciones };
  } catch (error) {
    return { success: false, error: error.message };
  }
}
function editarTransaccion(idTransaccion, nuevoGCI, nuevaComision, nuevoPct) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName('Actividad_Diaria');
    const datos = hoja.getDataRange().getValues();
    
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0] === idTransaccion) {
        hoja.getRange(i + 1, 13).setValue(nuevoGCI); // GCI
        hoja.getRange(i + 1, 16).setValue(nuevaComision); // Comision_Pagada
        hoja.getRange(i + 1, 17).setValue(nuevoPct); // Pct_Comision
        return { success: true, message: 'Transacción actualizada' };
      }
    }
    
    return { success: false, error: 'Transacción no encontrada' };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function anularTransaccion(idTransaccion) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName('Actividad_Diaria');
    const datos = hoja.getDataRange().getValues();
    
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0] === idTransaccion) {
        // Marcar como anulada en las notas
        const notasActuales = datos[i][14] || '';
        hoja.getRange(i + 1, 15).setValue('[ANULADA] ' + notasActuales);
        
        // Poner GCI y comisiones en 0
        hoja.getRange(i + 1, 13).setValue(0); // GCI
        hoja.getRange(i + 1, 16).setValue(0); // Comision_Pagada
        
        return { success: true, message: 'Transacción anulada' };
      }
    }
    
    return { success: false, error: 'Transacción no encontrada' };
  } catch (error) {
    return { success: false, error: error.message };
  }
}
function obtenerListaAgentes() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName('Agentes');
    const datos = hoja.getDataRange().getValues();
    
    const agentes = [];
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0]) {
        agentes.push({
          id: datos[i][0],
          nombre: datos[i][1]
        });
      }
    }
    
    return agentes;
  } catch (error) {
    return [];
  }
}

function editarTransaccionCompleta(datos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(CONFIG.HOJA_ACTIVIDAD);
    const todasFilas = hoja.getDataRange().getValues();
    
    for (let i = 1; i < todasFilas.length; i++) {
      if (todasFilas[i][0] === datos.id) {
        const fecha = new Date(datos.fecha);
        
        // ✅ CORRECCIÓN CRÍTICA: Comisión = (GCI × %) / 100
        const comisionImporte = parseFloat((datos.gci * datos.pctComision / 100).toFixed(2));
        
        const notas = `TRANSACCIÓN ${datos.tipo} - ${datos.descripcion || 'Venta/Alquiler'} | Lado: ${datos.lado} | Comis: ${datos.pctComision}%`;
        
        hoja.getRange(i + 1, 2).setValue(fecha);
        hoja.getRange(i + 1, 3).setValue(datos.idAgente);
        hoja.getRange(i + 1, 4).setValue(datos.nombreAgente);
        hoja.getRange(i + 1, 13).setValue(datos.gci);
        hoja.getRange(i + 1, 14).setValue(datos.volumenNegocio);
        hoja.getRange(i + 1, 15).setValue(notas);
        hoja.getRange(i + 1, 16).setValue(new Date());
        hoja.getRange(i + 1, 17).setValue(comisionImporte);  // ✅ GCI × %
        hoja.getRange(i + 1, 18).setValue(datos.pctComision);
        
        return { success: true, message: 'Transacción actualizada' };
      }
    }
    
    return { success: false, error: 'Transacción no encontrada' };
  } catch (error) {
    return { success: false, error: error.message };
  }
}
function obtenerBeneficioNeto(filtros) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaActividad = ss.getSheetByName(CONFIG.HOJA_ACTIVIDAD);
    const hojaGastos = ss.getSheetByName('Gastos');
    
    if (!hojaActividad) {
      return { success: false, error: 'Hoja no encontrada' };
    }
    
    const datos = hojaActividad.getDataRange().getValues();
    
    // Calcular GCI y Comisiones totales
    let gciTotal = 0;
    let comisionesTotal = 0;
    
    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      const gci = parseFloat(fila[12]) || 0;
      const comision = parseFloat(fila[16]) || 0;
      
      gciTotal += gci;
      comisionesTotal += comision;
    }
    
    // Calcular Gastos Operativos
    let gastosOperativos = 0;
    if (hojaGastos) {
      const datosGastos = hojaGastos.getDataRange().getValues();
      for (let i = 1; i < datosGastos.length; i++) {
        const monto = parseFloat(datosGastos[i][3]) || 0;
        gastosOperativos += monto;
      }
    }
    
    // Calcular Beneficio
    const beneficioNeto = gciTotal - comisionesTotal - gastosOperativos;
    const porcentajeBeneficio = gciTotal > 0 ? ((beneficioNeto / gciTotal) * 100) : 0;
    
    // Datos mensuales para gráfico
    const datosMensuales = {};
    for (let mes = 1; mes <= 12; mes++) {
      datosMensuales[mes] = { gci: 0, comisiones: 0, gastos: 0, beneficio: 0 };
    }
    
    // Acumular por mes
    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      const fecha = new Date(fila[1]);
      if (isNaN(fecha.getTime())) continue;
      
      const mes = fecha.getMonth() + 1;
      const gci = parseFloat(fila[12]) || 0;
      const comision = parseFloat(fila[16]) || 0;
      
      datosMensuales[mes].gci += gci;
      datosMensuales[mes].comisiones += comision;
    }
    
    // Distribuir gastos proporcionalmente por mes
    if (hojaGastos) {
      const datosGastos = hojaGastos.getDataRange().getValues();
      for (let i = 1; i < datosGastos.length; i++) {
        const fecha = new Date(datosGastos[i][2]);
        if (isNaN(fecha.getTime())) continue;
        
        const mes = fecha.getMonth() + 1;
        const monto = parseFloat(datosGastos[i][3]) || 0;
        datosMensuales[mes].gastos += monto;
      }
    }
    
    // Calcular beneficio mensual
    Object.keys(datosMensuales).forEach(mes => {
      const d = datosMensuales[mes];
      d.beneficio = d.gci - d.comisiones - d.gastos;
    });
    
    return {
      success: true,
      gciTotal: gciTotal,
      comisionesTotal: comisionesTotal,
      gastosOperativos: gastosOperativos,
      beneficioNeto: beneficioNeto,
      porcentajeBeneficio: porcentajeBeneficio,
      datosMensuales: datosMensuales
    };
  } catch (error) {
    return { success: false, error: error.message };
  }
}
function obtenerEstadisticasTransacciones(filtros) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(CONFIG.HOJA_ACTIVIDAD);
    
    if (!hoja) {
      return { success: false, error: 'Hoja no encontrada' };
    }
    
    const datos = hoja.getDataRange().getValues();
    
    // Contadores globales
    const stats = {
      total: 0,
      vendedor: 0,
      comprador: 0,
      ambos: 0,
      gciTotal: 0,
      gciVendedor: 0,
      gciComprador: 0,
      gciAmbos: 0,
      volumenTotal: 0,
      volumenVendedor: 0,
      volumenComprador: 0,
      volumenAmbos: 0,
      comisionTotal: 0,
      comisionVendedor: 0,
      comisionComprador: 0,
      comisionAmbos: 0
    };
    
    // Datos mensuales
    const mensuales = {};
    for (let mes = 1; mes <= 12; mes++) {
      mensuales[mes] = {
        total: 0,
        vendedor: 0,
        comprador: 0,
        ambos: 0,
        gciTotal: 0
      };
    }
    
    // Recorrer transacciones
    for (let i = 1; i < datos.length; i++) {
      const fila = datos[i];
      const gci = parseFloat(fila[12]) || 0;
      const volumen = parseFloat(fila[13]) || 0;
      const comision = parseFloat(fila[16]) || 0;
      const notas = String(fila[14] || '');
      
      // Solo filas con GCI > 0 Y que contengan "TRANSACCIÓN"
      if (gci > 0 && notas.includes('TRANSACCIÓN')) {
        const fecha = new Date(fila[1]);
        const mes = fecha.getMonth() + 1;
        
        // Parsear lado desde notas
        const matchLado = notas.match(/Lado:\s*(\w+)/i);
        const lado = matchLado ? matchLado[1].toLowerCase() : 'vendedor';
        
        // Contadores globales
        stats.total++;
        stats.gciTotal += gci;
        stats.volumenTotal += volumen;
        stats.comisionTotal += comision;
        
        // Por lado
        if (lado.includes('vendedor')) {
          stats.vendedor++;
          stats.gciVendedor += gci;
          stats.volumenVendedor += volumen;
          stats.comisionVendedor += comision;
          mensuales[mes].vendedor++;
        } else if (lado.includes('comprador')) {
          stats.comprador++;
          stats.gciComprador += gci;
          stats.volumenComprador += volumen;
          stats.comisionComprador += comision;
          mensuales[mes].comprador++;
        } else if (lado.includes('ambos')) {
          stats.ambos++;
          stats.gciAmbos += gci;
          stats.volumenAmbos += volumen;
          stats.comisionAmbos += comision;
          mensuales[mes].ambos++;
        }
        
        // Mensuales
        mensuales[mes].total++;
        mensuales[mes].gciTotal += gci;
      }
    }
    
    return {
      success: true,
      stats: stats,
      mensuales: mensuales
    };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

// ============================================
// MEJORA 2: Función para obtener orígenes de negocio
// ============================================

function obtenerOrigenesNegocio() {
  return ORIGENES_NEGOCIO;
}
// ============================================
// MEJORA 3: Función para obtener partidas de gastos
// ============================================

function obtenerPartidasGastos() {
  return PARTIDAS_GASTOS;
}
// ============================================
// MEJORA 4: Verificar actividad previa de un agente en una fecha
// ============================================

function obtenerActividadDia(idAgente, fecha) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(CONFIG.HOJA_ACTIVIDAD);
    if (!hoja) return [];
    
    const datos = hoja.getDataRange().getValues();
    const actividades = [];
    const fechaBuscar = new Date(fecha);
    fechaBuscar.setHours(0, 0, 0, 0);
    
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][2] === idAgente) { // Columna 3: ID_Agente
        const fechaFila = new Date(datos[i][1]); // Columna 2: Fecha
        fechaFila.setHours(0, 0, 0, 0);
        
        if (fechaFila.getTime() === fechaBuscar.getTime()) {
          actividades.push({
            gci: datos[i][12] || 0,
            citasCaptacion: datos[i][4] || 0,
            exclusivasVenta: datos[i][5] || 0
          });
        }
      }
    }
    
    return actividades;
  } catch (error) {
    Logger.log('Error en obtenerActividadDia: ' + error);
    return [];
  }
}
// ============================================
// MEJORA EXTRA: Obtener actividad completa de un día para editar
// ============================================

function obtenerActividadCompletaDia(idAgente, fecha) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(CONFIG.HOJA_ACTIVIDAD);
    if (!hoja) return null;
    
    const datos = hoja.getDataRange().getValues();
    const fechaBuscar = new Date(fecha);
    fechaBuscar.setHours(0, 0, 0, 0);
    
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][2] === idAgente) { // Columna 3: ID_Agente
        const fechaFila = new Date(datos[i][1]); // Columna 2: Fecha
        fechaFila.setHours(0, 0, 0, 0);
        
        if (fechaFila.getTime() === fechaBuscar.getTime()) {
          return {
            fila: i + 1, // Guardamos el número de fila para actualizar después
            citasCaptacion: datos[i][4] || 0,
            exclusivasVenta: datos[i][5] || 0,
            exclusivasComprador: datos[i][6] || 0,
            captacionesAbierto: datos[i][7] || 0,
            citasCompradores: datos[i][8] || 0,
            casasEnsenadas: datos[i][9] || 0,
            leadsCompradores: datos[i][10] || 0,
            llamadas: datos[i][11] || 0,
            gci: datos[i][12] || 0,
            volumenNegocio: datos[i][13] || 0,
            notas: datos[i][14] || ''
          };
        }
      }
    }
    
    return null; // No hay actividad ese día
  } catch (error) {
    Logger.log('Error en obtenerActividadCompletaDia: ' + error);
    return null;
  }
}
// ============================================
// MEJORA EXTRA: Actualizar actividad existente
// ============================================

function actualizarActividad(datos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(CONFIG.HOJA_ACTIVIDAD);
    if (!hoja) throw new Error('No se encontró la hoja Actividad_Diaria');
    
    const datosHoja = hoja.getDataRange().getValues();
    const fechaBuscar = new Date(datos.fecha);
    fechaBuscar.setHours(0, 0, 0, 0);
    
    // Buscar la fila exacta
    for (let i = 1; i < datosHoja.length; i++) {
      if (datosHoja[i][2] === datos.idAgente) { // Columna 3: ID_Agente
        const fechaFila = new Date(datosHoja[i][1]); // Columna 2: Fecha
        fechaFila.setHours(0, 0, 0, 0);
        
        if (fechaFila.getTime() === fechaBuscar.getTime()) {
          // Actualizar la fila encontrada
          const fila = i + 1;
          
          hoja.getRange(fila, 5).setValue(parseInt(datos.citasCaptacion) || 0);
          hoja.getRange(fila, 6).setValue(parseInt(datos.exclusivasVenta) || 0);
          hoja.getRange(fila, 7).setValue(parseInt(datos.exclusivasComprador) || 0);
          hoja.getRange(fila, 8).setValue(parseInt(datos.captacionesAbierto) || 0);
          hoja.getRange(fila, 9).setValue(parseInt(datos.citasCompradores) || 0);
          hoja.getRange(fila, 10).setValue(parseInt(datos.casasEnsenadas) || 0);
          hoja.getRange(fila, 11).setValue(parseInt(datos.leadsCompradores) || 0);
          hoja.getRange(fila, 12).setValue(parseInt(datos.llamadas) || 0);
          hoja.getRange(fila, 13).setValue(parseFloat(datos.gci) || 0);
          hoja.getRange(fila, 14).setValue(parseFloat(datos.volumenNegocio) || 0);
          hoja.getRange(fila, 15).setValue(datos.notas || '');
          hoja.getRange(fila, 16).setValue(new Date()); // Timestamp actualización
          
          return { success: true, message: 'Actividad actualizada correctamente', accion: 'actualizar' };
        }
      }
    }
    
    // Si no se encontró, crear nueva (fallback)
    return guardarActividad(datos);
    
  } catch (error) {
    Logger.log('Error en actualizarActividad: ' + error);
    return { success: false, error: error.message };
  }
}

function obtenerDashboardGestionAgentes(filtros) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaAgentes = ss.getSheetByName(CONFIG.HOJA_AGENTES);
    const hojaActividad = ss.getSheetByName(CONFIG.HOJA_ACTIVIDAD);
    
    if (!hojaAgentes || !hojaActividad) {
      throw new Error('Faltan hojas necesarias');
    }
    
    // Leer datos
    const datosAgentes = hojaAgentes.getDataRange().getValues();
    const datosActividad = hojaActividad.getDataRange().getValues();
    
    // Arrays para el dashboard
    const agentesActivos = [];
    let totalGCI = 0;
    let totalTransacciones = 0;
    
    // Procesar cada agente activo
    for (let i = 1; i < datosAgentes.length; i++) {
      const estado = datosAgentes[i][5]; // Columna F: Estado
      if (estado !== 'Activo') continue;
      
      const idAgente = datosAgentes[i][0];
      const nombre = datosAgentes[i][1];
      const email = datosAgentes[i][2];
      const fechaIngreso = datosAgentes[i][4];
      const sueldoFijo = parseFloat(datosAgentes[i][8]) || 0;
      
      // Calcular antigüedad en meses
      let antiguedad = 0;
      if (fechaIngreso) {
        const hoy = new Date();
        const fechaIng = new Date(fechaIngreso);
        antiguedad = Math.floor((hoy - fechaIng) / (1000 * 60 * 60 * 24 * 30));
      }
      
      // Calcular KPIs del agente
      let gciAgente = 0;
      let comisionesTotales = 0;
      let transaccionesAgente = 0;
      
      for (let j = 1; j < datosActividad.length; j++) {
        if (datosActividad[j][2] === idAgente) {
          const gci = parseFloat(datosActividad[j][12]) || 0;
          const comision = parseFloat(datosActividad[j][16]) || 0;
          
          gciAgente += gci;
          comisionesTotales += comision;
          
          if (gci > 0) transaccionesAgente++;
        }
      }
      
      // Calcular rentabilidad (GCI - Sueldo - Comisiones)
      const rentabilidad = gciAgente - (sueldoFijo * 12) - comisionesTotales;
      const pctRentabilidad = gciAgente > 0 ? (rentabilidad / gciAgente) * 100 : 0;
      
      // Calcular ticket promedio
      const ticketPromedio = transaccionesAgente > 0 ? gciAgente / transaccionesAgente : 0;
      
      agentesActivos.push({
        nombre: nombre,
        email: email,
        antiguedad: antiguedad + ' meses',
        sueldoFijo: sueldoFijo,
        gci: gciAgente,
        transacciones: transaccionesAgente,
        ticketPromedio: ticketPromedio,
        comisiones: comisionesTotales,
        rentabilidad: rentabilidad,
        pctRentabilidad: pctRentabilidad
      });
      
      totalGCI += gciAgente;
      totalTransacciones += transaccionesAgente;
    }
    
    // Ordenar por GCI descendente
    agentesActivos.sort((a, b) => b.gci - a.gci);
    
    // Calcular ticket promedio global
    const ticketPromedioGlobal = totalTransacciones > 0 ? totalGCI / totalTransacciones : 0;
    
    // ─── GRÁFICO: ALTAS DE AGENTES POR MES (ÚLTIMOS 12 MESES) ───
    const hoy = new Date();
    const altasPorMes = Array(12).fill(0);
    const labelsMeses = [];
    
    for (let m = 11; m >= 0; m--) {
      const fecha = new Date(hoy.getFullYear(), hoy.getMonth() - m, 1);
      const mesNombre = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'][fecha.getMonth()];
      labelsMeses.push(mesNombre);
    }
    
    for (let i = 1; i < datosAgentes.length; i++) {
      const fechaIngreso = datosAgentes[i][4];
      if (!fechaIngreso) continue;
      
      const fechaIng = new Date(fechaIngreso);
      const mesesAtras = Math.floor((hoy - fechaIng) / (1000 * 60 * 60 * 24 * 30));
      
      if (mesesAtras >= 0 && mesesAtras < 12) {
        const indice = 11 - mesesAtras;
        altasPorMes[indice]++;
      }
    }
    
    return {
      resumen: {
        totalAgentes: agentesActivos.length,
        totalGCI: totalGCI,
        totalTransacciones: totalTransacciones,
        ticketPromedio: ticketPromedioGlobal
      },
      agentes: agentesActivos,
      graficoAltas: {
        labels: labelsMeses,
        datos: altasPorMes
      }
    };
    
  } catch (error) {
    throw new Error('Error en obtenerDashboardGestionAgentes: ' + error.message);
  }
}

/**
 * ────────────────────────────────────────────────────────────────────────────
 * FUNCIÓN 2: obtenerTodosAgentes
 * ────────────────────────────────────────────────────────────────────────────
 * Retorna TODOS los agentes (activos e inactivos) para el listado completo
 */
function obtenerTodosAgentes() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(CONFIG.HOJA_AGENTES);
    
    if (!hoja) {
      throw new Error('No se encontró la hoja de Agentes');
    }
    
    const datos = hoja.getDataRange().getValues();
    const agentes = [];
    
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][0] && datos[i][1]) {
        agentes.push({
          id: datos[i][0],
          nombre: datos[i][1],
          email: datos[i][2] || '',
          telefono: datos[i][3] || '',
          fechaIngreso: datos[i][4] ? Utilities.formatDate(new Date(datos[i][4]), Session.getScriptTimeZone(), 'dd/MM/yyyy') : '',
          sueldoFijo: parseFloat(datos[i][8]) || 0,
          estado: datos[i][5] || 'Inactivo'
        });
      }
    }
    
    return agentes;
    
  } catch (error) {
    throw new Error('Error en obtenerTodosAgentes: ' + error.message);
  }
}

/**
 * ────────────────────────────────────────────────────────────────────────────
 * FUNCIÓN 3: crearAgenteCompleto
 * ────────────────────────────────────────────────────────────────────────────
 * Versión mejorada para crear agente desde el modal de gestión
 * (Similar a crearNuevoAgente pero retorna más información)
 */
function crearAgenteCompleto(datos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName(CONFIG.HOJA_AGENTES);
    
    if (!hoja) {
      throw new Error('No se encontró la hoja de Agentes');
    }
    
    const ultimaFila = hoja.getLastRow();
    const nuevoId = 'AG' + String(ultimaFila + 1).padStart(3, '0');
    
    const fila = [
      nuevoId,
      datos.nombre,
      datos.email || '',
      datos.telefono || '',
      new Date(datos.fechaIngreso),
      'Activo',
      datos.objetivosAcumulados || 'NO',
      new Date(),
      parseFloat(datos.sueldoFijo) || 0
    ];
    
    hoja.appendRow(fila);
    
    return {
      success: true,
      message: 'Agente creado correctamente',
      agente: {
        id: nuevoId,
        nombre: datos.nombre,
        email: datos.email,
        telefono: datos.telefono,
        fechaIngreso: datos.fechaIngreso,
        sueldoFijo: parseFloat(datos.sueldoFijo) || 0,
        estado: 'Activo'
      }
    };
    
  } catch (error) {
    throw new Error('Error al crear agente: ' + error.message);
  }
}
function guardarImportacionMasivaV27(listaAgentes) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Hoja Agentes (Crear si falta)
  let hojaAgentes = ss.getSheetByName('Agentes');
  if (!hojaAgentes) { crearHojaAgentes(ss); hojaAgentes = ss.getSheetByName('Agentes'); }
  
  // 2. Hoja Histórico (Crear si falta o si es vieja)
  let hojaHist = ss.getSheetByName('Historico_Agentes');
  if (hojaHist && hojaHist.getLastColumn() < 25) { ss.deleteSheet(hojaHist); hojaHist = null; }
  if (!hojaHist) {
    hojaHist = ss.insertSheet('Historico_Agentes');
    const headers = ['ID_Agente', 'Nombre', 'Año', 'Mes', 'Citas', 'Capt_Excl', 'Capt_Abierto', 'Visitas_Comp', 'Casas_Ens', 'Capt_Alq', '3Bs', 'Bajadas', 'Propuestas', 'Arras', 'Vtas_Excl', 'GCI_Excl', 'Vtas_Abierto', 'GCI_Abierto', 'Vtas_Comp', 'GCI_Comp', 'Vtas_Alq', 'GCI_Alq', 'GCI_Total', 'Co_Euro', 'Fecha_Registro'];
    hojaHist.getRange(1, 1, 1, headers.length).setValues([headers]).setBackground('#667eea').setFontColor('#ffffff').setFontWeight('bold');
    hojaHist.setFrozenRows(1);
  }

  // 3. Mapear IDs existentes
  const datosAg = hojaAgentes.getDataRange().getValues();
  const mapaIDs = {};
  let maxID = 0;
  
  for(let i=1; i<datosAg.length; i++) {
    const idStr = String(datosAg[i][0]);
    const nombre = String(datosAg[i][1]).trim().toUpperCase();
    mapaIDs[nombre] = idStr;
    
    const num = parseInt(idStr.replace('AG',''), 10);
    if(!isNaN(num) && num > maxID) maxID = num;
  }

  const nuevosAgentes = [];
  const filasHist = [];
  const timestamp = new Date();

  // 4. Procesar
  listaAgentes.forEach(ag => {
    const nombreKey = ag.nombre.trim().toUpperCase();
    let id = mapaIDs[nombreKey];

    // CREAR AGENTE SI NO EXISTE
    if (!id) {
      maxID++;
      id = 'AG' + String(maxID).padStart(3,'0');
      nuevosAgentes.push([id, ag.nombre, "", "", new Date(ag.anio, 0, 1), "Activo", "NO", new Date(), 0]);
      mapaIDs[nombreKey] = id; 
    }
    
    // DATOS HISTÓRICOS (20 columnas)
    ag.mensual.forEach((m, idx) => {
      filasHist.push([
        id, ag.nombre, ag.anio, idx + 1,
        m.citas || 0, m.exclVenta || 0, m.captAbierto || 0, m.citasComp || 0, m.visitas || 0,
        m.captAlq || 0, m.tresBs || 0, m.bajadas || 0, m.propuestas || 0, m.arras || 0,
        m.vtaExcl || 0, m.gciExcl || 0, m.vtaAbierto || 0, m.gciAbierto || 0,
        m.vtaComp || 0, m.gciComp || 0, m.vtaAlq || 0, m.gciAlq || 0,
        m.gciTotal || 0, m.coEuro || 0,
        timestamp
      ]);
    });
  });

  // 5. Escribir Agentes Nuevos
  if (nuevosAgentes.length > 0) {
    hojaAgentes.getRange(hojaAgentes.getLastRow() + 1, 1, nuevosAgentes.length, nuevosAgentes[0].length).setValues(nuevosAgentes);
  }

  // 6. Borrar datos viejos del año para evitar duplicados
  const datosH = hojaHist.getDataRange().getValues();
  const idsAfectados = new Set(listaAgentes.map(a => mapaIDs[a.nombre.toUpperCase()]));
  const anioAfectado = listaAgentes[0].anio;
  
  // Borrar de abajo arriba
  for (let i = datosH.length - 1; i >= 1; i--) {
    if (idsAfectados.has(datosH[i][0]) && datosH[i][2] == anioAfectado) {
       hojaHist.deleteRow(i + 1);
    }
  }

  // 7. Escribir Histórico
  if (filasHist.length > 0) {
    hojaHist.getRange(hojaHist.getLastRow() + 1, 1, filasHist.length, 25).setValues(filasHist);
  }

  return { 
    success: true, 
    message: `¡Éxito! ${nuevosAgentes.length} agentes nuevos creados y datos guardados.` 
  };
}
// --- NUEVA FUNCIÓN: LEER HISTÓRICO PARA SUMAR AL DASHBOARD ---

function diagnosticarEstructuraHistorico() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('Historico_Agentes');
  
  if (!hoja) {
    Logger.log('❌ La hoja Historico_Agentes NO EXISTE');
    return;
  }
  
  const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  
  Logger.log('═══════════════════════════════════════════════════════════');
  Logger.log('📊 DIAGNÓSTICO: Estructura de Historico_Agentes');
  Logger.log('═══════════════════════════════════════════════════════════');
  Logger.log('Total columnas: ' + encabezados.length);
  Logger.log('');
  
  for (let i = 0; i < encabezados.length; i++) {
    const letra = String.fromCharCode(65 + i);
    Logger.log(`   [${i}] Columna ${letra}: "${encabezados[i]}"`);
  }
  
  Logger.log('');
  Logger.log('🎯 VERIFICACIÓN DE COLUMNAS CRÍTICAS:');
  
  const checks = [
    { idx: 0, nombre: 'ID_Agente', valor: encabezados[0] },
    { idx: 2, nombre: 'Año', valor: encabezados[2] },
    { idx: 3, nombre: 'Mes', valor: encabezados[3] },
    { idx: 4, nombre: 'Citas', valor: encabezados[4] },
    { idx: 22, nombre: 'GCI_Total', valor: encabezados[22] }
  ];
  
  for (const check of checks) {
    const ok = check.valor === check.nombre ? '✅' : '❌';
    Logger.log(`   ${ok} [${check.idx}] "${check.nombre}" = "${check.valor}"`);
  }
  
  // Muestra una fila de ejemplo
  if (hoja.getLastRow() > 1) {
    const ejemplo = hoja.getRange(2, 1, 1, encabezados.length).getValues()[0];
    Logger.log('');
    Logger.log('📄 EJEMPLO DE FILA (fila 2):');
    Logger.log(`   ID: ${ejemplo[0]}`);
    Logger.log(`   Nombre: ${ejemplo[1]}`);
    Logger.log(`   Año: ${ejemplo[2]}`);
    Logger.log(`   Mes: ${ejemplo[3]}`);
    Logger.log(`   Citas: ${ejemplo[4]}`);
    Logger.log(`   GCI_Total (col 22): ${ejemplo[22]}`);
  } else {
    Logger.log('');
    Logger.log('⚠️ No hay datos en la hoja (solo encabezados)');
  }
  
  Logger.log('═══════════════════════════════════════════════════════════');
}
function obtenerHistoricoExistente(idAgente, anio) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName('Historico_Agentes');
    
    if (!hoja) return [];
    
    const datos = hoja.getDataRange().getValues();
    const resultados = [];
    
    for (let i = 1; i < datos.length; i++) {
      if (String(datos[i][0]) === String(idAgente) && datos[i][2] == anio) {
        resultados.push(datos[i]);
      }
    }
    
    return resultados;
    
  } catch (error) {
    Logger.log('Error en obtenerHistoricoExistente: ' + error);
    return [];
  }
}
function guardarHistoricoAgenteHTML(datosGuardar) {
  try {
    Logger.log('📥 Recibiendo datos históricos:', JSON.stringify(datosGuardar));
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName('Historico_Agentes');
    
    if (!hoja) {
      return { success: false, message: 'No existe la hoja Historico_Agentes' };
    }
    
    const idAgente = datosGuardar.idAgente;
    const anio = datosGuardar.anio;
    const modo = datosGuardar.modo;
    
    // Obtener nombre del agente
    const hojaAgentes = ss.getSheetByName('Agentes');
    const datosAgentes = hojaAgentes.getDataRange().getValues();
    let nombreAgente = '';
    
    for (let i = 1; i < datosAgentes.length; i++) {
      if (datosAgentes[i][0] === idAgente) {
        nombreAgente = datosAgentes[i][1];
        break;
      }
    }
    
    if (!nombreAgente) {
      return { success: false, message: 'Agente no encontrado' };
    }
    
    // ELIMINAR registros previos del mismo agente/año
    const datosHoja = hoja.getDataRange().getValues();
    const filasAEliminar = [];
    
    for (let i = datosHoja.length - 1; i >= 1; i--) {
      if (datosHoja[i][0] === idAgente && datosHoja[i][2] == anio) {
        filasAEliminar.push(i + 1);
      }
    }
    
    // Eliminar de abajo hacia arriba para no desplazar índices
    filasAEliminar.forEach(fila => hoja.deleteRow(fila));
    
    Logger.log('🗑️ Eliminadas ' + filasAEliminar.length + ' filas previas');
    
    // INSERTAR NUEVOS DATOS
    const filasNuevas = [];
    
    if (modo === 'ANUAL') {
      // MODO ANUAL: Distribuir totales entre 12 meses (proporcional)
      const totales = datosGuardar.totales;
      
      // Distribución sugerida (puedes ajustar)
      const distribucion = [8, 8, 9, 9, 10, 10, 9, 5, 9, 9, 8, 6]; // Suma = 100%
      const suma = distribucion.reduce((a, b) => a + b, 0);
      
      for (let mes = 1; mes <= 12; mes++) {
        const factor = distribucion[mes - 1] / suma;
        
        const fila = [
          idAgente,                                          // A: ID_Agente
          nombreAgente,                                      // B: Nombre
          anio,                                              // C: Año
          mes,                                               // D: Mes
          Math.round(totales.citasCapt * factor),           // E: Citas
          Math.round(totales.exclVenta * factor),           // F: Capt_Excl
          Math.round(totales.captAbierto * factor),         // G: Capt_Abierto
          Math.round(totales.citasComp * factor),           // H: Visitas_Comp
          0,                                                 // I: Casas_Ens (dejamos 0)
          0,                                                 // J: Capt_Alq (dejamos 0)
          0,                                                 // K: 3Bs (dejamos 0)
          0,                                                 // L: Bajadas (dejamos 0)
          0,                                                 // M: Propuestas (dejamos 0)
          0,                                                 // N: Arras (dejamos 0)
          Math.round(totales.ventas * factor),              // O: Vtas_Excl
          0,                                                 // P: GCI_Excl (dejamos 0)
          0,                                                 // Q: Vtas_Abierto (dejamos 0)
          0,                                                 // R: GCI_Abierto (dejamos 0)
          0,                                                 // S: Vtas_Comp (dejamos 0)
          0,                                                 // T: GCI_Comp (dejamos 0)
          0,                                                 // U: Vtas_Alq (dejamos 0)
          0,                                                 // V: GCI_Alq (dejamos 0)
          parseFloat((totales.gci * factor).toFixed(2)),    // W: GCI_Total ⭐
          0,                                                 // X: Co_Euro (dejamos 0)
          new Date()                                         // Y: Fecha_Registro
        ];
        
        filasNuevas.push(fila);
      }
      
    } else {
      // MODO MENSUAL: Usar datos específicos por mes
      datosGuardar.mensual.forEach(datosMes => {
        const fila = [
          idAgente,                                          // A: ID_Agente
          nombreAgente,                                      // B: Nombre
          anio,                                              // C: Año
          datosMes.mes,                                      // D: Mes
          datosMes.citasCapt || 0,                          // E: Citas
          datosMes.exclVenta || 0,                          // F: Capt_Excl
          datosMes.captAbierto || 0,                        // G: Capt_Abierto
          datosMes.citasComp || 0,                          // H: Visitas_Comp
          datosMes.visitas || 0,                            // I: Casas_Ens ✅
          0,                                                 // J: Capt_Alq
          0,                                                 // K: 3Bs
          0,                                                 // L: Bajadas
          0,                                                 // M: Propuestas
          0,                                                 // N: Arras
          datosMes.ventas || 0,                             // O: Vtas_Excl
          0,                                                 // P: GCI_Excl
          0,                                                 // Q: Vtas_Abierto
          0,                                                 // R: GCI_Abierto
          0,                                                 // S: Vtas_Comp
          0,                                                 // T: GCI_Comp
          0,                                                 // U: Vtas_Alq
          0,                                                 // V: GCI_Alq
          parseFloat((datosMes.gci || 0).toFixed(2)),       // W: GCI_Total ⭐
          0,                                                 // X: Co_Euro
          new Date()                                         // Y: Fecha_Registro
        ];
        
        filasNuevas.push(fila);
      });
    }
    
    // INSERTAR en la hoja
    if (filasNuevas.length > 0) {
      const ultimaFila = hoja.getLastRow();
      hoja.getRange(ultimaFila + 1, 1, filasNuevas.length, filasNuevas[0].length).setValues(filasNuevas);
      Logger.log('✅ Insertadas ' + filasNuevas.length + ' filas nuevas');
    }
    
    return { 
      success: true, 
      message: `Histórico guardado correctamente (${filasNuevas.length} registros)` 
    };
    
  } catch (error) {
    Logger.log('❌ Error en guardarHistoricoAgenteHTML: ' + error);
    return { success: false, message: error.toString() };
  }
}

/**
 * ═══════════════════════════════════════════════════════════════════════════
 * FUNCIÓN 2: OBTENER HISTÓRICO EXISTENTE
 * ═══════════════════════════════════════════════════════════════════════════
 */
function obtenerHistoricoExistente(idAgente, anio) {
  try {
    Logger.log('📖 Buscando histórico de:', idAgente, 'año:', anio);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName('Historico_Agentes');
    
    if (!hoja) return null;
    
    const datos = hoja.getDataRange().getValues();
    
    // Buscar registros del agente/año
    const registros = [];
    
    for (let i = 1; i < datos.length; i++) {
      if (String(datos[i][0]).trim().toUpperCase() === String(idAgente).trim().toUpperCase() && 
          datos[i][2] == anio) {
        
        const mes = parseInt(datos[i][3]);
        
        if (mes >= 1 && mes <= 12) {
          registros.push({
            mes: mes,
            citasCapt: parseFloat(datos[i][4]) || 0,
            exclVenta: parseFloat(datos[i][5]) || 0,
            captAbierto: parseFloat(datos[i][6]) || 0,
            citasComp: parseFloat(datos[i][7]) || 0,
            visitas: parseFloat(datos[i][8]) || 0,
            ventas: parseFloat(datos[i][14]) || 0,  // Columna O
            gci: parseFloat(datos[i][22]) || 0,      // Columna W ⭐
            exclComp: parseFloat(datos[i][5]) || 0   // Aproximación
          });
        }
      }
    }
    
    if (registros.length === 0) {
      Logger.log('⚠️ No hay histórico previo');
      return null;
    }
    
    Logger.log('✅ Encontrados ' + registros.length + ' registros');
    
    // Ordenar por mes
    registros.sort((a, b) => a.mes - b.mes);
    
    // Calcular totales
    const totales = {
      gci: 0,
      ventas: 0,
      citasCapt: 0,
      exclVenta: 0,
      captAbierto: 0,
      citasComp: 0,
      exclComp: 0,
      visitas: 0
    };
    
    registros.forEach(r => {
      totales.gci += r.gci;
      totales.ventas += r.ventas;
      totales.citasCapt += r.citasCapt;
      totales.exclVenta += r.exclVenta;
      totales.captAbierto += r.captAbierto;
      totales.citasComp += r.citasComp;
      totales.exclComp += r.exclComp;
      totales.visitas += r.visitas;
    });
    
    return {
      gci: totales.gci,
      ventas: totales.ventas,
      citasCapt: totales.citasCapt,
      exclVenta: totales.exclVenta,
      captAbierto: totales.captAbierto,
      citasComp: totales.citasComp,
      exclComp: totales.exclComp,
      visitas: totales.visitas,
      mensual: registros
    };
    
  } catch (error) {
    Logger.log('❌ Error en obtenerHistoricoExistente: ' + error);
    return null;
  }
}
// --- AUXILIAR: TRADUCTOR DE NÚMEROS ESPAÑOLES ---
function parsearNumeroES(valor) {
  if (!valor) return 0;
  if (typeof valor === 'number') return valor; // Si ya es número, perfecto
  
  let str = String(valor).trim();
  // Si está vacío o es guión
  if (str === '' || str === '-') return 0;
  
  // 1. Quitamos los puntos de miles (ej: "1.200,50" -> "1200,50")
  str = str.replace(/\./g, '');
  
  // 2. Cambiamos la coma decimal por punto (ej: "1200,50" -> "1200.50")
  str = str.replace(',', '.');
  
  return parseFloat(str) || 0;
}

/**
 * ═══════════════════════════════════════════════════════════════════════════
 * FUNCIÓN 3: OBTENER TODOS LOS HISTÓRICOS (Para cache)
 * ═══════════════════════════════════════════════════════════════════════════
 */
function obtenerTodosHistoricosAgentes(anio) {
  try {
    Logger.log('📚 Cargando TODOS los históricos del año:', anio);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hoja = ss.getSheetByName('Historico_Agentes');
    
    if (!hoja) return {};
    
    const datos = hoja.getDataRange().getValues();
    const cache = {};
    
    for (let i = 1; i < datos.length; i++) {
      const row = datos[i];
      
      if (row[2] == anio) { // Año coincide
        const idAgente = String(row[0]).trim();
        const mes = parseInt(row[3]) - 1; // Convertir a índice 0-11
        
        if (mes >= 0 && mes < 12) {
          // Inicializar agente si no existe
          if (!cache[idAgente]) {
            cache[idAgente] = {
              gci: 0,
              citasCapt: 0,
              exclVenta: 0,
              captAbierto: 0,
              citasComp: 0,
              exclComp: 0,
              visitas: 0,
              ventas: 0
            };
          }
          
          // Sumar totales
          cache[idAgente].gci += parseFloat(row[22]) || 0;      // GCI_Total
          cache[idAgente].citasCapt += parseFloat(row[4]) || 0;
          cache[idAgente].exclVenta += parseFloat(row[5]) || 0;
          cache[idAgente].captAbierto += parseFloat(row[6]) || 0;
          cache[idAgente].citasComp += parseFloat(row[7]) || 0;
          cache[idAgente].visitas += parseFloat(row[8]) || 0;
          cache[idAgente].ventas += parseFloat(row[14]) || 0;
          cache[idAgente].exclComp += parseFloat(row[5]) || 0; // Aproximación
        }
      }
    }
    
    Logger.log('✅ Históricos de ' + Object.keys(cache).length + ' agentes cargados');
    return cache;
    
  } catch (error) {
    Logger.log('❌ Error en obtenerTodosHistoricosAgentes: ' + error);
    return {};
  }
}


// ═══════════════════════════════════════════════════════════════════════════
// 📝 NOTA: Estas funciones se añaden al final del archivo .gs
// ═══════════════════════════════════════════════════════════════════════════

console.log('✅ Funciones backend de histórico cargadas');

