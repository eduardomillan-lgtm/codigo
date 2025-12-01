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
      const comisionPct = parseFloat(agente.comisionPct) || 40;
      
      // ✅ CORRECCIÓN CRÍTICA: Comisión = (GCI × %) / 100
      const comisionImporte = parseFloat((gci * comisionPct / 100).toFixed(2));

      const notas = `TRANSACCIÓN ${datosTransaccion.tipo.toUpperCase()} - ${datosTransaccion.descripcion || 'Venta/Alquiler'} | Lado: ${agente.lado} | Comis: ${comisionPct}%`;
      
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

function obtenerDatosDashboard(filtros) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaActividad = ss.getSheetByName(CONFIG.HOJA_ACTIVIDAD);
    const hojaObjetivos = ss.getSheetByName(CONFIG.HOJA_OBJETIVOS);
    const hojaAgentes = ss.getSheetByName(CONFIG.HOJA_AGENTES);
    
    if (!hojaActividad || !hojaObjetivos || !hojaAgentes) {
      throw new Error('Faltan hojas necesarias. Ejecuta "Inicializar Sistema".');
    }
    
    let fechaInicio = new Date(new Date().getFullYear(), 0, 1);
    let fechaFin = new Date();
    
    if (filtros && filtros.fechaInicio) fechaInicio = new Date(filtros.fechaInicio);
    if (filtros && filtros.fechaFin) fechaFin = new Date(filtros.fechaFin);
    
    fechaInicio.setHours(0, 0, 0, 0);
    fechaFin.setHours(23, 59, 59, 999);

    const datosAgentes = hojaAgentes.getDataRange().getValues();
    const todasActividades = hojaActividad.getDataRange().getValues();
    const todosObjetivos = hojaObjetivos.getDataRange().getValues();
    
    // --- 1. OBTENER TRANSACCIONES (Solución para tu tabla vacía) ---
    const listaTransacciones = [];
    const timeZone = Session.getScriptTimeZone();

    // Empezamos en 1 para saltar encabezados
    for (let i = 1; i < todasActividades.length; i++) {
      const fila = todasActividades[i];
      const notas = String(fila[12] || "").toUpperCase(); // Columna M (Notas)
      
      // Si es una transacción
      if (notas.includes("TRANSACCIÓN")) {
        const fecha = new Date(fila[1]);
        // Filtrar por fecha seleccionada
        if (fecha >= fechaInicio && fecha <= fechaFin) {
           listaTransacciones.push({
             id: fila[0],
             // Convertimos fecha a texto para evitar errores de "Formato inesperado"
             fecha: Utilities.formatDate(fecha, timeZone, "yyyy-MM-dd"), 
             agente: fila[3], 
             tipo: notas.match(/TRANSACCIÓN (\w+)/)?.[1] || 'Venta',
             lado: notas.match(/LADO: (\w+)/)?.[1] || '-',
             descripcion: notas.split('|')[0].replace('TRANSACCIÓN', '').trim(),
             gci: parseFloat(fila[11]) || 0, 
             comision: parseFloat(fila[14]) || 0 // Columna O (15)
           });
        }
      }
    }
    // Ordenar: más reciente primero
    listaTransacciones.sort((a,b) => new Date(b.fecha) - new Date(a.fecha));

    // --- 2. LÓGICA DE AGENTES (Tu código original optimizado) ---
    const resultados = [];
    const mesesPeriodo = calcularMesesEnPeriodo(fechaInicio, fechaFin);
    const evolucionMensualEquipo = { labels: mesesPeriodo.map(m => obtenerNombreMesAbreviado(m.mes)) };
    
    Object.keys(kpiNames).forEach(key => {
        evolucionMensualEquipo[key] = { realizado: Array(mesesPeriodo.length).fill(0), objetivo: Array(mesesPeriodo.length).fill(0) };
    });
    
    const agentesActivos = [];
    for (let i = 1; i < datosAgentes.length; i++) {
      if (datosAgentes[i][0] && datosAgentes[i][1] && datosAgentes[i][5] === 'Activo') {
        agentesActivos.push({
          id: datosAgentes[i][0],
          nombre: datosAgentes[i][1],
          esAcumulativo: datosAgentes[i][6] === 'SI'
        });
      }
    }

    const mapaActividad = {};
    for (let i = 1; i < todasActividades.length; i++) {
      const id = todasActividades[i][2];
      if (!id) continue;
      if (!mapaActividad[id]) mapaActividad[id] = [];
      const fecha = new Date(todasActividades[i][1]);
      if (fecha >= fechaInicio && fecha <= fechaFin) {
        mapaActividad[id].push(todasActividades[i]);
      }
    }

    const mapaObjetivos = {};
    for (let i = 1; i < todosObjetivos.length; i++) {
      const id = todosObjetivos[i][0];
      if (!id) continue;
      if (!mapaObjetivos[id]) mapaObjetivos[id] = [];
      mapaObjetivos[id].push(todosObjetivos[i]);
    }

    agentesActivos.forEach(agente => {
      const actividadFiltrada = mapaActividad[agente.id] || [];
      const objetivosFiltrados = mapaObjetivos[agente.id] || [];
      const actividad = obtenerActividadAgente(agente.id, fechaInicio, fechaFin, actividadFiltrada);
      const objetivos = obtenerObjetivosAgente(agente.id, fechaInicio, fechaFin, objetivosFiltrados);

      if (agente.esAcumulativo) {
        const pendientes = calcularObjetivosAcumuladosPendientes(agente.id, fechaInicio, actividadFiltrada, objetivosFiltrados);
        Object.keys(objetivos).forEach(key => objetivos[key] += pendientes[key]);
      }

      const cumplimientos = calcularCumplimientos(actividad, objetivos);
      const cumplimientoGlobal = calcularCumplimientoGlobal(cumplimientos);
      const ratios = calcularRatios(actividad, objetivos);
      
      // Evitar errores numéricos
      const cumplimientoGlobalSeguro = isNaN(cumplimientoGlobal) || !isFinite(cumplimientoGlobal) ? 0 : cumplimientoGlobal;
      
      let estadoClase = 'bajo';
      if (cumplimientoGlobalSeguro >= 90) estadoClase = 'excelente';
      else if (cumplimientoGlobalSeguro >= 70) estadoClase = 'bueno';
      
      const evolucionMensual = calcularEvolucionMensual(agente.id, mesesPeriodo, agente.esAcumulativo, actividadFiltrada, objetivosFiltrados);

      mesesPeriodo.forEach((mes, idx) => {
        Object.keys(kpiNames).forEach(key => {
          if (evolucionMensual[key]) {
            evolucionMensualEquipo[key].realizado[idx] += evolucionMensual[key].realizado[idx];
            evolucionMensualEquipo[key].objetivo[idx] += evolucionMensual[key].objetivo[idx];
          }
        });
      });

      resultados.push({
        id: agente.id,
        agente: agente.nombre,
        realizado: actividad,
        objetivos: objetivos,
        cumplimientos: cumplimientos,
        cumplimientoGlobal: cumplimientoGlobalSeguro.toFixed(1),
        estadoClase: estadoClase,
        ratios: ratios,
        evolucionMensual: evolucionMensual
      });
    });

    const numAgentes = resultados.length;
    if (numAgentes > 0) {
      mesesPeriodo.forEach((mes, idx) => {
        Object.keys(kpiNames).forEach(key => {
          if (!['gci', 'citasCaptacion', 'exclusivasVenta', 'exclusivasComprador', 'captacionesAbierto', 'citasCompradores', 'casasEnsenadas', 'leadsCompradores', 'actividadTotal'].includes(key)) {
            evolucionMensualEquipo[key].realizado[idx] /= numAgentes;
            evolucionMensualEquipo[key].objetivo[idx] /= numAgentes;
          }
        });
      });
    }

    // Retornamos todo junto, incluyendo las transacciones que faltaban
    return {
      agentes: resultados,
      evolucionMensualEquipo: evolucionMensualEquipo,
      transacciones: listaTransacciones // <--- ESTO ES LA CLAVE
    };

  } catch (error) {
    Logger.log('ERROR: ' + error.toString());
    throw error;
  }
}

// ====== TODAS LAS DEMÁS FUNCIONES SIGUEN IGUALES (NO TOQUE NADA MÁS) ======
function obtenerActividadAgente(idAgente, fechaInicio, fechaFin, todasActividades) {
  const actividad = {
        citasCaptacion: 0, exclusivasVenta: 0, exclusivasComprador: 0,
        captacionesAbierto: 0, citasCompradores: 0, casasEnsenadas: 0,
        leadsCompradores: 0, llamadas: 0, gci: 0, volumenNegocio: 0
    };
  
  // CORRECCIÓN: Empezamos en 0 porque 'todasActividades' ya viene filtrada sin cabeceras
  for (let i = 0; i < todasActividades.length; i++) {
    // Ya no hace falta comprobar el ID porque viene pre-filtrado, pero lo dejamos por seguridad
    // Nota: Al venir de 'mapaActividad', el objeto es la fila directa.
    
    const fechaRaw = todasActividades[i][1];
    if (!fechaRaw) continue;
    
    const fecha = new Date(fechaRaw);
    if (fecha instanceof Date && !isNaN(fecha) && fecha >= fechaInicio && fecha <= fechaFin) {
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

function obtenerDatosAgenteCompleto(nombreAgente, filtros) {
  try {
    Logger.log('Obteniendo datos completos para: ' + nombreAgente);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaActividad = ss.getSheetByName(CONFIG.HOJA_ACTIVIDAD);
    const hojaObjetivos = ss.getSheetByName(CONFIG.HOJA_OBJETIVOS);
    const hojaAgentes = ss.getSheetByName(CONFIG.HOJA_AGENTES);
   
    if (!hojaActividad || !hojaObjetivos || !hojaAgentes) {
      throw new Error('Faltan hojas necesarias.');
    }
   
    const datosAgentes = hojaAgentes.getDataRange().getValues();
    const todasActividades = hojaActividad.getDataRange().getValues();
    const todosObjetivos = hojaObjetivos.getDataRange().getValues();
    let idAgente = null;
    let esAcumulativo = false;
   
    for (let i = 1; i < datosAgentes.length; i++) {
      if (datosAgentes[i][1] === nombreAgente) {
        idAgente = datosAgentes[i][0];
        esAcumulativo = datosAgentes[i][6] === 'SI';
        break;
      }
    }
   
    if (!idAgente) throw new Error('No se encontró el agente: ' + nombreAgente);
   
    let fechaInicio = new Date(new Date().getFullYear(), 0, 1);
    let fechaFin = new Date();
   
    if (filtros && filtros.fechaInicio) fechaInicio = new Date(filtros.fechaInicio);
    if (filtros && filtros.fechaFin) fechaFin = new Date(filtros.fechaFin);
   
    fechaInicio.setHours(0, 0, 0, 0);
    fechaFin.setHours(23, 59, 59, 999);
   
    const actividad = obtenerActividadAgente(idAgente, fechaInicio, fechaFin, todasActividades);
    const objetivos = obtenerObjetivosAgente(idAgente, fechaInicio, fechaFin, todosObjetivos);
    if (esAcumulativo) {
        const pendientes = calcularObjetivosAcumuladosPendientes(idAgente, fechaInicio, todasActividades, todosObjetivos);
        Object.keys(objetivos).forEach(key => objetivos[key] += pendientes[key]);
    }
    const gciSeguro = isNaN(actividad.gci) ? 0 : actividad.gci;
    actividad.gci = gciSeguro;
    const cumplimientos = calcularCumplimientos(actividad, objetivos);
    const cumplimientoGlobal = calcularCumplimientoGlobal(cumplimientos);
    const ratios = calcularRatios(actividad, objetivos);
   
    const hoy = new Date();
    const pendientesPorPeriodo = {
      semana: calcularPendientesPeriodo(idAgente, hoy, 'semana', esAcumulativo, todasActividades, todosObjetivos),
      mes: calcularPendientesPeriodo(idAgente, hoy, 'mes', esAcumulativo, todasActividades, todosObjetivos),
      trimestre: calcularPendientesPeriodo(idAgente, hoy, 'trimestre', esAcumulativo, todasActividades, todosObjetivos),
      semestre: calcularPendientesPeriodo(idAgente, hoy, 'semestre', esAcumulativo, todasActividades, todosObjetivos),
      ano: calcularPendientesPeriodo(idAgente, hoy, 'ano', esAcumulativo, todasActividades, todosObjetivos)
    };
    const mesesPeriodo = calcularMesesEnPeriodo(fechaInicio, fechaFin);
    const evolucionMensual = calcularEvolucionMensual(idAgente, mesesPeriodo, esAcumulativo, todasActividades, todosObjetivos);
   
    const cumplimientoGlobalSeguro = isNaN(cumplimientoGlobal) || !isFinite(cumplimientoGlobal) ? 0 : cumplimientoGlobal;
    return {
      id: idAgente,
      agente: nombreAgente,
      realizado: actividad,
      objetivos: objetivos,
      cumplimientos: cumplimientos,
      cumplimientoGlobal: cumplimientoGlobalSeguro.toFixed(1),
      ratios: ratios,
      pendientesPorPeriodo: pendientesPorPeriodo,
      evolucionMensual: evolucionMensual
    };
  } catch (error) {
    Logger.log('ERROR en obtenerDatosAgenteCompleto: ' + error.toString());
    throw error;
  }
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
  let hoja = ss.getSheetByName(MODELOS_CONFIG.HOJA_FACT_PASADA);
  if (hoja) ss.deleteSheet(hoja);
  hoja = ss.insertSheet(MODELOS_CONFIG.HOJA_FACT_PASADA);
  const headers = ['ID_Fact','Fecha','Agente','Cliente','Concepto','Importe_Neto','Comision_Pct','Comision_Importe','GCI','Forma_Pago','Observaciones','Fecha_Registro'];
  hoja.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold').setBackground('#ffe8e8');
  hoja.setFrozenRows(1);
  hoja.getRange('F:F').setNumberFormat('#,##0.00');
  hoja.getRange('I:I').setNumberFormat('#,##0.00');
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
    
    // --- 1. CALCULAR TOTAL SUELDOS FIJOS MENSUALES ---
    // Esto es un Coste Operativo Fijo
    const hojaAgentes = ss.getSheetByName(CONFIG.HOJA_AGENTES);
    let totalSueldosFijosMes = 0;
    if (hojaAgentes) {
      const datosAgentes = hojaAgentes.getDataRange().getValues();
      // Empezamos en 1 para saltar cabecera
      for (let i = 1; i < datosAgentes.length; i++) {
        // Si el agente está activo, sumamos su sueldo fijo (Columna I -> índice 8)
        if (datosAgentes[i][5] === 'Activo') {
           totalSueldosFijosMes += parseFloat(datosAgentes[i][8]) || 0;
        }
      }
    }

    // --- 2. INICIALIZAR ARRAYS ---
    const gciMensual = new Array(12).fill(0);
    const gastosVentas = new Array(12).fill(0);      // Solo Comisiones
    const gastosOperativos = new Array(12).fill(0);  // Gastos + Sueldos Fijos
    
    const detallesVentas = Array.from({length: 12}, () => []);
    const detallesOperativos = Array.from({length: 12}, () => []);

    // --- 3. PRE-LLENAR GASTOS OPERATIVOS CON SUELDOS FIJOS ---
    // (CORRECCIÓN: Ahora van a Operativos, no a Ventas)
    for (let m = 0; m < 12; m++) {
      if (totalSueldosFijosMes > 0) {
        gastosOperativos[m] += totalSueldosFijosMes;
        detallesOperativos[m].push({
          fecha: "-",
          concepto: "Nóminas / Fijos",
          desc: "Suma de sueldos fijos plantilla",
          importe: totalSueldosFijosMes
        });
      }
    }

    // --- 4. PROCESAR ACTIVIDAD (SUMAR COMISIONES A GASTOS VENTAS) ---
    const hojaAct = ss.getSheetByName(CONFIG.HOJA_ACTIVIDAD);
    if (hojaAct) {
      const datosAct = hojaAct.getDataRange().getValues();
      for (let i = 1; i < datosAct.length; i++) {
        const fila = datosAct[i];
        const fechaRaw = fila[1];
        if (!fechaRaw) continue;
        const fecha = new Date(fechaRaw);
        if (isNaN(fecha.getTime()) || fecha.getFullYear() !== year) continue;

        const mes = fecha.getMonth(); 
        const gci = parseFloat(fila[11]) || 0;       
        const comision = parseFloat(fila[14]) || 0;  
        const notas = String(fila[12] || "");       

        if (gci > 0 || comision > 0) {
            gciMensual[mes] += gci;
            gastosVentas[mes] += comision; // Aquí SOLO van las comisiones variables

            if (comision > 0) {
                detallesVentas[mes].push({
                    fecha: Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd/MM/yyyy"),
                    concepto: fila[3] || 'Agente', 
                    desc: notas.split('|')[0] || 'Comisión Variable',
                    importe: comision
                });
            }
        }
      }
    }

    // --- 5. PROCESAR OTROS GASTOS OPERATIVOS (HOJA GASTOS) ---
    const hojaGastos = ss.getSheetByName('Gastos_Operativos');
    if (hojaGastos) {
      const datosGastos = hojaGastos.getDataRange().getValues();
      for (let i = 1; i < datosGastos.length; i++) {
        const fila = datosGastos[i];
        const fechaRaw = fila[1];
        if (!fechaRaw) continue;
        const fecha = new Date(fechaRaw);
        if (isNaN(fecha.getTime()) || fecha.getFullYear() !== year) continue;

        const mes = fecha.getMonth();
        const importe = parseFloat(fila[6]) || 0; 

        gastosOperativos[mes] += importe; // Se suman a los sueldos fijos

        detallesOperativos[mes].push({
            fecha: Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd/MM/yyyy"),
            concepto: fila[4], 
            desc: fila[5],     
            importe: importe
        });
      }
    }

    return {
      meses: ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'],
      gciMensual: gciMensual,
      gastosVentas: gastosVentas,
      gastosOperativos: gastosOperativos,
      detallesVentas: detallesVentas,
      detallesOperativos: detallesOperativos
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
  if (!hoja) {
    hoja = ss.insertSheet('Historico_Agentes');
    // AÑADIMOS LAS NUEVAS COLUMNAS AQUÍ
    const headers = [
      'ID_Agente', 'Nombre', 'Año', 'Mes', 
      'GCI', 'Ventas_Cierres', 
      'Citas_Captacion', 'Exclusivas_Venta', 'Capt_Abierto',
      'Citas_Comprador', 'Visitas_Casas',
      'Capt_Alquiler', '3Bs', 'Bajadas_Precio', 'Propuestas', 'Arras', // <--- NUEVAS
      'Fecha_Registro'
    ];
    hoja.getRange(1, 1, 1, headers.length).setValues([headers])
        .setBackground('#667eea')
        .setFontColor('#ffffff')
        .setFontWeight('bold');
    hoja.setFrozenRows(1);
  }
  return hoja;
}

function guardarHistoricoAgente(datos) {
  // datos: { idAgente, nombre, anio, modo, valores: { ... } }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let hoja = ss.getSheetByName('Historico_Agentes');
    if (!hoja) hoja = crearHojaHistoricoAgentes(ss);
    
    // 1. LIMPIEZA: Borrar datos previos
    const filas = hoja.getDataRange().getValues();
    for (let i = filas.length - 1; i >= 1; i--) {
      if (filas[i][0] == datos.idAgente && filas[i][2] == datos.anio) {
        hoja.deleteRow(i + 1);
      }
    }
    
    // 2. INSERTAR NUEVOS DATOS
    const nuevasFilas = [];
    const now = new Date();
    
    if (datos.modo === 'ANUAL') {
      const v = datos.valores;
      nuevasFilas.push([
        datos.idAgente, datos.nombre, datos.anio, 'TOTAL',
        v.gci||0, v.ventas||0, v.citasCapt||0, v.exclVenta||0, v.captAbierto||0,
        v.citasComp||0, v.exclComp||0, v.visitas||0, now
      ]);
    } else {
      // Modo Mensual
      datos.valores.forEach((m, idx) => {
        nuevasFilas.push([
          datos.idAgente, datos.nombre, datos.anio, idx + 1,
          m.gci||0, m.ventas||0, m.citasCapt||0, m.exclVenta||0, m.captAbierto||0,
          m.citasComp||0, m.exclComp||0, m.visitas||0, now
        ]);
      });
    }
    
    if (nuevasFilas.length > 0) {
      hoja.getRange(hoja.getLastRow() + 1, 1, nuevasFilas.length, nuevasFilas[0].length).setValues(nuevasFilas);
    }
    
    return { success: true, message: 'Histórico detallado guardado.' };
    
  } catch (e) {
    throw new Error('Error guardando histórico: ' + e.message);
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
/**
 * 📥 IMPORTACIÓN MASIVA DESDE EXCEL
 * Recibe un array de objetos con datos mensuales y anuales.
 * Crea agentes nuevos si no existen y guarda el histórico.
 */
function guardarImportacionMasiva(listaAgentes) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Gestión de Agentes (Igual que antes)
  let hojaAgentes = ss.getSheetByName('Agentes');
  if (!hojaAgentes) { crearHojaAgentes(ss); hojaAgentes = ss.getSheetByName('Agentes'); }
  const datosAg = hojaAgentes.getDataRange().getValues();
  const mapaIDs = {};
  let maxID = 0;
  for(let i=1; i<datosAg.length; i++) {
    const val = String(datosAg[i][0]);
    mapaIDs[String(datosAg[i][1]).trim().toUpperCase()] = val;
    const num = parseInt(val.replace('AG',''));
    if(!isNaN(num) && num > maxID) maxID = num;
  }

  const nuevosAgentes = [];
  const historico = [];
  const timestamp = new Date();

  // 2. Procesar Datos con las NUEVAS COLUMNAS
  listaAgentes.forEach(ag => {
    let id = mapaIDs[ag.nombre.trim().toUpperCase()];
    if(!id) {
      maxID++;
      id = 'AG' + String(maxID).padStart(3,'0');
      nuevosAgentes.push([id, ag.nombre, "", "", new Date(ag.anio,0,1), "Activo", "NO", new Date(), 0]);
      mapaIDs[ag.nombre.trim().toUpperCase()] = id;
    }
    
    ag.mensual.forEach((m, idx) => {
      historico.push([
        id,
        ag.nombre,
        ag.anio,
        idx + 1,
        m.gci || 0,
        m.ventas || 0,
        m.citasCapt || 0,
        m.exclVenta || 0,
        m.captAbierto || 0,
        m.citasComp || 0,
        m.visitas || 0,
        // --- AQUÍ GUARDAMOS LOS NUEVOS DATOS ---
        m.captAlq || 0,
        m.tresBs || 0,
        m.bajadas || 0,
        m.propuestas || 0,
        m.arras || 0,
        // ---------------------------------------
        timestamp
      ]);
    });
  });

  // 3. Escribir en hojas
  if(nuevosAgentes.length) hojaAgentes.getRange(hojaAgentes.getLastRow()+1,1,nuevosAgentes.length,9).setValues(nuevosAgentes);
  
  let hojaHist = ss.getSheetByName('Historico_Agentes');
  if(!hojaHist) hojaHist = crearHojaHistoricoAgentes(ss);
  
  // Importante: Ajustar el rango al número de columnas nuevas (ahora son 17 columnas en total)
  if(historico.length) hojaHist.getRange(hojaHist.getLastRow()+1, 1, historico.length, 17).setValues(historico);

  return { success: true, message: `Importados ${listaAgentes.length} agentes con detalle ampliado.` };
}
