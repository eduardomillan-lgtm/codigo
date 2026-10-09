/**
 * Pruebas de la lógica de reclutamiento.gs
 *
 * Apps Script no tiene framework de test, así que esto carga el archivo en un
 * contexto de Node con los servicios de Google simulados y comprueba las
 * funciones puras: teléfonos, deduplicación, robots.txt, scoring, plantillas
 * y supresión.
 *
 * Ejecutar:   node test_reclutamiento.js
 * Úsalo cada vez que toques los pesos del scoring o la normalización.
 */

const fs = require('fs'), vm = require('vm');


// ── Stubs mínimos de Apps Script ──────────────────────────────
const configFalsa = { TL_NOMBRE:'Laura Ruiz', TL_TELEFONO:'+34600111222',
  MC_EMAIL:'marbella@kwspain.es', MC_DIRECCION:'Av. Ricardo Soriano 12' };
const hojaConfig = {
  getLastRow: () => 1 + Object.keys(configFalsa).length,
  getRange: () => ({ getValues: () => Object.keys(configFalsa).map(k => [k, configFalsa[k]]) })
};
const ctx = {
  console,
  SpreadsheetApp: {
    getActiveSpreadsheet: () => ({ getSheetByName: n => n === 'Rec_Config' ? hojaConfig : null,
                                   toast: () => {} }),
    getUi: () => ({ alert: () => {}, ButtonSet:{OK:1}, Button:{YES:1} }),
    newDataValidation: () => ({ requireValueInList:()=>({setAllowInvalid:()=>({build:()=>({})})}) }),
    newConditionalFormatRule: () => ({ whenTextEqualTo:()=>({setBackground:()=>({setFontColor:()=>({setRanges:()=>({build:()=>({})})})})}) })
  },
  PropertiesService: { getScriptProperties: () => ({ getProperty: () => 'FAKE', setProperty: () => {} }) },
  CacheService: { getScriptCache: () => ({ get: () => null, put: () => {} }) },
  Utilities: { getUuid: () => 'abcdef12-3456', formatDate: () => '2026-10-08',
               base64EncodeWebSafe: s => Buffer.from(s).toString('base64'),
               sleep: () => {}, parseCsv: (t,d) => t.trim().split('\n').map(l => l.split(d||',')) },
  Session: { getScriptTimeZone: () => 'Europe/Madrid', getActiveUser: () => ({getEmail:()=>'x@y.z'}),
             getEffectiveUser: () => ({getEmail:()=>'x@y.z'}) },
  UrlFetchApp: { fetch: () => ({ getResponseCode: () => 200, getContentText: () => '{}' }) },
  Logger: { log: () => {} }, HtmlService: {}, DriveApp: {}, MailApp: {}, ScriptApp: {}, MimeType: {}
};
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(require('path').join(__dirname, 'reclutamiento.gs'), 'utf8'), ctx);

let ok = 0, fail = 0;
function t(nombre, real, esperado) {
  const pasa = JSON.stringify(real) === JSON.stringify(esperado);
  if (pasa) { ok++; console.log('  ✅ ' + nombre); }
  else { fail++; console.log('  ❌ ' + nombre + '\n       obtenido: ' + JSON.stringify(real) +
                             '\n       esperado: ' + JSON.stringify(esperado)); }
}
function tt(nombre, cond) { if (cond) { ok++; console.log('  ✅ ' + nombre); }
                            else { fail++; console.log('  ❌ ' + nombre); } }

console.log('\n── Normalización de teléfonos ──');
t('móvil español sin prefijo',   ctx.recNormalizarTelefono_('600 123 456'), '+34600123456');
t('fijo de Málaga',              ctx.recNormalizarTelefono_('952 90 00 00'), '+34952900000');
t('con prefijo ya puesto',       ctx.recNormalizarTelefono_('+34 600 123 456'), '+34600123456');
t('formato 0034',                ctx.recNormalizarTelefono_('0034600123456'), '+34600123456');
t('34 sin plus',                 ctx.recNormalizarTelefono_('34600123456'), '+34600123456');
t('británico (comprador UK)',    ctx.recNormalizarTelefono_('+44 7700 900123'), '+447700900123');
t('con ruido tipográfico',       ctx.recNormalizarTelefono_('Tel.: 600-123-456 (móvil)'), '+34600123456');
t('vacío',                       ctx.recNormalizarTelefono_(''), '');
t('demasiado corto se descarta', ctx.recNormalizarTelefono_('1234'), '');

console.log('\n── Clave de deduplicación ──');
t('prioriza el teléfono', ctx.recClaveDedupe_('Ana García','600123456','a@b.es'), 'T:+34600123456');
t('cae al email',         ctx.recClaveDedupe_('Ana García','','A@B.es'), 'E:a@b.es');
t('cae al nombre normalizado', ctx.recClaveDedupe_('Ana Mª Gárcía-López','',''), 'N:ana m garcialopez');
tt('mismo tel en dos formatos = misma clave',
   ctx.recClaveDedupe_('Ana','600123456','') === ctx.recClaveDedupe_('Ana G.','+34 600 12 34 56',''));

console.log('\n── robots.txt ──');
const r1 = 'User-agent: *\nDisallow: /equipo\nAllow: /equipo/publico';
tt('bloquea /equipo',            ctx.recRobotsPermite_(r1,'/equipo') === false);
tt('Allow más específico gana',  ctx.recRobotsPermite_(r1,'/equipo/publico') === true);
tt('permite ruta no listada',    ctx.recRobotsPermite_(r1,'/team') === true);
tt('robots vacío permite todo',  ctx.recRobotsPermite_('','/equipo') === true);
tt('ignora reglas de otro bot',
   ctx.recRobotsPermite_('User-agent: Googlebot\nDisallow: /\n','/equipo') === true);
tt('Disallow: / bloquea todo',   ctx.recRobotsPermite_('User-agent: *\nDisallow: /','/equipo') === false);
tt('respeta comentarios',        ctx.recRobotsPermite_('# nota\nUser-agent: *\nDisallow: /x # otra','/equipo') === true);

console.log('\n── Clasificación ──');
t('franquicia detectada', ctx.recClasificarModelo_('Engel & Völkers Marbella'), 'Gran franquicia');
t('independiente',        ctx.recClasificarModelo_('Villas de Lujo Marbella SL'), 'Independiente');
t('perfil autónomo',      ctx.recPerfilDesdeCargo_('Asesor inmobiliario freelance','Marca propia'), 'Agente autónomo / sin agencia');
t('perfil cambio sector', ctx.recPerfilDesdeCargo_('Guest Relations Manager','Hotel Puente Romano'), 'Cambio de sector (hostelería/lujo/banca)');
t('perfil KW se aparta',  ctx.recPerfilDesdeCargo_('Agente','Keller Williams Málaga'), 'Agente en otro MC KW');
t('perfil TL',            ctx.recPerfilDesdeCargo_('Team Leader','RE/MAX'), 'Team Leader / Broker');
t('idioma EN por apellido no hispano', ctx.recIdiomaProbable_({nombre:'Lars', apellidos:'Johansson'}), 'EN');
t('idioma ES por defecto', ctx.recIdiomaProbable_({nombre:'Ana', apellidos:'Gómez'}), 'ES');

console.log('\n── Scoring ──');
// los `const` de nivel superior no se cuelgan del contexto: hay que evaluarlos
const C = vm.runInContext('REC_COL', ctx), N = vm.runInContext('REC_N_COLS', ctx);
const PESOS = vm.runInContext('REC.PESOS', ctx);
function ficha(o) { const f = new Array(N).fill('');
  Object.keys(o).forEach(k => { f[C[k]-1] = o[k]; }); return f; }

const top = ficha({ INMUEBLES:30, PRECIO_MEDIO:2500000, ZONA:'Nueva Andalucía',
  PERFIL:'Agente autónomo / sin agencia', IDIOMAS:'Español, Inglés, Sueco',
  EXPERIENCIA:5, TELEFONO:'+34600111222', FUENTE:'Referido de agente',
  NOTAS:'comisión baja, sin exclusivas' });
const flojo = ficha({ PERFIL:'Recién titulado / sin experiencia', ZONA:'', FUENTE:'Web de agencia' });
const kw = ficha({ INMUEBLES:40, PRECIO_MEDIO:3000000, ZONA:'Marbella',
  PERFIL:'Agente en otro MC KW', IDIOMAS:'Español, Inglés', EXPERIENCIA:6, TELEFONO:'+34600111222' });
const medio = ficha({ INMUEBLES:10, PRECIO_MEDIO:600000, ZONA:'Estepona',
  PERFIL:'Agente en gran franquicia', IDIOMAS:'Español, Inglés', EXPERIENCIA:3, TELEFONO:'+34600111222' });

const sTop = ctx.recScoreCandidato_(top), sFlojo = ctx.recScoreCandidato_(flojo);
const sKw = ctx.recScoreCandidato_(kw), sMedio = ctx.recScoreCandidato_(medio);
console.log('     top=' + sTop + ' medio=' + sMedio + ' kw=' + sKw + ' flojo=' + sFlojo);
tt('el candidato ideal sale A',        ctx.recTemperatura_(sTop) === 'A');
tt('score dentro de 0-100',            sTop <= 100 && sFlojo >= 0);
tt('ordena top > medio > flojo',       sTop > sMedio && sMedio > sFlojo);
tt('un agente de otro MC KW baja mucho', sKw < sMedio);
t('umbrales de temperatura', [ctx.recTemperatura_(80), ctx.recTemperatura_(60),
   ctx.recTemperatura_(40), ctx.recTemperatura_(10)], ['A','B','C','D']);

console.log('\n── Plantillas y enlaces ──');
const render = ctx.recRenderPlantilla_(
  'Hola {{nombre}}, soy {{tl}} de {{mc}}. Vi tus {{inmuebles}} en {{zona}}. Tel: {{telefono_tl}}', top);
tt('sustituye datos del candidato', render.indexOf('Nueva Andalucía') !== -1 && render.indexOf('30') !== -1);
tt('sustituye datos de Rec_Config',  render.indexOf('Laura Ruiz') !== -1 && render.indexOf('+34600111222') !== -1);
tt('no deja marcadores sin resolver', render.indexOf('{{') === -1);
const enlace = ctx.recEnlaceWhatsApp_('+34 600 123 456', 'Hola Ana\n\n⚠️ RELLENA ESTO\nmás texto');
tt('wa.me sin signos en el número',  enlace.indexOf('https://wa.me/34600123456?text=') === 0);
tt('quita las notas internas del enlace', enlace.indexOf('RELLENA') === -1);
tt('conserva el saludo',             decodeURIComponent(enlace).indexOf('Hola Ana') !== -1);

console.log('\n── HTML a texto ──');
const txt = ctx.recHtmlATexto_('<div><script>var x=1</script><p>Ana G&oacute;mez</p><p>Asesora &amp; socia</p></div>');
tt('elimina el script',  txt.indexOf('var x') === -1);
tt('decodifica entidades', txt.indexOf('&') !== -1 && txt.indexOf('&amp;') === -1);
tt('conserva el nombre',  txt.indexOf('Ana') !== -1);

console.log('\n── Supresión ──');
const sup = { '+34600123456': true, 'ana@kw.es': true };
tt('bloquea por teléfono en otro formato',
   ctx.recEstaSuprimido_(sup, '600 123 456', '', '') === true);
tt('bloquea por email en mayúsculas', ctx.recEstaSuprimido_(sup, '', 'ANA@KW.ES', '') === true);
tt('no bloquea a quien no está',      ctx.recEstaSuprimido_(sup, '699999999', 'otro@x.es', '') === false);

console.log('\n── Smart Plan: integridad de los datos sembrados ──');
tt('los pesos del scoring suman 100',
   Object.keys(PESOS).reduce((a,k) => a + PESOS[k], 0) === 100);
tt('REC_COL cubre las 36 columnas', Object.keys(C).length === 36);
tt('índices de columna sin huecos ni repetidos',
   JSON.stringify(Object.keys(C).map(k=>C[k]).sort((a,b)=>a-b))
   === JSON.stringify(Array.from({length:36},(_,i)=>i+1)));

console.log('\n── X-ray: parser de perfiles de LinkedIn ──');
const p1 = ctx.recParsearPerfilLinkedIn_({
  title: 'Ana García - Asesora Inmobiliaria - Panorama Properties | LinkedIn',
  link: 'https://www.linkedin.com/in/anagarcia-marbella?trk=abc',
  snippet: 'Ubicación: Marbella · 6 años de experiencia · Experiencia: Panorama Properties'
});
t('nombre y apellidos separados', [p1.nombre, p1.apellidos], ['Ana','García']);
t('cargo extraído',               p1.cargo, 'Asesora Inmobiliaria');
t('agencia extraída',             p1.agencia, 'Panorama Properties');
t('zona del fragmento',           p1.zona, 'Marbella');
t('años de experiencia',          p1.experiencia, '6');
t('URL sin parámetros',           p1.linkedin, 'https://www.linkedin.com/in/anagarcia-marbella');

const p2 = ctx.recParsearPerfilLinkedIn_({
  title: 'Lars Nilsson - Marbella, Andalucía, España | Perfil profesional | LinkedIn',
  link: 'https://es.linkedin.com/in/larsnilsson', snippet: ''
});
tt('ubicación en el título no se confunde con el cargo',
   p2.cargo === '' && p2.zona.indexOf('Marbella') !== -1);

tt('descarta enlaces que no son de perfil',
   ctx.recParsearPerfilLinkedIn_({ title:'Empleos de inmobiliaria | LinkedIn',
     link:'https://www.linkedin.com/jobs/search', snippet:'' }) === null);
tt('descarta páginas de listado',
   ctx.recParsearPerfilLinkedIn_({ title:'Perfiles de asesor inmobiliario | LinkedIn',
     link:'https://www.linkedin.com/in/x', snippet:'' }) === null);

t('normaliza URL con subdominio y querystring',
  ctx.recNormalizarLinkedIn_('https://es.linkedin.com/in/AnaGarcia?trk=x'), 'linkedin.com/in/anagarcia');
tt('misma URL en dos formatos = misma clave',
  ctx.recNormalizarLinkedIn_('http://linkedin.com/in/ana/') ===
  ctx.recNormalizarLinkedIn_('https://www.linkedin.com/in/ANA?utm=1'));
t('URL no válida da vacío', ctx.recNormalizarLinkedIn_('https://facebook.com/ana'), '');

console.log('\n── X-ray: consultas generadas ──');
const qs = vm.runInContext('recConsultasXRay_()', ctx);
tt('genera consultas',            qs.length >= 10);
tt('todas limitadas a perfiles',  qs.every(q => q.query.indexOf('site:linkedin.com/in') === 0));
tt('todas acotadas geográficamente',
   qs.every(q => /Marbella|Estepona|Benahav|San Pedro|Nueva Andaluc|Mijas|Sotogrande|Costa del Sol/.test(q.query)));
tt('cubre el segmento de autónomos',
   qs.some(q => /autónomo|freelance/.test(q.query)));
tt('cubre el cambio de sector',
   qs.some(q => /concierge|yacht|private banker/.test(q.query)));

console.log('\n── Kelly: mapeo de paso a plantilla de Meta ──');
t('invitación a evento', ctx.recPlantillaKellyPara_('Invitación a evento'), 'recruit_invitacion_evento');
t('envío de calculadora', ctx.recPlantillaKellyPara_('Valor 3 — calculadora'), 'recruit_recurso');
t('dato de mercado',     ctx.recPlantillaKellyPara_('Valor 1 — dato de mercado'), 'recruit_valor_mercado');
t('informe trimestral',  ctx.recPlantillaKellyPara_('Informe trimestral'), 'recruit_recurso');
t('nurture por defecto', ctx.recPlantillaKellyPara_('Nurture mensual'), 'recruit_seguimiento');
t('tipo desconocido cae en seguimiento', ctx.recPlantillaKellyPara_('xyz'), 'recruit_seguimiento');

const PLANT = vm.runInContext('REC_PLANTILLAS_KELLY', ctx);
tt('toda plantilla de reclutamiento es MARKETING',
   Object.keys(PLANT).every(k => PLANT[k].categoria === 'MARKETING'));
tt('el mapeo solo devuelve plantillas declaradas',
   ['Invitación a evento','Valor 3 — calculadora','Valor 1 — dato de mercado','Nurture mensual','xyz']
     .every(x => PLANT[ctx.recPlantillaKellyPara_(x)] !== undefined));

console.log('\n── Kelly: rangos de botón a número ──');
t('rango simple',            ctx.recRangoAPunto_('6-15'), 11);
t('rango 1-5',               ctx.recRangoAPunto_('1-5'), 3);
t('euros con punto de miles', ctx.recRangoAPunto_('20.000-40.000 €'), 30000);
t('euros con coma de miles',  ctx.recRangoAPunto_('€10,000-20,000'), 15000);
t('tope superior ES',        ctx.recRangoAPunto_('Más de 30'), 42);
t('tope superior EN',        ctx.recRangoAPunto_('More than 30'), 42);
t('tope superior en euros',  ctx.recRangoAPunto_('Más de 40.000 €'), 56000);
t('tope inferior ES',        ctx.recRangoAPunto_('Menos de 10.000 €'), 6000);
t('tope inferior EN',        ctx.recRangoAPunto_('Under €10,000'), 6000);
t('sin números da nulo',     ctx.recRangoAPunto_('Nada ahora mismo'), null);
t('vacío da nulo',           ctx.recRangoAPunto_(''), null);
tt('un decimal real se conserva', Math.abs(ctx.recRangoAPunto_('1.5') - 2) < 0.6);

console.log('\n── Kelly: integridad del cuestionario ──');
const Q = vm.runInContext('REC_CUALIFICACION_KELLY', ctx);
tt('son 5 preguntas, no 10',   Q.length === 5);
tt('todas tienen ES y EN',     Q.every(q => q.pregunta_es && q.pregunta_en));
tt('el orden es 1..5',         JSON.stringify(Q.map(q=>q.orden)) === '[1,2,3,4,5]');
tt('NO se pregunta por el split — contradiría la promesa del día 28',
   !Q.some(q => /split|comisi[oó]n|cu[aá]nto te quedas|cu[aá]nto ganas|what you keep/i
                 .test(q.pregunta_es + ' ' + q.pregunta_en)));
tt('el dolor es pregunta abierta',
   Q.find(q => q.clave === 'dolor').botones.length === 0);
tt('la producción se pregunta antes del cierre',
   Q.find(q => q.clave === 'operaciones').orden < Q.find(q => q.clave === 'cita').orden);
tt('toda pregunta con botones los tiene en los dos idiomas',
   Q.every(q => q.botones.length === q.botones_en.length));

console.log('\n' + '─'.repeat(50));
console.log(fail === 0 ? '✅ ' + ok + ' PRUEBAS PASADAS' : '❌ ' + fail + ' FALLOS de ' + (ok+fail));
process.exit(fail ? 1 : 0);
