/**
 * Pruebas del endpoint del formulario de entrada (recDoPost / recRecibirFormulario).
 *
 * Es código expuesto a internet, así que lo que más se comprueba aquí es lo que
 * NO debe pasar: sin token no escribe, sin consentimiento no escribe, y una baja
 * previa manda sobre cualquier formulario nuevo.
 *
 * Ejecutar:  node test_formulario.js
 */

const fs=require('fs'), vm=require('vm'), path=require('path');

function hoja(cab){
  const d=[cab.slice()];
  const api={ _d:d,
    getLastRow:()=>d.length,
    getMaxColumns:()=>Math.max(cab.length, 26),
    insertColumnsAfter:()=>{}, setFrozenRows:()=>{}, deleteColumns:()=>{},
    setColumnWidth:()=>api, autoResizeColumns:()=>api, activate:()=>api,
    setDataValidation:()=>api, setConditionalFormatRules:()=>api,
    appendRow:(f)=>{ d.push(f.slice()); },
    getDataRange:()=>({getValues:()=>d}),
    getRange:(r,c,nr,nc)=>{
      nr = nr===undefined?1:nr; nc = nc===undefined?1:nc;
      const o={
        getValues:()=>{const out=[];for(let i=0;i<nr;i++){const f=d[r-1+i]||[];const g=[];
          for(let j=0;j<nc;j++)g.push(f[c-1+j]!==undefined?f[c-1+j]:'');out.push(g);}return out;},
        getValue:()=>{const f=d[r-1]||[];return f[c-1]!==undefined?f[c-1]:'';},
        setValues:(v)=>{v.forEach((f,i)=>{d[r-1+i]=d[r-1+i]||[];f.forEach((x,j)=>{d[r-1+i][c-1+j]=x;});});return o;},
        setValue:(v)=>{d[r-1]=d[r-1]||[];d[r-1][c-1]=v;return o;},
        setBackground:()=>o,setFontColor:()=>o,setFontWeight:()=>o,setFontSize:()=>o,
        setWrap:()=>o,setVerticalAlignment:()=>o,setDataValidation:()=>o,clearContent:()=>o,
        setNumberFormat:()=>o };
      return o; }
  };
  return api;
}

const CAB_CAND=['ID','Nombre','Apellidos','Teléfono','Email','LinkedIn','Instagram','Agencia_Actual',
'Cargo','Zona','Idiomas','Perfil','Inmuebles_Activos','Precio_Medio','Precio_Max','Años_Experiencia',
'Fuente','URL_Fuente','Fecha_Captura','Score','Temperatura','Estado','Responsable','Plan_Activo',
'Paso_Plan','Último_Toque','Próximo_Toque','Nº_Toques','Canal_Preferido','Idioma_Pref',
'Consentimiento_WA','Fecha_Consentimiento','Base_Legal','Info_RGPD_Enviada','ID_CommandMC','Notas'];

const hojas={
  'Rec_Candidatos': hoja(CAB_CAND),
  'Rec_Consentimientos': hoja(['ID','ID_Candidato','Nombre','Fecha','Canal','Texto_Exacto_Mostrado','Version','Idioma','Origen','Finalidades_Aceptadas']),
  'Rec_Toques': hoja(['ID_Toque','ID_Candidato','Nombre','Fecha','Hora','Canal','Paso_Plan','Tipo','Estado_Envío','Resultado','Mensaje','Usuario','Notas']),
  'Rec_Supresion': hoja(['Identificador','Tipo','Nombre','Fecha','Motivo','Origen','Usuario']),
  'Rec_Config': hoja(['Clave','Valor','Descripción'])
};
hojas['Rec_Config']._d.push(['MC_EMAIL','marbella@kwspain.es','']);

let correos=[];
const props={FORM_TOKEN:'secreto-largo-y-aleatorio'};
const ctx={console,
  SpreadsheetApp:{getActiveSpreadsheet:()=>({getSheetByName:n=>hojas[n]||null,insertSheet:n=>(hojas[n]=hoja([])),deleteSheet:()=>{},toast:()=>{}}),
    getUi:()=>({alert:()=>{},ButtonSet:{OK:1,YES_NO:2},Button:{YES:1,OK:1},prompt:()=>({getSelectedButton:()=>0,getResponseText:()=>''})}),
    newDataValidation:()=>({requireValueInList:()=>({setAllowInvalid:()=>({build:()=>({})})})}),
    newConditionalFormatRule:()=>({whenTextEqualTo:()=>({setBackground:()=>({setFontColor:()=>({setRanges:()=>({build:()=>({})})})})})})},
  PropertiesService:{getScriptProperties:()=>({getProperty:k=>props[k]||null,setProperty:(k,v)=>{props[k]=v;}})},
  CacheService:{getScriptCache:()=>({get:()=>null,put:()=>{}})},
  ContentService:{createTextOutput:t=>({_t:t,setMimeType:function(){return this;}}),MimeType:{JSON:'json'}},
  Utilities:{getUuid:()=>Math.random().toString(16).slice(2,10),formatDate:()=>'12:00',sleep:()=>{},
    base64EncodeWebSafe:s=>s,parseCsv:t=>t.trim().split('\n').map(l=>l.split(','))},
  Session:{getScriptTimeZone:()=>'Europe/Madrid',getActiveUser:()=>({getEmail:()=>'tl@kw.es'}),getEffectiveUser:()=>({getEmail:()=>'tl@kw.es'})},
  UrlFetchApp:{fetch:()=>({getResponseCode:()=>200,getContentText:()=>'{}'})},
  Logger:{log:()=>{}},MailApp:{sendEmail:o=>{correos.push(o);}},
  HtmlService:{createHtmlOutput:()=>({setWidth:()=>({setHeight:()=>({})})})},
  DriveApp:{},ScriptApp:{getProjectTriggers:()=>[]},MimeType:{CSV:'csv'}};
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(__dirname,'reclutamiento.gs'),'utf8'), ctx);

let ok=0,fail=0;
const t=(n,r,e)=>{const p=JSON.stringify(r)===JSON.stringify(e);
  if(p){ok++;console.log('  ✅ '+n);}else{fail++;console.log('  ❌ '+n+'\n       obtenido '+JSON.stringify(r)+'\n       esperado '+JSON.stringify(e));}};
const tt=(n,c)=>{if(c){ok++;console.log('  ✅ '+n);}else{fail++;console.log('  ❌ '+n);}};
const post=o=>JSON.parse(ctx.recDoPost({postData:{contents:JSON.stringify(o)}})._t);
const C=vm.runInContext('REC_COL',ctx);
const cand=()=>hojas['Rec_Candidatos']._d;

const BASE={token:'secreto-largo-y-aleatorio',nombre:'Marta',apellidos:'Ruiz',
  telefono:'600 111 222',email:'Marta@Agencia.ES',agencia:'Agencia X',zona:'Nueva Andalucía',
  consiente_contacto:true,consiente_whatsapp:true,operaciones:'16-30',
  honorarios_medios:'20.000-40.000 €',dolor:'no tengo apoyo para captar',
  texto_consentimiento:'Acepto que me contacten...',version_consentimiento:'v1-2026-10',
  campana:'calculadora-linkedin',origen:'https://kw.es/calculadora',idioma:'ES'};

console.log('\n── Seguridad del endpoint ──');
t('sin cuerpo falla',      post && JSON.parse(ctx.recDoPost({})._t).ok, false);
t('token incorrecto falla', post(Object.assign({},BASE,{token:'malo'})).ok, false);
t('mensaje de no autorizado', post(Object.assign({},BASE,{token:'malo'})).error, 'No autorizado');
tt('token incorrecto no escribe nada', cand().length === 1);
t('sin consentimiento falla', post(Object.assign({},BASE,{consiente_contacto:false})).ok, false);
tt('sin consentimiento no escribe nada', cand().length === 1);
t('sin teléfono ni email falla',
  post(Object.assign({},BASE,{telefono:'',email:''})).ok, false);

console.log('\n── Alta correcta ──');
const r1 = post(BASE);
t('devuelve ok',       r1.ok, true);
t('prioridad de llamada A por ser entrante', r1.temperatura, 'A');
t('el encaje real se reporta aparte', r1.encaje, 'B');
t('whatsapp consentido', r1.consentimiento_whatsapp, true);
tt('se ha creado la ficha', cand().length === 2);
const f = cand()[1];
t('teléfono normalizado',  f[C.TELEFONO-1], '+34600111222');
t('email en minúsculas',   f[C.EMAIL-1], 'marta@agencia.es');
tt('base legal es CONSENTIMIENTO, no interés legítimo',
   /6\.1\.a/.test(f[C.BASE_LEGAL-1]) && /[Cc]onsentimiento/.test(f[C.BASE_LEGAL-1]));
t('info RGPD marcada',     f[C.INFO_RGPD-1], 'SI');
t('consentimiento WA',     f[C.CONSENT_WA-1], 'SI');
t('estado conversación activa', f[C.ESTADO-1], 'Conversación activa');
tt('producción convertida del rango', Number(f[C.INMUEBLES-1]) === 23);
tt('honorarios convertidos', Number(f[C.PRECIO_MEDIO-1]) === 30000);
tt('el dolor queda en notas', String(f[C.NOTAS-1]).indexOf('apoyo para captar') !== -1);
tt('la campaña queda en la fuente', String(f[C.FUENTE-1]).indexOf('calculadora-linkedin') !== -1);
tt('temperatura A en la hoja', f[C.TEMPERATURA-1] === 'A');
tt('la nota explica que la A es por entrante, no por encaje',
   /Prioridad A por ser ENTRANTE/.test(String(f[C.NOTAS-1])));
tt('el score NO se infla: sigue siendo el de producción',
   Number(f[C.SCORE-1]) < 75);

console.log('\n── Prueba documental del consentimiento ──');
const con = hojas['Rec_Consentimientos']._d;
tt('se ha guardado una fila de consentimiento', con.length === 2);
tt('guarda el texto exacto que vio', String(con[1][5]).indexOf('Acepto que me contacten') !== -1);
t('guarda la versión',  con[1][6], 'v1-2026-10');
t('guarda el idioma',   con[1][7], 'ES');
tt('guarda el origen',  String(con[1][8]).indexOf('kw.es/calculadora') !== -1);
tt('registra las finalidades', String(con[1][9]).indexOf('WhatsApp') !== -1);

console.log('\n── Auditoría y aviso ──');
tt('queda registrado el toque entrante', hojas['Rec_Toques']._d.length === 2);
tt('avisa por email a la TL', correos.length === 1);
tt('el aviso dice que llame hoy', /llámale hoy/i.test(correos[0].htmlBody));

console.log('\n── Reenvío del formulario no duplica ──');
const antes = cand().length;
const r2 = post(BASE);
t('vuelve a devolver ok', r2.ok, true);
tt('no crea una ficha nueva', cand().length === antes);

console.log('\n── Una baja previa manda sobre el formulario ──');
hojas['Rec_Supresion']._d.push(['+34699888777','Teléfono','Pedro',new Date(),'baja','Manual','']);
const r3 = post(Object.assign({},BASE,{telefono:'699888777',email:'pedro@otra.es',nombre:'Pedro'}));
t('rechaza a quien pidió la baja', r3.ok, false);
tt('lo dice claramente', /baja/i.test(r3.error));
tt('no lo da de alta', !cand().some(x => x[C.NOMBRE-1] === 'Pedro'));

console.log('\n── Texto de consentimiento ──');
const txES = ctx.recTextoConsentimiento('ES'), txEN = ctx.recTextoConsentimiento('EN');
tt('ES cita el 6.1.a',      /6\.1\.a RGPD/.test(txES.contacto));
tt('ES menciona la AEPD',   /AEPD/.test(txES.contacto));
tt('ES ofrece la baja',     /BAJA/.test(txES.contacto));
tt('ES deja hueco a la razón social', /RAZÓN SOCIAL/.test(txES.contacto));
tt('EN cita el GDPR',       /6\.1\.a GDPR/.test(txEN.contacto));
tt('ambos llevan versión',  !!txES.version && txES.version === txEN.version);

console.log('\n'+'─'.repeat(52));
console.log(fail===0?'✅ '+ok+' COMPROBACIONES CORRECTAS':'❌ '+fail+' FALLOS de '+(ok+fail));
process.exit(fail?1:0);
