/**
 * Pruebas del importador de agencias, contra el CSV real de agencias_marbella_seed.csv.
 *
 * Simula la hoja de cálculo y el parser de CSV de Apps Script (incluidas las
 * comillas y las comas dentro de un campo, que es donde esto se rompe).
 *
 * Ejecutar:  node test_agencias.js
 */

const fs = require('fs'), vm = require('vm');

// CSV real de Apps Script: respeta comillas y comas dentro de campo
function parseCsv(txt, delim){
  delim = delim || ',';
  const out=[]; let fila=[], campo='', enQ=false;
  for(let i=0;i<txt.length;i++){
    const c=txt[i];
    if(enQ){
      if(c==='"'){ if(txt[i+1]==='"'){campo+='"';i++;} else enQ=false; }
      else campo+=c;
    } else if(c==='"'){ enQ=true; }
    else if(c===delim){ fila.push(campo); campo=''; }
    else if(c==='\n'){ fila.push(campo); out.push(fila); fila=[]; campo=''; }
    else if(c!=='\r'){ campo+=c; }
  }
  if(campo!==''||fila.length){ fila.push(campo); out.push(fila); }
  return out;
}

// Hoja simulada
function hojaFalsa(){
  const datos=[]; const anchos={};
  return {
    _datos: datos,
    getLastRow: ()=> datos.length,
    getRange: (r,c,nr,nc)=>({
      getValues: ()=>{ const o=[]; for(let i=0;i<nr;i++){ const f=datos[r-1+i]||[]; const g=[];
        for(let j=0;j<nc;j++) g.push(f[c-1+j]!==undefined?f[c-1+j]:''); o.push(g);} return o; },
      setValues: (v)=>{ v.forEach((f,i)=>{ datos[r-1+i]=datos[r-1+i]||[];
        f.forEach((x,j)=>{ datos[r-1+i][c-1+j]=x; }); }); return {setWrap:()=>({setVerticalAlignment:()=>{}})}; },
      setWrap: ()=>({setVerticalAlignment:()=>{}})
    }),
    setColumnWidth: (c,w)=>{ anchos[c]=w; return {setColumnWidth:()=>{}}; }
  };
}

const hAg = hojaFalsa();
hAg._datos.push(['ID','Agencia','Web','Teléfono','Email','Dirección','Zona','Google_Rating',
  'Google_Reviews','Agentes_Detectados','Modelo','Competidor_Directo','Prioridad',
  'Fuente','Fecha_Captura','Web_Rastreada','Notas']);

const ctx={console,
  SpreadsheetApp:{getActiveSpreadsheet:()=>({getSheetByName:n=> n==='Rec_Agencias'?hAg:null}),
    getUi:()=>({alert:()=>{}}),newDataValidation:()=>({requireValueInList:()=>({setAllowInvalid:()=>({build:()=>({})})})}),
    newConditionalFormatRule:()=>({whenTextEqualTo:()=>({setBackground:()=>({setFontColor:()=>({setRanges:()=>({build:()=>({})})})})})})},
  PropertiesService:{getScriptProperties:()=>({getProperty:()=>'X'})},
  CacheService:{getScriptCache:()=>({get:()=>null,put:()=>{}})},
  Utilities:{getUuid:()=>Math.random().toString(16).slice(2,10),formatDate:()=>'',sleep:()=>{},
    base64EncodeWebSafe:s=>s, parseCsv:parseCsv},
  Session:{getScriptTimeZone:()=>'Europe/Madrid',getActiveUser:()=>({getEmail:()=>''}),getEffectiveUser:()=>({getEmail:()=>''})},
  UrlFetchApp:{fetch:()=>({getResponseCode:()=>200,getContentText:()=>'{}'})},Logger:{log:()=>{}},
  HtmlService:{createHtmlOutput:()=>({setWidth:()=>({setHeight:()=>({})})})},
  DriveApp:{},MailApp:{},ScriptApp:{},MimeType:{}};
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(require('path').join(__dirname,'reclutamiento.gs'),'utf8'), ctx);

let ok=0, fail=0;
const t=(n,r,e)=>{ const p=JSON.stringify(r)===JSON.stringify(e);
  if(p){ok++;console.log('  ✅ '+n);} else {fail++;console.log('  ❌ '+n+'\n       obtenido '+JSON.stringify(r)+'\n       esperado '+JSON.stringify(e));}};
const tt=(n,c)=>{ if(c){ok++;console.log('  ✅ '+n);} else {fail++;console.log('  ❌ '+n);} };

const csv = fs.readFileSync(require('path').join(__dirname,'agencias_marbella_seed.csv'),'utf8');

console.log('\n── Importar el CSV real de 40 agencias ──');
const r1 = ctx.recImportarAgencias(csv);
t('40 agencias nuevas',        r1.nuevas, 40);
t('ninguna duplicada',         r1.duplicadas, 0);
t('25 con web',                r1.conWeb, 25);
t('15 sin web',                r1.sinWeb, 15);

const d = hAg._datos;
console.log('\n── Integridad de lo escrito ──');
tt('cabecera intacta',         d[0][1] === 'Agencia');
tt('41 filas en total',         d.length === 41);
const terra = d.find(f => f[1] === 'Terra Realty');
tt('Terra Realty presente',     !!terra);
t('web normalizada a https',    terra[2], 'https://terramarbella.com');
t('zona conservada',            terra[6], 'Nueva Andalucía');
t('prioridad conservada',       terra[12], 'Alta');
t('marcada como no rastreada',  terra[15], 'NO');
tt('nota con comas y comillas intacta',
   terra[16].indexOf('Fundada 1992') !== -1 && terra[16].indexOf('Nuestro Equipo') !== -1);

const inmo = d.find(f => f[1] === 'Inmobiliaria Marbella');
tt('nota con comas internas no se parte',
   inmo[16].indexOf('cargos e idiomas') !== -1 && inmo[16].indexOf('los socios no se reclutan') !== -1);

const lucas = d.find(f => f[1] === 'Lucas Fox Marbella');
t('franquicia detectada sola',  lucas[10], 'Gran franquicia');
const sinWeb = d.find(f => f[1] === 'Real Estate Nordica');
t('sin web queda vacío, no inventado', sinWeb[2], '');

console.log('\n── Reimportar la misma lista no duplica ──');
const r2 = ctx.recImportarAgencias(csv);
t('0 nuevas',   r2.nuevas, 0);
t('40 duplicadas', r2.duplicadas, 40);
tt('sigue habiendo 41 filas', hAg._datos.length === 41);

console.log('\n── Tolerancia de formato ──');
const r3 = ctx.recImportarAgencias('nombre,url\nAgencia Nueva SL,www.ejemplo.es\n');
t('acepta cabeceras alternativas', r3.nuevas, 1);
const nueva = hAg._datos.find(f => f[1] === 'Agencia Nueva SL');
t('url sin protocolo se arregla', nueva[2], 'https://www.ejemplo.es');
t('prioridad por defecto',        nueva[12], 'Media');

let err = null;
try { ctx.recImportarAgencias('col1,col2\na,b\n'); } catch(e){ err = e.message; }
tt('CSV sin columna de nombre falla con mensaje claro',
   err && err.indexOf('nombre de la agencia') !== -1);

console.log('\n'+'─'.repeat(52));
console.log(fail===0 ? '✅ '+ok+' COMPROBACIONES CORRECTAS' : '❌ '+fail+' FALLOS de '+(ok+fail));
process.exit(fail?1:0);
