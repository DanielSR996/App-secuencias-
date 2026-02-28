/**
 * Simula exactamente el flujo de la App2020 con reducido.xlsx
 * y verifica si E0 encuentra matches.
 */
const XLSX = require('xlsx');
const nH2020 = s => String(s??'').trim().toLowerCase().replace(/[\s_\-]/g,'');
const nFrac  = v => String(v??'').trim().replace(/^0+/,'') || '0';
const normStr= v => String(v??'').trim();
const isRealSec = v => { const s=normStr(v); return s!==''&&s!=='.'&&!isNaN(parseFloat(s)); };

const wb = XLSX.readFile('c:/Users/LCK_KATHIA/Desktop/reducido.xlsx', {cellStyles:false});
console.log('Hojas:', wb.SheetNames);

// ── Leer DS ───────────────────────────────────────────────────────────────────
const dsWs = wb.Sheets['DS'];
const dsAllRows = XLSX.utils.sheet_to_json(dsWs, {header:1, defval:''});
const dsHdr = dsAllRows[0].map(c => String(c??'').trim());
const dsIdx = {};
[['Pedimento2','Pedimento2'],['Fraccion','Fraccion'],['SecuenciaFraccion','SecuenciaFraccion'],
 ['DescripcionMercancia','DescripcionMercancia'],['CantidadUMComercial','CantidadUMComercial'],
 ['ValorDolares',['ValorDolares','Valor usd redondeado']],
 ['PaisOrigenDestino','PaisOrigenDestino'],['Candado 551',['Candado 551','Candado DS 551']]
].forEach(([key, aliases]) => {
  const als = Array.isArray(aliases) ? aliases : [aliases];
  for (const a of als) {
    const i = dsHdr.findIndex(h => nH2020(h)===nH2020(a));
    if(i>=0){dsIdx[key]=i;break;}
  }
});
console.log('\n[DS] colIdx:', JSON.stringify(dsIdx));
const dsRows = [];
for (let i=1;i<dsAllRows.length;i++) {
  const r=dsAllRows[i];
  if(!r||r.every(c=>c===''||c==null)) continue;
  const obj={_dsIdx:dsRows.length};
  for(const [k,ci] of Object.entries(dsIdx)) obj[k]=r[ci]??'';
  dsRows.push(obj);
}
console.log('[DS] filas:', dsRows.length);
dsRows.forEach((r,i)=>console.log(`  DS[${i}]: Ped2=${r.Pedimento2} Frac=${r.Fraccion} Sec=${r.SecuenciaFraccion} Cant=${r.CantidadUMComercial} Val=${r.ValorDolares} Candado="${r['Candado 551']}"`));

// ── Leer Layout ────────────────────────────────────────────────────────────────
const lyWs = wb.Sheets['Layout'];
const lyAllRows = XLSX.utils.sheet_to_json(lyWs, {header:1, defval:''});
// Detectar header row
const KNOWN = new Set(['pedimento','fraccionnico','seccalc','descripcion','paisorigen','pais_origen','valormpdolares','cantidadcomercial','cantidad_comercial','notas','estado']);
let hdrI=0, bHits=0;
for(let i=0;i<Math.min(10,lyAllRows.length);i++){
  const h=lyAllRows[i].filter(c=>KNOWN.has(nH2020(String(c??'')))).length;
  if(h>bHits){bHits=h;hdrI=i;}
}
const lyHdr = lyAllRows[hdrI].map(c=>String(c??'').trim());
const findFirst=(...names)=>{for(const n of names){const i=lyHdr.findIndex(h=>nH2020(h)===nH2020(n));if(i>=0)return i;}return -1;};
const findLast=(...names)=>{const ts=names.map(nH2020);return lyHdr.reduce((l,h,i)=>ts.includes(nH2020(h))?i:l,-1);};
const ci = {
  pedimento: findLast('pedimento'),
  frac:      findLast('FraccionNico','fraccionnico'),
  cant:      findLast('cantidad_comercial','cantidadcomercial'),
  val:       findFirst('ValorMPDolares','valormpdolares','valordolares','valor_me','valorme'),
  sec:       findFirst('SEC CALC','seccalc'),
  notasIn:   findFirst('NOTAS','notas'),
  notas:     findLast('NOTAS','notas'),
  pais:      findFirst('pais_origen','paisorigen'),
  desc:      findFirst('descripcion','clase_descripcion'),
  estado:    findFirst('ESTADO','estado'),
};
console.log('\n[Layout] hdrI:', hdrI, 'hits:', bHits);
console.log('[Layout] colIdx:', JSON.stringify(ci));
console.log(`  pedimento → col ${ci.pedimento}: "${lyHdr[ci.pedimento]}"`);
console.log(`  frac      → col ${ci.frac}: "${lyHdr[ci.frac]}"`);
console.log(`  sec       → col ${ci.sec}: "${lyHdr[ci.sec]}"`);
console.log(`  val       → col ${ci.val}: "${lyHdr[ci.val]}"`);
console.log(`  cant      → col ${ci.cant}: "${lyHdr[ci.cant]}"`);
console.log(`  notasIn   → col ${ci.notasIn}: "${lyHdr[ci.notasIn]}"`);
console.log(`  notas(out)→ col ${ci.notas}: "${lyHdr[ci.notas]}"`);

// Leer filas Layout
const lyRows=[];
for(let i=hdrI+1;i<lyAllRows.length;i++){
  const r=lyAllRows[i];
  const ped=normStr(r[ci.pedimento]??''), frac=normStr(r[ci.frac]??'');
  if(!ped&&!frac) continue;
  const notasV=normStr(r[ci.notasIn]??'').toUpperCase();
  lyRows.push({
    _idx:lyRows.length, Pedimento:ped, FraccionNico:frac,
    Descripcion:normStr(r[ci.desc]??''),
    PaisOrigen:normStr(r[ci.pais]??''),
    Cantidad:parseFloat(r[ci.cant]??0)||0,
    ValorUSD:parseFloat(r[ci.val]??0)||0,
    SecCalc:normStr(r[ci.sec]??''),
    noIncluir:notasV.includes('NO INCLUIR'),
    secIsReal:isRealSec(normStr(r[ci.sec]??'')),
  });
}
console.log(`\n[Layout] filas: ${lyRows.length}`);
console.log(`  Con sec real: ${lyRows.filter(r=>r.secIsReal).length}`);
console.log(`  NO INCLUIR:   ${lyRows.filter(r=>r.noIncluir).length}`);
console.log(`  A asignar:    ${lyRows.filter(r=>!r.secIsReal&&!r.noIncluir).length}`);

// ── Simular E0 ────────────────────────────────────────────────────────────────
const dsByCandado = new Map();
for(const r of dsRows){
  const c=normStr(r['Candado 551']);
  if(c) dsByCandado.set(c,r);
}
console.log('\n[E0] dsByCandado keys:', [...dsByCandado.keys()]);

let e0ok=0, e0fail=0;
for(const row of lyRows){
  if(row.noIncluir||!row.secIsReal) continue;
  const key=`${row.Pedimento}-${nFrac(row.FraccionNico)}-${row.SecCalc}`;
  const dsRow=dsByCandado.get(key);
  if(dsRow){
    e0ok++;
    if(e0ok<=3) console.log(`  [E0 OK] Layout fila ${row._idx}: key="${key}" → DS Sec=${dsRow.SecuenciaFraccion}`);
  } else {
    e0fail++;
    if(e0fail<=3) console.log(`  [E0 FAIL] Layout fila ${row._idx}: key="${key}" → NO MATCH`);
  }
}
console.log(`\n[E0 RESULTADO] Verificadas: ${e0ok} | Sin match en E0: ${e0fail}`);

// ── Grupos para E1-E5 (los que no pasaron E0) ─────────────────────────────────
const groups = new Map();
for(const row of lyRows){
  if(row.noIncluir) continue;
  if(/* ya verificado */ false) continue; // simplificado
  const key=`${row.Pedimento}|||${nFrac(row.FraccionNico)}|||${normStr(row.PaisOrigen)}`;
  if(!groups.has(key)) groups.set(key,[]);
  groups.get(key).push(row);
}
console.log('\n[E1-E5] Grupos a procesar:');
for(const [k,rows] of groups){
  const sumC=rows.reduce((a,r)=>a+r.Cantidad,0);
  const sumV=rows.reduce((a,r)=>a+r.ValorUSD,0);
  console.log(`  Grupo "${k}": ${rows.length} filas | sumCant=${sumC} | sumVal=${sumV.toFixed(2)}`);
}

// ── Análisis comercio exterior ─────────────────────────────────────────────────
console.log('\n=== ANÁLISIS CONSULTOR COMERCIO EXTERIOR ===');
// Grupos por fracción en DS
const dsByFrac = new Map();
for(const r of dsRows){
  const k=`${r.Pedimento2}|||${nFrac(String(r.Fraccion))}`;
  if(!dsByFrac.has(k)) dsByFrac.set(k,[]);
  dsByFrac.get(k).push(r);
}
// Comparar totales Layout vs DS por fracción
const layoutByFrac = new Map();
for(const r of lyRows){
  const k=`${r.Pedimento}|||${nFrac(r.FraccionNico)}`;
  if(!layoutByFrac.has(k)) layoutByFrac.set(k,{cant:0,val:0,noInc:0,rows:0});
  const g=layoutByFrac.get(k);
  g.cant+=r.Cantidad; g.val+=r.ValorUSD; g.rows++;
  if(r.noIncluir) g.noInc++;
}
console.log('\nComparación Layout vs DS por fracción:');
for(const [k,dsArr] of dsByFrac){
  const[ped,frac]=k.split('|||');
  const lyG=layoutByFrac.get(k)||{cant:0,val:0,noInc:0,rows:0};
  const dsCant=dsArr.reduce((a,r)=>a+(parseFloat(r.CantidadUMComercial)||0),0);
  const dsVal=dsArr.reduce((a,r)=>a+(parseFloat(r.ValorDolares)||0),0);
  const status=lyG.noInc===lyG.rows?'TODAS "NO INCLUIR"':lyG.noInc>0?'MIXTO':'ACTIVO';
  console.log(`  Frac ${frac}: DS=${dsCant}u/$${dsVal} | Layout(${lyG.rows} filas)=${lyG.cant}u/$${lyG.val.toFixed(2)} [${status}]`);
  if(Math.abs(dsCant-lyG.cant)>1||Math.abs(dsVal-lyG.val)>2){
    console.log(`    ⚠ DISCREPANCIA CANTIDAD: DS=${dsCant} vs Layout=${lyG.cant} (diff=${Math.abs(dsCant-lyG.cant)})`);
    console.log(`    ⚠ DISCREPANCIA VALOR USD: DS=${dsVal} vs Layout=${lyG.val.toFixed(2)} (diff=${Math.abs(dsVal-lyG.val).toFixed(2)})`);
  } else {
    console.log(`    ✓ Valores coinciden`);
  }
}
