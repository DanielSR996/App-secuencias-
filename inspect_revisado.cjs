/**
 * Analiza reducido.xlsx completo para entender qué motivos de no-match
 * existen y dónde está la columna REVISADO en el DS.
 */
const XLSX = require('xlsx');
const nH   = s => String(s??'').trim().toLowerCase().replace(/[\s_\-]/g,'');
const nFrac = v => String(v??'').trim().replace(/^0+/,'') || '0';
const normStr = v => String(v??'').trim();

const wb = XLSX.readFile('c:/Users/LCK_KATHIA/Desktop/reducido.xlsx', {cellStyles:false});
console.log('Hojas:', wb.SheetNames);

// ── DS ─────────────────────────────────────────────────────────────────────────
const dsWs = wb.Sheets['DS'];
const dsRows = XLSX.utils.sheet_to_json(dsWs, {header:1, defval:''});
const dsHdr = dsRows[0].map(c => String(c??'').trim());
console.log('\n[DS] Headers completos:');
dsHdr.forEach((h,i) => { if(h) console.log(`  [${i}] "${h}"`); });

// Encontrar col REVISADO en DS
const revisadoCol = dsHdr.findIndex(h => nH(h) === 'revisado');
const estadoCol   = dsHdr.findIndex(h => nH(h) === 'estado');
const ped2Col     = dsHdr.findIndex(h => nH(h) === 'pedimento2');
const fracCol     = dsHdr.findIndex(h => nH(h) === 'fraccion');
const secCol      = dsHdr.findIndex(h => nH(h) === 'secuenciafraccion');
const cantCol     = dsHdr.findIndex(h => nH(h) === 'cantidadumc' || nH(h) === 'cantidadumcomercial');
const valCol      = dsHdr.findIndex(h => ['valordolares','valorusdredondeado'].includes(nH(h)));
const candadoCol  = dsHdr.findIndex(h => nH(h) === 'candado551');
console.log(`\n[DS] REVISADO col: ${revisadoCol} | ESTADO col: ${estadoCol} | Pedimento2 col: ${ped2Col} | Fraccion col: ${fracCol}`);
console.log(`[DS] Sec col: ${secCol} | Cant col: ${cantCol} | Val col: ${valCol} | Candado col: ${candadoCol}`);

// DS data rows
const dsData = [];
for(let i=1;i<dsRows.length;i++){
  const r=dsRows[i];
  if(!r||r.every(c=>c===''||c==null)) continue;
  dsData.push({
    _rowI: i,
    Pedimento2: normStr(r[ped2Col]??''),
    Fraccion:   normStr(r[fracCol]??''),
    Sec:        normStr(r[secCol]??''),
    Cant:       parseFloat(r[cantCol]??0)||0,
    Val:        parseFloat(r[valCol]??0)||0,
    Candado:    normStr(r[candadoCol]??''),
    Estado:     normStr(r[estadoCol]??''),
    Revisado:   normStr(r[revisadoCol]??''),
  });
}
console.log('\n[DS] Filas de datos:', dsData.length);
dsData.forEach((r,i) => console.log(`  DS[${i}] row=${r._rowI}: Ped=${r.Pedimento2} Frac=${r.Fraccion} Sec=${r.Sec} Cant=${r.Cant} Val=${r.Val} Candado="${r.Candado}" Estado="${r.Estado}" Revisado="${r.Revisado}"`));

// ── Layout ─────────────────────────────────────────────────────────────────────
const lyWs = wb.Sheets['Layout'];
const lyAllRows = XLSX.utils.sheet_to_json(lyWs, {header:1, defval:''});
const KNOWN = new Set(['pedimento','fraccionnico','seccalc','descripcion','paisorigen','pais_origen','valormpdolares','cantidadcomercial','cantidad_comercial','notas','estado']);
let hdrI=0, bHits=0;
for(let i=0;i<Math.min(10,lyAllRows.length);i++){
  const h=lyAllRows[i].filter(c=>KNOWN.has(nH(String(c??'')))).length;
  if(h>bHits){bHits=h;hdrI=i;}
}
const lyHdr = lyAllRows[hdrI].map(c=>String(c??'').trim());
const fF=(...ns)=>{for(const n of ns){const i=lyHdr.findIndex(h=>nH(h)===nH(n));if(i>=0)return i;}return -1;};
const fL=(...ns)=>{const ts=ns.map(nH);return lyHdr.reduce((l,h,i)=>ts.includes(nH(h))?i:l,-1);};
const ci = {
  ped: fL('pedimento'), frac: fL('FraccionNico','fraccionnico'),
  cant: fL('cantidad_comercial','cantidadcomercial'),
  val:  fF('ValorMPDolares','valormpdolares','valordolares','valor_me','valorme'),
  sec:  fF('SEC CALC','seccalc'),
  notasIn: fF('NOTAS','notas'),
};

const lyData = [];
for(let i=hdrI+1;i<lyAllRows.length;i++){
  const r=lyAllRows[i];
  const ped=normStr(r[ci.ped]??''), frac=normStr(r[ci.frac]??'');
  if(!ped&&!frac) continue;
  const notasV=normStr(r[ci.notasIn]??'').toUpperCase();
  lyData.push({
    Pedimento:ped, FraccionNico:frac,
    Cant:parseFloat(r[ci.cant]??0)||0,
    Val:parseFloat(r[ci.val]??0)||0,
    SecCalc:normStr(r[ci.sec]??''),
    noIncluir:notasV.includes('NO INCLUIR'),
  });
}

// ── Resumen por grupo Layout ────────────────────────────────────────────────────
console.log('\n[Layout] Grupos por Pedimento+Fraccion+Pais:');
const grupos = new Map();
for(const r of lyData){
  const k=`${r.Pedimento}|||${nFrac(r.FraccionNico)}`;
  if(!grupos.has(k)) grupos.set(k,{cant:0,val:0,rows:0,noInc:0});
  const g=grupos.get(k);
  g.cant+=r.Cant; g.val+=r.Val; g.rows++;
  if(r.noIncluir) g.noInc++;
}
for(const [k,g] of grupos){
  const [ped,frac]=k.split('|||');
  console.log(`  Frac ${frac}: ${g.rows} filas | Cant=${g.cant.toLocaleString()} | Val=$${g.val.toFixed(2)} | noInc=${g.noInc}`);
}

// ── Analizar motivos de no-match por DS row ────────────────────────────────────
console.log('\n=== MOTIVOS DE NO-MATCH POR FILA DS ===');
for(const ds of dsData){
  const frac=nFrac(ds.Fraccion);
  const k=`${ds.Pedimento2}|||${frac}`;
  const g=grupos.get(k);
  let motivo = '';
  if(!g){
    // No hay ninguna fila Layout con esta fraccion/pedimento
    motivo = `Fracción ${ds.Fraccion} no encontrada en Layout para pedimento ${ds.Pedimento2}`;
  } else if(g.noInc===g.rows){
    motivo = `Todas las filas Layout (${g.rows}) marcadas NO INCLUIR`;
  } else {
    const diffCant = Math.abs(g.cant - ds.Cant);
    const diffVal  = Math.abs(g.val  - ds.Val);
    const pctCant  = ds.Cant>0 ? (diffCant/ds.Cant*100).toFixed(1) : '∞';
    const pctVal   = ds.Val>0  ? (diffVal/ds.Val*100).toFixed(1)  : '∞';
    motivo = `Cant.Layout=${g.cant.toLocaleString()} vs DS=${ds.Cant.toLocaleString()} (dif ${pctCant}%) | Valor Layout=$${g.val.toFixed(0)} vs DS=$${ds.Val} (dif ${pctVal}%)`;
  }
  console.log(`  DS Sec ${ds.Sec} Frac ${ds.Fraccion}: ${motivo}`);
}

// ── Layout sin DS ─────────────────────────────────────────────────────────────
console.log('\n=== GRUPOS LAYOUT SIN COBERTURA EN DS ===');
const dsFracs = new Set(dsData.map(r=>`${r.Pedimento2}|||${nFrac(r.Fraccion)}`));
for(const [k,g] of grupos){
  if(!dsFracs.has(k)){
    const[ped,frac]=k.split('|||');
    console.log(`  Frac ${frac} (${g.rows} filas, Cant=${g.cant.toLocaleString()}, Val=$${g.val.toFixed(2)}): SIN secuencia en DS`);
  }
}
