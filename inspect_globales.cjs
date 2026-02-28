/**
 * Análisis completo de concordancia cantidades/valores por secuencia
 * entre DS y Layout — para consultor de comercio exterior.
 */
const XLSX = require('xlsx');
const nH   = s => String(s??'').trim().toLowerCase().replace(/[\s_\-]/g,'');
const nFrac = v => String(v??'').trim().replace(/^0+/,'') || '0';
const ns    = v => String(v??'').trim();

const wb = XLSX.readFile('c:/Users/LCK_KATHIA/Desktop/reducido.xlsx', {cellStyles:false});

// ── DS ─────────────────────────────────────────────────────────────────────────
const dsWs = wb.Sheets['DS'];
const dsAll = XLSX.utils.sheet_to_json(dsWs, {header:1, defval:''});
const dsH   = dsAll[0].map(c => String(c??'').trim());

const fD = (...als) => { for(const a of als){ const i=dsH.findIndex(h=>nH(h)===nH(a)); if(i>=0)return i; } return -1; };
const dCI = {
  ped2:  fD('Pedimento2'),  frac:  fD('Fraccion'),
  sec:   fD('SecuenciaFraccion'), desc: fD('DescripcionMercancia'),
  cant:  fD('CantidadUMComercial'), pais: fD('PaisOrigenDestino'),
  val:   fD('ValorDolares','Valor usd redondeado','Valor Aduana Estadístico'),
  cand:  fD('Candado 551','Candado DS 551'),
  estado: fD('ESTADO'), revisado: fD('REVISADO'),
};

const dsData = [];
for(let i=1;i<dsAll.length;i++){
  const r=dsAll[i]; if(!r||r.every(c=>c===''||c==null)) continue;
  dsData.push({
    _rowI:i, Ped2:ns(r[dCI.ped2]), Frac:ns(r[dCI.frac]), Sec:ns(r[dCI.sec]),
    Desc:ns(r[dCI.desc]), Pais:ns(r[dCI.pais]),
    Cant:parseFloat(r[dCI.cant])||0, Val:parseFloat(r[dCI.val])||0,
    Candado:ns(r[dCI.cand]), Estado:ns(r[dCI.estado]),
  });
}

// ── Layout ─────────────────────────────────────────────────────────────────────
const lyWs = wb.Sheets['Layout'];
const lyAll = XLSX.utils.sheet_to_json(lyWs, {header:1, defval:''});
const KNOWN = new Set(['pedimento','fraccionnico','seccalc','descripcion','paisorigen','pais_origen','valormpdolares','cantidadcomercial','cantidad_comercial','notas','estado']);
let hI=0,bH=0;
for(let i=0;i<Math.min(10,lyAll.length);i++){
  const h=lyAll[i].filter(c=>KNOWN.has(nH(String(c??'')))).length;
  if(h>bH){bH=h;hI=i;}
}
const lyH = lyAll[hI].map(c=>String(c??'').trim());
const fF=(...ns2)=>{for(const n of ns2){const i=lyH.findIndex(h=>nH(h)===nH(n));if(i>=0)return i;}return -1;};
const fL=(...ns2)=>{const ts=ns2.map(nH);return lyH.reduce((l,h,i)=>ts.includes(nH(h))?i:l,-1);};
const lCI = {
  ped: fL('pedimento'), frac: fL('FraccionNico','fraccionnico'),
  sec: fF('SEC CALC','seccalc'), pais: fF('pais_origen','paisorigen'),
  cant: fL('cantidad_comercial','cantidadcomercial'),
  val:  fF('ValorMPDolares','valormpdolares','valordolares','valor_me','valorme'),
  desc: fF('descripcion','clase_descripcion'),
  notasIn: fF('NOTAS','notas'), estado: fF('ESTADO','estado'),
};

const lyData = [];
for(let i=hI+1;i<lyAll.length;i++){
  const r=lyAll[i];
  const ped=ns(r[lCI.ped]??''), frac=ns(r[lCI.frac]??'');
  if(!ped&&!frac) continue;
  const notasV=ns(r[lCI.notasIn]??'').toUpperCase();
  const secV = ns(r[lCI.sec]??'');
  lyData.push({
    _rowI:i+hI, Ped:ped, Frac:frac, Sec:secV,
    Pais:ns(r[lCI.pais]??''), Desc:ns(r[lCI.desc]??''),
    Cant:parseFloat(r[lCI.cant]??0)||0,
    Val:parseFloat(r[lCI.val]??0)||0,
    noIncluir:notasV.includes('NO INCLUIR'),
    secReal:secV!==''&&secV!=='.'&&!isNaN(parseFloat(secV)),
  });
}

console.log(`\n${'='.repeat(70)}`);
console.log('ANÁLISIS GLOBAL — reducido.xlsx');
console.log(`${'='.repeat(70)}`);

// ── Comparación por SECUENCIA ─────────────────────────────────────────────────
console.log('\n── DETALLE POR SECUENCIA (DS vs Layout) ────────────────────────────────');

// Totales del Layout agrupados por Pedimento + Fraccion + Secuencia
const lyBySec = new Map();
for(const r of lyData){
  if(r.noIncluir) continue;
  const sec = r.secReal ? r.Sec : '__SIN_SEC__';
  const k = `${r.Ped}|||${nFrac(r.Frac)}|||${sec}`;
  if(!lyBySec.has(k)) lyBySec.set(k,{cant:0,val:0,rows:0,paises:new Set(),descs:new Set(),sec});
  const g=lyBySec.get(k);
  g.cant+=r.Cant; g.val+=r.Val; g.rows++;
  g.paises.add(r.Pais); g.descs.add(r.Desc.slice(0,30));
}

// Construir candados DS
const dsByCandado = new Map();
for(const r of dsData) if(r.Candado) dsByCandado.set(r.Candado, r);

let totalDSCant=0, totalDSVal=0;
let totalLyCant=0, totalLyVal=0;

for(const ds of dsData){
  const sec = ns(ds.Sec);
  const frac = nFrac(ns(ds.Frac));
  const k = `${ds.Ped2}|||${frac}|||${sec}`;
  const g = lyBySec.get(k);

  totalDSCant += ds.Cant;
  totalDSVal  += ds.Val;

  const fmt = n => Number(n).toLocaleString('es-MX',{maximumFractionDigits:0});
  const fmtV= n => '$'+Number(n).toLocaleString('es-MX',{minimumFractionDigits:2,maximumFractionDigits:2});

  console.log(`\nSec ${ds.Sec} | Frac ${ds.Frac} | ${ds.Estado}`);
  console.log(`  Descripción DS: ${ds.Desc.slice(0,50)}`);
  console.log(`  DS  : Cant=${fmt(ds.Cant)} | Valor=${fmtV(ds.Val)} | Pais=${ds.Pais}`);
  if(g){
    const diffC=g.cant-ds.Cant, diffV=g.val-ds.Val;
    const pctC= ds.Cant>0?(diffC/ds.Cant*100).toFixed(1):'∞';
    const pctV= ds.Val>0?(diffV/ds.Val*100).toFixed(1):'∞';
    const statusC = Math.abs(diffC)<=1 ? '✓' : diffC>0?'▲ LAYOUT MAYOR':'▼ LAYOUT MENOR';
    const statusV = Math.abs(diffV)<=2 ? '✓' : diffV>0?'▲ LAYOUT MAYOR':'▼ LAYOUT MENOR';
    console.log(`  Lay : Cant=${fmt(g.cant)} | Valor=${fmtV(g.val)} | Países=[${[...g.paises].join(',')}] | ${g.rows} filas`);
    console.log(`  Diff: Cant=${diffC>0?'+':''}${fmt(diffC)} (${pctC}%) ${statusC} | Valor=${diffV>0?'+':''}${fmtV(diffV)} (${pctV}%) ${statusV}`);
    totalLyCant += g.cant; totalLyVal += g.val;
  } else {
    console.log(`  Lay : *** SIN FILAS LAYOUT PARA ESTA SECUENCIA ***`);
  }
}

// Filas Layout sin secuencia (a asignar)
const sinSec = lyBySec.get([...lyBySec.keys()].find(k=>k.includes('__SIN_SEC__'))||'__X__');
if(sinSec) {
  console.log(`\n[Sin asignar] Frac ??? : ${sinSec.rows} filas | Cant=${sinSec.cant.toLocaleString()} | Val=$${sinSec.val.toFixed(2)}`);
}

// Grupos sin sec separados por fracción
console.log('\n── GRUPOS LAYOUT SIN SECUENCIA ASIGNADA ────────────────────────────────');
const sinSecMap = new Map();
for(const r of lyData){
  if(r.noIncluir||r.secReal) continue;
  const k=`${r.Ped}|||${nFrac(r.Frac)}|||${r.Pais}`;
  if(!sinSecMap.has(k)) sinSecMap.set(k,{cant:0,val:0,rows:0,frac:r.Frac,ped:r.Ped,pais:r.Pais});
  const g=sinSecMap.get(k); g.cant+=r.Cant; g.val+=r.Val; g.rows++;
}
if(sinSecMap.size===0){
  console.log('  Todas las filas tienen secuencia asignada ✓');
} else {
  for(const [k,g] of sinSecMap){
    console.log(`  Frac ${g.frac} | País ${g.pais}: ${g.rows} filas | Cant=${g.cant.toLocaleString()} | Val=$${g.val.toFixed(2)}`);
    // Buscar candidatos DS para esta fracción
    const cands = dsData.filter(d=>nFrac(d.Frac)===nFrac(g.frac));
    if(cands.length===0){
      console.log(`    → Sin entrada DS para fracción ${g.frac}`);
    } else {
      for(const c of cands){
        const diffC=g.cant-c.Cant, diffV=g.val-c.Val;
        console.log(`    → DS Sec ${c.Sec}: Cant=${c.Cant.toLocaleString()} (dif ${diffC>0?'+':''}${diffC.toLocaleString()}) | Val=$${c.Val.toFixed(2)} (dif ${diffV>0?'+':''}$${diffV.toFixed(2)}) [${c.Estado}]`);
      }
    }
  }
}

const fmt = n=>Number(n).toLocaleString('es-MX',{maximumFractionDigits:0});
const fmtV= n=>'$'+Number(n).toLocaleString('es-MX',{minimumFractionDigits:2,maximumFractionDigits:2});
console.log('\n── TOTALES GLOBALES ────────────────────────────────────────────────────');
console.log(`  DS  total : Cant=${fmt(totalDSCant)} | Valor=${fmtV(totalDSVal)}`);
console.log(`  Layout(sec): Cant=${fmt(totalLyCant)} | Valor=${fmtV(totalLyVal)}`);
const gDC=totalLyCant-totalDSCant, gDV=totalLyVal-totalDSVal;
console.log(`  Diferencia : Cant=${gDC>0?'+':''}${fmt(gDC)} | Valor=${gDV>0?'+':''}\$${gDV.toFixed(2)}`);

// Total layout completo (incluyendo sin sec)
const totalAllLyCant = lyData.filter(r=>!r.noIncluir).reduce((a,r)=>a+r.Cant,0);
const totalAllLyVal  = lyData.filter(r=>!r.noIncluir).reduce((a,r)=>a+r.Val,0);
console.log(`  Layout total activo: Cant=${fmt(totalAllLyCant)} | Valor=${fmtV(totalAllLyVal)}`);
console.log(`  Layout NO INCLUIR: ${lyData.filter(r=>r.noIncluir).length} filas`);
