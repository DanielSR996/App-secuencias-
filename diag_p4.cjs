const XLSX = require('xlsx');
console.log('Leyendo...');
const wb = XLSX.readFile('c:/Users/LCK_KATHIA/Desktop/PRUEBA 1 (4).xlsx', {
  cellFormula:false,cellHTML:false,cellNF:false,cellStyles:false,
  sheetStubs:false,bookImages:false,bookVBA:false,
  sheets:['DS','Layout']
});

// Truncar rango a solo primeras 100 columnas para que sea rápido
function limitCols(sheet, maxCol) {
  const r = XLSX.utils.decode_range(sheet['!ref']);
  r.e.c = Math.min(r.e.c, maxCol - 1);
  sheet['!ref'] = XLSX.utils.encode_range(r);
}
limitCols(wb.Sheets['DS'],     100);
limitCols(wb.Sheets['Layout'], 100);

const dsRaw = XLSX.utils.sheet_to_json(wb.Sheets['DS'],     {header:1, raw:true, defval:''});
const lyRaw = XLSX.utils.sheet_to_json(wb.Sheets['Layout'], {header:1, raw:true, defval:''});
const dsHdr = dsRaw[0].map(c=>String(c).trim());
const lyHdr = lyRaw[0].map(c=>String(c).trim());
console.log('DS filas:', dsRaw.length-1, '| LY filas:', lyRaw.length-1);
console.log('DS headers (primeras 30):', dsHdr.slice(0,30).join(' | '));
console.log('LY headers (primeras 50):', lyHdr.slice(0,50).join(' | '));

const iDsCant = dsHdr.indexOf('CantidadUMComercial');
const iDsVal  = dsHdr.findIndex(h=>h==='Valor usd redondeado');
const iDsFrac = dsHdr.indexOf('Fraccion');
const iDsSec  = dsHdr.indexOf('SecuenciaFraccion');
const iDsDesc = dsHdr.indexOf('DescripcionMercancia');
const iDsPais = dsHdr.indexOf('PaisOrigenDestino');
const iDsPed2 = dsHdr.indexOf('Pedimento2');

const iLyCant = lyHdr.indexOf('cantidad_comercial');
const iLyVal  = lyHdr.indexOf('=REDONDEAR(FR1,0)');
const iLyFrac = lyHdr.indexOf('FraccionNico');
const iLySec  = lyHdr.indexOf('SEC CALC');
const iLyDesc = lyHdr.indexOf('descripcion');
const iLyPais = lyHdr.indexOf('pais_origen PARA LAYOUT');
const iLyPed  = lyHdr.indexOf('pedimento');

console.log('\nIndices DS:', {iDsCant,iDsVal,iDsFrac,iDsSec,iDsPed2,iDsPais,iDsDesc});
console.log('Indices LY:', {iLyCant,iLyVal,iLyFrac,iLySec,iLyPed,iLyPais,iLyDesc});

const nFrac = s => String(s||'').trim().replace(/\./g,'');
const nStr  = s => String(s||'').trim().toUpperCase();
const nDesc = s => String(s||'').trim().toLowerCase().replace(/\s+/g,' ');

// Totales globales
let dsTotalC=0,dsTotalV=0,lyTotalC=0,lyTotalV=0;
for(let i=1;i<dsRaw.length;i++){const r=dsRaw[i];const c=parseFloat(r[iDsCant])||0;const v=parseFloat(r[iDsVal])||0;if(c>0){dsTotalC+=c;dsTotalV+=v;}}
for(let i=1;i<lyRaw.length;i++){const r=lyRaw[i];const c=parseFloat(r[iLyCant])||0;if(!c)continue;lyTotalC+=c;lyTotalV+=parseFloat(r[iLyVal])||0;}
console.log('\n=== TOTALES ===');
console.log('DS  Cant:', dsTotalC.toLocaleString(), 'Val:', dsTotalV.toFixed(0));
console.log('LY  Cant:', lyTotalC.toLocaleString(), 'Val:', lyTotalV.toFixed(0));
console.log('Diff Cant:', lyTotalC-dsTotalC, '| Diff Val:', (lyTotalV-dsTotalV).toFixed(0));

// Agrupar Layout por Ped+Frac
const lyByPF = new Map();
for(let i=1;i<lyRaw.length;i++){
  const r=lyRaw[i];
  const c=parseFloat(r[iLyCant])||0; if(!c) continue;
  const k=nStr(r[iLyPed])+'|||'+nFrac(r[iLyFrac]);
  if(!lyByPF.has(k)) lyByPF.set(k,{cant:0,val:0,n:0,descs:[]});
  const g=lyByPF.get(k); g.cant+=c; g.val+=parseFloat(r[iLyVal])||0; g.n++;
  const d=nDesc(r[iLyDesc]||'').slice(0,25); if(!g.descs.includes(d)) g.descs.push(d);
}

// Agrupar DS por Ped2+Frac
const dsByPF = new Map();
for(let i=1;i<dsRaw.length;i++){
  const r=dsRaw[i];
  const c=parseFloat(r[iDsCant])||0; if(!c) continue;
  const k=nStr(r[iDsPed2])+'|||'+nFrac(r[iDsFrac]);
  if(!dsByPF.has(k)) dsByPF.set(k,[]);
  dsByPF.get(k).push({cant:c,val:parseFloat(r[iDsVal])||0,sec:String(r[iDsSec]||''),desc:nDesc(r[iDsDesc]||''),pais:nStr(r[iDsPais]||''),frac:nFrac(r[iDsFrac]||''),ped2:nStr(r[iDsPed2]||'')});
}

console.log('\n=== GRUPOS DS CON CANT DISTINTA AL LAYOUT ===');
let tot=0;
for(const [k,dsSecs] of dsByPF){
  const ly=lyByPF.get(k);
  const dsSumC=dsSecs.reduce((a,r)=>a+r.cant,0);
  const lySumC=ly?ly.cant:0;
  const diffC=Math.abs(lySumC-dsSumC);
  if(diffC>1||!ly){
    tot++;
    if(tot<=80){
      const f=dsSecs[0].frac,p=dsSecs[0].ped2.slice(-8);
      const info=ly?`LY=${lySumC.toLocaleString()} (${ly.n}filas) diffC=${diffC} lyDescs=[${ly.descs.slice(0,2).join(',')}]`:'SIN FILAS EN LAYOUT';
      console.log(`[${tot}] Frac ${f} (${p}) nsecs=${dsSecs.length} DS=${dsSumC.toLocaleString()} ${info}`);
      if(dsSecs.length<=5) for(const s of dsSecs) console.log(`     Sec${s.sec} cant=${s.cant.toLocaleString()} pais=${s.pais} desc="${s.desc.slice(0,35)}"`);
    }
  }
}
console.log('\nTotal grupos con diff:', tot);
console.log('Total grupos DS:', dsByPF.size, '| Total grupos LY:', lyByPF.size);
