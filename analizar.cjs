const XLSX = require('./node_modules/xlsx');
const wb = XLSX.readFile('c:/Users/LCK_KATHIA/Downloads/prueba1.xlsx');

const layout = XLSX.utils.sheet_to_json(wb.Sheets['Layout'], {header:1, defval:''});
const s551   = XLSX.utils.sheet_to_json(wb.Sheets['551'],    {header:1, defval:''});

const hL = layout[0];
const h5 = s551[0];

const secIdx  = hL.findIndex(h=>String(h).trim()==='SecuenciaPed');
const notasIdx= hL.findIndex(h=>String(h).trim()==='Notas');
const pedIdx  = hL.findIndex(h=>String(h).trim()==='Pedimento');
const fracIdx = hL.findIndex(h=>String(h).trim()==='FraccionNico');
const vcusdIdx= hL.findIndex(h=>String(h).trim()==='VCUSD');
const cantIdx = hL.findIndex(h=>String(h).trim()==='CantidadSaldo');
const paisIdx = hL.findIndex(h=>String(h).trim()==='PaisOrigen');

console.log('=== COLUMNAS LAYOUT ===');
console.log('SecuenciaPed col:', secIdx, '| VCUSD col:', vcusdIdx, '| CantidadSaldo col:', cantIdx);
console.log('Total 551 filas (sin header):', s551.length - 1);

// Índices en 551
const sPedIdx  = h5.findIndex(h=>String(h).trim()==='Pedimento');
const sFracIdx = h5.findIndex((h,i)=>{ // primera ocurrencia de Fraccion
  return String(h).trim()==='Fraccion' && i > 0;
});
const sSeqIdx  = h5.findIndex(h=>String(h).trim()==='SecuenciaFraccion');
const sCantIdx = h5.findIndex(h=>String(h).trim()==='CantidadUMComercial');
const sValIdx  = h5.findIndex(h=>String(h).trim()==='ValorDolares');
const sPaisIdx = h5.findIndex(h=>String(h).trim()==='PaisOrigenDestino');

console.log('551 -- Pedimento col:', sPedIdx, '| Fraccion col:', sFracIdx,
            '| SecuenciaFraccion col:', sSeqIdx,
            '| CantidadUMComercial col:', sCantIdx, '| ValorDolares col:', sValIdx);

// Contar Layout sin SecuenciaPed
let sinSec=0, conSec=0;
let ejemplosSin=[];
let vcusdZero=0, vcusdBlank=0;

for(let i=1;i<layout.length;i++){
  const row = layout[i];
  if(!row || row.every(c=>c===''||c==null)) continue;
  const sec  = row[secIdx];
  const nota = row[notasIdx];
  const vcusd= row[vcusdIdx];
  if(!sec || sec===''){
    sinSec++;
    if(ejemplosSin.length < 8)
      ejemplosSin.push({
        ped:  row[pedIdx],
        frac: row[fracIdx],
        pais: row[paisIdx],
        cant: row[cantIdx],
        vcusd,
        nota: String(nota||'').substring(0,100)
      });
  } else {
    conSec++;
  }
  if(vcusd===0 || vcusd==='0') vcusdZero++;
  if(vcusd===''||vcusd==null)  vcusdBlank++;
}
console.log('\n=== LAYOUT ===');
console.log('Con SecuenciaPed:', conSec, '| Sin SecuenciaPed:', sinSec);
console.log('Filas con VCUSD=0:', vcusdZero, '| VCUSD vacío:', vcusdBlank);
console.log('\nEjemplos sin secuencia:');
ejemplosSin.forEach((r,i)=>console.log(i+1+':', JSON.stringify(r)));

// Totales
const totalLayoutCant = layout.slice(1).reduce((a,r)=>a+(parseFloat(r[cantIdx])||0),0);
const totalLayoutVCUSD= layout.slice(1).reduce((a,r)=>a+(parseFloat(r[vcusdIdx])||0),0);
const total551Cant    = s551.slice(1).reduce((a,r)=>a+(parseFloat(r[sCantIdx])||0),0);
const total551Val     = s551.slice(1).reduce((a,r)=>a+(parseFloat(r[sValIdx])||0),0);

console.log('\n=== TOTALES GLOBALES ===');
console.log('Layout CantidadSaldo:', totalLayoutCant.toFixed(2));
console.log('551    CantUMComercial:', total551Cant.toFixed(2));
console.log('Diferencia Cantidad:', (totalLayoutCant - total551Cant).toFixed(2));
console.log('Layout VCUSD:', totalLayoutVCUSD.toFixed(2));
console.log('551    ValorDolares:', total551Val.toFixed(2));
console.log('Diferencia Valor:', (totalLayoutVCUSD - total551Val).toFixed(2));

// Revisar secuencias 551 vs las asignadas en Layout
const sec551Set = new Set();
for(let i=1;i<s551.length;i++){
  const r = s551[i];
  if(!r || r.every(c=>c===''||c==null)) continue;
  sec551Set.add(String(r[sSeqIdx]||'').trim());
}
const secLayoutSet = new Set();
for(let i=1;i<layout.length;i++){
  const r = layout[i];
  if(!r || r.every(c=>c===''||c==null)) continue;
  const s = String(r[secIdx]||'').trim();
  if(s) secLayoutSet.add(s);
}
console.log('\n=== SECUENCIAS ===');
console.log('Secuencias únicas en 551:', sec551Set.size);
console.log('Secuencias únicas en Layout:', secLayoutSet.size);
const en551NoLayout = [...sec551Set].filter(s=>!secLayoutSet.has(s));
console.log('En 551 pero NO en Layout:', en551NoLayout.length, '->', en551NoLayout.slice(0,20).join(', '));
const enLayoutNo551 = [...secLayoutSet].filter(s=>!sec551Set.has(s));
console.log('En Layout pero NO en 551:', enLayoutNo551.length, '->', enLayoutNo551.slice(0,20).join(', '));
