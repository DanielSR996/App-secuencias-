const XLSX = require('./node_modules/xlsx');
const wb = XLSX.readFile('c:/Users/LCK_KATHIA/Downloads/prueba1.xlsx');

const layout = XLSX.utils.sheet_to_json(wb.Sheets['Layout'], {header:1, defval:''});
const s551   = XLSX.utils.sheet_to_json(wb.Sheets['551'],    {header:1, defval:''});

const hL = layout[0];
const h5 = s551[0];

const secIdx   = hL.findIndex(h=>String(h).trim()==='SecuenciaPed');
const pedLIdx  = hL.findIndex(h=>String(h).trim()==='Pedimento');
const fracLIdx = hL.findIndex(h=>String(h).trim()==='FraccionNico');
const cantLIdx = hL.findIndex(h=>String(h).trim()==='CantidadSaldo');
const vcusdIdx = hL.findIndex(h=>String(h).trim()==='VCUSD');
const notasIdx = hL.findIndex(h=>String(h).trim()==='Notas');

const sPedIdx  = h5.findIndex(h=>String(h).trim()==='Pedimento');
const sFracIdx = h5.findIndex(h=>String(h).trim()==='Fraccion');
const sSeqIdx  = h5.findIndex(h=>String(h).trim()==='SecuenciaFraccion');
const sCantIdx = h5.findIndex(h=>String(h).trim()==='CantidadUMComercial');
const sValIdx  = h5.findIndex(h=>String(h).trim()==='ValorDolares');
const sPaisIdx = h5.findIndex(h=>String(h).trim()==='PaisOrigenDestino');

// Encontrar la línea separadora del orphan section
let orphanStartRow = -1;
for(let i=1;i<layout.length;i++){
  const r = layout[i];
  const ped = String(r[pedLIdx]||'');
  if(ped.includes('SECUENCIAS DEL 551 NO ASIGNADAS')) {
    orphanStartRow = i;
    break;
  }
}
console.log('Orphan section empieza en fila:', orphanStartRow, '(0-indexed)');
console.log('Total filas Layout originales (sin header, sin orphan):', orphanStartRow - 1);

// Separar filas originales del Layout vs filas orphan
const originalRows = layout.slice(1, orphanStartRow);
const orphanRows   = layout.slice(orphanStartRow + 2); // +1 separador +1 header

console.log('Filas orphan (551 sin match):', orphanRows.filter(r=>r.some(c=>c!==''&&c!=null)).length);

// Pedimentos únicos en Layout original
const pedSetLayout = new Set(originalRows.map(r=>String(r[pedLIdx]||'').trim()).filter(Boolean));
console.log('\nPedimentos únicos en Layout:', pedSetLayout.size, '->', [...pedSetLayout].join(' | '));

// Pedimentos únicos en 551
const pedSet551 = new Set(s551.slice(1).map(r=>String(r[sPedIdx]||'').trim()).filter(Boolean));
console.log('Pedimentos únicos en 551:   ', pedSet551.size, '->', [...pedSet551].join(' | '));

// ¿Todos los pedimentos del 551 están en Layout?
const pedIn551NoLayout = [...pedSet551].filter(p=>!pedSetLayout.has(p));
console.log('Pedimentos en 551 pero NO en Layout:', pedIn551NoLayout);

// Fracciones en 551 que NO tienen saldo en Layout (CantidadSaldo = 0)
const fracCantMap = new Map(); // ped+frac -> totalCantSaldo
for(const r of originalRows){
  if(!r.some(c=>c!==''&&c!=null)) continue;
  const key = String(r[pedLIdx]||'').trim() + '|||' + String(r[fracLIdx]||'').trim();
  fracCantMap.set(key, (fracCantMap.get(key)||0) + (parseFloat(r[cantLIdx])||0));
}

// Revisar cuántas filas 551 tienen su Ped+Frac en Layout vs no
let en551ConLayout=0, en551SinLayout=0;
let ejemplosOrfanos=[];
for(let i=1;i<s551.length;i++){
  const r = s551[i];
  if(!r||r.every(c=>c===''||c==null)) continue;
  const ped  = String(r[sPedIdx]||'').trim();
  const frac = String(r[sFracIdx]||'').trim().replace(/^0+/,'') || '0';
  const seq  = r[sSeqIdx];
  const cant = parseFloat(r[sCantIdx])||0;
  const val  = parseFloat(r[sValIdx])||0;
  const key  = ped + '|||' + frac;
  const layoutCant = fracCantMap.get(key) || 0;
  if(layoutCant > 0){
    en551ConLayout++;
  } else {
    en551SinLayout++;
    if(ejemplosOrfanos.length < 10)
      ejemplosOrfanos.push({seq, ped, frac, cant551: cant, val551: val, layoutCant});
  }
}
console.log('\n=== 551 ROWS con/sin Layout ===');
console.log('551 rows con Layout saldo:', en551ConLayout);
console.log('551 rows SIN Layout saldo:', en551SinLayout, '(estos son los que no pueden asignarse)');
console.log('\nEjemplos de 551 sin Layout saldo:');
ejemplosOrfanos.forEach((r,i)=>console.log(i+1+':', JSON.stringify(r)));

// Secuencias asignadas en Layout - distribución
const secDistMap = new Map();
for(const r of originalRows){
  if(!r.some(c=>c!==''&&c!=null)) continue;
  const sec = String(r[secIdx]||'').trim();
  if(!sec) continue;
  secDistMap.set(sec, (secDistMap.get(sec)||0)+1);
}
const secsConMuchasFilas = [...secDistMap.entries()].sort((a,b)=>b[1]-a[1]).slice(0,10);
console.log('\nTop 10 secuencias con más filas Layout asignadas:');
secsConMuchasFilas.forEach(([s,n])=>console.log(`  Sec ${s}: ${n} filas`));

// Revisar si el SecuenciaPed en Layout ya tenía valores ANTES del proceso
// (buscamos notas que digan "asignada por la app" vs "modificada" vs "sin cambio")
let noChange=0, added=0, modified=0, corrected=0, orphanNota=0;
for(const r of originalRows){
  if(!r.some(c=>c!==''&&c!=null)) continue;
  const nota = String(r[notasIdx]||'').toLowerCase();
  if(nota.includes('corrección')) corrected++;
  else if(nota.includes('asignada')) added++;
  else if(nota.includes('modificada')) modified++;
  else noChange++;
}
console.log('\n=== TIPO DE ASIGNACIÓN EN LAYOUT ===');
console.log('Sin cambio (secuencia ya correcta):', noChange);
console.log('Secuencia asignada (nueva):',         added);
console.log('Secuencia modificada:',               modified);
console.log('Con corrección de campo:',            corrected);
