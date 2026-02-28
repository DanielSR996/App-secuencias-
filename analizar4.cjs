const XLSX = require('./node_modules/xlsx');
const wb = XLSX.readFile('c:/Users/LCK_KATHIA/Downloads/prueba1.xlsx');

const layout = XLSX.utils.sheet_to_json(wb.Sheets['Layout'], {header:1, defval:''});
const s551   = XLSX.utils.sheet_to_json(wb.Sheets['551'],    {header:1, defval:''});
const hL = layout[0];
const h5 = s551[0];

// Buscar todas las columnas que puedan servir de clave directa
console.log('=== COLUMNAS LAYOUT (todas) ===');
hL.forEach((h,i)=>console.log(i+':', String(h).trim()));

console.log('\n=== COLUMNAS 551 (todas) ===');
h5.forEach((h,i)=>console.log(i+':', String(h).trim()));

// Buscar columna PED_SEC en Layout
const pedSecIdx = hL.findIndex(h=>String(h).trim()==='PED_SEC');
const secIdx    = hL.findIndex(h=>String(h).trim()==='SecuenciaPed');
// Buscar columna Secuencias en 551
const secuenciasIdx = h5.findIndex(h=>String(h).trim()==='Secuencias');
const seqFracIdx    = h5.findIndex(h=>String(h).trim()==='SecuenciaFraccion');

console.log('\nPED_SEC en Layout:', pedSecIdx);
console.log('Secuencias en 551:', secuenciasIdx, '| SecuenciaFraccion:', seqFracIdx);

// Ver si PED_SEC del Layout coincide con Secuencias del 551
if(pedSecIdx >= 0 && secuenciasIdx >= 0){
  // Construir set de Secuencias del 551
  const seqMap = new Map(); // Secuencias -> SecuenciaFraccion
  for(let i=1;i<s551.length;i++){
    const r = s551[i];
    if(!r||r.every(c=>c===''||c==null)) continue;
    const key = String(r[secuenciasIdx]||'').trim();
    const seq = String(r[seqFracIdx]||'').trim();
    if(key) seqMap.set(key, seq);
  }
  console.log('\nSecuencias únicas en 551:', seqMap.size);
  
  // Checar qué % de Layout.PED_SEC coincide con 551.Secuencias
  let match=0, noMatch=0, empty=0;
  let orphanStart = -1;
  for(let i=1;i<layout.length;i++){
    if(String(layout[i][hL.findIndex(h=>String(h).trim()==='Pedimento')]||'').includes('SECUENCIAS DEL 551')){
      orphanStart=i; break;
    }
  }
  const origLayout = layout.slice(1, orphanStart);
  
  const ejemplosMatch = [];
  const ejemplosNoMatch = [];
  
  for(const r of origLayout){
    if(!r.some(c=>c!==''&&c!=null)) continue;
    const pedSec = String(r[pedSecIdx]||'').trim();
    if(!pedSec){ empty++; continue; }
    if(seqMap.has(pedSec)){
      match++;
      if(ejemplosMatch.length < 3) ejemplosMatch.push({pedSec, seq551: seqMap.get(pedSec), seqLayout: String(r[secIdx]||'').trim()});
    } else {
      noMatch++;
      if(ejemplosNoMatch.length < 3) ejemplosNoMatch.push({pedSec, seqLayout: String(r[secIdx]||'').trim()});
    }
  }
  console.log('\n=== COINCIDENCIA PED_SEC vs Secuencias ===');
  console.log('Match directo:', match, '/', origLayout.filter(r=>r.some(c=>c!==''&&c!=null)).length);
  console.log('Sin match:', noMatch, '| PED_SEC vacío:', empty);
  console.log('\nEjemplos con match:');
  ejemplosMatch.forEach((r,i)=>console.log(i+1+':', JSON.stringify(r)));
  console.log('\nEjemplos sin match:');
  ejemplosNoMatch.forEach((r,i)=>console.log(i+1+':', JSON.stringify(r)));
} else {
  console.log('\nNo se encontraron columnas PED_SEC o Secuencias');
}
