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
const paisLIdx = hL.findIndex(h=>String(h).trim()==='PaisOrigen');
const notasIdx = hL.findIndex(h=>String(h).trim()==='Notas');

const sPedIdx  = h5.findIndex(h=>String(h).trim()==='Pedimento');
const sFracIdx = h5.findIndex(h=>String(h).trim()==='Fraccion');
const sSeqIdx  = h5.findIndex(h=>String(h).trim()==='SecuenciaFraccion');
const sCantIdx = h5.findIndex(h=>String(h).trim()==='CantidadUMComercial');
const sValIdx  = h5.findIndex(h=>String(h).trim()==='ValorDolares');
const sPaisIdx = h5.findIndex(h=>String(h).trim()==='PaisOrigenDestino');

// Solo filas originales del Layout (antes del separador)
let orphanStart = -1;
for(let i=1;i<layout.length;i++){
  if(String(layout[i][pedLIdx]||'').includes('SECUENCIAS DEL 551')){ orphanStart=i; break; }
}
const origLayout = layout.slice(1, orphanStart);

// Agrupar Layout por Ped+Frac
const layoutByFrac = new Map();
for(const r of origLayout){
  if(!r.some(c=>c!==''&&c!=null)) continue;
  const frac = String(r[fracLIdx]||'').trim().replace(/^0+/,'')||'0';
  if(!layoutByFrac.has(frac)) layoutByFrac.set(frac, []);
  layoutByFrac.get(frac).push({
    sec:  String(r[secIdx]||'').trim(),
    cant: parseFloat(r[cantLIdx])||0,
    vcusd:parseFloat(r[vcusdIdx])||0,
    pais: String(r[paisLIdx]||'').trim(),
    nota: String(r[notasIdx]||'').trim()
  });
}

// Agrupar 551 por Frac
const s551ByFrac = new Map();
for(let i=1;i<s551.length;i++){
  const r = s551[i];
  if(!r||r.every(c=>c===''||c==null)) continue;
  const frac = String(r[sFracIdx]||'').trim().replace(/^0+/,'')||'0';
  if(!s551ByFrac.has(frac)) s551ByFrac.set(frac, []);
  s551ByFrac.get(frac).push({
    seq:  String(r[sSeqIdx]||'').trim(),
    cant: parseFloat(r[sCantIdx])||0,
    val:  parseFloat(r[sValIdx])||0,
    pais: String(r[sPaisIdx]||'').trim()
  });
}

// Para cada fracción en 551, ver cuántas secuencias hay y cuántas se usan en Layout
console.log('\n=== FRACCIONES CON MÚLTIPLES SECUENCIAS EN 551 ===');
let totalOrphSeqs = 0;
const fracProblems = [];

for(const [frac, s551rows] of s551ByFrac){
  const layoutRows = layoutByFrac.get(frac) || [];
  const seqs551 = s551rows.map(r=>r.seq);
  const seqsUsedInLayout = new Set(layoutRows.map(r=>r.sec).filter(Boolean));
  const orphanSeqs = seqs551.filter(s=>!seqsUsedInLayout.has(s));
  
  if(orphanSeqs.length > 0){
    totalOrphSeqs += orphanSeqs.length;
    const totalCant551 = s551rows.reduce((a,r)=>a+r.cant,0);
    const totalCantL   = layoutRows.reduce((a,r)=>a+r.cant,0);
    const seqsUsed     = seqs551.filter(s=>seqsUsedInLayout.has(s));
    
    fracProblems.push({
      frac, 
      seqs551Count: seqs551.length,
      seqsUsed: seqsUsed.length,
      orphanSeqs: orphanSeqs.length,
      cant551Total: totalCant551,
      cantLayoutTotal: totalCantL,
      layoutRows: layoutRows.length,
      orphanSeqNums: orphanSeqs.slice(0,5).join(', '),
      usedSeqNums: seqsUsed.slice(0,5).join(', ')
    });
  }
}

fracProblems.sort((a,b)=>b.orphanSeqs-a.orphanSeqs);
console.log('Total secuencias orphan:', totalOrphSeqs);
console.log('\nTop fracciones con más orphans:');
fracProblems.slice(0,15).forEach(p=>{
  console.log(`\nFrac ${p.frac}:`);
  console.log(`  551 seqs total: ${p.seqs551Count} | Usadas: ${p.seqsUsed} | Orphan: ${p.orphanSeqs}`);
  console.log(`  551 total cant: ${p.cant551Total.toFixed(0)} | Layout total cant: ${p.cantLayoutTotal.toFixed(0)}`);
  console.log(`  Layout rows: ${p.layoutRows}`);
  console.log(`  Seqs orphan: ${p.orphanSeqNums}`);
  console.log(`  Seqs usadas: ${p.usedSeqNums}`);
});

// Analizar una fracción problemática en detalle
const topProb = fracProblems[0];
if(topProb){
  console.log('\n\n=== DETALLE FRACCIÓN ' + topProb.frac + ' ===');
  const s551fracRows = s551ByFrac.get(topProb.frac)||[];
  const lFracRows = layoutByFrac.get(topProb.frac)||[];
  console.log('551 sequences:');
  s551fracRows.forEach(r=>console.log(`  Seq ${r.seq}: cant=${r.cant.toFixed(0)}, val=${r.val.toFixed(2)}, pais=${r.pais}, up=${r.cant>0?(r.val/r.cant).toFixed(4):'N/A'}`));
  console.log('Layout rows (grouped by SecuenciaPed asignada):');
  const byAssignedSeq = new Map();
  lFracRows.forEach(r=>{
    if(!byAssignedSeq.has(r.sec)) byAssignedSeq.set(r.sec, {rows:0, cant:0, vcusd:0});
    const e = byAssignedSeq.get(r.sec);
    e.rows++; e.cant+=r.cant; e.vcusd+=r.vcusd;
  });
  for(const [seq, info] of byAssignedSeq){
    console.log(`  AsignadoSeq ${seq}: ${info.rows} filas, cant=${info.cant.toFixed(0)}, vcusd=${info.vcusd.toFixed(2)}, up=${info.cant>0?(info.vcusd/info.cant).toFixed(4):'N/A'}`);
  }
}
