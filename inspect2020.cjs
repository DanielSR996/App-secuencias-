const XLSX = require('xlsx');

const wb = XLSX.readFile('c:/Users/LCK_KATHIA/Desktop/reducido.xlsx', {
  cellStyles: false, cellNF: false, cellDates: false,
});

console.log('=== HOJAS ===', wb.SheetNames);

// ── DS ─────────────────────────────────────────────────────────────────────
const dsName = wb.SheetNames.find(n => n.toUpperCase().includes('DS'));
console.log('\n=== DS sheet:', dsName, '===');
if (wb.Sheets[dsName]) {
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[dsName], { header:1, defval:'' });
  console.log('Total filas:', rows.length);
  console.log('Fila 0 (headers):');
  (rows[0]||[]).forEach((h,i) => { if(String(h).trim()) console.log(`  [${i}] "${h}"`); });
  console.log('Fila 1 (primer dato):');
  (rows[0]||[]).forEach((h,i) => {
    const v = (rows[1]||[])[i];
    if(String(h).trim() && v !== '' && v != null) console.log(`  ${h}: ${JSON.stringify(v)}`);
  });
}

// ── Layout ─────────────────────────────────────────────────────────────────
const layName = wb.SheetNames.find(n => n.toLowerCase().includes('layout'));
console.log('\n=== Layout sheet:', layName, '===');
if (wb.Sheets[layName]) {
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[layName], { header:1, defval:'' });
  console.log('Total filas:', rows.length);
  // Mostrar primeras 4 filas con todos los valores
  for (let i = 0; i < Math.min(4, rows.length); i++) {
    const ne = rows[i].filter(c => c !== '' && c != null).length;
    console.log(`  fila[${i}] (${ne} no vacías): ${JSON.stringify(rows[i]).slice(0,400)}`);
  }
  console.log('\nHeaders fila[0]:');
  (rows[0]||[]).forEach((h,i) => { if(String(h).trim()) console.log(`  [${i}] "${h}"`); });
  console.log('\nFila 1 (primer dato):');
  (rows[0]||[]).forEach((h,i) => {
    const v = (rows[1]||[])[i];
    if(String(h).trim() && v !== '' && v != null) console.log(`  ${h}: ${JSON.stringify(v)}`);
  });
}
