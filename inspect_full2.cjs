const XLSX = require('xlsx');
const FILE = process.argv[2] || 'c:/Users/LCK_KATHIA/Desktop/0_Avance 2020 electronics 22012026 (1).xlsx';

console.log("Abriendo:", FILE);
const wb = XLSX.readFile(FILE, { cellStyles: false, cellNF: false, sheetRows: 8 });
console.log('Hojas:', wb.SheetNames);

// Inspeccionar CADA hoja para ver cuál tiene estructura de Layout real
for (const name of wb.SheetNames) {
  const ws = wb.Sheets[name];
  if (!ws) { console.log(`\n[${name}] → NO CARGÓ`); continue; }
  const rows = XLSX.utils.sheet_to_json(ws, { header:1, defval:'' });
  if (!rows.length) { console.log(`\n[${name}] → VACÍA`); continue; }
  const fila0 = rows[0].filter(c => c !== '' && c != null);
  console.log(`\n[${name}] ref:${ws['!ref']} | fila0 (${fila0.length} cols):`, JSON.stringify(rows[0]).slice(0,300));
  if (rows[1]) {
    const fila1 = rows[1].filter(c => c !== '' && c != null);
    console.log(`  fila1 (${fila1.length} cols):`, JSON.stringify(rows[1]).slice(0,300));
  }
}
