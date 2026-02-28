/**
 * Inspecciona la estructura del archivo completo multi-pedimento
 * para comparar contra reducido.xlsx
 */
const XLSX = require('xlsx');

// Prueba con el archivo completo — ajusta la ruta si es diferente
const FILE = process.argv[2] || 'c:/Users/LCK_KATHIA/Desktop/0_Avance 2020 electronics 22012026 (1).xlsx';

console.log("Abriendo:", FILE);

let wb;
try {
  wb = XLSX.readFile(FILE, {
    cellStyles: false, cellNF: false, cellDates: false,
    sheetRows: 10, // solo primeras 10 filas por hoja para ser rápido
  });
} catch(e) {
  console.error("ERROR al abrir:", e.message);
  process.exit(1);
}

console.log('\n=== HOJAS ===', wb.SheetNames);

// ── Detectar DS y Layout ──────────────────────────────────────────────────────
const dsName  = wb.SheetNames.find(n => n.toUpperCase().includes('DS'));
const layName = wb.SheetNames.find(n => n.toLowerCase().includes('layout'));
console.log('\nHoja DS  detectada:', dsName);
console.log('Hoja Layout detectada:', layName);

// ── Inspeccionar DS ───────────────────────────────────────────────────────────
if (dsName && wb.Sheets[dsName]) {
  console.log('\n=== DS:', dsName, '===');
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[dsName], { header:1, defval:'' });
  console.log('Filas leídas (max 10):', rows.length);
  // Mostrar primeras 3 filas
  for (let i = 0; i < Math.min(3, rows.length); i++) {
    const noEmpty = rows[i].filter(c => c !== '' && c != null);
    console.log(`  fila[${i}] (${noEmpty.length} cols):`, JSON.stringify(rows[i]).slice(0, 300));
  }
} else {
  console.log('\n[DS] No cargó o no existe');
}

// ── Inspeccionar Layout ───────────────────────────────────────────────────────
if (layName && wb.Sheets[layName]) {
  console.log('\n=== Layout:', layName, '===');
  const ws = wb.Sheets[layName];
  console.log('  !ref:', ws['!ref']);
  const rows = XLSX.utils.sheet_to_json(ws, { header:1, defval:'' });
  console.log('Filas leídas (max 10):', rows.length);
  for (let i = 0; i < Math.min(5, rows.length); i++) {
    const noEmpty = rows[i].filter(c => c !== '' && c != null);
    console.log(`  fila[${i}] (${noEmpty.length} cols):`, JSON.stringify(rows[i]).slice(0, 400));
  }
} else {
  console.log('\n[Layout] No cargó o no existe');
  // Si la hoja existe en SheetNames pero no en Sheets, es el problema del archivo grande
  if (layName) console.log('  → Hoja en SheetNames pero wb.Sheets[layName] es', wb.Sheets[layName]);
}
