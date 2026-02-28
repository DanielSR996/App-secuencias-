const XLSX = require('xlsx');
const nH = s => String(s??'').trim().toLowerCase().replace(/[\s_\-]/g,'');

const FILE = 'c:/Users/LCK_KATHIA/Desktop/2020.xlsx';
console.log("Abriendo:", FILE);
const wb = XLSX.readFile(FILE, { cellStyles:false, cellNF:false, sheetRows:6 });
console.log('Hojas:', wb.SheetNames);

const LAY_KNOWN = new Set([
  'pedimento','fraccionnico','seccalc','descripcion','paisorigen','pais_origen',
  'valormpdolares','cantidad_comercial','cantidadcomercial','notas','estado',
  'aduana_es','numero_parte','numeroparte','precio_unitario','valorme','fraccionmex',
]);

for (const name of wb.SheetNames) {
  const ws = wb.Sheets[name];
  if (!ws) { console.log(`\n[${name}] → NO CARGÓ`); continue; }
  console.log(`\n[${name}] ref:${ws['!ref']}`);
  const rows = XLSX.utils.sheet_to_json(ws, { header:1, defval:'', sheetRows:6 });
  let maxHits = 0;
  for (const row of rows) {
    const hits = row.filter(c => LAY_KNOWN.has(nH(String(c??'')))).length;
    if (hits > maxHits) maxHits = hits;
  }
  console.log(`  maxHits en LAY_KNOWN: ${maxHits}`);
  // Mostrar primeras 3 filas
  for (let i = 0; i < Math.min(3, rows.length); i++) {
    const ne = rows[i].filter(c=>c!==''&&c!=null).length;
    console.log(`  fila[${i}] (${ne} no vacías): ${JSON.stringify(rows[i]).slice(0,250)}`);
  }
}
