const XLSX = require('xlsx');
const FILE = process.argv[2] || 'c:/Users/LCK_KATHIA/Desktop/0_Avance 2020 electronics 22012026 (1).xlsx';
const wb = XLSX.readFile(FILE, { cellStyles:false, cellNF:false, sheetRows:3 });

// DS completo
const dsName = wb.SheetNames.find(n => n.toUpperCase().includes('DS'));
console.log('\n=== DS 2020 HEADERS COMPLETOS ===');
const dsRows = XLSX.utils.sheet_to_json(wb.Sheets[dsName], { header:1, defval:'' });
const dsHdr = dsRows[0];
dsHdr.forEach((h, i) => { if(String(h).trim()) console.log(`  [${i}] "${h}"`); });
console.log('\nFila 1 DS (datos):');
dsHdr.forEach((h, i) => {
  if(String(h).trim() && dsRows[1][i] !== '' && dsRows[1][i] != null)
    console.log(`  ${h}: ${JSON.stringify(dsRows[1][i])}`);
});

// Layout 2020 headers (buscar en fila correcta)
const layName = '2020';
console.log('\n=== LAYOUT 2020 HEADERS COMPLETOS ===');
const lyRows = XLSX.utils.sheet_to_json(wb.Sheets[layName], { header:1, defval:'' });
// Encontrar fila de headers
const nH = s => String(s??'').trim().toLowerCase().replace(/[\s_\-]/g,'');
const KNOWN = new Set(['pedimento','fraccionnico','seccalc','descripcion','paisorigen','pais_origen',
  'valormpdolares','cantidadcomercial','cantidad_comercial','notas','estado','aduana_es']);
let hdrI = 0, bHits = 0;
for (let i = 0; i < lyRows.length; i++) {
  const h = lyRows[i].filter(c => KNOWN.has(nH(String(c??'')))).length;
  if (h > bHits) { bHits = h; hdrI = i; }
}
console.log('hdrI:', hdrI, 'hits:', bHits);
const lyHdr = lyRows[hdrI];
lyHdr.forEach((h, i) => { if(String(h).trim()) console.log(`  [${i}] "${h}"`); });
