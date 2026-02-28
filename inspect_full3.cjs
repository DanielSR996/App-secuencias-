const XLSX = require('xlsx');
const nH2020 = s => String(s??'').trim().toLowerCase().replace(/[\s_\-]/g,'');

function resolveDS2020SheetNames(wb) {
  const names = wb.SheetNames || [];
  const dsName = names.find(n => n.toUpperCase().includes('DS'));
  const LAY_KNOWN = new Set([
    'pedimento','fraccionnico','seccalc','descripcion','paisorigen','pais_origen',
    'valormpdolares','cantidad_comercial','cantidadcomercial','notas','estado',
    'aduana_es','numero_parte','numeroparte','precio_unitario','valorme','fraccionmex',
  ]);
  let layName = null, bestHits = 0;
  for (const name of names) {
    if (name === dsName) continue;
    const ws = wb.Sheets[name];
    if (!ws) continue;
    try {
      const rows = XLSX.utils.sheet_to_json(ws, { header:1, defval:'', sheetRows:5 });
      for (const row of rows) {
        const hits = row.filter(c => LAY_KNOWN.has(nH2020(String(c??'')))).length;
        if (hits > bestHits) { bestHits = hits; layName = name; }
      }
    } catch(_){}
  }
  console.log('[resolve2020] dsName:', dsName, '| layName:', layName, '(hits:', bestHits, ')');
  return { dsName, layName };
}

const FILE = process.argv[2] || 'c:/Users/LCK_KATHIA/Desktop/0_Avance 2020 electronics 22012026 (1).xlsx';
console.log("Abriendo:", FILE);
const wb = XLSX.readFile(FILE, { cellStyles:false, cellNF:false, sheetRows:8 });
console.log('Hojas:', wb.SheetNames);
const { dsName, layName } = resolveDS2020SheetNames(wb);

// Mostrar headers del Layout detectado
if (layName && wb.Sheets[layName]) {
  const ws = wb.Sheets[layName];
  console.log('\n=== Layout detectado:', layName, '===');
  console.log('  !ref:', ws['!ref']);
  const rows = XLSX.utils.sheet_to_json(ws, { header:1, defval:'' });
  for (let i = 0; i < Math.min(4, rows.length); i++) {
    const ne = rows[i].filter(c => c !== '' && c != null).length;
    console.log(`  fila[${i}] (${ne} cols):`, JSON.stringify(rows[i]).slice(0,350));
  }
}

// Mostrar primeros datos DS
if (dsName && wb.Sheets[dsName]) {
  const ws = wb.Sheets[dsName];
  console.log('\n=== DS detectado:', dsName, '===');
  console.log('  !ref:', ws['!ref']);
  const rows = XLSX.utils.sheet_to_json(ws, { header:1, defval:'' });
  console.log('  fila[0]:', JSON.stringify(rows[0]).slice(0,400));
  console.log('  fila[1]:', JSON.stringify(rows[1]).slice(0,400));
}
