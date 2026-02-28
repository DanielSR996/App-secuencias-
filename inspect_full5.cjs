/**
 * Simula exactamente lo que hace App2020 con el archivo completo.
 */
const XLSX = require('xlsx');
const nH2020 = s => String(s??'').trim().toLowerCase().replace(/[\s_\-]/g,'');

// ── resolveDS2020SheetNames ───────────────────────────────────────────────────
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
  console.log('[resolve] dsName:', dsName, '| layName:', layName, '(hits:', bestHits, ')');
  return { dsName, layName };
}

// ── readLayout2020Sheet ───────────────────────────────────────────────────────
function readLayout2020Sheet(sheet) {
  if (!sheet || !sheet['!ref']) { console.log('[Layout] ERROR'); return {layoutRows:[]}; }
  const range = XLSX.utils.decode_range(sheet['!ref']);
  console.log('[Layout] ref:', sheet['!ref'], 'filas:', range.e.r+1);

  const hdrRange = { s:{r:0,c:range.s.c}, e:{r:Math.min(14,range.e.r),c:range.e.c} };
  const sampleRows = XLSX.utils.sheet_to_json(sheet, { header:1, defval:'', range:hdrRange });

  const KNOWN = new Set(['pedimento','fraccionnico','seccalc','descripcion',
    'paisorigen','valormpdolares','cantidadcomercial','cantidad_comercial','notas','estado']);
  let hdrI = 0, bestHits = 0;
  for (let i = 0; i < sampleRows.length; i++) {
    const hits = sampleRows[i].filter(c => KNOWN.has(nH2020(String(c??'')))).length;
    if (hits > bestHits) { bestHits = hits; hdrI = i; }
  }
  console.log('[Layout] hdrI:', hdrI, 'hits:', bestHits);

  const rawHeaders = (sampleRows[hdrI]||[]).map(c => String(c??'').trim());

  const findFirst = (...names) => {
    for (const name of names) {
      const n = nH2020(name);
      const idx = rawHeaders.findIndex(h => nH2020(h) === n);
      if (idx >= 0) return idx;
    }
    return -1;
  };
  const findLast = (...names) => {
    const ts = names.map(nH2020);
    return rawHeaders.reduce((last,h,i) => ts.includes(nH2020(h)) ? i : last, -1);
  };

  const colIdx = {
    pedimento: findLast('pedimento'),
    frac:      findLast('FraccionNico','fraccionnico'),
    cant:      findLast('cantidad_comercial','cantidadcomercial','cantidadumc'),
    notas:     findLast('NOTAS','notas'),
    desc:      findFirst('descripcion','clase_descripcion','descripcionmercancia'),
    pais:      findFirst('pais_origen','paisorigen','paisorigendestino'),
    val:       findFirst('ValorMPDolares','valormpdolares','valordolares','valor_me','valorme'),
    sec:       findFirst('SEC CALC','seccalc','secuenciaped'),
    notasIn:   findFirst('NOTAS','notas'),
    estado:    findFirst('ESTADO','estado'),
  };
  console.log('[Layout] colIdx:', JSON.stringify(colIdx));
  const show = (k) => rawHeaders[colIdx[k]] ? `"${rawHeaders[colIdx[k]]}"` : 'NO ENCONTRADO';
  console.log(`  pedimento  → col ${colIdx.pedimento}: ${show('pedimento')}`);
  console.log(`  frac       → col ${colIdx.frac}: ${show('frac')}`);
  console.log(`  sec        → col ${colIdx.sec}: ${show('sec')}`);
  console.log(`  val        → col ${colIdx.val}: ${show('val')}`);
  console.log(`  pais       → col ${colIdx.pais}: ${show('pais')}`);
  console.log(`  cant       → col ${colIdx.cant}: ${show('cant')}`);
  console.log(`  notasIn    → col ${colIdx.notasIn}: ${show('notasIn')}`);
  console.log(`  notas(out) → col ${colIdx.notas}: ${show('notas')}`);
  console.log(`  desc       → col ${colIdx.desc}: ${show('desc')}`);
  console.log(`  estado     → col ${colIdx.estado}: ${show('estado')}`);

  const cellVal = (r,c) => {
    if(c<0) return '';
    const cell = sheet[XLSX.utils.encode_cell({r,c})];
    if(!cell) return '';
    return String(cell.v??cell.w??'').trim();
  };
  const cellNum = (r,c) => {
    if(c<0) return 0;
    const cell = sheet[XLSX.utils.encode_cell({r,c})];
    return cell ? (parseFloat(cell.v)||0) : 0;
  };
  const isRealSec = v => { const s=String(v??'').trim(); return s!==''&&s!=='.'&&!isNaN(parseFloat(s)); };

  const layoutRows = [];
  for (let r = hdrI+1; r <= range.e.r; r++) {
    const pedVal = cellVal(r, colIdx.pedimento);
    const fracVal = cellVal(r, colIdx.frac);
    if (!pedVal && !fracVal) continue;
    const notasInVal = cellVal(r, colIdx.notasIn).toUpperCase();
    layoutRows.push({
      _idx: layoutRows.length, _rowI: r,
      Pedimento: pedVal, FraccionNico: fracVal,
      Descripcion: cellVal(r, colIdx.desc),
      PaisOrigen: cellVal(r, colIdx.pais),
      Cantidad: cellNum(r, colIdx.cant),
      ValorUSD: cellNum(r, colIdx.val),
      SecCalc: cellVal(r, colIdx.sec),
      noIncluir: notasInVal.includes('NO INCLUIR'),
      secIsReal: isRealSec(cellVal(r, colIdx.sec)),
    });
  }
  console.log('[Layout] layoutRows:', layoutRows.length);
  layoutRows.slice(0,5).forEach((r,i) => {
    console.log(`  Row[${i}]: ped=${r.Pedimento} frac=${r.FraccionNico} sec=${r.SecCalc} secReal=${r.secIsReal} noInc=${r.noIncluir} cant=${r.Cantidad} val=${r.ValorUSD}`);
  });
  return { layoutRows, colIdx };
}

// ── Main ──────────────────────────────────────────────────────────────────────
const FILE = process.argv[2] || 'c:/Users/LCK_KATHIA/Desktop/0_Avance 2020 electronics 22012026 (1).xlsx';
console.log('Archivo:', FILE);
const wb = XLSX.readFile(FILE, { cellStyles:false, cellNF:false });
console.log('Hojas:', wb.SheetNames);
const {dsName, layName} = resolveDS2020SheetNames(wb);
if (!layName) { console.log('ERROR: No Layout detectado'); process.exit(1); }
const {layoutRows} = readLayout2020Sheet(wb.Sheets[layName]);
console.log('\n=== RESUMEN ===');
console.log('Total filas Layout:', layoutRows.length);
console.log('Con sec real:', layoutRows.filter(r=>r.secIsReal).length);
console.log('NO INCLUIR:', layoutRows.filter(r=>r.noIncluir).length);
console.log('A asignar:', layoutRows.filter(r=>!r.secIsReal&&!r.noIncluir).length);
// Pedimentos únicos
const peds = [...new Set(layoutRows.map(r=>r.Pedimento))];
console.log('Pedimentos únicos:', peds.length, peds.slice(0,5));
