/**
 * Simula readDS2020Sheet + readLayout2020Sheet con reducido.xlsx
 * para confirmar que se leen correctamente antes de probar en el navegador.
 */
const XLSX = require('xlsx');

const nH2020 = (s) => String(s ?? "").trim().toLowerCase().replace(/[\s_\-]/g, "");

// ── readDS2020Sheet ───────────────────────────────────────────────────────────
function readDS2020Sheet(sheet) {
  if (!sheet) { console.log("[DS2020] sheet undefined"); return []; }
  const DS_COL_MAP = {
    Pedimento2:           ["Pedimento2"],
    Fraccion:             ["Fraccion"],
    SecuenciaFraccion:    ["SecuenciaFraccion"],
    DescripcionMercancia: ["DescripcionMercancia"],
    CantidadUMComercial:  ["CantidadUMComercial"],
    ValorDolares:         ["ValorDolares","Valor usd redondeado","Valor Aduana Estadístico","ValorAduana"],
    PaisOrigenDestino:    ["PaisOrigenDestino"],
    "Candado 551":        ["Candado 551","Candado DS 551"],
  };
  const knownNorms = new Set(Object.values(DS_COL_MAP).flat().map(nH2020));
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  let hdrI = 0, bestHits = 0;
  for (let i = 0; i < Math.min(rows.length, 5); i++) {
    const hits = rows[i].filter(c => knownNorms.has(nH2020(String(c ?? "")))).length;
    if (hits > bestHits) { bestHits = hits; hdrI = i; }
    if (hits >= 3) break;
  }
  const hdr = rows[hdrI].map(c => String(c ?? "").trim());
  console.log("[DS2020] hdrI:", hdrI, "hits:", bestHits, "headers:", hdr);
  const idx = {};
  for (const [internalName, aliases] of Object.entries(DS_COL_MAP)) {
    const aliasNorms = aliases.map(nH2020);
    const found = hdr.findIndex(h => aliasNorms.includes(nH2020(h)));
    if (found >= 0) idx[internalName] = found;
  }
  console.log("[DS2020] colIdx:", JSON.stringify(idx));
  const out = [];
  for (let i = hdrI + 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.every(c => c === "" || c == null)) continue;
    const obj = { _dsIdx: out.length };
    for (const [col, ci] of Object.entries(idx)) obj[col] = row[ci] ?? "";
    out.push(obj);
  }
  console.log("[DS2020] DS rows leídas:", out.length);
  out.forEach((r, i) => console.log(`  DS[${i}]:`, JSON.stringify(r)));
  return out;
}

// ── readLayout2020Sheet ───────────────────────────────────────────────────────
function readLayout2020Sheet(sheet) {
  if (!sheet || !sheet["!ref"]) { console.log("[Layout2020] ERROR no sheet/!ref"); return { layoutRows: [] }; }
  const range = XLSX.utils.decode_range(sheet["!ref"]);
  console.log("\n[Layout2020] ref:", sheet["!ref"], "filas:", range.e.r+1, "cols:", range.e.c+1);

  const hdrRange = { s:{r:0,c:range.s.c}, e:{r:Math.min(14,range.e.r),c:range.e.c} };
  const sampleRows = XLSX.utils.sheet_to_json(sheet, { header:1, defval:"", range:hdrRange });

  const KNOWN = new Set(["pedimento","fraccionnico","seccalc","descripcion",
                         "paisorigen","valormpdolares","cantidadcomercial","cantidad_comercial","notas","estado"]);
  let hdrI = 0, bestHits = 0;
  for (let i = 0; i < sampleRows.length; i++) {
    const hits = sampleRows[i].filter(c => KNOWN.has(nH2020(String(c ?? "")))).length;
    console.log(`  fila[${i}] hits:${hits}`, sampleRows[i].slice(0,5));
    if (hits > bestHits) { bestHits = hits; hdrI = i; }
    if (hits >= 4) break;
  }
  console.log("[Layout2020] hdrI:", hdrI, "bestHits:", bestHits);

  const rawHeaders = (sampleRows[hdrI] || []).map(c => String(c ?? "").trim());
  console.log("[Layout2020] headers:", rawHeaders);

  const findFirst = (...names) => {
    const ts = names.map(nH2020);
    for (let i = 0; i < rawHeaders.length; i++) { if (ts.includes(nH2020(rawHeaders[i]))) return i; }
    return -1;
  };
  const findLast = (...names) => {
    const ts = names.map(nH2020);
    return rawHeaders.reduce((last, h, i) => ts.includes(nH2020(h)) ? i : last, -1);
  };

  const colIdx = {
    pedimento: findFirst("pedimento"),
    frac:      findLast("FraccionNico","fraccionnico"),
    desc:      findFirst("descripcion","descripcionmercancia"),
    pais:      findFirst("pais_origen","paisorigen","paisorigendestino"),
    cant:      findFirst("cantidad_comercial","cantidadcomercial","cantidadumc"),
    val:       findFirst("ValorMPDolares","valormpdolares","valordolares","valor_me","valorme"),
    sec:       findFirst("SEC CALC","seccalc","secuenciaped"),
    notasIn:   findFirst("NOTAS","notas"),
    notas:     findLast("NOTAS","notas"),
    estado:    findFirst("ESTADO","estado"),
  };
  console.log("[Layout2020] colIdx:", JSON.stringify(colIdx));
  console.log("  pedimento col:", rawHeaders[colIdx.pedimento]);
  console.log("  frac col:", rawHeaders[colIdx.frac]);
  console.log("  sec col:", rawHeaders[colIdx.sec]);
  console.log("  val col:", rawHeaders[colIdx.val]);
  console.log("  pais col:", rawHeaders[colIdx.pais]);
  console.log("  cant col:", rawHeaders[colIdx.cant]);
  console.log("  notasIn col:", rawHeaders[colIdx.notasIn]);
  console.log("  notas col:", rawHeaders[colIdx.notas]);

  const cellVal = (r, c) => {
    if (c < 0) return "";
    const cell = sheet[XLSX.utils.encode_cell({r,c})];
    if (!cell) return "";
    return String(cell.v ?? cell.w ?? "").trim();
  };
  const cellNum = (r, c) => {
    if (c < 0) return 0;
    const cell = sheet[XLSX.utils.encode_cell({r,c})];
    return cell ? (parseFloat(cell.v) || 0) : 0;
  };

  const isRealSec = (v) => { const s = String(v??"").trim(); return s !== "" && s !== "." && !isNaN(parseFloat(s)); };

  const layoutRows = [];
  for (let r = hdrI + 1; r <= range.e.r; r++) {
    const pedVal  = cellVal(r, colIdx.pedimento);
    const fracVal = cellVal(r, colIdx.frac);
    if (!pedVal && !fracVal) continue;
    const notasInVal = cellVal(r, colIdx.notasIn).toUpperCase();
    const noIncluir  = notasInVal.includes("NO INCLUIR");
    layoutRows.push({
      _idx: layoutRows.length, _rowI: r,
      Pedimento: pedVal, FraccionNico: fracVal,
      Descripcion: cellVal(r, colIdx.desc),
      PaisOrigen: cellVal(r, colIdx.pais),
      Cantidad: cellNum(r, colIdx.cant),
      ValorUSD: cellNum(r, colIdx.val),
      SecCalc: cellVal(r, colIdx.sec),
      Notas: cellVal(r, colIdx.notas),
      Estado: cellVal(r, colIdx.estado),
      noIncluir,
      secIsReal: isRealSec(cellVal(r, colIdx.sec)),
    });
  }
  console.log("\n[Layout2020] layoutRows:", layoutRows.length);
  layoutRows.forEach((r, i) => {
    if (i < 10) console.log(`  Row[${i}]: ped=${r.Pedimento} frac=${r.FraccionNico} sec=${r.SecCalc} secReal=${r.secIsReal} noInc=${r.noIncluir} cant=${r.Cantidad} val=${r.ValorUSD}`);
  });
  return { layoutRows, colIdx };
}

// ── Main ──────────────────────────────────────────────────────────────────────
const wb = XLSX.readFile('c:/Users/LCK_KATHIA/Desktop/reducido.xlsx', { cellStyles: false });
console.log("Hojas:", wb.SheetNames);

const dsName  = wb.SheetNames.find(n => n.toUpperCase().includes("DS"));
const layName = wb.SheetNames.find(n => n.toLowerCase().includes("layout"));
console.log("dsName:", dsName, "layName:", layName);

console.log("\n=== DS ===");
const dsRows = readDS2020Sheet(wb.Sheets[dsName]);

console.log("\n=== LAYOUT ===");
const { layoutRows, colIdx } = readLayout2020Sheet(wb.Sheets[layName]);

console.log("\n=== RESUMEN ===");
console.log("DS rows:", dsRows.length);
console.log("Layout rows:", layoutRows.length);
console.log("Rows con sec real:", layoutRows.filter(r => r.secIsReal).length);
console.log("Rows NO INCLUIR:", layoutRows.filter(r => r.noIncluir).length);
console.log("Rows a asignar:", layoutRows.filter(r => !r.secIsReal && !r.noIncluir).length);
