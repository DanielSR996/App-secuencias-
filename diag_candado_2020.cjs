/**
 * Diagnóstico: por qué solo se asigna secuencia 2 cuando hay 5 en el DS.
 * Analiza candados en Layout (col AJ) y DS (col candado) y SecuenciaFraccion en DS.
 *
 * Uso: node diag_candado_2020.cjs "C:\ruta\archivo.xlsx"
 */
const XLSX = require("xlsx");

const FILE = process.argv[2];
if (!FILE) {
  console.error("Uso: node diag_candado_2020.cjs \"C:\\ruta\\archivo.xlsx\"");
  process.exit(1);
}

const wb = XLSX.readFile(FILE, { cellStyles: false, raw: true });
const nH = (s) => String(s ?? "").trim().toLowerCase().replace(/[\s_\-]/g, "");

// ─── DS ─────────────────────────────────────────────────────────────────
const dsAoa = XLSX.utils.sheet_to_json(wb.Sheets["DS"], { header: 1, defval: "" });
const dsHdr = (dsAoa[0] || []).map((c) => String(c ?? "").trim());
const idxCandadoDS = dsHdr.findIndex((h) => nH(h) === "candado551" || nH(h) === "candadods551" || nH(h) === "candado" || nH(h) === "clave" || nH(h) === "secuencias");
const idxSec = dsHdr.findIndex((h) => nH(h) === "secuenciafraccion");
const idxPed = dsHdr.findIndex((h) => nH(h) === "pedimento2" || nH(h) === "pedimento");
const idxCantDS = dsHdr.findIndex((h) => nH(h) === "cantidadumcomercial");
const idxValDS = dsHdr.findIndex((h) => nH(h) === "valordolares" || nH(h) === "valoragregado" || nH(h) === "valoraduanaestadístico");

console.log("=== DS ===");
console.log("Filas:", dsAoa.length);
console.log("Col candado (índice):", idxCandadoDS, idxCandadoDS >= 0 ? "(" + (dsHdr[idxCandadoDS] || "") + ")" : "NO ENCONTRADA");
console.log("Col SecuenciaFraccion:", idxSec, "| Cantidad:", idxCantDS, "| Valor:", idxValDS);
console.log("Cabeceras DS (primeras 20):", dsHdr.slice(0, 20).join(" | "));

const dsRows = [];
for (let i = 1; i < dsAoa.length; i++) {
  const row = dsAoa[i];
  if (row.every((c) => c === "" || c == null)) continue;
  const candado = String(idxCandadoDS >= 0 ? row[idxCandadoDS] ?? "" : "").trim();
  const sec = idxSec >= 0 ? row[idxSec] : "";
  const ped = idxPed >= 0 ? row[idxPed] : "";
  const cant = parseFloat(idxCantDS >= 0 ? row[idxCantDS] : 0) || 0;
  const val = parseFloat(idxValDS >= 0 ? row[idxValDS] : 0) || 0;
  dsRows.push({ candado, sec, ped, cant, val, _i: i + 1 });
}

const dsCandadosUnicos = new Map();
for (const r of dsRows) {
  if (r.candado) dsCandadosUnicos.set(r.candado, r);
}
console.log("\nDS: filas con datos:", dsRows.length);
console.log("DS: valores distintos de SecuenciaFraccion:", [...new Set(dsRows.map((r) => String(r.sec).trim()))].filter(Boolean));
console.log("DS: candados distintos (candado → sec):", dsCandadosUnicos.size);
for (const [c, r] of dsCandadosUnicos) {
  console.log("   ", c.slice(0, 50) + (c.length > 50 ? "..." : ""), "→ SecuenciaFraccion:", r.sec);
}
if (dsRows.length <= 15) {
  console.log("\nDS todas las filas (candado, sec):");
  dsRows.forEach((r, i) => console.log("   Fila", r._i, "| candado:", (r.candado || "(vacío)").slice(0, 45), "| sec:", r.sec));
}

// ─── Layout ─────────────────────────────────────────────────────────────
const layAoa = XLSX.utils.sheet_to_json(wb.Sheets["Layout"], { header: 1, defval: "" });
const KNOWN = new Set(["pedimento", "fraccionnico", "cantidadcomercial", "cantidad_comercial", "valormpdolares", "notas", "descripcion", "seccalc", "candado", "clave"]);
let layHdrI = 0;
let best = 0;
for (let r = 0; r < Math.min(20, layAoa.length); r++) {
  const row = layAoa[r] || [];
  const hits = row.filter((c) => KNOWN.has(nH(String(c ?? "")))).length;
  if (hits > best) { best = hits; layHdrI = r; }
}
const layHdr = (layAoa[layHdrI] || []).map((c) => String(c ?? "").trim());
const idxCandadoLay = layHdr.findIndex((h) => nH(h) === "candado" || nH(h) === "candadods551" || nH(h) === "candado551" || nH(h) === "clave");
const idxCantLay = layHdr.findIndex((h) => nH(h) === "cantidadcomercial" || nH(h) === "cantidad_comercial");
let idxValLay = layHdr.findIndex((h) => nH(h) === "valormpdolares" || nH(h) === "valordolares" || nH(h) === "valor_me");
if (idxValLay < 0) idxValLay = layHdr.findIndex((h) => String(h).toLowerCase().includes("valor"));

console.log("\n=== LAYOUT ===");
console.log("Fila encabezado:", layHdrI + 1);
console.log("Col candado:", idxCandadoLay, "| cantidad:", idxCantLay, "| valor:", idxValLay);

const layoutCandados = [];
for (let i = layHdrI + 1; i < layAoa.length; i++) {
  const row = layAoa[i];
  if (row.every((c) => c === "" || c == null)) continue;
  const candado = String(idxCandadoLay >= 0 ? row[idxCandadoLay] ?? "" : "").trim();
  const cant = parseFloat(idxCantLay >= 0 ? row[idxCantLay] : 0) || 0;
  const val = parseFloat(idxValLay >= 0 ? row[idxValLay] : 0) || 0;
  layoutCandados.push({ candado, cant, val, _row: i + 1 });
}

const layCandadosUnicos = new Map();
const laySumByCandado = new Map();
for (const r of layoutCandados) {
  if (!layCandadosUnicos.has(r.candado)) { layCandadosUnicos.set(r.candado, []); laySumByCandado.set(r.candado, { cant: 0, val: 0 }); }
  layCandadosUnicos.get(r.candado).push(r._row);
  const s = laySumByCandado.get(r.candado);
  s.cant += r.cant;
  s.val += r.val;
}

console.log("Layout: filas con datos:", layoutCandados.length);
console.log("Layout: candados distintos:", layCandadosUnicos.size);
const layoutCandadosList = [...layCandadosUnicos.keys()].filter(Boolean);
console.log("Layout: lista de candados únicos (primeros 20):", layoutCandadosList.slice(0, 20));

// Match
const conMatch = layoutCandadosList.filter((c) => dsCandadosUnicos.has(c));
const sinMatch = layoutCandadosList.filter((c) => !dsCandadosUnicos.has(c));
console.log("\n=== CRUCE CANDADO ===");
console.log("Layout candados que SÍ existen en DS:", conMatch.length, conMatch.slice(0, 10));
console.log("Layout candados que NO existen en DS:", sinMatch.length, sinMatch.slice(0, 10));

if (conMatch.length > 0) {
  console.log("\n=== CUADRE POR CANDADO (Layout suma vs DS) — tolerancia ±1 cant / ±5 USD ===");
  for (const c of conMatch) {
    const dsR = dsCandadosUnicos.get(c);
    const nFilas = layCandadosUnicos.get(c).length;
    const sumL = laySumByCandado.get(c) || { cant: 0, val: 0 };
    const diffC = sumL.cant - dsR.cant;
    const diffV = sumL.val - dsR.val;
    const okC = Math.abs(diffC) <= 1;
    const okV = Math.abs(diffV) <= 5;
    const cuadra = okC && okV;
    console.log("   Candado:", c.slice(0, 50), "→ Sec:", dsR.sec, "| filas Lay:", nFilas);
    console.log("      Layout suma: Cant=" + sumL.cant.toLocaleString() + " Val=" + sumL.val.toFixed(2));
    console.log("      DS:          Cant=" + dsR.cant.toLocaleString() + " Val=" + (dsR.val || 0).toFixed(2));
    console.log("      Diferencia:  ΔCant=" + diffC + " ΔVal=" + diffV.toFixed(2), cuadra ? "✓ CUADRA" : "✗ NO CUADRA");
  }
}

console.log("\n--- Conclusión ---");
if (dsCandadosUnicos.size === 1 && dsRows.length > 1) {
  console.log("PROBLEMA: En el DS hay", dsRows.length, "filas pero solo 1 candado distinto (las demás están vacías o repetidas). Solo se puede asignar 1 secuencia.");
}
if (layoutCandadosList.filter(Boolean).length === 1 && layoutCandados.length > 1) {
  console.log("PROBLEMA: En el Layout todas las filas tienen el mismo candado. Solo se asigna una secuencia a todas.");
}
if (sinMatch.length === layoutCandadosList.filter(Boolean).length && layoutCandadosList.filter(Boolean).length > 0) {
  console.log("PROBLEMA: Ningún candado del Layout coincide con el DS. Revisar formato del candado (columnas AJ vs DF).");
}
