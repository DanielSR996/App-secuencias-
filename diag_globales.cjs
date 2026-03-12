/**
 * Diagnóstico de totales globales: lee el mismo Excel que la app y muestra
 * cómo se calculan cantidad y valor (Layout vs 551/DS) para detectar por qué
 * "los valores totales" y "el conteo a nivel global" pueden no cuadrar.
 *
 * Soporta:
 * - Formato estándar: hojas "Layout" + "551"
 * - Formato 2020: hojas "Layout" + "DS" (el cuadre global EXCLUYE filas con NO INCLUIR en Notas)
 *
 * Uso: node diag_globales.cjs "C:\ruta\archivo.xlsx"
 */
const XLSX = require("xlsx");

const FILE = process.argv[2];
if (!FILE) {
  console.error("Uso: node diag_globales.cjs \"C:\\ruta\\archivo.xlsx\"");
  process.exit(1);
}

const wb = XLSX.readFile(FILE, { cellStyles: false, raw: true });
console.log("Hojas en el archivo:", wb.SheetNames.join(", "));

const is2020 = wb.SheetNames.includes("DS") && wb.SheetNames.includes("Layout");
const dataSheetName = is2020 ? "DS" : (wb.SheetNames.includes("551") ? "551" : null);
if (!wb.Sheets["Layout"] || !dataSheetName || !wb.Sheets[dataSheetName]) {
  console.error("Se requieren hojas 'Layout' y '551' (o 'DS' para formato 2020).");
  process.exit(1);
}
console.log("Formato detectado:", is2020 ? "2020 (Layout + DS)" : "estándar (Layout + 551)");

// Normalizar nombre de columna para búsqueda
function n(s) {
  return String(s ?? "").toLowerCase().replace(/\s/g, "").replace(/_/g, "");
}

// ─── LAYOUT ─────────────────────────────────────────────────────────────────
const layoutAoa = XLSX.utils.sheet_to_json(wb.Sheets["Layout"], { header: 1, defval: "" });
// En formato 2020 el encabezado puede no estar en la fila 0; buscar la fila con más columnas conocidas
const KNOWN_LAYOUT = new Set(["pedimento", "fraccionnico", "cantidadcomercial", "cantidad_comercial", "valormpdolares", "valordolares", "notas", "descripcion", "seccalc"]);
let hdrRowI = 0;
if (is2020 && layoutAoa.length > 1) {
  let best = 0;
  for (let r = 0; r < Math.min(20, layoutAoa.length); r++) {
    const row = layoutAoa[r] || [];
    const hits = row.filter((c) => KNOWN_LAYOUT.has(n(c))).length;
    if (hits > best) { best = hits; hdrRowI = r; }
  }
}
const layoutHeaders = (layoutAoa[hdrRowI] || []).map((c) => String(c ?? "").trim());
const layoutDataStart = is2020 ? hdrRowI + 1 : 1;

let cantCol, vcusdCol, pedCol, notasCol;
if (is2020) {
  cantCol = layoutHeaders.findIndex((h) => n(h) === "cantidadcomercial" || n(h) === "cantidad_comercial" || n(h) === "cantidadumc" || n(h) === "cantidad");
  vcusdCol = layoutHeaders.findIndex((h) => n(h) === "valormpdolares" || n(h) === "valordolares" || n(h) === "valor_me" || n(h) === "valorme");
  if (vcusdCol < 0) vcusdCol = layoutHeaders.findIndex((h) => String(h).toLowerCase().includes("valor"));
  notasCol = layoutHeaders.findIndex((h) => n(h) === "notas");
} else {
  cantCol = layoutHeaders.findIndex((h) => n(h) === "cantidadsaldo");
  vcusdCol = layoutHeaders.findIndex((h) => n(h) === "vcusd");
}
pedCol = layoutHeaders.findIndex((h) => n(h) === "pedimento");

console.log("\n--- LAYOUT ---");
if (is2020) console.log("Fila de encabezado detectada:", hdrRowI + 1);
console.log("Total filas en hoja (según Excel):", layoutAoa.length);
console.log("Columnas (primeras 22):", layoutHeaders.slice(0, 22).join(" | "));
// Listar TODAS las columnas que parecen "valor" o "cantidad" y su suma (para ver cuál tiene 7,635,950 vs 2,379,350)
if (is2020) {
  console.log("\n  Columnas que contienen 'valor' o 'cantidad' y su SUMA (para detectar cuál usar):");
  for (let c = 0; c < layoutHeaders.length; c++) {
    const name = String(layoutHeaders[c] || "").trim();
    if (!name) continue;
    const nn = n(name);
    if (nn.includes("valor") || nn.includes("cantidad")) {
      let sum = 0;
      for (let i = layoutDataStart; i < layoutAoa.length; i++) {
        const row = layoutAoa[i];
        if (row.every((cell) => cell === "" || cell == null)) continue;
        sum += parseFloat(row[c]) || 0;
      }
      console.log("    Col", c, ":", name, "  →  Suma =", sum.toLocaleString("es-MX", { maximumFractionDigits: 2 }));
    }
  }
}
console.log("\nÍndice cantidad (usado):", cantCol, "  valor USD (usado):", vcusdCol, vcusdCol >= 0 ? "(" + (layoutHeaders[vcusdCol] || "").trim() + ")" : "", "  pedimento:", pedCol, is2020 ? "  notas:" + notasCol : "");

let layoutRows = 0;
let layoutSumCant = 0;
let layoutSumVal = 0;
let layoutSumCantIncluir = 0;
let layoutSumValIncluir = 0;
let filasNoIncluir = 0;
let filasConTexto551NoAsignadas = 0;

for (let i = layoutDataStart; i < layoutAoa.length; i++) {
  const row = layoutAoa[i];
  if (row.every((c) => c === "" || c == null)) continue;
  layoutRows++;
  const cant = parseFloat(cantCol >= 0 ? row[cantCol] : 0) || 0;
  const val = parseFloat(vcusdCol >= 0 ? row[vcusdCol] : 0) || 0;
  layoutSumCant += cant;
  layoutSumVal += val;

  const ped = String(pedCol >= 0 ? row[pedCol] ?? "" : "").trim();
  if (ped && ped.includes("SECUENCIAS DEL 551")) filasConTexto551NoAsignadas++;

  if (is2020 && notasCol >= 0) {
    const notas = String(row[notasCol] ?? "").toUpperCase();
    const noIncluir = notas.includes("NO INCLUIR");
    if (noIncluir) filasNoIncluir++;
    else {
      layoutSumCantIncluir += cant;
      layoutSumValIncluir += val;
    }
  } else {
    layoutSumCantIncluir += cant;
    layoutSumValIncluir += val;
  }
}

console.log("Filas Layout (no vacías):", layoutRows);
console.log("  Suma cantidad (TODAS las filas):     ", layoutSumCant);
console.log("  Suma valor USD (TODAS las filas):    ", layoutSumVal);
if (is2020 && filasNoIncluir > 0) {
  console.log("  Filas con 'NO INCLUIR' en Notas:     ", filasNoIncluir);
  console.log("  Suma cantidad (sin NO INCLUIR):     ", layoutSumCantIncluir, "  ← esto es lo que usa la app para el cuadre");
  console.log("  Suma valor USD (sin NO INCLUIR):    ", layoutSumValIncluir);
} else if (is2020) {
  layoutSumCantIncluir = layoutSumCant;
  layoutSumValIncluir = layoutSumVal;
}
if (filasConTexto551NoAsignadas > 0) {
  console.log("  ⚠ Hay", filasConTexto551NoAsignadas, "fila(s) con 'SECUENCIAS DEL 551' en Pedimento (posible resultado anterior).");
}

// ─── 551 / DS ───────────────────────────────────────────────────────────────
const aoaData = XLSX.utils.sheet_to_json(wb.Sheets[dataSheetName], { header: 1, defval: "" });
const hData = (aoaData[0] || []).map((c) => String(c ?? "").trim());
const idxCant = hData.findIndex((h) => String(h).trim() === "CantidadUMComercial");
let idxVal = hData.findIndex((h) => String(h).trim() === "ValorDolares");
if (idxVal < 0) {
  const v = hData.findIndex((h) => String(h).toLowerCase().includes("valor"));
  if (v >= 0) idxVal = v;
}

// En DS 2020 el valor puede estar en "ValorDolares", "ValorAgregado", "Valor usd redondeado", etc.
if (idxVal < 0 && is2020) {
  idxVal = hData.findIndex((h) => n(h) === "valoragregado" || n(h) === "valorusdredondeado" || n(h) === "valoraduana");
}

console.log("\n---", dataSheetName, "---");
console.log("Columnas (primeras 14):", hData.slice(0, 14).join(" | "));
console.log("Índice CantidadUMComercial:", idxCant, "  Valor (Dolares/Agregado):", idxVal, idxVal >= 0 ? "(" + (hData[idxVal] || "").trim() + ")" : "");

let rowsData = 0;
let sumCantData = 0;
let sumValData = 0;
for (let i = 1; i < aoaData.length; i++) {
  const row = aoaData[i];
  if (row.every((c) => c === "" || c == null)) continue;
  rowsData++;
  sumCantData += parseFloat(idxCant >= 0 ? row[idxCant] : 0) || 0;
  sumValData += parseFloat(idxVal >= 0 ? row[idxVal] : 0) || 0;
}

console.log("Filas " + dataSheetName + " (no vacías):", rowsData);
console.log("  Suma CantidadUMComercial:", sumCantData);
console.log("  Suma ValorDolares:       ", sumValData);

// ─── CUADRE ───────────────────────────────────────────────────────────────────
const layoutCantParaCuadre = is2020 ? layoutSumCantIncluir : layoutSumCant;
const layoutValParaCuadre = is2020 ? layoutSumValIncluir : layoutSumVal;

console.log("\n--- CUADRE GLOBAL (lo que usa la app) ---");
console.log("  Layout total cantidad:", layoutCantParaCuadre, "  |  " + dataSheetName + " total cantidad:", sumCantData, "  → diferencia:", layoutCantParaCuadre - sumCantData);
console.log("  Layout total valor USD:", layoutValParaCuadre, "  |  " + dataSheetName + " total valor USD:", sumValData, "  → diferencia:", layoutValParaCuadre - sumValData);
const diffCant = layoutCantParaCuadre - sumCantData;
const diffVal = layoutValParaCuadre - sumValData;
console.log("  ¿Cuadra cantidad (±1)?", Math.abs(diffCant) <= 1 ? "SÍ" : "NO");
console.log("  ¿Cuadra valor (±5)?   ", Math.abs(diffVal) <= 5 ? "SÍ" : "NO");

if (is2020 && filasNoIncluir > 0 && Math.abs(layoutSumCant - layoutSumCantIncluir) > 0.001) {
  console.log("\n  ⚠ En formato 2020 la app solo suma filas SIN 'NO INCLUIR'. Si cuentas a mano todas las filas, los totales no coincidirán.");
}

console.log("\n--- RESUMEN ---");
console.log("  • La app asigna secuencias usando cantidad Y valor (con tolerancias).");
console.log("  • El cuadre global en formato 2020 usa solo filas del Layout que NO tengan 'NO INCLUIR' en Notas.");
if (Math.abs(diffCant) <= 1 && Math.abs(diffVal) > 5) {
  console.log("  • En este archivo: cantidad cuadra a nivel global; el valor no cuadra (p. ej. DS sin columna 'ValorDolares' o columna vacía).");
}
