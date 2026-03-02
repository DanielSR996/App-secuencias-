/* eslint-disable no-console */
const XLSX = require("xlsx");

function s(v) { return String(v ?? "").trim(); }
function nFrac(v) { return s(v).replace(/^0+/, "") || "0"; }
function nNum(v) {
  const x = typeof v === "number" ? v : parseFloat(String(v).replace(/,/g, ""));
  return Number.isFinite(x) ? x : 0;
}

function indexOfHeader(headers, ...names) {
  // Importante: preferir el orden de 'names' (prioridad funcional),
  // no el orden en el que aparecen las columnas en el Excel.
  const normHeaders = headers.map((h) => String(h ?? "").trim());
  for (const name of names) {
    const n = String(name ?? "").trim();
    const idx = normHeaders.findIndex((h) => h === n);
    if (idx >= 0) return idx;
  }
  return -1;
}

function ensureColumn(headers, name) {
  let idx = indexOfHeader(headers, name);
  if (idx >= 0) return idx;
  headers.push(name);
  return headers.length - 1;
}

function main() {
  const input = process.argv[2];
  const output = process.argv[3] || input.replace(/\.xlsx?$/i, "") + "_CORREGIDO.xlsx";
  if (!input) {
    console.error("Uso: node fix_dani.cjs \"C:\\ruta\\PEDIMENTOS VALIDADOR DANI.xlsx\" [salida.xlsx]");
    process.exit(2);
  }

  const wb = XLSX.readFile(input, { cellDates: true, raw: true });

  const layoutName =
    wb.SheetNames.find((n) => String(n).toLowerCase() === "layout") ||
    wb.SheetNames.find((n) => String(n).toLowerCase().includes("layout")) ||
    "layout";
  const dsName =
    wb.SheetNames.find((n) => String(n).toLowerCase() === "ds") ||
    wb.SheetNames.find((n) => String(n).toLowerCase().includes("551")) ||
    "ds";

  if (!wb.Sheets[layoutName]) throw new Error(`No encontré hoja de Layout (busqué 'layout'): ${layoutName}`);
  if (!wb.Sheets[dsName]) throw new Error(`No encontré hoja data stage/551 (busqué 'ds'): ${dsName}`);

  const lay = XLSX.utils.sheet_to_json(wb.Sheets[layoutName], { header: 1, defval: "" });
  const ds  = XLSX.utils.sheet_to_json(wb.Sheets[dsName], { header: 1, defval: "" });
  if (!lay.length || !ds.length) throw new Error("Alguna hoja está vacía.");

  const hL = lay[0].map((x) => s(x));
  const hD = ds[0].map((x) => s(x));

  // Layout (DANI) — columnas reales
  const L_PED  = indexOfHeader(hL, "Pedimento", "pedimento", "Pedimento2", "pedimento1");
  const L_FRAC = indexOfHeader(hL, "FraccionNico", "Fraccion Nico", "fraccionnico");
  const L_PAIS = indexOfHeader(hL, "PaisOrigen", "pais_origen", "Pais Origen");
  const L_CANT = indexOfHeader(hL, "CantidadSaldo", "cantidad_comercial", "cantidad_umc", "cantidad_umc ");
  const L_VAL  = indexOfHeader(hL, "VCUSD", "ValorTotalDolares", "valor_total_dlls", "ValorMPDolares");
  if (L_PED < 0 || L_FRAC < 0 || L_CANT < 0 || L_VAL < 0) {
    throw new Error(
      `Faltan columnas en Layout. Detecté idx: ped=${L_PED}, frac=${L_FRAC}, cant=${L_CANT}, val=${L_VAL}. ` +
      `Encabezados esperados: pedimento, FraccionNico, cantidad_comercial y ValorTotalDolares (o VCUSD).`
    );
  }

  // Columna destino: secuencias (CV) o SecuenciaPed
  const L_OUT = ensureColumn(hL, indexOfHeader(hL, "secuencias") >= 0 ? "secuencias" : "SecuenciaPed");
  const L_NOTAS = ensureColumn(hL, "NOTAS_ASIGNACION");

  // 551/DataStage
  const D_PED2 = indexOfHeader(hD, "Pedimento2", "Pedimento");
  const D_FRAC = indexOfHeader(hD, "Fraccion");
  const D_PAIS = indexOfHeader(hD, "PaisOrigenDestino");
  const D_CANT = indexOfHeader(hD, "CantidadUMComercial");
  const D_VAL  = indexOfHeader(hD, "ValorDolares");
  const D_SEQ  = indexOfHeader(hD, "SecuenciaFraccion");
  if (D_PED2 < 0 || D_FRAC < 0 || D_CANT < 0 || D_VAL < 0 || D_SEQ < 0) {
    throw new Error(`Faltan columnas en ds/551. Detecté idx: ped2=${D_PED2}, frac=${D_FRAC}, cant=${D_CANT}, val=${D_VAL}, seq=${D_SEQ}`);
  }

  // Determina si tiene sentido usar país (si en ds hay >1 país)
  const dsPaisSet = new Set();
  for (let i = 1; i < ds.length; i++) {
    const r = ds[i];
    if (!r || r.every((c) => c === "" || c == null)) continue;
    const p = s(r[D_PAIS]);
    if (p) dsPaisSet.add(p);
  }
  const usePais = dsPaisSet.size > 1 && D_PAIS >= 0;
  console.log("Países únicos en ds:", dsPaisSet.size, "->", [...dsPaisSet].slice(0, 10));
  console.log("Usar país en match:", usePais);

  // Index ds por claves exactas (ped2 + frac + (pais?) + cant + val) → lista de rows
  const tolVal = 0.05;
  const key = (ped2, frac, pais, cant, val) =>
    [s(ped2), nFrac(frac), usePais ? s(pais) : "", String(nNum(cant)), String(Math.round(nNum(val) * 100) / 100)].join("|||");

  const dsMap = new Map();
  for (let i = 1; i < ds.length; i++) {
    const r = ds[i];
    if (!r || r.every((c) => c === "" || c == null)) continue;
    const k = key(r[D_PED2], r[D_FRAC], r[D_PAIS], r[D_CANT], r[D_VAL]);
    if (!dsMap.has(k)) dsMap.set(k, []);
    dsMap.get(k).push({ rowIdx: i, seq: r[D_SEQ], cant: nNum(r[D_CANT]), val: nNum(r[D_VAL]), pais: s(r[D_PAIS]) });
  }

  // Para resolver duplicados, marcamos ds rows ya utilizados
  const usedDsRows = new Set();

  // Actualiza encabezados si agregamos columnas
  lay[0] = hL;

  let matched = 0, unmatched = 0, ambiguous = 0;
  const unmatchedRows = [];

  for (let i = 1; i < lay.length; i++) {
    const r = lay[i];
    if (!r || r.every((c) => c === "" || c == null)) continue;

    // Asegura que la fila tenga longitud de encabezados (por si agregamos columnas)
    if (r.length < hL.length) r.length = hL.length;

    const ped  = r[L_PED];
    const frac = r[L_FRAC];
    const pais = L_PAIS >= 0 ? r[L_PAIS] : "";
    const cant = nNum(r[L_CANT]);
    const val  = nNum(r[L_VAL]);

    // Si no hay datos mínimos, saltar
    if (!s(ped) || !s(frac) || cant === 0 || val === 0) {
      r[L_NOTAS] = "Sin datos mínimos (ped/frac/cant/valor).";
      continue;
    }

    // Match exacto por clave (valor redondeado a centavos)
    const k = key(ped, frac, pais, cant, val);
    const cands = (dsMap.get(k) || []).filter((x) => Math.abs(x.val - val) <= tolVal);

    if (!cands.length) {
      // Segundo intento: si el país del layout no existe en ds, ignóralo (ya está ignorado si usePais=false)
      unmatched++;
      r[L_OUT] = "";
      r[L_NOTAS] = "Sin match exacto en ds (por ped+frac+cant+valor).";
      unmatchedRows.push({ row: i + 1, ped: s(ped), frac: nFrac(frac), pais: s(pais), cant, val });
      continue;
    }

    // Resolver duplicados con 'first unused ds row'
    let pick = cands.find((c) => !usedDsRows.has(c.rowIdx)) || cands[0];
    if (cands.length > 1) ambiguous++;
    usedDsRows.add(pick.rowIdx);

    r[L_OUT] = pick.seq;
    r[L_NOTAS] = cands.length > 1
      ? `Match exacto. Había ${cands.length} candidatos iguales; se asignó el primero disponible.`
      : "Match exacto (ped+frac+cant+valor).";
    matched++;
  }

  // Re-crear la hoja layout con los cambios
  wb.Sheets[layoutName] = XLSX.utils.aoa_to_sheet(lay);

  // Reporte simple
  const rep = [
    ["REPORTE — CORRECCIÓN DE SECUENCIAS (DANI)"],
    ["Archivo", input],
    ["Layout", layoutName],
    ["ds/551", dsName],
    ["Matched", matched],
    ["Unmatched", unmatched],
    ["Ambiguous (duplicados en ds)", ambiguous],
    [],
    ["UNMATCHED (primeros 50)"],
    ["Row", "Pedimento", "Fraccion", "Pais", "Cantidad", "Valor"],
    ...unmatchedRows.slice(0, 50).map((u) => [u.row, u.ped, u.frac, u.pais, u.cant, u.val]),
  ];
  const wsRep = XLSX.utils.aoa_to_sheet(rep);
  XLSX.utils.book_append_sheet(wb, wsRep, "Reporte_Correccion");

  XLSX.writeFile(wb, output);
  console.log("Salida:", output);
  console.log({ matched, unmatched, ambiguous });
}

main();

