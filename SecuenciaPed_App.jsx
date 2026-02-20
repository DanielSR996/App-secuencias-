import { useState, useCallback, useRef } from "react";
import * as XLSX from "xlsx-js-style";

// ─── LECTURA EXCEL (hoja 551 con columnas duplicadas / nombres con espacios) ───
// Columnas necesarias del 551 (se busca la PRIMERA ocurrencia de cada nombre, por eso
// se usa header:1 y búsqueda por nombre trimado en lugar de sheet_to_json directamente).
const COLS_551 = [
  "Pedimento",
  "Fraccion",           // clave de cruce con FraccionNico del Layout
  "SecuenciaFraccion",  // valor a asignar en Layout.SecuenciaPed
  "PaisOrigenDestino",
  "CantidadUMComercial",
  "ValorDolares",
  "DescripcionMercancia",
];

function firstIndexByHeader(headerRow, colName) {
  const s = String(colName || "").trim();
  for (let i = 0; i < headerRow.length; i++) {
    if (String(headerRow[i] || "").trim() === s) return i;
  }
  return -1;
}

/** Lee la hoja 551 tomando la PRIMERA columna que coincida con cada nombre (maneja espacios y duplicados). */
function read551Sheet(sheet) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  if (!rows.length) return [];
  const headerRow = rows[0].map((c) => String(c ?? "").trim());
  const indices = {};
  for (const col of COLS_551) {
    const idx = firstIndexByHeader(headerRow, col);
    if (idx >= 0) indices[col] = idx;
  }
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (row.every((c) => c === "" || c == null)) continue; // saltar filas vacías
    const obj = {};
    for (const [col, idx] of Object.entries(indices)) {
      obj[col] = row[idx];
    }
    out.push(obj);
  }
  return out;
}

/** Lee el Layout y normaliza los nombres de columnas con espacios. */
function readLayoutSheet(sheet) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  if (!rows.length) return [];
  // Construir header normalizado (trim), guardando el índice de cada columna
  const rawHeaders = rows[0];
  const headers = rawHeaders.map((c) => String(c ?? "").trim());
  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (row.every((c) => c === "" || c == null)) continue;
    const obj = {};
    for (let j = 0; j < headers.length; j++) {
      if (headers[j]) obj[headers[j]] = row[j];
    }
    out.push(obj);
  }
  return out;
}

function resolve551SheetName(wb) {
  const names = wb.SheetNames || [];
  if (names.includes("551")) return "551";
  const lower = names.map((n) => String(n).toLowerCase());
  const i = lower.findIndex((n) => n.includes("data") && n.includes("stage"));
  if (i >= 0) return names[i];
  const j = lower.findIndex((n) => n === "datastage" || n.includes("551"));
  if (j >= 0) return names[j];
  return null;
}

// ─── MATCHING ENGINE ──────────────────────────────────────────────────────────
// Lógica real de cruce IMMEX:
//   Layout: Pedimento + FraccionNico + PaisOrigen → suma CantidadSaldo y VCUSD
//   551:    Pedimento + Fraccion    + PaisOrigenDestino → CantidadUMComercial y ValorDolares
//   Resultado: SecuenciaFraccion del 551 se asigna a SecuenciaPed del Layout

function groupBy(rows, keys) {
  const map = new Map();
  for (const row of rows) {
    const k = keys.map((k) => String(row[k] ?? "")).join("|||");
    if (!map.has(k)) map.set(k, { key: k, keyVals: keys.map((k) => row[k]), rows: [] });
    map.get(k).rows.push(row);
  }
  return [...map.values()];
}

function sumGroup(rows, cantCol, vcusdCol) {
  return rows.reduce(
    (acc, r) => ({
      cant: acc.cant + (parseFloat(r[cantCol]) || 0),
      vcusd: acc.vcusd + (parseFloat(r[vcusdCol]) || 0),
    }),
    { cant: 0, vcusd: 0 }
  );
}

function tryMatch(candidates, sumCant, sumVCUSD, tolCant = 1, tolVCUSD = 2) {
  for (const r of candidates) {
    const c551 = parseFloat(r["CantidadUMComercial"]) || 0;
    const v551 = parseFloat(r["ValorDolares"]) || 0;
    if (Math.abs(sumCant - c551) <= tolCant && Math.abs(sumVCUSD - v551) <= tolVCUSD) {
      return { seq: r["SecuenciaFraccion"], r551: r };
    }
  }
  return null;
}

function runCascade(layoutRows, s551Rows) {
  // ── Columnas del Layout (ya vienen normalizadas por readLayoutSheet) ─────
  const L_PED   = "Pedimento";
  const L_FRAC  = "FraccionNico";
  const L_PAIS  = "PaisOrigen";
  const L_CANT  = "CantidadSaldo";
  const L_VCUSD = "VCUSD";
  const L_SEC   = "SecuenciaPed";

  // ── Columnas del 551 ─────────────────────────────────────────────────────
  const S_PED  = "Pedimento";
  const S_FRAC = "Fraccion";
  const S_PAIS = "PaisOrigenDestino";
  const S_SEQ  = "SecuenciaFraccion";

  const nFrac = (v) => String(v ?? "").trim().replace(/^0+/, "") || "0";

  // Normaliza SecuenciaPed: si es texto (ej "Sin registro en 551") usa cadena vacía
  // para que esas filas se agrupen juntas en E2/E4 en lugar de cada una por separado
  const nSec = (v) => {
    const n = parseFloat(v);
    return isNaN(n) ? "" : String(Math.round(n));
  };

  const layout = layoutRows.map((r, i) => ({
    ...r,
    _idx: i,
    _frac: nFrac(r[L_FRAC]),
    _sec:  nSec(r[L_SEC]),
  }));

  const s551 = s551Rows.map((r, i) => ({
    ...r,
    _frac:    nFrac(r[S_FRAC]),
    _551idx:  i,
  }));

  // ── Set de todas las fracciones en el 551 (para diagnóstico) ────────────
  const fracSet551 = new Set(s551.map((r) => r._frac));

  // ── Lookups del 551 ───────────────────────────────────────────────────────
  const lookupPFP = new Map();  // Pedimento + Fraccion + Pais
  const lookupPF  = new Map();  // Pedimento + Fraccion (sin país)

  for (const r of s551) {
    const k1 = `${r[S_PED]}|||${r._frac}|||${String(r[S_PAIS] ?? "").trim()}`;
    if (!lookupPFP.has(k1)) lookupPFP.set(k1, []);
    lookupPFP.get(k1).push(r);

    const k2 = `${r[S_PED]}|||${r._frac}`;
    if (!lookupPF.has(k2)) lookupPF.set(k2, []);
    lookupPF.get(k2).push(r);
  }

  // ── Tracking ──────────────────────────────────────────────────────────────
  const assignment    = new Map();
  const assigned      = new Set();
  const used551       = new Set(); // índices (_551idx) de filas del 551 usadas en algún match
  const strategyStats = { E1: 0, E2: 0, E3: 0, E4: 0, E5: 0 };

  // Almacena también el registro 551 que generó el match (para hoja de cruce)
  const assignRows = (rows, seq, strategy, r551 = null) => {
    for (const r of rows) {
      if (!assigned.has(r._idx)) {
        assignment.set(r._idx, { seq, strategy, r551 });
        assigned.add(r._idx);
        strategyStats[strategy]++;
        if (r551?._551idx !== undefined) used551.add(r551._551idx);
      }
    }
  };

  // ── E1: Pedimento + Fracción + País, cantidades exactas (±1 ud / ±2 USD) ──
  for (const g of groupBy(layout, [L_PED, "_frac", L_PAIS])) {
    const k = `${g.keyVals[0]}|||${g.keyVals[1]}|||${String(g.keyVals[2] ?? "").trim()}`;
    const cands = lookupPFP.get(k) || [];
    const { cant, vcusd } = sumGroup(g.rows, L_CANT, L_VCUSD);
    const match = tryMatch(cands, cant, vcusd, 1, 2);
    if (match) assignRows(g.rows, match.seq, "E1", match.r551);
  }

  // ── E2: Mismo Ped+Frac+País, sub-grupo por SecuenciaPed (solo numérico) ───
  for (const g of groupBy(layout.filter((r) => !assigned.has(r._idx)), [L_PED, "_frac", L_PAIS, "_sec"])) {
    const k = `${g.keyVals[0]}|||${g.keyVals[1]}|||${String(g.keyVals[2] ?? "").trim()}`;
    const cands = lookupPFP.get(k) || [];
    const { cant, vcusd } = sumGroup(g.rows, L_CANT, L_VCUSD);
    const match = tryMatch(cands, cant, vcusd, 1, 2);
    if (match) assignRows(g.rows, match.seq, "E2", match.r551);
  }

  // ── E3: Pedimento + Fracción (sin País) ───────────────────────────────────
  for (const g of groupBy(layout.filter((r) => !assigned.has(r._idx)), [L_PED, "_frac"])) {
    const k = `${g.keyVals[0]}|||${g.keyVals[1]}`;
    const cands = lookupPF.get(k) || [];
    const { cant, vcusd } = sumGroup(g.rows, L_CANT, L_VCUSD);
    const match = tryMatch(cands, cant, vcusd, 1, 2);
    if (match) assignRows(g.rows, match.seq, "E3", match.r551);
  }

  // ── E4: Ped+Frac (sin País) + sub-grupo por SecuenciaPed (solo numérico) ──
  for (const g of groupBy(layout.filter((r) => !assigned.has(r._idx)), [L_PED, "_frac", "_sec"])) {
    const k = `${g.keyVals[0]}|||${g.keyVals[1]}`;
    const cands = lookupPF.get(k) || [];
    const { cant, vcusd } = sumGroup(g.rows, L_CANT, L_VCUSD);
    const match = tryMatch(cands, cant, vcusd, 1, 2);
    if (match) assignRows(g.rows, match.seq, "E4", match.r551);
  }

  // ── E5: Tolerancia ampliada ±5% ──────────────────────────────────────────
  for (const g of groupBy(layout.filter((r) => !assigned.has(r._idx)), [L_PED, "_frac", L_PAIS, "_sec"])) {
    const k = `${g.keyVals[0]}|||${g.keyVals[1]}`;
    const cands = lookupPF.get(k) || [];
    const { cant, vcusd } = sumGroup(g.rows, L_CANT, L_VCUSD);
    const tolCant  = Math.max(2, cant  * 0.05);
    const tolVCUSD = Math.max(5, vcusd * 0.05);
    const match = tryMatch(cands, cant, vcusd, tolCant, tolVCUSD);
    if (match) assignRows(g.rows, match.seq, "E5", match.r551);
  }

  // ── Layout lookup por Ped+Frac (para diagnóstico de orphans del 551) ─────
  const layoutPF = new Map();
  for (const r of layout) {
    const k = `${r[L_PED]}|||${r._frac}`;
    if (!layoutPF.has(k)) layoutPF.set(k, []);
    layoutPF.get(k).push(r);
  }

  // ── Secuencias del 551 que NO se usaron en ningún match (orphans) ─────────
  const getOrphanReason = (r) => {
    const cant = parseFloat(r["CantidadUMComercial"]);
    const val  = parseFloat(r["ValorDolares"]);
    const cantZero = isNaN(cant) || cant === 0;
    const valZero  = isNaN(val)  || val  === 0;
    const seq  = r["SecuenciaFraccion"] ?? "?";
    const frac = r._frac ?? r[S_FRAC] ?? "?";
    const ped  = r[S_PED] ?? "?";

    if (cantZero && valZero)
      return `Sec.${seq} — Sin cantidad ni valor: CantidadUMComercial=0 y ValorDolares=0`;
    if (cantZero)
      return `Sec.${seq} — CantidadUMComercial=0 (sin cantidad registrada en el 551)`;
    if (valZero)
      return `Sec.${seq} — ValorDolares=0 (sin valor en dólares registrado en el 551)`;

    const kPF = `${ped}|||${frac}`;
    if (!layoutPF.has(kPF))
      return `Sec.${seq} — Pedimento ${ped} / Fracción ${frac} no tiene partidas en Layout`;

    // Layout sí tiene esa combinación pero no cuadran cantidades
    const layoutCands = layoutPF.get(kPF);
    const sumCant = layoutCands.reduce((a, lr) => a + (parseFloat(lr[L_CANT]) || 0), 0);
    const sumVal  = layoutCands.reduce((a, lr) => a + (parseFloat(lr[L_VCUSD]) || 0), 0);
    return `Sec.${seq} — Cantidad/Valor no coinciden: Layout(${sumCant.toFixed(0)} ud / $${sumVal.toFixed(2)}) vs 551(${cant.toFixed(0)} ud / $${val.toFixed(2)})`;
  };

  const orphan551Rows = s551
    .filter((r) => !used551.has(r._551idx))
    .map((r)  => ({ ...r, _orphanReason: getOrphanReason(r) }));

  // ── Diagnóstico por grupo sin match ──────────────────────────────────────
  const computeGroupNote = (ped, frac, pais, cant, vcusd) => {
    if (!fracSet551.has(frac)) {
      return `Fracción arancelaria ${frac} no registrada en el 551`;
    }
    const candsPF = lookupPF.get(`${ped}|||${frac}`) || [];
    if (candsPF.length === 0) {
      const otrosPed = [...new Set(s551.filter((r) => r._frac === frac).map((r) => r[S_PED]))].join(", ");
      return `Fracción ${frac} no encontrada para pedimento ${ped}. Aparece en: ${otrosPed || "ninguno"}`;
    }
    const candsPFP = lookupPFP.get(`${ped}|||${frac}|||${String(pais ?? "").trim()}`) || [];
    if (candsPFP.length === 0) {
      const paises = [...new Set(candsPF.map((r) => r[S_PAIS]))].join(", ");
      return `País no coincide. Layout: ${pais} | 551 registra: ${paises}`;
    }
    // La fracción+pedimento+país existe pero las cantidades no cuadran
    let bestDiff = Infinity, best = null;
    for (const c of candsPFP) {
      const dc = Math.abs((parseFloat(c.CantidadUMComercial) || 0) - cant);
      const dv = Math.abs((parseFloat(c.ValorDolares) || 0) - vcusd);
      if (dc + dv < bestDiff) { bestDiff = dc + dv; best = c; }
    }
    if (best) {
      const c551c = parseFloat(best.CantidadUMComercial) || 0;
      const c551v = parseFloat(best.ValorDolares) || 0;
      const diffC = (cant - c551c).toFixed(0);
      const diffV = (vcusd - c551v).toFixed(2);
      return `Suma Layout: ${cant.toFixed(0)} ud / $${vcusd.toFixed(2)} | ` +
             `Entrada 551 (seq ${best[S_SEQ]}): ${c551c.toFixed(0)} ud / $${c551v.toFixed(2)} | ` +
             `Diferencia: ${diffC > 0 ? "+" : ""}${diffC} ud / $${diffV > 0 ? "+" : ""}${diffV} ` +
             `(hay ${candsPFP.length} candidatos — requiere sub-agrupación manual)`;
    }
    return "No se encontró correspondencia exacta en 551";
  };

  // Construir mapa rowIdx → nota para filas sin match
  const rowNotes = new Map();
  const unmatchedGroups = groupBy(
    layout.filter((r) => !assigned.has(r._idx)),
    [L_PED, "_frac", L_PAIS]
  );
  for (const g of unmatchedGroups) {
    const [ped, frac, pais] = g.keyVals;
    const { cant, vcusd } = sumGroup(g.rows, L_CANT, L_VCUSD);
    const nota = computeGroupNote(ped, frac, pais, cant, vcusd);
    for (const r of g.rows) rowNotes.set(r._idx, nota);
  }

  // ── Construir lista de sin-match para la UI ───────────────────────────────
  const unmatchedFinal = layout
    .filter((r) => !assigned.has(r._idx))
    .map((r) => ({
      Descripcion:           r["Descripcion"],
      FraccionNico:          r[L_FRAC],
      PaisOrigen:            r[L_PAIS],
      SecuenciaPed_Original: r[L_SEC],
      CantidadSaldo:         r[L_CANT],
      VCUSD:                 r[L_VCUSD],
      Nota:                  rowNotes.get(r._idx) || "",
    }));

  // ── Construir datos para la hoja Cruce_Layout_vs_551 ─────────────────────
  // Un registro por GRUPO (Ped + Frac + Pais + SecuenciaPedAsignada)
  const cruceData = [];

  // Grupos ASIGNADOS: agrupar por (seq asignada + frac + pais + ped)
  const matchedGroupMap = new Map();
  for (const [rowIdx, info] of assignment) {
    const r = layout[rowIdx];
    const gk = `${r[L_PED]}|||${r._frac}|||${String(r[L_PAIS] || "").trim()}|||${info.seq}`;
    if (!matchedGroupMap.has(gk)) {
      matchedGroupMap.set(gk, { rows: [], info, firstRow: r });
    }
    matchedGroupMap.get(gk).rows.push(r);
  }

  for (const [, g] of matchedGroupMap) {
    const { cant, vcusd }   = sumGroup(g.rows, L_CANT, L_VCUSD);
    const r551              = g.info.r551;
    const cant551           = r551 ? (parseFloat(r551.CantidadUMComercial) || 0) : null;
    const val551            = r551 ? (parseFloat(r551.ValorDolares)        || 0) : null;
    const diffCant          = cant551 !== null ? cant - cant551 : null;
    const diffVal           = val551  !== null ? vcusd - val551 : null;
    const okFrac = r551 ? (String(g.firstRow[L_FRAC]).trim() === String(r551[S_FRAC] || "").trim()) : false;
    const okPais = r551 ? (String(g.firstRow[L_PAIS] || "").trim() === String(r551[S_PAIS] || "").trim()) : false;
    const okCant = diffCant !== null && Math.abs(diffCant) <= 1;
    const okVal  = diffVal  !== null && Math.abs(diffVal)  <= 2;

    // Descripciones únicas de las partes en el grupo
    const descs = [...new Set(g.rows.map((r) => r["Descripcion"]).filter(Boolean))].join(" / ");

    cruceData.push({
      tipo:       "matched",
      estrategia: g.info.strategy,
      numFilas:   g.rows.length,
      pedimento:  g.firstRow[L_PED],
      // Layout
      layoutDesc:  descs,
      layoutFrac:  String(g.firstRow[L_FRAC] || ""),
      layoutPais:  g.firstRow[L_PAIS] || "",
      layoutCant:  cant,
      layoutVCUSD: vcusd,
      secOriginal: g.firstRow[L_SEC],
      secAsignada: g.info.seq,
      // 551
      s551Secuencias:  r551 ? (r551["Secuencias"] || `${r551[S_PED]}-${r551[S_FRAC]}-${r551[S_SEQ]}`) : "",
      s551Desc:        r551 ? (r551["DescripcionMercancia"] || "") : "",
      s551Frac:        r551 ? (r551[S_FRAC] || "") : "",
      s551Pais:        r551 ? (r551[S_PAIS] || "") : "",
      s551Cant:        cant551,
      s551Val:         val551,
      // Diferencias
      diffCant, diffVal,
      okFrac, okPais, okCant, okVal,
    });
  }

  // Grupos SIN MATCH: mostrar con el mejor candidato de 551
  for (const g of unmatchedGroups) {
    const [ped, frac, pais] = g.keyVals;
    const { cant, vcusd }   = sumGroup(g.rows, L_CANT, L_VCUSD);
    const candsPF  = lookupPF.get(`${ped}|||${frac}`) || [];
    let best = null, bestDiff = Infinity;
    for (const c of candsPF) {
      const d = Math.abs((parseFloat(c.CantidadUMComercial) || 0) - cant)
              + Math.abs((parseFloat(c.ValorDolares)        || 0) - vcusd);
      if (d < bestDiff) { bestDiff = d; best = c; }
    }
    const descs = [...new Set(g.rows.map((r) => r["Descripcion"]).filter(Boolean))].join(" / ");
    cruceData.push({
      tipo:       "unmatched",
      estrategia: "SIN MATCH",
      numFilas:   g.rows.length,
      pedimento:  ped,
      layoutDesc:  descs,
      layoutFrac:  frac,
      layoutPais:  pais,
      layoutCant:  cant,
      layoutVCUSD: vcusd,
      secOriginal: g.rows[0][L_SEC],
      secAsignada: "",
      s551Secuencias: best ? (best["Secuencias"] || `${best[S_PED]}-${best[S_FRAC]}-${best[S_SEQ]}`) : "",
      s551Desc:   best ? (best["DescripcionMercancia"] || "") : "— Sin candidato en 551 —",
      s551Frac:   best ? (best[S_FRAC] || "") : "",
      s551Pais:   best ? (best[S_PAIS] || "") : "",
      s551Cant:   best ? (parseFloat(best.CantidadUMComercial) || 0) : null,
      s551Val:    best ? (parseFloat(best.ValorDolares)        || 0) : null,
      diffCant:   best ? cant - (parseFloat(best.CantidadUMComercial) || 0) : null,
      diffVal:    best ? vcusd - (parseFloat(best.ValorDolares)       || 0) : null,
      okFrac: false, okPais: false, okCant: false, okVal: false,
      nota: rowNotes.get(g.rows[0]._idx) || "",
    });
  }

  // Ordenar: primero sin-match, luego por estrategia
  cruceData.sort((a, b) => {
    if (a.tipo !== b.tipo) return a.tipo === "unmatched" ? -1 : 1;
    return (a.estrategia || "").localeCompare(b.estrategia || "");
  });

  return { assignment, strategyStats, unmatchedFinal, total: layout.length, rowNotes, cruceData, orphan551Rows };
}

// ─── EXCEL BUILDER ────────────────────────────────────────────────────────────
function buildOutputExcel(workbook, layoutSheet, sheet551, sheet551Name, assignment, unmatchedFinal, stats, total, rowNotes, cruceData, orphan551Rows) {
  const wb = XLSX.utils.book_new();

  // ── Datos originales del Layout ─────────────────────────────────────────
  const layoutData = XLSX.utils.sheet_to_json(layoutSheet, { header: 1 });
  const rawHeaders  = layoutData[0] || [];

  // Buscar SecuenciaPed con trim (tolerante a espacios en el Excel)
  const secIdx  = rawHeaders.findIndex((h) => String(h ?? "").trim() === "SecuenciaPed");
  const notasIdx = rawHeaders.length;  // nueva columna al final
  const headers  = [...rawHeaders, "Notas"];

  // Normaliza SecuenciaPed para comparación (número o texto limpio)
  const normSeq = (v) => {
    const n = parseFloat(v);
    return isNaN(n) ? String(v ?? "").trim() : String(Math.round(n));
  };

  // ── Construir filas + registrar cambios ──────────────────────────────────
  // changeMap: rowIdx → { original, nuevo, tipoNota }
  //   tipoNota: "nuevo"    → la celda estaba vacía/texto y ahora tiene valor
  //             "cambio"   → el valor cambió de un número a otro
  //             "igual"    → el valor nuevo es igual al original (sin marcado)
  //             "sinmatch" → no se asignó secuencia
  const changeMap = new Map();
  const updatedRows = [headers];

  for (let i = 1; i < layoutData.length; i++) {
    const row      = [...layoutData[i]];
    const rowIdx   = i - 1;
    while (row.length <= notasIdx) row.push("");

    const originalRaw = secIdx >= 0 ? (layoutData[i][secIdx] ?? "") : "";
    const originalStr = normSeq(originalRaw);

    if (assignment.has(rowIdx)) {
      const rawSeq  = assignment.get(rowIdx).seq;
      const newSeq  = parseFloat(rawSeq) || rawSeq;
      const newStr  = normSeq(rawSeq);

      row[secIdx] = newSeq;

      if (newStr !== originalStr) {
        // Hubo cambio real: ¿era vacío/texto antes o era otro número?
        const wasEmpty = (originalStr === "" || isNaN(parseFloat(originalRaw)));
        const tipo     = wasEmpty ? "nuevo" : "cambio";
        const nota     = wasEmpty
          ? `Secuencia asignada por la app: ${newStr}`
          : `Secuencia modificada: ${originalStr} → ${newStr}`;
        row[notasIdx] = nota;
        changeMap.set(rowIdx, { original: originalStr, nuevo: newStr, tipo, nota });
      } else {
        // Mismo valor que el original — sin nota, sin color
        row[notasIdx] = "";
        changeMap.set(rowIdx, { original: originalStr, nuevo: newStr, tipo: "igual" });
      }
    } else {
      // Sin match: limpiar SecuenciaPed + nota diagnóstico
      row[secIdx]   = "";
      row[notasIdx] = rowNotes ? (rowNotes.get(rowIdx) || "") : "";
      changeMap.set(rowIdx, { tipo: "sinmatch" });
    }
    updatedRows.push(row);
  }

  // ── Filas extra al final: secuencias del 551 que no se asignaron al Layout ─
  const orphan551StartRow = updatedRows.length; // índice de fila en la hoja (0-based)

  if (orphan551Rows && orphan551Rows.length > 0) {
    // Fila separadora vacía
    updatedRows.push(Array(headers.length).fill(""));

    // Encabezado de sección
    const sectionHdr = Array(headers.length).fill("");
    sectionHdr[0] = `▼ SECUENCIAS DEL 551 NO ASIGNADAS AL LAYOUT  (${orphan551Rows.length} registros)`;
    updatedRows.push(sectionHdr);

    // Localizar índices de columnas del Layout para mapear los datos del 551
    const pedIdxL   = rawHeaders.findIndex((h) => String(h ?? "").trim() === "Pedimento");
    const fracIdxL  = rawHeaders.findIndex((h) => String(h ?? "").trim() === "FraccionNico");
    const paisIdxL  = rawHeaders.findIndex((h) => String(h ?? "").trim() === "PaisOrigen");
    const cantIdxL  = rawHeaders.findIndex((h) => String(h ?? "").trim() === "CantidadSaldo");
    const vcusdIdxL = rawHeaders.findIndex((h) => String(h ?? "").trim() === "VCUSD");

    for (const r551 of orphan551Rows) {
      const row = Array(headers.length).fill("");
      if (secIdx   >= 0) row[secIdx]   = r551["SecuenciaFraccion"] ?? "";
      if (pedIdxL  >= 0) row[pedIdxL]  = r551["Pedimento"] ?? "";
      if (fracIdxL >= 0) row[fracIdxL] = r551["Fraccion"]  ?? "";
      if (paisIdxL >= 0) row[paisIdxL] = r551["PaisOrigenDestino"] ?? "";
      if (cantIdxL >= 0) row[cantIdxL] = parseFloat(r551["CantidadUMComercial"]) || 0;
      if (vcusdIdxL >= 0) row[vcusdIdxL] = parseFloat(r551["ValorDolares"])      || 0;
      row[notasIdx] = r551._orphanReason || "";
      updatedRows.push(row);
    }
  }

  // ── Crear worksheet y aplicar estilos ────────────────────────────────────
  const wsLayout = XLSX.utils.aoa_to_sheet(updatedRows);

  // Cabecera SecuenciaPed
  if (secIdx >= 0) {
    const hCell = XLSX.utils.encode_cell({ r: 0, c: secIdx });
    wsLayout[hCell] = {
      v: "SecuenciaPed", t: "s",
      s: { font: { bold: true, color: { rgb: "FFFFFF" } },
           fill: { patternType: "solid", fgColor: { rgb: "1B3A6B" } },
           alignment: { horizontal: "center" } },
    };
  }
  // Cabecera Notas
  {
    const hNota = XLSX.utils.encode_cell({ r: 0, c: notasIdx });
    wsLayout[hNota] = {
      v: "Notas", t: "s",
      s: { font: { bold: true, color: { rgb: "FFFFFF" } },
           fill: { patternType: "solid", fgColor: { rgb: "7B2029" } } },
    };
  }

  // Estilos reutilizables
  const S_CHANGED = {           // rojo negrita: secuencia modificada/nueva
    font: { bold: true, sz: 11, color: { rgb: "C0392B" } },
    fill: { patternType: "solid", fgColor: { rgb: "FADBD8" } },
    alignment: { horizontal: "center" },
  };
  const S_EQUAL = {             // negro normal: mismo valor, sin tocar
    font: { color: { rgb: "000000" } },
    alignment: { horizontal: "center" },
  };
  const S_EMPTY = {             // gris claro: sin match
    font: { color: { rgb: "BBBBBB" } },
    alignment: { horizontal: "center" },
  };
  const S_NOTA_CAMBIO = {       // nota de cambio: fondo rojo muy claro
    font: { bold: true, sz: 10, color: { rgb: "922B21" } },
    fill: { patternType: "solid", fgColor: { rgb: "FADBD8" } },
    alignment: { wrapText: true },
  };
  const S_NOTA_SINMATCH = {     // nota de sin-match: fondo amarillo
    font: { italic: true, sz: 10, color: { rgb: "7D6608" } },
    fill: { patternType: "solid", fgColor: { rgb: "FEF9E7" } },
    alignment: { wrapText: true },
  };

  for (let rowI = 1; rowI < updatedRows.length; rowI++) {
    const rowIdx = rowI - 1;
    const info   = changeMap.get(rowIdx);
    if (!info) continue;

    // Estilo SecuenciaPed
    if (secIdx >= 0) {
      const addr = XLSX.utils.encode_cell({ r: rowI, c: secIdx });
      if (!wsLayout[addr]) wsLayout[addr] = { t: "s", v: "" };
      if (info.tipo === "nuevo" || info.tipo === "cambio") {
        wsLayout[addr].s = S_CHANGED;      // rojo negrita
      } else if (info.tipo === "igual") {
        wsLayout[addr].s = S_EQUAL;        // negro sin resaltar
      } else {
        wsLayout[addr].s = S_EMPTY;        // gris (vacío)
      }
    }

    // Estilo Notas
    const addrN = XLSX.utils.encode_cell({ r: rowI, c: notasIdx });
    if (!wsLayout[addrN]) wsLayout[addrN] = { t: "s", v: "" };
    if (info.tipo === "nuevo" || info.tipo === "cambio") {
      wsLayout[addrN].s = S_NOTA_CAMBIO;
    } else if (info.tipo === "sinmatch" && updatedRows[rowI][notasIdx]) {
      wsLayout[addrN].s = S_NOTA_SINMATCH;
    }
  }

  // ── Estilos para filas de orphan 551 ────────────────────────────────────
  if (orphan551Rows && orphan551Rows.length > 0) {
    const S_ORPHAN_HDR = {
      font: { bold: true, sz: 11, color: { rgb: "FFFFFF" } },
      fill: { patternType: "solid", fgColor: { rgb: "1A5276" } },
      alignment: { wrapText: true },
    };
    const S_ORPHAN_SEQ = {
      font: { bold: true, color: { rgb: "0D3349" } },
      fill: { patternType: "solid", fgColor: { rgb: "AED6F1" } },
      alignment: { horizontal: "center" },
    };
    const S_ORPHAN_DATA = {
      font: { color: { rgb: "1A5276" } },
      fill: { patternType: "solid", fgColor: { rgb: "EBF5FB" } },
    };
    const S_ORPHAN_NOTA = {
      font: { italic: true, sz: 10, color: { rgb: "0D3349" } },
      fill: { patternType: "solid", fgColor: { rgb: "AED6F1" } },
      alignment: { wrapText: true },
    };

    // Fila encabezado de sección (orphan551StartRow + 1)
    const sectionSheetRow = orphan551StartRow + 1;
    for (let c = 0; c < headers.length; c++) {
      const addr = XLSX.utils.encode_cell({ r: sectionSheetRow, c });
      if (!wsLayout[addr]) wsLayout[addr] = { t: "s", v: "" };
      wsLayout[addr].s = S_ORPHAN_HDR;
    }

    // Filas de datos orphan (orphan551StartRow + 2 en adelante)
    for (let o = 0; o < orphan551Rows.length; o++) {
      const shRowI = orphan551StartRow + 2 + o;
      for (let c = 0; c < headers.length; c++) {
        const addr = XLSX.utils.encode_cell({ r: shRowI, c });
        if (!wsLayout[addr]) wsLayout[addr] = { t: "s", v: "" };
        wsLayout[addr].s = c === secIdx ? S_ORPHAN_SEQ
                         : c === notasIdx ? S_ORPHAN_NOTA
                         : S_ORPHAN_DATA;
      }
    }
  }

  // Anchos de columna
  if (!wsLayout["!cols"]) wsLayout["!cols"] = [];
  if (secIdx >= 0) wsLayout["!cols"][secIdx] = { wch: 16 };
  wsLayout["!cols"][notasIdx] = { wch: 75 };

  XLSX.utils.book_append_sheet(wb, wsLayout, "Layout");

  // Copy original 551 / Data Stage sheet
  const ws551 = XLSX.utils.sheet_to_json(sheet551, { header: 1 });
  const ws551Sheet = XLSX.utils.aoa_to_sheet(ws551);
  XLSX.utils.book_append_sheet(wb, ws551Sheet, sheet551Name || "551");

  // ── Hoja Cruce_Layout_vs_551 ─────────────────────────────────────────────
  if (cruceData && cruceData.length > 0) {
    const pct2 = (v, total) => total > 0 ? ((v / total) * 100).toFixed(1) + "%" : "—";
    const fmt  = (v) => v == null ? "" : (typeof v === "number" ? Number(v.toFixed(4)) : v);
    const fmtDiff = (d) => d == null ? "" : (d > 0 ? `+${d.toFixed(2)}` : d.toFixed(2));
    const check = (ok) => ok ? "✓" : "✗";

    // Cabeceras con dos niveles
    const HEADERS = [
      // Bloque General
      "Estrategia", "N° Filas Layout", "Pedimento",
      // Bloque LAYOUT
      "Layout — Descripcion", "Layout — FraccionNico", "Layout — PaisOrigen",
      "Layout — Suma CantidadSaldo", "Layout — Suma VCUSD",
      "SecuenciaPed Original", "SecuenciaPed ASIGNADA",
      // Bloque 551
      "551 — Secuencias (clave)", "551 — DescripcionMercancia",
      "551 — Fraccion", "551 — PaisOrigenDestino",
      "551 — CantidadUMComercial", "551 — ValorDolares",
      // Comparación
      "¿Fracción coincide?", "¿País coincide?",
      "Dif. Cantidad (Layout−551)", "Dif. Valor USD (Layout−551)",
      "Estado match",
      // Nota (solo sin-match)
      "Notas / Motivo sin asignación",
    ];

    const cruceRows = [HEADERS];

    for (const d of cruceData) {
      const statusMatch = d.tipo === "unmatched"
        ? "❌ SIN MATCH"
        : (d.okCant && d.okVal ? "✅ MATCH EXACTO"
          : (Math.abs(d.diffCant || 0) / Math.max(1, d.layoutCant) < 0.05 ? "⚠ MATCH ~TOL.5%" : "⚠ MATCH TOLERADO"));

      cruceRows.push([
        d.estrategia,
        d.numFilas,
        d.pedimento,
        d.layoutDesc,
        d.layoutFrac,
        d.layoutPais,
        fmt(d.layoutCant),
        fmt(d.layoutVCUSD),
        d.secOriginal ?? "",
        d.secAsignada ?? "",
        d.s551Secuencias ?? "",
        d.s551Desc ?? "",
        d.s551Frac ?? "",
        d.s551Pais ?? "",
        fmt(d.s551Cant),
        fmt(d.s551Val),
        check(d.okFrac),
        check(d.okPais),
        fmtDiff(d.diffCant),
        fmtDiff(d.diffVal),
        statusMatch,
        d.nota || "",
      ]);
    }

    const wsCruce = XLSX.utils.aoa_to_sheet(cruceRows);

    // ── Estilos hoja Cruce ────────────────────────────────────────────────
    const colWidths = [10,12,22,40,16,10,20,20,18,18,38,45,16,10,22,22,16,14,20,20,20,60];
    wsCruce["!cols"] = colWidths.map((w) => ({ wch: w }));

    const S_HEAD = { font:{bold:true,color:{rgb:"FFFFFF"}}, fill:{patternType:"solid",fgColor:{rgb:"1B3A6B"}}, alignment:{horizontal:"center",wrapText:true} };
    const S_MATCH_OK  = { font:{color:{rgb:"155724"}}, fill:{patternType:"solid",fgColor:{rgb:"D4EDDA"}} };
    const S_MATCH_TOL = { font:{color:{rgb:"856404"}}, fill:{patternType:"solid",fgColor:{rgb:"FFF3CD"}} };
    const S_UNMATCH   = { font:{bold:true,color:{rgb:"842029"}}, fill:{patternType:"solid",fgColor:{rgb:"F8D7DA"}} };
    const S_CHECK_OK  = { font:{bold:true,color:{rgb:"155724"}}, alignment:{horizontal:"center"} };
    const S_CHECK_KO  = { font:{bold:true,color:{rgb:"842029"}}, alignment:{horizontal:"center"} };
    const S_DIFF_OK   = { font:{color:{rgb:"155724"}}, alignment:{horizontal:"right"} };
    const S_DIFF_BAD  = { font:{bold:true,color:{rgb:"842029"}}, alignment:{horizontal:"right"} };

    // Cabecera
    for (let c = 0; c < HEADERS.length; c++) {
      const addr = XLSX.utils.encode_cell({ r: 0, c });
      if (!wsCruce[addr]) wsCruce[addr] = { t: "s", v: HEADERS[c] };
      wsCruce[addr].s = S_HEAD;
    }

    // Datos
    for (let rowI = 1; rowI < cruceRows.length; rowI++) {
      const d = cruceData[rowI - 1];
      const isUnmatched = d.tipo === "unmatched";

      // Estilo de fila completa (columna estrategia y status)
      const addrEst = XLSX.utils.encode_cell({ r: rowI, c: 0 });
      if (wsCruce[addrEst]) wsCruce[addrEst].s = isUnmatched ? S_UNMATCH : (d.okCant && d.okVal ? S_MATCH_OK : S_MATCH_TOL);

      const addrStatus = XLSX.utils.encode_cell({ r: rowI, c: 20 });
      if (wsCruce[addrStatus]) wsCruce[addrStatus].s = isUnmatched ? S_UNMATCH : (d.okCant && d.okVal ? S_MATCH_OK : S_MATCH_TOL);

      // Columnas ✓/✗
      const fracOk = XLSX.utils.encode_cell({ r: rowI, c: 16 });
      const paisOk = XLSX.utils.encode_cell({ r: rowI, c: 17 });
      if (wsCruce[fracOk]) wsCruce[fracOk].s = d.okFrac ? S_CHECK_OK : S_CHECK_KO;
      if (wsCruce[paisOk]) wsCruce[paisOk].s = d.okPais ? S_CHECK_OK : S_CHECK_KO;

      // Diferencias
      const addrDC = XLSX.utils.encode_cell({ r: rowI, c: 18 });
      const addrDV = XLSX.utils.encode_cell({ r: rowI, c: 19 });
      if (wsCruce[addrDC]) wsCruce[addrDC].s = d.okCant ? S_DIFF_OK : S_DIFF_BAD;
      if (wsCruce[addrDV]) wsCruce[addrDV].s = d.okVal  ? S_DIFF_OK : S_DIFF_BAD;
    }

    XLSX.utils.book_append_sheet(wb, wsCruce, "Cruce_Layout_vs_551");
  }

  // Resultado_Validacion sheet
  const matched = total - unmatchedFinal.length;
  const pct = ((matched / total) * 100).toFixed(1);

  const reportRows = [
    ["RESULTADO DE VALIDACIÓN — CRUCE LAYOUT vs 551"],
    [],
    ["RESUMEN EJECUTIVO"],
    ["Total filas Layout", total],
    ["Filas con SecuenciaPed asignada", matched],
    ["Filas sin match", unmatchedFinal.length],
    ["PORCENTAJE DE ÉXITO", `${pct}%`],
    [],
    ["DESGLOSE POR ESTRATEGIA"],
    ["Estrategia", "Filas Asignadas", "Descripción"],
    ["E1 - Ped+Fracción+País exacto",    stats.E1, "Agrupa Layout por Pedimento+FraccionNico+País. Suma CantidadSaldo y VCUSD; busca en 551 con tolerancia ±1 ud / ±2 USD."],
    ["E2 - Sub-grupo por SecuenciaPed", stats.E2, "Misma clave E1 pero sub-divide por SecuenciaPed existente. Resuelve cuando la misma fracción+país tiene múltiples entradas en el 551."],
    ["E3 - Sin País (solo Ped+Fracción)", stats.E3, "Ignora PaisOrigen para manejar diferencias de código de país entre Layout y 551. Cantidades y valores exactos."],
    ["E4 - Sin País + Sub-SecuenciaPed",  stats.E4, "Combina E3 y E2: sin País + sub-agrupación por SecuenciaPed. Captura casos con variación de código país Y múltiples secuencias."],
    ["E5 - Tolerancia Ampliada ±5%",     stats.E5, "Último recurso con tolerancia ±5% en cantidad y valor. Resuelve diferencias de redondeo o conversión de unidades entre sistemas."],
    [],
  ];

  if (unmatchedFinal.length > 0) {
    reportRows.push(["GRUPOS SIN MATCH — REVISIÓN MANUAL REQUERIDA"]);
    reportRows.push(["Descripcion", "FraccionNico", "PaisOrigen", "SecuenciaPed_Original", "CantidadSaldo", "VCUSD", "Notas (motivo sin asignación)"]);
    for (const u of unmatchedFinal) {
      reportRows.push([u.Descripcion, u.FraccionNico, u.PaisOrigen, u.SecuenciaPed_Original, u.CantidadSaldo, u.VCUSD, u.Nota || ""]);
    }
  } else {
    reportRows.push(["✓ TODOS LOS GRUPOS TUVIERON MATCH EXITOSO"]);
  }

  if (orphan551Rows && orphan551Rows.length > 0) {
    reportRows.push([]);
    reportRows.push([`SECUENCIAS DEL 551 NO ASIGNADAS AL LAYOUT  (${orphan551Rows.length} registros)`]);
    reportRows.push(["SecuenciaFraccion", "Pedimento", "Fraccion", "PaisOrigenDestino", "CantidadUMComercial", "ValorDolares", "Motivo / Razón"]);
    for (const r of orphan551Rows) {
      reportRows.push([
        r["SecuenciaFraccion"] ?? "",
        r["Pedimento"] ?? "",
        r["Fraccion"]  ?? "",
        r["PaisOrigenDestino"] ?? "",
        parseFloat(r["CantidadUMComercial"]) || 0,
        parseFloat(r["ValorDolares"])        || 0,
        r._orphanReason || "",
      ]);
    }
  }

  const wsReport = XLSX.utils.aoa_to_sheet(reportRows);
  XLSX.utils.book_append_sheet(wb, wsReport, "Resultado_Validacion");

  return wb;
}

// ─── COMPONENTS ───────────────────────────────────────────────────────────────
const STRATEGIES = [
  {
    id: "E1",
    name: "Pedimento + Fracción + País",
    desc: "Agrupación exacta por Pedimento + FraccionNico + PaisOrigen. Suma CantidadSaldo vs CantidadUMComercial y VCUSD vs ValorDolares del 551 (tolerancia ±1 unidad / ±2 USD). Resuelve la mayoría de los casos.",
    color: "#22c55e",
    icon: "⬛",
  },
  {
    id: "E2",
    name: "Sub-agrupación por SecuenciaPed",
    desc: "Para grupos que fallaron E1, sub-divide usando el SecuenciaPed existente como guía. Resuelve casos donde la misma fracción+país tiene múltiples líneas en el 551 (ej: mismo material importado en dos fechas distintas).",
    color: "#3b82f6",
    icon: "⬛",
  },
  {
    id: "E3",
    name: "Sin filtro de País (Ped + Fracción)",
    desc: "Ignora PaisOrigen para manejar diferencias de captura de código de país entre Layout y 551 (ej: 'TWN' vs 'TAI', 'CHN' vs 'PRC'). Aplica las mismas tolerancias exactas de cantidad y valor.",
    color: "#f59e0b",
    icon: "⬛",
  },
  {
    id: "E4",
    name: "Sin País + Sub-SecuenciaPed",
    desc: "Combina E3 y E2: sin filtro de País y sub-agrupación por SecuenciaPed. Captura casos donde hay variación de código de país Y múltiples secuencias para la misma fracción.",
    color: "#a855f7",
    icon: "⬛",
  },
  {
    id: "E5",
    name: "Tolerancia Ampliada (±5%)",
    desc: "Último recurso: tolerancia ±5% en cantidad (mín 2 unidades) y ±5% en valor (mín 5 USD). Resuelve diferencias de redondeo, conversión de unidades UMC/UMT o tipos de cambio entre sistemas.",
    color: "#ef4444",
    icon: "⬛",
  },
];

function UploadZone({ onFile, isDragging, setIsDragging }) {
  const ref = useRef(null);
  const handleDrop = useCallback(
    (e) => {
      e.preventDefault();
      setIsDragging(false);
      const file = e.dataTransfer.files[0];
      if (file) onFile(file);
    },
    [onFile, setIsDragging]
  );

  return (
    <div
      onDragOver={(e) => { e.preventDefault(); setIsDragging(true); }}
      onDragLeave={() => setIsDragging(false)}
      onDrop={handleDrop}
      onClick={() => ref.current?.click()}
      style={{
        border: `2px dashed ${isDragging ? "#f59e0b" : "#334155"}`,
        borderRadius: 4,
        padding: "60px 40px",
        textAlign: "center",
        cursor: "pointer",
        background: isDragging ? "rgba(245,158,11,0.05)" : "rgba(15,23,42,0.6)",
        transition: "all 0.2s",
        position: "relative",
      }}
    >
      <input
        ref={ref}
        type="file"
        accept=".xlsx,.xls"
        style={{ display: "none" }}
        onChange={(e) => e.target.files[0] && onFile(e.target.files[0])}
      />
      <div style={{ fontSize: 48, marginBottom: 16 }}>📊</div>
      <div style={{ color: "#f8fafc", fontSize: 18, fontWeight: 700, marginBottom: 8, fontFamily: "Syne, sans-serif" }}>
        Sube tu archivo Excel
      </div>
      <div style={{ color: "#94a3b8", fontSize: 13 }}>
        Arrastra aquí o haz clic · Requiere hojas <span style={{ color: "#f59e0b", fontFamily: "monospace" }}>Layout</span> y <span style={{ color: "#f59e0b", fontFamily: "monospace" }}>551</span>
      </div>
    </div>
  );
}

function StrategyBar({ stats, total }) {
  const colors = { E1: "#22c55e", E2: "#3b82f6", E3: "#f59e0b", E4: "#a855f7", E5: "#ef4444" };
  const unmatched = total - Object.values(stats).reduce((a, b) => a + b, 0);
  const segs = [
    ...Object.entries(stats).map(([k, v]) => ({ label: k, val: v, color: colors[k] })),
    { label: "Sin match", val: unmatched, color: "#1e293b" },
  ];

  return (
    <div>
      <div style={{ display: "flex", height: 20, borderRadius: 2, overflow: "hidden", gap: 1, marginBottom: 10 }}>
        {segs.map((s) =>
          s.val > 0 ? (
            <div
              key={s.label}
              title={`${s.label}: ${s.val} filas`}
              style={{
                flex: s.val,
                background: s.color,
                opacity: s.label === "Sin match" ? 0.4 : 1,
              }}
            />
          ) : null
        )}
      </div>
      <div style={{ display: "flex", flexWrap: "wrap", gap: "8px 20px" }}>
        {segs.map((s) => (
          <div key={s.label} style={{ display: "flex", alignItems: "center", gap: 6 }}>
            <div style={{ width: 10, height: 10, background: s.color, borderRadius: 1, opacity: s.label === "Sin match" ? 0.4 : 1 }} />
            <span style={{ color: "#94a3b8", fontSize: 12, fontFamily: "monospace" }}>
              {s.label} <span style={{ color: "#f8fafc" }}>{s.val}</span>
            </span>
          </div>
        ))}
      </div>
    </div>
  );
}

function StatCard({ label, value, sub, accent }) {
  return (
    <div style={{
      background: "#0f172a",
      border: "1px solid #1e293b",
      borderLeft: `3px solid ${accent}`,
      borderRadius: 4,
      padding: "20px 24px",
    }}>
      <div style={{ color: "#94a3b8", fontSize: 11, textTransform: "uppercase", letterSpacing: "0.1em", marginBottom: 8, fontFamily: "Syne, sans-serif" }}>
        {label}
      </div>
      <div style={{ color: accent, fontSize: 36, fontWeight: 900, fontFamily: "DM Mono, monospace", lineHeight: 1 }}>
        {value}
      </div>
      {sub && <div style={{ color: "#475569", fontSize: 12, marginTop: 6, fontFamily: "monospace" }}>{sub}</div>}
    </div>
  );
}

// ─── MAIN APP ────────────────────────────────────────────────────────────────
export default function App() {
  const [phase, setPhase] = useState("upload"); // upload | processing | results
  const [isDragging, setIsDragging] = useState(false);
  const [results, setResults] = useState(null);
  const [outputWb, setOutputWb] = useState(null);
  const [fileName, setFileName] = useState("");
  const [error, setError] = useState(null);
  const [showUnmatched, setShowUnmatched] = useState(false);
  const [activeStrategy, setActiveStrategy] = useState(null);
  const [progress, setProgress] = useState(0);

  const processFile = useCallback(async (file) => {
    setError(null);
    setFileName(file.name);
    setPhase("processing");
    setProgress(0);

    try {
      const buf = await file.arrayBuffer();
      setProgress(20);
      const wb = XLSX.read(buf, { type: "array" });

      if (!wb.SheetNames.includes("Layout")) {
        throw new Error('El archivo debe contener la hoja "Layout"');
      }
      const sheet551Name = resolve551SheetName(wb);
      if (!sheet551Name) {
        throw new Error('El archivo debe contener una hoja "551" o "Data Stage" (datos del pedimento)');
      }

      setProgress(40);
      const layoutRows = readLayoutSheet(wb.Sheets["Layout"]);
      const s551Rows   = read551Sheet(wb.Sheets[sheet551Name]);
      setProgress(60);

      const { assignment, strategyStats, unmatchedFinal, total, rowNotes, cruceData, orphan551Rows } = runCascade(layoutRows, s551Rows);
      setProgress(80);

      const newWb = buildOutputExcel(wb, wb.Sheets["Layout"], wb.Sheets[sheet551Name], sheet551Name, assignment, unmatchedFinal, strategyStats, total, rowNotes, cruceData, orphan551Rows);
      setProgress(100);

      setResults({ strategyStats, unmatchedFinal, total, orphan551Count: orphan551Rows.length });
      setOutputWb(newWb);

      setTimeout(() => setPhase("results"), 400);
    } catch (e) {
      setError(e.message);
      setPhase("upload");
    }
  }, []);

  const downloadResult = () => {
    if (!outputWb) return;
    XLSX.writeFile(outputWb, fileName.replace(/\.xlsx?$/i, "") + "_Resultado.xlsx");
  };

  const reset = () => {
    setPhase("upload");
    setResults(null);
    setOutputWb(null);
    setError(null);
    setProgress(0);
  };

  const matched = results ? results.total - results.unmatchedFinal.length : 0;
  const pct = results ? ((matched / results.total) * 100).toFixed(1) : 0;

  return (
    <div style={{
      minHeight: "100vh",
      background: "#0f172a",
      fontFamily: "Syne, system-ui, sans-serif",
      color: "#f8fafc",
    }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800;900&family=DM+Mono:wght@400;500&display=swap');
        * { box-sizing: border-box; }
        ::-webkit-scrollbar { width: 6px; }
        ::-webkit-scrollbar-track { background: #0f172a; }
        ::-webkit-scrollbar-thumb { background: #334155; border-radius: 3px; }
        .strat-card:hover { border-color: #334155 !important; background: rgba(30,41,59,0.8) !important; }
        .dl-btn:hover { background: #d97706 !important; }
        .row-hover:hover { background: rgba(245,158,11,0.05) !important; }
        @keyframes spin { to { transform: rotate(360deg); } }
        @keyframes fadeUp { from { opacity:0; transform:translateY(12px); } to { opacity:1; transform:none; } }
        @keyframes progressFill { from { width: 0%; } to { width: 100%; } }
      `}</style>

      {/* Header */}
      <div style={{
        borderBottom: "1px solid #1e293b",
        padding: "18px 40px",
        display: "flex",
        alignItems: "center",
        justifyContent: "space-between",
        background: "rgba(15,23,42,0.95)",
        position: "sticky",
        top: 0,
        zIndex: 100,
      }}>
        <div style={{ display: "flex", alignItems: "center", gap: 16 }}>
          <div style={{
            width: 36, height: 36, background: "#f59e0b",
            display: "flex", alignItems: "center", justifyContent: "center",
            fontSize: 18, borderRadius: 4, flexShrink: 0,
          }}>🛃</div>
          <div>
            <div style={{ fontSize: 15, fontWeight: 800, letterSpacing: "-0.01em", color: "#f8fafc" }}>
              SECUENCIAPED MATCHER
            </div>
            <div style={{ fontSize: 11, color: "#475569", letterSpacing: "0.08em", fontFamily: "DM Mono, monospace" }}>
              COMERCIO EXTERIOR · INMEX · PEDIMENTO 551
            </div>
          </div>
        </div>
        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          {phase === "results" && (
            <>
              <button
                onClick={reset}
                style={{
                  background: "transparent", border: "1px solid #334155",
                  color: "#94a3b8", padding: "8px 16px", cursor: "pointer",
                  borderRadius: 4, fontSize: 13, fontFamily: "Syne, sans-serif",
                }}
              >
                ← Nuevo archivo
              </button>
              <button
                className="dl-btn"
                onClick={downloadResult}
                style={{
                  background: "#f59e0b", border: "none", color: "#0f172a",
                  padding: "8px 20px", cursor: "pointer", borderRadius: 4,
                  fontSize: 13, fontWeight: 800, fontFamily: "Syne, sans-serif",
                  transition: "background 0.2s",
                }}
              >
                ⬇ Descargar Excel
              </button>
            </>
          )}
        </div>
      </div>

      <div style={{ maxWidth: 1100, margin: "0 auto", padding: "40px 24px" }}>

        {/* UPLOAD PHASE */}
        {phase === "upload" && (
          <div style={{ animation: "fadeUp 0.4s ease" }}>
            <div style={{ textAlign: "center", marginBottom: 48 }}>
              <div style={{
                display: "inline-block",
                background: "rgba(245,158,11,0.1)",
                border: "1px solid rgba(245,158,11,0.2)",
                color: "#f59e0b",
                padding: "4px 14px",
                borderRadius: 20,
                fontSize: 11,
                letterSpacing: "0.12em",
                fontFamily: "DM Mono, monospace",
                marginBottom: 20,
              }}>
                5 ESTRATEGIAS EN CASCADA · MATCHING INTELIGENTE
              </div>
              <h1 style={{
                fontSize: 42, fontWeight: 900, margin: "0 0 16px",
                letterSpacing: "-0.02em", lineHeight: 1.1,
              }}>
                Asignación automática de{" "}
                <span style={{ color: "#f59e0b" }}>SecuenciaPed</span>
              </h1>
              <p style={{ color: "#64748b", fontSize: 16, maxWidth: 540, margin: "0 auto" }}>
                Cruza datos entre Layout y 551 aplicando metodología de consultor
                experto en pedimentos IMMEX.
              </p>
            </div>

            {error && (
              <div style={{
                background: "rgba(239,68,68,0.1)", border: "1px solid rgba(239,68,68,0.3)",
                borderRadius: 4, padding: "12px 16px", marginBottom: 20, color: "#fca5a5", fontSize: 13,
              }}>
                ⚠ {error}
              </div>
            )}

            <UploadZone onFile={processFile} isDragging={isDragging} setIsDragging={setIsDragging} />

            {/* Strategy cards */}
            <div style={{ marginTop: 48 }}>
              <div style={{ color: "#475569", fontSize: 11, letterSpacing: "0.1em", marginBottom: 20, fontFamily: "DM Mono, monospace" }}>
                METODOLOGÍA DE COINCIDENCIA — CASCADA DE 5 ESTRATEGIAS
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(300px, 1fr))", gap: 12 }}>
                {STRATEGIES.map((s) => (
                  <div
                    key={s.id}
                    className="strat-card"
                    style={{
                      background: "rgba(15,23,42,0.8)",
                      border: "1px solid #1e293b",
                      borderRadius: 4,
                      padding: "18px 20px",
                      transition: "all 0.2s",
                      cursor: "default",
                    }}
                  >
                    <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 10 }}>
                      <span style={{
                        background: s.color, color: "#000", fontSize: 10, fontWeight: 900,
                        padding: "2px 8px", borderRadius: 2, fontFamily: "DM Mono, monospace",
                        flexShrink: 0,
                      }}>
                        {s.id}
                      </span>
                      <span style={{ fontSize: 13, fontWeight: 700, color: "#f8fafc" }}>{s.name}</span>
                    </div>
                    <p style={{ color: "#64748b", fontSize: 12, lineHeight: 1.6, margin: 0 }}>{s.desc}</p>
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}

        {/* PROCESSING PHASE */}
        {phase === "processing" && (
          <div style={{ textAlign: "center", padding: "80px 0", animation: "fadeUp 0.3s ease" }}>
            <div style={{
              width: 60, height: 60, border: "3px solid #1e293b",
              borderTop: "3px solid #f59e0b", borderRadius: "50%",
              margin: "0 auto 32px",
              animation: "spin 0.8s linear infinite",
            }} />
            <div style={{ fontSize: 20, fontWeight: 800, marginBottom: 8 }}>Procesando archivo…</div>
            <div style={{ color: "#475569", fontSize: 13, marginBottom: 32 }}>{fileName}</div>
            <div style={{ maxWidth: 400, margin: "0 auto" }}>
              <div style={{ background: "#1e293b", borderRadius: 2, height: 4, overflow: "hidden" }}>
                <div style={{
                  height: "100%", background: "#f59e0b", borderRadius: 2,
                  width: `${progress}%`, transition: "width 0.4s ease",
                }} />
              </div>
              <div style={{ color: "#475569", fontSize: 12, marginTop: 8, fontFamily: "DM Mono, monospace" }}>
                Ejecutando cascada de estrategias · {progress}%
              </div>
            </div>
          </div>
        )}

        {/* RESULTS PHASE */}
        {phase === "results" && results && (
          <div style={{ animation: "fadeUp 0.5s ease" }}>
            {/* Headline stats */}
            <div style={{ marginBottom: 32 }}>
              <div style={{ color: "#475569", fontSize: 11, letterSpacing: "0.1em", fontFamily: "DM Mono, monospace", marginBottom: 12 }}>
                {fileName} · {results.total} filas procesadas
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(190px, 1fr))", gap: 12 }}>
                <StatCard label="Éxito global" value={`${pct}%`} sub={`${matched} de ${results.total} filas`} accent="#f59e0b" />
                <StatCard label="Filas asignadas" value={matched} sub="SecuenciaPed actualizada" accent="#22c55e" />
                <StatCard label="Sin match (Layout)" value={results.unmatchedFinal.length} sub="Requieren revisión manual" accent={results.unmatchedFinal.length > 0 ? "#ef4444" : "#22c55e"} />
                <StatCard label="Sec. 551 sin asignar" value={results.orphan551Count || 0} sub="Al final del Layout en Excel" accent={(results.orphan551Count || 0) > 0 ? "#3b82f6" : "#22c55e"} />
                <StatCard label="Estrategias activas" value={Object.values(results.strategyStats).filter((v) => v > 0).length} sub="de 5 disponibles" accent="#a855f7" />
              </div>
            </div>

            {/* Progress bar */}
            <div style={{
              background: "#0f172a", border: "1px solid #1e293b",
              borderRadius: 4, padding: "24px 28px", marginBottom: 20,
            }}>
              <div style={{ color: "#94a3b8", fontSize: 11, letterSpacing: "0.1em", fontFamily: "DM Mono, monospace", marginBottom: 16 }}>
                DISTRIBUCIÓN DE ASIGNACIÓN POR ESTRATEGIA
              </div>
              <StrategyBar stats={results.strategyStats} total={results.total} />
            </div>

            {/* Strategy breakdown */}
            <div style={{ marginBottom: 20 }}>
              <div style={{ color: "#475569", fontSize: 11, letterSpacing: "0.1em", fontFamily: "DM Mono, monospace", marginBottom: 12 }}>
                DETALLE POR ESTRATEGIA
              </div>
              <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                {STRATEGIES.map((s) => {
                  const count = results.strategyStats[s.id] || 0;
                  const pctS = results.total ? ((count / results.total) * 100).toFixed(1) : 0;
                  return (
                    <div
                      key={s.id}
                      onClick={() => setActiveStrategy(activeStrategy === s.id ? null : s.id)}
                      style={{
                        background: "#0f172a",
                        border: `1px solid ${activeStrategy === s.id ? s.color + "66" : "#1e293b"}`,
                        borderRadius: 4, padding: "14px 20px", cursor: "pointer",
                        transition: "all 0.2s",
                      }}
                    >
                      <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
                        <span style={{
                          background: count > 0 ? s.color : "#1e293b",
                          color: count > 0 ? "#000" : "#475569",
                          fontSize: 10, fontWeight: 900,
                          padding: "2px 8px", borderRadius: 2, fontFamily: "DM Mono, monospace",
                          flexShrink: 0,
                        }}>
                          {s.id}
                        </span>
                        <span style={{ fontSize: 13, fontWeight: 600, flex: 1, color: count > 0 ? "#f8fafc" : "#475569" }}>
                          {s.name}
                        </span>
                        <span style={{ fontFamily: "DM Mono, monospace", fontSize: 13, color: s.color, fontWeight: 700, marginRight: 8 }}>
                          {count > 0 ? `+${count}` : "0"}
                        </span>
                        <span style={{ fontFamily: "DM Mono, monospace", fontSize: 11, color: "#475569", width: 48, textAlign: "right" }}>
                          {pctS}%
                        </span>
                        <div style={{ width: 160, background: "#1e293b", borderRadius: 1, height: 4 }}>
                          <div style={{ width: `${pctS}%`, height: "100%", background: s.color, borderRadius: 1 }} />
                        </div>
                        <span style={{ color: "#334155", fontSize: 12 }}>{activeStrategy === s.id ? "▲" : "▼"}</span>
                      </div>
                      {activeStrategy === s.id && (
                        <div style={{ marginTop: 12, paddingTop: 12, borderTop: "1px solid #1e293b", color: "#64748b", fontSize: 12, lineHeight: 1.7 }}>
                          <strong style={{ color: "#94a3b8" }}>Metodología:</strong> {s.desc}
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            </div>

            {/* Unmatched table */}
            {results.unmatchedFinal.length > 0 && (
              <div style={{
                background: "#0f172a",
                border: "1px solid rgba(239,68,68,0.3)",
                borderRadius: 4, overflow: "hidden", marginBottom: 20,
              }}>
                <div
                  onClick={() => setShowUnmatched(!showUnmatched)}
                  style={{
                    padding: "16px 20px", cursor: "pointer",
                    display: "flex", alignItems: "center", justifyContent: "space-between",
                    background: "rgba(239,68,68,0.05)",
                  }}
                >
                  <div>
                    <span style={{ color: "#ef4444", fontWeight: 800, fontSize: 13 }}>
                      ⚠ {results.unmatchedFinal.length} filas sin match
                    </span>
                    <span style={{ color: "#475569", fontSize: 12, marginLeft: 12 }}>
                      Requieren revisión manual por un especialista
                    </span>
                  </div>
                  <span style={{ color: "#475569" }}>{showUnmatched ? "▲" : "▼"}</span>
                </div>

                {showUnmatched && (
                  <div style={{ overflowX: "auto" }}>
                    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                      <thead>
                        <tr style={{ background: "#1e293b" }}>
                          {["Fraccion", "País", "Cantidad", "VCUSD", "Notas — Motivo sin asignación"].map((h) => (
                            <th key={h} style={{
                              padding: "10px 16px", textAlign: "left",
                              color: "#64748b", fontFamily: "DM Mono, monospace",
                              fontSize: 10, letterSpacing: "0.06em", fontWeight: 500,
                            }}>
                              {h}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {results.unmatchedFinal.slice(0, 100).map((r, i) => (
                          <tr key={i} className="row-hover" style={{ borderTop: "1px solid #1e293b", transition: "background 0.1s" }}>
                            <td style={{ padding: "9px 16px", color: "#f59e0b", fontFamily: "monospace", whiteSpace: "nowrap" }}>{r.FraccionNico}</td>
                            <td style={{ padding: "9px 16px", color: "#94a3b8", fontFamily: "monospace" }}>{r.PaisOrigen}</td>
                            <td style={{ padding: "9px 16px", color: "#94a3b8", fontFamily: "monospace" }}>{r.CantidadSaldo?.toLocaleString()}</td>
                            <td style={{ padding: "9px 16px", color: "#94a3b8", fontFamily: "monospace" }}>{Number(r.VCUSD).toFixed(2)}</td>
                            <td style={{ padding: "9px 16px", color: "#fca5a5", fontSize: 11, lineHeight: 1.5, maxWidth: 400 }}>{r.Nota || "—"}</td>
                          </tr>
                        ))}
                        {results.unmatchedFinal.length > 100 && (
                          <tr>
                            <td colSpan={5} style={{ padding: "12px 16px", color: "#475569", textAlign: "center", fontFamily: "monospace", fontSize: 11 }}>
                              ... y {results.unmatchedFinal.length - 100} filas más (ver Excel descargado)
                            </td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                )}
              </div>
            )}

            {/* Download CTA */}
            <div style={{
              background: "linear-gradient(135deg, rgba(245,158,11,0.1), rgba(245,158,11,0.03))",
              border: "1px solid rgba(245,158,11,0.25)",
              borderRadius: 4, padding: "28px 32px",
              display: "flex", alignItems: "center", justifyContent: "space-between",
              flexWrap: "wrap", gap: 16,
            }}>
              <div>
                <div style={{ fontSize: 16, fontWeight: 800, marginBottom: 4 }}>
                  Archivo listo para descargar
                </div>
                <div style={{ color: "#64748b", fontSize: 13 }}>
                  Incluye Layout actualizado + hoja <code style={{ color: "#f59e0b", background: "rgba(245,158,11,0.1)", padding: "1px 6px", borderRadius: 2 }}>Resultado_Validacion</code>
                </div>
              </div>
              <button
                className="dl-btn"
                onClick={downloadResult}
                style={{
                  background: "#f59e0b", border: "none", color: "#0f172a",
                  padding: "12px 28px", cursor: "pointer", borderRadius: 4,
                  fontSize: 14, fontWeight: 900, fontFamily: "Syne, sans-serif",
                  transition: "background 0.2s",
                }}
              >
                ⬇ Descargar Excel Resultado
              </button>
            </div>

            {/* Recommendations for unmatched */}
            {results.unmatchedFinal.length > 0 && (
              <div style={{ marginTop: 24 }}>
                <div style={{ color: "#475569", fontSize: 11, letterSpacing: "0.1em", fontFamily: "DM Mono, monospace", marginBottom: 12 }}>
                  RECOMENDACIONES PARA FILAS SIN MATCH
                </div>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                  {[
                    { icon: "🔍", title: "Verificar Fracción Arancelaria", body: "Muchos casos sin match ocurren porque el mismo producto tiene múltiples fracciones (ej: 85322999 vs 85414004 para CAPACITORES). Agregar FraccionImpo como criterio de agrupación resolvería estos casos." },
                    { icon: "📋", title: "Revisar Pedimentos Pendientes", body: "Si la suma del Layout supera la cantidad del 551, es posible que parte del inventario provenga de pedimentos anteriores no incluidos en el archivo. Solicitar expediente completo." },
                    { icon: "⚖️", title: "Validar Unidades de Medida", body: "Diferencias de cantidad pueden deberse a conversiones UMC/UMT. Verificar si el 551 reporta en unidades distintas al Layout (piezas vs. lotes, kg vs. pzas)." },
                    { icon: "🔄", title: "Conciliación Parcial", body: "Para ARNES ELÉCTRICO y productos similares con múltiples registros en 551, hacer conciliación ítem por ítem comparando valor unitario (ValorDolares / CantidadUMComercial) como criterio discriminador." },
                  ].map((r) => (
                    <div key={r.title} style={{
                      background: "#0f172a", border: "1px solid #1e293b",
                      borderRadius: 4, padding: "18px 20px",
                    }}>
                      <div style={{ fontSize: 18, marginBottom: 8 }}>{r.icon}</div>
                      <div style={{ fontSize: 13, fontWeight: 700, marginBottom: 6, color: "#f8fafc" }}>{r.title}</div>
                      <div style={{ fontSize: 12, color: "#64748b", lineHeight: 1.6 }}>{r.body}</div>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}
