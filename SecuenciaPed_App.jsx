import { useState, useCallback, useRef } from "react";
import * as XLSX from "xlsx-js-style";

//  LECTURA EXCEL (hoja 551 con columnas duplicadas / nombres con espacios) 
// Columnas necesarias del 551 (se busca la PRIMERA ocurrencia de cada nombre, por eso
// se usa header:1 y búsqueda por nombre trimado en lugar de sheet_to_json directamente).
const COLS_551 = [
  "Pedimento",
  "Pedimento2",         // forma extendida; prioridad en pedCruceKey para alinear con Layout
  "Secuencias",         // clave compuesta Ped-Fraccion-SecuenciaFraccion (match directo con CANDADO DS 551 del Layout)
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

/**
 * Alinea pedimentos entre Layout (ej. 21-400-3459-1000958) y 551/DS
 * (ej. 400-3459-1000958 o 1000958) para claves de cruce.
 */
/**
 * Clave canónica de pedimento para cruzar Layout vs 551/DS.
 * En México el pedimento "real" (aduana-aa-patente-clave) es fijo: p. ej. 400-3459-1000958.
 * Los 12 dígitos iniciales (año, aduana distinta, etc. como 22- o 21- antes del núcleo)
 * NO cambian el pedimento y se ignoran para el cruce.
 * Prioridad: extraer el trío NNN-NNNN- por regex; si no, quitar solo el primer tramo de 12 dígitos.
 */
function pedCruceKey(raw) {
  const s = String(raw ?? "").trim();
  if (!s) return "";
  // Núcleo estándar: 3 cifras - 4 cifras - clave (410 dígitos)  p. ej. 400-3459-1000958
  const mCore = s.match(/(\d{3}-\d{4}-\d{4,10})/);
  if (mCore) return mCore[1];
  if (/^\d{6,10}$/.test(s)) return s;
  const p = s.split(/[\s-]+/).filter(Boolean);
  // 22-400-3459-1000958    descartar el primer tramo (12 dígitos) y unir 400-3459-1000958
  if (p.length >= 4 && p[0].length <= 2 && /^\d+$/.test(p[0])) {
    return p.slice(1, 4).join("-");
  }
  if (p.length === 3) return p.join("-");
  if (p.length === 2) return s;
  if (p.length === 1) return p[0];
  return s;
}

function scoreLayoutHeaderRow(cells) {
  const HITS = ["pediment", "fraccion", "francion", "frac", "pais", "candado", "cantidad", "descrip", "umc", "sec", "vcusd", "valor", "clave", "aduana"];
  let s = 0;
  for (const c of cells) {
    const t = String(c ?? "").trim().toLowerCase();
    if (!t) continue;
    for (const h of HITS) {
      if (t.includes(h)) { s += 1; break; }
    }
  }
  return s;
}

function findLayoutHeaderRowIndex(rows) {
  const maxR = Math.min(25, rows.length);
  let best = 0, bestS = -1;
  for (let r = 0; r < maxR; r++) {
    const sc = scoreLayoutHeaderRow(rows[r] || []);
    if (sc > bestS) { bestS = sc; best = r; }
  }
  return bestS >= 2 ? best : 0;
}

/**
 * Asegura nombres de columnas que usa runCascade (p. ej. pedimento  Pedimento,
 * cantidad_comercial  CantidadSaldo).
 */
function applyLayoutAliases(obj) {
  if (!obj || typeof obj !== "object") return obj;
  if ((obj.Pedimento == null || String(obj.Pedimento).trim() === "") && obj.pedimento != null)
    obj.Pedimento = obj.pedimento;
  if (obj.pedimento1 && (obj.Pedimento == null || String(obj.Pedimento).trim() === "")) obj.Pedimento = obj.pedimento1;
  if (obj.FraccionNico == null) {
    for (const a of [obj.fraccionNico, obj.Fraccion, obj.fraccion, obj.frac]) {
      if (a != null && String(a).trim() !== "") { obj.FraccionNico = a; break; }
    }
  }
  if (obj.PaisOrigen == null) {
    for (const a of [obj.pais_origen, obj.pais, obj.origen_destino]) {
      if (a != null && String(a).trim() !== "") { obj.PaisOrigen = a; break; }
    }
  }
  if (obj.CantidadSaldo == null) {
    for (const a of [obj.cantidad_comercial, obj.cantidad_saldo, obj.Cantidad]) {
      if (a != null && String(a).replace(/\s/g, "") !== "") { obj.CantidadSaldo = a; break; }
    }
  }
  if (obj.VCUSD == null) {
    for (const a of [obj.ValorTotalDolares, obj.vcusd, obj.ValorMPDolares, obj.valor_total_dlls]) {
      if (a != null && String(a).replace(/\s/g, "") !== "") { obj.VCUSD = a; break; }
    }
  }
  if (obj.Descripcion == null) {
    for (const a of [obj.descripcion, obj["Descripción"]]) {
      if (a != null && String(a).trim() !== "") { obj.Descripcion = a; break; }
    }
  }
  if (obj.SecuenciaPed == null) {
    for (const a of [obj["SEC CALC"], obj.secuencias, obj.SeqPed, obj.Seq, obj.secuencia_ped, obj.sSecuenciaPed, obj.Secuencia]) {
      if (a != null && String(a).replace(/\s/g, "") !== "") { obj.SecuenciaPed = a; break; }
    }
  }
  return obj;
}

/** Lee la hoja 551 tomando la PRIMERA columna que coincida con cada nombre (maneja espacios y duplicados).
 *  Caso especial: "ValorDolares" puede aparecer dos veces en el encabezado del 551.
 *  La primera columna suele ser el valor del lote/parcial; la segunda puede ser el valor en
 *  aduana o el acumulado. Usamos la primera que tenga un valor numérico no-cero;
 *  si ambas son 0 o vacías, usamos 0. Así evitamos que un ValorDolares=0 erróneo rompa el matching.
 */
function read551Sheet(sheet) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  if (!rows.length) return [];
  const headerRow = rows[0].map((c) => String(c ?? "").trim());
  const indices = {};
  for (const col of COLS_551) {
    const idx = firstIndexByHeader(headerRow, col);
    if (idx >= 0) indices[col] = idx;
  }

  // Recolectar TODOS los índices de columnas llamadas "ValorDolares"
  const vdIndices = headerRow.reduce((acc, h, i) => { if (h === "ValorDolares") acc.push(i); return acc; }, []);

  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (row.every((c) => c === "" || c == null)) continue;
    const obj = {};
    for (const [col, idx] of Object.entries(indices)) {
      obj[col] = row[idx];
    }
    // Resolver ValorDolares efectivo cuando hay columnas duplicadas:
    // recorrer todos los índices y quedarse con el primero que tenga valor numérico  0.
    if (vdIndices.length > 1) {
      let efectivo = null;
      for (const idx of vdIndices) {
        const n = parseFloat(row[idx]);
        if (!isNaN(n) && n !== 0) { efectivo = row[idx]; break; }
      }
      // Si todas son 0/vacío, usar la primera (mantiene semántica de 0)
      obj["ValorDolares"] = efectivo !== null ? efectivo : (row[vdIndices[0]] ?? 0);
    }
    out.push(obj);
  }
  return out;
}

/** Lee el Layout: detecta fila de encabezado (plantillas con fila 1 de títulos) y aplica alias. */
function readLayoutSheet(sheet) {
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  if (!rows.length) return [];
  const headerRowI = findLayoutHeaderRowIndex(rows);
  const rawHeaders = rows[headerRowI] || [];
  const headers = rawHeaders.map((c, j) => {
    const t = String(c ?? "").trim();
    return t || `__C${j}`;
  });
  const out = [];
  for (let i = headerRowI + 1; i < rows.length; i++) {
    const row = rows[i];
    if (row.every((c) => c === "" || c == null)) continue;
    const obj = {};
    for (let j = 0; j < headers.length; j++) {
      const key = headers[j] || `__C${j}`;
      obj[key] = row[j];
    }
    out.push(applyLayoutAliases(obj));
  }
  return out;
}

function resolve551SheetName(wb) {
  const names = wb.SheetNames || [];
  if (names.includes("551")) return "551";
  const lower = names.map((n) => String(n).toLowerCase());
  const d = lower.findIndex((n) => n === "ds" || n === "data" || n.startsWith("ds"));
  if (d >= 0) return names[d];
  const i = lower.findIndex((n) => n.includes("data") && n.includes("stage"));
  if (i >= 0) return names[i];
  const j = lower.findIndex((n) => n === "datastage" || n.includes("551"));
  if (j >= 0) return names[j];
  return null;
}

/** Nombre de hoja Layout (case-insensitive) o "Layout" por defecto. */
function resolveLayoutSheetName(wb) {
  const names = wb.SheetNames || [];
  if (names.includes("Layout")) return "Layout";
  const lower = names.map((n) => String(n).toLowerCase());
  const i = lower.findIndex((n) => n === "layout");
  if (i >= 0) return names[i];
  return null;
}

//  MATCHING ENGINE 
// Lógica real de cruce IMMEX:
//   Layout: Pedimento + FraccionNico + PaisOrigen  suma CantidadSaldo y VCUSD
//   551:    Pedimento + Fraccion    + PaisOrigenDestino  CantidadUMComercial y ValorDolares
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

/**
 * E6  Busca un subconjunto de 2 o 3 candidatos cuya suma  (sumCant, sumVCUSD).
 * Detecta el caso donde el mismo material entró en varios lotes del mismo pedimento.
 * Limita el pool a 12 para evitar explosión combinatoria (12C3 = 220 iteraciones máx.).
 */
function tryMatchCombination(candidates, sumCant, sumVCUSD, tolC = 1, tolV = 4) {
  const pool = candidates
    .filter((r) => (parseFloat(r["CantidadUMComercial"]) || 0) > 0)
    .slice(0, 12);
  const n = pool.length;

  // Pares
  for (let i = 0; i < n - 1; i++) {
    for (let j = i + 1; j < n; j++) {
      const c = (parseFloat(pool[i]["CantidadUMComercial"]) || 0)
              + (parseFloat(pool[j]["CantidadUMComercial"]) || 0);
      const v = (parseFloat(pool[i]["ValorDolares"]) || 0)
              + (parseFloat(pool[j]["ValorDolares"]) || 0);
      if (Math.abs(c - sumCant) <= tolC && Math.abs(v - sumVCUSD) <= tolV)
        return [pool[i], pool[j]];
    }
  }
  // Tríos
  for (let i = 0; i < n - 2; i++) {
    for (let j = i + 1; j < n - 1; j++) {
      for (let k = j + 1; k < n; k++) {
        const c = (parseFloat(pool[i]["CantidadUMComercial"]) || 0)
                + (parseFloat(pool[j]["CantidadUMComercial"]) || 0)
                + (parseFloat(pool[k]["CantidadUMComercial"]) || 0);
        const v = (parseFloat(pool[i]["ValorDolares"]) || 0)
                + (parseFloat(pool[j]["ValorDolares"]) || 0)
                + (parseFloat(pool[k]["ValorDolares"]) || 0);
        if (Math.abs(c - sumCant) <= tolC && Math.abs(v - sumVCUSD) <= tolV)
          return [pool[i], pool[j], pool[k]];
      }
    }
  }
  return null;
}

/**
 * E7  Precio unitario (ValorDolares / CantidadUMComercial) como discriminador.
 * Busca el candidato cuyo $/pieza sea más cercano al del Layout, con tolerancia ±tolPct.
 */
function tryMatchUnitPrice(candidates, sumCant, sumVCUSD, tolPct = 0.15) {
  if (sumCant <= 0) return null;
  const layoutUP = sumVCUSD / sumCant;
  let best = null, bestDiff = Infinity;
  for (const r of candidates) {
    const c = parseFloat(r["CantidadUMComercial"]) || 0;
    const v = parseFloat(r["ValorDolares"]) || 0;
    if (c <= 0) continue;
    const candUP = v / c;
    const ref    = Math.max(Math.abs(layoutUP), 0.0001);
    const diff   = Math.abs(candUP - layoutUP) / ref;
    if (diff <= tolPct && diff < bestDiff) { bestDiff = diff; best = r; }
  }
  return best ? { seq: best["SecuenciaFraccion"], r551: best } : null;
}

function normDescLite(s) {
  return String(s ?? "")
    .toLowerCase()
    .replace(/\s+/g, " ")
    .replace(/[^\p{L}\p{N}\s]/gu, " ")
    .trim();
}

function tokenOverlapRatio(a, b) {
  const ta = new Set(normDescLite(a).split(" ").filter((x) => x.length > 2));
  const tb = new Set(normDescLite(b).split(" ").filter((x) => x.length > 2));
  if (!ta.size || !tb.size) return 0;
  let inter = 0;
  for (const t of ta) if (tb.has(t)) inter++;
  return inter / Math.max(ta.size, tb.size);
}

/** Detecta si DS y Layout tienen pedimentos distintos (sin intersección).
 *  Retorna { ds, layout } con muestras si no hay match; null si hay coincidencia. */
function checkPedimentoMismatch(dsPedimentos, layoutPedimentos) {
  const norm = (p) => String(p ?? "").trim();
  const dsSet = new Set(dsPedimentos.map(norm).filter(Boolean));
  const laySet = new Set(layoutPedimentos.map(norm).filter(Boolean));
  const intersection = [...dsSet].filter((p) => laySet.has(p));
  if (intersection.length === 0 && dsSet.size > 0 && laySet.size > 0) {
    return { ds: [...dsSet].slice(0, 5), layout: [...laySet].slice(0, 5) };
  }
  return null;
}

function getPedimentosFromRows(rows, ...keys) {
  const out = new Set();
  for (const r of rows) {
    for (const k of keys) {
      const v = String(r[k] ?? "").trim();
      if (v) out.add(v);
    }
  }
  return [...out];
}

function runCascade(layoutRows, s551Rows) {
  //  Columnas del Layout (ya vienen normalizadas por readLayoutSheet) 
  const L_PED   = "Pedimento";
  const L_FRAC  = "FraccionNico";
  const L_PAIS  = "PaisOrigen";
  const L_CANT  = "CantidadSaldo";
  const L_VCUSD = "VCUSD";
  const L_SEC   = "SecuenciaPed";

  //  Columnas del 551 
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

  const L_PEDK = "_pedCruce";

  const layout = layoutRows.map((r, i) => {
    const pedRaw =
      r[L_PED] != null && String(r[L_PED]).trim() !== "" ? r[L_PED] : (r.pedimento ?? r.pedimento1 ?? "");
    const pedKey = pedCruceKey(pedRaw);
    const fracKey = nFrac(r[L_FRAC]);
    const cantNum = parseFloat(r[L_CANT]);
    const valNum = parseFloat(r[L_VCUSD]);
    const skipReasons = [];
    if (!pedKey) skipReasons.push("pedimento vacío o no interpretable");
    if (!fracKey || fracKey === "0") skipReasons.push("fracción vacía/no válida");
    if (!isFinite(cantNum) || cantNum === 0) skipReasons.push("CantidadSaldo vacía o 0");
    if (!isFinite(valNum) || valNum === 0) skipReasons.push("VCUSD vacío o 0");
    const skipReason = skipReasons.length > 0
      ? `No se buscó secuencia: ${skipReasons.join("; ")}.`
      : "";
    return {
      ...r,
      [L_PED]: pedRaw,
      _idx: i,
      _frac: fracKey,
      _sec:  nSec(r[L_SEC]),
      [L_PEDK]: pedKey,
      _skipReason: skipReason,
    };
  });

  const s551 = s551Rows.map((r, i) => {
    const pk = pedCruceKey(r.Pedimento2) || pedCruceKey(r[S_PED]) || "";
    return {
      ...r,
      _frac:   nFrac(r[S_FRAC]),
      _551idx: i,
      [L_PEDK]: pk,
    };
  });

  //  Totales globales para validar si Layout y 551 balancean 
  const globalTotals = {
    layoutCant:  layout.reduce((a, r) => a + (parseFloat(r[L_CANT])  || 0), 0),
    layoutVCUSD: layout.reduce((a, r) => a + (parseFloat(r[L_VCUSD]) || 0), 0),
    s551Cant:    s551.reduce((a, r) => a + (parseFloat(r["CantidadUMComercial"]) || 0), 0),
    s551Val:     s551.reduce((a, r) => a + (parseFloat(r["ValorDolares"])        || 0), 0),
  };

  //  Set de todas las fracciones en el 551 (para diagnóstico) 
  const fracSet551 = new Set(s551.map((r) => r._frac));

  //  Lookups del 551 
  const lookupPFP   = new Map();  // Pedimento + Fraccion + Pais
  const lookupPF    = new Map();  // Pedimento + Fraccion (sin país)
  const lookupP     = new Map();  // Solo Pedimento
  const lookupPChap = new Map();  // Pedimento + capítulo (primeros 4 dígitos de fracción)

  for (const r of s551) {
    const pK = String(r._pedCruce ?? "");
    const k1 = `${pK}|||${r._frac}|||${String(r[S_PAIS] ?? "").trim()}`;
    if (!lookupPFP.has(k1)) lookupPFP.set(k1, []);
    lookupPFP.get(k1).push(r);

    const k2 = `${pK}|||${r._frac}`;
    if (!lookupPF.has(k2)) lookupPF.set(k2, []);
    lookupPF.get(k2).push(r);

    const kP = String(pK ?? "");
    if (!lookupP.has(kP)) lookupP.set(kP, []);
    lookupP.get(kP).push(r);

    const chap = r._frac.slice(0, 4);
    const kPC  = `${pK}|||${chap}`;
    if (!lookupPChap.has(kPC)) lookupPChap.set(kPC, []);
    lookupPChap.get(kPC).push(r);
  }

  //  Tracking 
  const assignment    = new Map();
  const assigned      = new Set();
  const used551       = new Set();
  const correctionMap = new Map(); // rowIdx  [{field, original, corrected}]
  const rowNotes      = new Map();
  const strategyStats = { E0: 0, E1: 0, E2: 0, E3: 0, E4: 0, E5: 0, E6: 0, E7: 0, E8: 0, E9: 0, E10: 0, E11: 0, R1: 0, R2: 0, R3: 0, R4: 0, R5: 0 };
  const pendingRows = () => layout.filter((r) => !assigned.has(r._idx) && !r._skipReason);

  // Registrar desde el inicio por qué NO se busca secuencia en ciertas filas
  for (const r of layout) {
    if (r._skipReason) rowNotes.set(r._idx, r._skipReason);
  }

  // Asigna filas a secuencia 551. NO modifica país, fracción ni descripción del Layout.
  const assignRows = (rows, seq, strategy, r551 = null, corrections = []) => {
    for (const r of rows) {
      if (!assigned.has(r._idx)) {
        assignment.set(r._idx, { seq, strategy, r551 });
        assigned.add(r._idx);
        strategyStats[strategy]++;
        if (r551?._551idx !== undefined) used551.add(r551._551idx);
        if (corrections && corrections.length > 0) {
          correctionMap.set(r._idx, corrections);
        }
      }
    }
  };

  //  E0: Match directo por clave compuesta CANDADO DS 551  551.Secuencias 
  // La columna "CANDADO DS 551" del Layout contiene la clave compuesta
  // "Pedimento-Fraccion-SecuenciaFraccion" que corresponde EXACTAMENTE a la
  // columna "Secuencias" del 551. Esto da asignación perfecta sin ningún cálculo.
  if (!STRICT_DS_RULES) {
    // Construir lookup: 551.Secuencias  fila del 551
    const lookupSecuencias = new Map();
    for (const r of s551) {
      const clave = String(r["Secuencias"] ?? "").trim();
      if (clave) lookupSecuencias.set(clave, r);
    }
    // Leer columna "CANDADO DS 551" de cada fila del Layout y asignar directamente
    for (const r of layout) {
      if (r._skipReason) continue;
      if (assigned.has(r._idx)) continue;
      const candado = String(r["CANDADO DS 551"] ?? "").trim();
      if (!candado) continue;
      const r551match = lookupSecuencias.get(candado);
      if (!r551match) continue;
      const seq = r551match["SecuenciaFraccion"];
      if (!seq && seq !== 0) continue;
      assignRows([r], seq, "E0", r551match);
    }
  }

  //  E1: Pedimento + Fracción + País, cantidades exactas (±1 ud / ±4 USD) 
  for (const g of groupBy(layout.filter((r) => !r._skipReason), ['_pedCruce', "_frac", L_PAIS])) {
    const k = `${g.keyVals[0]}|||${g.keyVals[1]}|||${String(g.keyVals[2] ?? "").trim()}`;
    const cands = lookupPFP.get(k) || [];
    const { cant, vcusd } = sumGroup(g.rows, L_CANT, L_VCUSD);
    const match = tryMatch(cands, cant, vcusd, 1, 4);
    if (match) assignRows(g.rows, match.seq, "E1", match.r551);
  }

  //  E2: Mismo Ped+Frac+País, sub-grupo por SecuenciaPed (solo numérico) 
  for (const g of groupBy(pendingRows(), ['_pedCruce', "_frac", L_PAIS, "_sec"])) {
    const k = `${g.keyVals[0]}|||${g.keyVals[1]}|||${String(g.keyVals[2] ?? "").trim()}`;
    const cands = lookupPFP.get(k) || [];
    const { cant, vcusd } = sumGroup(g.rows, L_CANT, L_VCUSD);
    const match = tryMatch(cands, cant, vcusd, 1, 4);
    if (match) assignRows(g.rows, match.seq, "E2", match.r551);
  }

  //  E3: Pedimento + Fracción (sin País) 
  for (const g of groupBy(pendingRows(), ['_pedCruce', "_frac"])) {
    const k = `${g.keyVals[0]}|||${g.keyVals[1]}`;
    const cands = lookupPF.get(k) || [];
    const { cant, vcusd } = sumGroup(g.rows, L_CANT, L_VCUSD);
    const match = tryMatch(cands, cant, vcusd, 1, 4);
    if (match) assignRows(g.rows, match.seq, "E3", match.r551);
  }

  //  E4: Ped+Frac (sin País) + sub-grupo por SecuenciaPed (solo numérico) 
  for (const g of groupBy(pendingRows(), ['_pedCruce', "_frac", "_sec"])) {
    const k = `${g.keyVals[0]}|||${g.keyVals[1]}`;
    const cands = lookupPF.get(k) || [];
    const { cant, vcusd } = sumGroup(g.rows, L_CANT, L_VCUSD);
    const match = tryMatch(cands, cant, vcusd, 1, 4);
    if (match) assignRows(g.rows, match.seq, "E4", match.r551);
  }

  //  E5: Tolerancia estricta ±1 ud / ±4 USD (igual que E1-E4) 
  for (const g of groupBy(pendingRows(), ['_pedCruce', "_frac", L_PAIS, "_sec"])) {
    const k = `${g.keyVals[0]}|||${g.keyVals[1]}`;
    const cands = lookupPF.get(k) || [];
    const { cant, vcusd } = sumGroup(g.rows, L_CANT, L_VCUSD);
    const match = tryMatch(cands, cant, vcusd, 1, 4);
    if (match) assignRows(g.rows, match.seq, "E5", match.r551);
  }

  //  E6: Suma de combinaciones (2 o 3 secuencias del 551) 
  // Resuelve cuando el mismo material ingresó en múltiples lotes del mismo pedimento
  // y la suma de 2 o 3 entradas del 551 iguala la suma del grupo Layout (±1 cant / ±4 val).
  for (const g of groupBy(pendingRows(), ['_pedCruce', "_frac", L_PAIS])) {
    const kPFP = `${g.keyVals[0]}|||${g.keyVals[1]}|||${String(g.keyVals[2] ?? "").trim()}`;
    const kPF  = `${g.keyVals[0]}|||${g.keyVals[1]}`;
    // Preferir candidatos con país coincidente; si no hay suficientes, usar sin país
    const pfpCands = lookupPFP.get(kPFP) || [];
    const cands    = pfpCands.length >= 2 ? pfpCands : (lookupPF.get(kPF) || []);
    if (cands.length < 2) continue;
    const { cant, vcusd } = sumGroup(g.rows, L_CANT, L_VCUSD);
    const combo = tryMatchCombination(cands, cant, vcusd, 1, 4);
    if (!combo) continue;

    // Particionar filas del Layout entre los candidatos: greedy por quota restante
    const sorted  = [...g.rows].sort((a, b) => (parseFloat(b[L_CANT]) || 0) - (parseFloat(a[L_CANT]) || 0));
    const targets = combo.map((r) => ({
      r551:      r,
      remaining: parseFloat(r["CantidadUMComercial"]) || 0,
      rows:      [],
    }));
    for (const row of sorted) {
      const rowCant = parseFloat(row[L_CANT]) || 0;
      // Asignar a la partición con mayor quota restante
      let bestTi = 0;
      for (let ti = 1; ti < targets.length; ti++) {
        if (targets[ti].remaining > targets[bestTi].remaining) bestTi = ti;
      }
      targets[bestTi].rows.push(row);
      targets[bestTi].remaining -= rowCant;
    }
    for (const t of targets) {
      if (t.rows.length > 0) assignRows(t.rows, t.r551[S_SEQ], "E6", t.r551);
    }
  }

  //  E7: Precio unitario ($/pieza) como discriminador  solo si Cant±1 y Val±4 
  for (const g of groupBy(pendingRows(), ['_pedCruce', "_frac", L_PAIS, "_sec"])) {
    const kPFP = `${g.keyVals[0]}|||${g.keyVals[1]}|||${String(g.keyVals[2] ?? "").trim()}`;
    const kPF  = `${g.keyVals[0]}|||${g.keyVals[1]}`;
    const cands = (lookupPFP.get(kPFP) || []).length > 0
      ? lookupPFP.get(kPFP)
      : (lookupPF.get(kPF) || []);
    const { cant, vcusd } = sumGroup(g.rows, L_CANT, L_VCUSD);
    const match = tryMatchUnitPrice(cands, cant, vcusd, 0.15, 1, 4);
    if (match) assignRows(g.rows, match.seq, "E7", match.r551);
  }

  //  E8: Asignación por eliminación  solo si Cant±1 y Val±4 
  for (const g of groupBy(pendingRows(), ['_pedCruce', "_frac"])) {
    const k     = `${g.keyVals[0]}|||${g.keyVals[1]}`;
    const cands = (lookupPF.get(k) || []).filter((r) => !used551.has(r._551idx));
    if (cands.length === 0) continue;
    const { cant, vcusd } = sumGroup(g.rows, L_CANT, L_VCUSD);
    const match = cands.length === 1 && Math.abs(cant - (parseFloat(cands[0]["CantidadUMComercial"])||0)) <= 1 && Math.abs(vcusd - (parseFloat(cands[0]["ValorDolares"])||0)) <= 4
      ? { seq: cands[0][S_SEQ], r551: cands[0] }
      : tryMatchUnitPrice(cands, cant, vcusd, 0.30, 1, 4);
    if (match) assignRows(g.rows, match.seq, "E8", match.r551);
  }

  //  E9: Mismo capítulo arancelario (4 dígitos)  sin modificar Layout 
  for (const g of groupBy(pendingRows(), ['_pedCruce', "_frac", L_PAIS])) {
    const chap  = String(g.keyVals[1]).slice(0, 4);
    const kPC   = `${g.keyVals[0]}|||${chap}`;
    const cands = (lookupPChap.get(kPC) || []).filter((r) => !used551.has(r._551idx));
    if (!cands.length) continue;
    const { cant, vcusd } = sumGroup(g.rows, L_CANT, L_VCUSD);
    const match = tryMatch(cands, cant, vcusd, 1, 4) || tryMatchUnitPrice(cands, cant, vcusd, 0.20, 1, 4);
    if (match) assignRows(g.rows, match.seq, "E9", match.r551);
  }

  //  E10: Solo Pedimento + precio unitario  solo si Cant±1 y Val±4 
  for (const g of groupBy(pendingRows(), ['_pedCruce', "_frac", L_PAIS])) {
    const kP    = String(g.keyVals[0]);
    const cands = (lookupP.get(kP) || []).filter((r) => !used551.has(r._551idx));
    if (!cands.length) continue;
    const { cant, vcusd } = sumGroup(g.rows, L_CANT, L_VCUSD);
    const matchUP = tryMatchUnitPrice(cands, cant, vcusd, 0.25, 1, 4);
    const match   = matchUP || (cands.length === 1 && Math.abs(cant - (parseFloat(cands[0]["CantidadUMComercial"])||0)) <= 1 && Math.abs(vcusd - (parseFloat(cands[0]["ValorDolares"])||0)) <= 4
      ? { seq: cands[0][S_SEQ], r551: cands[0] } : null);
    if (match) assignRows(g.rows, match.seq, "E10", match.r551);
  }

  //  E11: Fuerza 1:1 greedy por precio unitario  solo si Cant±1 y Val±4 
  {
    const pendGrps = groupBy(pendingRows(), ['_pedCruce', "_frac", L_PAIS]);
    const pedMap   = new Map();
    for (const g of pendGrps) {
      const ped = String(g.keyVals[0]);
      if (!pedMap.has(ped)) pedMap.set(ped, []);
      const { cant, vcusd } = sumGroup(g.rows, L_CANT, L_VCUSD);
      pedMap.get(ped).push({ rows: g.rows, keyVals: g.keyVals,
        up: vcusd / Math.max(cant, 0.0001) });
    }
    for (const [ped, grps] of pedMap) {
      const avail = (lookupP.get(ped) || [])
        .filter((r) => !used551.has(r._551idx))
        .map((r) => ({ r551: r,
          up: (parseFloat(r["ValorDolares"]) || 0) /
              Math.max(parseFloat(r["CantidadUMComercial"]) || 1, 0.0001) }));
      if (!avail.length) continue;
      const usedInE11 = new Set();
      for (const grp of grps) {
        const unassigned = grp.rows.filter((r) => !assigned.has(r._idx));
        if (!unassigned.length) continue;
        const { cant, vcusd } = sumGroup(grp.rows, L_CANT, L_VCUSD);
        let best = null, bestDiff = Infinity;
        for (let si = 0; si < avail.length; si++) {
          if (usedInE11.has(si)) continue;
          const r = avail[si].r551;
          const c551 = parseFloat(r["CantidadUMComercial"]) || 0;
          const v551 = parseFloat(r["ValorDolares"]) || 0;
          if (Math.abs(cant - c551) > 1 || Math.abs(vcusd - v551) > 4) continue;
          const diff = Math.abs(avail[si].up - grp.up) / Math.max(grp.up, 0.0001);
          if (diff < bestDiff) { bestDiff = diff; best = si; }
        }
        if (best === null) continue;
        usedInE11.add(best);
        assignRows(unassigned, avail[best].r551[S_SEQ], "E11", avail[best].r551);
      }
    }
  }

  //  R4: Cobertura total por score (pedimento) 
  // Si aún queda algo sin asignar, elige la mejor secuencia del MISMO pedimento
  // minimizando una función de costo (cantidad, valor, $/pieza y similitud de descripción).
  {
    const pendR4 = groupBy(pendingRows(), ['_pedCruce', "_frac", L_PAIS]);
    for (const g of pendR4) {
      const [ped, frac] = g.keyVals;
      const cands = lookupP.get(String(ped)) || [];
      if (!cands.length) continue;
      const { cant, vcusd } = sumGroup(g.rows, L_CANT, L_VCUSD);
      const upL = vcusd / Math.max(cant, 0.0001);
      const descL = [...new Set(g.rows.map((r) => r["Descripcion"]).filter(Boolean))].join(" ");

      let best = null;
      let bestScore = Infinity;
      for (const c of cands) {
        const cCant = parseFloat(c["CantidadUMComercial"]) || 0;
        const cVal = parseFloat(c["ValorDolares"]) || 0;
        const upC = cVal / Math.max(cCant, 0.0001);
        const descC = c["DescripcionMercancia"] || "";

        const relCant = Math.abs(cant - cCant) / Math.max(cant, cCant, 1);
        const relVal = Math.abs(vcusd - cVal) / Math.max(vcusd, cVal, 1);
        const relUP = Math.abs(upL - upC) / Math.max(Math.abs(upL), Math.abs(upC), 0.0001);
        const descPenalty = 1 - tokenOverlapRatio(descL, descC);
        const fracPenalty = String(frac) === String(c._frac) ? 0 : 0.25;
        const score = relCant * 0.45 + relVal * 0.35 + relUP * 0.15 + descPenalty * 0.05 + fracPenalty;
        if (score < bestScore) { bestScore = score; best = c; }
      }
      if (!best) continue;
      const corrFrac = String(g.keyVals[1]) !== best._frac
        ? [{ field: "FraccionNico", original: String(g.keyVals[1]), corrected: best._frac }]
        : [];
      assignRows(g.rows, best[S_SEQ], "R4", best, corrFrac);
    }
  }

  //  R5: Cobertura total absoluta (global, incluso sin pedimento) 
  // ltimo recurso de emergencia: si no hay candidatos por pedimento, busca globalmente
  // por fracción y luego en todo el 551 para NO dejar filas sin secuencia.
  {
    const pendR5 = groupBy(pendingRows(), ['_pedCruce', "_frac", L_PAIS]);
    for (const g of pendR5) {
      const [, frac] = g.keyVals;
      const poolFrac = s551.filter((r) => String(r._frac) === String(frac));
      const cands = poolFrac.length > 0 ? poolFrac : s551;
      if (!cands.length) continue;
      const { cant, vcusd } = sumGroup(g.rows, L_CANT, L_VCUSD);
      const upL = vcusd / Math.max(cant, 0.0001);
      let best = null;
      let bestScore = Infinity;
      for (const c of cands) {
        const cCant = parseFloat(c["CantidadUMComercial"]) || 0;
        const cVal = parseFloat(c["ValorDolares"]) || 0;
        const upC = cVal / Math.max(cCant, 0.0001);
        const relCant = Math.abs(cant - cCant) / Math.max(cant, cCant, 1);
        const relVal = Math.abs(vcusd - cVal) / Math.max(vcusd, cVal, 1);
        const relUP = Math.abs(upL - upC) / Math.max(Math.abs(upL), Math.abs(upC), 0.0001);
        const score = relCant * 0.5 + relVal * 0.35 + relUP * 0.15;
        if (score < bestScore) { bestScore = score; best = c; }
      }
      if (!best) continue;
      const corrFrac = String(g.keyVals[1]) !== best._frac
        ? [{ field: "FraccionNico", original: String(g.keyVals[1]), corrected: best._frac }]
        : [];
      assignRows(g.rows, best[S_SEQ], "R5", best, corrFrac);
    }
  }

  //  Layout lookup por Ped+Frac (para diagnóstico de orphans del 551) 
  const layoutPF = new Map();
  for (const r of layout) {
    const k = `${r._pedCruce}|||${r._frac}`;
    if (!layoutPF.has(k)) layoutPF.set(k, []);
    layoutPF.get(k).push(r);
  }

  //  R1: Barrido inverso  precio unitario ±30%, SIN filtro used551 
  // Las E anteriores solo usan secuencias 551 "libres". R1 permite reusar cualquier
  // secuencia 551 (incluso asignada) con tolerancia más amplia de precio/pieza.
  for (const g of groupBy(pendingRows(), ['_pedCruce', "_frac", L_PAIS, "_sec"])) {
    const kPFP = `${g.keyVals[0]}|||${g.keyVals[1]}|||${String(g.keyVals[2] ?? "").trim()}`;
    const kPF  = `${g.keyVals[0]}|||${g.keyVals[1]}`;
    const cands = (lookupPFP.get(kPFP) || []).length > 0
      ? lookupPFP.get(kPFP)
      : (lookupPF.get(kPF) || []);
    if (!cands.length) continue;
    const { cant, vcusd } = sumGroup(g.rows, L_CANT, L_VCUSD);
    const match = tryMatch(cands, cant, vcusd, 3, 8)
               || tryMatchUnitPrice(cands, cant, vcusd, 0.30);
    if (match) assignRows(g.rows, match.seq, "R1", match.r551);
  }

  //  R2: Solo Pedimento  sin filtro used551  corrige Fracción y País 
  // Busca en todo el pedimento sin importar fracción ni país, con precio ±40%.
  for (const g of groupBy(pendingRows(), ['_pedCruce', "_frac", L_PAIS])) {
    const kP    = String(g.keyVals[0]);
    const cands = lookupP.get(kP) || [];
    if (!cands.length) continue;
    const { cant, vcusd } = sumGroup(g.rows, L_CANT, L_VCUSD);
    const match = tryMatchUnitPrice(cands, cant, vcusd, 0.40)
               || (cands.length === 1 ? { seq: cands[0][S_SEQ], r551: cands[0] } : null);
    if (!match) continue;
    const corrFrac = String(g.keyVals[1]) !== match.r551._frac
      ? [{ field: "FraccionNico", original: String(g.keyVals[1]), corrected: match.r551._frac }]
      : [];
    assignRows(g.rows, match.seq, "R2", match.r551, corrFrac);
  }

  //  R3: Fuerza total greedy  sin filtro, sin tolerancia 
  // ltimo recurso absoluto: empareja cada grupo Layout restante con la secuencia
  // 551 del mismo pedimento cuyo precio/pieza sea más cercano (sin importar diferencia).
  {
    const pendR3  = groupBy(pendingRows(), ['_pedCruce', "_frac", L_PAIS]);
    const pedMapR = new Map();
    for (const g of pendR3) {
      const ped = String(g.keyVals[0]);
      if (!pedMapR.has(ped)) pedMapR.set(ped, []);
      const { cant, vcusd } = sumGroup(g.rows, L_CANT, L_VCUSD);
      pedMapR.get(ped).push({ rows: g.rows, keyVals: g.keyVals,
        up: vcusd / Math.max(cant, 0.0001) });
    }
    for (const [ped, grps] of pedMapR) {
      const avail = (lookupP.get(ped) || []).map((r) => ({
        r551: r,
        up: (parseFloat(r["ValorDolares"]) || 0) /
            Math.max(parseFloat(r["CantidadUMComercial"]) || 1, 0.0001),
      }));
      if (!avail.length) continue;
      for (const grp of grps) {
        const unassigned = grp.rows.filter((r) => !assigned.has(r._idx));
        if (!unassigned.length) continue;
        let best = null, bestDiff = Infinity;
        for (const a of avail) {
          const diff = Math.abs(a.up - grp.up) / Math.max(grp.up, 0.0001);
          if (diff < bestDiff) { bestDiff = diff; best = a; }
        }
        if (!best) continue;
        const corrFrac = String(grp.keyVals[1]) !== best.r551._frac
          ? [{ field: "FraccionNico", original: String(grp.keyVals[1]), corrected: best.r551._frac }]
          : [];
        assignRows(unassigned, best.r551[S_SEQ], "R3", best.r551, corrFrac);
      }
    }
  }

  //  Secuencias del 551 que NO se usaron en ningún match (orphans) 
  const getOrphanReason = (r) => {
    const cant = parseFloat(r["CantidadUMComercial"]);
    const val  = parseFloat(r["ValorDolares"]);
    const cantZero = isNaN(cant) || cant === 0;
    const valZero  = isNaN(val)  || val  === 0;
    const seq  = r["SecuenciaFraccion"] ?? "?";
    const frac = r._frac ?? r[S_FRAC] ?? "?";
    const ped  = r._pedCruce != null && String(r._pedCruce).trim() !== "" ? r._pedCruce : (r[S_PED] ?? "?");

    if (cantZero && valZero)
      return `Sec.${seq}  Sin cantidad ni valor: CantidadUMComercial=0 y ValorDolares=0`;
    if (cantZero)
      return `Sec.${seq}  CantidadUMComercial=0 (sin cantidad registrada en el 551)`;
    if (valZero)
      return `Sec.${seq}  ValorDolares=0 (sin valor en dólares registrado en el 551)`;

    const kPF = `${ped}|||${frac}`;
    if (!layoutPF.has(kPF))
      return `Sec.${seq}  Pedimento (cruce) ${ped} / Fracción ${frac} no tiene partidas en Layout`;

    // Esta secuencia 551 no recibió ninguna fila del Layout. El total Layout es de TODA la fracción (otras secuencias sí pueden tener asignación).
    const layoutCands = layoutPF.get(kPF);
    const sumCant = layoutCands.reduce((a, lr) => a + (parseFloat(lr[L_CANT]) || 0), 0);
    const sumVal  = layoutCands.reduce((a, lr) => a + (parseFloat(lr[L_VCUSD]) || 0), 0);
    const sinAsignar = layoutCands.filter((lr) => !assigned.has(lr._idx));
    const sumSinAsignarC = sinAsignar.reduce((a, lr) => a + (parseFloat(lr[L_CANT]) || 0), 0);
    const sumSinAsignarV = sinAsignar.reduce((a, lr) => a + (parseFloat(lr[L_VCUSD]) || 0), 0);
    const parte = sumSinAsignarC > 0 || sumSinAsignarV > 0
      ? ` Layout sin asignar (esta fracción): ${sumSinAsignarC.toFixed(0)} ud / $${sumSinAsignarV.toFixed(2)}.`
      : "";
    return `Sec.${seq}  Ninguna fila del Layout asignada a esta secuencia. Total Layout (ped+fracción): ${sumCant.toFixed(0)} ud / $${sumVal.toFixed(2)}; esta línea 551: ${cant.toFixed(0)} ud / $${val.toFixed(2)}.${parte}`;
  };

  const orphan551Rows = s551
    .filter((r) => !used551.has(r._551idx))
    .map((r)  => ({ ...r, _orphanReason: getOrphanReason(r) }));

  //  Diagnóstico por grupo sin match 
  const computeGroupNote = (ped, frac, pais, cant, vcusd) => {
    if (!fracSet551.has(frac)) {
      return `Fracción arancelaria ${frac} no registrada en el 551`;
    }
    const candsPF = lookupPF.get(`${ped}|||${frac}`) || [];
    if (candsPF.length === 0) {
      const otrosPed = [...new Set(s551.filter((r) => r._frac === frac).map((r) => r._pedCruce || r[S_PED]))].join(", ");
      return `Fracción ${frac} no encontrada para clave de pedimento ${ped}. Aparece en: ${otrosPed || "ninguno"}`;
    }
    const candsPFValid = candsPF.filter((r) => (parseFloat(r.CantidadUMComercial) || 0) > 0 || (parseFloat(r.ValorDolares) || 0) > 0);
    if (candsPFValid.length === 0) {
      return `No se buscó secuencia con datos de 551 para ${ped}/${frac}: los candidatos tienen cantidad y valor en 0 o vacíos.`;
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
             `(hay ${candsPFP.length} candidatos; requiere sub-agrupación manual)`;
    }
    return "No se encontró correspondencia exacta en 551";
  };

  // Construir mapa rowIdx  nota para filas sin match
  // rowNotes se inicializa arriba para conservar razones de filas bloqueadas
  const unmatchedGroups = groupBy(
    pendingRows(),
    ['_pedCruce', "_frac", L_PAIS]
  );
  for (const g of unmatchedGroups) {
    const [ped, frac, pais] = g.keyVals;
    const { cant, vcusd } = sumGroup(g.rows, L_CANT, L_VCUSD);
    const nota = computeGroupNote(ped, frac, pais, cant, vcusd);
    for (const r of g.rows) rowNotes.set(r._idx, nota);
  }

  //  Construir lista de sin-match para la UI 
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

  //  Construir datos para la hoja Cruce_Layout_vs_551 
  // Un registro por GRUPO (Ped + Frac + Pais + SecuenciaPedAsignada)
  const cruceData = [];

  // rowMatchMap: rowIdx  { okCant, okVal, diffCant, diffVal, cant551, val551, layoutCant, layoutVCUSD }
  // Permite que buildOutputExcel marque en rojo las filas donde los totales Layout  551.
  const rowMatchMap = new Map();

  // Grupos ASIGNADOS: agrupar por (seq asignada + frac + pais + ped)
  const matchedGroupMap = new Map();
  for (const [rowIdx, info] of assignment) {
    const r = layout[rowIdx];
    const gk = `${r._pedCruce}|||${r._frac}|||${String(r[L_PAIS] || "").trim()}|||${info.seq}`;
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

    const softStrat = new Set(["R1", "R2", "R3", "R4", "R5", "E6", "E7", "E8", "E9", "E10", "E11"]);
    const estr = g.info.strategy || "";
    const weakByDelta = !okCant || !okVal;
    const weakByStrat = softStrat.has(estr);
    const motivoCruce = weakByDelta
      ? `Diferencia grupo vs renglón 551 asignado: �cant = ${diffCant != null ? diffCant.toFixed(0) : "�"} ud; �USD = $${diffVal != null ? diffVal.toFixed(2) : "�"} (supera tolerancia estricta: ±1 ud y ±$2 al nivel de grupo).`
      : weakByStrat
        ? `Asignado con ${estr} (criterio flexible/combinación/respaldo; puede exceder tolerancias estrictas, validar manualmente).`
        : "Cruce estricto con el renglón 551 asignado (cantidad y valor dentro de tolerancias al agrupar).";
    const severidad = weakByDelta ? "alta" : (weakByStrat ? "media" : "ok");

    for (const row of g.rows) {
      rowMatchMap.set(row._idx, {
        okCant, okVal, diffCant, diffVal, cant551, val551, layoutCant: cant, layoutVCUSD: vcusd,
        strategy: estr, motivo: motivoCruce, severidad, weakByDelta, weakByStrat,
      });
    }

    // Descripciones únicas de las partes en el grupo
    const descs = [...new Set(g.rows.map((r) => r["Descripcion"]).filter(Boolean))].join(" / ");

    cruceData.push({
      tipo:       "matched",
      estrategia: g.info.strategy,
      numFilas:   g.rows.length,
      pedimento:  g.firstRow[L_PED],
      pedimentoCruce: g.firstRow._pedCruce,
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
      motivo: motivoCruce,
      severidad,
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
      pedimento:  g.rows[0][L_PED] || String(ped),
      pedimentoCruce: String(ped),
      layoutDesc:  descs,
      layoutFrac:  frac,
      layoutPais:  pais,
      layoutCant:  cant,
      layoutVCUSD: vcusd,
      secOriginal: g.rows[0][L_SEC],
      secAsignada: "",
      s551Secuencias: best ? (best["Secuencias"] || `${best[S_PED]}-${best[S_FRAC]}-${best[S_SEQ]}`) : "",
      s551Desc:   best ? (best["DescripcionMercancia"] || "") : " Sin candidato en 551 ",
      s551Frac:   best ? (best[S_FRAC] || "") : "",
      s551Pais:   best ? (best[S_PAIS] || "") : "",
      s551Cant:   best ? (parseFloat(best.CantidadUMComercial) || 0) : null,
      s551Val:    best ? (parseFloat(best.ValorDolares)        || 0) : null,
      diffCant:   best ? cant - (parseFloat(best.CantidadUMComercial) || 0) : null,
      diffVal:    best ? vcusd - (parseFloat(best.ValorDolares)       || 0) : null,
      okFrac: false, okPais: false, okCant: false, okVal: false,
      nota: rowNotes.get(g.rows[0]._idx) || "",
      motivo: rowNotes.get(g.rows[0]._idx) || "Ninguna estrategia pudo asignar secuencia a este grupo (revisa totales, país, fracción o columnas clave).",
      severidad: "alta",
    });
  }

  // Ordenar: primero sin-match, luego por estrategia
  cruceData.sort((a, b) => {
    if (a.tipo !== b.tipo) return a.tipo === "unmatched" ? -1 : 1;
    return (a.estrategia || "").localeCompare(b.estrategia || "");
  });

  return { assignment, strategyStats, unmatchedFinal, total: layout.length, rowNotes, cruceData, orphan551Rows, correctionMap, globalTotals, rowMatchMap };
}

//  EXCEL BUILDER 
function buildOutputExcel(workbook, layoutSheet, sheet551, sheet551Name, assignment, unmatchedFinal, stats, total, rowNotes, cruceData, orphan551Rows, correctionMap, globalTotals, rowMatchMap = new Map(), layoutOutName = "Layout") {
  const wb = XLSX.utils.book_new();

  //  Datos originales del Layout 
  const layoutData = XLSX.utils.sheet_to_json(layoutSheet, { header: 1 });
  const headerI = findLayoutHeaderRowIndex(layoutData);
  const rawHeaders  = layoutData[headerI] || [];

  // Helper: busca la primera columna cuyo nombre (sin espacios, en minúsculas) coincida
  // con alguno de los nombres candidatos. Tolera variaciones de mayúsculas/minúsculas y espacios.
  const normH = (s) => String(s ?? "").trim().toLowerCase().replace(/[\s_\-]/g, "");
  const findCol = (...names) => {
    const targets = names.map(normH);
    return rawHeaders.findIndex((h) => targets.includes(normH(h)));
  };

  // Buscar columnas clave con tolerancia a variaciones de nombre
  const secIdx      = findCol("SecuenciaPed", "Secuencia Ped", "Secuencia");
  const paisIdx     = findCol("PaisOrigen", "Pais Origen", "PaisOrigenDestino", "Pais");
  const fracIdx     = findCol("FraccionNico", "Fraccion Nico", "FraccionArancelaria", "Fraccion");
  const descIdx     = findCol("Descripcion", "DescripcionMercancia", "Descripcion Mercancia");
  const pedIdx      = findCol("Pedimento", "NumPedimento", "Num Pedimento");
  const cantIdx     = findCol("CantidadSaldo", "Cantidad Saldo", "CantidadUMComercial", "Cantidad");
  const vcusdIdx    = findCol("VCUSD", "ValorDolares", "Valor Dolares", "ValorComercialUSD");
  const candadoIdx  = findCol("CANDADO DS 551", "CANDADODS551", "CandadoDS551");
  const notasIdx    = rawHeaders.length;  // nueva columna al final
  const headers     = [...rawHeaders, "Notas"];

  // Normaliza SecuenciaPed para comparación (número o texto limpio)
  const normSeq = (v) => {
    const n = parseFloat(v);
    return isNaN(n) ? String(v ?? "").trim() : String(Math.round(n));
  };

  //  Construir filas + registrar cambios 
  // changeMap: rowIdx  { original, nuevo, tipoNota }
  //   tipoNota: "nuevo"     la celda estaba vacía/texto y ahora tiene valor
  //             "cambio"    el valor cambió de un número a otro
  //             "igual"     el valor nuevo es igual al original (sin marcado)
  //             "sinmatch"  no se asignó secuencia
  const changeMap = new Map();
  const updatedRows = [headers];

  // Helper local para normalizar fracción (igual que en runCascade)
  const nFracLocal = (v) => String(v ?? "").trim().replace(/^0+/, "") || "0";

  // rowIdx paralelo que sólo cuenta filas NO vacías, igual que readLayoutSheet/_idx
  let dataRowIdx = 0;

  for (let i = headerI + 1; i < layoutData.length; i++) {
    const rawDataRow = layoutData[i] || [];
    const isEmpty    = rawDataRow.every((c) => c === "" || c == null);
    const row        = [...rawDataRow];
    while (row.length <= notasIdx) row.push("");

    if (isEmpty) {
      // Fila vacía: conservar sin tocar y no consumir índice
      updatedRows.push(row);
      continue;
    }

    const rowIdx      = dataRowIdx++;   // sincronizado con _idx de runCascade
    const originalRaw = secIdx >= 0 ? (rawDataRow[secIdx] ?? "") : "";
    const originalStr = normSeq(originalRaw);

    //  Correcciones extra que vienen de estrategias específicas (E9, E10, R2) 
    const extraCorrs = correctionMap ? (correctionMap.get(rowIdx) || []) : [];
    for (const corr of extraCorrs) {
      const colIdx = corr.field === "PaisOrigen"   ? paisIdx
                   : corr.field === "FraccionNico" ? fracIdx
                   : corr.field === "Descripcion"  ? descIdx : -1;
      if (colIdx >= 0) row[colIdx] = corr.corrected;
    }

    //  Correcciones DIRECTAS desde r551 (siempre, para todos los campos clave) 
    // Independientes de correctionMap para garantizar que siempre se apliquen.
    const directCorrs = [];
    if (assignment.has(rowIdx)) {
      const r551 = assignment.get(rowIdx).r551;
      if (r551) {
        const camposDef = [
          { field: "PaisOrigen",   colIdx: paisIdx,    s551Key: "PaisOrigenDestino",    equal: (a, b) => a === b },
          { field: "FraccionNico", colIdx: fracIdx,    s551Key: "Fraccion",             equal: (a, b) => nFracLocal(a) === nFracLocal(b) },
          { field: "Descripcion",  colIdx: descIdx,    s551Key: "DescripcionMercancia", equal: (a, b) => a === b },
          { field: "Pedimento",    colIdx: pedIdx,     s551Key: "Pedimento",            equal: (a, b) => String(a).trim() === String(b).trim() },
          { field: "CandadoDS551", colIdx: candadoIdx, s551Key: "Secuencias",           equal: (a, b) => a === b },
        ];
        for (const def of camposDef) {
          const val551 = String(r551[def.s551Key] ?? "").trim();
          if (!val551) continue;
          const valL   = String(def.colIdx >= 0 ? (row[def.colIdx] ?? "") : "").trim();
          if (!def.equal(valL, val551)) {
            // No duplicar si extraCorrs ya tiene este campo
            if (!extraCorrs.some((c) => c.field === def.field)) {
              directCorrs.push({ field: def.field, original: valL, corrected: val551 });
            }
            if (def.colIdx >= 0) row[def.colIdx] = val551;  // escribir valor del 551
          }
        }
      }
    }

    const corrections = [...extraCorrs, ...directCorrs];
    const corrNote = corrections.length > 0
      ? corrections.map((c) =>
          `[CORRECCI�N ${c.field}] '${c.original}' -> '${c.corrected}' (fuente: 551)`
        ).join(" | ")
      : "";

    if (assignment.has(rowIdx)) {
      const rawSeq  = assignment.get(rowIdx).seq;
      const newSeq  = parseFloat(rawSeq) || rawSeq;
      const newStr  = normSeq(rawSeq);

      row[secIdx] = newSeq;

      if (newStr !== originalStr) {
        const wasEmpty = (originalStr === "" || isNaN(parseFloat(originalRaw)));
        const tipo = corrections.length > 0
          ? (wasEmpty ? "nuevo_corr" : "cambio_corr")
          : (wasEmpty ? "nuevo" : "cambio");
        let nota = wasEmpty
          ? `Secuencia asignada: ${newStr}`
          : `Secuencia modificada: ${originalStr} -> ${newStr}`;
        if (corrNote) nota += `. ${corrNote}`;
        row[notasIdx] = nota;
        changeMap.set(rowIdx, { original: originalStr, nuevo: newStr, tipo, nota, corrections });
      } else {
        row[notasIdx] = corrNote || "";
        const tipo = corrections.length > 0 ? "igual_corr" : "igual";
        changeMap.set(rowIdx, { original: originalStr, nuevo: newStr, tipo, corrections });
      }
    } else {
      row[secIdx]   = "";
      row[notasIdx] = rowNotes ? (rowNotes.get(rowIdx) || "") : "";
      changeMap.set(rowIdx, { tipo: "sinmatch" });
    }

    // Señalar discrepancia de totales Layout vs 551 si la hay
    if (assignment.has(rowIdx) && rowMatchMap) {
      const mq = rowMatchMap.get(rowIdx);
      if (mq && (!mq.okCant || !mq.okVal)) {
        const dC = mq.diffCant != null ? ((mq.diffCant > 0 ? "+" : "") + mq.diffCant.toFixed(0)) : "?";
        const dV = mq.diffVal  != null ? ((mq.diffVal  > 0 ? "+" : "") + mq.diffVal.toFixed(2))  : "?";
        const badNote = `�a� DIFERENCIA SOBRE TOLERANCIA: Layout(${(mq.layoutCant||0).toFixed(0)} ud / $${(mq.layoutVCUSD||0).toFixed(2)}) vs 551(${mq.cant551 != null ? mq.cant551.toFixed(0) : "?"} ud / $${mq.val551 != null ? mq.val551.toFixed(2) : "?"}) | �cant:${dC} | �val:${dV}. Motivo: asignación de respaldo o datos no conciliables al umbral estricto.`;
        row[notasIdx] = row[notasIdx] ? `${row[notasIdx]} | ${badNote}` : badNote;
        const cm = changeMap.get(rowIdx);
        if (cm) cm.hasBadMatch = true;
      }
    }

    updatedRows.push(row);
  }

  //  Filas extra al final: secuencias del 551 que no se asignaron al Layout 
  const orphan551StartRow = updatedRows.length; // índice de fila en la hoja (0-based)

  if (orphan551Rows && orphan551Rows.length > 0) {
    // Fila separadora vacía
    updatedRows.push(Array(headers.length).fill(""));

    // Encabezado de sección
    const sectionHdr = Array(headers.length).fill("");
    sectionHdr[0] = ` FILAS AADIDAS DESDE EL 551  Secuencias sin partida en Layout (${orphan551Rows.length} registros)  campos llenados con datos del 551`;
    updatedRows.push(sectionHdr);

    // Usar los índices flexibles ya definidos arriba (findCol) + mapeo 551Layout
    const orphanFieldMap = [
      { colIdx: secIdx,     val: (r) => r["SecuenciaFraccion"]       ?? "" },
      { colIdx: pedIdx,     val: (r) => r["Pedimento"]               ?? "" },
      { colIdx: fracIdx,    val: (r) => r["Fraccion"]                ?? "" },
      { colIdx: paisIdx,    val: (r) => r["PaisOrigenDestino"]       ?? "" },
      { colIdx: cantIdx,    val: (r) => parseFloat(r["CantidadUMComercial"]) || 0 },
      { colIdx: vcusdIdx,   val: (r) => parseFloat(r["ValorDolares"])        || 0 },
      { colIdx: descIdx,    val: (r) => r["DescripcionMercancia"]    ?? "" },
      { colIdx: candadoIdx, val: (r) => r["Secuencias"]              ?? "" },
    ];

    for (const r551 of orphan551Rows) {
      const row = Array(headers.length).fill("");
      for (const { colIdx, val } of orphanFieldMap) {
        if (colIdx >= 0) row[colIdx] = val(r551);
      }
      row[notasIdx] = `FILA AADIDA DESDE 551  ${r551._orphanReason || "secuencia sin partida en Layout"}`;
      updatedRows.push(row);
    }
  }

  //  Crear worksheet y aplicar estilos 
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
  const S_CORRECTED_CELL = {   // celda corregida por la app (País/Fracción/Descripcion): rojo claro
    font: { bold: true, color: { rgb: "7B241C" } },
    fill: { patternType: "solid", fgColor: { rgb: "FFCCCC" } },
    alignment: { horizontal: "center", wrapText: true },
  };
  const S_NOTA_CORRECCION = {  // nota de corrección: fondo naranja claro
    font: { bold: true, sz: 10, color: { rgb: "784212" } },
    fill: { patternType: "solid", fgColor: { rgb: "FDEBD0" } },
    alignment: { wrapText: true },
  };
  const S_SEQ_DISCREPANCIA = {  // SecuenciaPed con diferencia de totales: fondo rojo intenso, texto blanco
    font: { bold: true, sz: 11, color: { rgb: "FFFFFF" } },
    fill: { patternType: "solid", fgColor: { rgb: "922B21" } },
    alignment: { horizontal: "center" },
  };
  const S_NOTA_DISCREPANCIA = {  // nota de diferencia de totales: rojo oscuro
    font: { bold: true, sz: 10, color: { rgb: "641E16" } },
    fill: { patternType: "solid", fgColor: { rgb: "FADBD8" } },
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
      if (info.tipo === "nuevo"  || info.tipo === "cambio" ||
          info.tipo === "nuevo_corr" || info.tipo === "cambio_corr") {
        wsLayout[addr].s = S_CHANGED;
      } else if (info.tipo === "igual" || info.tipo === "igual_corr") {
        wsLayout[addr].s = S_EQUAL;
      } else {
        wsLayout[addr].s = S_EMPTY;
      }
    }

    // Estilo Notas
    const addrN = XLSX.utils.encode_cell({ r: rowI, c: notasIdx });
    if (!wsLayout[addrN]) wsLayout[addrN] = { t: "s", v: "" };
    if (info.tipo === "nuevo"  || info.tipo === "cambio" ||
        info.tipo === "nuevo_corr" || info.tipo === "cambio_corr") {
      wsLayout[addrN].s = S_NOTA_CAMBIO;
    } else if (info.tipo === "igual_corr") {
      wsLayout[addrN].s = S_NOTA_CORRECCION;
    } else if (info.tipo === "sinmatch" && updatedRows[rowI][notasIdx]) {
      wsLayout[addrN].s = S_NOTA_SINMATCH;
    }

    // Pintar en rojo las celdas de campos corregidos (PaisOrigen, FraccionNico)
    if (info.corrections && info.corrections.length > 0) {
      for (const corr of info.corrections) {
        const colIdx = corr.field === "PaisOrigen"   ? paisIdx
                     : corr.field === "FraccionNico" ? fracIdx
                     : corr.field === "Descripcion"  ? descIdx : -1;
        if (colIdx < 0) continue;
        const addrCorr = XLSX.utils.encode_cell({ r: rowI, c: colIdx });
        if (!wsLayout[addrCorr]) wsLayout[addrCorr] = { t: "s", v: corr.corrected };
        wsLayout[addrCorr].s = S_CORRECTED_CELL;
      }
    }

    // Señalar en rojo cuando los totales Layout  551 para esta fila asignada
    if (info.hasBadMatch) {
      if (secIdx >= 0) {
        const addrSeqDisc = XLSX.utils.encode_cell({ r: rowI, c: secIdx });
        if (!wsLayout[addrSeqDisc]) wsLayout[addrSeqDisc] = { t: "s", v: "" };
        wsLayout[addrSeqDisc].s = S_SEQ_DISCREPANCIA;
      }
      const addrNotaDisc = XLSX.utils.encode_cell({ r: rowI, c: notasIdx });
      if (!wsLayout[addrNotaDisc]) wsLayout[addrNotaDisc] = { t: "s", v: "" };
      wsLayout[addrNotaDisc].s = S_NOTA_DISCREPANCIA;
    }
  }

  //  Estilos para filas de orphan 551 
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

  XLSX.utils.book_append_sheet(wb, wsLayout, layoutOutName);

  // Copy original 551 / Data Stage sheet
  const ws551 = XLSX.utils.sheet_to_json(sheet551, { header: 1 });
  const ws551Sheet = XLSX.utils.aoa_to_sheet(ws551);
  XLSX.utils.book_append_sheet(wb, ws551Sheet, sheet551Name || "551");

  //  Hoja Cruce_Layout_vs_551 
  if (cruceData && cruceData.length > 0) {
    const pct2 = (v, total) => total > 0 ? ((v / total) * 100).toFixed(1) + "%" : "";
    const fmt  = (v) => v == null ? "" : (typeof v === "number" ? Number(v.toFixed(4)) : v);
    const fmtDiff = (d) => d == null ? "" : (d > 0 ? `+${d.toFixed(2)}` : d.toFixed(2));
    const check = (ok) => ok ? "" : "";

    // Cabeceras con dos niveles
    const HEADERS = [
      // Bloque General
      "Estrategia", "N° Filas Layout", "Pedimento",
      // Bloque LAYOUT
      "Layout  Descripcion", "Layout  FraccionNico", "Layout  PaisOrigen",
      "Layout  Suma CantidadSaldo", "Layout  Suma VCUSD",
      "SecuenciaPed Original", "SecuenciaPed ASIGNADA",
      // Bloque 551
      "551  Secuencias (clave)", "551  DescripcionMercancia",
      "551  Fraccion", "551  PaisOrigenDestino",
      "551  CantidadUMComercial", "551  ValorDolares",
      // Comparación
      "¿Fracción coincide?", "¿País coincide?",
      "Dif. Cantidad (Layout551)", "Dif. Valor USD (Layout551)",
      "Estado match",
      // Nota (solo sin-match)
      "Notas / Motivo sin asignación",
    ];

    const cruceRows = [HEADERS];

    for (const d of cruceData) {
      const statusMatch = d.tipo === "unmatched"
        ? " SIN MATCH"
        : (d.okCant && d.okVal ? " MATCH EXACTO"
          : (Math.abs(d.diffCant || 0) / Math.max(1, d.layoutCant) < 0.05 ? " MATCH ~TOL.5%" : " MATCH TOLERADO"));

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

    //  Estilos hoja Cruce 
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

      // Columnas /
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

  //  Hoja Reporte_Secuencias_vs_551: totales globales + detalle por secuencia asignada 
  const matchedOnly = (cruceData || []).filter((d) => d.tipo === "matched");
  const reportSecRows = [];
  reportSecRows.push(["REPORTE  SECUENCIAS ASIGNADAS vs DATA 551"]);
  reportSecRows.push([]);

  // Total global Layout vs 551
  if (globalTotals) {
    reportSecRows.push(["TOTALES GLOBALES"]);
    reportSecRows.push(["", "Cantidad", "Valor USD"]);
    reportSecRows.push(["Layout (suma total)", Number(globalTotals.layoutCant.toFixed(4)), Number(globalTotals.layoutVCUSD.toFixed(4))]);
    reportSecRows.push(["551 (suma total)", Number(globalTotals.s551Cant.toFixed(4)), Number(globalTotals.s551Val.toFixed(4))]);
    const dC = globalTotals.layoutCant - globalTotals.s551Cant;
    const dV = globalTotals.layoutVCUSD - globalTotals.s551Val;
    reportSecRows.push(["Diferencia (Layout  551)", Number(dC.toFixed(4)), Number(dV.toFixed(4))]);
    const okGlobal = Math.abs(dC) < 1 && Math.abs(dV) < 2;
    reportSecRows.push(["¿Totales coinciden?", okGlobal ? "Sí" : "No"]);
    reportSecRows.push([]);
  }

  // Detalle por secuencia asignada: Layout vs 551 (cantidad, valor, país, descripción)
  reportSecRows.push(["DETALLE POR SECUENCIA ASIGNADA (Layout vs 551)"]);
  const detalleHeaders = [
    "Estrategia", "Secuencia asignada", "Pedimento",
    "Layout  Descripción", "Layout  Cantidad", "Layout  Valor USD",
    "551  Clave (Secuencias)", "551  Descripción", "551  Cantidad", "551  Valor USD",
    "¿Cantidad coincide?", "¿Valor coincide?", "¿País coincide?", "¿Descripción coincide?",
    "Estado match",
  ];
  reportSecRows.push(detalleHeaders);

  const fmtR = (v) => v == null ? "" : (typeof v === "number" ? Number(v.toFixed(4)) : String(v ?? ""));
  const checkR = (ok) => (ok ? "Sí" : "No");
  const descCoincide = (layoutDesc, s551Desc) => {
    const a = String(layoutDesc ?? "").trim();
    const b = String(s551Desc ?? "").trim();
    if (!a && !b) return true;
    return a === b;
  };

  for (const d of matchedOnly) {
    const statusMatch = d.okCant && d.okVal ? "Match exacto" : "Tolerado / diferencia";
    const descOk = descCoincide(d.layoutDesc, d.s551Desc);
    reportSecRows.push([
      d.estrategia ?? "",
      d.secAsignada ?? "",
      d.pedimento ?? "",
      (d.layoutDesc ?? "").toString().slice(0, 200),
      fmtR(d.layoutCant),
      fmtR(d.layoutVCUSD),
      (d.s551Secuencias ?? "").toString().slice(0, 80),
      (d.s551Desc ?? "").toString().slice(0, 200),
      fmtR(d.s551Cant),
      fmtR(d.s551Val),
      checkR(d.okCant),
      checkR(d.okVal),
      checkR(d.okPais),
      checkR(descOk),
      statusMatch,
    ]);
  }

  if (matchedOnly.length === 0) {
    reportSecRows.push(["Sin secuencias asignadas para detallar."]);
  }

  const wsReportSec = XLSX.utils.aoa_to_sheet(reportSecRows);
  const colWidthsReportSec = [10, 10, 14, 42, 14, 14, 38, 42, 14, 14, 14, 14, 14, 18, 18];
  wsReportSec["!cols"] = colWidthsReportSec.map((w) => ({ wch: Math.min(w, 50) }));
  XLSX.utils.book_append_sheet(wb, wsReportSec, "Reporte_Secuencias_vs_551");

  // Resultado_Validacion sheet
  const matched = total - unmatchedFinal.length;
  const pct = ((matched / total) * 100).toFixed(1);

  const reportRows = [
    ["RESULTADO DE VALIDACIN  CRUCE LAYOUT vs 551"],
    [],
    ["RESUMEN EJECUTIVO"],
    ["Total filas Layout", total],
    ["Filas con SecuenciaPed asignada", matched],
    ["Filas sin match", unmatchedFinal.length],
    ["PORCENTAJE DE XITO", `${pct}%`],
    [],
    ["DESGLOSE POR ESTRATEGIA"],
    ["Estrategia", "Filas Asignadas", "Descripción"],
    ["E0 - Match directo CANDADO DS 551",   stats.E0, "Usa columna 'CANDADO DS 551' del Layout como clave compuesta directa hacia 'Secuencias' del 551. Asignación perfecta sin cálculos. Requiere que el Layout tenga esta columna poblada."],
    ["E1 - Ped+Fracción+País exacto",      stats.E1, "Agrupa Layout por Pedimento+FraccionNico+País. Suma CantidadSaldo y VCUSD; busca en 551 con tolerancia ±1 ud / ±2 USD."],
    ["E2 - Sub-grupo por SecuenciaPed",   stats.E2, "Misma clave E1 pero sub-divide por SecuenciaPed existente. Resuelve cuando la misma fracción+país tiene múltiples entradas en el 551."],
    ["E3 - Sin País (solo Ped+Fracción)", stats.E3, "Ignora PaisOrigen para manejar diferencias de código de país entre Layout y 551. Cantidades y valores exactos."],
    ["E4 - Sin País + Sub-SecuenciaPed",  stats.E4, "Combina E3 y E2: sin País + sub-agrupación por SecuenciaPed. Captura casos con variación de código país Y múltiples secuencias."],
    ["E5 - Tolerancia Ampliada ±5%",      stats.E5, "Tolerancia ±5% en cantidad y valor. Resuelve diferencias de redondeo o conversión de unidades entre sistemas."],
    ["E6 - Suma de combinaciones (2+3)",  stats.E6, "Evalúa si la suma de 2 o 3 secuencias del 551 iguala el total del grupo Layout (±2%). Detecta materiales importados en múltiples lotes del mismo pedimento y particiona las filas del Layout por cuota."],
    ["E7 - Precio unitario ±15%",         stats.E7, "Usa el precio por unidad ($/pieza) como discriminador con tolerancia ±15%. Resuelve casos donde las cantidades totales no coinciden pero el precio unitario confirma la correspondencia correcta."],
    ["E8 - Eliminación por descarte",     stats.E8, "Filtra candidatos del 551 ya usados y asigna el único remanente (o el más cercano en precio unitario ±30%). Válido cuando el material no tiene otra posible correspondencia."],
    ["E9 - Corrección de Fracción (capítulo)", stats.E9, "Si la fracción del Layout no existe en el 551 para ese pedimento pero sí existe otra del mismo capítulo (4 dígitos), corrige FraccionNico con el valor del 551. La celda corregida aparece en ROJO en el Excel."],
    ["E10 - Solo Pedimento + precio unitario", stats.E10, "Búsqueda en todo el pedimento ignorando fracción y país. Precio unitario ±25% o único candidato disponible. Corrige FraccionNico y/o PaisOrigen según el 551. Celdas corregidas en ROJO."],
    ["E11 - Fuerza 1:1 greedy",           stats.E11, "Empareja grupos Layout pendientes con secuencias 551 no usadas del mismo pedimento, ordenando por precio unitario (sin límite de tolerancia). Aplica todas las correcciones necesarias. Celdas corregidas en ROJO."],
    ["R1 - Barrido inverso Ped+Frac ±30%", stats.R1, "Barrido inverso sin filtro used551: usa cualquier sec. 551 del mismo Ped+Frac, precio ±30% o exacto. Elimina restricción de secuencias 'ya usadas' para maximizar asignación."],
    ["R2 - Barrido inverso solo Pedimento ±40%", stats.R2, "Busca en todo el pedimento sin restricción de fracción/país, precio ±40%. Corrige Fracción y País si difieren del 551. Para casos con fracción completamente diferente en Layout."],
    ["R3 - Fuerza total greedy sin filtro", stats.R3, "Fuerza máxima: empareja todos los grupos Layout restantes con cualquier sec. 551 del mismo pedimento, sin filtro ni tolerancia, eligiendo siempre la de precio/pieza más cercano."],
    ["R4 - Cobertura total por score (mismo pedimento)", stats.R4, "Optimización tardía: para cada grupo pendiente busca la mejor secuencia del mismo pedimento minimizando score combinado de diferencia en cantidad, valor, precio unitario y similitud de descripción."],
    ["R5 - Cobertura total absoluta (global)", stats.R5, "ltimo recurso extremo: si no hay candidatos por pedimento, busca por fracción a nivel global y, si tampoco existe, contra todo el 551. Prioriza mínima diferencia de cantidad/valor/precio para no dejar filas sin secuencia."],
    [],
  ];

  if (globalTotals) {
    const fmt4 = (n) => Number(n.toFixed(4));
    const diffCant  = globalTotals.layoutCant  - globalTotals.s551Cant;
    const diffVal   = globalTotals.layoutVCUSD - globalTotals.s551Val;
    const pctC = globalTotals.s551Cant  > 0 ? ((diffCant  / globalTotals.s551Cant)  * 100).toFixed(2) + "%" : "N/A";
    const pctV = globalTotals.s551Val   > 0 ? ((diffVal   / globalTotals.s551Val)   * 100).toFixed(2) + "%" : "N/A";
    const balance = Math.abs(diffCant) < 1 && Math.abs(diffVal) < 2 ? " BALANCE EXACTO" : " DIFERENCIA";
    reportRows.push(["VALIDACIN DE TOTALES GLOBALES (Layout vs 551)"]);
    reportRows.push(["", "Cantidad total", "Valor USD total"]);
    reportRows.push(["Layout (suma CantidadSaldo + VCUSD)", fmt4(globalTotals.layoutCant), fmt4(globalTotals.layoutVCUSD)]);
    reportRows.push(["551   (suma CantidadUMC + ValorDolares)", fmt4(globalTotals.s551Cant),  fmt4(globalTotals.s551Val)]);
    reportRows.push(["Diferencia (Layout  551)", fmt4(diffCant), fmt4(diffVal)]);
    reportRows.push(["Diferencia %", pctC, pctV]);
    reportRows.push([balance]);
    reportRows.push([]);
  }

  if (unmatchedFinal.length > 0) {
    reportRows.push(["GRUPOS SIN MATCH  REVISIN MANUAL REQUERIDA"]);
    reportRows.push(["Descripcion", "FraccionNico", "PaisOrigen", "SecuenciaPed_Original", "CantidadSaldo", "VCUSD", "Notas (motivo sin asignación)"]);
    for (const u of unmatchedFinal) {
      reportRows.push([u.Descripcion, u.FraccionNico, u.PaisOrigen, u.SecuenciaPed_Original, u.CantidadSaldo, u.VCUSD, u.Nota || ""]);
    }
  } else {
    reportRows.push([" TODOS LOS GRUPOS TUVIERON MATCH EXITOSO"]);
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

//  COMPONENTS 
const STRATEGIES = [
  {
    id: "E0",
    name: "Match directo  CANDADO DS 551  Secuencias 551",
    desc: "Estrategia prioritaria: usa la columna 'CANDADO DS 551' del Layout (clave compuesta Ped-Fracción-Secuencia) para hacer match DIRECTO con la columna 'Secuencias' del 551. Asignación perfecta sin cálculos. Resuelve el 99%+ de los casos cuando el Layout tiene esta columna poblada.",
    color: "#00d4aa",
    icon: "",
  },
  {
    id: "E1",
    name: "Pedimento + Fracción + País",
    desc: "Agrupación exacta por Pedimento + FraccionNico + PaisOrigen. Suma CantidadSaldo vs CantidadUMComercial y VCUSD vs ValorDolares del 551 (tolerancia ±1 unidad / ±2 USD). Resuelve la mayoría de los casos.",
    color: "#22c55e",
    icon: "",
  },
  {
    id: "E2",
    name: "Sub-agrupación por SecuenciaPed",
    desc: "Para grupos que fallaron E1, sub-divide usando el SecuenciaPed existente como guía. Resuelve casos donde la misma fracción+país tiene múltiples líneas en el 551 (ej: mismo material importado en dos fechas distintas).",
    color: "#3b82f6",
    icon: "",
  },
  {
    id: "E3",
    name: "Sin filtro de País (Ped + Fracción)",
    desc: "Ignora PaisOrigen para manejar diferencias de captura de código de país entre Layout y 551 (ej: 'TWN' vs 'TAI', 'CHN' vs 'PRC'). Aplica las mismas tolerancias exactas de cantidad y valor.",
    color: "#f59e0b",
    icon: "",
  },
  {
    id: "E4",
    name: "Sin País + Sub-SecuenciaPed",
    desc: "Combina E3 y E2: sin filtro de País y sub-agrupación por SecuenciaPed. Captura casos donde hay variación de código de país Y múltiples secuencias para la misma fracción.",
    color: "#a855f7",
    icon: "",
  },
  {
    id: "E5",
    name: "Tolerancia Ampliada (±5%)",
    desc: "Tolerancia ±5% en cantidad (mín 2 unidades) y ±5% en valor (mín 5 USD). Resuelve diferencias de redondeo, conversión de unidades UMC/UMT o tipos de cambio entre sistemas.",
    color: "#ef4444",
    icon: "",
  },
  {
    id: "E6",
    name: "Suma de combinaciones (2 + 3 lotes)",
    desc: "Cuando el Layout suma más que cualquier secuencia individual del 551, evalúa si la combinación de 2 o 3 secuencias suma al total del grupo Layout (±2%). Detecta materiales importados en múltiples lotes dentro del mismo pedimento. Particiona las filas del Layout entre los lotes por cuota restante.",
    color: "#06b6d4",
    icon: "",
  },
  {
    id: "E7",
    name: "Precio unitario ($/pieza ±15%)",
    desc: "Usa el precio por unidad (ValorDolares / CantidadUMComercial) como discriminador con tolerancia ±15%. Resuelve casos donde los totales no coinciden pero el precio unitario confirma el material correcto (saldo acumulado de pedimentos anteriores, diferencias de conversión UMC/UMT).",
    color: "#f97316",
    icon: "",
  },
  {
    id: "E8",
    name: "Eliminación por descarte",
    desc: "Filtra candidatos del 551 ya usados y asigna el único remanente (o el más cercano en precio unitario ±30%). Válido cuando el material no tiene otra posible correspondencia.",
    color: "#8b5cf6",
    icon: "",
  },
  {
    id: "E9",
    name: "Mismo capítulo arancelario  corrige Fracción",
    desc: "Si la fracción del Layout no está en el 551 pero existe otra fracción del mismo capítulo (mismos 4 dígitos) en el mismo pedimento, la corrige usando el 551 como fuente oficial. Usa tolerancia de cantidades ±2 ud / ±5 USD o precio unitario ±20%. La fracción corregida aparece en rojo en el Excel.",
    color: "#ec4899",
    icon: "",
  },
  {
    id: "E10",
    name: "Solo Pedimento + precio unitario  corrige Fracción y País",
    desc: "Busca en todo el pedimento sin restricción de fracción ni país. Usa precio por pieza (±25%) como discriminador o asigna el único candidato disponible. Corrige tanto FraccionNico como PaisOrigen si difieren del 551. Las correcciones aparecen en rojo en el Excel.",
    color: "#14b8a6",
    icon: "",
  },
  {
    id: "E11",
    name: "Fuerza 1:1 greedy  emparejamiento por precio unitario",
    desc: "ltimo recurso total: empareja los grupos Layout pendientes con las secuencias 551 no usadas del mismo pedimento, ordenando ambas listas por precio unitario y asignando en pares (greedy). Aplica todas las correcciones necesarias de fracción y país. Sin límite de tolerancia.",
    color: "#64748b",
    icon: "",
  },
  {
    id: "R1",
    name: "Barrido inverso  precio unitario ±30% (cualquier sec. 551)",
    desc: "Primer barrido inverso: usa cualquier secuencia del 551 del mismo Pedimento+Fracción (sin importar si ya fue asignada antes). Criterio: precio por pieza ±30% o cantidades exactas ±3 ud / ±8 USD. Elimina el sesgo del filtro 'used551'.",
    color: "#0ea5e9",
    icon: "",
  },
  {
    id: "R2",
    name: "Barrido inverso  solo Pedimento, precio ±40%, corrige Fracción",
    desc: "Busca en todo el pedimento ignorando fracción y país. Precio unitario ±40% o único candidato disponible. Corrige FraccionNico y PaisOrigen si difieren del 551. Para los casos donde el Layout tiene fracción completamente diferente.",
    color: "#10b981",
    icon: "",
  },
  {
    id: "R3",
    name: "Fuerza total  greedy global sin tolerancia",
    desc: "Fuerza máxima: empareja TODOS los grupos Layout restantes con cualquier secuencia 551 del mismo pedimento, sin filtro ni tolerancia, eligiendo siempre la de precio unitario más cercano. Garantiza cobertura máxima aunque las correcciones sean necesarias.",
    color: "#f43f5e",
    icon: "",
  },
  {
    id: "R4",
    name: "Cobertura total por score (mismo pedimento)",
    desc: "Optimización de cierre: para cada grupo pendiente calcula un score de distancia contra todas las secuencias del mismo pedimento (cantidad, valor, precio unitario y similitud de descripción) y asigna la mejor. Diseñada para archivos grandes con variaciones de captura.",
    color: "#06b6d4",
    icon: "",
  },
  {
    id: "R5",
    name: "Cobertura total absoluta (global)",
    desc: "ltimo recurso extremo: si no existe candidato por pedimento, busca por fracción a nivel global y luego contra todo el 551. Prioriza la menor diferencia de cantidad/valor/precio para evitar filas sin secuencia.",
    color: "#f59e0b",
    icon: "",
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
      <div style={{ fontSize: 48, marginBottom: 16 }}></div>
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
  const colors = { E0: "#00d4aa", E1: "#22c55e", E2: "#3b82f6", E3: "#f59e0b", E4: "#a855f7", E5: "#ef4444",
                   E6: "#06b6d4", E7: "#f97316", E8: "#8b5cf6", E9: "#ec4899", E10: "#14b8a6", E11: "#64748b",
                   R1: "#0ea5e9", R2: "#10b981", R3: "#f43f5e", R4: "#06b6d4", R5: "#f59e0b" };
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

// 
// MDULO 2020  DS multi-pedimento (hoja DS*, Layout*)
// Lectura celda-por-celda para manejar hojas de >600 MB sin colapsar memoria.
// 

//  Helper compartido 
const nH2020 = (s) => String(s ?? "").trim().toLowerCase().replace(/[\s_\-]/g, "");

/**
 * Encuentra la hoja DS y la hoja Layout dentro del workbook.
 *
 * Estrategia en dos pasos:
 * 1. Detección por CONTENIDO para hojas que SÍ cargaron en memoria.
 *    Evita elegir tablas pivot/resumen ("td layout") cuando el Layout real
 *    se llama "2020" u otro nombre sin la palabra "layout".
 * 2. Fallback por NOMBRE ("layout" en el nombre) para hojas que NO cargaron
 *    (sheets demasiado grandes que xlsx no puede parsear completamente,
 *    ej. "Layout 2020" en archivos de 600 MB).
 */
function resolveDS2020SheetNames(wb) {
  const names = wb.SheetNames || [];

  //  DS: primera hoja con "DS" en el nombre 
  const dsName = names.find(n => n.toUpperCase().includes("DS"));

  //  Layout: paso 1  detección por contenido 
  const LAY_KNOWN = new Set([
    "pedimento","fraccionnico","seccalc","descripcion","paisorigen","pais_origen",
    "valormpdolares","cantidad_comercial","cantidadcomercial","notas","estado",
    "aduana_es","numero_parte","numeroparte","precio_unitario","valorme","fraccionmex",
  ]);

  let layName = null, bestHits = 0;
  for (const name of names) {
    if (name === dsName) continue;
    const ws = wb.Sheets[name];
    if (!ws) continue; // hoja no cargada  se intentará en paso 2
    try {
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "", sheetRows: 5 });
      for (const row of rows) {
        const hits = row.filter(c => LAY_KNOWN.has(nH2020(String(c ?? "")))).length;
        if (hits > bestHits) { bestHits = hits; layName = name; }
      }
    } catch (_) { /* skip hojas que lancen error */ }
  }

  //  Layout: paso 2  fallback por nombre (hojas grandes no cargadas) 
  // Si no se encontró por contenido, buscar hojas con "layout" en el nombre
  // que NO cargaron (típico de archivos con hojas >500 MB).
  if (!layName) {
    const fallback = names.find(n => n !== dsName && n.toLowerCase().includes("layout"));
    if (fallback) layName = fallback;
  }

  console.log("[resolve2020] dsName:", dsName, "| layName:", layName,
    layName ? `(hits: ${bestHits}, cargada: ${!!wb.Sheets[layName]})` : "(NO ENCONTRADA)");
  return { dsName, layName };
}

/**
 * Lee la hoja DS.
 * Soporta variantes de nombre de columna (ej. "Valor usd redondeado" en lugar de "ValorDolares").
 */
function readDS2020Sheet(sheet) {
  if (!sheet) return [];

  // Aliases internos  posibles nombres en el Excel (normalizados con nH2020)
  const DS_COL_MAP = {
    Pedimento2:           ["Pedimento2"],
    Fraccion:             ["Fraccion"],
    SubdivisionFraccion:  ["SubdivisionFraccion","Subdivision Fraccion","NICO"],
    SecuenciaFraccion:    ["SecuenciaFraccion"],
    DescripcionMercancia: ["DescripcionMercancia"],
    CantidadUMComercial:  ["CantidadUMComercial"],
    ValorDolares:         ["ValorDolares","Valor usd redondeado","ValorAduana","Valor Aduana Estadístico","ValorAgregado"],
    PaisOrigenDestino:    ["PaisOrigenDestino"],
    "Candado 551":        ["Candado 551","Candado DS 551","Candado","Clave","Secuencias"],
  };
  const knownNorms = new Set(Object.values(DS_COL_MAP).flat().map(nH2020));

  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
  if (!rows.length) return [];

  // Detectar fila header: la que tenga más columnas conocidas
  let hdrI = 0, bestHits = 0;
  for (let i = 0; i < Math.min(rows.length, 5); i++) {
    const hits = rows[i].filter(c => knownNorms.has(nH2020(String(c ?? "")))).length;
    if (hits > bestHits) { bestHits = hits; hdrI = i; }
    if (hits >= 3) break;
  }
  const hdr = rows[hdrI].map(c => String(c ?? "").trim());
  console.log("[DS2020] hdrI:", hdrI, "hits:", bestHits, "headers:", hdr.slice(0, 16));

  // Mapear: nombre interno  primer índice que coincida, respetando ORDEN de aliases
  const idx = {};
  for (const [internalName, aliases] of Object.entries(DS_COL_MAP)) {
    let found = -1;
    for (const alias of aliases) {
      const n = nH2020(alias);
      const i = hdr.findIndex(h => nH2020(h) === n);
      if (i >= 0) { found = i; break; }
    }
    if (found >= 0) idx[internalName] = found;
  }
  console.log("[DS2020] colIdx:", JSON.stringify(idx));

  // Columna REVISADO en el DS (para escribir motivos de no-match)
  const revisadoColIdx = hdr.reduce((l, h, i) => nH2020(h) === "revisado" ? i : l, -1);

  const out = [];
  for (let i = hdrI + 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.every(c => c === "" || c == null)) continue;
    const obj = { _dsIdx: out.length, _rowI: i }; // _rowI = índice 0-based en el sheet (para encode_cell)
    for (const [col, ci] of Object.entries(idx)) obj[col] = row[ci] ?? "";
    out.push(obj);
  }
  out._hdrI         = hdrI;
  out._revisadoCol  = revisadoColIdx;
  // Columnas no encontradas (para avisar al usuario cómo deben nombrarse)
  out._missingColumns = Object.entries(DS_COL_MAP)
    .filter(([k]) => idx[k] === undefined)
    .map(([name, aliases]) => ({ name, aceptados: aliases }));
  console.log("[DS2020] DS rows:", out.length, "REVISADO col:", revisadoColIdx,
    "missing:", out._missingColumns?.length, "muestra:", out.slice(0,2).map(r=>({ped:r.Pedimento2,frac:r.Fraccion,sec:r.SecuenciaFraccion,candado:r["Candado 551"]})));
  return out;
}

/**
 * Lee la hoja Layout 2020 celda-por-celda (solo las columnas necesarias).
 * Usa sheet_to_json solo para detectar encabezados (resuelve shared strings),
 * luego lee dato a dato para eficiencia con hojas grandes.
 */
function readLayout2020Sheet(sheet) {
  if (!sheet || !sheet["!ref"]) {
    console.log("[Layout2020] ERROR: sheet undefined o sin !ref");
    return { layoutRows: [], headerRowIdx: 0, colIdx: {} };
  }

  const range = XLSX.utils.decode_range(sheet["!ref"]);
  console.log("[Layout2020] ref:", sheet["!ref"], "filas totales:", range.e.r + 1, "cols:", range.e.c + 1);

  // Asegurar que leemos suficientes columnas para encontrar candado (ej. AJ = col 35); si !ref es corto, extendemos
  // Limitar columnas leídas para el header: si el sheet es más ancho que 400 cols
  // (ej. A1:XEM5065), la mayoría de columnas útiles están en las primeras ~300.
  // Escaneamos en bloques de 300 cols hasta encontrar los hits suficientes (4).
  const MIN_COLS_CANDADO = 60;
  const MAX_HDR_COLS     = 400; // límite razonable para encontrar todos los headers
  const maxCol = Math.min(Math.max(range.e.c, MIN_COLS_CANDADO - 1), MAX_HDR_COLS - 1);

  //  1. Primeras 15 filas con sheet_to_json (resuelve shared strings) 
  const hdrRange = { s: { r: 0, c: range.s.c }, e: { r: Math.min(14, range.e.r), c: maxCol } };
  const sampleRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", range: hdrRange });
  console.log("[Layout2020] sampleRows leídas:", sampleRows.length, "cols hasta:", maxCol + 1);

  //  2. Detectar fila de encabezado 
  const KNOWN = new Set(["pedimento","fraccionnico","seccalc","descripcion",
                         "paisorigen","valormpdolares","cantidadcomercial",
                         "cantidad_comercial","notas","estado","candado","clave","nico"]);
  let hdrI = 0, bestHits = 0;
  for (let i = 0; i < sampleRows.length; i++) {
    const hits = sampleRows[i].filter(c => KNOWN.has(nH2020(String(c ?? "")))).length;
    console.log(`[Layout2020] fila[${i}] hits:${hits}`, sampleRows[i].slice(0,6).map(v=>String(v).slice(0,12)));
    if (hits > bestHits) { bestHits = hits; hdrI = i; }
    if (hits >= 4) break;
  }
  console.log("[Layout2020] hdrI:", hdrI, "bestHits:", bestHits);

  //  3. Mapear columnas 
  const rawHeaders = (sampleRows[hdrI] || []).map(c => String(c ?? "").trim());
  console.log("[Layout2020] headers:", rawHeaders);

  // findFirst: prueba cada alias EN ORDEN y devuelve la primera columna que coincida
  // Esto garantiza que el alias más específico (primero en la lista) gana.
  const findFirst = (...names) => {
    for (const name of names) {
      const n = nH2020(name);
      const idx = rawHeaders.findIndex(h => nH2020(h) === n);
      if (idx >= 0) return idx;
    }
    return -1;
  };
  // findLast: última columna que coincida (para duplicados donde queremos el último)
  const findLast = (...names) => {
    const ts = names.map(nH2020);
    return rawHeaders.reduce((last, h, i) => ts.includes(nH2020(h)) ? i : last, -1);
  };

  const colIdx = {
    // findLast para columnas que pueden estar duplicadas y la versión útil es la LTIMA
    // (ej. "pedimento" en col 3 es "20-400-3459-..." pero col 174 es "400-3459-..." = formato DS)
    pedimento:  findLast("pedimento"),
    frac:       findLast("FraccionNico","fraccionnico"),
    // findFirst para cantidad: cuando hay varias "cantidad_comercial" la primera suele ser la columna principal de datos (la que cuadra con DS); la última puede ser un duplicado o subtotal
    cant:       findFirst("cantidad_comercial","cantidadcomercial","cantidadumc"),
    notas:      findLast("NOTAS","notas"),    // última NOTAS = columna de salida
    // findFirst para columnas sin ambigüedad importante
    desc:       findFirst("descripcion","clase_descripcion","descripcionmercancia"),
    pais:       findFirst("pais_origen","paisorigen","paisorigendestino"),
    val:        findFirst("ValorMPDolares","valormpdolares","valordolares","valor_me","valorme"),
    nico:       findFirst("NICO","nico","SubdivisionFraccion","subdivisionfraccion"),
    sec:        findFirst("SEC CALC","seccalc","secuenciaped","Secuencia","secuencia","SEC"),
    notasIn:    findFirst("NOTAS","notas"),   // primera NOTAS = flags de entrada ("NO INCLUIR")
    estado:     findFirst("ESTADO","estado"),
    // Candado: priorizar "CANDADO 551" / "Candado 551" (columna que cuadra con DS); si no existe, usar findLast para evitar columna "Clave" con valor único.
    candado:    (() => { const exact = findFirst("CANDADO 551","Candado 551","Candado DS 551"); return exact >= 0 ? exact : findLast("Candado","Clave","candado"); })(),
  };
  // Columnas esperadas y nombres aceptados (para avisar si no se encuentran)
  const LAYOUT_COL_SPEC = [
    { key: "pedimento", label: "Pedimento", aceptados: ["Pedimento", "pedimento"] },
    { key: "frac", label: "FraccionNico", aceptados: ["FraccionNico", "fraccionnico"] },
    { key: "sec", label: "SEC CALC / Secuencia", aceptados: ["SEC CALC", "seccalc", "secuenciaped", "Secuencia", "secuencia", "SEC"] },
    { key: "desc", label: "Descripcion", aceptados: ["descripcion", "clase_descripcion", "descripcionmercancia"] },
    { key: "pais", label: "PaisOrigen", aceptados: ["pais_origen", "paisorigen", "paisorigendestino"] },
    { key: "cant", label: "Cantidad", aceptados: ["cantidad_comercial", "cantidadcomercial", "cantidadumc"] },
    { key: "val", label: "Valor USD", aceptados: ["ValorMPDolares", "valormpdolares", "valordolares", "valor_me", "valorme"] },
    { key: "notas", label: "NOTAS", aceptados: ["NOTAS", "notas"] },
    // Candado 551 no se exige: el sistema lo usa si existe, pero no se avisa como faltante (se puede crear/calcular)
  ];
  const layoutMissing = LAYOUT_COL_SPEC.filter(({ key }) => colIdx[key] < 0).map(({ label, aceptados }) => ({ name: label, aceptados }));
  console.log("[Layout2020] colIdx:", JSON.stringify(colIdx), "missing:", layoutMissing.length);

  //  4. Helpers lectura celda 
  const cellVal = (r, c) => {
    if (c < 0) return "";
    const cell = sheet[XLSX.utils.encode_cell({ r, c })];
    if (!cell) return "";
    const v = cell.v ?? cell.w ?? "";
    return String(v).trim();
  };
  const cellNum = (r, c) => {
    if (c < 0) return 0;
    const cell = sheet[XLSX.utils.encode_cell({ r, c })];
    return cell ? (parseFloat(cell.v) || 0) : 0;
  };

  //  5. Leer filas de datos celda-por-celda (saltar filas sin pedimento ni fracción) 
  const layoutRows = [];
  for (let r = hdrI + 1; r <= range.e.r; r++) {
    const pedVal  = cellVal(r, colIdx.pedimento);
    const fracVal = cellVal(r, colIdx.frac);

    // Saltar filas completamente vacías (sin pedimento ni fracción)
    if (!pedVal && !fracVal) continue;

    const notasInVal = cellVal(r, colIdx.notasIn).toUpperCase();
    const noIncluir  = notasInVal.includes("NO INCLUIR");

    layoutRows.push({
      _idx:        layoutRows.length,
      _rowI:       r,
      Pedimento:   pedVal,
      FraccionNico:fracVal,
      NICO:        colIdx.nico >= 0 ? cellVal(r, colIdx.nico) : "",
      Descripcion: cellVal(r, colIdx.desc),
      PaisOrigen:  cellVal(r, colIdx.pais),
      Cantidad:    cellNum(r, colIdx.cant),
      ValorUSD:    cellNum(r, colIdx.val),
      SecCalc:     cellVal(r, colIdx.sec),
      Notas:       cellVal(r, colIdx.notas),
      Estado:      cellVal(r, colIdx.estado),
      Candado:     colIdx.candado >= 0 ? cellVal(r, colIdx.candado) : "",
      noIncluir,
    });
  }
  console.log("[Layout2020] layoutRows:", layoutRows.length,
    "muestra:", layoutRows.slice(0,3).map(r=>({ped:r.Pedimento,frac:r.FraccionNico,sec:r.SecCalc,noInc:r.noIncluir})));
  return { layoutRows, headerRowIdx: hdrI, colIdx, _missingColumns: layoutMissing };
}

/** Cascade 2020: verifica secuencias existentes y asigna las que faltan.
 *
 *  PRINCIPIO (como consultor de comercio exterior):
 *  Cada secuencia del DS tiene Fraccion + Descripcion + PaisOrigenDestino +
 *  CantidadUMComercial + ValorDolares. En el Layout, las filas con la misma
 *  fracción y descripción (normalizada) forman el grupo que debe coincidir con
 *  esa secuencia DS, SIN importar si el país difiere  el DS puede registrar
 *  país MEX mientras el Layout tiene MEX+CHN+USA+TWN (todos entran con ese país
 *  de la sec, no se corrige el país).
 *
 *  Tolerancia permitida: ±1 unidad en Cantidad, ±4 USD en Valor.
 *  Si los totales globales DS = Layout cuadran, todas las filas deben asignarse.
 */
function runCascade2020(layoutRows, dsRows) {
  const nFrac = (v) => String(v ?? "").trim().replace(/^0+/, "") || "0";
  const normStr = (v) => String(v ?? "").trim();
  // NICO normalizado (2 dígitos). Si no viene explícito, intenta derivarlo de FraccionNico (10 dígitos).
  const nNico = (nicoVal, fracNicoVal = "") => {
    let s = normStr(nicoVal);
    if (!s || s === ".") {
      const digits = normStr(fracNicoVal).replace(/\D/g, "");
      if (digits.length >= 10) s = digits.slice(-2);
    }
    if (!s || s === ".") return "00";
    const d = s.replace(/\D/g, "");
    if (d) return d.slice(-2).padStart(2, "0");
    const n = parseInt(s, 10);
    if (Number.isFinite(n)) return String(n).padStart(2, "0");
    return s.toUpperCase();
  };
  // Normaliza descripción: lowercase, collapse whitespace, quita paréntesis redundantes
  // nDesc: normaliza descripción  espacios, "/" sin espacios alrededor, paréntesis
  const nDesc = (s) => String(s ?? "").trim().toLowerCase()
    .replace(/\s*\/\s*/g, "/")   // "hoja / divisores"  "hoja/divisores"
    .replace(/\s*\(\s*/g, "(")   // "resistencia ( las"  "resistencia (las"
    .replace(/\s*\)\s*/g, ")")
    .replace(/\s+/g, " ");

  const dsTotalC = dsRows.reduce((a,r)=>a+(parseFloat(r["CantidadUMComercial"])||0),0);
  const dsTotalV = dsRows.reduce((a,r)=>a+(parseFloat(r["ValorDolares"])||0),0);
  const lyTotalC = layoutRows.filter(r=>!r.noIncluir).reduce((a,r)=>a+r.Cantidad,0);
  const lyTotalV = layoutRows.filter(r=>!r.noIncluir).reduce((a,r)=>a+r.ValorUSD,0);
  const diffC = Math.abs(lyTotalC - dsTotalC);
  const diffV = Math.abs(lyTotalV - dsTotalV);
  // globalCuadra se basa SOLO en cantidad  si cant global coincide, todo layout debe asignarse
  // El valor USD puede tener diferencias pequeñas por redondeos entre el DS y el Layout
  const globalCantCuadra = diffC <= 1;
  const globalValCuadra  = diffV <= 5;
  const globalCuadra     = globalCantCuadra && globalValCuadra; // para lógica de asignación, solo cuenta cantidad
  const STRICT_DS_RULES = true; // En DS: sin forzados fuera de tolerancia
  if (diffC > 1) {
    console.warn("[Cascade2020] Cantidades globales NO coinciden: Layout=", lyTotalC, "DS=", dsTotalC, "diff=", diffC);
  }
  // Una secuencia es "real" si es un número válido (no ".", no vacío)
  const isRealSec = (v) => {
    const s = normStr(v);
    return s !== "" && s !== "." && !isNaN(parseFloat(s));
  };

  // getDSFrac: obtiene la fracción del DS row.
  // Algunos archivos tienen la columna Fraccion vacía pero la fracción está
  // embebida en el Candado 551 con formato "Ped-Frac-Sec".
  // Fallback: extraerla del candado cuando Fraccion está vacío.
  const getDSFrac = (dsRow) => {
    const fracDirect = nFrac(normStr(dsRow["Fraccion"]));
    if (fracDirect) return fracDirect;
    // Intentar extraer del candado: formato "PED-FRAC-SEC"
    const candado = normStr(dsRow["Candado 551"]);
    if (candado) {
      // Separar por "-" y reconstruir: primer token = aduana, segundo = num, tercero = año, cuarto = frac, quinto = sec
      // Formato real: "400-3459-0002178-85322999-104"  partes [400, 3459, 0002178, 85322999, 104]
      const parts = candado.split("-");
      // La fracción siempre es la parte de 8 dígitos (índice 3 en el formato estándar)
      for (let i = 0; i < parts.length; i++) {
        const p = parts[i].replace(/\./g, "");
        if (p.length === 8 && /^\d+$/.test(p)) return p;
      }
    }
    return "";
  };

  //  Lookups del DS 
  const dsByCandado  = new Map(); // "Candado 551"  dsRow
  const dsByPFP      = new Map(); // Pedimento2|||Fraccion|||Pais  [dsRow]
  const dsByPF       = new Map(); // Pedimento2|||Fraccion  [dsRow]
  const dsByPFN      = new Map(); // Pedimento2|||Fraccion|||NICO  [dsRow]
  const dsByPFDesc   = new Map(); // Pedimento2|||Fraccion|||DescNorm  [dsRow]
  const usedDS       = new Set();

  // Cadena de prioridad: Pedimento2-Fraccion-DescripcionMercancia-PaisOrigen-CantidadUMComercial-ValorUSDredondeado (para match explícito y mensaje "Coincide por cadena")
  const dsByCadenaPrefix = new Map(); // ped|||frac|||desc|||pais  [dsRow]
  for (const r of dsRows) {
    const candado = normStr(r["Candado 551"]);
    if (candado) dsByCandado.set(candado, r);

    const ped2 = normStr(r["Pedimento2"]);
    const frac = getDSFrac(r);
    const nico = nNico(r["SubdivisionFraccion"], r["Fraccion"]);
    const pais = normStr(r["PaisOrigenDestino"]);
    const desc = nDesc(r["DescripcionMercancia"]);

    const kCadena = `${ped2}|||${frac}|||${desc}|||${pais}`;
    if (!dsByCadenaPrefix.has(kCadena)) dsByCadenaPrefix.set(kCadena, []);
    dsByCadenaPrefix.get(kCadena).push(r);

    const kPFP = `${ped2}|||${frac}|||${pais}`;
    if (!dsByPFP.has(kPFP)) dsByPFP.set(kPFP, []);
    dsByPFP.get(kPFP).push(r);

    const kPF = `${ped2}|||${frac}`;
    if (!dsByPF.has(kPF)) dsByPF.set(kPF, []);
    dsByPF.get(kPF).push(r);

    const kPFN = `${ped2}|||${frac}|||${nico}`;
    if (!dsByPFN.has(kPFN)) dsByPFN.set(kPFN, []);
    dsByPFN.get(kPFN).push(r);

    const kPFD = `${ped2}|||${frac}|||${desc}`;
    if (!dsByPFDesc.has(kPFD)) dsByPFDesc.set(kPFD, []);
    dsByPFDesc.get(kPFD).push(r);
  }

  // Match por cantidades: tolerancia estricta ±1 ud / ±4 USD (comercio exterior)
  const tryMatchQty = (cands, sumCant, sumVal, tolC = 1, tolV = 4) => {
    for (const r of cands) {
      if (usedDS.has(r._dsIdx)) continue;
      const c = parseFloat(r["CantidadUMComercial"]) || 0;
      const v = parseFloat(r["ValorDolares"])        || 0;
      if (Math.abs(sumCant - c) <= tolC && Math.abs(sumVal - v) <= tolV) return r;
    }
    return null;
  };

  const tryMatchUP = (cands, sumCant, sumVal, tolPct = 0.15) => {
    if (sumCant <= 0) return null;
    const up = sumVal / sumCant;
    let best = null, bestDiff = Infinity;
    for (const r of cands) {
      if (usedDS.has(r._dsIdx)) continue;
      const c = parseFloat(r["CantidadUMComercial"]) || 0;
      const v = parseFloat(r["ValorDolares"])        || 0;
      if (c <= 0) continue;
      const diff = Math.abs((v / c) - up) / Math.max(up, 0.001);
      if (diff <= tolPct && diff < bestDiff) { bestDiff = diff; best = r; }
    }
    return best;
  };

  // findSubset: busca un subconjunto de filas que sume exactamente (dsCant, dsVal)
  // Estrategia: greedy rápido primero, luego backtracking con poda de sufijo.
  const findSubset = (rows, dsCant, dsVal, tolC = 1, tolV = 4) => {
    if (rows.length === 0) return null;

    // 1. El grupo completo coincide
    const totalC = rows.reduce((a,r)=>a+r.Cantidad,0);
    if (Math.abs(totalC - dsCant) <= tolC) {
      const totalV = rows.reduce((a,r)=>a+r.ValorUSD,0);
      if (Math.abs(totalV - dsVal) <= tolV) return rows;
    }
    // Si la suma total es menor que el objetivo, imposible
    if (totalC < dsCant - tolC) return null;

    // 2. Greedy consecutivo (rápido para casos simples)
    const sorted = [...rows].sort((a,b) => a.Cantidad - b.Cantidad);
    for (let start = 0; start < Math.min(sorted.length, 60); start++) {
      let sumC = 0, sumV = 0, subset = [];
      for (let j = start; j < sorted.length; j++) {
        if (sumC + sorted[j].Cantidad > dsCant + tolC) break;
        subset.push(sorted[j]);
        sumC += sorted[j].Cantidad;
        sumV += sorted[j].ValorUSD;
        if (Math.abs(sumC - dsCant) <= tolC && Math.abs(sumV - dsVal) <= tolV) return subset;
      }
    }
    const sortedDesc = [...rows].sort((a,b) => b.Cantidad - a.Cantidad);
    for (let start = 0; start < Math.min(sortedDesc.length, 60); start++) {
      let sumC = 0, sumV = 0, subset = [];
      for (let j = start; j < sortedDesc.length; j++) {
        if (sumC + sortedDesc[j].Cantidad > dsCant + tolC && subset.length > 0) break;
        subset.push(sortedDesc[j]);
        sumC += sortedDesc[j].Cantidad;
        sumV += sortedDesc[j].ValorUSD;
        if (Math.abs(sumC - dsCant) <= tolC && Math.abs(sumV - dsVal) <= tolV) return subset;
      }
    }

    // 3. Backtracking con poda de sufijo (branch & bound).
    //    sortedDesc ya calculado arriba.
    //    Precalcular sufijos: suffixC[i] = suma Cantidad desde i hasta el final.
    //    Si sumC + suffixC[i] < dsCant - tolC  imposible llegar  podar.
    const suffixC = new Array(sortedDesc.length + 1).fill(0);
    for (let i = sortedDesc.length - 1; i >= 0; i--) {
      suffixC[i] = suffixC[i + 1] + sortedDesc[i].Cantidad;
    }

    let btResult = null;
    let btNodes  = 0;
    const MAX_BT = 500000;

    const bt = (idx, sumC, sumV, current) => {
      if (btResult || btNodes > MAX_BT) return;
      btNodes++;
      if (Math.abs(sumC - dsCant) <= tolC && Math.abs(sumV - dsVal) <= tolV && current.length > 0) {
        btResult = [...current];
        return;
      }
      if (idx >= sortedDesc.length) return;
      if (sumC > dsCant + tolC) return;                       // poda superior
      if (sumC + suffixC[idx] < dsCant - tolC) return;       // poda inferior (sufijo)

      const r = sortedDesc[idx];
      if (sumC + r.Cantidad <= dsCant + tolC) {
        current.push(r);
        bt(idx + 1, sumC + r.Cantidad, sumV + r.ValorUSD, current);
        current.pop();
      }
      if (!btResult) bt(idx + 1, sumC, sumV, current);
    };

    bt(0, 0, 0, []);
    return btResult;
  };

  // findSubsetCantOnly: igual que findSubset pero ignora valor USD (para fase B2).
  // Mucho más rápido: solo necesita que la cantidad cuadre ±1.
  const findSubsetCantOnly = (rows, dsCant, tolC = 1) => {
    if (rows.length === 0) return null;
    const totalC = rows.reduce((a,r)=>a+r.Cantidad,0);
    if (Math.abs(totalC - dsCant) <= tolC) return rows;
    if (totalC < dsCant - tolC) return null;

    // Greedy ascendente
    const sorted = [...rows].sort((a,b) => a.Cantidad - b.Cantidad);
    for (let start = 0; start < Math.min(sorted.length, 60); start++) {
      let sumC = 0, subset = [];
      for (let j = start; j < sorted.length; j++) {
        if (sumC + sorted[j].Cantidad > dsCant + tolC) break;
        subset.push(sorted[j]);
        sumC += sorted[j].Cantidad;
        if (Math.abs(sumC - dsCant) <= tolC) return subset;
      }
    }
    // Greedy descendente
    const sortedDesc = [...rows].sort((a,b) => b.Cantidad - a.Cantidad);
    for (let start = 0; start < Math.min(sortedDesc.length, 60); start++) {
      let sumC = 0, subset = [];
      for (let j = start; j < sortedDesc.length; j++) {
        if (sumC + sortedDesc[j].Cantidad > dsCant + tolC && subset.length > 0) break;
        subset.push(sortedDesc[j]);
        sumC += sortedDesc[j].Cantidad;
        if (Math.abs(sumC - dsCant) <= tolC) return subset;
      }
    }
    // Backtracking con poda de sufijo
    const suffixC = new Array(sortedDesc.length + 1).fill(0);
    for (let i = sortedDesc.length - 1; i >= 0; i--) suffixC[i] = suffixC[i+1] + sortedDesc[i].Cantidad;
    let result = null, nodes = 0;
    const bt2 = (idx, sumC, cur) => {
      if (result || nodes++ > 300000) return;
      if (Math.abs(sumC - dsCant) <= tolC && cur.length > 0) { result = [...cur]; return; }
      if (idx >= sortedDesc.length) return;
      if (sumC > dsCant + tolC) return;
      if (sumC + suffixC[idx] < dsCant - tolC) return;
      const r = sortedDesc[idx];
      if (sumC + r.Cantidad <= dsCant + tolC) { cur.push(r); bt2(idx+1, sumC+r.Cantidad, cur); cur.pop(); }
      if (!result) bt2(idx+1, sumC, cur);
    };
    bt2(0, 0, []);
    return result;
  };

  // Suma de 2 o 3 secuencias DS que coincida con total (tolerancia ±1 cant, ±4 val)
  const tryMatchCombo = (cands, sumCant, sumVal, tolC = 1, tolV = 4) => {
    const pool = cands.filter(r => !usedDS.has(r._dsIdx)).slice(0, 12);
    for (let i = 0; i < pool.length - 1; i++) {
      for (let j = i + 1; j < pool.length; j++) {
        const c = (parseFloat(pool[i]["CantidadUMComercial"]) || 0) + (parseFloat(pool[j]["CantidadUMComercial"]) || 0);
        const v = (parseFloat(pool[i]["ValorDolares"]) || 0) + (parseFloat(pool[j]["ValorDolares"]) || 0);
        if (Math.abs(c - sumCant) <= tolC && Math.abs(v - sumVal) <= tolV) return [pool[i], pool[j]];
      }
    }
    for (let i = 0; i < pool.length - 2; i++) {
      for (let j = i + 1; j < pool.length - 1; j++) {
        for (let k = j + 1; k < pool.length; k++) {
          const c = (parseFloat(pool[i]["CantidadUMComercial"]) || 0) + (parseFloat(pool[j]["CantidadUMComercial"]) || 0) + (parseFloat(pool[k]["CantidadUMComercial"]) || 0);
          const v = (parseFloat(pool[i]["ValorDolares"]) || 0) + (parseFloat(pool[j]["ValorDolares"]) || 0) + (parseFloat(pool[k]["ValorDolares"]) || 0);
          if (Math.abs(c - sumCant) <= tolC && Math.abs(v - sumVal) <= tolV) return [pool[i], pool[j], pool[k]];
        }
      }
    }
    return null;
  };

  const pickSubset = (pool, dsCant, dsVal) => {
    const exact = findSubset(pool, dsCant, dsVal, 1, 4);
    if (exact) return exact;
    if (!STRICT_DS_RULES && globalCantCuadra) return findSubsetCantOnly(pool, dsCant);
    return null;
  };

  // assignment: _idx  { status, newSec, dsRow, corrections[], reason }
  // status: "ok"|"corrected"|"new"|"unmatched"
  const assignment = new Map();

  // Helper: asignar filas a una secuencia DS (usado por Cadena, E0candado y fases siguientes)
  const assignRows = (rows, dsRow, estrategia, fracCorr = null) => {
    const newSec = normStr(dsRow["SecuenciaFraccion"]);
    const sumC = rows.reduce((a,r)=>a+r.Cantidad,0);
    const sumV = rows.reduce((a,r)=>a+r.ValorUSD,0);
    const dsCant = parseFloat(dsRow["CantidadUMComercial"])||0;
    const dsVal  = parseFloat(dsRow["ValorDolares"])||0;
    usedDS.add(dsRow._dsIdx);
    const CADENA_FIELDS = "Pedimento2-Fraccion-DescripcionMercancia-PaisOrigen-CantidadUMComercial-ValorUSDredondeado";
    for (const row of rows) {
      const fracOriginal = nFrac(row.FraccionNico);
      const dsFrac       = getDSFrac(dsRow);
      const isCrossFrac  = fracCorr !== null || fracOriginal !== dsFrac;
      const secOriginal  = isRealSec(row.SecCalc) ? normStr(row.SecCalc) : null;
      const isCorrected = secOriginal !== null && secOriginal !== newSec;
      let reason = estrategia === "Cadena"
        ? `Coincide por cadena (${CADENA_FIELDS}). [Cadena] Sec=${newSec}  Layout Cant=${sumC.toLocaleString()} Val=$${sumV.toFixed(0)} | DS Cant=${dsCant.toLocaleString()} Val=$${dsVal.toFixed(0)}`
        : `[${estrategia}] Sec=${newSec}`;
      if (!(estrategia === "Cadena") && isCorrected)  reason += ` (corregido de ${secOriginal})`;
      if (!(estrategia === "Cadena") && isCrossFrac)  reason += ` [Frac ${fracOriginal}${fracCorr||dsFrac}]`;
      if (estrategia !== "Cadena") reason += `  Layout Cant=${sumC.toLocaleString()} Val=$${sumV.toFixed(0)} | DS Cant=${dsCant.toLocaleString()} Val=$${dsVal.toFixed(0)}`;
      if (estrategia !== "Cadena") reason = `Alternativa: ${reason}`;
      assignment.set(row._idx, {
        status: secOriginal !== null ? "corrected" : "new",
        newSec, dsRow, corrections: [], estrategia,
        secOrig: isCorrected ? secOriginal : null,
        fracCorr: isCrossFrac ? (fracCorr || dsFrac) : null,
        fracOrig: isCrossFrac ? fracOriginal : null,
        reason,
      });
    }
  };

  //  Prioridad por NICO: Pedimento2+Fraccion+NICO (grupos CE) 
  // Consultoría CE: primero cerrar por grupos NICO, luego bajar a reglas más amplias.
  const groupsByNico = new Map(); // ped|||frac|||nico  rows
  for (const row of layoutRows) {
    if (row.noIncluir) continue;
    const ped  = normStr(row.Pedimento);
    const frac = nFrac(row.FraccionNico);
    const nico = nNico(row.NICO, row.FraccionNico);
    const kN   = `${ped}|||${frac}|||${nico}`;
    if (!groupsByNico.has(kN)) groupsByNico.set(kN, []);
    groupsByNico.get(kN).push(row);
  }

  for (const [kN, rows] of groupsByNico) {
    const pend = rows.filter(r => !assignment.has(r._idx));
    if (!pend.length) continue;

    const dsListBase = (dsByPFN.get(kN) || []).filter(ds => !usedDS.has(ds._dsIdx));
    if (!dsListBase.length) continue;

    // Caso simple: solo una secuencia DS para este NICO
    if (dsListBase.length === 1) {
      const dsRow = dsListBase[0];
      const dsCant = parseFloat(dsRow["CantidadUMComercial"]) || 0;
      const dsVal  = parseFloat(dsRow["ValorDolares"]) || 0;
      const dsSinValor = dsVal === 0 || !Number.isFinite(dsVal);
      const sumC = pend.reduce((a, r) => a + r.Cantidad, 0);
      const sumV = pend.reduce((a, r) => a + r.ValorUSD, 0);
      const okC = Math.abs(sumC - dsCant) <= 1 || (dsCant > 1000 && Math.abs(sumC - dsCant) / dsCant <= 0.005);
      const okV = Math.abs(sumV - dsVal) <= 5 || dsSinValor;
      if (okC && okV) {
        assignRows(pend, dsRow, "NICO1");
      } else {
        const sub = pickSubset(pend, dsCant, dsVal);
        if (sub) assignRows(sub, dsRow, "NICO1_sub");
      }
      continue;
    }

    // Caso múltiple: repartir filas del mismo NICO entre varias secuencias DS del NICO
    const dsList = [...dsListBase].sort((a, b) => (parseFloat(a["CantidadUMComercial"]) || 0) - (parseFloat(b["CantidadUMComercial"]) || 0));
    let remaining = [...pend];

    for (let i = 0; i < dsList.length; i++) {
      const dsRow = dsList[i];
      if (usedDS.has(dsRow._dsIdx) || !remaining.length) continue;

      const dsCant = parseFloat(dsRow["CantidadUMComercial"]) || 0;
      const dsVal  = parseFloat(dsRow["ValorDolares"]) || 0;
      const dsDesc = nDesc(dsRow["DescripcionMercancia"]);
      const dsPais = normStr(dsRow["PaisOrigenDestino"]);
      const isLast = i === dsList.length - 1;

      let sub = null;
      const byDP = remaining.filter(r => nDesc(r.Descripcion) === dsDesc && normStr(r.PaisOrigen) === dsPais);
      if (byDP.length) sub = pickSubset(byDP, dsCant, dsVal);

      if (!sub) {
        const byD = remaining.filter(r => nDesc(r.Descripcion) === dsDesc);
        if (byD.length) sub = pickSubset(byD, dsCant, dsVal);
      }

      if (!sub && !isLast) {
        sub = pickSubset(remaining, dsCant, dsVal);
      }

      if (!sub && isLast) {
        const remC = remaining.reduce((a, r) => a + r.Cantidad, 0);
        const remV = remaining.reduce((a, r) => a + r.ValorUSD, 0);
        const dsSinValor = dsVal === 0 || !Number.isFinite(dsVal);
        const okC = Math.abs(remC - dsCant) <= 1 || (dsCant > 1000 && Math.abs(remC - dsCant) / dsCant <= 0.005);
        const okV = Math.abs(remV - dsVal) <= 5 || dsSinValor;
        if (okC && okV) sub = remaining;
        if (!sub) sub = pickSubset(remaining, dsCant, dsVal);
      }

      if (sub && sub.length) {
        assignRows(sub, dsRow, "NICOm");
        const usedSet = new Set(sub.map(r => r._idx));
        remaining = remaining.filter(r => !usedSet.has(r._idx));
      }
    }
  }

  //  Prioridad por cadena: Pedimento2-Fraccion-DescripcionMercancia-PaisOrigen-CantidadUMComercial-ValorUSDredondeado 
  // Si los campos en Layout y DS coinciden (misma cadena + totales dentro de tolerancia), asignar y reportar "Coincide por cadena".
  // Agrupar Layout por (Pedimento, FraccionNico, Descripcion norm, PaisOrigen); buscar en DS por la misma cadena y totales ±1 cant / ±5 USD.
  const groupsByCadena = new Map();
  for (const row of layoutRows) {
    if (row.noIncluir) continue;
    const ped  = normStr(row.Pedimento);
    const frac = nFrac(row.FraccionNico);
    const desc = nDesc(row.Descripcion ?? "");
    const pais = normStr(row.PaisOrigen ?? "");
    const k    = `${ped}|||${frac}|||${desc}|||${pais}`;
    if (!groupsByCadena.has(k)) groupsByCadena.set(k, []);
    groupsByCadena.get(k).push(row);
  }
  const tolCadenaC = 1, tolCadenaV = 5;
  const tolCadenaPct = 0.005;
  for (const [kCadena, rows] of groupsByCadena) {
    if (rows.some(r => assignment.has(r._idx))) continue;
    const sumC = rows.reduce((a, r) => a + r.Cantidad, 0);
    const sumV = rows.reduce((a, r) => a + r.ValorUSD, 0);
    const cands = dsByCadenaPrefix.get(kCadena) || [];
    let dsRow = null;
    for (const r of cands) {
      if (usedDS.has(r._dsIdx)) continue;
      const dsCant = parseFloat(r["CantidadUMComercial"]) || 0;
      const dsVal  = parseFloat(r["ValorDolares"]) || 0;
      const dsSinValor = dsVal === 0 || !Number.isFinite(dsVal);
      const okC = Math.abs(sumC - dsCant) <= tolCadenaC || (dsCant > 1000 && Math.abs(sumC - dsCant) / dsCant <= tolCadenaPct);
      const okV = Math.abs(sumV - dsVal) <= tolCadenaV || dsSinValor;
      if (okC && okV) { dsRow = r; break; }
    }
    if (dsRow) assignRows(rows, dsRow, "Cadena");
  }

  //  E0candado: Asignar por candado (Layout col AJ = DS col candado) 
  // Solo si el total del grupo Layout (cantidad + valor) cuadra con el DS (±1 ud / ±5 USD).
  // Así se evita "sobrepasar por mucho": no asignar 54 filas Layout que suman 1M ud a un DS de 100 ud.
  const layoutByCandado = new Map();
  for (const row of layoutRows) {
    if (row.noIncluir) continue;
    const candadoVal = normStr(row.Candado ?? "");
    if (!candadoVal) continue;
    if (!layoutByCandado.has(candadoVal)) layoutByCandado.set(candadoVal, []);
    layoutByCandado.get(candadoVal).push(row);
  }
  // Tolerancia: cantidad ±1 ud o ±0.5% cuando son grandes (redondeos); valor ±5 USD o no exigir si DS no tiene valor
  const tolCandadoC = 1, tolCandadoV = 5;
  const tolCandadoPct = 0.005; // 0.5% para cantidades grandes
  for (const [candadoVal, rows] of layoutByCandado) {
    const dsRow = dsByCandado.get(candadoVal);
    if (!dsRow || usedDS.has(dsRow._dsIdx)) continue;
    const sumC = rows.reduce((a, r) => a + r.Cantidad, 0);
    const sumV = rows.reduce((a, r) => a + r.ValorUSD, 0);
    const dsCant = parseFloat(dsRow["CantidadUMComercial"]) || 0;
    const dsVal  = parseFloat(dsRow["ValorDolares"]) || 0;
    const dsSinValor = dsVal === 0 || !Number.isFinite(dsVal);
    const cuadraCant = Math.abs(sumC - dsCant) <= tolCandadoC
      || (dsCant > 1000 && Math.abs(sumC - dsCant) / dsCant <= tolCandadoPct);
    const cuadraVal  = Math.abs(sumV - dsVal) <= tolCandadoV || dsSinValor;
    if (cuadraCant && cuadraVal) {
      assignRows(rows, dsRow, "E0candado");
    }
    // Si no cuadra: las filas pasan a FASE A (A1, A2, findSubset, etc.)
  }

  //  E0: Verificar SEC CALC existente contra DS 
  // Intento 1: por "Candado 551" exacto (Pedimento-Fraccion-Secuencia)
  // Intento 2: por Ped+Frac+SecuenciaFraccion directamente (cubre formatos distintos de candado)
  // Solo para filas que aún no se asignaron por E0candado y que ya tienen secuencia en Layout.
  for (const row of layoutRows) {
    if (row.noIncluir) continue;
    if (assignment.has(row._idx)) continue; // ya asignada por E0candado
    if (!isRealSec(row.SecCalc)) continue;
    const ped2 = normStr(row.Pedimento);
    const frac = nFrac(row.FraccionNico);
    const sec  = normStr(row.SecCalc);

    // Intento 1: candado 551
    const candado = `${row.Pedimento}-${frac}-${row.SecCalc}`;
    let dsRow = dsByCandado.get(candado);

    // Intento 2: buscar en DS por Ped+Frac+Sec directamente
    if (!dsRow) {
      const cands = dsByPF.get(`${ped2}|||${frac}`) || [];
      dsRow = cands.find(ds => normStr(ds["SecuenciaFraccion"]) === sec) || null;
    }

    if (dsRow && !usedDS.has(dsRow._dsIdx)) {
      // DS sec encontrada y aún no usada  verificar OK
      usedDS.add(dsRow._dsIdx);
      assignment.set(row._idx, { status: "ok", newSec: row.SecCalc, dsRow, corrections: [],
        reason: "OK  Secuencia verificada contra DS" });
    }
    // Si la DS sec ya fue usada por otra fila (duplicado con misma sec), esta fila
    // pasa a fases siguientes para encontrar su secuencia correcta (ej: Sec22Sec23)
  }

  //  E1E7: Asignar filas sin secuencia (o cuya sec no coincidió en E0) 
  // Agrupamos por DOS criterios en paralelo:
  //   1. Ped+Frac+Pais  (tradicional, para casos exactos por país)
  //   2. Ped+Frac+DescNorm  (nuevo, para casos donde DS agrupa por descripción
  //      y el Layout mezcla países con la misma descripción)

  const groupsByPais = new Map(); // Ped|||Frac|||Pais  rows
  const groupsByDesc = new Map(); // Ped|||Frac|||DescNorm  rows
  for (const row of layoutRows) {
    if (row.noIncluir) continue;
    if (assignment.has(row._idx)) continue; // ya asignada (E0candado, E0, E1E7, etc.)
    const a = assignment.get(row._idx);
    if (a?.status === "ok") continue;
    const ped  = row.Pedimento;
    const frac = nFrac(row.FraccionNico);
    const pais = normStr(row.PaisOrigen);
    const desc = nDesc(row.Descripcion);
    const kP  = `${ped}|||${frac}|||${pais}`;
    const kD  = `${ped}|||${frac}|||${desc}`;
    if (!groupsByPais.has(kP)) groupsByPais.set(kP, []);
    groupsByPais.get(kP).push(row);
    if (!groupsByDesc.has(kD)) groupsByDesc.set(kD, []);
    groupsByDesc.get(kD).push(row);
  }

  //  FASE A: Para cada DS row, buscar Layout por Ped+Frac+Desc 
  // Esto resuelve el caso PRUEBA 1: DS 85049099/CAJA vs Layout 85049099/CAJA MEX+CHN+USA+TWN
  // Ordena secuencias DS de menor a mayor cant (asigna las más pequeñas primero)
  const dsSorted = [...dsRows].filter(r => !usedDS.has(r._dsIdx))
    .sort((a,b) => (parseFloat(a["CantidadUMComercial"])||0) - (parseFloat(b["CantidadUMComercial"])||0));

  for (const dsRow of dsSorted) {
    if (usedDS.has(dsRow._dsIdx)) continue;
    const ped2  = normStr(dsRow["Pedimento2"]);
    const frac  = getDSFrac(dsRow);
    const desc  = nDesc(dsRow["DescripcionMercancia"]);
    const dsCant = parseFloat(dsRow["CantidadUMComercial"]) || 0;
    const dsVal  = parseFloat(dsRow["ValorDolares"])        || 0;

    // --- A1: Grupo Layout cuya suma Ped+Frac+Desc coincida exactamente ---
    const kD = `${ped2}|||${frac}|||${desc}`;
    const grpDesc = (groupsByDesc.get(kD) || []).filter(r => !assignment.has(r._idx));
    if (grpDesc.length > 0) {
      const sumC = grpDesc.reduce((a,r)=>a+r.Cantidad,0);
      const sumV = grpDesc.reduce((a,r)=>a+r.ValorUSD,0);
      if (Math.abs(sumC - dsCant) <= 1 && Math.abs(sumV - dsVal) <= 4) {
        assignRows(grpDesc, dsRow, "A1");
        continue;
      }
      // Quizá solo es un subconjunto (misma desc pero DS tiene varias secs)
      //  ver si un subconjunto de grpDesc suma exactamente (búsqueda greedy)
      const subset = findSubset(grpDesc, dsCant, dsVal, 1, 4);
      if (subset) { assignRows(subset, dsRow, "A1b"); continue; }
    }

    // --- A2: Grupo Layout por Ped+Frac+Pais exacto ---
    const pais = normStr(dsRow["PaisOrigenDestino"]);
    const kP   = `${ped2}|||${frac}|||${pais}`;
    const grpPais = (groupsByPais.get(kP) || []).filter(r => !assignment.has(r._idx));
    if (grpPais.length > 0) {
      const sumC = grpPais.reduce((a,r)=>a+r.Cantidad,0);
      const sumV = grpPais.reduce((a,r)=>a+r.ValorUSD,0);
      if (Math.abs(sumC - dsCant) <= 1 && Math.abs(sumV - dsVal) <= 4) {
        assignRows(grpPais, dsRow, "A2");
        continue;
      }
      const subset = findSubset(grpPais, dsCant, dsVal, 1, 4);
      if (subset) { assignRows(subset, dsRow, "A2b"); continue; }
    }

    // --- A3: Ped+Frac sin restricción (total o subconjunto = DS) ---
    const kPF  = `${ped2}|||${frac}`;
    const allPFRows = [...groupsByPais.entries()]
      .filter(([k]) => k.startsWith(kPF + "|||"))
      .flatMap(([,rows]) => rows)
      .filter(r => !assignment.has(r._idx));
    if (allPFRows.length > 0) {
      const sumC = allPFRows.reduce((a,r)=>a+r.Cantidad,0);
      const sumV = allPFRows.reduce((a,r)=>a+r.ValorUSD,0);
      if (Math.abs(sumC - dsCant) <= 1 && Math.abs(sumV - dsVal) <= 4) {
        assignRows(allPFRows, dsRow, "A3");
        continue;
      }
      const subset = findSubset(allPFRows, dsCant, dsVal, 1, 4);
      if (subset) { assignRows(subset, dsRow, "A3b"); continue; }
    }

    // --- A1_fuzzy: Desc parcial  primer 50% de la desc DS coincide ---
    // Para casos donde DS tiene desc corta y Layout tiene desc más larga
    const descKey50 = desc.slice(0, Math.max(8, Math.floor(desc.length * 0.6)));
    for (const [k, grp] of groupsByDesc) {
      if (!k.startsWith(`${ped2}|||${frac}|||`)) continue;
      const lyDesc = k.split("|||")[2] || "";
      if (!lyDesc.startsWith(descKey50) && !descKey50.startsWith(lyDesc.slice(0, descKey50.length))) continue;
      const rows2 = grp.filter(r => !assignment.has(r._idx));
      if (!rows2.length) continue;
      const sumC = rows2.reduce((a,r)=>a+r.Cantidad,0);
      const sumV = rows2.reduce((a,r)=>a+r.ValorUSD,0);
      if (Math.abs(sumC - dsCant) <= 1 && Math.abs(sumV - dsVal) <= 4) {
        assignRows(rows2, dsRow, "A1f");
        break;
      }
      const subset = findSubset(rows2, dsCant, dsVal, 1, 4);
      if (subset) { assignRows(subset, dsRow, "A1fb"); break; }
    }
    if (usedDS.has(dsRow._dsIdx)) continue;
  }

  //  FASE A2: Segundo paso  DS aún no usadas, buscar por Ped+Frac total con combo 
  // Para casos donde Layout total Ped+Frac = suma de 2-3 DS secs
  const dsSortedB = [...dsRows].filter(r => !usedDS.has(r._dsIdx))
    .sort((a,b) => (parseFloat(a["CantidadUMComercial"])||0) - (parseFloat(b["CantidadUMComercial"])||0));

  // Agrupar DS no usadas por Ped+Frac
  const unusedByPF = new Map();
  for (const dsRow of dsSortedB) {
    const kPF = `${normStr(dsRow["Pedimento2"])}|||${getDSFrac(dsRow)}`;
    if (!unusedByPF.has(kPF)) unusedByPF.set(kPF, []);
    unusedByPF.get(kPF).push(dsRow);
  }

  for (const [kPF, dsList] of unusedByPF) {
    const allPFRows = [...groupsByPais.entries()]
      .filter(([k]) => k.startsWith(kPF + "|||"))
      .flatMap(([,rows]) => rows)
      .filter(r => !assignment.has(r._idx));
    if (!allPFRows.length) continue;

    if (dsList.length === 1) {
      // Ya fue intentado en Fase A, si llegó aquí no hay match exacto
      const dsR = dsList[0];
      const dsCant = parseFloat(dsR["CantidadUMComercial"])||0;
      const dsVal  = parseFloat(dsR["ValorDolares"])||0;
      const sumC = allPFRows.reduce((a,r)=>a+r.Cantidad,0);
      const sumV = allPFRows.reduce((a,r)=>a+r.ValorUSD,0);
      // Solo asignar si cantidad cuadra estrictamente ±1 y valor ±4
      if (Math.abs(sumC - dsCant) <= 1 && Math.abs(sumV - dsVal) <= 4) {
        assignRows(allPFRows, dsR, "A_forzado");
      }
      continue;
    }

    // Múltiples DS secs: distribuir filas por descripción/país  subconjunto  resto
    const dsSort = [...dsList].sort((a,b)=>(parseFloat(a["CantidadUMComercial"])||0)-(parseFloat(b["CantidadUMComercial"])||0));
    let remaining = [...allPFRows];

    for (let i = 0; i < dsSort.length; i++) {
      if (!remaining.length) break;
      const m = dsSort[i];
      const dsCant     = parseFloat(m["CantidadUMComercial"])||0;
      const dsVal      = parseFloat(m["ValorDolares"])||0;
      const dsDescNorm = nDesc(m["DescripcionMercancia"]);
      const dsPaisNorm = normStr(m["PaisOrigenDestino"]);
      const esUltimo   = (i === dsSort.length - 1);

      let matched = false;

      //  Estrategia 1: desc + país exacto 
      const byDP = remaining.filter(r => nDesc(r.Descripcion) === dsDescNorm && normStr(r.PaisOrigen) === dsPaisNorm);
      if (!matched && byDP.length > 0) {
        const sC = byDP.reduce((a,r)=>a+r.Cantidad,0), sV = byDP.reduce((a,r)=>a+r.ValorUSD,0);
        if (Math.abs(sC - dsCant) <= 1 && Math.abs(sV - dsVal) <= 4) {
          assignRows(byDP, m, "A2m_dp"); matched = true;
        } else {
          const sub = findSubset(byDP, dsCant, dsVal, 1, 4);
          if (sub) { assignRows(sub, m, "A2m_dp_sub"); matched = true; }
        }
      }

      //  Estrategia 2: solo descripción 
      if (!matched) {
        const byD = remaining.filter(r => nDesc(r.Descripcion) === dsDescNorm);
        if (byD.length > 0) {
          const sC = byD.reduce((a,r)=>a+r.Cantidad,0), sV = byD.reduce((a,r)=>a+r.ValorUSD,0);
          if (Math.abs(sC - dsCant) <= 1 && Math.abs(sV - dsVal) <= 4) {
            assignRows(byD, m, "A2m_d"); matched = true;
          } else {
            const sub = findSubset(byD, dsCant, dsVal, 1, 4);
            if (sub) { assignRows(sub, m, "A2m_d_sub"); matched = true; }
          }
        }
      }

      //  Estrategia 3: solo país 
      if (!matched) {
        const byP = remaining.filter(r => normStr(r.PaisOrigen) === dsPaisNorm);
        if (byP.length > 0) {
          const sC = byP.reduce((a,r)=>a+r.Cantidad,0), sV = byP.reduce((a,r)=>a+r.ValorUSD,0);
          if (Math.abs(sC - dsCant) <= 1 && Math.abs(sV - dsVal) <= 4) {
            assignRows(byP, m, "A2m_p"); matched = true;
          } else {
            const sub = findSubset(byP, dsCant, dsVal, 1, 4);
            if (sub) { assignRows(sub, m, "A2m_p_sub"); matched = true; }
          }
        }
      }

      //  Estrategia 4: subconjunto del total remaining 
      if (!matched) {
        if (esUltimo) {
          // ltimo: si remaining cuadra exacto, asignar todo
          const sC = remaining.reduce((a,r)=>a+r.Cantidad,0), sV = remaining.reduce((a,r)=>a+r.ValorUSD,0);
          if (Math.abs(sC - dsCant) <= 1 && Math.abs(sV - dsVal) <= 4) {
            assignRows(remaining, m, "A2m_last"); matched = true;
          } else {
            const sub = findSubset(remaining, dsCant, dsVal, 1, 4);
            if (sub) { assignRows(sub, m, "A2m_last_sub"); matched = true; }
          }
        } else {
          const sub = findSubset(remaining, dsCant, dsVal, 1, 4);
          if (sub) { assignRows(sub, m, "A2m_sub"); matched = true; }
        }
      }

      // Actualizar remaining quitando las filas asignadas
      if (matched) {
        remaining = remaining.filter(r => !assignment.has(r._idx));
      }
    }

    // Si quedan filas sin asignar y hay DS sin usar, intentar asignar al último DS si cuadra
    if (remaining.length > 0) {
      const dsLast  = dsSort[dsSort.length - 1];
      const dsCantL = parseFloat(dsLast["CantidadUMComercial"])||0;
      const dsValL  = parseFloat(dsLast["ValorDolares"])||0;
      const sumCR   = remaining.reduce((a,r)=>a+r.Cantidad,0);
      const sumVR   = remaining.reduce((a,r)=>a+r.ValorUSD,0);
      if (!usedDS.has(dsLast._dsIdx) && Math.abs(sumCR - dsCantL) <= 1 && Math.abs(sumVR - dsValL) <= 4) {
        assignRows(remaining, dsLast, "A2m_rest");
      } else {
        for (const row of remaining) {
          if (!assignment.has(row._idx)) {
            assignment.set(row._idx, { status:"unmatched", newSec:row.SecCalc||"", dsRow:null, corrections:[],
              reason: `Sin match estricto (Cant/Val no coinciden con ninguna sec DS restante)` });
          }
        }
      }
    }
  }

  //  FASE B: E0b  completar remanente de sec ya verificada en E0 
  for (const [kP, rows] of groupsByPais) {
    const pendingRows = rows.filter(r => !assignment.has(r._idx));
    if (pendingRows.length === 0) continue;
    const [ped, frac] = kP.split("|||").slice(0,2);
    const sumCant = pendingRows.reduce((a,r)=>a+r.Cantidad,0);
    for (const dsR of (dsByPF.get(`${ped}|||${frac}`) || [])) {
      if (!usedDS.has(dsR._dsIdx)) continue;
      const dsCant = parseFloat(dsR["CantidadUMComercial"]) || 0;
      let verifiedCant = 0;
      for (const lRow of layoutRows) {
        const a0 = assignment.get(lRow._idx);
        if (a0?.dsRow?._dsIdx === dsR._dsIdx) verifiedCant += lRow.Cantidad;
      }
      const proyectado = verifiedCant + sumCant;
      if (sumCant > 0 && Math.abs(proyectado - dsCant) <= 1) {
        assignRows(pendingRows, dsR, "E0b");
        break;
      }
    }
  }

  //  FASE B2: Fallback cantidad-only (solo si cant global coincide) 
  // Si la cantidad global cuadra pero quedan DS sin asignar (por diferencias en USD
  // o backtracking con subconjuntos complejos), asignar priorizando cantidad ±1.
  // Orden: primero desc+pais, luego desc, luego pais, luego PF completo.
  if (globalCuadra) {
    const dsPendientes = [...dsRows.filter(r => !usedDS.has(r._dsIdx) && (parseFloat(r["CantidadUMComercial"])||0) > 0)]
      .sort((a,b) => (parseFloat(a["CantidadUMComercial"])||0) - (parseFloat(b["CantidadUMComercial"])||0));

    for (const dsRow of dsPendientes) {
      if (usedDS.has(dsRow._dsIdx)) continue;
      const dsCant     = parseFloat(dsRow["CantidadUMComercial"]) || 0;
      const dsVal      = parseFloat(dsRow["ValorDolares"])        || 0;
      const ped2       = normStr(dsRow["Pedimento2"]);
      const frac       = getDSFrac(dsRow);
      const dsDescNorm = nDesc(dsRow["DescripcionMercancia"]);
      const dsPaisNorm = normStr(dsRow["PaisOrigenDestino"]);
      if (dsCant === 0) continue;

      // Filas layout sin asignar del mismo Ped+Frac
      const lyPend = layoutRows.filter(r =>
        !r.noIncluir && !assignment.has(r._idx) &&
        normStr(r.Pedimento) === ped2 && nFrac(r.FraccionNico) === frac
      );
      if (!lyPend.length) continue;

      let matched = false;

      // 1. desc + país, con val ±4
      if (!matched) {
        const pool = lyPend.filter(r => nDesc(r.Descripcion) === dsDescNorm && normStr(r.PaisOrigen) === dsPaisNorm);
        if (pool.length) {
          const sub = findSubset(pool, dsCant, dsVal, 1, 4);
          if (sub) { assignRows(sub, dsRow, "B2_dp"); matched = true; }
        }
      }
      // 2. solo descripción
      if (!matched) {
        const pool = lyPend.filter(r => nDesc(r.Descripcion) === dsDescNorm);
        if (pool.length) {
          const sub = findSubset(pool, dsCant, dsVal, 1, 4);
          if (sub) { assignRows(sub, dsRow, "B2_d"); matched = true; }
        }
      }
      // 3. solo país
      if (!matched) {
        const pool = lyPend.filter(r => normStr(r.PaisOrigen) === dsPaisNorm);
        if (pool.length) {
          const sub = findSubset(pool, dsCant, dsVal, 1, 4);
          if (sub) { assignRows(sub, dsRow, "B2_p"); matched = true; }
        }
      }
      // 4. Ped+Frac completo (cantidad-only)
      if (!matched) {
        const sub = findSubset(lyPend, dsCant, dsVal, 1, 4);
        if (sub) { assignRows(sub, dsRow, "B2_pf"); matched = true; }
      }
    }
  }

  //  FASE B3: Partición PF-total cuando cant PF coincide exactamente 
  // Para grupos donde el total del Ped+Frac entre DS y Layout coincide ±1 en
  // cantidad, pero el backtracking no pudo dividir (ej: 85340004, 48081001).
  // Estrategia: distribuir Layout rows entre DS secs por país, luego por greedy.
  {
    const dsPendB3 = [...dsRows.filter(r => !usedDS.has(r._dsIdx) && (parseFloat(r["CantidadUMComercial"])||0) > 0)]
      .sort((a,b) => (parseFloat(a["CantidadUMComercial"])||0) - (parseFloat(b["CantidadUMComercial"])||0));

    // Agrupar DS pendientes por PF
    const b3ByPF = new Map();
    for(const ds of dsPendB3){
      const k = `${normStr(ds["Pedimento2"])}|||${nFrac(normStr(ds["Fraccion"]))}`;
      if(!b3ByPF.has(k)) b3ByPF.set(k,[]);
      b3ByPF.get(k).push(ds);
    }

    for(const [kPF, dsList] of b3ByPF){
      // Layout pendiente del mismo PF
      const lyPend = layoutRows.filter(r =>
        !r.noIncluir && !assignment.has(r._idx) &&
        `${normStr(r.Pedimento)}|||${nFrac(r.FraccionNico)}` === kPF
      );
      if(!lyPend.length) continue;

      const lySumC = lyPend.reduce((a,r)=>a+r.Cantidad,0);
      const dsSumC = dsList.reduce((a,r)=>a+(parseFloat(r["CantidadUMComercial"])||0),0);
      // Solo aplicar si el total del PF coincide ±1
      if(Math.abs(lySumC - dsSumC) > 1) continue;

      // Distribuir Layout rows entre DS secs: por país, luego greedy
      let remaining = [...lyPend];
      for(let i=0; i<dsList.length; i++){
        if(!remaining.length) break;
        const ds      = dsList[i];
        const dsCant  = parseFloat(ds["CantidadUMComercial"])||0;
        const dsVal   = parseFloat(ds["ValorDolares"])||0;
        const dsPais  = normStr(ds["PaisOrigenDestino"]);
        const isLast  = i === dsList.length - 1;

        // Si es el último, asignar todo lo que queda si suma coincide ±1
        if(isLast){
          const remSum = remaining.reduce((a,r)=>a+r.Cantidad,0);
          if(Math.abs(remSum - dsCant) <= 1) assignRows(remaining, ds, "B3_last");
          break;
        }

        // Intentar asignar por país exacto primero
        const byPais = remaining.filter(r => normStr(r.PaisOrigen) === dsPais);
        if(byPais.length > 0){
          const bpSum = byPais.reduce((a,r)=>a+r.Cantidad,0);
          if(Math.abs(bpSum - dsCant) <= 1){
            assignRows(byPais, ds, "B3_pais");
            const used = new Set(byPais.map(r=>r._idx));
            remaining = remaining.filter(r=>!used.has(r._idx));
            continue;
          }
          // Subconjunto del grupo pais
          const sub = null;
          if(sub){
            assignRows(sub, ds, "B3_pais_sub");
            const used = new Set(sub.map(r=>r._idx));
            remaining = remaining.filter(r=>!used.has(r._idx));
            continue;
          }
        }

        // Subconjunto del total remaining
        const sub2 = findSubset(remaining, dsCant, dsVal, 1, 4);
        if(sub2){
          assignRows(sub2, ds, "B3_sub");
          const used = new Set(sub2.map(r=>r._idx));
          remaining = remaining.filter(r=>!used.has(r._idx));
          continue;
        }

        // Greedy: acumular filas hasta llenar la cuota de cantidad
        // Solo si la cuota es factible (suma de las más pequeñas alcanza dsCant)
        const sorted = [...remaining].sort((a,b)=>a.Cantidad-b.Cantidad);
        let acc=0; const greedyRows=[];
        for(const r of sorted){
          if(acc + r.Cantidad > dsCant + 1) continue; // saltar si excede
          greedyRows.push(r);
          acc += r.Cantidad;
          if(Math.abs(acc - dsCant) <= 1) break;
        }
        if(Math.abs(acc - dsCant) <= 1 && greedyRows.length > 0){
          assignRows(greedyRows, ds, "B3_greedy");
          const used = new Set(greedyRows.map(r=>r._idx));
          remaining = remaining.filter(r=>!used.has(r._idx));
        }
      }
    }
  }

  //  FASE B4: Búsqueda cruzada de fracciones (cross-fraction) 
  // Algunos Layout tienen la fracción incorrecta respecto al DS.
  // Ej: DS sec `85332101` no tiene Layout, pero `85332999` tiene 46k extra que
  //     en el DS corresponden a `85332101`.
  // Estrategia: para DS secs sin Layout en su fracción, buscar en cualquier
  //             fila del MISMO pedimento que esté sin asignar.
  //             Solo si cantidad global coincide (no inventamos unidades).
  if (globalCuadra) {
    const dsPendB4 = [...dsRows.filter(r => !usedDS.has(r._dsIdx) && (parseFloat(r["CantidadUMComercial"])||0) > 0)]
      .sort((a,b) => (parseFloat(b["CantidadUMComercial"])||0) - (parseFloat(a["CantidadUMComercial"])||0)); // desc: primero los más grandes

    for (const dsRow of dsPendB4) {
      if (usedDS.has(dsRow._dsIdx)) continue;
      const dsCant     = parseFloat(dsRow["CantidadUMComercial"]) || 0;
      const dsVal      = parseFloat(dsRow["ValorDolares"])        || 0;
      const ped2       = normStr(dsRow["Pedimento2"]);
      const frac       = getDSFrac(dsRow);
      const dsDescNorm = nDesc(dsRow["DescripcionMercancia"]);
      const dsPaisNorm = normStr(dsRow["PaisOrigenDestino"]);
      if (dsCant === 0) continue;

      // Buscar en TODO el pedimento (cualquier fracción), sin asignar
      const lyAllPed = layoutRows.filter(r =>
        !r.noIncluir && !assignment.has(r._idx) &&
        normStr(r.Pedimento) === ped2
      );
      if (!lyAllPed.length) continue;

      let matched = false;

      // 1. desc exacta + país (cualquier fracción)
      if (!matched) {
        const pool = lyAllPed.filter(r => nDesc(r.Descripcion) === dsDescNorm && normStr(r.PaisOrigen) === dsPaisNorm);
        if (pool.length) {
          const sub = findSubset(pool, dsCant, dsVal, 1, 4);
          if (sub) { assignRows(sub, dsRow, "B4_dp", frac); matched = true; }
        }
      }
      // 2. solo descripción exacta (cualquier fracción, cualquier país)
      if (!matched) {
        const pool = lyAllPed.filter(r => nDesc(r.Descripcion) === dsDescNorm);
        if (pool.length) {
          const sub = findSubset(pool, dsCant, dsVal, 1, 4);
          if (sub) { assignRows(sub, dsRow, "B4_d", frac); matched = true; }
        }
      }
      // 3. descripción parcial (prefijo 60%) + país
      if (!matched) {
        const pref = dsDescNorm.slice(0, Math.max(8, Math.floor(dsDescNorm.length * 0.6)));
        const pool = lyAllPed.filter(r => {
          const ld = nDesc(r.Descripcion);
          return normStr(r.PaisOrigen) === dsPaisNorm &&
                 (ld.startsWith(pref) || pref.startsWith(ld.slice(0, pref.length)));
        });
        if (pool.length) {
          const sub = findSubset(pool, dsCant, dsVal, 1, 4);
          if (sub) { assignRows(sub, dsRow, "B4_fp", frac); matched = true; }
        }
      }
      // 4. descripción parcial (cualquier país)
      if (!matched) {
        const pref = dsDescNorm.slice(0, Math.max(8, Math.floor(dsDescNorm.length * 0.6)));
        const pool = lyAllPed.filter(r => {
          const ld = nDesc(r.Descripcion);
          return ld.startsWith(pref) || pref.startsWith(ld.slice(0, pref.length));
        });
        if (pool.length) {
          const sub = findSubset(pool, dsCant, dsVal, 1, 4);
          if (sub) { assignRows(sub, dsRow, "B4_fd", frac); matched = true; }
        }
      }
      // 5. solo país + cantidad
      if (!matched) {
        const pool = lyAllPed.filter(r => normStr(r.PaisOrigen) === dsPaisNorm);
        if (pool.length) {
          const sub = findSubset(pool, dsCant, dsVal, 1, 4);
          if (sub) { assignRows(sub, dsRow, "B4_p", frac); matched = true; }
        }
      }
      // 6. Modo estricto: se elimina fallback "solo cantidad"
    }
  }

  //  FASE B6: Relleno final por similitud (permite reutilizar secuencias) 
  // Objetivo: cerrar asignación al 100% en archivos donde los métodos estrictos
  // no alcanzan por diferencias operativas entre Layout y DS.
  // Criterio de mejor candidato: misma fracción (si existe), descripción, país
  // y cercanía de precio unitario. Aquí SÍ se permite reutilizar secuencia DS.
  {
    const dsByPed = new Map();
    for (const ds of dsRows) {
      const ped = normStr(ds["Pedimento2"]);
      if (!dsByPed.has(ped)) dsByPed.set(ped, []);
      dsByPed.get(ped).push(ds);
    }

    const descPrefixMatch = (a, b) => {
      const x = nDesc(a), y = nDesc(b);
      if (!x || !y) return false;
      const pref = Math.max(8, Math.floor(Math.min(x.length, y.length) * 0.6));
      return x.slice(0, pref) === y.slice(0, pref);
    };

    const chooseBestDsForRow = (row) => {
      const ped  = normStr(row.Pedimento);
      const frac = nFrac(row.FraccionNico);
      const rDesc = nDesc(row.Descripcion || "");
      const rPais = normStr(row.PaisOrigen || "");
      const rUp   = row.Cantidad > 0 ? (row.ValorUSD / row.Cantidad) : 0;

      const byPF = dsByPF.get(`${ped}|||${frac}`) || [];
      const cands = byPF.length ? byPF : (dsByPed.get(ped) || []);
      if (!cands.length) return null;

      let best = null;
      let bestScore = -Infinity;

      for (const ds of cands) {
        const dsFrac = getDSFrac(ds);
        const dsDesc = nDesc(ds["DescripcionMercancia"] || "");
        const dsPais = normStr(ds["PaisOrigenDestino"] || "");
        const dsCant = parseFloat(ds["CantidadUMComercial"]) || 0;
        const dsVal  = parseFloat(ds["ValorDolares"]) || 0;
        const dsUp   = dsCant > 0 ? (dsVal / dsCant) : 0;

        let score = 0;
        if (dsFrac === frac) score += 30;
        if (dsPais && rPais && dsPais === rPais) score += 20;
        if (dsDesc && rDesc && dsDesc === rDesc) score += 80;
        else if (descPrefixMatch(dsDesc, rDesc)) score += 35;

        if (rUp > 0 && dsUp > 0) {
          const upDiff = Math.abs(dsUp - rUp) / Math.max(rUp, 0.0001);
          score -= Math.min(60, upDiff * 100);
        }

        // Pequeño sesgo para no preferir secuencias con cantidad cero.
        if (dsCant <= 0) score -= 20;

        if (score > bestScore) {
          bestScore = score;
          best = ds;
        }
      }
      return best;
    };

    for (const row of layoutRows) {
      if (row.noIncluir || assignment.has(row._idx)) continue;

      const dsBest = chooseBestDsForRow(row);
      if (!dsBest) continue;

      const newSec = normStr(dsBest["SecuenciaFraccion"]);
      const fracOriginal = nFrac(row.FraccionNico);
      const dsFrac       = getDSFrac(dsBest);
      const isCrossFrac  = fracOriginal !== dsFrac;

      assignment.set(row._idx, {
        status: "new",
        newSec,
        dsRow: dsBest,
        corrections: [],
        estrategia: "B6_relleno",
        secOrig: null,
        fracCorr: isCrossFrac ? dsFrac : null,
        fracOrig: isCrossFrac ? fracOriginal : null,
        reason: `Alternativa: [B6_relleno] Sec=${newSec}  asignación por mejor similitud (desc/pais/precio unitario)${isCrossFrac ? ` [Frac ${fracOriginal}${dsFrac}]` : ""}`,
      });
    }
  }

  // Validaci�n estricta final: rechazar asignaciones fuera de tolerancia
  {
    const byDs = new Map(); // dsIdx -> rows[]
    for (const row of layoutRows) {
      const a = assignment.get(row._idx);
      if (!a || !a.dsRow || a.status === "unmatched") continue;
      const di = a.dsRow._dsIdx;
      if (!byDs.has(di)) byDs.set(di, []);
      byDs.get(di).push(row);
    }
    for (const [di, rows] of byDs.entries()) {
      const a0 = assignment.get(rows[0]._idx);
      const ds = a0?.dsRow;
      if (!ds) continue;
      const dsCant = parseFloat(ds["CantidadUMComercial"]) || 0;
      const dsVal  = parseFloat(ds["ValorDolares"]) || 0;
      const sumC = rows.reduce((a, r) => a + (parseFloat(r.Cantidad) || 0), 0);
      const sumV = rows.reduce((a, r) => a + (parseFloat(r.ValorUSD) || 0), 0);
      const diffCant = sumC - dsCant;
      const diffVal  = sumV - dsVal;
      const okCant = Math.abs(diffCant) <= 1;
      const okVal  = Math.abs(diffVal) <= 4;
      if (!(okCant && okVal)) {
        usedDS.delete(di);
        for (const row of rows) {
          assignment.set(row._idx, {
            status: "unmatched",
            newSec: row.SecCalc || "",
            dsRow: null,
            corrections: [],
            reason: `[RECHAZADO ESTRICTO] Sec ${normStr(ds["SecuenciaFraccion"])} descartada por exceder tolerancia: �Cant=${diffCant.toLocaleString("es-MX")} (tol �1), �USD=${diffVal.toFixed(2)} (tol �4). Estrategia previa: ${a0?.estrategia || "N/A"}.`,
          });
        }
      }
    }
  }

  //  FASE C: Marcar sin match las filas que quedaron sin asignar 
  for (const row of layoutRows) {
    if (row.noIncluir || assignment.has(row._idx)) continue;
    assignment.set(row._idx, {
      status: "unmatched", newSec: row.SecCalc || "", dsRow: null, corrections: [],
      reason: `Sin match en DS para Ped ${row.Pedimento} / Frac ${row.FraccionNico} / Desc "${(row.Descripcion||"").slice(0,30)}"`,
    });
  }


  //  Totales asignados por sec (para verificación) 
  // Secuencias DS no usadas: solo contar las que tienen cantidad real (>0)
  const unusedDS = dsRows.filter(r => !usedDS.has(r._dsIdx) && (parseFloat(r["CantidadUMComercial"]) || 0) > 0);
  const stats = { verified: 0, corrected: 0, newAssigned: 0, unmatched: 0 };
  for (const a of assignment.values()) {
    if (a.status === "ok")             stats.verified++;
    else if (a.status === "corrected") stats.corrected++;
    else if (a.status === "new")       stats.newAssigned++;
    else                               stats.unmatched++;
  }

  //  Calcular motivos de no-match para DS no usadas 
  // Construir totales del Layout por Pedimento+Fraccion (para el diagnóstico)
  const layoutTotals = new Map(); // "ped|||frac"  { cant, val, rows, noInc }
  for (const row of layoutRows) {
    const k = `${row.Pedimento}|||${nFrac(row.FraccionNico)}`;
    if (!layoutTotals.has(k)) layoutTotals.set(k, { cant: 0, val: 0, rows: 0, noInc: 0 });
    const g = layoutTotals.get(k);
    g.cant += row.Cantidad; g.val += row.ValorUSD; g.rows++;
    if (row.noIncluir) g.noInc++;
  }

  const fmt = n => Number(n).toLocaleString("es-MX", { maximumFractionDigits: 0 });

  // Calcular totales Layout ASIGNADOS por DS row (para comparar con DS)
  const assignedTotalsByDS = new Map(); // _dsIdx  { cant, val }
  for (const row of layoutRows) {
    const a = assignment.get(row._idx);
    if (!a || a.status === "unmatched" || !a.dsRow) continue;
    const di = a.dsRow._dsIdx;
    if (!assignedTotalsByDS.has(di)) assignedTotalsByDS.set(di, { cant: 0, val: 0 });
    const g = assignedTotalsByDS.get(di);
    g.cant += row.Cantidad; g.val += row.ValorUSD;
  }

  const mismatchReasons = new Map(); // _dsIdx  reason string
  const unusedDSDetails = [];
  const layoutByPed = new Map(); // ped -> [{ frac, cant, val, rows, noInc }]
  for (const [k, g] of layoutTotals.entries()) {
    const [pedK, fracK] = String(k).split("|||");
    if (!layoutByPed.has(pedK)) layoutByPed.set(pedK, []);
    layoutByPed.get(pedK).push({ frac: fracK, ...g });
  }
  // Para DS usadas: verificar si el total Layout coincide con DS
  for (const dsRow of dsRows) {
    const dsCant = parseFloat(dsRow["CantidadUMComercial"]) || 0;
    if (dsCant === 0) continue; // ignorar filas sin cantidad (totales, vacíos)
    const dsVal  = parseFloat(dsRow["ValorDolares"])        || 0;
    const asg    = assignedTotalsByDS.get(dsRow._dsIdx);

    if (!usedDS.has(dsRow._dsIdx)) {
      // DS no usada  calcular motivo
      const frac = getDSFrac(dsRow);
      const k    = `${normStr(dsRow["Pedimento2"])}|||${frac}`;
      const g    = layoutTotals.get(k);
      let reason;
      if (!g) {
        reason = `Fracción ${dsRow["Fraccion"]} no encontrada en Layout (ped. ${dsRow["Pedimento2"]})`;
      } else if (g.noInc === g.rows) {
        reason = `Todas las filas Layout (${g.rows}) marcadas NO INCLUIR`;
      } else {
        const diffCant = g.cant - dsCant;
        const diffVal  = g.val  - dsVal;
        const pctCant  = dsCant > 0 ? (diffCant / dsCant * 100).toFixed(0) : "";
        const pctVal   = dsVal  > 0 ? (diffVal  / dsVal  * 100).toFixed(0) : "";
        reason = `Sin concordancia  Cant.Layout=${fmt(g.cant)} vs DS=${fmt(dsCant)} (${diffCant>=0?"+":""}${pctCant}%) | Valor Layout=$${fmt(g.val)} vs DS=$${fmt(dsVal)} (${diffVal>=0?"+":""}${pctVal}%)`;
      }
      mismatchReasons.set(dsRow._dsIdx, reason);
      // Probable candidato Layout (misma clave de pedimento) para explicar "DS vacío"
      let probable = "";
      const pedK = normStr(dsRow["Pedimento2"]);
      const candsPed = layoutByPed.get(pedK) || [];
      if (!g && candsPed.length > 0) {
        let best = null;
        let bestScore = Infinity;
        for (const c of candsPed) {
          const relC = Math.abs(c.cant - dsCant) / Math.max(c.cant, dsCant, 1);
          const relV = Math.abs(c.val - dsVal) / Math.max(c.val, dsVal, 1);
          const score = relC * 0.6 + relV * 0.4;
          if (score < bestScore) { bestScore = score; best = c; }
        }
        if (best) {
          const prob = Math.max(5, Math.round((1 / (1 + bestScore)) * 100));
          probable = `Probable relación en Layout: fracción ${best.frac}, Cant=${fmt(best.cant)}, Val=$${fmt(best.val)} (confianza aprox. ${prob}%).`;
        }
      }
      unusedDSDetails.push({
        dsIdx: dsRow._dsIdx,
        secuencia: normStr(dsRow["SecuenciaFraccion"]),
        pedimento: normStr(dsRow["Pedimento2"] || dsRow["Pedimento"]),
        fraccion: normStr(dsRow["Fraccion"] || frac),
        pais: normStr(dsRow["PaisOrigenDestino"]),
        cantidadDS: dsCant,
        valorDS: dsVal,
        razon: reason,
        probable,
      });
    } else if (asg) {
      // DS usada: verificar si el total asignado coincide con DS
      const diffCant = Math.abs(asg.cant - dsCant);
      const diffVal  = Math.abs(asg.val  - dsVal);
      const tolCant  = 1; // tolerancia estricta comercio exterior
      const tolVal   = 4;
      if (diffCant > tolCant || diffVal > tolVal) {
        const pctC = dsCant > 0 ? ((asg.cant - dsCant) / dsCant * 100).toFixed(1) : "";
        const pctV = dsVal  > 0 ? ((asg.val  - dsVal)  / dsVal  * 100).toFixed(1) : "";
        const signo = (v) => v >= 0 ? "+" : "";
        mismatchReasons.set(dsRow._dsIdx,
          ` DISCREPANCIA  Cant.Layout=${fmt(asg.cant)} vs DS=${fmt(dsCant)} (${signo((asg.cant-dsCant))}${pctC}%) | Valor Layout=$${fmt(asg.val)} vs DS=$${fmt(dsVal)} (${signo((asg.val-dsVal))}${pctV}%)`
        );
      }
    }
  }

  return { assignment, stats, unusedDS, unusedDSDetails, layoutRows, dsRows, mismatchReasons,
    globalTotals: { dsCant: dsTotalC, dsVal: dsTotalV, lyCant: lyTotalC, lyVal: lyTotalV, cuadra: globalCantCuadra, cuadraVal: globalValCuadra } };
}

/** Construye el Excel 2020 de salida: modifica celdas específicas del Layout. */
/**
 * Modifica DIRECTAMENTE las celdas del worksheet original (sin reconstruirlo).
 * Luego crea un libro de salida con:
 *   - La hoja Layout modificada (celdas SEC CALC + NOTAS en verde/rojo)
 *   - La hoja DS sin cambios
 *   - Una hoja "Reporte_2020" con el detalle
 *
 * Escribir celdas directamente es ~1000x más eficiente que sheet_to_json  aoa_to_sheet
 * sobre una hoja de 22 millones de celdas.
 */
function buildOutput2020Excel(workbook, layoutSheetName, dsSheetName,
                               layout2020Data, assignment, mismatchReasons) {
  const { layoutRows, colIdx } = layout2020Data;
  const ws = workbook.Sheets[layoutSheetName]; // referencia al worksheet original

  //  Estilos 
  const S_OK_SEC   = { font:{bold:true,color:{rgb:"145A32"}}, fill:{patternType:"solid",fgColor:{rgb:"D5F5E3"}}, alignment:{horizontal:"center"} };
  const S_NEW_SEC  = { font:{bold:true,color:{rgb:"7B241C"}}, fill:{patternType:"solid",fgColor:{rgb:"FFCCCC"}}, alignment:{horizontal:"center"} };
  const S_OK_NOTA  = { font:{italic:true,sz:10,color:{rgb:"145A32"}}, fill:{patternType:"solid",fgColor:{rgb:"D5F5E3"}}, alignment:{wrapText:true} };
  const S_NEW_NOTA = { font:{bold:true,sz:10,color:{rgb:"641E16"}}, fill:{patternType:"solid",fgColor:{rgb:"FADBD8"}}, alignment:{wrapText:true} };
  const S_CORR_FLD = { font:{bold:true,color:{rgb:"7B241C"}}, fill:{patternType:"solid",fgColor:{rgb:"FFCCCC"}}, alignment:{horizontal:"center",wrapText:true} };
  const styleAmarillo = { font:{bold:true,color:{rgb:"7D6608"}}, fill:{patternType:"solid",fgColor:{rgb:"FCF3CF"}}, alignment:{horizontal:"center"} };
  const styleAmarilloNota = { font:{italic:true,sz:10,color:{rgb:"7D6608"}}, fill:{patternType:"solid",fgColor:{rgb:"FCF3CF"}}, alignment:{wrapText:true} };
  // Estilo morado/azul para fracción corregida cross-fraction
  const S_FRAC_CORR     = { font:{bold:true,color:{rgb:"4A235A"}}, fill:{patternType:"solid",fgColor:{rgb:"E8DAEF"}}, alignment:{horizontal:"center"} };
  const S_FRAC_CORR_NOTA= { font:{italic:true,sz:10,color:{rgb:"4A235A"}}, fill:{patternType:"solid",fgColor:{rgb:"E8DAEF"}}, alignment:{wrapText:true} };
  // Cadena Layout vs DS: verde si coinciden tal cual, naranja si no
  const S_CADENA_OK   = { font:{sz:9,color:{rgb:"145A32"}}, fill:{patternType:"solid",fgColor:{rgb:"D5F5E3"}}, alignment:{wrapText:true,vertical:"top"} };
  const S_CADENA_DIFF = { font:{sz:9,color:{rgb:"92400E"}}, fill:{patternType:"solid",fgColor:{rgb:"FFEDD5"}}, alignment:{wrapText:true,vertical:"top"} };

  // Construir cadena (Pedimento2-Fraccion-Descripcion-Pais-Cantidad-Valor) para comparación Layout vs DS
  const normStr = (v) => String(v ?? "").trim();
  const nFrac = (v) => normStr(v).replace(/^0+/, "") || "0";
  const nDesc = (s) => String(s ?? "").trim().toLowerCase().replace(/\s*\/\s*/g, "/").replace(/\s+/g, " ");
  const getDSFrac = (dsRow) => {
    const f = nFrac(dsRow["Fraccion"]);
    if (f) return f;
    const candado = normStr(dsRow["Candado 551"] ?? "");
    if (!candado) return "";
    const parts = candado.split("-");
    for (let i = 0; i < parts.length; i++) {
      const p = parts[i].replace(/\./g, "");
      if (p.length === 8 && /^\d+$/.test(p)) return p;
    }
    return "";
  };
  const buildCadenaLayout = (row) => {
    const ped = normStr(row.Pedimento);
    const frac = nFrac(row.FraccionNico ?? "");
    const desc = nDesc(row.Descripcion ?? "");
    const pais = normStr(row.PaisOrigen ?? "");
    const cant = Math.round(parseFloat(row.Cantidad) || 0);
    const val  = Math.round(parseFloat(row.ValorUSD) || 0);
    return `${ped}|${frac}|${desc}|${pais}|${cant}|${val}`;
  };
  const buildCadenaDS = (dsRow) => {
    const ped  = normStr(dsRow["Pedimento2"] ?? "");
    const frac = getDSFrac(dsRow);
    const desc = nDesc(dsRow["DescripcionMercancia"] ?? "");
    const pais = normStr(dsRow["PaisOrigenDestino"] ?? "");
    const cant = Math.round(parseFloat(dsRow["CantidadUMComercial"]) || 0);
    const val  = Math.round(parseFloat(dsRow["ValorDolares"]) || 0);
    return `${ped}|${frac}|${desc}|${pais}|${cant}|${val}`;
  };

  const cadenaCol = colIdx.val >= 0 ? colIdx.val + 1 : -1;
  const notasModifCol = colIdx.notas >= 0 ? colIdx.notas + 1 : -1;
  const hdrRowI = layoutRows.length > 0 ? layoutRows[0]._rowI - 1 : 0;

  // Notas de modificaciones (igual que en la tabla del sistema): correcciones o "OK"
  const normT = (s) => String(s ?? "").trim().toUpperCase();
  const buildNotasModifExcel = (layoutRow, a) => {
    if (!a || a.status === "unmatched" || !a.dsRow) return "OK";
    const ds = a.dsRow;
    const partes = [];
    const paisLy = normT(layoutRow.PaisOrigen || "");
    const paisDs = normT(ds["PaisOrigenDestino"] || "");
    if (paisLy && paisDs && paisLy !== paisDs) partes.push(`Se corrigió país ${paisLy} a ${paisDs}`);
    const descLy = String(layoutRow.Descripcion || "").trim();
    const descDs = String(ds["DescripcionMercancia"] || "").trim();
    if (descLy !== descDs && descDs) partes.push(`Se corrigió descripción "${descLy}" a "${descDs}"`);
    if (a.fracCorr && a.fracOrig) partes.push(`Se corrigió fracción ${a.fracOrig} a ${a.fracCorr}`);
    if (a.secOrig && a.newSec && String(a.secOrig) !== String(a.newSec)) partes.push(`Se corrigió secuencia ${a.secOrig} a ${a.newSec}`);
    return partes.length ? partes.join("; ") : "OK";
  };

  // Filas repetidas (mismo Ped+Frac+Pais+Cant+Val)  pintar amarillo (solo si tienen datos)
  const keyDup = (row) => `${row.Pedimento}|||${(row.FraccionNico||"").trim()}|||${row.PaisOrigen}|||${row.Cantidad}|||${row.ValorUSD}`;
  const keyCounts = new Map();
  for (const row of layoutRows) {
    if (!row.Pedimento && !row.FraccionNico) continue; // no marcar filas vacías como duplicadas
    const k = keyDup(row);
    keyCounts.set(k, (keyCounts.get(k) || 0) + 1);
  }
  const dupIdx = new Set();
  for (const row of layoutRows) {
    if (!row.Pedimento && !row.FraccionNico) continue;
    if (keyCounts.get(keyDup(row)) > 1) dupIdx.add(row._idx);
  }

  const setCell = (ws, r, c, val, style) => {
    if (c < 0 || !ws) return;
    const addr = XLSX.utils.encode_cell({ r, c });
    const t = (typeof val === "number") ? "n" : "s";
    ws[addr] = { t, v: val, s: style };
  };

  if (ws && cadenaCol >= 0) {
    setCell(ws, hdrRowI, cadenaCol, "Cadena (Layout vs DS)", { font: { bold: true, sz: 10 }, fill: { patternType: "solid", fgColor: { rgb: "E5E7EB" } }, alignment: { wrapText: true } });
  }
  if (ws && notasModifCol >= 0) {
    setCell(ws, hdrRowI, notasModifCol, "Notas (modificaciones)", { font: { bold: true, sz: 10 }, fill: { patternType: "solid", fgColor: { rgb: "E5E7EB" } }, alignment: { wrapText: true } });
  }

  //  Modificar celdas en el worksheet original 
  if (ws) {
    for (const row of layoutRows) {
      const a = assignment.get(row._idx);
      const esDup = dupIdx.has(row._idx);
      const r = row._rowI;

      if (a && a.status !== "unmatched") {
        const newSecVal  = isNaN(parseFloat(a.newSec)) ? a.newSec : parseFloat(a.newSec);
        const isOk       = a.status === "ok";
        const isCrossFrac = !!a.fracCorr;

        // Elegir estilo: morado=cross-fraction, amarillo=duplicado, verde=ok, rojo=nuevo
        const styleSec  = isCrossFrac ? S_FRAC_CORR
                        : esDup       ? styleAmarillo
                        : isOk        ? S_OK_SEC
                        :               S_NEW_SEC;
        const styleNota = isCrossFrac ? S_FRAC_CORR_NOTA
                        : esDup       ? styleAmarilloNota
                        : isOk        ? S_OK_NOTA
                        :               S_NEW_NOTA;

        // SEC CALC  si la sec fue corregida, mostrar "Nueva (antes: Vieja)"
        let secDisplayVal = newSecVal;
        if (a.secOrig && !isOk) {
          // Escribir solo el número nuevo en la celda (el "antes" va en NOTAS)
          secDisplayVal = newSecVal;
        }
        setCell(ws, r, colIdx.sec, secDisplayVal, styleSec);

        // Sincronizar campos maestros con DS (551): Fracción, Descripción y País
        // En CE, el DS es la referencia final para estos datos.
        const dsFracOut = getDSFrac(a.dsRow);
        const dsDescOut = normStr(a.dsRow["DescripcionMercancia"] || "");
        const dsPaisOut = normStr(a.dsRow["PaisOrigenDestino"] || "");
        const lyFracOut = normStr(row.FraccionNico || "");
        const lyDescOut = normStr(row.Descripcion || "");
        const lyPaisOut = normStr(row.PaisOrigen || "");

        if (colIdx.frac >= 0 && dsFracOut && lyFracOut !== dsFracOut) {
          setCell(ws, r, colIdx.frac, dsFracOut, S_CORR_FLD);
        } else if (isCrossFrac && colIdx.frac >= 0) {
          // Compatibilidad con resaltado histórico de cross-fraction
          setCell(ws, r, colIdx.frac, a.fracCorr, S_FRAC_CORR);
        }
        if (colIdx.desc >= 0 && dsDescOut && lyDescOut !== dsDescOut) {
          setCell(ws, r, colIdx.desc, dsDescOut, S_CORR_FLD);
        }
        if (colIdx.pais >= 0 && dsPaisOut && lyPaisOut !== dsPaisOut) {
          setCell(ws, r, colIdx.pais, dsPaisOut, S_CORR_FLD);
        }

        // NOTAS  incluir "(corregido de X)" si había secuencia previa incorrecta
        let notaText = a.reason;
        if (a.secOrig && !isOk && a.secOrig !== String(newSecVal)) {
          notaText = ` Sec anterior: ${a.secOrig}  Corregida a ${newSecVal} | ${a.reason}`;
        }
        setCell(ws, r, colIdx.notas, notaText, styleNota);

        // Cadena (Layout vs DS) al lado de Valor USD: verde si coinciden tal cual, naranja si no
        if (cadenaCol >= 0) {
          const cadenaLay = buildCadenaLayout(row);
          const cadenaDs  = buildCadenaDS(a.dsRow);
          const coinciden = cadenaLay === cadenaDs;
          const txtCadena = `L: ${cadenaLay}\nDS: ${cadenaDs}`;
          setCell(ws, r, cadenaCol, txtCadena, coinciden ? S_CADENA_OK : S_CADENA_DIFF);
        }
        // Notas (modificaciones): correcciones de país, descripción, fracción, secuencia o OK
        if (notasModifCol >= 0) setCell(ws, r, notasModifCol, buildNotasModifExcel(row, a), S_OK_NOTA);
      } else if (esDup) {
        // Fila repetida sin match: pintar SEC CALC y NOTAS en amarillo para visibilidad
        setCell(ws, r, colIdx.sec, row.SecCalc || ".", styleAmarillo);
        const notaUnmatched = a?.reason || `Sin match en DS para Ped ${row.Pedimento} / Frac ${row.FraccionNico}`;
        setCell(ws, r, colIdx.notas, notaUnmatched, styleAmarilloNota);
        if (cadenaCol >= 0) {
          const cadenaLay = buildCadenaLayout(row);
          setCell(ws, r, cadenaCol, `L: ${cadenaLay}\nDS: `, S_CADENA_DIFF);
        }
        if (notasModifCol >= 0) setCell(ws, r, notasModifCol, "OK", styleAmarilloNota);
      } else if (row.SecCalc && String(row.SecCalc).trim() !== "" && String(row.SecCalc).trim() !== "." && !isNaN(parseFloat(String(row.SecCalc)))) {
        // Fila con secuencia existente que NO pudo verificarse ni corregirse
        //  marcar en naranja oscuro para que el usuario la revise manualmente
        const S_SEC_REVISAR  = { font:{bold:true,color:{rgb:"6E2C00"}}, fill:{patternType:"solid",fgColor:{rgb:"FDEBD0"}}, alignment:{horizontal:"center"} };
        const S_NOTA_REVISAR = { font:{italic:true,sz:10,color:{rgb:"6E2C00"}}, fill:{patternType:"solid",fgColor:{rgb:"FDEBD0"}}, alignment:{wrapText:true} };
        setCell(ws, r, colIdx.sec, `${row.SecCalc} `, S_SEC_REVISAR);
        const motivo = a?.reason || `Sec ${row.SecCalc} no coincide con DS  revisar manualmente`;
        setCell(ws, r, colIdx.notas, ` REVISAR: ${motivo}`, S_NOTA_REVISAR);
        if (cadenaCol >= 0) {
          const cadenaLay = buildCadenaLayout(row);
          const cadenaDs  = a?.dsRow ? buildCadenaDS(a.dsRow) : "";
          setCell(ws, r, cadenaCol, `L: ${cadenaLay}\nDS: ${cadenaDs}`, S_CADENA_DIFF);
        }
        if (notasModifCol >= 0) setCell(ws, r, notasModifCol, buildNotasModifExcel(row, a) || "OK", S_NOTA_REVISAR);
      } else {
        if (cadenaCol >= 0) {
          const cadenaLay = buildCadenaLayout(row);
          setCell(ws, r, cadenaCol, `L: ${cadenaLay}\nDS: `, S_CADENA_DIFF);
        }
        if (notasModifCol >= 0) setCell(ws, r, notasModifCol, "OK", { font: { sz: 10 }, alignment: { wrapText: true } });
      }
    }
  }

  //  Escribir motivos de no-match en columna REVISADO del DS 
  const dsWsOut = workbook.Sheets[dsSheetName];
  const dsRows  = layout2020Data.dsRows;  // las filas del DS con _rowI y _dsIdx
  let revCol    = dsRows?._revisadoCol ?? -1;

  // Si no existe columna REVISADO, crearla al final del encabezado
  if (dsWsOut && dsRows && revCol < 0) {
    const range = dsWsOut["!ref"] ? XLSX.utils.decode_range(dsWsOut["!ref"]) : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
    revCol = range.e.c + 1;
    const hdrI = dsRows._hdrI ?? 0;
    const addrHdr = XLSX.utils.encode_cell({ r: hdrI, c: revCol });
    const styleRevisadoHdr = { font: { bold: true }, fill: { patternType: "solid", fgColor: { rgb: "E8DAEF" } } };
    dsWsOut[addrHdr] = { t: "s", v: "REVISADO", s: styleRevisadoHdr };
    if (!dsWsOut["!cols"]) dsWsOut["!cols"] = [];
    dsWsOut["!cols"][revCol] = { wch: 80 };
    range.e.c = revCol;
    dsWsOut["!ref"] = XLSX.utils.encode_range(range);
  }

  const S_REVISADO_FAIL = {
    font:  { bold: true, color: { rgb: "7B241C" }, sz: 10 },
    fill:  { patternType: "solid", fgColor: { rgb: "FADBD8" } },
    alignment: { wrapText: true },
  };
  const S_REVISADO_OK = {
    font:  { bold: true, color: { rgb: "145A32" }, sz: 10 },
    fill:  { patternType: "solid", fgColor: { rgb: "D5F5E3" } },
    alignment: { horizontal: "center" },
  };

  if (dsWsOut && revCol >= 0 && dsRows) {
    for (const dsRow of dsRows) {
      if (typeof dsRow._rowI !== "number") continue;
      if ((parseFloat(dsRow["CantidadUMComercial"]) || 0) === 0) continue; // saltar filas sin cantidad
      const addr = XLSX.utils.encode_cell({ r: dsRow._rowI, c: revCol });
      let reason = mismatchReasons?.get(dsRow._dsIdx);
      if (reason) {
        // Fila DS sin match o con discrepancia: escribir motivo en rojo
        dsWsOut[addr] = { t: "s", v: reason, s: S_REVISADO_FAIL };
      } else {
        // Fila DS sí fue usada: marcar OK en verde (solo si estaba vacía)
        const existing = dsWsOut[addr];
        if (!existing || !String(existing.v ?? "").trim()) {
          dsWsOut[addr] = { t: "s", v: "OK  Secuencia verificada/asignada", s: S_REVISADO_OK };
        }
      }
    }
    // Ajustar ancho de columna REVISADO
    if (!dsWsOut["!cols"]) dsWsOut["!cols"] = [];
    dsWsOut["!cols"][revCol] = { wch: 80 };
  }

  //  Construir libro de salida 
  const wb = XLSX.utils.book_new();

  // Añadir la hoja Layout modificada (referencia al ws ya modificado arriba)
  if (ws) {
    if (!ws["!cols"]) ws["!cols"] = [];
    if (colIdx.sec   >= 0) ws["!cols"][colIdx.sec]   = { wch: 14 };
    if (colIdx.notas >= 0) ws["!cols"][colIdx.notas] = { wch: 80 };
    if (cadenaCol >= 0) ws["!cols"][cadenaCol] = { wch: 55 };
    if (notasModifCol >= 0) ws["!cols"][notasModifCol] = { wch: 50 };
    const safeLayName = (layoutSheetName.slice(0, 17) + " (Actualiz.)").slice(0, 31);
    XLSX.utils.book_append_sheet(wb, ws, safeLayName);
  }

  // Añadir hoja DS sin cambios
  if (workbook.Sheets[dsSheetName]) {
    try { XLSX.utils.book_append_sheet(wb, workbook.Sheets[dsSheetName], dsSheetName.slice(0,31)); } catch(_){}
  }

  //  Hoja Reporte_2020 
  const reportRows2020 = [
    ["REPORTE MDULO 2020  Verificación y Asignación de Secuencias"],
    [],
    ["Hoja DS usada", dsSheetName || "DS *"],
    ["Hoja Layout usada", layoutSheetName],
    [],
    ["LEYENDA"],
    ["Verde en SEC CALC", "Secuencia verificada y correcta (ya estaba bien)"],
    ["Rojo en SEC CALC",  "Secuencia asignada o corregida por la app"],
    [],
    ["LEYENDA DE COLORES"],
    ["Verde (celda SEC CALC)", "Secuencia verificada  coincide con DS"],
    ["Rojo (celda SEC CALC)", "Secuencia asignada  Cant±1, Val±4 = DS (sin modificar país/fracción)"],
    ["Amarillo (celda)", "Fila repetida  mismo Ped+Frac+Pais+Cant+Val que otra(s)"],
    [],
    ["DETALLE POR FILA"],
    ["Pedimento","FraccionNico","PaisOrigen","Cantidad","ValorUSD","SEC CALC anterior","SEC CALC nuevo","Estado","Notas / Razón"],
  ];
  for (const row of layoutRows) {
    const a = assignment.get(row._idx);
    reportRows2020.push([
      row.Pedimento, row.FraccionNico, row.PaisOrigen,
      row.Cantidad, row.ValorUSD,
      row.SecCalc, a?.newSec ?? "",
      a?.status ?? "", a?.reason ?? "",
    ]);
  }
  const wsReport2020 = XLSX.utils.aoa_to_sheet(reportRows2020);
  wsReport2020["!cols"] = [22,14,8,12,12,14,14,12,80].map(w => ({ wch: w }));
  XLSX.utils.book_append_sheet(wb, wsReport2020, "Reporte_2020");

  return wb;
}

//  COMPONENTE APP2020 
function App2020() {
  const [phase2020, setPhase2020]     = useState("upload");
  const [isDragging2020, setIsDrag2020] = useState(false);
  const [error2020, setError2020]     = useState(null);
  const [fileName2020, setFileName2020] = useState("");
  const [results2020, setResults2020] = useState(null);
  const [outputWb2020, setOutputWb2020] = useState(null);
  const [progress2020, setProgress2020] = useState(0);
  const [tableData2020, setTableData2020] = useState(null);
  const [filterPed2020, setFilterPed2020] = useState("TODOS");
  const [copiedMsg, setCopiedMsg] = useState("");
  const [showInstrucciones2020, setShowInstrucciones2020] = useState(false);
  const [showUnusedDS2020, setShowUnusedDS2020] = useState(false);
  const [unusedPasteMsg2020, setUnusedPasteMsg2020] = useState("");
  const inputRef2020 = useRef(null);

  const process2020 = useCallback(async (file) => {
    setError2020(null);
    setFileName2020(file.name);
    setPhase2020("processing");
    setProgress2020(0);
    try {
      const buf = await file.arrayBuffer();
      setProgress2020(20);
      const wb = XLSX.read(buf, { type: "array" });
      setProgress2020(30);

      const { dsName, layName } = resolveDS2020SheetNames(wb);
      if (!dsName)  throw new Error('No se encontró hoja DS (debe contener "DS" en el nombre)');
      if (!layName) throw new Error('No se encontró hoja Layout (debe tener columnas de Layout: pedimento, FraccionNico, SEC CALC, etc.)');

      setProgress2020(40);
      const dsRows      = readDS2020Sheet(wb.Sheets[dsName]);
      const layout2020  = readLayout2020Sheet(wb.Sheets[layName]);
      layout2020.dsRows = dsRows; // pasar ref al DS para que buildOutput pueda leer _rowI y _revisadoCol

      // Avisar si faltan columnas: indicar cuáles no se encontraron y cómo deben nombrarse
      const missingDS = dsRows._missingColumns || [];
      const missingLy = layout2020._missingColumns || [];
      if (missingDS.length > 0 || missingLy.length > 0) {
        const parts = [];
        if (missingDS.length > 0) {
          parts.push("Hoja DS  columnas no encontradas:\n" + missingDS.map(m =>
            ` ${m.name}\n  Nómbrela así (o use un alias aceptado): ${m.aceptados.join(", ")}`
          ).join("\n\n"));
        }
        if (missingLy.length > 0) {
          parts.push("Hoja Layout  columnas no encontradas:\n" + missingLy.map(m =>
            ` ${m.name}\n  Nómbrela así (o use un alias aceptado): ${m.aceptados.join(", ")}`
          ).join("\n\n"));
        }
        throw new Error(parts.join("\n\n") + "\n\nConsulte «Instrucciones» (botón arriba) para la lista completa de nombres requeridos.");
      }

      setProgress2020(60);

      const pedMismatch = checkPedimentoMismatch(
        getPedimentosFromRows(dsRows, "Pedimento2"),
        getPedimentosFromRows(layout2020.layoutRows, "Pedimento", "pedimento")
      );

      const { assignment, stats, unusedDS, unusedDSDetails, mismatchReasons, globalTotals: gt2020 } = runCascade2020(layout2020.layoutRows, dsRows);
      setProgress2020(80);

      const newWb = buildOutput2020Excel(wb, layName, dsName, layout2020, assignment, mismatchReasons);
      setProgress2020(100);

      // Construir datos de tabla para vista in-app
      const nDescT = s => String(s ?? "").trim().toLowerCase().replace(/\s+/g, " ");
      const normT  = s => String(s ?? "").trim().toUpperCase();
      // Cadena Layout vs DS (mismo criterio que Excel: verde si coinciden, naranja si no)
      const normStrCad = (v) => String(v ?? "").trim();
      const nFracCad   = (v) => normStrCad(v).replace(/^0+/, "") || "0";
      const nDescCad   = (s) => String(s ?? "").trim().toLowerCase().replace(/\s*\/\s*/g, "/").replace(/\s+/g, " ");
      const getDSFracCad = (dsRow) => {
        const f = nFracCad(dsRow["Fraccion"]);
        if (f) return f;
        const candado = normStrCad(dsRow["Candado 551"] ?? "");
        if (!candado) return "";
        const parts = candado.split("-");
        for (let i = 0; i < parts.length; i++) {
          const p = parts[i].replace(/\./g, "");
          if (p.length === 8 && /^\d+$/.test(p)) return p;
        }
        return "";
      };
      const buildCadenaLayoutRow = (row) => {
        const ped = normStrCad(row.Pedimento);
        const frac = nFracCad(row.FraccionNico ?? "");
        const desc = nDescCad(row.Descripcion ?? "");
        const pais = normStrCad(row.PaisOrigen ?? "");
        const cant = Math.round(parseFloat(row.Cantidad) || 0);
        const val  = Math.round(parseFloat(row.ValorUSD) || 0);
        return `${ped}|${frac}|${desc}|${pais}|${cant}|${val}`;
      };
      const buildCadenaDSRow = (dsRow) => {
        const ped  = normStrCad(dsRow["Pedimento2"] ?? "");
        const frac = getDSFracCad(dsRow);
        const desc = nDescCad(dsRow["DescripcionMercancia"] ?? "");
        const pais = normStrCad(dsRow["PaisOrigenDestino"] ?? "");
        const cant = Math.round(parseFloat(dsRow["CantidadUMComercial"]) || 0);
        const val  = Math.round(parseFloat(dsRow["ValorDolares"]) || 0);
        return `${ped}|${frac}|${desc}|${pais}|${cant}|${val}`;
      };

      // Helper: texto de notas de modificaciones (correcciones) o "OK"
      const buildNotasModificaciones = (layoutRow, a, ds) => {
        if (!a) return "Sin evaluación de secuencia para esta fila.";
        if (a.status === "unmatched" || !ds) {
          return `Sin match (DS sin cadena para esta fila): ${a.reason || "No se encontró candidato en DS con pedimento/fracción/cantidad/valor compatibles."}`;
        }
        const partes = [];
        const paisLy  = normT(layoutRow.PaisOrigen || "");
        const paisDs  = normT(ds["PaisOrigenDestino"] || "");
        if (paisLy && paisDs && paisLy !== paisDs) partes.push(`Se corrigió país ${paisLy} a ${paisDs}`);
        const descLy  = String(layoutRow.Descripcion || "").trim();
        const descDs  = String(ds["DescripcionMercancia"] || "").trim();
        if (descLy !== descDs && descDs) partes.push(`Se corrigió descripción "${descLy}" a "${descDs}"`);
        if (a.fracCorr && a.fracOrig) partes.push(`Se corrigió fracción ${a.fracOrig} a ${a.fracCorr}`);
        if (a.secOrig && a.newSec && String(a.secOrig) !== String(a.newSec)) partes.push(`Se corrigió secuencia ${a.secOrig} a ${a.newSec}`);
        return partes.length ? partes.join("; ") : "OK";
      };

      // Paso 1: construir filas base con datos DS
      const tRowsBase = layout2020.layoutRows.map(r => {
        const a   = assignment.get(r._idx);
        const ds  = a?.dsRow || null;
        const pais = normT(r.PaisOrigen || r["Pais Origen"] || "");
        const desc = String(r.Descripcion || r["DescripcionMercancia"] || "");
        const cadenaLay = buildCadenaLayoutRow(r);
        const cadenaDs  = ds ? buildCadenaDSRow(ds) : "";
        const cadenaCoincide = !!ds && cadenaLay === cadenaDs;
        const notasModificaciones = buildNotasModificaciones(r, a, ds);
        return {
          idx:        r._idx,
          ped:        String(r.Pedimento  || ""),
          frac:       String(r.FraccionNico || ""),
          pais,
          desc,
          cant:       r.Cantidad || 0,
          val:        r.ValorUSD  || 0,
          secOrig:    String(r.SecCalc || ""),
          secNueva:   a?.newSec  || "",
          secPrevCorr: a?.secOrig || null,   // sec original que fue reemplazada
          status:     a?.status  || "unmatched",
          estrategia: a?.estrategia || "",
          reason:     a?.reason   || "Sin match",
          fracCorr:   a?.fracCorr || null,   // fracción corregida (cross-fraction)
          fracOrig:   a?.fracOrig || null,   // fracción original en Layout
          notasModificaciones,
          // Cadena Layout vs DS (verde si coinciden, naranja si no)
          cadenaLayout:   cadenaLay,
          cadenaDS:       cadenaDs,
          cadenaCoincide,
          // Datos del DS para comparación
          dsCant:     ds ? (parseFloat(ds["CantidadUMComercial"]) || 0) : null,
          dsVal:      ds ? (parseFloat(ds["ValorDolares"])        || 0) : null,
          dsPais:     ds ? normT(ds["PaisOrigenDestino"] || "") : null,
          dsDesc:     ds ? String(ds["DescripcionMercancia"] || "") : null,
          dsFrac:     ds ? String(ds["Fraccion"] || "") : null,
          // Clave de grupo para sumar cant/val de todas las filas con la misma secuencia asignada
          groupKey:   a?.newSec ? `${r.Pedimento}||${r.FraccionNico}||${a.newSec}||${ds?._dsIdx ?? ""}` : null,
        };
      });

      // Paso 2: calcular sumas por grupo y secuencias involucradas en cada grupo
      const groupSums = new Map();
      const groupSecuencias = new Map(); // groupKey  [secNueva, secNueva, ...]
      for (const r of tRowsBase) {
        if (!r.groupKey) continue;
        if (!groupSums.has(r.groupKey)) groupSums.set(r.groupKey, { sumCant: 0, sumVal: 0 });
        const g = groupSums.get(r.groupKey);
        g.sumCant += r.cant;
        g.sumVal  += r.val;
        if (!groupSecuencias.has(r.groupKey)) groupSecuencias.set(r.groupKey, []);
        groupSecuencias.get(r.groupKey).push(String(r.secNueva || "").trim() || "");
      }

      // Paso 3: añadir totales de grupo, cadena total (grupo sumado vs DS) y secuencias del grupo
      const tRows = tRowsBase.map(r => {
        const groupSumCant = r.groupKey ? (groupSums.get(r.groupKey)?.sumCant ?? null) : null;
        const groupSumVal  = r.groupKey ? (groupSums.get(r.groupKey)?.sumVal  ?? null) : null;
        const cadenaTotalLayout = (groupSumCant != null && groupSumVal != null)
          ? `${normStrCad(r.ped)}|${nFracCad(r.frac)}|${nDescCad(r.desc)}|${normStrCad(r.pais)}|${Math.round(groupSumCant)}|${Math.round(groupSumVal)}`
          : "";
        const cadenaTotalCoincide = (r.dsCant != null && r.dsVal != null && groupSumCant != null && groupSumVal != null)
          && (Math.round(groupSumCant) === Math.round(r.dsCant) && Math.round(groupSumVal) === Math.round(r.dsVal));
        const secuenciasEnGrupo = r.groupKey ? (groupSecuencias.get(r.groupKey) || []) : [];
        const diffCantGrupo = (r.dsCant != null && groupSumCant != null) ? (groupSumCant - r.dsCant) : null;
        const diffValGrupo  = (r.dsVal  != null && groupSumVal  != null) ? (groupSumVal  - r.dsVal)  : null;
        const absDiffCant = diffCantGrupo == null ? null : Math.abs(diffCantGrupo);
        const absDiffVal  = diffValGrupo  == null ? null : Math.abs(diffValGrupo);
        const strategy = r.estrategia || "";
        const isForced = ["B6_relleno","A_forzado","R2","R3","R4","R5","A3b","A2b","A1b"].includes(strategy) || r.status === "new";
        const fueraRegla = (absDiffCant != null && absDiffCant > 1) || (absDiffVal != null && absDiffVal > 4);
        const motivoDetallado = (() => {
          if (r.status === "unmatched") {
            return `Sin match: ${r.reason || "No se encontró secuencia con reglas del DS."}`;
          }
          const partes = [];
          partes.push(`Estrategia: ${strategy || "N/A"}`);
          if (!r.cadenaDS) partes.push("DS sin cadena para esta fila (no hubo renglón DS equivalente para comparación 1:1).");
          if (isForced) partes.push("Asignación de respaldo/forzada para no dejar fila sin secuencia.");
          if (fueraRegla) {
            partes.push(`Fuera de regla: �Cant=${diffCantGrupo == null ? "N/A" : diffCantGrupo.toLocaleString("es-MX")} (tol ±1), �USD=${diffValGrupo == null ? "N/A" : diffValGrupo.toFixed(2)} (tol ±4).`);
          } else {
            partes.push("Dentro de reglas de tolerancia de comparación por grupo.");
          }
          if (r.reason && r.reason !== "Sin match") partes.push(`Motor: ${r.reason}`);
          return partes.join(" | ");
        })();
        return {
          ...r,
          groupSumCant,
          groupSumVal,
          cadenaTotalLayout,
          cadenaTotalCoincide: !!cadenaTotalCoincide,
          secuenciasEnGrupo,
          diffCantGrupo,
          diffValGrupo,
          isForced,
          fueraRegla,
          motivoDetallado,
        };
      });

      const secuenciasCompletasAsignadas = (() => {
        const seen = new Set();
        for (const r of tRows) {
          if (r.groupKey && r.status !== "unmatched" && r.cadenaTotalCoincide) {
            seen.add(r.groupKey);
          }
        }
        return seen.size;
      })();

      setTableData2020(tRows);
      setFilterPed2020("TODOS");

      setResults2020({
        stats,
        unusedDSCount: unusedDS.length,
        unusedDSDetails: unusedDSDetails || [],
        total: layout2020.layoutRows.length,
        secuenciasCompletasAsignadas,
        dsName,
        layName,
        pedMismatch,
        globalTotals: gt2020
      });
      setOutputWb2020(newWb);
      setTimeout(() => setPhase2020("results"), 400);
    } catch (e) {
      setError2020(e.message);
      setPhase2020("upload");
    }
  }, []);

  const onFile2020 = useCallback((file) => {
    if (!file?.name?.match(/\.(xlsx|xls)$/i)) { setError2020("Solo archivos Excel (.xlsx / .xls)"); return; }
    process2020(file);
  }, [process2020]);

  const download2020 = () => {
    if (!outputWb2020) return;
    const out = XLSX.write(outputWb2020, { bookType: "xlsx", type: "array" });
    const blob = new Blob([out], { type: "application/octet-stream" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a"); a.href = url;
    a.download = fileName2020.replace(/\.xlsx?$/i, "") + "_2020_secuencias.xlsx";
    a.click(); URL.revokeObjectURL(url);
  };

  const reset2020 = () => { setPhase2020("upload"); setResults2020(null); setOutputWb2020(null); setError2020(null); setProgress2020(0); setTableData2020(null); setFilterPed2020("TODOS"); setCopiedMsg(""); setShowUnusedDS2020(false); setUnusedPasteMsg2020(""); };

  return (
    <div>
      {/* Botones header */}
      {phase2020 === "results" && (
        <div style={{ display:"flex", gap:8, marginBottom:24 }}>
          <button onClick={reset2020} style={{ background:"transparent",border:"1px solid #334155",color:"#94a3b8",padding:"8px 16px",cursor:"pointer",borderRadius:4,fontSize:13 }}> Nuevo archivo</button>
          <button onClick={download2020} style={{ background:"#22c55e",border:"none",color:"#0f172a",padding:"8px 20px",cursor:"pointer",borderRadius:4,fontSize:13,fontWeight:800 }}> Descargar Excel 2020</button>
        </div>
      )}

      {phase2020 === "upload" && (
        <div style={{ animation:"fadeUp 0.4s ease" }}>
          <div style={{ textAlign:"center", marginBottom:40 }}>
            <div style={{ display:"inline-block",background:"rgba(34,197,94,0.1)",border:"1px solid rgba(34,197,94,0.2)",color:"#22c55e",padding:"4px 14px",borderRadius:20,fontSize:11,letterSpacing:"0.12em",fontFamily:"DM Mono, monospace",marginBottom:16 }}>
              MULTI-PEDIMENTO · VERIFICACIN + ASIGNACIN
            </div>
            <h2 style={{ fontSize:32,fontWeight:900,margin:"0 0 12px",letterSpacing:"-0.02em" }}>
              Módulo <span style={{color:"#22c55e"}}>DS</span>  Secuencias Multi-Pedimento
            </h2>
            <p style={{ color:"#64748b",fontSize:14,maxWidth:520,margin:"0 auto" }}>
              Sube un Excel con hojas <b style={{color:"#22c55e"}}>DS *</b> (Data Stage) y <b style={{color:"#22c55e"}}>Layout *</b>.
              La app verifica secuencias existentes y asigna las faltantes por pedimento.
            </p>
            <button
              type="button"
              onClick={() => setShowInstrucciones2020(!showInstrucciones2020)}
              style={{ marginTop:12, background:"transparent", border:"1px solid #334155", color:"#94a3b8", padding:"8px 16px", borderRadius:6, cursor:"pointer", fontSize:13, fontWeight:600 }}
            >
              {showInstrucciones2020 ? " Ocultar instrucciones" : " Instrucciones  requisitos del Excel"}
            </button>
          </div>

          {showInstrucciones2020 && (
            <div style={{ marginBottom:24, padding:"20px 24px", background:"#0f172a", border:"1px solid #334155", borderRadius:8, textAlign:"left", maxWidth:720 }}>
              <div style={{ color:"#22c55e", fontSize:14, fontWeight:800, marginBottom:16, letterSpacing:"0.05em" }}>REQUISITOS DEL ARCHIVO EXCEL</div>
              <p style={{ color:"#94a3b8", fontSize:13, marginBottom:16 }}>El archivo debe tener <strong style={{color:"#f8fafc"}}>dos hojas</strong> con los nombres y columnas indicados. Use exactamente estos nombres de columnas (o los alternativos aceptados) para que el sistema lea correctamente los valores.</p>

              <div style={{ marginBottom:20 }}>
                <div style={{ color:"#f8fafc", fontSize:14, fontWeight:700, marginBottom:8 }}>1. Hoja DS (Data Stage)</div>
                <p style={{ color:"#64748b", fontSize:12, marginBottom:8 }}>Nombre de la hoja: debe contener <strong>"DS"</strong> (ej. <code style={{background:"#1e293b",padding:"2px 6px",borderRadius:4}}>DS</code>).</p>
                <div style={{ fontSize:12, color:"#cbd5e1" }}>
                  <table style={{ width:"100%", borderCollapse:"collapse", fontSize:11 }}>
                    <thead><tr style={{ borderBottom:"1px solid #334155" }}><th style={{ padding:"6px 8px", textAlign:"left", color:"#64748b" }}>Columna (recomendado)</th><th style={{ padding:"6px 8px", textAlign:"left", color:"#64748b" }}>Nombres aceptados</th></tr></thead>
                    <tbody>
                      <tr style={{ borderBottom:"1px solid #1e293b" }}><td style={{ padding:"6px 8px", fontFamily:"monospace" }}>Pedimento2</td><td style={{ padding:"6px 8px" }}>Pedimento2</td></tr>
                      <tr style={{ borderBottom:"1px solid #1e293b" }}><td style={{ padding:"6px 8px", fontFamily:"monospace" }}>Fraccion</td><td style={{ padding:"6px 8px" }}>Fraccion</td></tr>
                      <tr style={{ borderBottom:"1px solid #1e293b" }}><td style={{ padding:"6px 8px", fontFamily:"monospace" }}>SecuenciaFraccion</td><td style={{ padding:"6px 8px" }}>SecuenciaFraccion</td></tr>
                      <tr style={{ borderBottom:"1px solid #1e293b" }}><td style={{ padding:"6px 8px", fontFamily:"monospace" }}>SubdivisionFraccion (NICO)</td><td style={{ padding:"6px 8px" }}>SubdivisionFraccion, Subdivision Fraccion, NICO</td></tr>
                      <tr style={{ borderBottom:"1px solid #1e293b" }}><td style={{ padding:"6px 8px", fontFamily:"monospace" }}>DescripcionMercancia</td><td style={{ padding:"6px 8px" }}>DescripcionMercancia</td></tr>
                      <tr style={{ borderBottom:"1px solid #1e293b" }}><td style={{ padding:"6px 8px", fontFamily:"monospace" }}>CantidadUMComercial</td><td style={{ padding:"6px 8px" }}>CantidadUMComercial</td></tr>
                      <tr style={{ borderBottom:"1px solid #1e293b" }}><td style={{ padding:"6px 8px", fontFamily:"monospace" }}>ValorDolares</td><td style={{ padding:"6px 8px" }}>ValorDolares, Valor usd redondeado, ValorAduana, Valor Aduana Estadístico, ValorAgregado</td></tr>
                      <tr style={{ borderBottom:"1px solid #1e293b" }}><td style={{ padding:"6px 8px", fontFamily:"monospace" }}>PaisOrigenDestino</td><td style={{ padding:"6px 8px" }}>PaisOrigenDestino</td></tr>
                      <tr style={{ borderBottom:"1px solid #1e293b" }}><td style={{ padding:"6px 8px", fontFamily:"monospace" }}>Candado 551</td><td style={{ padding:"6px 8px" }}>Candado 551, Candado DS 551, Candado, Clave, Secuencias</td></tr>
                    </tbody>
                  </table>
                </div>
              </div>

              <div>
                <div style={{ color:"#f8fafc", fontSize:14, fontWeight:700, marginBottom:8 }}>2. Hoja Layout</div>
                <p style={{ color:"#64748b", fontSize:12, marginBottom:8 }}>Nombre de la hoja: debe contener <strong>"Layout"</strong> o tener columnas típicas de Layout (Pedimento, FraccionNico, SEC CALC, etc.).</p>
                <div style={{ fontSize:12, color:"#cbd5e1" }}>
                  <table style={{ width:"100%", borderCollapse:"collapse", fontSize:11 }}>
                    <thead><tr style={{ borderBottom:"1px solid #334155" }}><th style={{ padding:"6px 8px", textAlign:"left", color:"#64748b" }}>Columna (recomendado)</th><th style={{ padding:"6px 8px", textAlign:"left", color:"#64748b" }}>Nombres aceptados</th></tr></thead>
                    <tbody>
                      <tr style={{ borderBottom:"1px solid #1e293b" }}><td style={{ padding:"6px 8px", fontFamily:"monospace" }}>Pedimento</td><td style={{ padding:"6px 8px" }}>Pedimento</td></tr>
                      <tr style={{ borderBottom:"1px solid #1e293b" }}><td style={{ padding:"6px 8px", fontFamily:"monospace" }}>FraccionNico</td><td style={{ padding:"6px 8px" }}>FraccionNico, fraccionnico</td></tr>
                      <tr style={{ borderBottom:"1px solid #1e293b" }}><td style={{ padding:"6px 8px", fontFamily:"monospace" }}>NICO (opcional recomendado)</td><td style={{ padding:"6px 8px" }}>NICO, nico, SubdivisionFraccion, subdivisionfraccion</td></tr>
                      <tr style={{ borderBottom:"1px solid #1e293b" }}><td style={{ padding:"6px 8px", fontFamily:"monospace" }}>SEC CALC / Secuencia</td><td style={{ padding:"6px 8px" }}>SEC CALC, seccalc, secuenciaped, Secuencia, secuencia, SEC</td></tr>
                      <tr style={{ borderBottom:"1px solid #1e293b" }}><td style={{ padding:"6px 8px", fontFamily:"monospace" }}>Descripcion</td><td style={{ padding:"6px 8px" }}>descripcion, clase_descripcion, descripcionmercancia</td></tr>
                      <tr style={{ borderBottom:"1px solid #1e293b" }}><td style={{ padding:"6px 8px", fontFamily:"monospace" }}>PaisOrigen</td><td style={{ padding:"6px 8px" }}>pais_origen, paisorigen, paisorigendestino</td></tr>
                      <tr style={{ borderBottom:"1px solid #1e293b" }}><td style={{ padding:"6px 8px", fontFamily:"monospace" }}>Cantidad</td><td style={{ padding:"6px 8px" }}>cantidad_comercial, cantidadcomercial, cantidadumc</td></tr>
                      <tr style={{ borderBottom:"1px solid #1e293b" }}><td style={{ padding:"6px 8px", fontFamily:"monospace" }}>Valor USD</td><td style={{ padding:"6px 8px" }}>ValorMPDolares, valormpdolares, valordolares, valor_me, valorme</td></tr>
                      <tr style={{ borderBottom:"1px solid #1e293b" }}><td style={{ padding:"6px 8px", fontFamily:"monospace" }}>NOTAS</td><td style={{ padding:"6px 8px" }}>NOTAS, notas</td></tr>
                      <tr style={{ borderBottom:"1px solid #1e293b" }}><td style={{ padding:"6px 8px", fontFamily:"monospace" }}>Candado 551</td><td style={{ padding:"6px 8px" }}>CANDADO 551, Candado 551, Candado DS 551, Candado, Clave, candado</td></tr>
                    </tbody>
                  </table>
                </div>
              </div>

              <p style={{ color:"#64748b", fontSize:11, marginTop:16 }}>La primera fila con la mayoría de estas columnas se toma como encabezado. Homologue los nombres en su Excel a uno de los aceptados para evitar errores de lectura.</p>
            </div>
          )}

          {error2020 && (
            <div style={{ background:"rgba(239,68,68,0.1)",border:"1px solid #ef4444",borderRadius:4,padding:"12px 18px",marginBottom:20,color:"#fca5a5",fontSize:13,whiteSpace:"pre-wrap",textAlign:"left" }}>
               {error2020}
            </div>
          )}

          <div
            onClick={() => inputRef2020.current?.click()}
            onDragOver={e => { e.preventDefault(); setIsDrag2020(true); }}
            onDragLeave={() => setIsDrag2020(false)}
            onDrop={e => { e.preventDefault(); setIsDrag2020(false); const f = e.dataTransfer.files[0]; if(f) onFile2020(f); }}
            style={{ border:`2px dashed ${isDragging2020?"#22c55e":"#334155"}`,borderRadius:8,padding:"48px 32px",textAlign:"center",cursor:"pointer",background:isDragging2020?"rgba(34,197,94,0.05)":"transparent",transition:"all 0.2s" }}
          >
            <input ref={inputRef2020} type="file" accept=".xlsx,.xls" style={{display:"none"}} onChange={e => e.target.files[0] && onFile2020(e.target.files[0])} />
            <div style={{fontSize:40,marginBottom:12}}></div>
            <div style={{color:"#f8fafc",fontSize:16,fontWeight:700,marginBottom:8}}>Sube tu archivo Excel 2020</div>
            <div style={{color:"#94a3b8",fontSize:12}}>Requiere hojas <span style={{color:"#22c55e",fontFamily:"monospace"}}>DS</span> y <span style={{color:"#22c55e",fontFamily:"monospace"}}>Layout</span></div>
          </div>

          <div style={{marginTop:28,padding:"18px 20px",background:"rgba(34,197,94,0.05)",border:"1px solid rgba(34,197,94,0.15)",borderRadius:6}}>
            <div style={{color:"#22c55e",fontSize:12,fontWeight:700,marginBottom:10,letterSpacing:"0.08em"}}>LEYENDA DE COLORES EN EL EXCEL DE SALIDA</div>
            {[["Verde en SEC CALC","Secuencia existente VERIFICADA  coincide con DS"],["Rojo en SEC CALC","Secuencia NUEVA asignada o CORREGIDA (Cant±1, Val±4 = DS)"],["Naranja en SEC CALC ","Secuencia existente NO verificada  revisar manualmente"],["Amarillo en celda","Fila REPETIDA  mismo Ped+Frac+Pais+Cant+Val (se conservan todas)"],["Morado en celda","Fracción CORREGIDA  Layout tenía fracción diferente al DS (cross-fraction)"]].map(([c,d]) => (
              <div key={c} style={{display:"flex",gap:10,marginBottom:6,fontSize:12}}>
                <span style={{color:"#22c55e",fontWeight:700,minWidth:180}}>{c}</span>
                <span style={{color:"#64748b"}}>{d}</span>
              </div>
            ))}
          </div>
        </div>
      )}

      {phase2020 === "processing" && (
        <div style={{textAlign:"center",padding:"80px 0",animation:"fadeUp 0.3s ease"}}>
          <div style={{width:48,height:48,border:"3px solid #22c55e",borderTopColor:"transparent",borderRadius:"50%",animation:"spin 0.8s linear infinite",margin:"0 auto 24px"}} />
          <div style={{color:"#f8fafc",fontSize:18,fontWeight:700,marginBottom:8}}>Procesando DS...</div>
          <div style={{color:"#64748b",fontSize:13,marginBottom:24}}>{fileName2020}</div>
          <div style={{width:280,height:4,background:"#1e293b",borderRadius:2,margin:"0 auto",overflow:"hidden"}}>
            <div style={{height:"100%",background:"#22c55e",width:`${progress2020}%`,transition:"width 0.4s ease",borderRadius:2}} />
          </div>
        </div>
      )}

      {phase2020 === "results" && results2020 && (
        <div style={{animation:"fadeUp 0.4s ease"}}>
          <div style={{color:"#94a3b8",fontSize:12,marginBottom:16,fontFamily:"monospace"}}>
            DS: <span style={{color:"#22c55e"}}>{results2020.dsName}</span> · Layout: <span style={{color:"#22c55e"}}>{results2020.layName}</span>
          </div>
          {results2020.pedMismatch && (
            <div style={{background:"rgba(245,158,11,0.12)",border:"1px solid #f59e0b",borderRadius:6,padding:"14px 18px",marginBottom:24,fontSize:13,color:"#fcd34d"}}>
              <strong> Pedimentos distintos entre DS y Layout</strong>
              <p style={{margin:"8px 0 0",color:"#fde68a",lineHeight:1.5}}>
                El DS tiene pedimentos que no aparecen en el Layout (y viceversa). Por eso no hay match.
                <br />DS: <code style={{background:"rgba(0,0,0,0.2)",padding:"2px 6px",borderRadius:3}}>{results2020.pedMismatch.ds.join(", ")}</code>
                <br />Layout: <code style={{background:"rgba(0,0,0,0.2)",padding:"2px 6px",borderRadius:3}}>{results2020.pedMismatch.layout.join(", ")}</code>
                <br /><span style={{fontSize:12,opacity:0.9}}>Verifica que ambas hojas correspondan al mismo pedimento o incluyan los mismos pedimentos.</span>
              </p>
            </div>
          )}
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12,marginBottom:32}}>
            {[
              { label:"Verificadas (OK)", value: results2020.stats.verified,    accent:"#22c55e",  sub:"Verde en Excel" },
              { label:"Corregidas",        value: results2020.stats.corrected,   accent:"#ef4444",  sub:"Rojo  cambio aplicado" },
              { label:"Nuevas asignadas",  value: results2020.stats.newAssigned, accent:"#f59e0b",  sub:"Rojo  asignación nueva" },
              { label:"Sin match",         value: results2020.stats.unmatched,   accent:"#64748b",  sub:"Revisar manualmente" },
            ].map(c => <StatCard key={c.label} label={c.label} value={c.value} sub={c.sub} accent={c.accent} />)}
          </div>
          {results2020.globalTotals && (() => {
            const gt    = results2020.globalTotals;
            const cantOk = gt.cuadra;
            const valOk  = gt.cuadraVal ?? (Math.abs(gt.lyVal - gt.dsVal) <= 5);
            const fmt  = n => Number(n).toLocaleString("es-MX", {maximumFractionDigits:0});
            const fmtV = n => Number(n).toLocaleString("es-MX", {maximumFractionDigits:2});
            const diffC = gt.lyCant - gt.dsCant;
            const diffV = gt.lyVal  - gt.dsVal;
            return (
              <div style={{borderRadius:6, padding:"16px 20px", marginBottom:16, fontSize:13,
                background: cantOk ? "rgba(34,197,94,0.08)" : "rgba(239,68,68,0.12)",
                border:`1px solid ${cantOk ? "rgba(34,197,94,0.3)" : "#ef4444"}`}}>
                {/* Título cantidad */}
                <div style={{fontWeight:700, marginBottom:6, fontSize:14,
                  color: cantOk ? "#22c55e" : "#ef4444"}}>
                  {cantOk ? " Cantidad global coincide  todas las filas serán asignadas"
                           : " Cantidad global NO coincide  habrá filas sin asignar"}
                </div>
                {/* Subtítulo valor */}
                <div style={{marginBottom:12, fontSize:12,
                  color: valOk ? "#86efac" : "#fbbf24"}}>
                  {valOk ? " Valor USD global también coincide"
                          : ` Valor USD difiere $${Math.abs(diffV).toFixed(2)}  aceptable por redondeos, las secuencias se asignarán por cantidad`}
                </div>
                <div style={{display:"grid", gridTemplateColumns:"1fr 1fr 1fr 1fr", gap:12}}>
                  {[
                    {label:"DS  Cantidad",      val: fmt(gt.dsCant),        color:"#94a3b8"},
                    {label:"Layout  Cantidad",  val: fmt(gt.lyCant),        color: cantOk ? "#22c55e" : "#ef4444"},
                    {label:"DS  Valor USD",     val: `$${fmtV(gt.dsVal)}`,  color:"#94a3b8"},
                    {label:"Layout  Valor USD", val: `$${fmtV(gt.lyVal)}`,  color: valOk ? "#22c55e" : "#fbbf24"},
                  ].map(item=>(
                    <div key={item.label}>
                      <div style={{color:"#475569",fontSize:10,fontFamily:"DM Mono, monospace",marginBottom:4}}>{item.label}</div>
                      <div style={{color:item.color,fontFamily:"DM Mono, monospace",fontSize:14,fontWeight:700}}>{item.val}</div>
                    </div>
                  ))}
                </div>
                {!cantOk && (
                  <div style={{marginTop:10,color:"#fca5a5",fontSize:12}}>
                    Diferencia Cantidad: {diffC > 0 ? "+" : ""}{diffC.toLocaleString("es-MX")}
                    {" · "}Verifica que ambas hojas correspondan al mismo pedimento y no falten filas.
                  </div>
                )}
                {!valOk && cantOk && (
                  <div style={{marginTop:10,color:"#fde68a",fontSize:12}}>
                    Diferencia USD: {diffV > 0 ? "+" : ""}${diffV.toFixed(2)} global  cada secuencia se asigna por cantidad (±1); el valor puede tener pequeñas variaciones de redondeo.
                  </div>
                )}
              </div>
            );
          })()}
          <div style={{background:"rgba(34,197,94,0.08)",border:"1px solid rgba(34,197,94,0.2)",borderRadius:6,padding:"12px 20px",marginBottom:24,fontSize:13,color:"#94a3b8"}}>
            Total filas Layout: <b style={{color:"#f8fafc"}}>{results2020.total}</b> &nbsp;·&nbsp;
            Secuencias DS no usadas: <b style={{color: results2020.unusedDSCount>0?"#ef4444":"#22c55e"}}>{results2020.unusedDSCount}</b>
            &nbsp;·&nbsp;
            Secuencias completas asignadas:{" "}
            <b style={{color:"#22c55e"}} title="Grupos distintos donde la suma de cantidad y valor del Layout coincide con el renglón DS asignado (misma cadena agregada vs DS).">
              {results2020.secuenciasCompletasAsignadas ?? 0}
            </b>
          </div>

          {results2020.unusedDSCount > 0 && (
            <div style={{ marginBottom: 18 }}>
              <div
                onClick={() => setShowUnusedDS2020(!showUnusedDS2020)}
                style={{
                  cursor: "pointer",
                  background: "rgba(239,68,68,0.08)",
                  border: "1px solid rgba(239,68,68,0.25)",
                  borderRadius: 6,
                  padding: "10px 14px",
                  display: "flex",
                  justifyContent: "space-between",
                  alignItems: "center",
                  fontSize: 13,
                  color: "#fecaca"
                }}
              >
                <span><b>Secuencias DS no usadas:</b> detalle y razón probabilística</span>
                <span style={{ color: "#fda4af" }}>{showUnusedDS2020 ? "\u25B2" : "\u25BC"}</span>
              </div>
              {showUnusedDS2020 && (
                <>
                  <div
                    style={{
                      marginTop: 10,
                      display: "flex",
                      flexWrap: "wrap",
                      alignItems: "center",
                      gap: 10,
                      padding: "8px 10px",
                      background: "rgba(15,23,42,0.6)",
                      border: "1px solid #3f1d1d",
                      borderRadius: 6,
                      borderBottomLeftRadius: 0,
                      borderBottomRightRadius: 0,
                      borderBottom: "none"
                    }}
                  >
                    <button
                      type="button"
                      onClick={async () => {
                        const details = results2020?.unusedDSDetails || [];
                        const raw = details.map((u) => String(u.secuencia ?? "").trim()).filter(Boolean);
                        const nums = [...new Set(raw)];
                        nums.sort((a, b) => {
                          const na = parseFloat(String(a).replace(/,/g, ""));
                          const nb = parseFloat(String(b).replace(/,/g, ""));
                          if (Number.isFinite(na) && Number.isFinite(nb)) return na - nb;
                          return String(a).localeCompare(String(b), "es", { numeric: true });
                        });
                        const text = ["Secuencia", ...nums].join(String.fromCharCode(13, 10));
                        try {
                          await navigator.clipboard.writeText(text);
                          setUnusedPasteMsg2020(`${nums.length} secuencia(s) copiadas. Pega en Excel (columna A).`);
                          setTimeout(() => setUnusedPasteMsg2020(""), 4000);
                        } catch {
                          setUnusedPasteMsg2020("No se pudo copiar al portapapeles (permiso del navegador).");
                          setTimeout(() => setUnusedPasteMsg2020(""), 5000);
                        }
                      }}
                      style={{
                        background: "#22c55e",
                        color: "#0f172a",
                        border: "none",
                        borderRadius: 4,
                        padding: "8px 14px",
                        fontSize: 12,
                        fontWeight: 700,
                        cursor: "pointer"
                      }}
                    >
                      Copiar lista de secuencias para Excel
                    </button>
                    <span style={{ fontSize: 11, color: "#94a3b8", maxWidth: 420, lineHeight: 1.4 }}>
                      Una fila por número (encabezado &quot;Secuencia&quot;). Incluye todas las no usadas, no solo las visibles en la tabla.
                    </span>
                    {unusedPasteMsg2020 ? (
                      <span style={{ fontSize: 12, color: "#86efac", fontWeight: 600 }}>{unusedPasteMsg2020}</span>
                    ) : null}
                  </div>
                <div style={{ marginTop: 0, border: "1px solid #3f1d1d", borderRadius: 6, borderTopLeftRadius: 0, borderTopRightRadius: 0, overflowX: "auto", maxHeight: 320, overflowY: "auto" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
                    <thead style={{ position: "sticky", top: 0, background: "#2b1212" }}>
                      <tr>
                        {["Sec", "Pedimento", "Fracción", "País", "Cant DS", "Valor DS", "Razón", "Probable vínculo Layout"].map(h => (
                          <th key={h} style={{ textAlign: "left", padding: "8px 10px", color: "#fca5a5", borderBottom: "1px solid #3f1d1d", fontFamily: "DM Mono, monospace", fontSize: 10 }}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {(results2020.unusedDSDetails || []).slice(0, 250).map((u, i) => (
                        <tr key={`${u.dsIdx}-${i}`} style={{ borderTop: "1px solid #2b1414" }}>
                          <td style={{ padding: "7px 10px", color: "#f87171", fontFamily: "monospace" }}>{u.secuencia || "\u2014"}</td>
                          <td style={{ padding: "7px 10px", color: "#e2e8f0" }}>{u.pedimento || "\u2014"}</td>
                          <td style={{ padding: "7px 10px", color: "#f59e0b" }}>{u.fraccion || "\u2014"}</td>
                          <td style={{ padding: "7px 10px", color: "#94a3b8" }}>{u.pais || "\u2014"}</td>
                          <td style={{ padding: "7px 10px", color: "#94a3b8", fontFamily: "monospace" }}>{Number(u.cantidadDS || 0).toLocaleString("es-MX")}</td>
                          <td style={{ padding: "7px 10px", color: "#94a3b8", fontFamily: "monospace" }}>${Number(u.valorDS || 0).toLocaleString("es-MX", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
                          <td style={{ padding: "7px 10px", color: "#fecaca", lineHeight: 1.4, maxWidth: 360 }}>{u.razon || "Sin razón calculada"}</td>
                          <td style={{ padding: "7px 10px", color: "#fde68a", lineHeight: 1.4, maxWidth: 320 }}>{u.probable || "Sin candidato probable en Layout"}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  {(results2020.unusedDSDetails || []).length > 250 && (
                    <div style={{ padding: "8px 10px", color: "#fca5a5", fontSize: 11 }}>
                      ... y {(results2020.unusedDSDetails || []).length - 250} secuencias DS no usadas adicionales (ver columna REVISADO en el Excel).
                    </div>
                  )}
                </div>
                </>
              )}
            </div>
          )}

          {/*  TABLA IN-APP  */}
          {tableData2020 && (() => {
            const pedList = ["TODOS", ...Array.from(new Set(tableData2020.map(r => r.ped).filter(Boolean))).sort()];
            const filtered = filterPed2020 === "TODOS" ? tableData2020 : tableData2020.filter(r => r.ped === filterPed2020);

            const statusColor = s => s === "ok" ? "#22c55e" : s === "new" ? "#f59e0b" : s === "corrected" ? "#fb923c" : "#ef4444";
            const statusLabel = s => s === "ok" ? "OK" : s === "new" ? "NUEVA" : s === "corrected" ? "CORR" : "";
            const rowBg       = s => s === "ok" ? "rgba(34,197,94,0.06)" : s === "new" ? "rgba(245,158,11,0.07)" : s === "corrected" ? "rgba(251,146,60,0.07)" : "rgba(239,68,68,0.07)";

            const nDescCmp = s => String(s ?? "").trim().toLowerCase().replace(/\s+/g, " ");
            const normCmp  = s => String(s ?? "").trim().toUpperCase();

            // Helpers de comparación por pedimento (suma del grupo vs DS)
            const cantInfo = r => {
              if (r.dsCant === null || r.groupSumCant === null) return { color:"#64748b", label: Number(r.cant).toLocaleString("es-MX") };
              const diff    = r.groupSumCant - r.dsCant;
              const absDiff = Math.abs(diff);
              const lbl     = Number(r.cant).toLocaleString("es-MX");
              if (absDiff <= 1) return { color:"#22c55e", label: lbl };
              // Cualquier diferencia > 1 en cantidad es problema  mostrar en rojo con el delta
              return { color:"#ef4444", label: `${lbl} (${diff>0?"+":""}${diff.toLocaleString("es-MX")})` };
            };
            const valInfo = r => {
              if (r.dsVal === null || r.groupSumVal === null) return { color:"#64748b", label: `$${Number(r.val).toLocaleString("es-MX",{minimumFractionDigits:2,maximumFractionDigits:2})}` };
              const diff    = r.groupSumVal - r.dsVal;
              const absDiff = Math.abs(diff);
              const lbl     = `$${Number(r.val).toLocaleString("es-MX",{minimumFractionDigits:2,maximumFractionDigits:2})}`;
              if (absDiff <= 1)  return { color:"#22c55e", label: lbl };           // óptimo
              if (absDiff <= 4)  return { color:"#f97316", label: `${lbl} (${diff>0?"+":""}${diff.toFixed(2)})` }; // aceptable
              return { color:"#ef4444", label: `${lbl} (${diff>0?"+":""}${diff.toFixed(2)})` }; // fuera de tolerancia
            };
            const paisInfo = r => {
              if (!r.dsPais) return { color:"#94a3b8", label: r.pais };
              const match = normCmp(r.pais) === normCmp(r.dsPais);
              if (match) return { color:"#22c55e", label: r.pais };
              return { color:"#fbbf24", label: `${r.pais}`, sub: `DS: ${r.dsPais}` };
            };
            const descInfo = r => {
              if (!r.dsDesc) return { color:"#94a3b8", label: r.desc };
              const match = nDescCmp(r.desc) === nDescCmp(r.dsDesc);
              if (match) return { color:"#22c55e", label: r.desc };
              return { color:"#fbbf24", label: r.desc, sub: `DS: ${r.dsDesc.slice(0,60)}${r.dsDesc.length>60?"":""}` };
            };

            const copyTSV = () => {
              const hdr = "SEC CALC\tPedimento\tFraccion\tPais\tDescripcion\tCantidad\tValor USD\tCadena (L vs DS)\tCadena total (grupo vs DS)\tSecuencias en grupo\tEstado\tNotas\tMotivo detallado";
              const body = filtered.map(r => [
                r.secNueva, r.ped, r.frac, r.pais,
                r.desc, r.cant, r.val.toFixed(2),
                (r.cadenaLayout ? `L: ${r.cadenaLayout} | DS: ${r.cadenaDS || ""}` : ""),
                (r.cadenaTotalLayout ? `L(total): ${r.cadenaTotalLayout} | DS: ${r.cadenaDS || ""}${r.cadenaTotalCoincide ? "  grupo" : ""}` : ""),
                (r.secuenciasEnGrupo?.length ? (() => { const c={}; for (const s of r.secuenciasEnGrupo) { const v=s||""; c[v]=(c[v]||0)+1; } return Object.entries(c).map(([sec,n]) => n>1 ? `${sec} (se repite ${n})` : sec).join("; "); })() : ""),
                statusLabel(r.status),
                r.notasModificaciones ?? "OK",
                r.motivoDetallado ?? ""
              ].join("\t")).join("\n");
              navigator.clipboard.writeText(hdr + "\n" + body).then(() => {
                setCopiedMsg("¡Tabla copiada! Pega en Excel con Ctrl+V");
                setTimeout(() => setCopiedMsg(""), 3000);
              });
            };

            const copySecs = () => {
              const seqs = filtered.map(r => r.secNueva || "").join("\n");
              navigator.clipboard.writeText(seqs).then(() => {
                setCopiedMsg("¡Secuencias copiadas! Pega en Excel con Ctrl+V");
                setTimeout(() => setCopiedMsg(""), 3000);
              });
            };

            const copyPaises = () => {
              // Copia el país "correcto": si el DS tiene país, usa ese; si no, usa el del Layout
              const lines = filtered.map(r => {
                const lyP = r.pais || "";
                const dsP = r.dsPais || "";
                return dsP ? dsP : lyP;
              }).join("\n");
              navigator.clipboard.writeText(lines).then(() => {
                setCopiedMsg("¡Países copiados! (país del DS cuando existe, si no el del Layout). Pega en Excel con Ctrl+V");
                setTimeout(() => setCopiedMsg(""), 3500);
              });
            };

            const copyFracciones = () => {
              // Copia la fracción "correcta": si hubo corrección (fracCorr), esa; si no, la del Layout
              const lines = filtered.map(r => (r.fracCorr != null ? r.fracCorr : (r.frac || ""))).join("\n");
              navigator.clipboard.writeText(lines).then(() => {
                setCopiedMsg("¡Fracciones copiadas! (corregida cuando aplica, si no la del Layout). Pega en Excel con Ctrl+V");
                setTimeout(() => setCopiedMsg(""), 3500);
              });
            };

            const copyDescripciones = () => {
              // Copia la descripción "correcta": si el DS tiene descripción, usa esa; si no, la del Layout
              const lines = filtered.map(r => (r.dsDesc ? r.dsDesc : (r.desc || ""))).join("\n");
              navigator.clipboard.writeText(lines).then(() => {
                setCopiedMsg("¡Descripciones copiadas! (del DS cuando existe, si no la del Layout). Pega en Excel con Ctrl+V");
                setTimeout(() => setCopiedMsg(""), 3500);
              });
            };

            return (
              <div style={{marginTop:8}}>
                {/* Barra de herramientas */}
                <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:12,flexWrap:"wrap"}}>
                  <div style={{color:"#94a3b8",fontSize:12,fontWeight:700,letterSpacing:"0.08em"}}>TABLA DE SECUENCIAS IN-APP</div>
                  <div style={{flex:1}} />
                  {/* Filtro pedimento */}
                  <select
                    value={filterPed2020}
                    onChange={e => setFilterPed2020(e.target.value)}
                    style={{background:"#0f172a",color:"#f8fafc",border:"1px solid #334155",borderRadius:4,padding:"5px 10px",fontSize:12,cursor:"pointer"}}
                  >
                    {pedList.map(p => <option key={p} value={p}>{p === "TODOS" ? "Todos los pedimentos" : p}</option>)}
                  </select>
                  {/* Botón copiar tabla */}
                  <button onClick={copyTSV} style={{background:"#1e40af",border:"none",color:"#bfdbfe",padding:"6px 14px",cursor:"pointer",borderRadius:4,fontSize:12,fontWeight:700}}>
                     Copiar tabla (Excel)
                  </button>
                  {/* Botón copiar solo SECs */}
                  <button onClick={copySecs} style={{background:"#14532d",border:"none",color:"#86efac",padding:"6px 14px",cursor:"pointer",borderRadius:4,fontSize:12,fontWeight:700}}>
                    # Copiar solo SECs
                  </button>
                  {/* Botón copiar países */}
                  <button onClick={copyPaises} style={{background:"#78350f",border:"none",color:"#fde68a",padding:"6px 14px",cursor:"pointer",borderRadius:4,fontSize:12,fontWeight:700}}>
                     Copiar países
                  </button>
                  {/* Botón copiar fracciones corregidas */}
                  <button onClick={copyFracciones} style={{background:"#581c87",border:"none",color:"#e9d5ff",padding:"6px 14px",cursor:"pointer",borderRadius:4,fontSize:12,fontWeight:700}}>
                     Copiar fracciones
                  </button>
                  {/* Botón copiar descripciones corregidas */}
                  <button onClick={copyDescripciones} style={{background:"#4c1d95",border:"none",color:"#c4b5fd",padding:"6px 14px",cursor:"pointer",borderRadius:4,fontSize:12,fontWeight:700}}>
                     Copiar descripciones
                  </button>
                </div>

                {/* Leyenda comparación */}
                <div style={{display:"flex",gap:16,flexWrap:"wrap",marginBottom:10,fontSize:11,color:"#64748b"}}>
                  {[["#22c55e","Cant ±1 / Val ±1 (correcto)"],["#f97316","Val ±2 a ±4 (aceptable)"],["#ef4444","Cant >1 o Val >4 (fuera de tolerancia)"],["#fbbf24","País/Desc distinto al DS"],["#a855f7","Asignación forzada/respaldo"]].map(([c,t])=>(
                    <span key={t}><span style={{display:"inline-block",width:8,height:8,borderRadius:"50%",background:c,marginRight:4}} />{t}</span>
                  ))}
                </div>

                {/* Mensaje de copiado */}
                {copiedMsg && (
                  <div style={{background:"rgba(34,197,94,0.15)",border:"1px solid #22c55e",borderRadius:4,padding:"8px 14px",marginBottom:10,color:"#86efac",fontSize:12}}>
                     {copiedMsg}
                  </div>
                )}

                {/* Tabla  scroll horizontal para ver cadenas completas */}
                <div style={{overflowX:"auto",overflowY:"auto",borderRadius:6,border:"1px solid #1e293b",maxHeight:500,WebkitOverflowScrolling:"touch"}}>
                  <table style={{width:"100%",minWidth:1200,borderCollapse:"collapse",fontSize:12,fontFamily:"DM Mono, monospace"}}>
                    <thead>
                      <tr style={{background:"#0f172a",position:"sticky",top:0,zIndex:2}}>
                        {["SEC CALC","Pedimento","Fracción","País","Descripción","Cantidad","Valor USD","Cadena (L vs DS)","Cadena total (grupo vs DS)","Secuencias en grupo","Estado","Notas","Motivo detallado"].map(h => (
                          <th key={h} style={{padding:"8px 10px",textAlign:["Cantidad","Valor USD"].includes(h)?"right":"left",color:"#64748b",fontWeight:700,borderBottom:"1px solid #1e293b",whiteSpace:"nowrap",fontSize:11}}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {filtered.map(r => {
                        const ci = cantInfo(r);
                        const vi = valInfo(r);
                        const pi = paisInfo(r);
                        const di = descInfo(r);
                        return (
                          <tr key={r.idx} style={{background:rowBg(r.status),borderBottom:"1px solid rgba(30,41,59,0.8)"}}>
                            {/* SEC CALC  muestra sec anterior si fue corregida */}
                            <td style={{padding:"6px 10px",fontWeight:900,fontSize:14,color:statusColor(r.status),minWidth:80}}>
                              {r.secNueva || <span style={{color:"#475569"}}></span>}
                              {r.secPrevCorr && r.secPrevCorr !== r.secNueva && (
                                <div style={{fontSize:10,fontWeight:400,color:"#f97316",marginTop:2}}>
                                  antes: {r.secPrevCorr}
                                </div>
                              )}
                            </td>
                            <td style={{padding:"6px 10px",color:"#cbd5e1",whiteSpace:"nowrap"}}>{r.ped.slice(-6)}</td>
                            {/* Fracción  morado si fue corregida cross-fraction */}
                            <td style={{padding:"6px 10px",minWidth:80}}>
                              {r.fracCorr ? (
                                <span title={`Fracción original en Layout: ${r.fracOrig}`}>
                                  <span style={{color:"#c084fc",fontWeight:700}}>{r.fracCorr}</span>
                                  <div style={{color:"#9333ea",fontSize:10,opacity:0.8}}>orig: {r.fracOrig}</div>
                                </span>
                              ) : (
                                <span style={{color:"#cbd5e1"}}>{r.frac}</span>
                              )}
                            </td>
                            {/* País  verde si coincide, amarillo si difiere */}
                            <td style={{padding:"6px 10px",minWidth:60}}>
                              <span style={{color:pi.color,fontWeight:pi.sub?700:400}}>{r.pais||""}</span>
                              {pi.sub && <div style={{color:"#fbbf24",fontSize:10,opacity:0.8}}>{pi.sub}</div>}
                            </td>
                            {/* Descripción  verde si coincide, amarillo si difiere */}
                            <td style={{padding:"6px 10px",maxWidth:260}} title={r.desc}>
                              <div style={{color:di.color,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>
                                {r.desc.slice(0,55)}{r.desc.length>55?"":""}
                              </div>
                              {di.sub && <div style={{color:"#fbbf24",fontSize:10,opacity:0.8,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{di.sub}</div>}
                            </td>
                            {/* Cantidad  color según diferencia grupo vs DS */}
                            <td style={{padding:"6px 10px",textAlign:"right",color:ci.color,fontWeight:600,whiteSpace:"nowrap"}}>{ci.label}</td>
                            {/* Valor USD  color según diferencia grupo vs DS */}
                            <td style={{padding:"6px 10px",textAlign:"right",color:vi.color,fontWeight:600,whiteSpace:"nowrap"}}>{vi.label}</td>
                            {/* Cadena Layout vs DS  siempre muestra L y DS cuando hay secuencia; verde si coinciden, naranja si no */}
                            <td style={{padding:"6px 10px",minWidth:280,maxWidth:380,fontSize:10,background:(r.secNueva || r.cadenaDS) ? (r.cadenaCoincide?"#145A3222":"#92400E22") : "#1e293b",color:(r.secNueva || r.cadenaDS) ? (r.cadenaCoincide?"#145A32":"#92400E") : "#64748b",borderLeft:(r.secNueva || r.cadenaDS) ? (r.cadenaCoincide?"2px solid #22c55e":"2px solid #f97316") : "1px solid #334155",whiteSpace:"pre-wrap",wordBreak:"break-all"}} title={`Layout: ${r.cadenaLayout}\nDS: ${r.cadenaDS||""}`}>
                              <>
                                <div style={{fontWeight:600}}>L: {r.cadenaLayout}</div>
                                <div style={{marginTop:2,opacity:0.9}}>DS: {r.cadenaDS || ""}</div>
                                {r.cadenaCoincide && (r.secNueva || r.cadenaDS) && <div style={{marginTop:4,fontSize:9,color:"#22c55e",fontWeight:700}}> Coinciden</div>}
                              </>
                            </td>
                            {/* Cadena total (grupo vs DS): cuando hay varias líneas, suma del grupo vs DS  verde si coinciden totales, naranja si no */}
                            <td style={{padding:"6px 10px",minWidth:280,maxWidth:380,fontSize:10,background:(r.cadenaTotalLayout && !r.cadenaCoincide) ? (r.cadenaTotalCoincide ? "#145A3222" : "#92400E22") : "transparent",color:(r.cadenaTotalLayout && !r.cadenaCoincide) ? (r.cadenaTotalCoincide ? "#145A32" : "#92400E") : "#64748b",borderLeft:(r.cadenaTotalLayout && !r.cadenaCoincide) ? (r.cadenaTotalCoincide ? "2px solid #22c55e" : "2px solid #f97316") : "1px solid #334155",whiteSpace:"pre-wrap",wordBreak:"break-all"}} title={r.cadenaTotalLayout ? `L (total grupo): ${r.cadenaTotalLayout}\nDS: ${r.cadenaDS||""}` : ""}>
                              {r.cadenaCoincide ? (
                                <span style={{color:"#64748b",fontSize:10}}></span>
                              ) : r.cadenaTotalLayout ? (
                                <>
                                  <div style={{fontWeight:600}}>L (total): {r.cadenaTotalLayout}</div>
                                  <div style={{marginTop:2,opacity:0.9}}>DS: {r.cadenaDS || ""}</div>
                                  {r.cadenaTotalCoincide && <div style={{marginTop:4,fontSize:9,color:"#22c55e",fontWeight:700}}> Coincide a nivel grupo</div>}
                                </>
                              ) : (
                                <span style={{color:"#64748b"}}></span>
                              )}
                            </td>
                            {/* Secuencias involucradas en el grupo: valor único con "(se repite N)" en vez de repetir */}
                            <td style={{padding:"6px 10px",minWidth:120,fontSize:11,color:"#94a3b8",whiteSpace:"pre-wrap",wordBreak:"break-word"}} title={r.secuenciasEnGrupo?.length ? `Secuencias en este grupo: ${r.secuenciasEnGrupo.join(", ")}` : ""}>
                              {r.secuenciasEnGrupo?.length ? (() => {
                                const countBy = {};
                                for (const s of r.secuenciasEnGrupo) { const v = s || ""; countBy[v] = (countBy[v] || 0) + 1; }
                                const txt = Object.entries(countBy).map(([sec, n]) => n > 1 ? `${sec} (se repite ${n})` : sec).join("; ");
                                return <div style={{fontWeight:600,color:"#cbd5e1"}}>{txt}</div>;
                              })() : (
                                <span style={{color:"#64748b"}}></span>
                              )}
                            </td>
                            <td style={{padding:"6px 10px"}}>
                              <span style={{background:statusColor(r.status)+"22",color:statusColor(r.status),padding:"2px 7px",borderRadius:3,fontSize:10,fontWeight:700}}>
                                {statusLabel(r.status)}
                              </span>
                            </td>
                            {/* Notas: correcciones (país, descripción, fracción, secuencia) o OK */}
                            <td style={{padding:"6px 10px",minWidth:180,maxWidth:320,fontSize:11,color:"#94a3b8",whiteSpace:"pre-wrap",wordBreak:"break-word"}} title={r.notasModificaciones}>
                              {r.notasModificaciones}
                            </td>
                            {/* Motivo detallado: incluye si fue forzada y si excede tolerancias */}
                            <td style={{padding:"6px 10px",minWidth:300,maxWidth:450,fontSize:10,color:r.fueraRegla?"#fca5a5":"#cbd5e1",whiteSpace:"pre-wrap",wordBreak:"break-word"}} title={r.motivoDetallado}>
                              {r.motivoDetallado}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                  {filtered.length === 0 && (
                    <div style={{padding:"32px",textAlign:"center",color:"#475569",fontSize:13}}>Sin filas para el filtro seleccionado</div>
                  )}
                </div>
                <div style={{marginTop:8,color:"#475569",fontSize:11}}>
                  {filtered.length} filas · La diferencia en Cant/Val es la suma del grupo asignado vs el DS 551
                </div>
              </div>
            );
          })()}
        </div>
      )}
    </div>
  );
}

function stripAccents(s) {
  return String(s || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function normalizeDescText(s) {
  return stripAccents(String(s || ""))
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function normSimpleText(s) {
  return stripAccents(String(s || ""))
    .toUpperCase()
    .replace(/\s+/g, " ")
    .trim();
}

function normCodeText(s) {
  return String(s || "")
    .toUpperCase()
    .replace(/[\s.\-_/]/g, "")
    .trim();
}

function descTokens(s) {
  const stop = new Set(["de", "la", "el", "los", "las", "y", "o", "para", "con", "sin", "del", "al", "the", "and"]);
  const syn = {
    balastra: "balasto",
    balasto: "balasto",
    ballast: "balasto",
    tablilla: "tablilla",
    tarjeta: "tablilla",
    modulo: "modulo",
    mod: "modulo",
    lampara: "lampara",
    lamparas: "lampara",
    foco: "lampara",
    focos: "lampara",
    led: "led",
    driver: "driver",
    resistencia: "resistencia",
    capacitor: "capacitor",
    capacitores: "capacitor",
    circuito: "circuito",
    integrado: "integrado",
    monolitico: "monolitico",
    monolitica: "monolitico",
  };
  return normalizeDescText(s)
    .split(" ")
    .filter((t) => t && !stop.has(t))
    .map((t) => syn[t] || t);
}

function tokenSignature(s) {
  return [...new Set(descTokens(s))].sort().join("|");
}

function compareDescriptions(a, b) {
  const na = normalizeDescText(a);
  const nb = normalizeDescText(b);
  if (!na && !nb) return { status: "empty", score: 1 };
  if (!na || !nb) return { status: "red", score: 0 };
  if (na === nb) return { status: "green", score: 1 };

  const ta = [...new Set(descTokens(a))];
  const tb = [...new Set(descTokens(b))];
  if (!ta.length || !tb.length) return { status: "red", score: 0 };

  const setA = new Set(ta);
  const setB = new Set(tb);
  let inter = 0;
  for (const t of setA) if (setB.has(t)) inter++;
  const union = new Set([...ta, ...tb]).size || 1;
  const jaccard = inter / union;
  const covA = inter / setA.size;
  const covB = inter / setB.size;

  if (jaccard >= 0.95 || (covA === 1 && covB === 1)) return { status: "green", score: jaccard };
  if (jaccard >= 0.45 || covA >= 0.7 || covB >= 0.7 || na.includes(nb) || nb.includes(na)) {
    return { status: "yellow", score: jaccard };
  }
  return { status: "red", score: jaccard };
}

function paintCellWithColor(cell, rgb) {
  const prev = cell?.s || {};
  return {
    ...prev,
    fill: {
      patternType: "solid",
      fgColor: { rgb },
      bgColor: { rgb },
    },
  };
}

function AppDescripcion() {
  const [phaseDesc, setPhaseDesc] = useState("upload");
  const [isDraggingDesc, setIsDraggingDesc] = useState(false);
  const [resultsDesc, setResultsDesc] = useState(null);
  const [outputWbDesc, setOutputWbDesc] = useState(null);
  const [fileNameDesc, setFileNameDesc] = useState("");
  const [errorDesc, setErrorDesc] = useState(null);
  const [progressDesc, setProgressDesc] = useState(0);
  const [filterPedDesc, setFilterPedDesc] = useState("TODOS");
  const [copiedMsgDesc, setCopiedMsgDesc] = useState("");

  const processDescripcion = useCallback(async (file) => {
    setErrorDesc(null);
    setFileNameDesc(file.name);
    setPhaseDesc("processing");
    setProgressDesc(0);

    try {
      const buf = await file.arrayBuffer();
      setProgressDesc(20);
      const newWb = XLSX.read(buf, { type: "array", cellStyles: true });

      const stats = { green: 0, yellow: 0, red: 0, purple: 0, empty: 0, total: 0 };
      const statsExtra = {
        paisOk: 0, paisBad: 0,
        fracOk: 0, fracBad: 0,
        nicoOk: 0, nicoBad: 0,
      };
      const sheetsTouched = [];
      const previewRows = [];
      const pedimentosSet = new Set();

      for (const sheetName of newWb.SheetNames) {
        const ws = newWb.Sheets[sheetName];
        const ref = ws["!ref"];
        if (!ref) continue;
        const range = XLSX.utils.decode_range(ref);
        const pedHeaderAliases = ["pedimento", "ped", "numero de pedimento", "num pedimento"];
        let pedCol = 0; // fallback: columna A
        const headerRow = [];
        for (let c = range.s.c; c <= range.e.c; c++) {
          const addr = XLSX.utils.encode_cell({ r: 0, c });
          headerRow[c] = normalizeDescText(String(ws[addr]?.v ?? ""));
        }
        for (let c = range.s.c; c <= range.e.c; c++) {
          const h = headerRow[c] || "";
          if (pedHeaderAliases.some((a) => h.includes(a))) {
            pedCol = c;
            break;
          }
        }

        const pairs = [];
        let lastPed = "";
        for (let r = 1; r <= range.e.r; r++) {
          const addrJ = XLSX.utils.encode_cell({ r, c: 9 });  // J
          const addrK = XLSX.utils.encode_cell({ r, c: 10 }); // K
          const addrM = XLSX.utils.encode_cell({ r, c: 12 }); // M
          const addrN = XLSX.utils.encode_cell({ r, c: 13 }); // N
          const addrP = XLSX.utils.encode_cell({ r, c: 15 }); // P
          const addrQ = XLSX.utils.encode_cell({ r, c: 16 }); // Q
          const addrS = XLSX.utils.encode_cell({ r, c: 18 }); // S
          const addrT = XLSX.utils.encode_cell({ r, c: 19 }); // T
          const addrPed = XLSX.utils.encode_cell({ r, c: pedCol });
          const vJ = String(ws[addrJ]?.v ?? "").trim();
          const vK = String(ws[addrK]?.v ?? "").trim();
          const vM = String(ws[addrM]?.v ?? "").trim();
          const vN = String(ws[addrN]?.v ?? "").trim();
          const vP = String(ws[addrP]?.v ?? "").trim();
          const vQ = String(ws[addrQ]?.v ?? "").trim();
          const vS = String(ws[addrS]?.v ?? "").trim();
          const vT = String(ws[addrT]?.v ?? "").trim();
          const pedRaw = String(ws[addrPed]?.v ?? "").trim();
          const ped = pedRaw || lastPed || "SIN_PEDIMENTO";
          if (pedRaw) lastPed = pedRaw;
          if (!vJ && !vK && !vM && !vN && !vP && !vQ && !vS && !vT) continue;
          pairs.push({
            r, ped,
            addrJ, addrK, addrM, addrN, addrP, addrQ, addrS, addrT,
            vJ, vK, vM, vN, vP, vQ, vS, vT,
            sigJ: tokenSignature(vJ), sigK: tokenSignature(vK),
            classif: null,
          });
        }
        if (!pairs.length) continue;

        for (const p of pairs) {
          p.classif = compareDescriptions(p.vJ, p.vK).status;
          p.revueltaWith = null;
        }

        // Morado: descripciones "revueltas" entre filas contiguas
        for (let i = 0; i < pairs.length - 1; i++) {
          const a = pairs[i];
          const b = pairs[i + 1];
          if (a.ped !== b.ped) continue; // No mezclar pedimentos
          if (!a.sigJ || !a.sigK || !b.sigJ || !b.sigK) continue;
          const swap = a.sigJ === b.sigK && b.sigJ === a.sigK;
          if (swap && a.sigJ !== a.sigK && b.sigJ !== b.sigK) {
            a.classif = "purple";
            b.classif = "purple";
            a.revueltaWith = { sheetName, ped: b.ped, rowExcel: b.r + 1 };
            b.revueltaWith = { sheetName, ped: a.ped, rowExcel: a.r + 1 };
          }
        }

        for (const p of pairs) {
          stats.total++;
          const rgb =
            p.classif === "green" ? "FF22C55E" :
            p.classif === "yellow" ? "FFFACC15" :
            p.classif === "purple" ? "FF9333EA" :
            p.classif === "empty" ? "FF94A3B8" :
            "FFEF4444";

          if (p.classif === "green") stats.green++;
          else if (p.classif === "yellow") stats.yellow++;
          else if (p.classif === "purple") stats.purple++;
          else if (p.classif === "empty") stats.empty++;
          else stats.red++;

          ws[p.addrJ] = ws[p.addrJ] || { t: "s", v: p.vJ || "" };
          ws[p.addrK] = ws[p.addrK] || { t: "s", v: p.vK || "" };
          ws[p.addrJ].s = paintCellWithColor(ws[p.addrJ], rgb);
          ws[p.addrK].s = paintCellWithColor(ws[p.addrK], rgb);

          const paisOK = normSimpleText(p.vM) === normSimpleText(p.vN);
          const fracOK = normCodeText(p.vP) === normCodeText(p.vQ);
          const nicoOK = normCodeText(p.vS) === normCodeText(p.vT);
          statsExtra[paisOK ? "paisOk" : "paisBad"]++;
          statsExtra[fracOK ? "fracOk" : "fracBad"]++;
          statsExtra[nicoOK ? "nicoOk" : "nicoBad"]++;

          const rgbOk = "FF22C55E";
          const rgbBad = "FFEF4444";
          ws[p.addrM] = ws[p.addrM] || { t: "s", v: p.vM || "" };
          ws[p.addrN] = ws[p.addrN] || { t: "s", v: p.vN || "" };
          ws[p.addrP] = ws[p.addrP] || { t: "s", v: p.vP || "" };
          ws[p.addrQ] = ws[p.addrQ] || { t: "s", v: p.vQ || "" };
          ws[p.addrS] = ws[p.addrS] || { t: "s", v: p.vS || "" };
          ws[p.addrT] = ws[p.addrT] || { t: "s", v: p.vT || "" };
          ws[p.addrM].s = paintCellWithColor(ws[p.addrM], paisOK ? rgbOk : rgbBad);
          ws[p.addrN].s = paintCellWithColor(ws[p.addrN], paisOK ? rgbOk : rgbBad);
          ws[p.addrP].s = paintCellWithColor(ws[p.addrP], fracOK ? rgbOk : rgbBad);
          ws[p.addrQ].s = paintCellWithColor(ws[p.addrQ], fracOK ? rgbOk : rgbBad);
          ws[p.addrS].s = paintCellWithColor(ws[p.addrS], nicoOK ? rgbOk : rgbBad);
          ws[p.addrT].s = paintCellWithColor(ws[p.addrT], nicoOK ? rgbOk : rgbBad);

          if (p.ped && p.ped !== "SIN_PEDIMENTO") pedimentosSet.add(p.ped);
          if (previewRows.length < 5000) {
            previewRows.push({
              sheetName,
              rowExcel: p.r + 1,
              pedimento: p.ped || "SIN_PEDIMENTO",
              descJ: p.vJ || "",
              descK: p.vK || "",
              classif: p.classif || "red",
              paisL: p.vM || "",
              paisR: p.vN || "",
              paisOK,
              fracL: p.vP || "",
              fracR: p.vQ || "",
              fracOK,
              nicoL: p.vS || "",
              nicoR: p.vT || "",
              nicoOK,
              revueltaWith: p.revueltaWith || null,
            });
          }
        }

        sheetsTouched.push({ sheetName, rows: pairs.length });
      }

      if (!sheetsTouched.length) {
        throw new Error("No se detectaron descripciones en columnas J y K. Verifica el formato del archivo.");
      }

      setProgressDesc(100);
      setOutputWbDesc(newWb);
      const pedimentos = [...pedimentosSet].sort();
      setResultsDesc({ stats, statsExtra, sheetsTouched, previewRows, pedimentos, pedCount: pedimentos.length });
      setFilterPedDesc("TODOS");
      setTimeout(() => setPhaseDesc("results"), 350);
    } catch (e) {
      setErrorDesc(e.message);
      setPhaseDesc("upload");
    }
  }, []);

  const onFileDesc = useCallback((file) => {
    if (!file?.name?.match(/\.(xlsx|xls)$/i)) {
      setErrorDesc("Solo archivos Excel (.xlsx / .xls)");
      return;
    }
    processDescripcion(file);
  }, [processDescripcion]);

  const downloadDesc = () => {
    if (!outputWbDesc) return;
    XLSX.writeFile(outputWbDesc, fileNameDesc.replace(/\.xlsx?$/i, "") + "_Comparacion_Descripciones.xlsx");
  };

  const resetDesc = () => {
    setPhaseDesc("upload");
    setResultsDesc(null);
    setOutputWbDesc(null);
    setErrorDesc(null);
    setProgressDesc(0);
    setFilterPedDesc("TODOS");
    setCopiedMsgDesc("");
  };

  const getPreviewRowsFiltered = () => {
    const rows = resultsDesc?.previewRows || [];
    return filterPedDesc === "TODOS" ? rows : rows.filter((r) => r.pedimento === filterPedDesc);
  };

  const copyDescColorsForJK = async () => {
    const rows = getPreviewRowsFiltered();
    const htmlRows = rows.map((r) => {
      const bg =
        r.classif === "green" ? "#22c55e" :
        r.classif === "yellow" ? "#facc15" :
        r.classif === "purple" ? "#a855f7" :
        r.classif === "empty" ? "#94a3b8" :
        "#ef4444";
      const fg = (r.classif === "yellow") ? "#111827" : "#f8fafc";
      const vJ = String(r.descJ ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
      const vK = String(r.descK ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
      return `<tr><td style="background:${bg};color:${fg}">${vJ}</td><td style="background:${bg};color:${fg}">${vK}</td></tr>`;
    }).join("");
    const html = `<table><tbody>${htmlRows}</tbody></table>`;
    const plain = rows.map((r) => {
      return `${String(r.descJ ?? "")}\t${String(r.descK ?? "")}`;
    }).join("\n");
    try {
      if (window.ClipboardItem) {
        const item = new ClipboardItem({
          "text/html": new Blob([html], { type: "text/html" }),
          "text/plain": new Blob([plain], { type: "text/plain" }),
        });
        await navigator.clipboard.write([item]);
      } else {
        await navigator.clipboard.writeText(plain);
      }
      setCopiedMsgDesc("J/K copiados con sus valores originales y color. Pega en Excel sobre el rango destino.");
      setTimeout(() => setCopiedMsgDesc(""), 3200);
    } catch {
      await navigator.clipboard.writeText(plain);
      setCopiedMsgDesc("J/K copiados en texto. Si no respeta color, usa el Excel generado.");
      setTimeout(() => setCopiedMsgDesc(""), 3200);
    }
  };

  return (
    <div>
      {phaseDesc === "results" && (
        <div style={{ display:"flex", gap:8, marginBottom:24 }}>
          <button onClick={resetDesc} style={{ background:"transparent",border:"1px solid #334155",color:"#94a3b8",padding:"8px 16px",cursor:"pointer",borderRadius:4,fontSize:13 }}>← Nuevo archivo</button>
          <button onClick={downloadDesc} style={{ background:"#a855f7",border:"none",color:"#0f172a",padding:"8px 20px",cursor:"pointer",borderRadius:4,fontSize:13,fontWeight:800 }}>⬇ Descargar Excel comparado</button>
        </div>
      )}

      {phaseDesc === "upload" && (
        <div style={{ animation: "fadeUp 0.4s ease" }}>
          <div style={{ textAlign: "center", marginBottom: 40 }}>
            <div style={{ display:"inline-block", background:"rgba(168,85,247,0.12)", border:"1px solid rgba(168,85,247,0.4)", color:"#c084fc", padding:"4px 14px", borderRadius:20, fontSize:11, letterSpacing:"0.12em", fontFamily:"DM Mono, monospace", marginBottom:20 }}>
              COMPARADOR DE DESCRIPCIONES · J vs K
            </div>
            <h1 style={{ fontSize: 36, fontWeight: 900, margin: "0 0 12px", letterSpacing: "-0.02em", lineHeight: 1.1 }}>
              Comparación inteligente de <span style={{ color: "#a855f7" }}>descripciones</span>
            </h1>
            <p style={{ color:"#64748b", fontSize:15, maxWidth:700, margin:"0 auto" }}>
              Se colorean columnas J y K respetando formato del archivo: verde (igual), amarillo (sinónimos/orden distinto), rojo (no relacionadas) y morado (revueltas entre filas).
            </p>
          </div>

          {errorDesc && (
            <div style={{ background:"rgba(239,68,68,0.1)", border:"1px solid rgba(239,68,68,0.35)", borderRadius:4, padding:"12px 16px", marginBottom:20, color:"#fca5a5", fontSize:13 }}>
              ⚠ {errorDesc}
            </div>
          )}

          <UploadZone onFile={onFileDesc} isDragging={isDraggingDesc} setIsDragging={setIsDraggingDesc} />
        </div>
      )}

      {phaseDesc === "processing" && (
        <div style={{ textAlign:"center", padding:"80px 0", animation:"fadeUp 0.3s ease" }}>
          <div style={{ width:60, height:60, border:"3px solid #1e293b", borderTop:"3px solid #a855f7", borderRadius:"50%", margin:"0 auto 32px", animation:"spin 0.8s linear infinite" }} />
          <div style={{ fontSize:20, fontWeight:800, marginBottom:8 }}>Analizando descripciones J/K</div>
          <div style={{ color:"#475569", fontSize:13, marginBottom:24 }}>{fileNameDesc}</div>
          <div style={{ maxWidth:400, margin:"0 auto" }}>
            <div style={{ background:"#1e293b", borderRadius:2, height:4, overflow:"hidden" }}>
              <div style={{ height:"100%", background:"#a855f7", borderRadius:2, width:`${progressDesc}%`, transition:"width 0.4s ease" }} />
            </div>
            <div style={{ color:"#475569", fontSize:12, marginTop:8, fontFamily:"DM Mono, monospace" }}>{progressDesc}%</div>
          </div>
        </div>
      )}

      {phaseDesc === "results" && resultsDesc && (
        <div style={{ animation:"fadeUp 0.5s ease" }}>
          <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit, minmax(160px, 1fr))", gap:12, marginBottom:20 }}>
            <StatCard label="Filas comparadas" value={resultsDesc.stats.total} sub="Con datos en J/K" accent="#a855f7" />
            <StatCard label="Pedimentos detectados" value={resultsDesc.pedCount || 0} sub="Procesados por separado" accent="#38bdf8" />
            <StatCard label="Exactas" value={resultsDesc.stats.green} sub="Verde" accent="#22c55e" />
            <StatCard label="Similares" value={resultsDesc.stats.yellow} sub="Amarillo" accent="#facc15" />
            <StatCard label="Revueltas" value={resultsDesc.stats.purple} sub="Morado" accent="#9333ea" />
            <StatCard label="No relacionadas" value={resultsDesc.stats.red} sub="Rojo" accent="#ef4444" />
            <StatCard label="País OK" value={resultsDesc.statsExtra?.paisOk ?? 0} sub="M/N en verde" accent="#22c55e" />
            <StatCard label="Fracción OK" value={resultsDesc.statsExtra?.fracOk ?? 0} sub="P/Q en verde" accent="#22c55e" />
            <StatCard label="NICO OK" value={resultsDesc.statsExtra?.nicoOk ?? 0} sub="S/T en verde" accent="#22c55e" />
          </div>

          <div style={{ background:"#0f172a", border:"1px solid #1e293b", borderRadius:6, padding:"14px 16px", color:"#94a3b8", fontSize:12 }}>
            Hojas evaluadas: {resultsDesc.sheetsTouched.map(s => `${s.sheetName} (${s.rows})`).join(" · ")}
          </div>

          <div style={{ marginTop: 12, background:"#0f172a", border:"1px solid #1e293b", borderRadius:6, padding:"12px 14px", color:"#94a3b8", fontSize:12 }}>
            Pedimentos en archivo: {resultsDesc.pedCount || 0}
            {(resultsDesc.pedimentos || []).length > 0 && (
              <span> · {resultsDesc.pedimentos.join(" · ")}</span>
            )}
          </div>

          <div style={{ marginTop: 10, display:"flex", gap:8, flexWrap:"wrap", alignItems:"center" }}>
            <button onClick={copyDescColorsForJK} style={{ background:"#7c3aed", border:"none", color:"#ede9fe", padding:"6px 12px", cursor:"pointer", borderRadius:4, fontSize:12, fontWeight:700 }}>
              🎨 Copiar J/K con color (2 col)
            </button>
            {copiedMsgDesc && (
              <span style={{ color:"#86efac", fontSize:12 }}>{copiedMsgDesc}</span>
            )}
          </div>

          {(() => {
            const rows = resultsDesc.previewRows || [];
            const pedList = ["TODOS", ...(resultsDesc.pedimentos || [])];
            const filtered = filterPedDesc === "TODOS" ? rows : rows.filter(r => r.pedimento === filterPedDesc);
            const tone = (c) => c === "green" ? "#22c55e" : c === "yellow" ? "#facc15" : c === "purple" ? "#a855f7" : c === "empty" ? "#94a3b8" : "#ef4444";
            const bgTone = (c) => c === "green" ? "rgba(34,197,94,0.18)" : c === "yellow" ? "rgba(250,204,21,0.22)" : c === "purple" ? "rgba(168,85,247,0.2)" : c === "empty" ? "rgba(148,163,184,0.14)" : "rgba(239,68,68,0.16)";
            const label = (c) => c === "green" ? "Exacta" : c === "yellow" ? "Similar" : c === "purple" ? "Revueltas" : c === "empty" ? "Vacía" : "No relacionada";
            return (
              <div style={{ marginTop: 14 }}>
                <div style={{ display:"flex", alignItems:"center", gap:10, marginBottom:10, flexWrap:"wrap" }}>
                  <div style={{ color:"#94a3b8", fontSize:12, fontWeight:700, letterSpacing:"0.08em" }}>VISTA PREVIA</div>
                  <div style={{ flex:1 }} />
                  <select
                    value={filterPedDesc}
                    onChange={(e) => setFilterPedDesc(e.target.value)}
                    style={{ background:"#0f172a",color:"#f8fafc",border:"1px solid #334155",borderRadius:4,padding:"5px 10px",fontSize:12,cursor:"pointer" }}
                  >
                    {pedList.map(p => <option key={p} value={p}>{p === "TODOS" ? "Todos los pedimentos" : p}</option>)}
                  </select>
                </div>

                <div style={{ overflowX:"auto", borderRadius:6, border:"1px solid #1e293b", maxHeight:420, overflowY:"auto" }}>
                  <table style={{ width:"100%", borderCollapse:"collapse", fontSize:12, fontFamily:"DM Mono, monospace" }}>
                    <thead>
                      <tr style={{ background:"#0f172a", position:"sticky", top:0, zIndex:2 }}>
                        {["Hoja", "Fila", "Pedimento", "Descripción J", "Descripción K", "Clasificación", "País M/N", "Fracción P/Q", "NICO S/T", "Detalle revuelta"].map(h => (
                          <th key={h} style={{ padding:"8px 10px", textAlign:"left", color:"#64748b", fontWeight:700, borderBottom:"1px solid #1e293b", whiteSpace:"nowrap", fontSize:11 }}>{h}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {filtered.map((r, i) => (
                        <tr key={`${r.sheetName}-${r.rowExcel}-${i}`} style={{ borderBottom:"1px solid rgba(30,41,59,0.8)" }}>
                          <td style={{ padding:"6px 10px", color:"#93c5fd", whiteSpace:"nowrap" }}>{r.sheetName}</td>
                          <td style={{ padding:"6px 10px", color:"#94a3b8" }}>{r.rowExcel}</td>
                          <td style={{ padding:"6px 10px", color:"#cbd5e1", whiteSpace:"nowrap" }}>{r.pedimento}</td>
                          <td style={{ padding:"6px 10px", color:"#e2e8f0", maxWidth:320, whiteSpace:"pre-wrap", wordBreak:"break-word", background:bgTone(r.classif) }}>{r.descJ}</td>
                          <td style={{ padding:"6px 10px", color:"#e2e8f0", maxWidth:320, whiteSpace:"pre-wrap", wordBreak:"break-word", background:bgTone(r.classif) }}>{r.descK}</td>
                          <td style={{ padding:"6px 10px" }}>
                            <span style={{ background: tone(r.classif) + "22", color: tone(r.classif), padding:"2px 7px", borderRadius:3, fontSize:10, fontWeight:700 }}>
                              {label(r.classif)}
                            </span>
                          </td>
                          <td style={{ padding:"6px 10px", minWidth:180 }}>
                            <div style={{ color:"#cbd5e1", marginBottom:2 }}>{r.paisL} ↔ {r.paisR}</div>
                            <span style={{ background:(r.paisOK ? "#22c55e" : "#ef4444") + "22", color:r.paisOK ? "#22c55e" : "#ef4444", padding:"2px 7px", borderRadius:3, fontSize:10, fontWeight:700 }}>
                              {r.paisOK ? "Coincide" : "No coincide"}
                            </span>
                          </td>
                          <td style={{ padding:"6px 10px", minWidth:180 }}>
                            <div style={{ color:"#cbd5e1", marginBottom:2 }}>{r.fracL} ↔ {r.fracR}</div>
                            <span style={{ background:(r.fracOK ? "#22c55e" : "#ef4444") + "22", color:r.fracOK ? "#22c55e" : "#ef4444", padding:"2px 7px", borderRadius:3, fontSize:10, fontWeight:700 }}>
                              {r.fracOK ? "Coincide" : "No coincide"}
                            </span>
                          </td>
                          <td style={{ padding:"6px 10px", minWidth:180 }}>
                            <div style={{ color:"#cbd5e1", marginBottom:2 }}>{r.nicoL} ↔ {r.nicoR}</div>
                            <span style={{ background:(r.nicoOK ? "#22c55e" : "#ef4444") + "22", color:r.nicoOK ? "#22c55e" : "#ef4444", padding:"2px 7px", borderRadius:3, fontSize:10, fontWeight:700 }}>
                              {r.nicoOK ? "Coincide" : "No coincide"}
                            </span>
                          </td>
                          <td style={{ padding:"6px 10px", color:r.classif==="purple" ? "#d8b4fe" : "#64748b", fontSize:11, minWidth:210, maxWidth:260, whiteSpace:"pre-wrap", wordBreak:"break-word" }}>
                            {r.classif === "purple" && r.revueltaWith
                              ? `Revueltas con fila ${r.revueltaWith.rowExcel} (${r.revueltaWith.sheetName} · ${r.revueltaWith.ped})`
                              : "—"}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  {filtered.length === 0 && (
                    <div style={{ padding:"24px", textAlign:"center", color:"#475569", fontSize:12 }}>Sin filas para el filtro seleccionado</div>
                  )}
                </div>
                <div style={{ marginTop:8, color:"#475569", fontSize:11 }}>
                  {filtered.length} filas en vista previa{rows.length >= 5000 ? " (limitada a 5000 para rendimiento)" : ""}
                </div>
              </div>
            );
          })()}
        </div>
      )}
    </div>
  );
}

//  MAIN APP 
export default function App() {
  const [activeTab, setActiveTab] = useState("551"); // "551" | "2020" | "desc"
  const [phase, setPhase] = useState("upload"); // upload | processing | results
  const [isDragging, setIsDragging] = useState(false);
  const [results, setResults] = useState(null);
  const [outputWb, setOutputWb] = useState(null);
  const [fileName, setFileName] = useState("");
  const [error, setError] = useState(null);
  const [showUnmatched, setShowUnmatched] = useState(false);
  const [showCruce, setShowCruce] = useState(true);
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

      const layoutName = resolveLayoutSheetName(wb);
      if (!layoutName) {
        throw new Error('El archivo debe contener una hoja "Layout" (sin importar mayúsculas).');
      }
      const sheet551Name = resolve551SheetName(wb);
      if (!sheet551Name) {
        throw new Error('El archivo debe contener una hoja "551", "DS" o "Data Stage" (datos del pedimento)');
      }

      setProgress(40);
      const layoutRows = readLayoutSheet(wb.Sheets[layoutName]);
      const s551Rows   = read551Sheet(wb.Sheets[sheet551Name]);
      setProgress(60);

      const dsKeys = [...new Set(s551Rows.map((r) => pedCruceKey(r.Pedimento2 || r.Pedimento)).filter(Boolean))];
      const lyKeys = [...new Set(layoutRows.map((r) => pedCruceKey(r.Pedimento || r.pedimento || r.pedimento1)).filter(Boolean))];
      const inters = dsKeys.filter((k) => lyKeys.includes(k));
      const pedMismatch551 = inters.length === 0 && dsKeys.length > 0 && lyKeys.length > 0
        ? { ds: dsKeys.slice(0, 5), layout: lyKeys.slice(0, 5) }
        : null;

      const { assignment, strategyStats, unmatchedFinal, total, rowNotes, cruceData, orphan551Rows, correctionMap, globalTotals, rowMatchMap } = runCascade(layoutRows, s551Rows);
      setProgress(80);

      const newWb = buildOutputExcel(wb, wb.Sheets[layoutName], wb.Sheets[sheet551Name], sheet551Name, assignment, unmatchedFinal, strategyStats, total, rowNotes, cruceData, orphan551Rows, correctionMap, globalTotals, rowMatchMap, layoutName);
      setProgress(100);

      const alertasCruce = (cruceData || []).filter((d) => d.tipo === "matched" && (d.severidad === "alta" || d.severidad === "media")).length;
      setResults({ strategyStats, unmatchedFinal, cruceData, total, orphan551Count: orphan551Rows.length, correctionCount: correctionMap.size, globalTotals, pedMismatch: pedMismatch551, alertasCruce, rowMatchMap });
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
    setShowCruce(true);
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
          }}></div>
          <div>
            <div style={{ fontSize: 15, fontWeight: 800, letterSpacing: "-0.01em", color: "#f8fafc" }}>
              LCK consultores
            </div>
            <div style={{ fontSize: 11, color: "#475569", letterSpacing: "0.08em", fontFamily: "DM Mono, monospace" }}>
              COMERCIO EXTERIOR · INMEX · PEDIMENTO 551
            </div>
          </div>
          {/* Tab selector */}
          <div style={{ display:"flex", gap:4, marginLeft:24, background:"#0f172a", border:"1px solid #1e293b", borderRadius:6, padding:4 }}>
            {[{ id:"551", label:"Módulo 551" }, { id:"2020", label:"Módulo DS" }, { id:"desc", label:"Comparar descripciones" }].map(t => (
              <button key={t.id} onClick={() => setActiveTab(t.id)} style={{
                background: activeTab===t.id ? (t.id==="2020"?"#22c55e":(t.id==="desc"?"#a855f7":"#f59e0b")) : "transparent",
                border:"none", color: activeTab===t.id ? "#0f172a" : "#64748b",
                padding:"6px 16px", cursor:"pointer", borderRadius:4,
                fontSize:12, fontWeight:700, fontFamily:"Syne, sans-serif",
                transition:"all 0.2s",
              }}>{t.label}</button>
            ))}
          </div>
        </div>
        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          {activeTab==="551" && phase === "results" && (
            <>
              <button
                onClick={reset}
                style={{
                  background: "transparent", border: "1px solid #334155",
                  color: "#94a3b8", padding: "8px 16px", cursor: "pointer",
                  borderRadius: 4, fontSize: 13, fontFamily: "Syne, sans-serif",
                }}
              >
                 Nuevo archivo
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
                 Descargar Excel
              </button>
            </>
          )}
        </div>
      </div>

      <div style={{ maxWidth: 1100, margin: "0 auto", padding: "40px 24px" }}>

        {/* MDULO DS */}
        {activeTab === "2020" && <App2020 />}
        {activeTab === "desc" && <AppDescripcion />}

        {/* MDULO 551  original */}
        {activeTab === "551" && <>

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
                 {error}
              </div>
            )}

            <UploadZone onFile={processFile} isDragging={isDragging} setIsDragging={setIsDragging} />

            {/* Strategy cards */}
            <div style={{ marginTop: 48 }}>
              <div style={{ color: "#475569", fontSize: 11, letterSpacing: "0.1em", marginBottom: 20, fontFamily: "DM Mono, monospace" }}>
                METODOLOGÍA DE COINCIDENCIA  CASCADA DE 5 ESTRATEGIAS
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
            <div style={{ fontSize: 20, fontWeight: 800, marginBottom: 8 }}>Procesando archivo</div>
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
            {results.pedMismatch && (
              <div style={{background:"rgba(245,158,11,0.12)",border:"1px solid #f59e0b",borderRadius:6,padding:"14px 18px",marginBottom:24,fontSize:13,color:"#fcd34d"}}>
                <strong> Pedimentos distintos entre 551 y Layout</strong>
                <p style={{margin:"8px 0 0",color:"#fde68a",lineHeight:1.5}}>
                  El 551 tiene pedimentos que no aparecen en el Layout (y viceversa). Por eso no hay match.
                  <br />551: <code style={{background:"rgba(0,0,0,0.2)",padding:"2px 6px",borderRadius:3}}>{results.pedMismatch.ds.join(", ")}</code>
                  <br />Layout: <code style={{background:"rgba(0,0,0,0.2)",padding:"2px 6px",borderRadius:3}}>{results.pedMismatch.layout.join(", ")}</code>
                  <br /><span style={{fontSize:12,opacity:0.9}}>Verifica que ambas hojas correspondan al mismo pedimento o incluyan los mismos pedimentos.</span>
                </p>
              </div>
            )}
            {/* Headline stats */}
            <div style={{ marginBottom: 32 }}>
              <div style={{ color: "#475569", fontSize: 11, letterSpacing: "0.1em", fontFamily: "DM Mono, monospace", marginBottom: 12 }}>
                {fileName} · {results.total} filas procesadas
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(175px, 1fr))", gap: 12 }}>
                <StatCard label="xito global" value={`${pct}%`} sub={`${matched} de ${results.total} filas`} accent="#f59e0b" />
                <StatCard label="Filas asignadas" value={matched} sub="SecuenciaPed actualizada" accent="#22c55e" />
                <StatCard label="Sin match" value={results.unmatchedFinal.length} sub="Revisión manual" accent={results.unmatchedFinal.length > 0 ? "#ef4444" : "#22c55e"} />
                <StatCard label="Correcciones" value={results.correctionCount || 0} sub="Campos ajustados por 551" accent={(results.correctionCount || 0) > 0 ? "#f97316" : "#22c55e"} />
                <StatCard label="Sec. 551 sin asignar" value={results.orphan551Count || 0} sub="Al final del Layout" accent={(results.orphan551Count || 0) > 0 ? "#3b82f6" : "#22c55e"} />
                <StatCard label="Estrategias activas" value={Object.values(results.strategyStats).filter((v) => v > 0).length} sub="de 15 disponibles" accent="#a855f7" />
                <StatCard label="Alertas de cruce" value={results.alertasCruce ?? 0} sub=" grande o respaldo" accent={(results.alertasCruce ?? 0) > 0 ? "#f97316" : "#22c55e"} />
              </div>
            </div>

            {/* Totales globales Layout vs 551 */}
            {results.globalTotals && (() => {
              const gt = results.globalTotals;
              const diffC = gt.layoutCant  - gt.s551Cant;
              const diffV = gt.layoutVCUSD - gt.s551Val;
              const balanced = Math.abs(diffC) < 1 && Math.abs(diffV) < 2;
              return (
                <div style={{
                  background: "#0f172a",
                  border: `1px solid ${balanced ? "rgba(34,197,94,0.4)" : "rgba(245,158,11,0.4)"}`,
                  borderRadius: 4, padding: "20px 28px", marginBottom: 20,
                }}>
                  <div style={{ color: "#94a3b8", fontSize: 11, letterSpacing: "0.1em", fontFamily: "DM Mono, monospace", marginBottom: 14 }}>
                    BALANCE GLOBAL  LAYOUT vs 551
                  </div>
                  <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 12 }}>
                    {[
                      { label: "Layout  Cantidad total", val: gt.layoutCant.toLocaleString("es-MX", {maximumFractionDigits: 0}), color: "#94a3b8" },
                      { label: "551  Cantidad total", val: gt.s551Cant.toLocaleString("es-MX", {maximumFractionDigits: 0}), color: "#94a3b8" },
                      { label: "Layout  Valor USD total", val: `$${gt.layoutVCUSD.toLocaleString("es-MX", {maximumFractionDigits: 2})}`, color: "#94a3b8" },
                      { label: "551  Valor USD total", val: `$${gt.s551Val.toLocaleString("es-MX", {maximumFractionDigits: 2})}`, color: "#94a3b8" },
                    ].map((item) => (
                      <div key={item.label}>
                        <div style={{ color: "#475569", fontSize: 10, fontFamily: "DM Mono, monospace", marginBottom: 4 }}>{item.label}</div>
                        <div style={{ color: item.color, fontFamily: "DM Mono, monospace", fontSize: 14, fontWeight: 700 }}>{item.val}</div>
                      </div>
                    ))}
                  </div>
                  <div style={{ marginTop: 14, paddingTop: 12, borderTop: "1px solid #1e293b", display: "flex", gap: 32, alignItems: "center" }}>
                    <span style={{ color: "#475569", fontSize: 11, fontFamily: "DM Mono, monospace" }}>
                      Dif. Cantidad: <span style={{ color: Math.abs(diffC) < 1 ? "#22c55e" : "#f59e0b", fontWeight: 700 }}>
                        {diffC >= 0 ? "+" : ""}{diffC.toFixed(0)} ud
                      </span>
                    </span>
                    <span style={{ color: "#475569", fontSize: 11, fontFamily: "DM Mono, monospace" }}>
                      Dif. Valor: <span style={{ color: Math.abs(diffV) < 2 ? "#22c55e" : "#f59e0b", fontWeight: 700 }}>
                        {diffV >= 0 ? "+" : ""}${diffV.toFixed(2)}
                      </span>
                    </span>
                    <span style={{
                      marginLeft: "auto",
                      background: balanced ? "rgba(34,197,94,0.1)" : "rgba(245,158,11,0.1)",
                      border: `1px solid ${balanced ? "rgba(34,197,94,0.3)" : "rgba(245,158,11,0.3)"}`,
                      color: balanced ? "#22c55e" : "#f59e0b",
                      padding: "3px 12px", borderRadius: 20, fontSize: 11, fontFamily: "DM Mono, monospace",
                    }}>
                      {balanced ? " BALANCE EXACTO" : " TOTALES DIFIEREN  revisar pedimentos faltantes"}
                    </span>
                  </div>
                </div>
              );
            })()}

            {results.cruceData && results.cruceData.length > 0 && (
              <div style={{ marginBottom: 20 }}>
                <div
                  onClick={() => setShowCruce(!showCruce)}
                  style={{
                    padding: "12px 16px", cursor: "pointer", marginBottom: showCruce ? 10 : 0,
                    display: "flex", alignItems: "center", justifyContent: "space-between",
                    background: "rgba(30,64,175,0.2)", border: "1px solid #334155", borderRadius: 4,
                  }}
                >
                  <div>
                    <span style={{ color: "#f8fafc", fontWeight: 800, fontSize: 13 }}>Motivos de asignación o sin asignar (cruce Layout vs 551)</span>
                    <span style={{ color: "#64748b", fontSize: 12, marginLeft: 12 }}>{results.cruceData.length} grupos · {results.alertasCruce ?? 0} con alerta</span>
                  </div>
                  <span style={{ color: "#94a3b8" }}>{showCruce ? "" : ""}</span>
                </div>
                {showCruce && (
                  <div style={{ overflowX: "auto", border: "1px solid #1e293b", borderRadius: 4, maxHeight: 420, overflowY: "auto" }}>
                    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 11 }}>
                      <thead style={{ position: "sticky", top: 0, background: "#1e293b" }}>
                        <tr>
                          {["Tipo", "Estrat.", "Ped.", "Fracc. L", "Sec.", " cant", " $", "Motivo / nota"].map((h) => (
                            <th key={h} style={{ padding: "8px 10px", textAlign: "left", color: "#94a3b8", fontFamily: "DM Mono, monospace", fontSize: 9, textTransform: "uppercase" }}>{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {results.cruceData.slice(0, 200).map((d, i) => {
                          const sev = d.severidad;
                          const bg = d.tipo === "unmatched" ? "rgba(127,29,29,0.2)"
                            : sev === "alta" ? "rgba(194,65,12,0.2)"
                            : sev === "media" ? "rgba(30,64,175,0.2)" : "transparent";
                          return (
                            <tr key={i} style={{ borderTop: "1px solid #1e293b", background: bg }}>
                              <td style={{ padding: "6px 10px", color: d.tipo === "unmatched" ? "#f87171" : "#22c55e" }}>{d.tipo === "unmatched" ? "Sin" : "OK"}</td>
                              <td style={{ padding: "6px 10px", color: "#f8fafc", fontFamily: "monospace" }}>{d.estrategia}</td>
                              <td style={{ padding: "6px 10px", color: "#e2e8f0" }}>{String(d.pedimento || "").slice(0, 28)}</td>
                              <td style={{ padding: "6px 10px", color: "#f59e0b" }}>{String(d.layoutFrac || "")}</td>
                              <td style={{ padding: "6px 10px", color: "#94a3b8" }}>{d.secAsignada !== "" && d.secAsignada != null ? d.secAsignada : ""}</td>
                              <td style={{ padding: "6px 10px", color: "#94a3b8", fontFamily: "monospace" }}>{d.diffCant == null ? "" : Number(d.diffCant).toFixed(0)}</td>
                              <td style={{ padding: "6px 10px", color: "#94a3b8", fontFamily: "monospace" }}>{d.diffVal == null ? "" : Number(d.diffVal).toFixed(2)}</td>
                              <td style={{ padding: "6px 10px", color: d.tipo === "unmatched" ? "#fca5a5" : "#e2e8f0", lineHeight: 1.45, maxWidth: 420 }}>{(d.motivo || d.nota || "")}</td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                    {results.cruceData.length > 200 && (
                      <div style={{ padding: 8, textAlign: "center", color: "#64748b", fontSize: 10 }}>y {results.cruceData.length - 200} grupos más (ver hoja Cruce en Excel descargado)</div>
                    )}
                  </div>
                )}
              </div>
            )}

            {/* Progress bar */}
            <div style={{
              background: "#0f172a", border: "1px solid #1e293b",
              borderRadius: 4, padding: "24px 28px", marginBottom: 20,
            }}>
              <div style={{ color: "#94a3b8", fontSize: 11, letterSpacing: "0.1em", fontFamily: "DM Mono, monospace", marginBottom: 16 }}>
                DISTRIBUCIN DE ASIGNACIN POR ESTRATEGIA
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
                        <span style={{ color: "#334155", fontSize: 12 }}>{activeStrategy === s.id ? "" : ""}</span>
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
                       {results.unmatchedFinal.length} filas sin match
                    </span>
                    <span style={{ color: "#475569", fontSize: 12, marginLeft: 12 }}>
                      Requieren revisión manual por un especialista
                    </span>
                  </div>
                  <span style={{ color: "#475569" }}>{showUnmatched ? "" : ""}</span>
                </div>

                {showUnmatched && (
                  <div style={{ overflowX: "auto" }}>
                    <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12 }}>
                      <thead>
                        <tr style={{ background: "#1e293b" }}>
                          {["Fraccion", "País", "Cantidad", "VCUSD", "Notas  Motivo sin asignación"].map((h) => (
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
                            <td style={{ padding: "9px 16px", color: "#fca5a5", fontSize: 11, lineHeight: 1.5, maxWidth: 400 }}>{r.Nota || ""}</td>
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
                 Descargar Excel Resultado
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
                    { icon: "", title: "Verificar Fracción Arancelaria", body: "Muchos casos sin match ocurren porque el mismo producto tiene múltiples fracciones (ej: 85322999 vs 85414004 para CAPACITORES). Agregar FraccionImpo como criterio de agrupación resolvería estos casos." },
                    { icon: "", title: "Revisar Pedimentos Pendientes", body: "Si la suma del Layout supera la cantidad del 551, es posible que parte del inventario provenga de pedimentos anteriores no incluidos en el archivo. Solicitar expediente completo." },
                    { icon: "️", title: "Validar Unidades de Medida", body: "Diferencias de cantidad pueden deberse a conversiones UMC/UMT. Verificar si el 551 reporta en unidades distintas al Layout (piezas vs. lotes, kg vs. pzas)." },
                    { icon: "", title: "Conciliación Parcial", body: "Para ARNES ELCTRICO y productos similares con múltiples registros en 551, hacer conciliación ítem por ítem comparando valor unitario (ValorDolares / CantidadUMComercial) como criterio discriminador." },
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

        {/* Cierre módulo 551 */}
        </>}

      </div>
    </div>
  );
}



