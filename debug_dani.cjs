/* eslint-disable no-console */
const XLSX = require("xlsx");

function normStr(v) {
  return String(v ?? "").trim();
}

function normPed(v) {
  // Mantén solo dígitos; útil cuando un pedimento viene con espacios/guiones
  const s = normStr(v);
  const digits = s.replace(/\D+/g, "");
  return digits || s;
}

function findSheet(wb, pred) {
  const names = wb.SheetNames || [];
  for (const n of names) if (pred(String(n))) return n;
  return null;
}

function colLetterToIndex(letter) {
  // "A"->0, "Z"->25, "AA"->26, ...
  const s = String(letter).toUpperCase().trim();
  let n = 0;
  for (let i = 0; i < s.length; i++) {
    n = n * 26 + (s.charCodeAt(i) - 64);
  }
  return n - 1;
}

function listInterestingHeaders(headers) {
  const want = [
    "ped", "pedi",
    "frac",
    "pais", "origen",
    "cant",
    "vc", "usd", "valor",
    "sec", "secu",
    "desc",
    "candado",
  ];
  return headers
    .map((h, i) => ({ i, h: normStr(h) }))
    .filter(({ h }) => h && want.some((w) => h.toLowerCase().includes(w)));
}

function main() {
  const file = process.argv[2];
  if (!file) {
    console.error("Uso: node debug_dani.cjs \"C:\\ruta\\archivo.xlsx\"");
    process.exit(2);
  }

  const wb = XLSX.readFile(file, { cellDates: true, raw: true });
  console.log("Archivo:", file);
  console.log("Hojas:", wb.SheetNames);

  const layoutName = findSheet(wb, (n) => n.toLowerCase().includes("layout")) || wb.SheetNames[0];
  const s551Name =
    findSheet(wb, (n) => n.toLowerCase() === "551") ||
    findSheet(wb, (n) => n.toLowerCase().includes("551")) ||
    findSheet(wb, (n) => n.toLowerCase().includes("data") && n.toLowerCase().includes("stage"));

  console.log("Layout detectado:", layoutName);
  console.log("551/DataStage detectado:", s551Name);

  const layoutSheet = wb.Sheets[layoutName];
  const layoutAoa = XLSX.utils.sheet_to_json(layoutSheet, { header: 1, defval: "" });
  const layoutHeaders = (layoutAoa[0] || []).map((h) => normStr(h));
  console.log("Layout columnas:", layoutHeaders.length);

  const cvIdx = colLetterToIndex("CV");
  console.log("Columna CV idx:", cvIdx, "encabezado:", layoutHeaders[cvIdx] ?? "(sin encabezado en CV)");

  console.log("Encabezados relevantes en Layout (idx: nombre):");
  for (const { i, h } of listInterestingHeaders(layoutHeaders)) {
    console.log(String(i).padStart(3, " "), ":", h);
  }

  // Muestra primeras filas con columnas clave si existen
  const pick = (name) => layoutHeaders.findIndex((h) => h === name);
  const idxPed = pick("Pedimento");
  const idxFrac = pick("FraccionNico");
  const idxPais = pick("PaisOrigen");
  const idxCant = pick("CantidadSaldo");
  const idxVal = pick("VCUSD");
  const idxSec = pick("SecuenciaPed");

  console.log("Índices esperados (Layout):", { idxPed, idxFrac, idxPais, idxCant, idxVal, idxSec });
  console.log("Muestra Layout (primeras 5 filas):");
  for (let r = 1; r <= Math.min(5, layoutAoa.length - 1); r++) {
    const row = layoutAoa[r];
    console.log({
      ped: idxPed >= 0 ? row[idxPed] : undefined,
      pedNorm: idxPed >= 0 ? normPed(row[idxPed]) : undefined,
      frac: idxFrac >= 0 ? row[idxFrac] : undefined,
      pais: idxPais >= 0 ? row[idxPais] : undefined,
      cant: idxCant >= 0 ? row[idxCant] : undefined,
      vcusd: idxVal >= 0 ? row[idxVal] : undefined,
      sec: idxSec >= 0 ? row[idxSec] : undefined,
      CV: row[cvIdx],
    });
  }

  if (s551Name) {
    const s551Sheet = wb.Sheets[s551Name];
    const s551Aoa = XLSX.utils.sheet_to_json(s551Sheet, { header: 1, defval: "" });
    const s551Headers = (s551Aoa[0] || []).map((h) => normStr(h));
    console.log("551 columnas:", s551Headers.length);
    console.log("Encabezados relevantes en 551 (idx: nombre):");
    for (const { i, h } of listInterestingHeaders(s551Headers)) {
      console.log(String(i).padStart(3, " "), ":", h);
    }

    const idx551Ped = s551Headers.findIndex((h) => h === "Pedimento");
    const idx551Frac = s551Headers.findIndex((h) => h === "Fraccion");
    const idx551Pais = s551Headers.findIndex((h) => h === "PaisOrigenDestino");
    const idx551Sec = s551Headers.findIndex((h) => h === "SecuenciaFraccion");
    const idx551Cant = s551Headers.findIndex((h) => h === "CantidadUMComercial");
    const idx551Val = s551Headers.findIndex((h) => h === "ValorDolares");
    console.log("Índices esperados (551):", { idx551Ped, idx551Frac, idx551Pais, idx551Sec, idx551Cant, idx551Val });

    // Validación rápida: ¿hay intersección de pedimentos?
    const layPeds = new Set();
    for (let r = 1; r < layoutAoa.length; r++) if (idxPed >= 0) layPeds.add(normPed(layoutAoa[r][idxPed]));
    const dsPeds = new Set();
    for (let r = 1; r < s551Aoa.length; r++) if (idx551Ped >= 0) dsPeds.add(normPed(s551Aoa[r][idx551Ped]));
    const inter = [...layPeds].filter((p) => dsPeds.has(p));
    console.log("Pedimentos Layout:", layPeds.size, "Pedimentos 551:", dsPeds.size, "Intersección:", inter.length);
    console.log("Ejemplos intersección:", inter.slice(0, 10));
  }
}

main();

