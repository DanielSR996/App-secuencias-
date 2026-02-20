// Ejecutar con: node inspect-excel.js [ruta.xlsx]
import XLSX from "xlsx";
const path = process.argv[2] || "C:\\Users\\LCK_KATHIA\\Desktop\\AYUDA.xlsx";

const wb = XLSX.readFile(path);
console.log("Hojas encontradas:", wb.SheetNames);

wb.SheetNames.forEach((name) => {
  const sheet = wb.Sheets[name];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 });
  const headers = rows[0] || [];
  console.log("\n--- Hoja:", name, "---");
  console.log("Headers:", JSON.stringify(headers, null, 0));
  console.log("Total filas (con header):", rows.length);
  if (rows[1]) console.log("Primera fila datos (ejemplo):", JSON.stringify(rows[1].slice(0, 8)));
});
