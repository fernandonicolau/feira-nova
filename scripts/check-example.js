const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");

const ROOT_DIR = process.cwd();
const GENERATED_DIR = fs
  .readdirSync(ROOT_DIR)
  .filter((name) => /^output-\d{4}-\d{2}-\d{2}-\d{2}-\d{2}$/.test(name))
  .sort()
  .at(-1);

if (!GENERATED_DIR) {
  console.error("Nenhuma pasta output-* encontrada.");
  process.exit(1);
}

async function workbookSnapshot(filePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const worksheet = workbook.worksheets[0];
  const snapshot = [];

  for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    const values = [];
    row.eachCell({ includeEmpty: false }, (cell, columnNumber) => {
      values.push(`${columnNumber}:${String(cell.value ?? "")}`);
    });

    if (values.length) {
      snapshot.push(`${rowNumber}|${values.join("|")}`);
    }
  }

  return snapshot.join("\n");
}

async function compareFile(fileName, expectedName = fileName) {
  const actualPath = path.join(ROOT_DIR, GENERATED_DIR, fileName);
  const expectedPath = path.join(ROOT_DIR, "exemplo", "output", "output mapa", expectedName);

  const actual = await workbookSnapshot(actualPath);
  const expected = await workbookSnapshot(expectedPath);

  if (actual === expected) {
    console.log(`${fileName}: OK`);
    return;
  }

  console.log(`${fileName}: DIFERENTE`);
}

(async () => {
  await compareFile("MAPA.xlsx");
  await compareFile("MAPA2.xlsx", "MAPA 2.xlsx");
  await compareFile("MAPA3.xlsx", "MAPA 3.xlsx");
})().catch((error) => {
  console.error(error.message);
  process.exit(1);
});
