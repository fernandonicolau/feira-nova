import ExcelJS from "exceljs";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { structuredSpreadsheet } from "./fixtures/spreadsheet.fixture";

const { processBatch } = require("../src/index.js") as {
  processBatch(options: {
    entries: Array<{ fileName: string; buffer: Buffer; store: string }>;
    now: Date;
  }): Promise<{ artifacts: Array<{ fileName: string; buffer: Buffer }> }>;
};
const { generateSupplierFiles } = require("../scripts/generate-fornecedores.js") as {
  generateSupplierFiles(options: {
    mapDir: string;
    templateDir: string;
    outputDir: string;
    now: Date;
  }): Promise<{ files: string[]; unmatched: unknown[]; unmatchedFile: unknown }>;
};

async function supplierTemplate(filePath: string, supplier: string): Promise<void> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Pedido");
  sheet.addRow([supplier, "01/01/2026"]);
  sheet.addRow(["PRODUTO", "CERAMICA", "TOTAL"]);
  sheet.addRow(["CEBOLA ROXA", null, { formula: "SUM(B3:B3)" }]);
  await workbook.xlsx.writeFile(filePath);
}

describe("supplier regression", () => {
  it("generates a supplier workbook from isolated maps and templates", async () => {
    const root = fs.mkdtempSync(path.join(os.tmpdir(), "feira-nova-supplier-"));
    const mapDir = path.join(root, "maps");
    const templateDir = path.join(root, "templates");
    const outputDir = path.join(root, "output");
    fs.mkdirSync(mapDir);
    fs.mkdirSync(templateDir);

    try {
      const batch = await processBatch({
        entries: [{
          fileName: "ceramica.xlsx",
          store: "Cerâmica",
          buffer: await structuredSpreadsheet("Cerâmica", [["CEBOLA ROXA", 6]]),
        }],
        now: new Date("2026-10-02T12:00:00Z"),
      });
      for (const artifact of batch.artifacts) {
        fs.writeFileSync(path.join(mapDir, artifact.fileName), artifact.buffer);
      }
      await supplierTemplate(path.join(templateDir, "adonai.xlsx"), "ADONAI");
      await supplierTemplate(path.join(templateDir, "Milanes.xlsx"), "MILANES");

      const result = await generateSupplierFiles({
        mapDir,
        templateDir,
        outputDir,
        now: new Date("2026-10-02T12:00:00Z"),
      });
      expect(result.files).toContain("adonai.xlsx");
      expect(result.unmatched).toHaveLength(0);
      expect(result.unmatchedFile).toBeNull();

      const generated = new ExcelJS.Workbook();
      await generated.xlsx.readFile(path.join(outputDir, "adonai.xlsx"));
      const sheet = generated.worksheets[0];
      expect(sheet.getCell("B3").value).toBe(6);
      expect(sheet.getCell("B1").value).toBe("03/10/2026");
    } finally {
      fs.rmSync(root, { recursive: true, force: true });
    }
  });
});
