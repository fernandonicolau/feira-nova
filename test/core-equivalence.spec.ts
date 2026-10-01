import ExcelJS from "exceljs";

const { processBatch } = require("../src/index.js") as {
  processBatch(options: {
    entries: Array<{ fileName: string; buffer: Buffer; store?: string }>;
    now?: Date;
  }): Promise<{
    summary: { entries: number; items: number; artifacts: number };
    artifacts: Array<{ fileName: string; buffer: Buffer }>;
  }>;
};

describe("spreadsheet core", () => {
  it("processes an in-memory workbook without input or output directories", async () => {
    const input = new ExcelJS.Workbook();
    const sheet = input.addWorksheet("Pedido");
    sheet.addRow(["FILIAL", "CERAMICA"]);
    sheet.addRow(["PRODUTO", "QUANTIDADE"]);
    sheet.addRow(["ABACATE", 7]);

    const result = await processBatch({
      entries: [
        {
          fileName: "ceramica.xlsx",
          buffer: Buffer.from(await input.xlsx.writeBuffer()),
        },
      ],
      now: new Date("2026-10-01T12:00:00Z"),
    });

    expect(result.summary).toEqual({ entries: 1, items: 1, artifacts: 3 });
    expect(result.artifacts.map((artifact) => artifact.fileName)).toEqual([
      "MAPA.xlsx",
      "MAPA2.xlsx",
      "MAPA3.xlsx",
    ]);

    const generated = new ExcelJS.Workbook();
    await generated.xlsx.load(result.artifacts[0].buffer as never);
    const outputSheet = generated.worksheets[0];
    let matchedQuantity: unknown;

    for (let row = 1; row <= outputSheet.rowCount; row += 1) {
      if (String(outputSheet.getCell(`E${row}`).value ?? "").includes("ABACATE")) {
        matchedQuantity = outputSheet.getCell(`F${row}`).value;
        break;
      }
    }

    expect(matchedQuantity).toBe(7);
  });
});
