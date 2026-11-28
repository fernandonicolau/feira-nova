import ExcelJS from "exceljs";
import { inspectFile, validateLayout } from "./file-recognition";
import { structuredSpreadsheet, looseSpreadsheet } from "../../../test/fixtures/spreadsheet.fixture";

async function workbookBuffer(rows: unknown[][]): Promise<Buffer> {
  const book = new ExcelJS.Workbook();
  const sheet = book.addWorksheet("Pedido");
  rows.forEach((row) => sheet.addRow(row));
  return Buffer.from(await book.xlsx.writeBuffer());
}

describe("identificação de planilhas", () => {
  it("recognizes a known store from the header without a selection", async () => {
    const result = await inspectFile("pedido.xlsx", await structuredSpreadsheet("Cerâmica", [["ABACATE", 5]]));
    expect(result).toMatchObject({ store: "Cerâmica", source: "header", needsReview: false });
    expect(result.layout?.pairs).toEqual([{ nameColumn: "A", quantityColumn: "B" }]);
    expect(result.preview.itemCount).toBe(1);
  });

  it("uses a known filename when the spreadsheet has no branch header", async () => {
    const result = await inspectFile("pedido-coelho.xlsx", await looseSpreadsheet([["ABACATE", 2]]));
    expect(result).toMatchObject({ store: "Coelho", source: "filename", needsReview: false });
  });

  it("flags conflicting store evidence for review", async () => {
    const result = await inspectFile("pedido-coelho.xlsx", await structuredSpreadsheet("Cerâmica", [["ABACATE", 5]]));
    expect(result.needsReview).toBe(true);
    expect(result.source).toBe("conflict");
    expect(result.preview.warnings).toEqual(expect.arrayContaining([expect.stringContaining("lojas diferentes")]));
  });

  it("finds two name and quantity pairs and honors a manual correction", async () => {
    const buffer = await workbookBuffer([["FILIAL", "Cerâmica"], ["PRODUTO", "QTD", "NOME", "QUANTIDADE"], ["ABACATE", 2, "BANANA PRATA", 4]]);
    const result = await inspectFile("pedido.xlsx", buffer);
    expect(result.layout?.pairs).toEqual([{ nameColumn: "A", quantityColumn: "B" }, { nameColumn: "C", quantityColumn: "D" }]);
    expect(result.preview.itemCount).toBe(2);
    const core = require("../../index.js") as { workbookEntryFromBuffer(input: { fileName: string; buffer: Buffer; store: string; layout?: { headerRow: number; pairs: Array<{ nameColumn: string; quantityColumn: string }> } }): Promise<{ items: Array<{ productName: string; quantity: number }> }> };
    const automatic = await core.workbookEntryFromBuffer({ fileName: "pedido.xlsx", buffer, store: "Cerâmica" });
    expect(automatic.items).toHaveLength(2);

    const unusual = await workbookBuffer([["ABACATE", 2, "BANANA", 4]]);
    const corrected = await inspectFile("sem-loja.xlsx", unusual, { store: "Cerâmica", layout: { headerRow: 0, pairs: [{ nameColumn: "C", quantityColumn: "D" }] } });
    expect(corrected).toMatchObject({ store: "Cerâmica", source: "manual", needsReview: false });
    expect(corrected.preview.sampleItems).toEqual([{ name: "BANANA", quantity: 4 }]);
    expect(() => validateLayout({ headerRow: 0, pairs: [{ nameColumn: "A", quantityColumn: "A" }] })).toThrow();
    const parsed = await core.workbookEntryFromBuffer({ fileName: "sem-loja.xlsx", buffer: unusual, store: "Cerâmica", layout: { headerRow: 0, pairs: [{ nameColumn: "C", quantityColumn: "D" }] } });
    expect(parsed.items).toEqual([{ productName: "BANANA", quantity: 4 }]);
  });
});
