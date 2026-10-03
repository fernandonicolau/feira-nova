import ExcelJS from "exceljs";
import { looseSpreadsheet, structuredSpreadsheet } from "./fixtures/spreadsheet.fixture";

const core = require("../src/index.js") as {
  canonicalizeProductName(value: unknown): string;
  canonicalStoreName(value: unknown): string;
  textEntryFromContent(entry: { id: string; store: string; text: string }): {
    items: Array<{ productName: string; quantity: number }>;
    warnings: Array<{ entryId: string; line: number; originalText: string }>;
  };
  workbookEntryFromBuffer(entry: { fileName: string; buffer: Buffer; store?: string }): Promise<{
    storeKey: string;
    items: Array<{ productName: string; quantity: number }>;
  }>;
  processBatch(options: {
    entries: Array<{ fileName: string; buffer: Buffer; store?: string }>;
    now?: Date;
  }): Promise<{ artifacts: Array<{ fileName: string; buffer: Buffer }> }>;
};

async function quantityFor(buffer: Buffer, product: string, column: string): Promise<unknown> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer as never);
  const sheet = workbook.worksheets[0];
  for (let row = 1; row <= sheet.rowCount; row += 1) {
    if (String(sheet.getCell(`A${row}`).value ?? "") === product) return sheet.getCell(`${column}${row}`).value;
    if (String(sheet.getCell(`E${row}`).value ?? "") === product) return sheet.getCell(`${column}${row}`).value;
  }
  return undefined;
}

describe("core regression", () => {
  it("parses manual quantities before or after the product and preserves invalid line details", () => {
    const entry = core.textEntryFromContent({ id: "manual", store: "Cerâmica", text: "2 BANANA PRATA\nABACATE 5\nrevisar esta linha" });
    expect(entry.items).toEqual([
      { productName: "BANANA PRATA", quantity: 2 },
      { productName: "ABACATE", quantity: 5 },
    ]);
    expect(entry.warnings).toEqual([
      expect.objectContaining({ entryId: "manual", line: 3, originalText: "revisar esta linha" }),
    ]);
  });
  it.each([
    ["beringela kg", "BERINJELA"],
    ["limão tahiti", "LIMAO"],
    ["banana terra", "BANANA DA TERRA"],
    ["maçã fugi", "MACA FUJI"],
  ])("normalizes %s as %s", (input, expected) => {
    expect(core.canonicalizeProductName(input)).toBe(expected);
  });

  it.each([
    ["Cerâmica", "CERAMICA"],
    ["Nova Iguaçu", "NOVA IGUACU"],
    ["Santa Cruz", "SANTA CRUZ"],
  ])("normalizes store %s", (input, expected) => {
    expect(core.canonicalStoreName(input)).toBe(expected);
  });

  it("extracts structured quantities, including localized decimal strings", async () => {
    const entry = await core.workbookEntryFromBuffer({
      fileName: "ceramica.xlsx",
      buffer: await structuredSpreadsheet("Cerâmica", [["ABACATE", 7], ["BANANA PRATA", "2,5"]]),
    });
    expect(entry.storeKey).toBe("CERAMICA");
    expect(entry.items).toEqual([{ productName: "ABACATE", quantity: 7 }, { productName: "BANANA PRATA", quantity: 2.5 }]);
  });

  it("extracts loose rows and honors an explicit store over the filename", async () => {
    const entry = await core.workbookEntryFromBuffer({
      fileName: "origem-desconhecida.xlsx",
      store: "Coelho",
      buffer: await looseSpreadsheet([["ABACATE 4"], ["BANANA PRATA", 3]]),
    });
    expect(entry.storeKey).toBe("COELHO");
    expect(entry.items).toEqual([{ productName: "ABACATE", quantity: 4 }, { productName: "BANANA PRATA", quantity: 3 }]);
  });

  it("aggregates known products into the correct store columns without filesystem outputs", async () => {
    const result = await core.processBatch({
      entries: [
        { fileName: "ceramica.xlsx", store: "Cerâmica", buffer: await structuredSpreadsheet("Cerâmica", [["ABACATE", 7]]) },
        { fileName: "coelho.xlsx", store: "Coelho", buffer: await structuredSpreadsheet("Coelho", [["ABACATE", 3]]) },
      ],
      now: new Date("2026-10-02T12:00:00Z"),
    });
    const map = result.artifacts.find((artifact) => artifact.fileName === "MAPA.xlsx");
    expect(map).toBeDefined();
    expect(await quantityFor(map!.buffer, "ABACATE", "F")).toBe(7);
    expect(await quantityFor(map!.buffer, "ABACATE", "G")).toBe(3);
  });
});
