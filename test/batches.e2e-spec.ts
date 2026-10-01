import "reflect-metadata";
import { INestApplication } from "@nestjs/common";
import { Test } from "@nestjs/testing";
import ExcelJS from "exceljs";
import request from "supertest";
import { AppModule } from "../src/api/app.module";
import { HttpExceptionFilter } from "../src/api/common/filters/http-exception.filter";

async function spreadsheet(store: string, product: string, quantity: number): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Pedido");
  sheet.addRow(["FILIAL", store]);
  sheet.addRow(["PRODUTO", "QUANTIDADE"]);
  sheet.addRow([product, quantity]);
  return Buffer.from(await workbook.xlsx.writeBuffer());
}

describe("Batch file processing", () => {
  let app: INestApplication;

  beforeAll(async () => {
    const moduleRef = await Test.createTestingModule({ imports: [AppModule] }).compile();
    app = moduleRef.createNestApplication();
    app.useGlobalFilters(new HttpExceptionFilter());
    await app.init();
  });

  afterAll(async () => app.close());

  it("processes multiple associated spreadsheets without an input directory", async () => {
    const batch = {
      entries: [
        { id: "ceramica", type: "file", store: "Cerâmica", fileRef: "sheet1" },
        { id: "coelho", type: "file", store: "Coelho", fileRef: "sheet2" },
      ],
    };

    const response = await request(app.getHttpServer())
      .post("/api/v1/batches/process")
      .field("batch", JSON.stringify(batch))
      .attach("sheet1", await spreadsheet("Cerâmica", "ABACATE", 7), "ceramica.xlsx")
      .attach("sheet2", await spreadsheet("Coelho", "ABACATE", 3), "coelho.xlsx")
      .expect("content-type", /json/)
      .expect(201);

    expect(response.body.requestId).toEqual(expect.any(String));
    expect(response.body.data.summary).toEqual({
      entries: 2,
      files: 2,
      texts: 0,
      items: 2,
      artifacts: 3,
    });
    expect(response.body.data.artifacts).toHaveLength(3);
  });

  it("rejects mismatched file references with a structured error", async () => {
    const response = await request(app.getHttpServer())
      .post("/api/v1/batches/process")
      .field(
        "batch",
        JSON.stringify({
          entries: [{ id: "ceramica", type: "file", store: "Cerâmica", fileRef: "sheet1" }],
        }),
      )
      .attach("wrong", await spreadsheet("Cerâmica", "ABACATE", 7), "ceramica.xlsx")
      .expect(400);

    expect(response.body).toMatchObject({
      requestId: expect.any(String),
      error: { code: "VALIDATION_ERROR" },
    });
  });

  it("processes a mixed file and text batch and returns text warnings", async () => {
    const response = await request(app.getHttpServer())
      .post("/api/v1/batches/process")
      .field(
        "batch",
        JSON.stringify({
          entries: [
            { id: "ceramica", type: "file", store: "Cerâmica", fileRef: "sheet1" },
            {
              id: "manual-coelho",
              type: "text",
              store: "Coelho",
              text: "ABACATE 2\nobservação sem quantidade",
            },
          ],
        }),
      )
      .attach("sheet1", await spreadsheet("Cerâmica", "ABACATE", 7), "ceramica.xlsx")
      .expect(201);

    expect(response.body.data.summary).toMatchObject({ entries: 2, files: 1, texts: 1, items: 2 });
    expect(response.body.warnings).toEqual([
      expect.objectContaining({ code: "UNPARSED_TEXT_LINE", entryId: "manual-coelho" }),
    ]);
  });

  it("processes a text-only batch without multipart files", async () => {
    const response = await request(app.getHttpServer())
      .post("/api/v1/batches/process")
      .field(
        "batch",
        JSON.stringify({
          entries: [{ id: "manual", type: "text", store: "Queimados", text: "ABACATE 5" }],
        }),
      )
      .expect(201);

    expect(response.body.data.summary).toMatchObject({ files: 0, texts: 1, items: 1 });
  });
});
