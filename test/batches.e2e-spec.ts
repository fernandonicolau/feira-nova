import "reflect-metadata";
import { INestApplication } from "@nestjs/common";
import { Test } from "@nestjs/testing";
import JSZip from "jszip";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import request from "supertest";
import { AppModule } from "../src/api/app.module";
import { HttpExceptionFilter } from "../src/api/common/filters/http-exception.filter";
import { structuredSpreadsheet } from "./fixtures/spreadsheet.fixture";

async function spreadsheet(store: string, product: string, quantity: number): Promise<Buffer> {
  return structuredSpreadsheet(store, [[product, quantity]]);
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
    expect(response.body.data.summary).toMatchObject({
      entries: 2,
      files: 2,
      texts: 0,
      items: 2,
    });
    expect(response.body.data.summary.artifacts).toBeGreaterThanOrEqual(4);
    expect(response.body.data.artifacts).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ fileName: "MAPA.xlsx", downloadUrl: "/api/v1/batches/process/download" }),
        expect.objectContaining({ fileName: "MAPA2.xlsx" }),
        expect.objectContaining({ fileName: "MAPA3.xlsx" }),
      ]),
    );
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

  it.each([
    ["missing batch", undefined, "Multipart field 'batch' is required"],
    ["invalid JSON", "{not-json", "Multipart field 'batch' must be valid JSON"],
  ])("returns a structured error for %s", async (_case, batch, message) => {
    let call = request(app.getHttpServer()).post("/api/v1/batches/process");
    if (batch !== undefined) call = call.field("batch", batch);
    const response = await call.expect(400);
    expect(response.body).toMatchObject({ requestId: expect.any(String), error: { code: "VALIDATION_ERROR", message } });
  });

  it("rejects unsupported file formats", async () => {
    const batch = { entries: [{ id: "manual", type: "file", store: "Cerâmica", fileRef: "sheet1" }] };
    const response = await request(app.getHttpServer())
      .post("/api/v1/batches/process")
      .field("batch", JSON.stringify(batch))
      .attach("sheet1", Buffer.from("not a workbook"), "pedido.csv")
      .expect(400);
    expect(response.body).toMatchObject({ requestId: expect.any(String), error: { code: "VALIDATION_ERROR" } });
  });

  it("rejects an invalid spreadsheet with a safe structured response", async () => {
    const batch = { entries: [{ id: "manual", type: "file", store: "Cerâmica", fileRef: "sheet1" }] };
    const response = await request(app.getHttpServer())
      .post("/api/v1/batches/process")
      .field("batch", JSON.stringify(batch))
      .attach("sheet1", Buffer.from("invalid xlsx"), "pedido.xlsx")
      .expect(400);
    expect(response.body.error).toMatchObject({ code: "VALIDATION_ERROR" });
    expect(response.body.error.message).not.toContain("node_modules");
  });

  it("downloads maps, supplier outputs and a structured manifest in one ZIP", async () => {
    const temporaryDirectoriesBefore = new Set(
      fs.readdirSync(os.tmpdir()).filter((name) => name.startsWith("feira-nova-")),
    );
    const batch = {
      name: "Download E2E",
      entries: [{ id: "ceramica", type: "file", store: "Cerâmica", fileRef: "sheet1" }],
    };
    const response = await request(app.getHttpServer())
      .post("/api/v1/batches/process/download")
      .field("batch", JSON.stringify(batch))
      .attach("sheet1", await spreadsheet("Cerâmica", "CEBOLA ROXA", 6), "ceramica.xlsx")
      .buffer(true)
      .parse((res, callback) => {
        const chunks: Buffer[] = [];
        res.on("data", (chunk: Buffer) => chunks.push(chunk));
        res.on("end", () => callback(null, Buffer.concat(chunks)));
      })
      .expect("content-type", /application\/zip/)
      .expect("content-disposition", /attachment; filename="feira-nova-[^"]+\.zip"/)
      .expect(201);

    const zip = await JSZip.loadAsync(response.body as Buffer);
    expect(Object.keys(zip.files)).toEqual(
      expect.arrayContaining([
        "mapas/MAPA.xlsx",
        "mapas/MAPA2.xlsx",
        "mapas/MAPA3.xlsx",
        "fornecedores/adonai.xlsx",
        "manifest.json",
      ]),
    );
    const manifest = JSON.parse(await zip.file("manifest.json")!.async("string")) as {
      requestId: string;
      data: { summary: { artifacts: number } };
      warnings: unknown[];
    };
    expect(manifest.requestId).toBe(response.headers["x-request-id"]);
    expect(manifest.data.summary.artifacts).toBeGreaterThanOrEqual(4);
    expect(Array.isArray(manifest.warnings)).toBe(true);
    const temporaryDirectoriesAfter = fs
      .readdirSync(os.tmpdir())
      .filter((name) => name.startsWith("feira-nova-"))
      .filter((name) => !temporaryDirectoriesBefore.has(name));
    expect(temporaryDirectoriesAfter.map((name) => path.join(os.tmpdir(), name))).toEqual([]);
  });
});
