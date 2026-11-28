import { INestApplication } from "@nestjs/common";
import { Test } from "@nestjs/testing";
import ExcelJS from "exceljs";
import { randomUUID } from "node:crypto";
import request from "supertest";
import { AppModule } from "../src/api/app.module";
import { HttpExceptionFilter } from "../src/api/common/filters/http-exception.filter";
import { structuredSpreadsheet } from "./fixtures/spreadsheet.fixture";

const local = /localhost|127\.0\.0\.1/.test(process.env.SUPABASE_URL ?? "");
const run = process.env.RUN_DAY_INTEGRATION === "1" && local && !!process.env.SUPABASE_SERVICE_ROLE_KEY;

(run ? describe : describe.skip)("file recognition with local Supabase", () => {
  let app: INestApplication;
  const day = "2030-01-03";
  const old = { code: process.env.ACCESS_CODE, secret: process.env.SESSION_SECRET };
  beforeAll(async () => {
    process.env.ACCESS_CODE = "123456";
    process.env.SESSION_SECRET = "local-recognition-integration-secret-123456";
    const moduleRef = await Test.createTestingModule({ imports: [AppModule] }).compile();
    app = moduleRef.createNestApplication();
    app.useGlobalFilters(new HttpExceptionFilter());
    await app.init();
  });
  afterAll(async () => {
    await app.close();
    if (old.code === undefined) delete process.env.ACCESS_CODE; else process.env.ACCESS_CODE = old.code;
    if (old.secret === undefined) delete process.env.SESSION_SECRET; else process.env.SESSION_SECRET = old.secret;
  });

  it("suggests known files and saves human corrections without changing originals", async () => {
    const login = await request(app.getHttpServer()).post("/api/v1/access/session").send({ code: "123456" }).expect(201);
    const cookie = login.headers["set-cookie"][0];
    await request(app.getHttpServer()).get("/api/v1/days/stores").expect(401);
    const stores = await request(app.getHttpServer()).get("/api/v1/days/stores").set("Cookie", cookie).expect(200);
    expect(stores.body).toContain("Cerâmica");
    await request(app.getHttpServer()).put(`/api/v1/days/${day}`).set("Cookie", cookie).expect(200);
    const known = await request(app.getHttpServer()).post(`/api/v1/days/${day}/files`).set("Cookie", cookie).field("clientUploadId", randomUUID()).attach("files", await structuredSpreadsheet("Cerâmica", [["ABACATE", 5]]), "pedido.xlsx").expect(201);
    expect(known.body.entries[0]).toMatchObject({ store: "Cerâmica", detection_source: "header", review_required: false, state: "received" });
    const conflict = await request(app.getHttpServer()).post(`/api/v1/days/${day}/files`).set("Cookie", cookie).field("clientUploadId", randomUUID()).attach("files", await structuredSpreadsheet("Cerâmica", [["ABACATE", 5]]), "pedido-coelho.xlsx").expect(201);
    const conflicted = conflict.body.entries[0];
    expect(conflicted).toMatchObject({ review_required: true, state: "needs_review" });
    const corrected = await request(app.getHttpServer()).patch(`/api/v1/days/${day}/entries/${conflicted.id}/review`).set("Cookie", cookie).send({ store: "Coelho" }).expect(200);
    expect(corrected.body).toMatchObject({ store: "Coelho", detection_source: "manual", review_required: false, state: "received", sha256: conflicted.sha256 });

    const book = new ExcelJS.Workbook();
    const sheet = book.addWorksheet("Pedido");
    sheet.addRow(["?", "?", "ABACATE", 3]);
    const unknown = await request(app.getHttpServer()).post(`/api/v1/days/${day}/files`).set("Cookie", cookie).field("clientUploadId", randomUUID()).attach("files", Buffer.from(await book.xlsx.writeBuffer()), "pedido-olinda.xlsx").expect(201);
    expect(unknown.body.entries[0].review_required).toBe(true);
    const mapped = await request(app.getHttpServer()).patch(`/api/v1/days/${day}/entries/${unknown.body.entries[0].id}/review`).set("Cookie", cookie).send({ store: "Olinda", layout: { headerRow: 0, pairs: [{ nameColumn: "C", quantityColumn: "D" }] } }).expect(200);
    expect(mapped.body).toMatchObject({ store: "Olinda", review_required: false, state: "received", sha256: unknown.body.entries[0].sha256 });
    expect(mapped.body.preview.itemCount).toBe(1);
    const summary = await request(app.getHttpServer()).get(`/api/v1/days/${day}`).set("Cookie", cookie).expect(200);
    expect(summary.body.events.filter((event: { action: string }) => event.action === "reviewed")).toHaveLength(2);
  }, 120_000);
});
