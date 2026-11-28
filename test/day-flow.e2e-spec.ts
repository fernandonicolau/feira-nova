import { INestApplication } from "@nestjs/common";
import { Test } from "@nestjs/testing";
import JSZip from "jszip";
import { randomUUID } from "node:crypto";
import request from "supertest";
import { AppModule } from "../src/api/app.module";
import { HttpExceptionFilter } from "../src/api/common/filters/http-exception.filter";
import { structuredSpreadsheet } from "./fixtures/spreadsheet.fixture";

const local = /localhost|127\.0\.0\.1/.test(process.env.SUPABASE_URL ?? "");
const run = process.env.RUN_DAY_INTEGRATION === "1" && local && !!process.env.SUPABASE_SERVICE_ROLE_KEY;

(run ? describe : describe.skip)("daily product with local Supabase", () => {
  let app: INestApplication;
  const code = "123456";
  const day = "2030-01-02";
  const old = { code: process.env.ACCESS_CODE, secret: process.env.SESSION_SECRET };

  async function start(): Promise<void> {
    const moduleRef = await Test.createTestingModule({ imports: [AppModule] }).compile();
    app = moduleRef.createNestApplication();
    app.useGlobalFilters(new HttpExceptionFilter());
    await app.init();
  }

  beforeAll(async () => {
    process.env.ACCESS_CODE = code;
    process.env.SESSION_SECRET = "local-daily-integration-test-secret-123456";
    await start();
  });
  afterAll(async () => {
    await app.close();
    if (old.code === undefined) delete process.env.ACCESS_CODE; else process.env.ACCESS_CODE = old.code;
    if (old.secret === undefined) delete process.env.SESSION_SECRET; else process.env.SESSION_SECRET = old.secret;
  });

  it("stores files, ZIP members and text, retries safely, and survives app restart", async () => {
    await request(app.getHttpServer()).get(`/api/v1/days/${day}`).expect(401);
    const login = await request(app.getHttpServer()).post("/api/v1/access/session").send({ code }).expect(201);
    const cookie = login.headers["set-cookie"][0];
    await request(app.getHttpServer()).put(`/api/v1/days/${day}`).set("Cookie", cookie).expect(200);
    const file = await structuredSpreadsheet("Cerâmica", [["ABACATE", 5]]);
    const zip = new JSZip();
    zip.file("loja-a/pedido.xlsx", file);
    zip.file("loja-b/pedido.xlsx", file);
    const archive = await zip.generateAsync({ type: "nodebuffer" });
    const uploadId = randomUUID();
    const send = () => request(app.getHttpServer()).post(`/api/v1/days/${day}/files`).set("Cookie", cookie).field("clientUploadId", uploadId).attach("files", archive, "pedidos.zip");
    const first = await send().expect(201);
    expect(first.body.failures).toEqual([]);
    expect(first.body.entries).toHaveLength(2);
    const retry = await send().expect(201);
    expect(retry.body.entries.map((entry: { id: string }) => entry.id)).toEqual(first.body.entries.map((entry: { id: string }) => entry.id));
    const textId = randomUUID();
    const manual = await request(app.getHttpServer()).post(`/api/v1/days/${day}/texts`).set("Cookie", cookie).send({ clientEntryId: textId, text: "ABACATE 5" }).expect(201);
    expect(manual.body.kind).toBe("text");
    const manualRetry = await request(app.getHttpServer()).post(`/api/v1/days/${day}/texts`).set("Cookie", cookie).send({ clientEntryId: textId, text: "ABACATE 5" }).expect(201);
    expect(manualRetry.body.id).toBe(manual.body.id);
    const original = await request(app.getHttpServer()).get(`/api/v1/days/${day}/uploads/${first.body.entries[0].upload_id}/original`).set("Cookie", cookie).expect(200);
    expect(original.headers["content-type"]).toContain("application/zip");
    await app.close();
    await start();
    const persisted = await request(app.getHttpServer()).get(`/api/v1/days/${day}`).set("Cookie", cookie).expect(200);
    expect(persisted.body.entries).toEqual(expect.arrayContaining([expect.objectContaining({ id: manual.body.id })]));
    expect(persisted.body.entries.filter((entry: { upload_id: string }) => entry.upload_id === first.body.entries[0].upload_id)).toHaveLength(2);
    const partial = await request(app.getHttpServer()).post(`/api/v1/days/${day}/files`).set("Cookie", cookie).field("clientUploadId", randomUUID()).attach("files", file, "avulso.xlsx").attach("files", Buffer.from("csv"), "invalido.csv").expect(201);
    expect(partial.body.entries).toHaveLength(1);
    expect(partial.body.failures).toHaveLength(1);
    const replacement = await request(app.getHttpServer()).post(`/api/v1/days/${day}/entries/${partial.body.entries[0].id}/replace`).set("Cookie", cookie).field("clientUploadId", randomUUID()).attach("files", file, "corrigido.xlsx").expect(201);
    expect(replacement.body.failures).toEqual([]);
    await request(app.getHttpServer()).delete(`/api/v1/days/${day}/entries/${manual.body.id}`).set("Cookie", cookie).expect(200);
    const afterDelete = await request(app.getHttpServer()).get(`/api/v1/days/${day}`).set("Cookie", cookie).expect(200);
    expect(afterDelete.body.entries.some((entry: { id: string }) => entry.id === manual.body.id)).toBe(false);
    expect(afterDelete.body.events.some((event: { action: string }) => event.action === "removed")).toBe(true);
    expect(afterDelete.body.events.some((event: { action: string }) => event.action === "replaced")).toBe(true);
  }, 120_000);
});
