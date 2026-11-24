import { INestApplication } from "@nestjs/common";
import { Test } from "@nestjs/testing";
import request from "supertest";
import { AppModule } from "../src/api/app.module";
import { HttpExceptionFilter } from "../src/api/common/filters/http-exception.filter";

describe("Acesso ao produto", () => {
  let app: INestApplication;
  const old = { code: process.env.ACCESS_CODE, secret: process.env.SESSION_SECRET, origins: process.env.CORS_ORIGINS };

  beforeAll(async () => {
    process.env.ACCESS_CODE = "123456";
    process.env.SESSION_SECRET = "a-long-secret-at-least-thirty-two-characters";
    process.env.CORS_ORIGINS = "http://localhost:5173";
    const moduleRef = await Test.createTestingModule({ imports: [AppModule] }).compile();
    app = moduleRef.createNestApplication();
    app.useGlobalFilters(new HttpExceptionFilter());
    await app.init();
  });

  afterAll(async () => {
    await app.close();
    for (const [key, value] of Object.entries({ ACCESS_CODE: old.code, SESSION_SECRET: old.secret, CORS_ORIGINS: old.origins })) {
      if (value === undefined) delete process.env[key]; else process.env[key] = value;
    }
  });

  it("keeps health public and blocks the batch endpoints without a session", async () => {
    await request(app.getHttpServer()).get("/health").expect(200);
    await request(app.getHttpServer()).post("/api/v1/batches/process").expect(401);
    await request(app.getHttpServer()).post("/api/v1/batches/process/download").expect(401);
  });

  it("creates a cookie, accepts it and clears it on logout", async () => {
    const login = await request(app.getHttpServer()).post("/api/v1/access/session").send({ code: "123456" }).expect(201);
    const cookie = login.headers["set-cookie"]?.[0];
    expect(cookie).toContain("HttpOnly");
    expect(cookie).not.toContain("123456");
    await request(app.getHttpServer()).get("/api/v1/access/session").set("Cookie", cookie).expect(200);
    await request(app.getHttpServer()).post("/api/v1/batches/process").set("Cookie", cookie).expect(400);
    await request(app.getHttpServer()).post("/api/v1/batches/process").set("Cookie", cookie).set("Origin", "https://other.example").expect(403);
    const logout = await request(app.getHttpServer()).delete("/api/v1/access/session").set("Cookie", cookie).expect(200);
    expect(logout.headers["set-cookie"]?.[0]).toContain("Max-Age=0");
  });
});
