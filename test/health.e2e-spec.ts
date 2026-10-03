import { INestApplication, ValidationPipe } from "@nestjs/common";
import { Test } from "@nestjs/testing";
import request from "supertest";
import { AppModule } from "../src/api/app.module";
import { HttpExceptionFilter } from "../src/api/common/filters/http-exception.filter";
import { parseCorsOrigins } from "../src/api/config/env.validation";

describe("Health endpoint", () => {
  let app: INestApplication;

  beforeAll(async () => {
    const moduleRef = await Test.createTestingModule({ imports: [AppModule] }).compile();
    app = moduleRef.createNestApplication();
    app.enableCors({
      origin: parseCorsOrigins("https://web.homolog.feiranova.example"),
      methods: ["GET", "POST", "OPTIONS"],
      allowedHeaders: ["Accept", "Content-Type", "X-Request-Id"],
      exposedHeaders: ["X-Request-Id"],
      credentials: false,
      maxAge: 600,
    });
    app.useGlobalPipes(new ValidationPipe({ transform: true, whitelist: true }));
    app.useGlobalFilters(new HttpExceptionFilter());
    await app.init();
  });

  afterAll(async () => {
    await app.close();
  });

  it("GET /health", async () => {
    await request(app.getHttpServer())
      .get("/health")
      .expect("x-request-id", /.+/)
      .expect(200)
      .expect({ status: "ok" });
  });

  it("GET /health/ready", async () => {
    await request(app.getHttpServer())
      .get("/health/ready")
      .expect(200)
      .expect(({ body }) => {
        expect(body).toEqual({ status: "ready", uptimeSeconds: expect.any(Number) });
      });
  });

  it("allows the configured web origin", async () => {
    await request(app.getHttpServer())
      .options("/health")
      .set("Origin", "https://web.homolog.feiranova.example")
      .set("Access-Control-Request-Method", "GET")
      .expect("access-control-allow-origin", "https://web.homolog.feiranova.example")
      .expect("access-control-allow-methods", "GET,POST,OPTIONS")
      .expect(204);
  });

  it("does not allow an unconfigured origin", async () => {
    await request(app.getHttpServer())
      .options("/health")
      .set("Origin", "https://untrusted.example")
      .set("Access-Control-Request-Method", "GET")
      .expect((response) => {
        if (response.headers["access-control-allow-origin"]) {
          throw new Error("Unconfigured origin received a CORS allow header");
        }
      })
      .expect(204);
  });
});
