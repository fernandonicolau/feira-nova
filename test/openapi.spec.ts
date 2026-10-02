import { Test } from "@nestjs/testing";
import { DocumentBuilder, SwaggerModule } from "@nestjs/swagger";
import { AppModule } from "../src/api/app.module";

describe("OpenAPI contract", () => {
  it("publishes health, processing and ZIP download operations", async () => {
    const moduleRef = await Test.createTestingModule({ imports: [AppModule] }).compile();
    const app = moduleRef.createNestApplication();
    await app.init();
    try {
      const document = SwaggerModule.createDocument(
        app,
        new DocumentBuilder().setTitle("Feira Nova API").setVersion("1.0").build(),
      );
      expect(document.paths["/health"]?.get).toBeDefined();
      expect(document.paths["/api/v1/batches/process"]?.post).toBeDefined();
      const download = document.paths["/api/v1/batches/process/download"]?.post;
      expect(download).toBeDefined();
      expect(download?.requestBody).toBeDefined();
      expect(download?.responses[201]).toBeDefined();
    } finally {
      await app.close();
    }
  });
});
