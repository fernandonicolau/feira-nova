import "reflect-metadata";
import { Logger, ValidationPipe } from "@nestjs/common";
import { ConfigService } from "@nestjs/config";
import { NestFactory } from "@nestjs/core";
import { AppModule } from "./app.module";
import { HttpExceptionFilter } from "./common/filters/http-exception.filter";
import { parseCorsOrigins } from "./config/env.validation";

export async function bootstrap(): Promise<void> {
  const app = await NestFactory.create(AppModule);
  const config = app.get(ConfigService);

  app.enableCors({
    origin: parseCorsOrigins(config.getOrThrow<string>("CORS_ORIGINS")),
    methods: ["GET", "POST", "OPTIONS"],
    allowedHeaders: ["Accept", "Content-Type", "X-Request-Id"],
    exposedHeaders: ["X-Request-Id"],
    credentials: false,
    maxAge: 600,
  });

  app.useGlobalPipes(
    new ValidationPipe({
      transform: true,
      whitelist: true,
      forbidNonWhitelisted: true,
    }),
  );
  app.useGlobalFilters(new HttpExceptionFilter());
  app.enableShutdownHooks();

  const port = config.get<number>("PORT", 3000);
  await app.listen(port, "0.0.0.0");
  Logger.log(`API listening on port ${port}`, "Bootstrap");
}

if (require.main === module) {
  void bootstrap();
}
