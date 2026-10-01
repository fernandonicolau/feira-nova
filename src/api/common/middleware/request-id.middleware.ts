import { Injectable, Logger, NestMiddleware } from "@nestjs/common";
import type { NextFunction, Request, Response } from "express";
import { randomUUID } from "node:crypto";

export type RequestWithId = Request & { requestId: string };

@Injectable()
export class RequestIdMiddleware implements NestMiddleware {
  private readonly logger = new Logger(RequestIdMiddleware.name);

  use(request: RequestWithId, response: Response, next: NextFunction): void {
    const startedAt = process.hrtime.bigint();
    const providedRequestId = request.header("x-request-id")?.trim();
    request.requestId = providedRequestId || randomUUID();
    response.setHeader("x-request-id", request.requestId);

    response.once("finish", () => {
      const durationMs = Number(process.hrtime.bigint() - startedAt) / 1_000_000;
      this.logger.log(
        `requestId=${request.requestId} method=${request.method} path=${request.path} status=${response.statusCode} durationMs=${durationMs.toFixed(1)}`,
      );
    });

    next();
  }
}
