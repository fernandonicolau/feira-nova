import { Injectable, NestMiddleware } from "@nestjs/common";
import type { NextFunction, Request, Response } from "express";
import { randomUUID } from "node:crypto";

export type RequestWithId = Request & { requestId: string };

@Injectable()
export class RequestIdMiddleware implements NestMiddleware {
  use(request: RequestWithId, response: Response, next: NextFunction): void {
    const providedRequestId = request.header("x-request-id")?.trim();
    request.requestId = providedRequestId || randomUUID();
    response.setHeader("x-request-id", request.requestId);
    next();
  }
}
