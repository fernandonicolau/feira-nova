import {
  ArgumentsHost,
  Catch,
  ExceptionFilter,
  HttpException,
  HttpStatus,
  Logger,
} from "@nestjs/common";
import type { Response } from "express";
import type { RequestWithId } from "../middleware/request-id.middleware";

@Catch()
export class HttpExceptionFilter implements ExceptionFilter {
  private readonly logger = new Logger(HttpExceptionFilter.name);

  catch(exception: unknown, host: ArgumentsHost): void {
    const context = host.switchToHttp();
    const request = context.getRequest<RequestWithId>();
    const response = context.getResponse<Response>();
    const status =
      exception instanceof HttpException
        ? exception.getStatus()
        : HttpStatus.INTERNAL_SERVER_ERROR;
    const exceptionResponse =
      exception instanceof HttpException ? exception.getResponse() : undefined;

    if (!(exception instanceof HttpException)) {
      this.logger.error(
        `requestId=${request.requestId} ${request.method} ${request.url}`,
        exception instanceof Error ? exception.stack : undefined,
      );
    }

    response.status(status).json({
      requestId: request.requestId,
      error: {
        code: this.resolveCode(status),
        message: this.resolveMessage(exceptionResponse, status),
        details: this.resolveDetails(exceptionResponse),
      },
    });
  }

  private resolveCode(status: number): string {
    if (status === HttpStatus.BAD_REQUEST) return "VALIDATION_ERROR";
    if (status === HttpStatus.NOT_FOUND) return "NOT_FOUND";
    return status >= 500 ? "INTERNAL_ERROR" : "REQUEST_ERROR";
  }

  private resolveMessage(value: unknown, status: number): string {
    if (typeof value === "string") return value;
    if (value && typeof value === "object" && "message" in value) {
      const message = (value as { message: unknown }).message;
      if (typeof message === "string") return message;
    }
    return status >= 500 ? "Internal server error" : "Request failed";
  }

  private resolveDetails(value: unknown): unknown[] {
    if (value && typeof value === "object" && "message" in value) {
      const message = (value as { message: unknown }).message;
      return Array.isArray(message) ? message : [];
    }
    return [];
  }
}
