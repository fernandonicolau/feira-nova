import { CanActivate, ExecutionContext, ForbiddenException, Injectable, UnauthorizedException } from "@nestjs/common";
import type { Request } from "express";
import { AccessService } from "./access.service";

export function sessionCookie(request: Request): string | undefined {
  const item = request.headers.cookie?.split(";").map((part) => part.trim()).find((part) => part.startsWith("feira_session="));
  return item?.slice("feira_session=".length);
}

@Injectable()
export class AccessGuard implements CanActivate {
  constructor(private readonly access: AccessService) {}

  canActivate(context: ExecutionContext): boolean {
    const request = context.switchToHttp().getRequest<Request>();
    if (!this.access.valid(sessionCookie(request))) {
      throw new UnauthorizedException("Informe o código de acesso");
    }
    const origin = request.headers.origin;
    if (origin && !String(process.env.CORS_ORIGINS ?? "http://localhost:5173").split(",").map((item) => item.trim().replace(/\/$/, "")).includes(origin)) {
      throw new ForbiddenException("Origem não permitida");
    }
    return true;
  }
}

@Injectable()
export class LegacyAccessGuard extends AccessGuard {
  canActivate(context: ExecutionContext): boolean {
    if (!process.env.ACCESS_CODE) return true;
    return super.canActivate(context);
  }
}
