import { Body, Controller, Delete, Get, Post, Req, Res, UnauthorizedException } from "@nestjs/common";
import type { Request, Response } from "express";
import { sessionCookie } from "./access.guard";
import { AccessService } from "./access.service";

@Controller("api/v1/access/session")
export class AccessController {
  constructor(private readonly access: AccessService) {}

  @Post()
  login(@Body("code") code: string, @Req() request: Request, @Res({ passthrough: true }) response: Response): { authenticated: true } {
    const token = this.access.login(code, request.ip || request.socket.remoteAddress || "unknown");
    response.setHeader("Set-Cookie", this.access.cookie(token));
    return { authenticated: true };
  }

  @Get()
  session(@Req() request: Request): { authenticated: true } {
    if (!this.access.valid(sessionCookie(request))) throw new UnauthorizedException("Informe o código de acesso");
    return { authenticated: true };
  }

  @Delete()
  logout(@Res({ passthrough: true }) response: Response): { authenticated: false } {
    response.setHeader("Set-Cookie", this.access.cookie("", true));
    return { authenticated: false };
  }
}
