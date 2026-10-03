import { Controller, Get } from "@nestjs/common";

@Controller("health")
export class HealthController {
  @Get()
  check(): { status: "ok" } {
    return { status: "ok" };
  }

  @Get("ready")
  readiness(): { status: "ready"; uptimeSeconds: number } {
    return { status: "ready", uptimeSeconds: Math.floor(process.uptime()) };
  }
}
