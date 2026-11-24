import { Module } from "@nestjs/common";
import { AccessController } from "./access.controller";
import { AccessGuard, LegacyAccessGuard } from "./access.guard";
import { AccessService } from "./access.service";

@Module({ controllers: [AccessController], providers: [AccessService, AccessGuard, LegacyAccessGuard], exports: [AccessService, AccessGuard, LegacyAccessGuard] })
export class AccessModule {}
