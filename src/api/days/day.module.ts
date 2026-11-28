import { Module } from "@nestjs/common";
import { AccessModule } from "../access/access.module";
import { DayController } from "./day.controller";
import { DayService } from "./day.service";
import { SupabaseStore } from "./supabase.store";

@Module({ imports: [AccessModule], controllers: [DayController], providers: [DayService, SupabaseStore] })
export class DayModule {}
