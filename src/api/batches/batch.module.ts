import { Module } from "@nestjs/common";
import { BatchController } from "./batch.controller";
import { BatchService } from "./batch.service";
import { AccessModule } from "../access/access.module";

@Module({
  imports: [AccessModule],
  controllers: [BatchController],
  providers: [BatchService],
})
export class BatchModule {}
