import {
  Body,
  Controller,
  Post,
  Req,
  UploadedFiles,
  UseInterceptors,
} from "@nestjs/common";
import { AnyFilesInterceptor } from "@nestjs/platform-express";
import type { RequestWithId } from "../common/middleware/request-id.middleware";
import { BatchService } from "./batch.service";
import { BATCH_LIMITS, type ProcessBatchResponse } from "./contracts";

@Controller("api/v1/batches")
export class BatchController {
  constructor(private readonly batchService: BatchService) {}

  @Post("process")
  @UseInterceptors(
    AnyFilesInterceptor({
      limits: {
        files: BATCH_LIMITS.maxFiles,
        fileSize: BATCH_LIMITS.maxFileBytes,
        fields: 1,
      },
    }),
  )
  process(
    @Body("batch") batch: string | undefined,
    @UploadedFiles() files: Express.Multer.File[] = [],
    @Req() request: RequestWithId,
  ): Promise<ProcessBatchResponse> {
    return this.batchService.process(batch, files, request.requestId);
  }
}
