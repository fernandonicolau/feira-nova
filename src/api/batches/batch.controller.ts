import {
  Body,
  Controller,
  Post,
  Req,
  Res,
  StreamableFile,
  UploadedFiles,
  UseInterceptors,
} from "@nestjs/common";
import { AnyFilesInterceptor } from "@nestjs/platform-express";
import type { Response } from "express";
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

  @Post("process/download")
  @UseInterceptors(
    AnyFilesInterceptor({
      limits: {
        files: BATCH_LIMITS.maxFiles,
        fileSize: BATCH_LIMITS.maxFileBytes,
        fields: 1,
      },
    }),
  )
  async download(
    @Body("batch") batch: string | undefined,
    @UploadedFiles() files: Express.Multer.File[] = [],
    @Req() request: RequestWithId,
    @Res({ passthrough: true }) response: Response,
  ): Promise<StreamableFile> {
    const archive = await this.batchService.download(batch, files, request.requestId);
    response.set({
      "Content-Type": "application/zip",
      "Content-Disposition": `attachment; filename="${archive.fileName}"`,
      "X-Request-Id": request.requestId,
    });
    return new StreamableFile(archive.buffer);
  }
}
