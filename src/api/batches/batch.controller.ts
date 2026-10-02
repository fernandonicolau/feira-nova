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
import { ApiBody, ApiConsumes, ApiCreatedResponse, ApiProduces, ApiTags } from "@nestjs/swagger";
import type { RequestWithId } from "../common/middleware/request-id.middleware";
import { BatchService } from "./batch.service";
import { BATCH_LIMITS, type ProcessBatchResponse } from "./contracts";

const multipartBatchSchema: Parameters<typeof ApiBody>[0] = {
  schema: {
    type: "object",
    required: ["batch"],
    properties: {
      batch: {
        type: "string",
        description: "Contrato JSON do lote; entradas file referenciam o nome do campo multipart.",
        example: JSON.stringify({
          name: "Pedido da manhã",
          entries: [{ id: "manual-1", type: "text", store: "Cerâmica", text: "ABACATE 5" }],
        }),
      },
      files: { type: "array", items: { type: "string", format: "binary" } },
    },
  },
};

@ApiTags("batches")
@Controller("api/v1/batches")
export class BatchController {
  constructor(private readonly batchService: BatchService) {}

  @Post("process")
  @ApiConsumes("multipart/form-data")
  @ApiBody(multipartBatchSchema)
  @ApiCreatedResponse({ description: "Resumo, avisos e manifesto de todos os artefatos." })
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
  @ApiConsumes("multipart/form-data")
  @ApiProduces("application/zip")
  @ApiBody(multipartBatchSchema)
  @ApiCreatedResponse({ description: "ZIP com mapas, fornecedores, pendências e manifest.json." })
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
