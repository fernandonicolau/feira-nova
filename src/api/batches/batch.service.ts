import { BadRequestException, Injectable } from "@nestjs/common";
import { plainToInstance } from "class-transformer";
import { validate } from "class-validator";
import { randomUUID } from "node:crypto";
import path from "node:path";
import {
  BATCH_LIMITS,
  BatchEntryType,
  ProcessBatchDto,
  type ProcessBatchResponse,
} from "./contracts";

interface CoreArtifact {
  fileName: string;
  mediaType: string;
  buffer: Buffer;
}

interface CoreResult {
  summary: { entries: number; items: number; artifacts: number };
  warnings: Array<{ code: string; message: string; entryId?: string }>;
  artifacts: CoreArtifact[];
}

type CoreProcessBatch = (options: {
  entries: Array<
    | { fileName: string; buffer: Buffer; store: string }
    | { id: string; text: string; store: string }
  >;
}) => Promise<CoreResult>;

@Injectable()
export class BatchService {
  async process(
    rawBatch: string | undefined,
    files: Express.Multer.File[],
    requestId: string,
  ): Promise<ProcessBatchResponse> {
    const batch = await this.parseContract(rawBatch);

    this.validateFiles(batch, files);
    const filesByRef = new Map(files.map((file) => [file.fieldname, file]));
    const processBatch = this.loadCore();

    let result: CoreResult;
    try {
      result = await processBatch({
        entries: batch.entries.map((entry) => {
          if (entry.type === BatchEntryType.TEXT) {
            return { id: entry.id, text: entry.text!, store: entry.store };
          }

          const file = filesByRef.get(entry.fileRef!);
          return { fileName: file!.originalname, buffer: file!.buffer, store: entry.store };
        }),
      });
    } catch (error) {
      throw new BadRequestException(
        error instanceof Error ? `Invalid spreadsheet: ${error.message}` : "Invalid spreadsheet",
      );
    }

    const batchId = randomUUID();
    return {
      requestId,
      data: {
        batchId,
        summary: {
          entries: batch.entries.length,
          files: batch.entries.filter((entry) => entry.type === BatchEntryType.FILE).length,
          texts: batch.entries.filter((entry) => entry.type === BatchEntryType.TEXT).length,
          items: result.summary.items,
          artifacts: result.artifacts.length,
        },
        artifacts: result.artifacts.map((artifact, index) => ({
          id: `artifact-${index + 1}`,
          fileName: artifact.fileName,
          mediaType: artifact.mediaType,
          size: artifact.buffer.length,
          downloadUrl: "/api/v1/batches/download",
        })),
      },
      warnings: result.warnings,
    };
  }

  private async parseContract(rawBatch: string | undefined): Promise<ProcessBatchDto> {
    if (!rawBatch) {
      throw new BadRequestException("Multipart field 'batch' is required");
    }

    let value: unknown;
    try {
      value = JSON.parse(rawBatch);
    } catch {
      throw new BadRequestException("Multipart field 'batch' must be valid JSON");
    }

    const batch = plainToInstance(ProcessBatchDto, value);
    const errors = await validate(batch, {
      whitelist: true,
      forbidNonWhitelisted: true,
      validationError: { target: false, value: false },
    });
    if (errors.length > 0) {
      throw new BadRequestException(errors);
    }

    return batch;
  }

  private validateFiles(batch: ProcessBatchDto, files: Express.Multer.File[]): void {
    const fileEntries = batch.entries.filter((entry) => entry.type === BatchEntryType.FILE);
    if (fileEntries.length > 0 && files.length === 0) {
      throw new BadRequestException("At least one spreadsheet file is required");
    }

    const expectedRefs = new Set(fileEntries.map((entry) => entry.fileRef));
    const receivedRefs = new Set(files.map((file) => file.fieldname));
    const totalBytes = files.reduce((total, file) => total + file.size, 0);

    if (
      expectedRefs.size !== receivedRefs.size ||
      [...expectedRefs].some((reference) => !receivedRefs.has(reference!))
    ) {
      throw new BadRequestException("Uploaded files must match every fileRef exactly");
    }

    if (totalBytes > BATCH_LIMITS.maxTotalFileBytes) {
      throw new BadRequestException("Total upload size exceeds the configured limit");
    }

    for (const file of files) {
      if (!/\.(xlsx|xlsm)$/i.test(file.originalname)) {
        throw new BadRequestException(`Unsupported spreadsheet format: ${file.originalname}`);
      }
      if (!file.buffer?.length) {
        throw new BadRequestException(`Spreadsheet is empty: ${file.originalname}`);
      }
    }
  }

  private loadCore(): CoreProcessBatch {
    const corePath = path.join(process.cwd(), "src", "index.js");
    return (require(corePath) as { processBatch: CoreProcessBatch }).processBatch;
  }
}
