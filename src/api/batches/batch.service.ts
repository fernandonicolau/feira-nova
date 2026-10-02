import { BadRequestException, Injectable } from "@nestjs/common";
import { plainToInstance } from "class-transformer";
import { validate } from "class-validator";
import { randomUUID } from "node:crypto";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import JSZip from "jszip";
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

interface PreparedArtifact extends CoreArtifact {
  folder: "mapas" | "fornecedores";
}

interface SupplierResult {
  files: string[];
  unmatched: unknown[];
  unmatchedFile: { fileName: string } | null;
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
    const prepared = await this.prepare(rawBatch, files, requestId);
    return prepared.response;
  }

  async download(
    rawBatch: string | undefined,
    files: Express.Multer.File[],
    requestId: string,
  ): Promise<{ buffer: Buffer; fileName: string }> {
    const prepared = await this.prepare(rawBatch, files, requestId);
    const zip = new JSZip();
    for (const artifact of prepared.artifacts) {
      zip.file(`${artifact.folder}/${artifact.fileName}`, artifact.buffer);
    }
    zip.file(
      "manifest.json",
      JSON.stringify(
        {
          requestId,
          data: prepared.response.data,
          warnings: prepared.response.warnings,
        },
        null,
        2,
      ),
    );

    return {
      buffer: await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" }),
      fileName: `feira-nova-${prepared.response.data.batchId}.zip`,
    };
  }

  private async prepare(
    rawBatch: string | undefined,
    files: Express.Multer.File[],
    requestId: string,
  ): Promise<{ response: ProcessBatchResponse; artifacts: PreparedArtifact[] }> {
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

    const supplierArtifacts = await this.generateSupplierArtifacts(result.artifacts);
    const artifacts: PreparedArtifact[] = [
      ...result.artifacts.map((artifact) => ({ ...artifact, folder: "mapas" as const })),
      ...supplierArtifacts.artifacts,
    ];
    const warnings = [...result.warnings];
    if (supplierArtifacts.unmatched > 0) {
      warnings.push({
        code: "UNMATCHED_SUPPLIER_ASSOCIATIONS",
        message: `${supplierArtifacts.unmatched} associações não foram atribuídas a fornecedores.`,
      });
    }

    const batchId = randomUUID();
    const response: ProcessBatchResponse = {
      requestId,
      data: {
        batchId,
        summary: {
          entries: batch.entries.length,
          files: batch.entries.filter((entry) => entry.type === BatchEntryType.FILE).length,
          texts: batch.entries.filter((entry) => entry.type === BatchEntryType.TEXT).length,
          items: result.summary.items,
          artifacts: artifacts.length,
        },
        artifacts: artifacts.map((artifact, index) => ({
          id: `artifact-${index + 1}`,
          fileName: artifact.fileName,
          mediaType: artifact.mediaType,
          size: artifact.buffer.length,
          downloadUrl: "/api/v1/batches/process/download",
        })),
      },
      warnings,
    };
    return { response, artifacts };
  }

  private async generateSupplierArtifacts(
    mapArtifacts: CoreArtifact[],
  ): Promise<{ artifacts: PreparedArtifact[]; unmatched: number }> {
    const requestDir = await fs.promises.mkdtemp(path.join(os.tmpdir(), "feira-nova-"));
    const mapDir = path.join(requestDir, "maps");
    const outputDir = path.join(requestDir, "fornecedores");
    const templateDir = path.join(process.cwd(), "template", "fornecedores");

    try {
      await fs.promises.mkdir(mapDir);
      await Promise.all(
        mapArtifacts.map((artifact) =>
          fs.promises.writeFile(path.join(mapDir, artifact.fileName), artifact.buffer),
        ),
      );
      const generatorPath = path.join(process.cwd(), "scripts", "generate-fornecedores.js");
      const { generateSupplierFiles } = require(generatorPath) as {
        generateSupplierFiles(options: {
          mapDir: string;
          templateDir: string;
          outputDir: string;
          now: Date;
        }): Promise<SupplierResult>;
      };
      const generated = await generateSupplierFiles({
        mapDir,
        templateDir,
        outputDir,
        now: new Date(),
      });
      const fileNames = await fs.promises.readdir(outputDir);
      const artifacts = await Promise.all(
        fileNames
          .filter((fileName) => fileName.toLowerCase().endsWith(".xlsx"))
          .sort((a, b) => a.localeCompare(b, "pt-BR"))
          .map(async (fileName): Promise<PreparedArtifact> => ({
            fileName,
            mediaType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            buffer: await fs.promises.readFile(path.join(outputDir, fileName)),
            folder: "fornecedores",
          })),
      );
      return { artifacts, unmatched: generated.unmatched.length };
    } finally {
      await fs.promises.rm(requestDir, { recursive: true, force: true });
    }
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
