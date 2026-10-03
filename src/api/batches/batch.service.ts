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
  warnings: Array<{ code: string; message: string; entryId?: string; line?: number; originalText?: string }>;
  artifacts: CoreArtifact[];
}

type CoreProcessBatch = (options: {
  entries: Array<
    | { fileName: string; buffer: Buffer; store: string }
    | { id: string; text: string; store: string }
  >;
}) => Promise<CoreResult>;

type CoreEntry = Parameters<CoreProcessBatch>[0]["entries"][number];

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
    const coreEntries = await this.expandEntries(batch, filesByRef);

    let result: CoreResult;
    try {
      result = await processBatch({
        entries: coreEntries,
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
          entries: coreEntries.length,
          files: coreEntries.filter((entry) => "buffer" in entry).length,
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
      if (!/\.(xlsx|xlsm|zip)$/i.test(file.originalname)) {
        throw new BadRequestException(`Unsupported upload format: ${file.originalname}`);
      }
      if (!file.buffer?.length) {
        throw new BadRequestException(`Spreadsheet is empty: ${file.originalname}`);
      }
    }
  }

  private async expandEntries(
    batch: ProcessBatchDto,
    filesByRef: Map<string, Express.Multer.File>,
  ): Promise<CoreEntry[]> {
    const entries: CoreEntry[] = [];
    let expandedBytes = 0;

    for (const entry of batch.entries) {
      if (entry.type === BatchEntryType.TEXT) {
        entries.push({ id: entry.id, text: entry.text!, store: entry.store });
        continue;
      }

      const upload = filesByRef.get(entry.fileRef!)!;
      if (!/\.zip$/i.test(upload.originalname)) {
        entries.push({ fileName: upload.originalname, buffer: upload.buffer, store: entry.store });
        expandedBytes += upload.buffer.length;
        continue;
      }

      let zip: JSZip;
      try {
        zip = await JSZip.loadAsync(upload.buffer);
      } catch {
        throw new BadRequestException(`Invalid ZIP file: ${upload.originalname}`);
      }

      const members = Object.values(zip.files).filter((member) => !member.dir);
      if (members.length === 0) throw new BadRequestException(`ZIP is empty: ${upload.originalname}`);
      for (const member of members) {
        const originalName = member.unsafeOriginalName ?? member.name;
        if (
          originalName.startsWith("/") ||
          originalName.startsWith("\\") ||
          originalName.split(/[\\/]/).includes("..")
        ) {
          throw new BadRequestException(`Unsafe ZIP path: ${originalName}`);
        }
        if (!/\.(xlsx|xlsm)$/i.test(member.name)) {
          throw new BadRequestException(`Unsupported file inside ZIP: ${member.name}`);
        }
        const buffer = await member.async("nodebuffer");
        if (!buffer.length || buffer.length > BATCH_LIMITS.maxFileBytes) {
          throw new BadRequestException(`Invalid file size inside ZIP: ${member.name}`);
        }
        expandedBytes += buffer.length;
        entries.push({ fileName: path.basename(member.name), buffer, store: entry.store });
      }
    }

    const fileCount = entries.filter((entry) => "buffer" in entry).length;
    if (fileCount > BATCH_LIMITS.maxFiles || expandedBytes > BATCH_LIMITS.maxTotalFileBytes) {
      throw new BadRequestException("Expanded ZIP contents exceed the configured limits");
    }
    return entries;
  }

  private loadCore(): CoreProcessBatch {
    const corePath = path.join(process.cwd(), "src", "index.js");
    return (require(corePath) as { processBatch: CoreProcessBatch }).processBatch;
  }
}
