import { Body, Controller, Delete, Get, Param, Patch, Post, Put, Res, StreamableFile, UploadedFiles, UseGuards, UseInterceptors } from "@nestjs/common";
import { AnyFilesInterceptor } from "@nestjs/platform-express";
import { ApiConsumes, ApiTags } from "@nestjs/swagger";
import type { Response } from "express";
import { AccessGuard } from "../access/access.guard";
import { BATCH_LIMITS } from "../batches/contracts";
import { DayService } from "./day.service";
import { KNOWN_STORES, type ColumnLayout } from "./file-recognition";

@ApiTags("days")
@Controller("api/v1/days")
@UseGuards(AccessGuard)
export class DayController {
  constructor(private readonly days: DayService) {}

  @Get()
  list() { return this.days.list(); }

  @Get("stores")
  stores() { return KNOWN_STORES; }

  @Put(":day")
  create(@Param("day") day: string) { return this.days.create(day); }

  @Get(":day")
  get(@Param("day") day: string) { return this.days.get(day); }

  @Post(":day/files")
  @ApiConsumes("multipart/form-data")
  @UseInterceptors(AnyFilesInterceptor({ limits: { files: BATCH_LIMITS.maxFiles, fileSize: BATCH_LIMITS.maxFileBytes, fieldSize: 256, fields: 1, parts: BATCH_LIMITS.maxFiles + 1 } }))
  files(@Param("day") day: string, @Body("clientUploadId") clientUploadId: string, @UploadedFiles() files: Express.Multer.File[]) {
    return this.days.addFiles(day, files ?? [], clientUploadId);
  }

  @Post(":day/entries/:id/replace")
  @ApiConsumes("multipart/form-data")
  @UseInterceptors(AnyFilesInterceptor({ limits: { files: 1, fileSize: BATCH_LIMITS.maxFileBytes, fieldSize: 256, fields: 1, parts: 2 } }))
  replace(@Param("day") day: string, @Param("id") id: string, @Body("clientUploadId") clientUploadId: string, @UploadedFiles() files: Express.Multer.File[]) {
    return this.days.addFiles(day, files ?? [], clientUploadId, id);
  }

  @Post(":day/texts")
  text(@Param("day") day: string, @Body("text") text: string, @Body("clientEntryId") clientEntryId: string, @Body("store") store?: string) {
    return this.days.addText(day, text, clientEntryId, store);
  }

  @Delete(":day/entries/:id")
  async remove(@Param("day") day: string, @Param("id") id: string) {
    await this.days.removeEntry(day, id);
    return { removed: true };
  }

  @Patch(":day/entries/:id/review")
  review(@Param("day") day: string, @Param("id") id: string, @Body() body: { store?: string; layout?: ColumnLayout }) {
    return this.days.reviewEntry(day, id, body);
  }

  @Get(":day/uploads/:id/original")
  async original(@Param("day") day: string, @Param("id") id: string, @Res({ passthrough: true }) response: Response): Promise<StreamableFile> {
    const original = await this.days.downloadUpload(day, id);
    response.set({
      "Content-Type": original.contentType,
      "Content-Disposition": `attachment; filename*=UTF-8''${encodeURIComponent(original.fileName)}`,
    });
    return new StreamableFile(original.buffer);
  }
}
