import { BadRequestException, ConflictException, Injectable, NotFoundException, ServiceUnavailableException } from "@nestjs/common";
import { createHash, randomUUID } from "node:crypto";
import { BATCH_LIMITS } from "../batches/contracts";
import { SupabaseStore } from "./supabase.store";
import { planUpload, safeFileName } from "./upload-plan";
import { inspectFile, resolveStore, validateLayout, type ColumnLayout, type FileInspection } from "./file-recognition";

const STORAGE_CEILING = 900 * 1024 * 1024;

export interface DayRecord { id: string; day: string; revision: number; submissions_closed: boolean; created_at: string }
export interface DayEntry { id: string; day_id: string; kind: "file" | "text"; file_name: string | null; source_file_name?: string | null; byte_size: number | null; sha256: string | null; manual_text: string | null; store: string | null; detected_store: string | null; detection_source: string | null; detection_confidence: number | null; layout: ColumnLayout | null; preview: unknown; review_required: boolean; state: string; upload_id: string | null; created_at: string; deleted_at: string | null }

function assertDay(day: string): void {
  const parsed = /^\d{4}-\d{2}-\d{2}$/.test(day) ? new Date(`${day}T00:00:00.000Z`) : null;
  if (!parsed || Number.isNaN(parsed.getTime()) || parsed.toISOString().slice(0, 10) !== day) {
    throw new BadRequestException("Data inválida; use AAAA-MM-DD");
  }
}

function failed(): ServiceUnavailableException { return new ServiceUnavailableException("Não foi possível acessar os dados do dia"); }

@Injectable()
export class DayService {
  constructor(private readonly storage: SupabaseStore) {}

  async list(): Promise<DayRecord[]> {
    const { data, error } = await this.storage.db.from("operation_days").select("id,day,revision,submissions_closed,created_at").order("day", { ascending: false }).limit(90);
    if (error) throw failed();
    return (data ?? []) as DayRecord[];
  }

  async create(day: string): Promise<DayRecord> {
    assertDay(day);
    const existing = await this.find(day);
    if (existing) return existing;
    const { data, error } = await this.storage.db.from("operation_days").insert({ day }).select("id,day,revision,submissions_closed,created_at").single();
    if (error?.code === "23505") return (await this.find(day))!;
    if (error || !data) throw failed();
    return data as DayRecord;
  }

  async get(day: string): Promise<{ day: DayRecord; entries: DayEntry[]; events: unknown[] }> {
    const record = await this.required(day);
    const [entryResult, eventResult, uploadResult] = await Promise.all([
      this.storage.db.from("day_entries").select("id,day_id,kind,file_name,byte_size,sha256,manual_text,store,detected_store,detection_source,detection_confidence,layout,preview,review_required,state,upload_id,created_at,deleted_at").eq("day_id", record.id).is("deleted_at", null).order("created_at"),
      this.storage.db.from("day_events").select("id,entry_id,action,details,created_at").eq("day_id", record.id).order("created_at", { ascending: false }).limit(30),
      this.storage.db.from("day_uploads").select("id,file_name").eq("day_id", record.id),
    ]);
    if (entryResult.error || eventResult.error || uploadResult.error) throw failed();
    const names = new Map((uploadResult.data ?? []).map((upload) => [upload.id, upload.file_name]));
    const entries = ((entryResult.data ?? []) as DayEntry[]).map((entry) => ({ ...entry, source_file_name: entry.upload_id ? names.get(entry.upload_id) ?? null : null }));
    return { day: record, entries, events: eventResult.data ?? [] };
  }

  async addFiles(day: string, files: Express.Multer.File[], clientUploadId: string, replaceEntryId?: string): Promise<{ entries: DayEntry[]; failures: Array<{ fileName: string; message: string }> }> {
    const record = await this.required(day);
    if (!/^[a-zA-Z0-9_-]{8,100}$/.test(clientUploadId)) throw new BadRequestException("Identificador de envio inválido");
    if (!files.length || files.length > BATCH_LIMITS.maxFiles || files.reduce((sum, file) => sum + file.size, 0) > BATCH_LIMITS.maxTotalFileBytes) {
      throw new BadRequestException("Quantidade ou tamanho de arquivos excedido");
    }
    if (replaceEntryId && files.length !== 1) throw new BadRequestException("Substitua apenas um arquivo por vez");
    if (replaceEntryId) {
      const { data: oldEntry, error } = await this.storage.db.from("day_entries").select("id,kind").eq("id", replaceEntryId).eq("day_id", record.id).is("deleted_at", null).maybeSingle();
      if (error) throw failed();
      if (!oldEntry || oldEntry.kind !== "file") throw new NotFoundException("Arquivo original não encontrado para substituição");
    }
    const entries: DayEntry[] = [];
    const failures: Array<{ fileName: string; message: string }> = [];
    for (const [index, file] of files.entries()) {
      try {
        const saved = await this.addOneFile(record, file, `${clientUploadId}-${index}`);
        entries.push(...saved);
        if (replaceEntryId) {
          const { data, error } = await this.storage.db.from("day_entries").update({ deleted_at: new Date().toISOString(), replaced_by: saved[0].id }).eq("id", replaceEntryId).eq("day_id", record.id).is("deleted_at", null).select("id").maybeSingle();
          if (error) throw failed();
          if (!data) throw new NotFoundException("Arquivo original não encontrado para substituição");
        }
      } catch (error) {
        failures.push({ fileName: file.originalname || "arquivo", message: error instanceof BadRequestException || error instanceof ConflictException || error instanceof NotFoundException ? error.message : "Falha ao salvar; tente novamente" });
      }
    }
    return { entries, failures };
  }

  async addText(day: string, text: string, clientEntryId: string, store?: string): Promise<DayEntry> {
    const record = await this.required(day);
    const content = String(text ?? "").trim();
    if (!content || content.length > BATCH_LIMITS.maxTextCharacters) throw new BadRequestException("Texto vazio ou grande demais");
    if (!/^[a-zA-Z0-9_-]{8,100}$/.test(clientEntryId)) throw new BadRequestException("Identificador de entrada inválido");
    const normalizedStore = store?.trim() || null;
    if (normalizedStore && normalizedStore.length > 120) throw new BadRequestException("Nome da loja grande demais");
    const resolvedStore = normalizedStore ? resolveStore(normalizedStore) : null;
    const { data: existing, error: lookupError } = await this.storage.db.from("day_entries").select("*").eq("day_id", record.id).eq("client_entry_id", clientEntryId).maybeSingle();
    if (lookupError) throw failed();
    if (existing) {
      if (existing.manual_text !== content || existing.detected_store !== normalizedStore) throw new ConflictException("Identificador já usado para outro texto");
      return existing as DayEntry;
    }
    const { data, error } = await this.storage.db.from("day_entries").insert({ day_id: record.id, kind: "text", manual_text: content, store: resolvedStore, detected_store: normalizedStore, detection_source: resolvedStore ? "manual" : "none", detection_confidence: resolvedStore ? 1 : 0, client_entry_id: clientEntryId, state: resolvedStore ? "received" : "needs_review", review_required: !resolvedStore, preview: { itemCount: content.split(/\r?\n/).filter(Boolean).length, warnings: resolvedStore ? [] : ["Selecione uma loja antes de processar."] } }).select("*").single();
    if (error?.code === "23505") return this.addText(day, text, clientEntryId, store);
    if (error || !data) throw failed();
    return data as DayEntry;
  }

  async removeEntry(day: string, id: string): Promise<void> {
    const record = await this.required(day);
    const { data, error } = await this.storage.db.from("day_entries").update({ deleted_at: new Date().toISOString() }).eq("id", id).eq("day_id", record.id).is("deleted_at", null).select("id").maybeSingle();
    if (error) throw failed();
    if (!data) throw new NotFoundException("Entrada não encontrada");
  }

  async reviewEntry(day: string, id: string, input: { store?: string; layout?: ColumnLayout }): Promise<DayEntry> {
    const record = await this.required(day);
    if (!input || (input.store === undefined && input.layout === undefined)) throw new BadRequestException("Informe loja ou colunas para corrigir");
    const { data: entry, error: lookupError } = await this.storage.db.from("day_entries").select("*").eq("id", id).eq("day_id", record.id).is("deleted_at", null).maybeSingle();
    if (lookupError) throw failed();
    if (!entry) throw new NotFoundException("Entrada não encontrada");
    const correctedStore = input.store !== undefined ? resolveStore(input.store) : entry.store;
    if (input.store !== undefined && !correctedStore) throw new BadRequestException("Loja não cadastrada");
    if (entry.kind === "text") {
      if (input.layout !== undefined) throw new BadRequestException("Texto manual não usa colunas");
      const { data, error } = await this.storage.db.from("day_entries").update({ store: correctedStore, state: correctedStore ? "received" : "needs_review", review_required: !correctedStore, detection_source: "manual", detection_confidence: correctedStore ? 1 : 0, preview: { ...(entry.preview || {}), warnings: correctedStore ? [] : ["Selecione uma loja antes de processar."] }, reviewed_at: new Date().toISOString() }).eq("id", id).select("*").single();
      if (error || !data) throw failed();
      return data as DayEntry;
    }
    const layout = input.layout !== undefined ? validateLayout(input.layout) : entry.layout as ColumnLayout | null;
    let inspection: FileInspection;
    try {
      inspection = await inspectFile(entry.file_name, await this.storage.get(entry.object_path), { store: correctedStore || undefined, layout: layout || undefined });
    } catch (error) {
      if (error instanceof BadRequestException) throw error;
      throw new BadRequestException("Não foi possível ler esta planilha; substitua o arquivo");
    }
    const { data, error } = await this.storage.db.from("day_entries").update({ store: inspection.store, layout: inspection.layout, preview: inspection.preview, review_required: inspection.needsReview, state: inspection.needsReview ? "needs_review" : "received", detection_source: inspection.source, detection_confidence: inspection.confidence, reviewed_at: new Date().toISOString() }).eq("id", id).select("*").single();
    if (error || !data) throw failed();
    return data as DayEntry;
  }

  async downloadUpload(day: string, uploadId: string): Promise<{ buffer: Buffer; fileName: string; contentType: string }> {
    const record = await this.required(day);
    const { data, error } = await this.storage.db.from("day_uploads").select("file_name,object_path").eq("id", uploadId).eq("day_id", record.id).maybeSingle();
    if (error) throw failed();
    if (!data) throw new NotFoundException("Arquivo original não encontrado");
    return { buffer: await this.storage.get(data.object_path), fileName: safeFileName(data.file_name), contentType: /\.zip$/i.test(data.file_name) ? "application/zip" : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" };
  }

  private async find(day: string): Promise<DayRecord | null> {
    assertDay(day);
    const { data, error } = await this.storage.db.from("operation_days").select("id,day,revision,submissions_closed,created_at").eq("day", day).maybeSingle();
    if (error) throw failed();
    return data as DayRecord | null;
  }

  private async required(day: string): Promise<DayRecord> {
    const found = await this.find(day);
    if (!found) throw new NotFoundException("Dia não encontrado");
    return found;
  }

  private async addOneFile(day: DayRecord, file: Express.Multer.File, clientId: string): Promise<DayEntry[]> {
    const name = safeFileName(file.originalname);
    const originalHash = createHash("sha256").update(file.buffer).digest("hex");
    const { data: prior, error: lookupError } = await this.storage.db.from("day_uploads").select("id,file_name,sha256,byte_size").eq("day_id", day.id).eq("client_upload_id", clientId).maybeSingle();
    if (lookupError) throw failed();
    if (prior) {
      if (prior.sha256 !== originalHash || prior.file_name !== name || prior.byte_size !== file.buffer.length) throw new ConflictException("Identificador já usado para outro arquivo");
      const { data, error } = await this.storage.db.from("day_entries").select("*").eq("upload_id", prior.id).is("deleted_at", null).order("created_at");
      if (error) throw failed();
      return (data ?? []) as DayEntry[];
    }
    const parts = await planUpload(name, file.buffer);
    const used = await this.storage.db.rpc("daily_storage_bytes");
    if (used.error) throw failed();
    const projected = file.buffer.length + (/\.zip$/i.test(name) ? parts.reduce((sum, part) => sum + part.buffer.length, 0) : 0);
    if (Number(used.data ?? 0) + projected > STORAGE_CEILING) throw new BadRequestException("Limite de armazenamento atingido; procure o responsável pela operação");

    const uploadId = randomUUID();
    const originalPath = `days/${day.day}/uploads/${uploadId}/original`;
    const stored: string[] = [];
    let metadataSaved = false;
    try {
      await this.storage.put(originalPath, file.buffer, /\.zip$/i.test(name) ? "application/zip" : "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      stored.push(originalPath);
      const rows = [];
      for (const part of parts) {
        const memberPath = /\.zip$/i.test(name) ? `days/${day.day}/uploads/${uploadId}/members/${part.id}` : originalPath;
        if (memberPath !== originalPath) { await this.storage.put(memberPath, part.buffer, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"); stored.push(memberPath); }
        let inspection: FileInspection;
        try { inspection = await inspectFile(part.fileName, part.buffer); }
        catch { inspection = { store: null, detectedStore: null, source: "none", confidence: 0, layout: null, needsReview: true, preview: { itemCount: 0, sampleItems: [], warnings: ["Não foi possível ler a planilha; substitua ou revise o arquivo."], format: "unknown" } }; }
        rows.push({ id: part.id, day_id: day.id, upload_id: uploadId, kind: "file", file_name: part.fileName, object_path: memberPath, sha256: createHash("sha256").update(part.buffer).digest("hex"), byte_size: part.buffer.length, store: inspection.store, detected_store: inspection.detectedStore, detection_source: inspection.source, detection_confidence: inspection.confidence, layout: inspection.layout, preview: inspection.preview, review_required: inspection.needsReview, state: inspection.needsReview ? "needs_review" : "received" });
      }
      const { error: uploadError } = await this.storage.db.from("day_uploads").insert({ id: uploadId, day_id: day.id, client_upload_id: clientId, file_name: name, object_path: originalPath, sha256: originalHash, byte_size: file.buffer.length });
      if (uploadError?.code === "23505") throw new ConflictException("Envio simultâneo com o mesmo identificador; tente novamente");
      if (uploadError) throw failed();
      metadataSaved = true;
      const { data, error: entryError } = await this.storage.db.from("day_entries").insert(rows).select("*");
      if (entryError || !data) throw failed();
      return data as DayEntry[];
    } catch (error) {
      if (metadataSaved) await this.storage.db.from("day_uploads").delete().eq("id", uploadId);
      await Promise.allSettled(stored.map((objectPath) => this.storage.remove(objectPath)));
      throw error;
    }
  }
}
