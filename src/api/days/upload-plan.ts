import { BadRequestException } from "@nestjs/common";
import { randomUUID } from "node:crypto";
import path from "node:path";
import { Readable } from "node:stream";
import JSZip from "jszip";
import { BATCH_LIMITS } from "../batches/contracts";

export interface UploadPart {
  id: string;
  fileName: string;
  buffer: Buffer;
}

async function readMember(member: JSZip.JSZipObject, maxBytes: number): Promise<Buffer> {
  const chunks: Buffer[] = [];
  let size = 0;
  const stream = new Readable({ read() {} }).wrap(member.nodeStream());
  for await (const chunk of stream) {
    const buffer = Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk);
    size += buffer.length;
    if (size > maxBytes) {
      stream.destroy();
      throw new BadRequestException(`Arquivo muito grande no ZIP: ${member.name}`);
    }
    chunks.push(buffer);
  }
  return Buffer.concat(chunks);
}

export function safeFileName(raw: string): string {
  const name = raw.replace(/\\/g, "/").split("/").pop()?.trim() ?? "";
  if (!name || name.length > 180 || /[\x00-\x1f]/.test(name)) throw new BadRequestException("Nome de arquivo inválido");
  return name;
}

export async function planUpload(fileName: string, buffer: Buffer): Promise<UploadPart[]> {
  const name = safeFileName(fileName);
  if (!buffer.length || buffer.length > BATCH_LIMITS.maxFileBytes) throw new BadRequestException(`Tamanho inválido: ${name}`);
  if (/\.(xlsx|xlsm)$/i.test(name)) return [{ id: randomUUID(), fileName: name, buffer }];
  if (!/\.zip$/i.test(name)) throw new BadRequestException(`Formato não aceito: ${name}`);

  let zip: JSZip;
  try { zip = await JSZip.loadAsync(buffer); } catch { throw new BadRequestException(`ZIP inválido: ${name}`); }
  const members = Object.values(zip.files).filter((member) => !member.dir);
  if (!members.length || members.length > BATCH_LIMITS.maxFiles) throw new BadRequestException(`Quantidade de arquivos inválida no ZIP: ${name}`);
  const parts: UploadPart[] = [];
  let total = 0;
  for (const member of members) {
    const original = member.unsafeOriginalName ?? member.name;
    if (original.startsWith("/") || original.startsWith("\\") || original.split(/[\\/]/).includes("..")) {
      throw new BadRequestException(`Caminho inseguro no ZIP: ${original}`);
    }
    if (!/\.(xlsx|xlsm)$/i.test(member.name)) throw new BadRequestException(`Arquivo não aceito no ZIP: ${member.name}`);
    const content = await readMember(member, Math.min(BATCH_LIMITS.maxFileBytes, BATCH_LIMITS.maxTotalFileBytes - total));
    total += content.length;
    if (!content.length || content.length > BATCH_LIMITS.maxFileBytes || total > BATCH_LIMITS.maxTotalFileBytes) {
      throw new BadRequestException(`Tamanho inválido no ZIP: ${member.name}`);
    }
    parts.push({ id: randomUUID(), fileName: path.posix.basename(member.name), buffer: content });
  }
  return parts;
}
