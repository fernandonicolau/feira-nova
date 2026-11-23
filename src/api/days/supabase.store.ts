import { Injectable, ServiceUnavailableException } from "@nestjs/common";
import { createClient, type SupabaseClient } from "@supabase/supabase-js";

export const DAILY_BUCKET = "feira-nova-private";

@Injectable()
export class SupabaseStore {
  readonly client: SupabaseClient | null;

  constructor() {
    const url = process.env.SUPABASE_URL;
    const key = process.env.SUPABASE_SERVICE_ROLE_KEY;
    this.client = url && key ? createClient(url, key, {
      auth: { autoRefreshToken: false, persistSession: false },
    }) : null;
  }

  get db(): SupabaseClient {
    if (!this.client) throw new ServiceUnavailableException("Supabase não configurado na API");
    return this.client;
  }

  async put(path: string, body: Buffer, contentType: string): Promise<void> {
    const { error } = await this.db.storage.from(DAILY_BUCKET).upload(path, body, {
      contentType, upsert: false,
    });
    if (error) throw new ServiceUnavailableException(`Falha ao guardar arquivo: ${error.message}`);
  }

  async get(path: string): Promise<Buffer> {
    const { data, error } = await this.db.storage.from(DAILY_BUCKET).download(path);
    if (error || !data) throw new ServiceUnavailableException("Arquivo indisponível no armazenamento");
    return Buffer.from(await data.arrayBuffer());
  }

  async remove(path: string): Promise<void> {
    await this.db.storage.from(DAILY_BUCKET).remove([path]);
  }
}
