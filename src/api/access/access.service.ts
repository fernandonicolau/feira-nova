import { ForbiddenException, HttpException, HttpStatus, Injectable, ServiceUnavailableException } from "@nestjs/common";
import { createHmac, randomBytes, scryptSync, timingSafeEqual } from "node:crypto";

const SESSION_MS = 8 * 60 * 60 * 1000;
const WINDOW_MS = 15 * 60 * 1000;

@Injectable()
export class AccessService {
  private readonly attempts = new Map<string, { count: number; until: number }>();

  private config(): { code: string; secret: string } {
    const code = process.env.ACCESS_CODE ?? "";
    const secret = process.env.SESSION_SECRET ?? "";
    if (!/^\d{6}$/.test(code) || secret.length < 32) {
      throw new ServiceUnavailableException("Acesso ainda não configurado");
    }
    return { code, secret };
  }

  private mac(value: string, secret: string): Buffer {
    return createHmac("sha256", secret).update(value).digest();
  }

  login(input: string, address: string): string {
    const { code, secret } = this.config();
    const key = address || "unknown";
    const now = Date.now();
    if (this.attempts.size > 5000) {
      for (const [addressKey, record] of this.attempts) if (record.until <= now) this.attempts.delete(addressKey);
    }
    const attempt = this.attempts.get(key);
    if (attempt && attempt.until > now && attempt.count >= 5) throw new HttpException("Muitas tentativas. Tente novamente em 15 minutos.", HttpStatus.TOO_MANY_REQUESTS);
    if (!/^\d{6}$/.test(input) || !timingSafeEqual(scryptSync(input, secret, 32), scryptSync(code, secret, 32))) {
      this.attempts.set(key, { count: (attempt?.until && attempt.until > now ? attempt.count : 0) + 1, until: now + WINDOW_MS });
      throw new ForbiddenException("Código incorreto");
    }
    this.attempts.delete(key);
    const expires = now + SESSION_MS;
    const nonce = randomBytes(12).toString("hex");
    const payload = `v1.${expires}.${nonce}`;
    const signature = this.mac(`${payload}.${code}`, secret).toString("hex");
    return `${payload}.${signature}`;
  }

  valid(token: string | undefined): boolean {
    if (!token) return false;
    const { code, secret } = this.config();
    const parts = token.split(".");
    if (parts.length !== 4 || parts[0] !== "v1" || !/^\d+$/.test(parts[1])) return false;
    const expires = Number(parts[1]);
    if (!Number.isSafeInteger(expires) || expires <= Date.now() || expires > Date.now() + SESSION_MS) return false;
    const provided = Buffer.from(parts[3], "hex");
    const expected = this.mac(`${parts.slice(0, 3).join(".")}.${code}`, secret);
    return provided.length === expected.length && timingSafeEqual(provided, expected);
  }

  cookie(token: string, clear = false): string {
    const secure = process.env.COOKIE_SECURE === "true" || process.env.NODE_ENV === "production";
    return `feira_session=${token}; HttpOnly; Path=/api/v1; Max-Age=${clear ? 0 : SESSION_MS / 1000}; SameSite=${secure ? "None" : "Lax"}${secure ? "; Secure" : ""}`;
  }
}
