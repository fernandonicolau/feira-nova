export interface AppEnvironment {
  NODE_ENV: string;
  PORT: number;
  CORS_ORIGINS: string;
}

export function parseCorsOrigins(value: string): string[] {
  const origins = value
    .split(",")
    .map((origin) => origin.trim().replace(/\/$/, ""))
    .filter(Boolean);

  if (origins.length === 0) {
    throw new Error("CORS_ORIGINS must contain at least one origin");
  }

  for (const origin of origins) {
    if (origin === "*") {
      throw new Error("CORS_ORIGINS does not allow wildcard origins");
    }

    let parsed: URL;
    try {
      parsed = new URL(origin);
    } catch {
      throw new Error(`Invalid CORS origin: ${origin}`);
    }

    if (!["http:", "https:"].includes(parsed.protocol) || parsed.origin !== origin) {
      throw new Error(`Invalid CORS origin: ${origin}`);
    }
  }

  return [...new Set(origins)];
}

export function validateEnvironment(
  raw: Record<string, unknown>,
): AppEnvironment & Record<string, unknown> {
  const port = Number(raw.PORT ?? 3000);

  if (!Number.isInteger(port) || port < 1 || port > 65535) {
    throw new Error("PORT must be an integer between 1 and 65535");
  }

  const corsOrigins = parseCorsOrigins(
    String(raw.CORS_ORIGINS ?? "http://localhost:5173"),
  );

  return {
    ...raw,
    NODE_ENV: String(raw.NODE_ENV ?? "development"),
    PORT: port,
    CORS_ORIGINS: corsOrigins.join(","),
  };
}
