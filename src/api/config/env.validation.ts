export interface AppEnvironment {
  NODE_ENV: string;
  PORT: number;
  CORS_ORIGINS: string;
}

export function validateEnvironment(
  raw: Record<string, unknown>,
): AppEnvironment & Record<string, unknown> {
  const port = Number(raw.PORT ?? 3000);

  if (!Number.isInteger(port) || port < 1 || port > 65535) {
    throw new Error("PORT must be an integer between 1 and 65535");
  }

  return {
    ...raw,
    NODE_ENV: String(raw.NODE_ENV ?? "development"),
    PORT: port,
    CORS_ORIGINS: String(raw.CORS_ORIGINS ?? "http://localhost:5173"),
  };
}
