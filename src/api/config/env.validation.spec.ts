import { parseCorsOrigins, validateEnvironment } from "./env.validation";

describe("environment validation", () => {
  it("normalizes and deduplicates explicit CORS origins", () => {
    expect(
      parseCorsOrigins(
        "https://web.example.com/, http://localhost:5173,https://web.example.com",
      ),
    ).toEqual(["https://web.example.com", "http://localhost:5173"]);
  });

  it.each(["*", "", "web.example.com", "https://web.example.com/path"])(
    "rejects an unsafe CORS origin list: %s",
    (value) => {
      expect(() => parseCorsOrigins(value)).toThrow();
    },
  );

  it("returns a canonical environment", () => {
    expect(
      validateEnvironment({
        NODE_ENV: "production",
        PORT: "10000",
        CORS_ORIGINS: "https://web.example.com/",
        REQUEST_TIMEOUT_MS: "90000",
      }),
    ).toMatchObject({
      NODE_ENV: "production",
      PORT: 10000,
      CORS_ORIGINS: "https://web.example.com",
      REQUEST_TIMEOUT_MS: 90000,
    });
  });

  it("rejects invalid request timeouts", () => {
    expect(() =>
      validateEnvironment({
        CORS_ORIGINS: "https://web.example.com",
        REQUEST_TIMEOUT_MS: 0,
      }),
    ).toThrow("REQUEST_TIMEOUT_MS must be a positive integer");
  });
});
