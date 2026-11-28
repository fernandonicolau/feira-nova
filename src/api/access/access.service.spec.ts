import { AccessService } from "./access.service";

describe("AccessService", () => {
  const original = { code: process.env.ACCESS_CODE, secret: process.env.SESSION_SECRET };
  beforeEach(() => {
    process.env.ACCESS_CODE = "123456";
    process.env.SESSION_SECRET = "a-long-secret-at-least-thirty-two-characters";
  });
  afterAll(() => {
    if (original.code === undefined) delete process.env.ACCESS_CODE; else process.env.ACCESS_CODE = original.code;
    if (original.secret === undefined) delete process.env.SESSION_SECRET; else process.env.SESSION_SECRET = original.secret;
  });

  it("accepts a correct code and rejects a changed code", () => {
    const access = new AccessService();
    const token = access.login("123456", "127.0.0.1");
    expect(access.valid(token)).toBe(true);
    process.env.ACCESS_CODE = "654321";
    expect(access.valid(token)).toBe(false);
  });

  it("limits repeated wrong codes", () => {
    const access = new AccessService();
    for (let i = 0; i < 5; i++) expect(() => access.login("000000", "127.0.0.2")).toThrow("Código incorreto");
    expect(() => access.login("123456", "127.0.0.2")).toThrow("Muitas tentativas");
  });
});
