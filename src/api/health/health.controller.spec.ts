import { HealthController } from "./health.controller";

describe("HealthController", () => {
  it("reports that the process is healthy", () => {
    expect(new HealthController().check()).toEqual({ status: "ok" });
  });

  it("reports readiness without depending on local storage", () => {
    expect(new HealthController().readiness()).toEqual({
      status: "ready",
      uptimeSeconds: expect.any(Number),
    });
  });
});
