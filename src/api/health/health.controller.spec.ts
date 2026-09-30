import { HealthController } from "./health.controller";

describe("HealthController", () => {
  it("reports that the process is healthy", () => {
    expect(new HealthController().check()).toEqual({ status: "ok" });
  });
});
