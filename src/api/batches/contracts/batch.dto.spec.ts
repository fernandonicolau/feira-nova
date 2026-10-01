import "reflect-metadata";
import { plainToInstance } from "class-transformer";
import { validate } from "class-validator";
import { BatchEntryType, ProcessBatchDto } from "./batch.dto";

async function validateContract(value: unknown) {
  return validate(plainToInstance(ProcessBatchDto, value), {
    whitelist: true,
    forbidNonWhitelisted: true,
  });
}

describe("ProcessBatchDto", () => {
  it("accepts multiple files, texts, and a mixed batch", async () => {
    const errors = await validateContract({
      name: "Pedido da manhã",
      entries: [
        { id: "file-1", type: BatchEntryType.FILE, store: "Cerâmica", fileRef: "sheet1" },
        { id: "file-2", type: BatchEntryType.FILE, store: "Coelho", fileRef: "sheet2" },
        { id: "text-1", type: BatchEntryType.TEXT, store: "Queimados", text: "Banana 5" },
        { id: "text-2", type: BatchEntryType.TEXT, store: "Anchieta", text: "Maçã 3" },
      ],
    });

    expect(errors).toHaveLength(0);
  });

  it.each([
    [{ entries: [] }],
    [{ entries: [{ id: "file-1", type: "file", store: "Cerâmica" }] }],
    [{ entries: [{ id: "text-1", type: "text", store: "", text: "Banana 5" }] }],
    [{ entries: [{ id: "text-1", type: "text", store: "Queimados", text: "" }] }],
    [{ entries: [{ id: "file-1", type: "file", store: "Cerâmica", fileRef: "sheet1", text: "Banana 5" }] }],
    [{ entries: [
      { id: "same", type: "text", store: "Cerâmica", text: "Banana 5" },
      { id: "same", type: "text", store: "Coelho", text: "Maçã 3" },
    ] }],
  ])("rejects an invalid batch: %j", async (value) => {
    expect(await validateContract(value)).not.toHaveLength(0);
  });
});
