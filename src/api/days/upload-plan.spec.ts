import JSZip from "jszip";
import { planUpload } from "./upload-plan";

describe("planUpload", () => {
  it("keeps each ZIP member traceable and rejects unsafe paths", async () => {
    const zip = new JSZip();
    zip.file("loja-a/pedido.xlsx", Buffer.from("planilha-a"));
    zip.file("loja-b/pedido.xlsx", Buffer.from("planilha-b"));
    const members = await planUpload("pedidos.zip", await zip.generateAsync({ type: "nodebuffer" }));
    expect(members).toHaveLength(2);
    expect(members[0].fileName).toBe("pedido.xlsx");
    expect(members[0].id).not.toBe(members[1].id);
    expect(members.map((member) => member.buffer.toString())).toEqual(["planilha-a", "planilha-b"]);

    const unsafe = new JSZip();
    unsafe.file("../segredo.xlsx", "conteúdo");
    await expect(planUpload("unsafe.zip", await unsafe.generateAsync({ type: "nodebuffer" }))).rejects.toThrow("Caminho inseguro");
  });

  it("rejects unsupported uploads without interpreting them", async () => {
    await expect(planUpload("pedido.csv", Buffer.from("abc"))).rejects.toThrow("Formato não aceito");
  });
});
