import ExcelJS from "exceljs";

export async function structuredSpreadsheet(
  store: string,
  items: Array<[product: string, quantity: number | string]>,
): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Pedido");
  sheet.addRow(["FILIAL", store]);
  sheet.addRow(["PRODUTO", "QUANTIDADE"]);
  items.forEach((item) => sheet.addRow(item));
  return Buffer.from(await workbook.xlsx.writeBuffer());
}

export async function looseSpreadsheet(
  rows: Array<[description: string, quantity?: number | string]>,
): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Pedido");
  rows.forEach((row) => sheet.addRow(row));
  return Buffer.from(await workbook.xlsx.writeBuffer());
}
