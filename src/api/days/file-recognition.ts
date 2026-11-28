import { BadRequestException } from "@nestjs/common";
import ExcelJS from "exceljs";

export interface ColumnPair { nameColumn: string; quantityColumn: string }
export interface ColumnLayout { headerRow: number; pairs: ColumnPair[] }
export interface FileInspection {
  store: string | null;
  detectedStore: string | null;
  source: "header" | "filename" | "conflict" | "none" | "manual";
  confidence: number;
  layout: ColumnLayout | null;
  needsReview: boolean;
  preview: { itemCount: number; sampleItems: Array<{ name: string; quantity: number }>; warnings: string[]; format: string };
}

export const KNOWN_STORES = ["Cerâmica", "Coelho", "Queimados", "Piabetá", "Anchieta", "Olinda", "Santa Cruz", "Irajá", "Cachambi", "Santos", "Freguesia"] as const;

const ALIASES: Record<string, string> = {
  CERAMICA: "Cerâmica", CERAM: "Cerâmica", "NOVA IGUACU": "Cerâmica",
  COELHO: "Coelho", QUEIMADOS: "Queimados", QUEIM: "Queimados",
  PIABETA: "Piabetá", ANCHIETA: "Anchieta", ANCH: "Anchieta", ANCHEITA: "Anchieta",
  OLINDA: "Olinda", "STA CRUZ": "Santa Cruz", STACRUZ: "Santa Cruz", "SANTA CRUZ": "Santa Cruz",
  IRAJA: "Irajá", CACHAMBI: "Cachambi", CACH: "Cachambi", SANTOS: "Santos",
  "TODOS OS SANTOS": "Santos", FREGUESIA: "Freguesia", FREG: "Freguesia",
};
const NAME_LABELS = new Set(["PRODUTO", "NOME", "ITEM", "DESCRICAO", "NOME DO PRODUTO"]);
const QUANTITY_LABELS = new Set(["QTD", "QTDE", "QUANTIDADE", "QUANT"]);

function normalize(value: unknown): string {
  return String(value ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toUpperCase().replace(/[_./,:;()\-]/g, " ").replace(/\s+/g, " ").trim();
}

function cellText(value: ExcelJS.CellValue): string {
  if (value == null) return "";
  if (typeof value === "object") {
    if ("text" in value && value.text != null) return String(value.text);
    if ("richText" in value && value.richText) return value.richText.map((part) => part.text).join("");
    if ("result" in value && value.result != null) return String(value.result);
    if ("formula" in value) return "";
  }
  return String(value);
}

function column(number: number): string {
  let result = "";
  for (let value = number; value > 0; value = Math.floor((value - 1) / 26)) result = String.fromCharCode(65 + (value - 1) % 26) + result;
  return result;
}

function matchesStore(raw: string): string[] {
  const normalized = ` ${normalize(raw)} `;
  const found = new Set<string>();
  for (const [alias, store] of Object.entries(ALIASES)) if (normalized.includes(` ${alias} `)) found.add(store);
  return [...found];
}

function headerStore(sheet: ExcelJS.Worksheet): string {
  for (let row = 1; row <= Math.min(sheet.rowCount, 20); row++) {
    for (let col = 1; col <= Math.min(sheet.getRow(row).cellCount || 30, 30); col++) {
      const raw = cellText(sheet.getRow(row).getCell(col).value);
      if (normalize(raw) === "FILIAL") {
        for (let offset = 1; offset <= 3; offset++) {
          const value = cellText(sheet.getRow(row).getCell(col + offset).value).trim();
          if (value) return value;
        }
      }
      if (/^FILIAL\s*[:\-]/i.test(raw)) return raw.replace(/^FILIAL\s*[:\-]\s*/i, "").trim();
    }
  }
  return "";
}

function detectLayout(sheet: ExcelJS.Worksheet): ColumnLayout | null {
  let best: ColumnLayout | null = null;
  for (let row = 1; row <= Math.min(sheet.rowCount, 20); row++) {
    const names: number[] = [];
    const quantities: number[] = [];
    for (let col = 1; col <= Math.min(Math.max(sheet.getRow(row).cellCount, 2), 30); col++) {
      const value = normalize(cellText(sheet.getRow(row).getCell(col).value));
      if (NAME_LABELS.has(value)) names.push(col);
      if (QUANTITY_LABELS.has(value)) quantities.push(col);
    }
    const used = new Set<number>();
    const pairs = names.map((name) => {
      const quantity = quantities.find((q) => q > name && q - name <= 3 && !used.has(q));
      if (!quantity) return null;
      used.add(quantity);
      return { nameColumn: column(name), quantityColumn: column(quantity) };
    }).filter((pair): pair is ColumnPair => pair !== null);
    if (pairs.length > (best?.pairs.length ?? 0)) best = { headerRow: row, pairs };
  }
  if (best) return best;
  let simpleRows = 0;
  for (let row = 1; row <= Math.min(sheet.rowCount, 40); row++) {
    if (cellText(sheet.getCell(`A${row}`).value).trim() && quantity(sheet.getCell(`B${row}`).value) !== null) simpleRows++;
  }
  return simpleRows ? { headerRow: 0, pairs: [{ nameColumn: "A", quantityColumn: "B" }] } : null;
}

function quantity(value: ExcelJS.CellValue): number | null {
  const text = cellText(value).trim().replace(/\./g, "").replace(",", ".");
  if (!text) return null;
  const number = Number(text);
  return Number.isFinite(number) ? number : null;
}

export function validateLayout(value: unknown): ColumnLayout {
  if (!value || typeof value !== "object" || !("headerRow" in value) || !("pairs" in value)) throw new BadRequestException("Mapeamento de colunas inválido");
  const layout = value as ColumnLayout;
  if (!Number.isInteger(layout.headerRow) || layout.headerRow < 0 || layout.headerRow > 50 || !Array.isArray(layout.pairs) || layout.pairs.length < 1 || layout.pairs.length > 4) throw new BadRequestException("Mapeamento de colunas inválido");
  const used = new Set<string>();
  for (const pair of layout.pairs) {
    if (!pair || !/^[A-Z]{1,2}$/.test(pair.nameColumn) || !/^[A-Z]{1,2}$/.test(pair.quantityColumn) || pair.nameColumn === pair.quantityColumn || used.has(pair.nameColumn) || used.has(pair.quantityColumn)) throw new BadRequestException("Colunas duplicadas ou inválidas");
    used.add(pair.nameColumn); used.add(pair.quantityColumn);
  }
  return layout;
}

export function resolveStore(raw: string): string | null {
  const options = matchesStore(raw);
  return options.length === 1 ? options[0] : null;
}

export async function inspectFile(fileName: string, buffer: Buffer, override?: { store?: string; layout?: ColumnLayout }): Promise<FileInspection> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer as never);
  const sheet = workbook.worksheets[0];
  if (!sheet) throw new BadRequestException("Planilha sem abas");
  const header = headerStore(sheet);
  const fromHeader = matchesStore(header);
  const fromName = matchesStore(fileName.replace(/\.(xlsx|xlsm)$/i, ""));
  const warnings: string[] = [];
  let source: FileInspection["source"] = "none";
  let store: string | null = null;
  let confidence = 0;
  if (fromHeader.length > 1 || (fromHeader.length === 1 && fromName.length === 1 && fromHeader[0] !== fromName[0])) {
    source = "conflict"; warnings.push("Cabeçalho e nome do arquivo indicam lojas diferentes.");
  } else if (fromHeader.length === 1) { store = fromHeader[0]; source = "header"; confidence = 0.98; }
  else if (fromName.length === 1 && !header) { store = fromName[0]; source = "filename"; confidence = 0.86; }
  else if (fromName.length === 1) { store = fromName[0]; source = "filename"; confidence = 0.65; warnings.push("O cabeçalho informa uma filial não reconhecida; confirme a loja."); }
  else warnings.push("Loja não identificada; selecione a loja antes de processar.");
  if (override?.store) {
    const corrected = resolveStore(override.store);
    if (!corrected) throw new BadRequestException("Loja não cadastrada");
    store = corrected;
    source = "manual";
    for (let index = warnings.length - 1; index >= 0; index--) {
      if (/loja|filial|cabeçalho/i.test(warnings[index])) warnings.splice(index, 1);
    }
  }
  const layout = override?.layout ? validateLayout(override.layout) : detectLayout(sheet);
  if (!layout) warnings.push("Colunas de produto e quantidade não identificadas.");
  const items: Array<{ name: string; quantity: number }> = [];
  let skipped = 0;
  if (layout) {
    for (let row = layout.headerRow + 1; row <= sheet.rowCount; row++) {
      for (const pair of layout.pairs) {
        const name = cellText(sheet.getCell(`${pair.nameColumn}${row}`).value).trim();
        const amount = quantity(sheet.getCell(`${pair.quantityColumn}${row}`).value);
        if (!name && amount === null) continue;
        if (!name || amount === null) { skipped++; continue; }
        items.push({ name, quantity: amount });
      }
    }
  }
  if (skipped) warnings.push(`${skipped} linha(s) com produto ou quantidade incompletos não foram interpretadas.`);
  if (!items.length) warnings.push("Nenhum item foi reconhecido na planilha.");
  const needsReview = !store || confidence < 0.8 && !override?.store || !layout || !items.length || source === "conflict";
  return { store, detectedStore: fromHeader[0] ?? fromName[0] ?? null, source, confidence: override?.store ? 1 : confidence, layout, needsReview, preview: { itemCount: items.length, sampleItems: items.slice(0, 8), warnings, format: layout ? "columns" : "unknown" } };
}
