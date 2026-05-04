const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");

const ROOT_DIR = process.cwd();
const DEFAULT_MAP_DIR = path.join(ROOT_DIR, "output");
const DEFAULT_TEMPLATE_DIR = path.join(ROOT_DIR, "exemplo");
const PRODUCT_START_ROW = 9;

const MAP_FILES = [
  {
    fileName: "MAPA.xlsx",
    sections: [
      { productColumn: 1, storeColumns: { CERAMICA: 2, COELHO: 3, QUEIMADOS: 4 } },
      { productColumn: 5, storeColumns: { CERAMICA: 6, COELHO: 7, QUEIMADOS: 8 } },
    ],
  },
  {
    fileName: "MAPA2.xlsx",
    sections: [
      { productColumn: 1, storeColumns: { PIABETA: 2, ANCHIETA: 3, OLINDA: 4, "SANTA CRUZ": 5 } },
      { productColumn: 6, storeColumns: { PIABETA: 7, ANCHIETA: 8, OLINDA: 9, "SANTA CRUZ": 10 } },
    ],
  },
  {
    fileName: "MAPA3.xlsx",
    sections: [
      { productColumn: 1, storeColumns: { IRAJA: 2, CACHAMBI: 3, SANTOS: 4, FREGUESIA: 5 } },
      { productColumn: 6, storeColumns: { IRAJA: 7, CACHAMBI: 8, SANTOS: 9, FREGUESIA: 10 } },
    ],
  },
];

const STORE_ALIASES = new Map([
  ["ANCH", "ANCHIETA"],
  ["ANCHIETA", "ANCHIETA"],
  ["CACH", "CACHAMBI"],
  ["CACHAMBI", "CACHAMBI"],
  ["CERAM", "CERAMICA"],
  ["CERAMICA", "CERAMICA"],
  ["COELHO", "COELHO"],
  ["COELHO DA ROCHA", "COELHO"],
  ["C ROCHA", "COELHO"],
  ["CROCHA", "COELHO"],
  ["FREG", "FREGUESIA"],
  ["FREGUE", "FREGUESIA"],
  ["FREGUESIA", "FREGUESIA"],
  ["IRAJA", "IRAJA"],
  ["OLINDA", "OLINDA"],
  ["PIABETA", "PIABETA"],
  ["QUEIM", "QUEIMADOS"],
  ["QUEIMADOS", "QUEIMADOS"],
  ["S CRUZ", "SANTA CRUZ"],
  ["STA CRUZ", "SANTA CRUZ"],
  ["STACRUZ", "SANTA CRUZ"],
  ["SANTA CRUZ", "SANTA CRUZ"],
  ["SANTOS", "SANTOS"],
]);

const PRODUCT_REPLACEMENTS = [
  [/\bABACAXI UNID\b/g, "ABACAXI"],
  [/\bABOBORA BAHIANA\b/g, "ABOBORA BAIANA"],
  [/\bBATATA BAROA BDJ\b/g, "BATATA BAROA"],
  [/\bBANANA D AGUA\b/g, "BANANA DAGUA"],
  [/\bBANANA DAGUA\b/g, "BANANA DAGUA"],
  [/\bCOCO SECO UN\b/g, "COCO SECO"],
  [/\bGOIABA GRANEL\b/g, "GOIABA"],
  [/\bKIWI KG\b/g, "KIWI"],
  [/\bLARANJA SELETA\b/g, "LARANJA SELETA"],
  [/\bLIMAO THAITI\b/g, "LIMAO"],
  [/\bMACA RED IMPORT\b/g, "MACA RED"],
  [/\bMACA VERDE GRAN\b/g, "MACA VERDE"],
  [/\bMACA GALA 850G\b/g, "MACA 850G"],
  [/\bMAMAO PAPAYA\b/g, "MAMAO HAVAI"],
  [/\bMELAO CANT\b/g, "MELAO CANTALOUPE"],
  [/\bMILHO VERDE BDJ 3\b/g, "MILHO BDJ"],
  [/\bMILHO VERDE\b/g, "MILHO"],
  [/\bOVOS BRANCO C 20\b/g, "OVOS BRANCOS C 20"],
  [/\bOVOS BRANCOS 30\b/g, "OVOS BRANCOS C 30"],
  [/\bOVOS CODORNA C 30\b/g, "OVOS CODORNA"],
  [/\bOVOS VERMELHOS C 12\b/g, "OVOS VERMELHO C 12"],
  [/\bPERA WILLIANS\b/g, "PERA WILLIAMS"],
  [/\bPEPINO COMUM\b/g, "PEPINO"],
  [/\bPIMENTAO VERDE\b/g, "PIMENTAO"],
  [/\bQUIABO 300G\b/g, "QUIABO BDJ"],
  [/\bQUIABO BANDEJA 300G\b/g, "QUIABO BDJ"],
  [/\bQUIABO EMBALADOS\b/g, "QUIABO BDJ"],
  [/\bTANGERINA IMP\b/g, "TANGERINA IMPORTADA"],
  [/\bTOMATE SWEET 180\b/g, "TOMATE SWEET"],
  [/\bUVA ITALIA\b/g, "UVA ITALIA"],
];

const ALWAYS_SUPPLIER_PRODUCTS = [
  {
    fornecedor: "adonai",
    produtos: new Set(["CEBOLA ROXA"]),
  },
  {
    fornecedor: "FAISÃO",
    produtos: new Set(["TANGERINA IMPORTADA"]),
  },
  {
    fornecedor: "Kifrut",
    produtos: new Set(["MACA 850G"]),
  },
];

function worksheetValueToString(value) {
  if (value == null) {
    return "";
  }
  if (typeof value === "object") {
    if (value.richText) {
      return value.richText.map((part) => part.text).join("");
    }
    if (value.text) {
      return String(value.text);
    }
    if (value.result != null) {
      return String(value.result);
    }
    if (value.formula) {
      return "";
    }
  }
  return String(value);
}

function normalizeText(value) {
  return worksheetValueToString(value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/[.'’"()]/g, " ")
    .replace(/[-/,:;]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeStore(value) {
  const normalized = normalizeText(value);
  return STORE_ALIASES.get(normalized) ?? STORE_ALIASES.get(normalized.replace(/\s+/g, "")) ?? null;
}

function normalizeProduct(value) {
  let normalized = normalizeText(value);

  for (const [pattern, replacement] of PRODUCT_REPLACEMENTS) {
    normalized = normalized.replace(pattern, replacement);
  }

  return normalized.replace(/\s+/g, " ").trim();
}

function quantityForCell(value) {
  if (value == null || value === "") {
    return null;
  }
  if (typeof value !== "number") {
    return value;
  }
  if (Number.isInteger(value)) {
    return value;
  }
  return Number(value.toFixed(2).replace(/\.?0+$/, ""));
}

function lookupKey(product, store) {
  return `${normalizeProduct(product)}|${store}`;
}

function isFilledQuantity(value) {
  return value != null && value !== "";
}

async function loadMapQuantities(mapDir) {
  const quantities = new Map();
  const entries = [];

  for (const mapFile of MAP_FILES) {
    const mapPath = path.join(mapDir, mapFile.fileName);
    if (!fs.existsSync(mapPath)) {
      throw new Error(`Mapa nao encontrado: ${mapPath}`);
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(mapPath);
    const worksheet = workbook.worksheets[0];

    for (const section of mapFile.sections) {
      for (let rowNumber = PRODUCT_START_ROW; rowNumber <= worksheet.rowCount; rowNumber += 1) {
        const product = worksheetValueToString(worksheet.getRow(rowNumber).getCell(section.productColumn).value).trim();
        if (!product) {
          continue;
        }

        for (const [store, columnNumber] of Object.entries(section.storeColumns)) {
          const value = worksheet.getRow(rowNumber).getCell(columnNumber).value;
          if (isFilledQuantity(value)) {
            const key = lookupKey(product, store);
            const quantity = quantityForCell(value);
            quantities.set(key, quantity);
            entries.push({
              key,
              produtoMapa: product,
              produtoNormalizado: normalizeProduct(product),
              loja: store,
              quantidade: quantity,
              mapa: mapFile.fileName,
              celula: `${columnNumberToName(columnNumber)}${rowNumber}`,
            });
          }
        }
      }
    }
  }

  return { quantities, entries };
}

function findHeaderRow(worksheet) {
  let best = null;

  for (let rowNumber = 1; rowNumber <= Math.min(worksheet.rowCount, 10); rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    const storeColumns = [];
    let totalColumn = null;

    for (let columnNumber = 1; columnNumber <= worksheet.columnCount; columnNumber += 1) {
      const text = normalizeText(row.getCell(columnNumber).value);
      const store = normalizeStore(text);

      if (store) {
        storeColumns.push({ store, columnNumber });
      } else if (text === "TOTAL") {
        totalColumn = columnNumber;
      }
    }

    if (!best || storeColumns.length > best.storeColumns.length) {
      best = { rowNumber, storeColumns, totalColumn };
    }
  }

  if (!best || !best.storeColumns.length) {
    throw new Error(`Nao encontrei linha de lojas na aba ${worksheet.name}.`);
  }

  return best;
}

function isTitleCell(value) {
  const text = normalizeText(value);
  return !text || text === "PRODUTO" || text === "PRODUTOS" || text === "TOTAL" || /\d{2}\s+\d{2}\s+\d{4}/.test(text);
}

function columnNumberToName(columnNumber) {
  let dividend = columnNumber;
  let columnName = "";

  while (dividend > 0) {
    const modulo = (dividend - 1) % 26;
    columnName = String.fromCharCode(65 + modulo) + columnName;
    dividend = Math.floor((dividend - modulo) / 26);
  }

  return columnName;
}

function supplierNameFromFile(fileName) {
  return path.basename(fileName, path.extname(fileName));
}

function isAlwaysSupplierProduct(fileName, product) {
  const fornecedor = normalizeText(supplierNameFromFile(fileName));
  const produto = normalizeProduct(product);

  return ALWAYS_SUPPLIER_PRODUCTS.some((rule) => normalizeText(rule.fornecedor) === fornecedor && rule.produtos.has(produto));
}

function rowHasStoreQuantity(row, storeColumns) {
  return storeColumns.some(({ columnNumber }) => isFilledQuantity(row.getCell(columnNumber).value));
}

function updateTotalFormulas(worksheet, header) {
  if (!header.totalColumn) {
    return;
  }

  const firstStoreColumn = header.storeColumns[0]?.columnNumber;
  const lastStoreColumn = header.storeColumns.at(-1)?.columnNumber;
  if (!firstStoreColumn || !lastStoreColumn) {
    return;
  }

  const firstStoreName = columnNumberToName(firstStoreColumn);
  const lastStoreName = columnNumberToName(lastStoreColumn);

  for (let rowNumber = header.rowNumber + 1; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    const product = worksheetValueToString(row.getCell(1).value).trim();

    if (isTitleCell(product)) {
      continue;
    }

    row.getCell(header.totalColumn).value = {
      formula: `SUM(${firstStoreName}${rowNumber}:${lastStoreName}${rowNumber})`,
    };
  }
}

function removeProductRowsWithoutOrders(worksheet, header) {
  for (let rowNumber = worksheet.rowCount; rowNumber > header.rowNumber; rowNumber -= 1) {
    const row = worksheet.getRow(rowNumber);
    const product = worksheetValueToString(row.getCell(1).value).trim();

    if (!product && !rowHasStoreQuantity(row, header.storeColumns)) {
      worksheet.spliceRows(rowNumber, 1);
      continue;
    }

    if (isTitleCell(product)) {
      continue;
    }

    if (!rowHasStoreQuantity(row, header.storeColumns)) {
      worksheet.spliceRows(rowNumber, 1);
    }
  }

  updateTotalFormulas(worksheet, header);
}

function clearAndFillSupplierWorksheet(worksheet, quantities, fileName, consumedKeys) {
  const header = findHeaderRow(worksheet);
  const mappedCells = [];

  for (let rowNumber = header.rowNumber + 1; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    const product = worksheetValueToString(row.getCell(1).value).trim();

    if (isTitleCell(product)) {
      continue;
    }

    for (const { store, columnNumber } of header.storeColumns) {
      const cell = row.getCell(columnNumber);
      if (isFilledQuantity(cell.value) || isAlwaysSupplierProduct(fileName, product)) {
        mappedCells.push({ rowNumber, columnNumber, product, store });
      }
      cell.value = null;
    }
  }

  for (const { rowNumber, columnNumber, product, store } of mappedCells) {
    const key = lookupKey(product, store);
    const quantity = quantities.get(key);

    if (quantity == null) {
      continue;
    }

    worksheet.getRow(rowNumber).getCell(columnNumber).value = quantity;
    consumedKeys.add(key);
  }

  removeProductRowsWithoutOrders(worksheet, header);
}

async function writeUnmatchedWorkbook(outputDir, pendingEntries) {
  if (!pendingEntries.length) {
    return null;
  }

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Associacoes pendentes");

  worksheet.columns = [
    { header: "Mapa", key: "mapa", width: 14 },
    { header: "Celula", key: "celula", width: 10 },
    { header: "Produto no mapa", key: "produtoMapa", width: 32 },
    { header: "Produto normalizado", key: "produtoNormalizado", width: 32 },
    { header: "Loja", key: "loja", width: 16 },
    { header: "Quantidade", key: "quantidade", width: 12 },
    { header: "Chave buscada", key: "chaveBusca", width: 42 },
    { header: "Fornecedor correto", key: "fornecedorCorreto", width: 24 },
    { header: "Como tratar", key: "comoTratar", width: 42 },
  ];

  worksheet.getRow(1).font = { bold: true };
  worksheet.views = [{ state: "frozen", ySplit: 1 }];
  worksheet.autoFilter = {
    from: "A1",
    to: "I1",
  };

  pendingEntries.forEach((item) => {
    worksheet.addRow({
      ...item,
      chaveBusca: item.key,
      fornecedorCorreto: "",
      comoTratar: "",
    });
  });

  const fileName = "associacoes-pendentes.xlsx";
  const filePath = path.join(outputDir, fileName);
  await workbook.xlsx.writeFile(filePath);

  return {
    fileName,
    filePath,
  };
}

async function generateSupplierFiles(options = {}) {
  const mapDir = options.mapDir ?? DEFAULT_MAP_DIR;
  const templateDir = options.templateDir ?? DEFAULT_TEMPLATE_DIR;
  const outputDir = options.outputDir ?? path.join(mapDir, "fornecedores");

  if (!fs.existsSync(templateDir)) {
    throw new Error("Pasta exemplo nao encontrada.");
  }

  const { quantities, entries } = await loadMapQuantities(mapDir);
  fs.rmSync(outputDir, { recursive: true, force: true });
  fs.mkdirSync(outputDir, { recursive: true });
  const consumedKeys = new Set();

  const supplierFiles = fs
    .readdirSync(templateDir)
    .filter((fileName) => /\.xlsx$/i.test(fileName))
    .sort((a, b) => a.localeCompare(b, "pt-BR"));

  for (const fileName of supplierFiles) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(path.join(templateDir, fileName));

    if (workbook.worksheets[0]) {
      clearAndFillSupplierWorksheet(workbook.worksheets[0], quantities, fileName, consumedKeys);
    }

    await workbook.xlsx.writeFile(path.join(outputDir, fileName));
  }

  const pendingEntries = entries.filter((entry) => !consumedKeys.has(entry.key));
  const unmatchedFile = await writeUnmatchedWorkbook(outputDir, pendingEntries);

  return {
    outputDir,
    files: supplierFiles,
    unmatched: pendingEntries,
    unmatchedFile,
  };
}

module.exports = {
  generateSupplierFiles,
};

if (require.main === module) {
  generateSupplierFiles()
    .then(({ outputDir, files, unmatchedFile }) => {
      console.log(`Arquivos de fornecedores gerados em: ${outputDir}`);
      for (const fileName of files) {
        console.log(`- ${fileName}`);
      }
      if (unmatchedFile) {
        console.log(`Associacoes pendentes: ${unmatchedFile.fileName}`);
      }
    })
    .catch((error) => {
      console.error(error.message);
      process.exitCode = 1;
    });
}
