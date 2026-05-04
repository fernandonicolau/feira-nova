(function () {
  const PRODUCT_LIST_FILES = {
    legumes: "data/produtos-legumes.json",
    frutas: "data/produtos-frutas.json",
  };
  const PRODUCT_START_ROW = 9;
  let productListCachePromise = null;
  const MAP_CONFIGS = [
    {
      outputName: "MAPA.xlsx",
      templateName: "MAPA.xlsx",
      sections: [
        {
          productColumn: "A",
          storeColumns: { CERAM: "B", COELHO: "C", QUEIM: "D" },
          productListKey: "legumes",
        },
        {
          productColumn: "E",
          storeColumns: { CERAM: "F", COELHO: "G", QUEIM: "H" },
          productListKey: "frutas",
        },
      ],
      storeAliases: {
        CERAMICA: "CERAM",
        CERAM: "CERAM",
        COELHO: "COELHO",
        QUEIMADOS: "QUEIM",
        QUEIM: "QUEIM",
      },
    },
    {
      outputName: "MAPA2.xlsx",
      templateName: "MAPA2.xlsx",
      sections: [
        {
          productColumn: "A",
          storeColumns: { PIABETA: "B", ANCH: "C", OLINDA: "D", "STA CRUZ": "E" },
          productListKey: "legumes",
        },
        {
          productColumn: "F",
          storeColumns: { PIABETA: "G", ANCH: "H", OLINDA: "I", "STA CRUZ": "J" },
          productListKey: "frutas",
        },
      ],
      storeAliases: {
        PIABETA: "PIABETA",
        ANCHIETA: "ANCH",
        ANCH: "ANCH",
        OLINDA: "OLINDA",
        "STA CRUZ": "STA CRUZ",
        STACRUZ: "STA CRUZ",
        "SANTA CRUZ": "STA CRUZ",
      },
    },
    {
      outputName: "MAPA3.xlsx",
      templateName: "MAPA3.xlsx",
      sections: [
        {
          productColumn: "A",
          storeColumns: { IRAJA: "B", CACH: "C", SANTOS: "D", FREG: "E" },
          productListKey: "legumes",
        },
        {
          productColumn: "F",
          storeColumns: { IRAJA: "G", CACH: "H", SANTOS: "I", FREG: "J" },
          productListKey: "frutas",
        },
      ],
      storeAliases: {
        IRAJA: "IRAJA",
        CACHAMBI: "CACH",
        CACH: "CACH",
        SANTOS: "SANTOS",
        FREGUESIA: "FREG",
        FREG: "FREG",
      },
    },
  ];

  const STOP_WORDS = new Set([
    "KG",
    "KILO",
    "KILOS",
    "UN",
    "UNI",
    "UNID",
    "UNIDS",
    "UNIDADE",
    "UNIDADES",
    "UND",
    "UNDS",
    "PCT",
    "PACOTE",
    "BANDEJA",
    "BDJ",
    "CX",
    "C",
    "G",
    "GR",
    "GRAMA",
    "GRAMAS",
    "EMBALADO",
    "IMPORTADA",
    "IMPORTADO",
    "NACIONAL",
    "ARGENTINA",
    "SEM",
    "CAROCO",
    "GRANDE",
    "ORGANICO",
    "CAIPIRA",
    "SITIO",
    "RAIAR",
    "SEAL",
    "COCORICO",
    "FOZ",
  ]);

  const PHRASE_REPLACEMENTS = [
    [/ABOBORA JAP\b/g, "ABOBORA JAPONESA"],
    [/ABOBORA MOR\b/g, "ABOBORA MORANGA"],
    [/ABOBORA SERG\b/g, "ABOBORA SERGIPANA"],
    [/ABOBORA SERGIPANA\b/g, "ABOBORA SERGIPANA"],
    [/ABOBRINHA ITALIANA\b/g, "ABOBRINHA"],
    [/AMORA(?:\s+BANDEJA)?\s+100G\b/g, "AMORA"],
    [/BANANA D AGUA\b/g, "BANANA DAGUA"],
    [/BANANA DAGUA\b/g, "BANANA DAGUA"],
    [/BATATA BAROA BANDEJA\b/g, "BATATA BAROA"],
    [/BATATA BAROA BANDEJA \d+G\b/g, "BATATA BAROA"],
    [/BATATA CALABRESA\b/g, "BATATA BOLINHA"],
    [/BATATA BOLINHA PACOTE\b/g, "BATATA BOLINHA"],
    [/BATATA BOLINHA PACOTE \d+KG\b/g, "BATATA BOLINHA"],
    [/BATATA BOLINHA PACOTE \d+G\b/g, "BATATA BOLINHA"],
    [/BATATA BOLINHA PCT\b/g, "BATATA BOLINHA"],
    [/BATATA BAROA BDJ\b/g, "BATATA BAROA"],
    [/BATATA ESCOVADA\b/g, "BATATA SUJA"],
    [/ALHO DESCASCADO\b/g, "ALHO DESCASCADO"],
    [/ALHO FOZ DESCASCADO\b/g, "ALHO DESCASCADO"],
    [/ALHO DENTE\b/g, "ALHO DESCASCADO"],
    [/CAJU BANDEJA\b/g, "CAJU"],
    [/CAQUI RAMA FORTE BANDEJA\b/g, "CAQUI RAMA FORTE"],
    [/CARAMBOLA BANDEJA\b/g, "CARAMBOLA"],
    [/CEBOLINHA PIRULITO\b/g, "CEBOLA PIRULITO"],
    [/COGUMELO PARIS(?:\s+BANDEJA)?(?:\s+\d+G)?\b/g, "COGUMELO PARIS"],
    [/COGUMELO PORTOBELLO(?:\s+BANDEJA)?(?:\s+\d+G)?\b/g, "COGUMELO PORTOBELLO"],
    [/COGUMELO SHIMEJI(?:\s+BANDEJA)?(?:\s+\d+G)?\b/g, "COGUMELO SHIMEJI"],
    [/COGUMELO SHITAKE(?:\s+BANDEJA)?(?:\s+\d+G)?\b/g, "COGUMELO SHITAKE"],
    [/GOIABA VERMELHA\b/g, "GOIABA"],
    [/GOIABA GRANEL\b/g, "GOIABA"],
    [/COCO SECO\b/g, "COCO SECO"],
    [/COCO VERDE\b/g, "COCO VERDE"],
    [/LIMAO THAITI\b/g, "LIMAO"],
    [/LIMAO TAHITI\b/g, "LIMAO"],
    [/MACA GALA SUPER K\b/g, "MACA 850G"],
    [/MACA GALA BENASSI\b/g, "MACA 850G"],
    [/MACA PACOTE\b/g, "MACA 850G"],
    [/MACA PCT\b/g, "MACA 850G"],
    [/MACA FUG\b/g, "MACA FUJI"],
    [/MACA GALA NACIONAL\b/g, "MACA GALA"],
    [/MACA RED(?: IMPORT)?\b/g, "MACA RED IMPORT"],
    [/MACA GRASMIT\b/g, "MACA VERDE GRAN"],
    [/MACA GRANSMHT\b/g, "MACA VERDE GRAN"],
    [/MACA GRANSMITH\b/g, "MACA VERDE GRAN"],
    [/MAMAO PAPAYA\b/g, "MAMAO HAVAI"],
    [/MAMAO FORMOSA\b/g, "MAMAO FORMOSA"],
    [/MANGA TOMY\b/g, "MANGA TOMMY"],
    [/GOIABA\b/g, "GOIABA"],
    [/KIWI KILO\b/g, "KIWI"],
    [/KIWI IMPORTADO\b/g, "KIWI"],
    [/LARANJA BAIA\b/g, "LARANJA BAHIA"],
    [/LARANJA SELETA\b/g, "LARANJA SELETA"],
    [/LARANJA PERA\b/g, "LARANJA PERA"],
    [/MELAO PELE(?: DE)? SAPO\b/g, "MELAO VERDE"],
    [/MILHO VERDE BANDEJA\b/g, "MILHO VERDE"],
    [/MILHO (?:BAND|BDJ)\b/g, "MILHO VERDE"],
    [/MORANGO BJ\b/g, "MORANGO"],
    [/MORANGO BANDEJA\b/g, "MORANGO"],
    [/PITAYA(?:\s+BANDEJA)?(?:\s+\d+G)?\b/g, "PITAYA"],
    [/OVO[S]?\s+BRANCO\s+GRANDE\s+12\s*UN\b/g, "OVOS 12"],
    [/OVO[S]?\s+VERMELH[OA]?\s+GRANDE\s+12\s*UN\b/g, "OVOS VERMELHOS 12"],
    [/OVO[S]?\s+BRANCO.*30\b/g, "OVOS BRANCOS 30"],
    [/OVO[S]?\s+BRANCO.*20\b/g, "OVOS BRANCO 20"],
    [/OVO[S]?.*CODORNA.*30\b/g, "OVOS CODORNA 30"],
    [/CODORNA\b/g, "OVOS CODORNA"],
    [/OVO[S]?\s+VERM\b.*12\b/g, "OVOS VERMELHOS 12"],
    [/OVO[S]?.*VERMELH.*12\b/g, "OVOS VERMELHOS 12"],
    [/OVO[S]?.*VERMELH.*20\b/g, "OVOS VERMELHOS 20"],
    [/OVO[S]?.*VERMELH.*30\b/g, "OVOS VERMELHOS 30"],
    [/OVO[S]?.*\bC\b.*12\b/g, "OVOS 12"],
    [/PERA WILLIAM[SN]?\b/g, "PÊRA WILLIANS"],
    [/PERA WILLIANS\b/g, "PÊRA WILLIANS"],
    [/PERA PORTUGUESA\b/g, "PERA PORTUGUESA"],
    [/PIMENTAO AMARELO\b/g, "PIMENTAO AMARELO"],
    [/PIMENTAO VERDE\b/g, "PIMENTAO"],
    [/PIMENTAO VERMEL(?:HO)?\b/g, "PIMENTAO"],
    [/PIMENTAO BRANCO\b/g, "PIMENTAO"],
    [/TANGERINA POKAN\b/g, "TANGERINA PONKAN"],
    [/TANGERINA IMP\b/g, "TANGERINA IMPORTADA"],
    [/TANGERINA IMPORTADA\b/g, "TANGERINA IMPORTADA"],
    [/TANGERINA MORGOTE\b/g, "TANGERINA MORCOTE"],
    [/TOMATE SWEET GRAPE\b/g, "TOMATE SWEET 180"],
    [/TOMATE SWEET\b/g, "TOMATE SWEET 180"],
    [/UVA CRINSON\b/g, "UVA CRIMSON"],
    [/UVA ITALIA\b/g, "UVA ITALIA"],
    [/UVA RED GLOBE\b/g, "UVA RED GLOB"],
    [/UVA REDGLOBE\b/g, "UVA RED GLOB"],
    [/UVA THOMPSON VERDE\b/g, "UVA THOMPSON"],
    [/UVA VITORIA SEM\b/g, "UVA VITORIA"],
    [/QUIABO(?:\s+(?:BAND|BANDEJA|EMBALADO))?(?:\s+\d+G)?\b/g, "QUIABO 300G"],
    [/VAGEM MACARRAO\b/g, "VAGEM MANT"],
    [/VAGEM MANTEIGA\b/g, "VAGEM MANT"],
  ];

  const state = {
    inputFiles: [],
    generatedZipBlob: null,
    generatedFiles: [],
  };

  const elements = {
    input: document.getElementById("input-files"),
    dropzone: document.getElementById("dropzone"),
    fileCount: document.getElementById("file-count"),
    fileList: document.getElementById("file-list"),
    resultList: document.getElementById("result-list"),
    statusCard: document.getElementById("status-card"),
    generateButton: document.getElementById("generate-button"),
    downloadButton: document.getElementById("download-button"),
    clearButton: document.getElementById("clear-button"),
  };

  function normalizeText(value) {
    return String(value ?? "")
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .toUpperCase()
      .replace(/[.'â€™"()]/g, " ")
      .replace(/[-/,:;]/g, " ")
      .replace(/\s+/g, " ")
      .trim();
  }

  function canonicalizeProductName(rawName) {
    let text = normalizeText(rawName);

    if (/^OVO[S]?\s+RAIAR\s+GRANDE\s+ORGANICO\s+20\s*UN$/.test(text)) {
      return text;
    }
    if (/^OVO[S]?\s+SITIO\s+COCORICO\s+CAIPIRA\s+10\s*UN$/.test(text)) {
      return text;
    }

    for (const [pattern, replacement] of PHRASE_REPLACEMENTS) {
      text = text.replace(pattern, replacement);
    }

    if (text.startsWith("TANGERINA IMPORTADA")) {
      return "TANGERINA IMPORTADA";
    }

    text = text.replace(/\s+/g, " ").trim();

    const filtered = text
      .split(" ")
      .filter((token) => token && !STOP_WORDS.has(token))
      .join(" ")
      .trim();

    let canonical = filtered
      .replace(/\bD AGUA\b/g, "DAGUA")
      .replace(/\bSERGIPANA\b/g, "ABOBORA SERGIPANA")
      .replace(/\bJAPONESA\b/g, "ABOBORA JAPONESA")
      .replace(/\bMORANGA\b/g, "ABOBORA MORANGA")
      .replace(/\bPESCOCO\b/g, "ABOBORA PESCOCO")
      .replace(/\bBAIANA\b/g, "ABOBORA BAIANA")
      .replace(/^ALHO\b(?!.*DESCASCADO).*$/, "ALHO")
      .replace(/^ALHO .*DESCASCADO.*$/, "ALHO DESCASCADO")
      .replace(/^BATATA BAROA.*$/, "BATATA BAROA")
      .replace(/^BATATA BOLINHA.*$/, "BATATA BOLINHA")
      .replace(/^BATATA ASTERIX.*$/, "BATATA ASTERIX")
      .replace(/^BATATA DOCE.*$/, "BATATA DOCE")
      .replace(/^BATATA INGLESA.*$/, "BATATA INGLESA")
      .replace(/^BATATA SUJA.*$/, "BATATA SUJA")
      .replace(/\s+/g, " ")
      .trim();

    if (canonical.includes("ABOBORA")) {
      if (canonical.includes("SERGIPANA")) {
        return "ABOBORA SERGIPANA";
      }
      if (canonical.includes("JAPONESA")) {
        return "ABOBORA JAPONESA";
      }
      if (canonical.includes("MORANGA")) {
        return "ABOBORA MORANGA";
      }
      if (canonical.includes("PESCOCO")) {
        return "ABOBORA PESCOCO";
      }
      if (canonical.includes("BAIANA")) {
        return "ABOBORA BAIANA";
      }
    }

    if (canonical.includes("CODORNA")) {
      return "OVOS CODORNA";
    }

    if (
      canonical.includes("MACA") &&
      (canonical.includes("850G") ||
        canonical.includes("SUPER K") ||
        canonical.includes("BENASSI") ||
        canonical.includes("PCT"))
    ) {
      return "MACA 850G";
    }
    if (canonical === "MACA") {
      return "MACA GALA";
    }

    if (canonical.includes("UVA")) {
      if (canonical.includes("THOMPSON")) {
        return "UVA THOMPSON";
      }
      if (canonical.includes("RED GLOB") || canonical.includes("REDGLOB")) {
        return "UVA RED GLOB";
      }
      if (canonical.includes("CRIMSON") || canonical.includes("CRINSON")) {
        return "UVA CRIMSON";
      }
      if (canonical.includes("ITALIA")) {
        return "UVA ITALIA";
      }
      if (canonical.includes("ROSADA")) {
        return "UVA ROSADA";
      }
      if (canonical.includes("VITORIA")) {
        return "UVA VITORIA";
      }
    }

    if (canonical.includes("CAJU")) {
      return "CAJU";
    }

    if (canonical.startsWith("CAQUI")) {
      return "CAQUI";
    }

    if (canonical.includes("CARAMBOLA")) {
      return "CARAMBOLA";
    }

    if (canonical.includes("MILHO VERDE")) {
      return "MILHO VERDE";
    }

    if (canonical.startsWith("MORANGO")) {
      return "MORANGO";
    }

    if (canonical.startsWith("ROMA")) {
      return "ROMA";
    }

    if (canonical.includes("MELAO") && canonical.includes("SAPO")) {
      return "MELAO VERDE";
    }

    if (canonical.includes("TOMATE SWEET")) {
      return "TOMATE SWEET 180";
    }

    if (canonical.startsWith("TANGERINA PONKAN")) {
      return "TANGERINA PONKAN";
    }
    if (canonical.startsWith("TANGERINA IMPORTADA")) {
      return "TANGERINA IMPORTADA";
    }

    return canonical;
  }

  function outputProductName(rawName, productKey) {
    const text = normalizeText(rawName);

    if (/^OVO[S]?\s+RAIAR\s+GRANDE\s+ORGANICO\s+20\s*UN$/.test(text)) {
      return text;
    }
    if (/^OVO[S]?\s+SITIO\s+COCORICO\s+CAIPIRA\s+10\s*UN$/.test(text)) {
      return text;
    }
    if (/^OVO[S]?\s+BRANCO\s+GRANDE\s+12\s*UN$/.test(text)) {
      return "OVOS C/12";
    }
    if (/^OVO[S]?\s+VERMELH[OA]?\s+GRANDE\s+12\s*UN$/.test(text)) {
      return "OVOS VERMELHOS C/12";
    }
    if (productKey === "TANGERINA IMPORTADA") {
      return "TANGERINA IMPORTADA";
    }
    if (productKey === "COGUMELO PARIS") {
      return "COGUMELO PARIS BANDEJA 200g";
    }
    if (productKey === "COGUMELO PORTOBELLO") {
      return "COGUMELO PORTOBELLO BANDEJA 200g";
    }
    if (productKey === "COGUMELO SHIMEJI") {
      return "COGUMELO SHIMEJI BANDEJA 200g";
    }
    if (productKey === "COGUMELO SHITAKE") {
      return "COGUMELO SHITAKE BANDEJA 200g";
    }
    if (productKey === "OVOS BRANCO 20") {
      return "OVOS BRANCO C/ 20";
    }
    if (productKey === "OVOS BRANCOS 30") {
      return "OVOS  BRANCOS C/30";
    }
    if (productKey === "QUIABO 300G") {
      return "QUIABO BANDEJA 300G";
    }

    return productKey || text;
  }

  function parseLooseItem(text) {
    const cleaned = String(text ?? "").trim();
    const match = cleaned.match(
      /^(.*?)(?:\s*[-]?\s*)(\d+(?:[.,]\d+)?)\s*(?:KG|KILO|UNID|UNIDADE|UN|UND|UNID\.|CX|PCT|PACOTE|BDJ)?\s*$/i,
    );

    if (!match) {
      return null;
    }

    if (/C\/\s*$/i.test(match[1])) {
      return null;
    }

    return {
      productName: match[1].trim(),
      quantity: normalizeQuantity(match[2]),
    };
  }

  function normalizeQuantity(value) {
    if (value == null || value === "") {
      return null;
    }

    const text = String(value).trim().replace(/\./g, "").replace(",", ".");
    const parsed = Number(text);
    return Number.isNaN(parsed) ? null : parsed;
  }

  function quantityForCell(value) {
    if (value == null) {
      return null;
    }

    if (Number.isInteger(value)) {
      return value;
    }

    return Number(value.toFixed(2).replace(/\.?0+$/, ""));
  }

  function formatDate(value) {
    const year = value.getFullYear();
    const month = String(value.getMonth() + 1).padStart(2, "0");
    const day = String(value.getDate()).padStart(2, "0");
    return `${day}/${month}/${year}`;
  }

  function canonicalStoreName(rawName) {
    const normalized = normalizeText(rawName).replace(/\d+\s*/g, "");
    return normalized
      .replace(/\bDA ROCHA\b/g, "")
      .replace(/\bII\b/g, "")
      .replace(/\bSTA CRUZ DA SERRA\b/g, "STA CRUZ")
      .replace(/\s+/g, " ")
      .trim();
  }

  function worksheetValueToString(value) {
    if (value == null) {
      return "";
    }

    if (typeof value === "object") {
      if (value.text != null) {
        return String(value.text);
      }

      if (value.richText) {
        return value.richText.map((item) => item.text ?? "").join("");
      }

      if (value.result != null) {
        return String(value.result);
      }

      if (value.formula != null && value.result == null) {
        return "";
      }
    }

    return String(value);
  }

  function findBranchName(worksheet) {
    for (let rowNumber = 1; rowNumber <= Math.min(worksheet.rowCount, 20); rowNumber += 1) {
      for (
        let columnNumber = 1;
        columnNumber <= Math.min(worksheet.getRow(rowNumber).cellCount || 30, 30);
        columnNumber += 1
      ) {
        const current = normalizeText(
          worksheetValueToString(worksheet.getRow(rowNumber).getCell(columnNumber).value),
        );
        if (current !== "FILIAL") {
          continue;
        }

        for (let offset = 1; offset <= 3; offset += 1) {
          const sibling = worksheetValueToString(
            worksheet.getRow(rowNumber).getCell(columnNumber + offset).value,
          ).trim();
          if (sibling) {
            return sibling;
          }
        }
      }
    }

    return "";
  }

  function getStoreKey(workbook, fileName) {
    const worksheet = workbook.worksheets[0];
    const fallbackName = fileName.replace(/\.[^.]+$/i, "");
    const branchName = findBranchName(worksheet) || fallbackName;
    return canonicalStoreName(branchName);
  }

  function findStructuredColumns(worksheet) {
    for (let rowNumber = 1; rowNumber <= Math.min(worksheet.rowCount, 20); rowNumber += 1) {
      let productColumn = null;
      let quantityColumn = null;
      const row = worksheet.getRow(rowNumber);
      const maxColumns = Math.max(row.cellCount || 0, 30);

      for (let columnNumber = 1; columnNumber <= maxColumns; columnNumber += 1) {
        const cellText = normalizeText(worksheetValueToString(row.getCell(columnNumber).value));
        if (cellText === "PRODUTO") {
          productColumn = columnNumber;
        }

        if (cellText === "QTDE" || cellText === "QUANTIDADE" || cellText === "QTD") {
          quantityColumn = columnNumber;
        }
      }

      if (productColumn && quantityColumn) {
        return { headerRow: rowNumber, productColumn, quantityColumn };
      }
    }

    return null;
  }

  function parseStructuredItems(worksheet, columns) {
    const items = [];

    for (let rowNumber = columns.headerRow + 1; rowNumber <= worksheet.rowCount; rowNumber += 1) {
      const productName = worksheetValueToString(
        worksheet.getRow(rowNumber).getCell(columns.productColumn).value,
      ).trim();
      const quantity = normalizeQuantity(
        worksheet.getRow(rowNumber).getCell(columns.quantityColumn).value,
      );

      if (!productName || quantity == null) {
        continue;
      }

      items.push({ productName, quantity });
    }

    return items;
  }

  function parseWorkbookItems(workbook) {
    const worksheet = workbook.worksheets[0];
    const structuredColumns = findStructuredColumns(worksheet);

    if (structuredColumns) {
      return parseStructuredItems(worksheet, structuredColumns);
    }

    const items = [];
    for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber += 1) {
      const rawValue = worksheet.getCell(`A${rowNumber}`).value;
      if (!rawValue) {
        continue;
      }

      const parsed = parseLooseItem(rawValue);
      if (parsed) {
        items.push(parsed);
        continue;
      }

      const productName = worksheetValueToString(rawValue).trim();
      const quantity = normalizeQuantity(worksheet.getCell(`B${rowNumber}`).value);
      if (productName && quantity != null) {
        items.push({ productName, quantity });
      }
    }

    return items;
  }

  function buildTemplateMap(worksheet, sections) {
    const map = new Map();

    sections.forEach((section, sectionIndex) => {
      for (let rowNumber = PRODUCT_START_ROW; rowNumber <= worksheet.rowCount; rowNumber += 1) {
        const label = worksheet.getCell(`${section.productColumn}${rowNumber}`).value;
        if (!label) {
          continue;
        }

        const canonical = canonicalizeProductName(label);
        if (!canonical || map.has(canonical)) {
          continue;
        }

        map.set(canonical, { rowNumber, label: String(label), sectionIndex });
      }
    });

    return map;
  }

  function clearSectionStoreValues(worksheet, sections) {
    sections.forEach((section) => {
      const columns = Object.values(section.storeColumns);
      for (const column of columns) {
        for (let rowNumber = PRODUCT_START_ROW; rowNumber <= worksheet.rowCount; rowNumber += 1) {
          worksheet.getCell(`${column}${rowNumber}`).value = null;
        }
      }
    });
  }

  function cloneStyle(style) {
    return JSON.parse(JSON.stringify(style ?? {}));
  }

  function findLastConfiguredRow(worksheet, section) {
    for (let rowNumber = worksheet.rowCount; rowNumber >= PRODUCT_START_ROW; rowNumber -= 1) {
      const productText = worksheetValueToString(
        worksheet.getCell(`${section.productColumn}${rowNumber}`).value,
      ).trim();
      if (productText) {
        return rowNumber;
      }
    }

    return PRODUCT_START_ROW;
  }

  function applyReviewFill(cell) {
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFFF00" },
    };
  }

  function isReviewFill(fill) {
    if (!fill || fill.type !== "pattern" || fill.pattern !== "solid") {
      return false;
    }

    const color = fill.fgColor?.argb ?? fill.bgColor?.argb ?? "";
    return color.toUpperCase() === "FFFFFF00";
  }

  function buildSectionCategories(worksheet, section) {
    const categories = [];
    let currentCategory = [];

    for (let rowNumber = PRODUCT_START_ROW; rowNumber <= worksheet.rowCount; rowNumber += 1) {
      const productText = worksheetValueToString(
        worksheet.getCell(`${section.productColumn}${rowNumber}`).value,
      ).trim();

      if (!productText) {
        if (currentCategory.length) {
          categories.push(currentCategory);
          currentCategory = [];
        }
        continue;
      }

      currentCategory.push(rowNumber);
    }

    if (currentCategory.length) {
      categories.push(currentCategory);
    }

    return categories;
  }

  async function loadProductLists() {
    if (!productListCachePromise) {
      productListCachePromise = Promise.all(
        Object.entries(PRODUCT_LIST_FILES).map(async ([key, filePath]) => {
          const response = await fetch(filePath);
          if (!response.ok) {
            throw new Error(`Nao foi possivel carregar a lista ${filePath}.`);
          }

          return [key, await response.json()];
        }),
      ).then((entries) => Object.fromEntries(entries));
    }

    return productListCachePromise;
  }

  function getConfiguredSectionProducts(config, section, productLists) {
    const sectionProducts = productLists[section.productListKey];
    if (!Array.isArray(sectionProducts)) {
      throw new Error(`Lista de produtos nao configurada: ${section.productListKey}`);
    }
    return sectionProducts;
  }

  function cloneSectionRowLayout(worksheet, section, sourceRowNumber, targetRowNumber) {
    worksheet.getRow(targetRowNumber).height = worksheet.getRow(sourceRowNumber).height;

    const productCell = worksheet.getCell(`${section.productColumn}${targetRowNumber}`);
    const prototypeProductCell = worksheet.getCell(`${section.productColumn}${sourceRowNumber}`);
    productCell.style = cloneStyle(prototypeProductCell.style);

    Object.values(section.storeColumns).forEach((column) => {
      const cell = worksheet.getCell(`${column}${targetRowNumber}`);
      const prototypeCell = worksheet.getCell(`${column}${sourceRowNumber}`);
      cell.style = cloneStyle(prototypeCell.style);
      cell.value = null;
    });
  }

  function applyConfiguredProductLists(worksheet, config, productLists) {
    config.sections.forEach((section) => {
      const products = getConfiguredSectionProducts(config, section, productLists);
      const lastConfiguredRow = Math.max(findLastConfiguredRow(worksheet, section), PRODUCT_START_ROW);
      const lastRequiredRow = PRODUCT_START_ROW + products.length - 1;

      for (let rowNumber = PRODUCT_START_ROW; rowNumber <= lastRequiredRow; rowNumber += 1) {
        if (rowNumber > lastConfiguredRow) {
          cloneSectionRowLayout(worksheet, section, lastConfiguredRow, rowNumber);
        }

        const label = products[rowNumber - PRODUCT_START_ROW];
        worksheet.getCell(`${section.productColumn}${rowNumber}`).value = label || null;
      }

      for (
        let rowNumber = Math.max(lastRequiredRow + 1, PRODUCT_START_ROW);
        rowNumber <= lastConfiguredRow;
        rowNumber += 1
      ) {
        worksheet.getCell(`${section.productColumn}${rowNumber}`).value = null;
      }
    });
  }

  function snapshotSectionRow(worksheet, section, rowNumber) {
    const storeColumns = Object.values(section.storeColumns);
    const productCell = worksheet.getCell(`${section.productColumn}${rowNumber}`);
    const quantities = storeColumns.map((column) => worksheet.getCell(`${column}${rowNumber}`));

    return {
      isReview:
        isReviewFill(productCell.fill) || quantities.some((cell) => isReviewFill(cell.fill)),
      product: {
        value: productCell.value,
        style: cloneStyle(productCell.style),
      },
      quantities: Object.fromEntries(
        storeColumns.map((column) => {
          const cell = worksheet.getCell(`${column}${rowNumber}`);
          return [
            column,
            {
              value: cell.value,
              style: cloneStyle(cell.style),
            },
          ];
        }),
      ),
    };
  }

  function hasRowQuantity(row, storeColumns) {
    return storeColumns.some(
      (column) => row.quantities[column].value != null && row.quantities[column].value !== "",
    );
  }

  function getSectionRows(worksheet, section, categoryRows) {
    const storeColumns = Object.values(section.storeColumns);

    return categoryRows
      .map((rowNumber) => snapshotSectionRow(worksheet, section, rowNumber))
      .filter(
        (row) => worksheetValueToString(row.product.value).trim() && hasRowQuantity(row, storeColumns),
      );
  }

  function rewriteSectionRows(worksheet, section, rows, categoryCount) {
    const storeColumns = Object.values(section.storeColumns);
    const startRow = 9;
    const lastRow = worksheet.rowCount;

    for (let rowNumber = startRow; rowNumber <= lastRow; rowNumber += 1) {
      worksheet.getCell(`${section.productColumn}${rowNumber}`).value = null;
      worksheet.getCell(`${section.productColumn}${rowNumber}`).fill =
        cloneStyle(worksheet.getCell(`${section.productColumn}${rowNumber}`).style).fill;

      storeColumns.forEach((column) => {
        worksheet.getCell(`${column}${rowNumber}`).value = null;
        worksheet.getCell(`${column}${rowNumber}`).fill =
          cloneStyle(worksheet.getCell(`${column}${rowNumber}`).style).fill;
      });
    }

    const writeRow = (row, targetRow) => {
      const productCell = worksheet.getCell(`${section.productColumn}${targetRow}`);
      productCell.style = cloneStyle(row.product.style);
      productCell.value = row.product.value;

      storeColumns.forEach((column) => {
        const quantityCell = worksheet.getCell(`${column}${targetRow}`);
        quantityCell.style = cloneStyle(row.quantities[column].style);
        quantityCell.value = row.quantities[column].value;
      });
    };

    let currentRow = startRow;

    rows.forEach((group, groupIndex) => {
      group.forEach((row) => {
        writeRow(row, currentRow);
        currentRow += 1;
      });

      if (group.length && groupIndex < categoryCount - 1) {
        currentRow += 1;
      }
    });

    for (let rowNumber = currentRow; rowNumber <= lastRow; rowNumber += 1) {
      const productCell = worksheet.getCell(`${section.productColumn}${rowNumber}`);
      productCell.style = {};
      productCell.value = null;

      storeColumns.forEach((column) => {
        const quantityCell = worksheet.getCell(`${column}${rowNumber}`);
        quantityCell.style = {};
        quantityCell.value = null;
      });
    }
  }

  function compactWorksheetSections(worksheet, sections, sectionCategoryMap) {
    sections.forEach((section, sectionIndex) => {
      const categories = sectionCategoryMap[sectionIndex] ?? [];
      const groupedRows = categories
        .map((categoryRows) => getSectionRows(worksheet, section, categoryRows))
        .filter((rows) => rows.length);

      const reviewRows = getSectionRows(
        worksheet,
        section,
        Array.from(
          { length: Math.max(worksheet.rowCount - (PRODUCT_START_ROW - 1), 0) },
          (_, index) => index + PRODUCT_START_ROW,
        ),
      ).filter((row) => row.isReview);

      if (reviewRows.length) {
        groupedRows.push(reviewRows);
      }

      rewriteSectionRows(worksheet, section, groupedRows, groupedRows.length);
    });
  }

  function createUnknownItemWriter(worksheet, config) {
    const firstSection = config.sections[0];
    const prototypeRow = findLastConfiguredRow(worksheet, firstSection);
    const unknownRows = new Map();
    let nextRowNumber = worksheet.rowCount + 1;

    return (productKey, originalName, storeKey, quantity) => {
      let rowNumber = unknownRows.get(productKey);

      if (!rowNumber) {
        rowNumber = nextRowNumber;
        nextRowNumber += 1;
        unknownRows.set(productKey, rowNumber);

        worksheet.getRow(rowNumber).height = worksheet.getRow(prototypeRow).height;

        const productCell = worksheet.getCell(`${firstSection.productColumn}${rowNumber}`);
        const prototypeProductCell = worksheet.getCell(
          `${firstSection.productColumn}${prototypeRow}`,
        );
        productCell.style = cloneStyle(prototypeProductCell.style);
        productCell.value = outputProductName(originalName, productKey);
        applyReviewFill(productCell);

        Object.values(firstSection.storeColumns).forEach((column) => {
          const cell = worksheet.getCell(`${column}${rowNumber}`);
          const prototypeCell = worksheet.getCell(`${column}${prototypeRow}`);
          cell.style = cloneStyle(prototypeCell.style);
        });
      }

      const targetColumn = firstSection.storeColumns[storeKey];
      if (!targetColumn) {
        return;
      }

      const quantityCell = worksheet.getCell(`${targetColumn}${rowNumber}`);
      quantityCell.value = quantityForCell(quantity);
      applyReviewFill(quantityCell);
    };
  }

  function replaceDateText(existingValue, formattedDate) {
    const currentText = String(existingValue ?? "");
    const normalizedText = normalizeText(currentText);

    if (normalizedText.startsWith("PEDIDOS CEASA")) {
      return `PEDIDOS CEASA ${formattedDate}`;
    }
    if (normalizedText.startsWith("DATA")) {
      return `DATA ${formattedDate}`;
    }

    if (!currentText) {
      return `Data ${formattedDate}`;
    }

    if (/\d{2}\/\d{2}\/\d{4}/.test(currentText)) {
      return currentText.replace(/\d{2}\/\d{2}\/\d{4}/, formattedDate);
    }

    return `${currentText} ${formattedDate}`.trim();
  }

  function updateHeaderDates(worksheet, formattedDate) {
    for (let rowNumber = 1; rowNumber <= Math.min(worksheet.rowCount, 3); rowNumber += 1) {
      const row = worksheet.getRow(rowNumber);
      const maxColumns = Math.max(row.cellCount || 0, 10);

      for (let columnNumber = 1; columnNumber <= maxColumns; columnNumber += 1) {
        const cell = row.getCell(columnNumber);
        const text = worksheetValueToString(cell.value);
        const normalizedText = normalizeText(text);

        if (normalizedText.startsWith("PEDIDOS CEASA") || normalizedText.startsWith("DATA")) {
          cell.value = replaceDateText(text, formattedDate);
        }
      }
    }
  }

  async function readWorkbookFromBuffer(buffer) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);
    return workbook;
  }

  async function loadInputEntries(files) {
    const inputEntries = [];

    for (const uploaded of files) {
      const file = uploaded.file;

      if (!/\.(xlsx|xlsm)$/i.test(uploaded.name)) {
        continue;
      }

      const buffer = await file.arrayBuffer();
      const workbook = await readWorkbookFromBuffer(buffer);
      const storeKey = getStoreKey(workbook, uploaded.name);
      const items = parseWorkbookItems(workbook);

      inputEntries.push({
        fileName: uploaded.name,
        storeKey,
        items,
      });
    }

    return inputEntries;
  }

  async function fetchTemplateBuffer(templateName) {
    const response = await fetch(`template/mapa/${encodeURIComponent(templateName)}`);
    if (!response.ok) {
      throw new Error(`Nao foi possivel carregar o template ${templateName}.`);
    }

    return response.arrayBuffer();
  }

  async function generateMap(config, inputs, now) {
    const templateBuffer = await fetchTemplateBuffer(config.templateName);
    const workbook = await readWorkbookFromBuffer(templateBuffer);
    const worksheet = workbook.worksheets[0];
    const productLists = await loadProductLists();
    applyConfiguredProductLists(worksheet, config, productLists);
    const templateMap = buildTemplateMap(worksheet, config.sections);
    const sectionCategoryMap = config.sections.map((section) => buildSectionCategories(worksheet, section));
    const writeUnknownItem = createUnknownItemWriter(worksheet, config);

    clearSectionStoreValues(worksheet, config.sections);
    updateHeaderDates(worksheet, formatDate(now));

    for (const input of inputs) {
      const resolvedStore = config.storeAliases[input.storeKey];
      if (!resolvedStore) {
        continue;
      }

      for (const item of input.items) {
        const productKey = canonicalizeProductName(item.productName);
        if (!productKey) {
          continue;
        }

        const row = templateMap.get(productKey) ?? null;
        if (!row) {
          writeUnknownItem(productKey, item.productName, resolvedStore, item.quantity);
          continue;
        }

        const targetColumn = config.sections[row.sectionIndex].storeColumns[resolvedStore];
        if (!targetColumn) {
          continue;
        }

        if (productKey === "TANGERINA IMPORTADA") {
          worksheet.getCell(`${config.sections[row.sectionIndex].productColumn}${row.rowNumber}`).value =
            "TANGERINA IMPORTADA";
        }

        worksheet.getCell(`${targetColumn}${row.rowNumber}`).value = quantityForCell(item.quantity);
      }
    }

    compactWorksheetSections(worksheet, config.sections, sectionCategoryMap);

    const buffer = await workbook.xlsx.writeBuffer();
    return {
      fileName: config.outputName,
      buffer,
    };
  }

  function setStatus(message, tone) {
    elements.statusCard.textContent = message;
    elements.statusCard.classList.remove("is-error", "is-success");
    if (tone) {
      elements.statusCard.classList.add(tone);
    }
  }

  function formatBytes(bytes) {
    if (bytes < 1024) {
      return `${bytes} B`;
    }
    if (bytes < 1024 * 1024) {
      return `${(bytes / 1024).toFixed(1)} KB`;
    }
    return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
  }

  function renderFileList() {
    elements.fileCount.textContent = `${state.inputFiles.length} arquivo${state.inputFiles.length === 1 ? "" : "s"}`;

    if (!state.inputFiles.length) {
      elements.fileList.innerHTML = '<li class="empty-state">Nenhum arquivo carregado ainda.</li>';
      return;
    }

    elements.fileList.innerHTML = state.inputFiles
      .map(
        (file) => `
          <li class="file-item">
            <div class="file-meta">
              <span class="file-name">${escapeHtml(file.name)}</span>
              <span class="file-subtitle">${escapeHtml(file.source)} • ${formatBytes(file.size)}</span>
            </div>
          </li>
        `,
      )
      .join("");
  }

  function renderResults() {
    if (!state.generatedFiles.length) {
      elements.resultList.innerHTML = '<li class="empty-state">Nenhum mapa gerado ainda.</li>';
      return;
    }

    elements.resultList.innerHTML = state.generatedFiles
      .map(
        (file) => `
          <li class="result-item">
            <div class="result-meta">
              <span class="result-name">${escapeHtml(file.fileName)}</span>
              <span class="result-subtitle">${formatBytes(file.size)}</span>
            </div>
          </li>
        `,
      )
      .join("");
  }

  function escapeHtml(value) {
    return String(value)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function updateButtons() {
    elements.generateButton.disabled = !state.inputFiles.length;
    elements.downloadButton.disabled = !state.generatedZipBlob;
  }

  function clearState() {
    state.inputFiles = [];
    state.generatedZipBlob = null;
    state.generatedFiles = [];
    elements.input.value = "";
    renderFileList();
    renderResults();
    updateButtons();
    setStatus("Arquivos limpos. Voce pode enviar um novo lote.", null);
  }

  async function filesFromZip(file) {
    const zip = await JSZip.loadAsync(await file.arrayBuffer());
    const extracted = [];

    for (const [entryName, entry] of Object.entries(zip.files)) {
      if (entry.dir || !/\.(xlsx|xlsm)$/i.test(entryName)) {
        continue;
      }

      const buffer = await entry.async("arraybuffer");
      const cleanName = entryName.split("/").pop() || entryName;
      extracted.push(
        new File([buffer], cleanName, {
          type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        }),
      );
    }

    return extracted;
  }

  async function addFiles(fileList) {
    const collected = [];

    for (const originalFile of Array.from(fileList)) {
      if (/\.zip$/i.test(originalFile.name)) {
        const extracted = await filesFromZip(originalFile);
        extracted.forEach((file) => {
          collected.push({
            name: file.name,
            size: file.size,
            file,
            source: `extraido de ${originalFile.name}`,
          });
        });
        continue;
      }

      if (/\.(xlsx|xlsm)$/i.test(originalFile.name)) {
        collected.push({
          name: originalFile.name,
          size: originalFile.size,
          file: originalFile,
          source: "upload direto",
        });
      }
    }

    const deduped = new Map();
    [...state.inputFiles, ...collected].forEach((file) => {
        const key = `${file.name}::${file.size}`;
        deduped.set(key, file);
    });

    state.inputFiles = Array.from(deduped.values());
    state.generatedZipBlob = null;
    state.generatedFiles = [];
    renderFileList();
    renderResults();
    updateButtons();

    if (!state.inputFiles.length) {
      setStatus("Nenhuma planilha valida foi encontrada no envio.", "is-error");
      return;
    }

    setStatus(`${state.inputFiles.length} arquivo(s) prontos para processamento.`, null);
  }

  async function handleGenerate() {
    if (!state.inputFiles.length) {
      setStatus("Envie pelo menos uma planilha antes de gerar os mapas.", "is-error");
      return;
    }

    elements.generateButton.disabled = true;
    elements.downloadButton.disabled = true;
    setStatus("Gerando mapas a partir dos arquivos enviados...", null);

    try {
      const inputs = await loadInputEntries(state.inputFiles);
      if (!inputs.length) {
        throw new Error("Nenhuma planilha Excel valida foi lida.");
      }

      const now = new Date();
      const outputs = [];

      for (const config of MAP_CONFIGS) {
        const generated = await generateMap(config, inputs, now);
        outputs.push({
          fileName: generated.fileName,
          size: generated.buffer.byteLength,
          buffer: generated.buffer,
        });
      }

      const zip = new JSZip();
      outputs.forEach((file) => {
        zip.file(file.fileName, file.buffer);
      });

      state.generatedZipBlob = await zip.generateAsync({ type: "blob" });
      state.generatedFiles = outputs.map(({ fileName, size }) => ({ fileName, size }));
      renderResults();
      updateButtons();
      setStatus("Mapas gerados com sucesso. O download ja esta liberado.", "is-success");
    } catch (error) {
      console.error(error);
      state.generatedZipBlob = null;
      state.generatedFiles = [];
      renderResults();
      updateButtons();
      setStatus(error.message || "Ocorreu um erro ao gerar os mapas.", "is-error");
    } finally {
      elements.generateButton.disabled = !state.inputFiles.length;
    }
  }

  function handleDownload() {
    if (!state.generatedZipBlob) {
      return;
    }

    const now = new Date();
    const stamp = [
      now.getFullYear(),
      String(now.getMonth() + 1).padStart(2, "0"),
      String(now.getDate()).padStart(2, "0"),
      String(now.getHours()).padStart(2, "0"),
      String(now.getMinutes()).padStart(2, "0"),
    ].join("-");

    const url = URL.createObjectURL(state.generatedZipBlob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = `mapas-gerados-${stamp}.zip`;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    URL.revokeObjectURL(url);
  }

  function wireDragAndDrop() {
    ["dragenter", "dragover"].forEach((eventName) => {
      elements.dropzone.addEventListener(eventName, (event) => {
        event.preventDefault();
        elements.dropzone.classList.add("is-dragover");
      });
    });

    ["dragleave", "drop"].forEach((eventName) => {
      elements.dropzone.addEventListener(eventName, (event) => {
        event.preventDefault();
        elements.dropzone.classList.remove("is-dragover");
      });
    });

    elements.dropzone.addEventListener("drop", async (event) => {
      await addFiles(event.dataTransfer.files);
    });
  }

  function init() {
    renderFileList();
    renderResults();
    updateButtons();
    wireDragAndDrop();

    elements.input.addEventListener("change", async (event) => {
      await addFiles(event.target.files);
    });
    elements.generateButton.addEventListener("click", handleGenerate);
    elements.downloadButton.addEventListener("click", handleDownload);
    elements.clearButton.addEventListener("click", clearState);
  }

  init();
})();
