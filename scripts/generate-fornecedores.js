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

const STORE_HEADER_LABELS = new Map([
  ["ANCHIETA", "ANCHIETA"],
  ["CACHAMBI", "CACHAMBI"],
  ["CERAMICA", "CERAMICA"],
  ["COELHO", "COELHO"],
  ["FREGUESIA", "FREGUESIA"],
  ["IRAJA", "IRAJÁ"],
  ["OLINDA", "OLINDA"],
  ["PIABETA", "PIABETA"],
  ["QUEIMADOS", "QUEIMADOS"],
  ["SANTA CRUZ", "STA.CRUZ"],
  ["SANTOS", "SANTOS"],
]);

const STORE_ORDER = [
  "IRAJA",
  "CACHAMBI",
  "SANTOS",
  "FREGUESIA",
  "ANCHIETA",
  "OLINDA",
  "CERAMICA",
  "COELHO",
  "PIABETA",
  "SANTA CRUZ",
  "QUEIMADOS",
];

const PRODUCT_REPLACEMENTS = [
  [/\bORGANIC[AO]S?\b/g, "ORGANICO"],
  [/^ABACAXI\b.*$/g, "ABACAXI"],
  [/\bABACAXI UNID\b/g, "ABACAXI"],
  [/\bABOBORA BAHIANA\b/g, "ABOBORA BAIANA"],
  [/\bBATATA BAROA BDJ\b/g, "BATATA BAROA"],
  [/\bBANANA D AGUA\b/g, "BANANA DAGUA"],
  [/\bBANANA DAGUA\b/g, "BANANA DAGUA"],
  [/^BANANA DA TERRA\b.*$/g, "BANANA DA TERRA"],
  [/^BANANA DAGUA\b.*$/g, "BANANA DAGUA"],
  [/^BANANA MACA\b.*$/g, "BANANA MACA"],
  [/^BANANA OURO\b.*$/g, "BANANA OURO"],
  [/^BANANA PRATA\s+ORGANICO\b.*$/g, "BANANA PRATA ORGANICO"],
  [/^BANANA PRATA(?!\s+ORGANICO\b)\b.*$/g, "BANANA PRATA"],
  [/\bCOCO SECO UN\b/g, "COCO SECO"],
  [/\bCEREJA(?:\s+BANDEJA)?\s+250G\b/g, "CEREJA"],
  [/\bATEMOIA(?:\s+BANDEJA)?(?:\s+(?:\d+\s*)?KG)?\b/g, "ATEMOIA"],
  [/\bFIGO ROXO(?:\s+BANDEJA)?\s+300G\b/g, "FIGO"],
  [/\bFRAMBOESA(?:\s+BANDEJA)?\s+100G\b/g, "FRAMBOESA"],
  [/\bGOIABA GRANEL\b/g, "GOIABA"],
  [/\bJAMBO ROSA(?:\s+BANDEJA)?\s+300G\b/g, "JAMBO"],
  [/\bKIWI KG\b/g, "KIWI"],
  [/\bLARANJA LIMA(?: DA)? PERSIA\b/g, "LIMA DA PERSIA"],
  [/\bLARANJA SELETA\b/g, "LARANJA SELETA"],
  [/\bLIMAO THAITI\b/g, "LIMAO"],
  [/\bMACA RED IMPORT\b/g, "MACA RED"],
  [/\bMACA VERDE GRAN\b/g, "MACA VERDE"],
  [/\bMACA GALA 850G\b/g, "MACA 850G"],
  [/\bMACA BENNI\b/g, "MACA 850G"],
  [/\bMAMAO PAPAYA\b/g, "MAMAO HAVAI"],
  [/\bMELAO CANT\b/g, "MELAO CANTALOUPE"],
  [/\bMELAO REI REDE\b/g, "MELAO REI"],
  [/\bMELAO AMARELO REDE\b/g, "MELAO REI"],
  [/\bMELAO REDE\b/g, "MELAO REI"],
  [/\bMILHO VERDE 3\s+500G\b/g, "MILHO BDJ"],
  [/\bMILHO VERDE BANDEJA\b/g, "MILHO BDJ"],
  [/\bMILHO VERDE BDJ 3\b/g, "MILHO BDJ"],
  [/^MILHO VERDE$/g, "MILHO ESPIGA"],
  [/\bMIRTILLO(?:\s+BANDEJA)?(?:\s+\d+G)?(?:\s+<<<\s+REVISAR\s+>>>)?/g, "MIRTILO"],
  [/\bOVO BRANCO DZ\b/g, "OVOS BRANCOS DZ"],
  [/\bOVOS BRANCO C 20\b/g, "OVOS BRANCOS C 20"],
  [/\bOVOS BRANCO 30\b/g, "OVOS BRANCOS C 30"],
  [/\bOVOS BRANCOS DZ\b/g, "OVOS BRANCOS DZ"],
  [/\bOVOS BRANCOS C 20\b/g, "OVOS BRANCOS C 20"],
  [/\bOVOS BRANCOS 20\b/g, "OVOS BRANCOS C 20"],
  [/\bOVOS BRANCOS 30\b/g, "OVOS BRANCOS C 30"],
  [/\bOVOS CODORNA C 30\b/g, "OVOS CODORNA"],
  [/\bOVOS VERMELHOS C 12\b/g, "OVOS VERMELHO C 12"],
  [/\bPERA D ANJOUR\b/g, "PERA DANJOUR"],
  [/\bPERA DANJOUR\b/g, "PERA DANJOUR"],
  [/\bPERA WILHIA[MN][SN]?\b/g, "PERA WILLIAMS"],
  [/\bPERA WIL?LIANS\b/g, "PERA WILLIAMS"],
  [/\bPEPINO COMUM\b/g, "PEPINO"],
  [/\bPITAYA\b(?:\s+BANDEJA\b)?(?:\s+(?:\d+\s*)?(?:G|KG))?(?:\s+<<<\s+REVISAR\s+>>>)?/g, "PITAYA"],
  [/\bPIMENTAO VERDE(?:\s+BANDEJA)?(?:\s+\d+G)?\b/g, "PIMENTAO"],
  [/\bPIMENTAO AMARELO(?:\s+BANDEJA)?(?:\s+\d+G)?\b/g, "PIMENTAO AMARELO"],
  [/\bPIMENTAO VERMEL(?:HO)?(?:\s+BANDEJA)?(?:\s+\d+G)?\b/g, "PIMENTAO VERMELHO"],
  [/\bPIMENTAO BRANCO(?:\s+BANDEJA)?(?:\s+\d+G)?\b/g, "PIMENTAO BRANCO"],
  [/\bREPOLHO VERDE\b/g, "REPOLHO"],
  [/\bCEBOLA MIUDA(?:\s+(?:PACOTE|PCT))?(?:\s+\d+\s*(?:KG|UNID?|UNIDADE|UNIDADES))?\b/g, "CEBOLA CALABRESA"],
  [/\bCEBOLA CALABRESA(?:\s+(?:PACOTE|PCT))?(?:\s+\d+\s*(?:KG|UNID?|UNIDADE|UNIDADES))?\b/g, "CEBOLA CALABRESA"],
  [/\bCEBOLA PIRULITO(?:\s+(?:PACOTE|PCT))?(?:\s+\d+\s*(?:KG|UNID?|UNIDADE|UNIDADES))?\b/g, "CEBOLA PIRULITO"],
  [/\bCEBOLA PIRUTLIO\b/g, "CEBOLA PIRULITO"],
  [/\bCOGUMELO PARIS(?:\s+(?:BANDEJA|INTEIRO))*(?:\s+\d+G)?\b/g, "COGUMELO PARIS"],
  [/\bCOGUMELO PORTO\s*BELLO(?:\s+(?:BANDEJA|INTEIRO))*(?:\s+\d+G)?\b/g, "COGUMELO PORTOBELLO"],
  [/\bCOGUMELO PORTOBELLO(?:\s+(?:BANDEJA|INTEIRO))*(?:\s+\d+G)?\b/g, "COGUMELO PORTOBELLO"],
  [/\bCOGUMELO SHIME(?:J|G)I(?:\s+(?:BANDEJA|BRANCO|PRETO))*(?:\s+\d+G)?\b/g, "COGUMELO SHIMEJI"],
  [/\bCOGUMELO SHITAKE(?:\s+(?:BANDEJA|INTEIRO))*(?:\s+\d+G)?\b/g, "COGUMELO SHITAKE"],
  [/\bERVILHA(?:\s+BANDEJA)?\s+200G\b/g, "ERVILHA"],
  [/\bQUIABO 300G\b/g, "QUIABO BDJ"],
  [/\bQUIABO BANDEJA 300G\b/g, "QUIABO BDJ"],
  [/\bQUIABO EMBALADOS\b/g, "QUIABO BDJ"],
  [/\bSAPOTI(?:\s+BANDEJA)?\s+300G\b/g, "SAPOTI"],
  [/\bSERIGUELA(?:\s+BANDEJA)?\s+250G\s+<<<\s+REVISAR\s+>>>/g, "SERIGUELA"],
  [/\bSERIGUELA\s+600G\s+<<<\s+REVISAR\s+>>>/g, "SERIGUELA"],
  [/\bTAMARINDO(?:\s+BANDEJA)?\s+300G\b/g, "TAMARINDO"],
  [/\bTANGERINA IMP(?:ORT)?\b/g, "TANGERINA IMPORTADA"],
  [/\bTOMATE GRAPE AMARELO(?:\s+BANDEJA)?\s+250G(?:\s*<<<\s*REVISAR\s*>>>)?/g, "TOMATE GRAPE AMARELO 250G"],
  [/\bTOMATE GRAPE MISTO(?:\s+BANDEJA)?\s+250G(?:\s*<<<\s*REVISAR\s*>>>)?/g, "TOMATE GRAPE MISTO 250G"],
  [/\bTOMATINHO DO BENI\s+250G\b/g, "TOMATE SWEET"],
  [/\bTOMATE SWEET 180\b/g, "TOMATE SWEET"],
  [/\bUVA ITALIA\b/g, "UVA ITALIA"],
  [/\bVAGEM MANTEIGA\b/g, "VAGEM MANT"],
];

const ALWAYS_SUPPLIER_PRODUCTS = [
  {
    fornecedor: "adonai",
    produtos: new Set(["CEBOLA ROXA"]),
  },
  {
    fornecedor: "adonai",
    produtos: new Set(["BATATA ASTERIX", "BATATA BOLINHA"]),
    lojas: new Set(["ANCHIETA"]),
  },
  {
    fornecedor: "adonai",
    produtos: new Set(["CEBOLA CALABRESA", "CEBOLA PIRULITO"]),
    lojas: new Set(["ANCHIETA"]),
  },
  {
    fornecedor: "adonai",
    produtos: new Set(["CEBOLA CALABRESA", "CEBOLA PIRULITO"]),
    lojas: new Set(["FREGUESIA", "PIABETA"]),
  },
  {
    fornecedor: "adonai",
    produtos: new Set(["BATATA BOLINHA"]),
    lojas: new Set(["PIABETA"]),
  },
  {
    fornecedor: "adonai",
    produtos: new Set(["BATATA BOLINHA", "CEBOLA CALABRESA", "CEBOLA PIRULITO"]),
    lojas: new Set(["OLINDA"]),
  },
  {
    fornecedor: "adonai",
    produtos: new Set(["CEBOLA", "CEBOLA CALABRESA", "CEBOLA PIRULITO"]),
    lojas: new Set(["IRAJA"]),
  },
  {
    fornecedor: "adonai",
    produtos: new Set(["BATATA BOLINHA", "BATATA INGLESA", "CEBOLA"]),
    lojas: new Set(["CACHAMBI", "FREGUESIA", "SANTOS"]),
  },
  {
    fornecedor: "adonai",
    produtos: new Set(["BATATA SUJA"]),
    lojas: new Set(["IRAJA"]),
  },
  {
    fornecedor: "adonai",
    produtos: new Set(["BATATA ASTERIX", "BATATA BOLINHA", "CEBOLA CALABRESA", "CEBOLA PIRULITO"]),
    lojas: new Set(["SANTA CRUZ"]),
  },
  {
    fornecedor: "Delorenze",
    produtos: new Set(["BATATA INGLESA", "CEBOLA"]),
    lojas: new Set(["SANTA CRUZ"]),
  },
  {
    fornecedor: "Agrocomercial",
    produtos: new Set(["BATATA INGLESA", "CEBOLA"]),
    lojas: new Set(["OLINDA"]),
  },
  {
    fornecedor: "alevan",
    produtos: new Set(["PINHA"]),
  },
  {
    fornecedor: "Baixinho",
    produtos: new Set(["MILHO BDJ", "MILHO ESPIGA", "QUIABO BDJ"]),
  },
  {
    fornecedor: "BENASSI",
    produtos: new Set([
      "COCO VERDE",
      "COGUMELO",
      "COGUMELO PARIS",
      "COGUMELO PORTOBELLO",
      "COGUMELO SHIMEJI",
      "COGUMELO SHITAKE",
      "MACA 850G",
      "MACA VERDE",
      "MELANCIA PINGO AM",
      "MELANCIA PINGO VER",
      "MELAO REI",
      "UVA BRASIL",
      "UVA CRIMSON",
    ]),
  },
  {
    fornecedor: "Casa dina",
    produtos: new Set([
      "ABOBORA BAIANA",
      "ABOBORA JAPONESA",
      "ABOBORA MORANGA",
      "ABOBORA PESCOCO",
      "ABOBORA SERGIPANA",
      "GENGIBRE",
      "MELANCIA",
    ]),
  },
  {
    fornecedor: "FAISÃO",
    produtos: new Set([
      "AMEIXA",
      "AMORA",
      "CAJA",
      "CARAMBOLA",
      "CEREJA",
      "FIGO",
      "FRAMBOESA",
      "GOIABA",
      "JAMBO",
      "LIMAO SICILIANO",
      "MACA RED",
      "MELANCIA BABY",
      "MELAO GALIA",
      "MELAO ORANGE",
      "MELAO VERDE",
      "MIRTILO",
      "PERA DANJOUR",
      "PITAYA",
      "ROMA",
      "SAPOTI",
      "SERIGUELA",
      "TAMARINDO",
      "TANGERINA IMPORTADA",
      "UVA ITALIA",
    ]),
  },
  {
    fornecedor: "FAISÃO",
    produtos: new Set(["GRAVIOLA"]),
    lojas: new Set(["SANTOS"]),
  },
  {
    fornecedor: "FAISÃO",
    produtos: new Set(["LIMA DA PERSIA"]),
  },
  {
    fornecedor: "Cia dos ovos",
    produtos: new Set([
      "OVO CODORNA",
      "OVO VERM C 12",
      "OVO VERM C 20",
      "OVO VERM C 30",
      "OVOS C 12",
      "OVOS C 20",
      "OVOS C 30",
      "OVOS BRANCOS C 20",
      "OVOS BRANCOS C 30",
      "OVOS BRANCOS DZ",
      "OVOS CODORNA",
      "OVOS VERMELHO C 12",
      "OVOS VERMELHOS C 20",
      "OVOS VERMELHOS C 30",
    ]),
  },
  {
    fornecedor: "BENASSI",
    produtos: new Set(["BETERRABA"]),
    lojas: new Set(["SANTA CRUZ"]),
  },
  {
    fornecedor: "BENASSI",
    produtos: new Set(["VAGEM MANT"]),
    lojas: new Set(["SANTA CRUZ"]),
  },
  {
    fornecedor: "BENASSI",
    produtos: new Set(["VAGEM MACARRAO"]),
    lojas: new Set(["IRAJA"]),
  },
  {
    fornecedor: "BENASSI",
    produtos: new Set(["REPOLHO ROXO"]),
    lojas: new Set(["COELHO"]),
  },
  {
    fornecedor: "CRT",
    produtos: new Set(["BETERRABA"]),
    lojas: new Set(["ANCHIETA"]),
  },
  {
    fornecedor: "CRT",
    produtos: new Set(["BETERRABA"]),
    lojas: new Set(["CACHAMBI"]),
  },
  {
    fornecedor: "CRT",
    produtos: new Set(["MAXIXE"]),
    lojas: new Set(["CACHAMBI", "SANTOS"]),
  },
  {
    fornecedor: "CRT",
    produtos: new Set(["CARA"]),
    lojas: new Set(["CACHAMBI", "FREGUESIA", "IRAJA"]),
  },
  {
    fornecedor: "CRT",
    produtos: new Set(["BERINJELA"]),
    lojas: new Set(["ANCHIETA", "IRAJA"]),
  },
  {
    fornecedor: "CRT",
    produtos: new Set(["BATATA DOCE"]),
    lojas: new Set(["ANCHIETA", "FREGUESIA", "SANTOS"]),
  },
  {
    fornecedor: "CRT",
    produtos: new Set(["PEPINO"]),
    lojas: new Set(["ANCHIETA", "SANTA CRUZ"]),
  },
  {
    fornecedor: "CRT",
    produtos: new Set(["PEPINO JAPONES"]),
    lojas: new Set(["IRAJA"]),
  },
  {
    fornecedor: "CRT",
    produtos: new Set(["PEPINO", "PEPINO JAPONES", "VAGEM MANT"]),
    lojas: new Set(["CACHAMBI", "SANTOS", "FREGUESIA", "IRAJA", "OLINDA"]),
  },
  {
    fornecedor: "CRT",
    produtos: new Set(["TOMATE"]),
    lojas: new Set(["ANCHIETA", "CACHAMBI", "FREGUESIA", "IRAJA", "OLINDA", "QUEIMADOS", "SANTOS"]),
  },
  {
    fornecedor: "CRT",
    produtos: new Set(["REPOLHO", "REPOLHO ROXO"]),
    lojas: new Set(["OLINDA"]),
  },
  {
    fornecedor: "CRT",
    produtos: new Set([
      "PIMENTA ARDIDA BANDEJA 100G",
      "PIMENTA BIQUINHO 100G",
      "PIMENTA BIQUINHO VERMELHA BANDEJA 100G",
      "PIMENTA CAMBUCI BANDEJA 100G",
      "PIMENTA CAMBUCI BANDEJA 250G",
      "PIMENTA CHEIRO BANDEJA 250G",
      "PIMENTA DE CHEIRO DOCE BANDEJA 100G",
      "PIMENTA DEDO MOCA BANDEJA 100G",
      "PIMENTA DEDO MOCA BANDEJA 200G",
      "PIMENTA MALAGUETA BANDEJA 150G",
    ]),
    lojas: new Set(["COELHO"]),
  },
  {
    fornecedor: "CRT",
    produtos: new Set(["ABOBRINHA", "BATATA DOCE", "BERINJELA", "BETERRABA", "CENOURA", "CHUCHU", "INHAME", "PEPINO"]),
    lojas: new Set(["QUEIMADOS"]),
  },
  {
    fornecedor: "JACUBA",
    produtos: new Set(["LIMAO"]),
  },
  {
    fornecedor: "Kifrut",
    produtos: new Set(["ATEMOIA", "CAJU", "UVA RED GLOB", "UVA ROSADA"]),
  },
  {
    fornecedor: "LTB",
    produtos: new Set(["BATATA ASTERIX"]),
    lojas: new Set(["CERAMICA", "COELHO"]),
  },
  {
    fornecedor: "LTB",
    produtos: new Set(["BATATA BOLINHA"]),
    lojas: new Set(["CERAMICA"]),
  },
  {
    fornecedor: "LTB",
    produtos: new Set(["CEBOLA CALABRESA", "CEBOLA PIRULITO"]),
    lojas: new Set(["CERAMICA", "COELHO"]),
  },
  {
    fornecedor: "MIBA",
    produtos: new Set(["ABACATE", "COCO SECO"]),
  },
  {
    fornecedor: "MIBA",
    produtos: new Set(["BANANA PRATA"]),
    lojas: new Set(["IRAJA"]),
  },
  {
    fornecedor: "Milanes",
    produtos: new Set(["LARANJA LIMA", "LARANJA SELETA", "TANGERINA MORCOTE"]),
  },
  {
    fornecedor: "Laranja pera",
    produtos: new Set(["LARANJA PERA"]),
  },
  {
    fornecedor: "NIPPO",
    produtos: new Set(["ABACAXI", "MELAO AMARELO", "UVA THOMPSON", "UVA VITORIA"]),
  },
  {
    fornecedor: "Rio Minas",
    produtos: new Set(["BATATA BAROA"]),
  },
  {
    fornecedor: "ROSSI",
    produtos: new Set(["MAMAO FORMOSA", "MAMAO HAVAI"]),
  },
  {
    fornecedor: "Real",
    produtos: new Set(["BANANA DA TERRA", "BANANA DAGUA", "BANANA MACA", "BANANA OURO", "BANANA PRATA"]),
    lojas: new Set(["CACHAMBI", "SANTOS", "FREGUESIA"]),
  },
  {
    fornecedor: "SEAL",
    produtos: new Set(["PIMENTAO", "PIMENTAO AMARELO", "PIMENTAO VERMELHO", "TOMATE COQUETEL", "TOMATE GRAPE AMARELO 250G", "TOMATE GRAPE MISTO 250G", "TOMATE ITALIANO", "TOMATE SWEET"]),
  },
  {
    fornecedor: "SEAL",
    produtos: new Set(["TOMATE"]),
    lojas: new Set(["CERAMICA", "COELHO", "PIABETA", "SANTA CRUZ"]),
  },
  {
    fornecedor: "uvale",
    produtos: new Set(["BANANA DA TERRA", "BANANA DAGUA", "BANANA MACA", "BANANA OURO", "BANANA PRATA"]),
    lojas: new Set(["ANCHIETA", "CERAMICA", "COELHO", "OLINDA", "PIABETA", "QUEIMADOS", "SANTA CRUZ"]),
  },
  {
    fornecedor: "uvale",
    produtos: new Set(["MANGA PALMER", "MANGA TOMMY"]),
  },
  {
    fornecedor: "Vitoria",
    produtos: new Set(["KIWI", "LARANJA BAHIA", "MACA FUJI", "MACA GALA", "PERA PORTUGUESA", "PERA WILLIAMS"]),
  },
  {
    fornecedor: "Rio Minas",
    produtos: new Set(["MORANGO"]),
  },
  {
    fornecedor: "BENASSI",
    produtos: new Set([
      "AIPIM",
      "BATATA DOCE",
      "BERINJELA",
      "BETERRABA",
      "CENOURA",
      "CHUCHU",
      "INHAME",
      "MAXIXE",
      "PEPINO",
      "PEPINO JAPONES",
      "VAGEM MANT",
    ]),
    lojas: new Set(["CERAMICA", "COELHO", "PIABETA"]),
  },
  {
    fornecedor: "BENASSI",
    produtos: new Set(["BATATA DOCE", "BERINJELA", "BETERRABA", "CENOURA", "CHUCHU", "INHAME", "MAXIXE", "VAGEM MANT"]),
    lojas: new Set(["SANTA CRUZ"]),
  },
  {
    fornecedor: "brasnica",
    produtos: new Set(["BANANA PRATA ORGANICO"]),
    lojas: new Set(["IRAJA", "SANTOS"]),
  },
  {
    fornecedor: "brasnica",
    produtos: new Set(["TANGERINA PONKAN"]),
    lojas: new Set(["ANCHIETA", "CACHAMBI", "CERAMICA", "COELHO", "FREGUESIA", "IRAJA", "OLINDA", "PIABETA", "QUEIMADOS", "SANTA CRUZ", "SANTOS"]),
  },
  {
    fornecedor: "Veneza",
    produtos: new Set(["AIPIM"]),
    lojas: new Set(["QUEIMADOS"]),
  },
  {
    fornecedor: "Veneza",
    produtos: new Set(["ABOBRINHA"]),
    lojas: new Set(["CERAMICA", "COELHO", "PIABETA", "SANTA CRUZ"]),
  },
  {
    fornecedor: "Veneza",
    produtos: new Set(["JILO", "MARACUJA"]),
  },
  {
    fornecedor: "CRT",
    produtos: new Set(["ERVILHA"]),
  },
];

const FORCED_PENDING_ASSOCIATIONS = new Set([]);

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

  if (/^TOMATE GRAPE AMARELO(?:\s+BANDEJA)?\s+250G\b/.test(normalized)) {
    return "TOMATE GRAPE AMARELO 250G";
  }
  if (/^TOMATE GRAPE MISTO(?:\s+BANDEJA)?\s+250G\b/.test(normalized)) {
    return "TOMATE GRAPE MISTO 250G";
  }

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

function formatDate(value) {
  const year = value.getFullYear();
  const month = String(value.getMonth() + 1).padStart(2, "0");
  const day = String(value.getDate()).padStart(2, "0");
  return `${day}/${month}/${year}`;
}

function addDays(value, days) {
  const date = new Date(value);
  date.setDate(date.getDate() + days);
  return date;
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

function addStoreColumn(worksheet, header, store) {
  const insertAt = header.totalColumn || (header.storeColumns.at(-1)?.columnNumber ?? worksheet.columnCount) + 1;
  const sourceColumn = Math.max(insertAt - 1, 1);

  worksheet.spliceColumns(insertAt, 0, []);
  const source = worksheet.getColumn(sourceColumn);
  const target = worksheet.getColumn(insertAt);
  target.width = source.width;
  target.hidden = source.hidden;
  target.outlineLevel = source.outlineLevel || 0;

  for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const sourceCell = worksheet.getRow(rowNumber).getCell(sourceColumn);
    const targetCell = worksheet.getRow(rowNumber).getCell(insertAt);
    targetCell.style = cloneStyle(sourceCell.style);
  }

  worksheet.getRow(header.rowNumber).getCell(insertAt).value = STORE_HEADER_LABELS.get(store) || store;
  header.storeColumns = header.storeColumns.map(({ store: currentStore, columnNumber }) => ({
    store: currentStore,
    columnNumber: columnNumber >= insertAt ? columnNumber + 1 : columnNumber,
  }));
  header.storeColumns.push({ store, columnNumber: insertAt });
  header.storeColumns.sort((a, b) => a.columnNumber - b.columnNumber);

  if (header.totalColumn && header.totalColumn >= insertAt) {
    header.totalColumn += 1;
  }
}

function ensureStoreColumnsForAlwaysSupplierProducts(worksheet, header, quantities, fileName) {
  const existingStores = new Set(header.storeColumns.map(({ store }) => store));
  const missingStores = new Set();

  for (const rule of getAlwaysSupplierRules(fileName)) {
    for (const product of rule.produtos) {
      for (const store of STORE_ORDER) {
        if (existingStores.has(store) || (rule.lojas && !rule.lojas.has(store))) {
          continue;
        }

        if (quantities.has(lookupKey(product, store))) {
          missingStores.add(store);
        }
      }
    }
  }

  for (const store of STORE_ORDER) {
    if (!missingStores.has(store)) {
      continue;
    }

    addStoreColumn(worksheet, header, store);
    existingStores.add(store);
  }
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

function isForcedPendingAssociation(product, store) {
  return FORCED_PENDING_ASSOCIATIONS.has(lookupKey(product, store));
}

function isAlwaysSupplierProduct(fileName, product, store) {
  const fornecedor = normalizeText(supplierNameFromFile(fileName));
  const produto = normalizeProduct(product);

  return ALWAYS_SUPPLIER_PRODUCTS.some((rule) => {
    if (normalizeText(rule.fornecedor) !== fornecedor || !rule.produtos.has(produto)) {
      return false;
    }
    return !rule.lojas || rule.lojas.has(store);
  });
}

function isAssignedToOtherSupplier(fileName, product, store) {
  const fornecedor = normalizeText(supplierNameFromFile(fileName));
  const produto = normalizeProduct(product);

  return ALWAYS_SUPPLIER_PRODUCTS.some((rule) => {
    if (normalizeText(rule.fornecedor) === fornecedor || !rule.produtos.has(produto)) {
      return false;
    }
    return !rule.lojas || rule.lojas.has(store);
  });
}

function getAlwaysSupplierRules(fileName) {
  const fornecedor = normalizeText(supplierNameFromFile(fileName));
  return ALWAYS_SUPPLIER_PRODUCTS.filter((rule) => normalizeText(rule.fornecedor) === fornecedor);
}

function findLastProductRow(worksheet, header) {
  for (let rowNumber = worksheet.rowCount; rowNumber > header.rowNumber; rowNumber -= 1) {
    const product = worksheetValueToString(worksheet.getRow(rowNumber).getCell(1).value).trim();
    if (product && !isTitleCell(product)) {
      return rowNumber;
    }
  }
  return header.rowNumber + 1;
}

function cloneStyle(style) {
  return JSON.parse(JSON.stringify(style ?? {}));
}

function addSupplierProductRow(worksheet, header, product) {
  const sourceRowNumber = findLastProductRow(worksheet, header);
  const targetRowNumber = worksheet.rowCount + 1;
  const sourceRow = worksheet.getRow(sourceRowNumber);
  const targetRow = worksheet.getRow(targetRowNumber);

  targetRow.height = sourceRow.height;

  for (let columnNumber = 1; columnNumber <= worksheet.columnCount; columnNumber += 1) {
    const sourceCell = sourceRow.getCell(columnNumber);
    const targetCell = targetRow.getCell(columnNumber);
    targetCell.style = cloneStyle(sourceCell.style);
    targetCell.value = null;
  }

  targetRow.getCell(1).value = product;
  return targetRowNumber;
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

function updateWorksheetDates(worksheet, formattedDate) {
  const datePattern = /\b\d{2}\/\d{2}\/\d{4}\b/;
  const allDatesPattern = /\b\d{2}\/\d{2}\/\d{4}\b/g;

  worksheet.eachRow((row) => {
    row.eachCell((cell) => {
      const text = worksheetValueToString(cell.value);
      if (!datePattern.test(text)) {
        return;
      }

      cell.value = text.replace(allDatesPattern, formattedDate);
    });
  });
}

function updateWorksheetSupplierName(worksheet, sourceName, supplierName) {
  const sourcePattern = new RegExp(sourceName, "gi");
  worksheet.eachRow((row) => {
    row.eachCell((cell) => {
      if (typeof cell.value === "string") {
        cell.value = cell.value.replace(sourcePattern, supplierName);
      }
    });
  });
}

function sortProductRowsAlphabetically(worksheet, header) {
  const productRows = [];

  for (let rowNumber = header.rowNumber + 1; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    const product = worksheetValueToString(row.getCell(1).value).trim();

    if (isTitleCell(product)) {
      continue;
    }

    const values = [];
    for (let columnNumber = 1; columnNumber <= worksheet.columnCount; columnNumber += 1) {
      values[columnNumber] = row.getCell(columnNumber).value;
    }

    productRows.push({
      product,
      rowNumber,
      values,
    });
  }

  const sortedRows = [...productRows].sort((a, b) => {
    return a.product.localeCompare(b.product, "pt-BR");
  });

  productRows.forEach((target, index) => {
    const source = sortedRows[index];
    const row = worksheet.getRow(target.rowNumber);

    for (let columnNumber = 1; columnNumber <= worksheet.columnCount; columnNumber += 1) {
      row.getCell(columnNumber).value = source.values[columnNumber] ?? null;
    }
  });
}

function moveTotalRowToEnd(worksheet, header) {
  let totalRowNumber = null;

  for (let rowNumber = header.rowNumber + 1; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const label = normalizeText(worksheetValueToString(worksheet.getRow(rowNumber).getCell(1).value));
    if (label === "TOTAL") {
      totalRowNumber = rowNumber;
      break;
    }
  }

  if (!totalRowNumber) {
    return;
  }

  const sourceRow = worksheet.getRow(totalRowNumber);
  const sourceHeight = sourceRow.height;
  const sourceCells = [];
  for (let columnNumber = 1; columnNumber <= worksheet.columnCount; columnNumber += 1) {
    const cell = sourceRow.getCell(columnNumber);
    sourceCells[columnNumber] = {
      style: cloneStyle(cell.style),
      value: cell.value,
    };
  }

  if (totalRowNumber !== worksheet.rowCount) {
    worksheet.spliceRows(totalRowNumber, 1);
    totalRowNumber = worksheet.rowCount + 1;
  }

  const totalRow = worksheet.getRow(totalRowNumber);
  totalRow.height = sourceHeight;
  for (let columnNumber = 1; columnNumber <= worksheet.columnCount; columnNumber += 1) {
    totalRow.getCell(columnNumber).style = sourceCells[columnNumber].style;
    totalRow.getCell(columnNumber).value = sourceCells[columnNumber].value;
  }

  totalRow.getCell(1).value = "TOTAL";
  const firstProductRow = header.rowNumber + 1;
  const lastProductRow = totalRowNumber - 1;
  const columnsToSum = [...header.storeColumns.map(({ columnNumber }) => columnNumber), header.totalColumn].filter(Boolean);

  for (const columnNumber of columnsToSum) {
    const columnName = columnNumberToName(columnNumber);
    totalRow.getCell(columnNumber).value = {
      formula: `SUM(${columnName}${firstProductRow}:${columnName}${lastProductRow})`,
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

  sortProductRowsAlphabetically(worksheet, header);
  moveTotalRowToEnd(worksheet, header);
  updateTotalFormulas(worksheet, header);
}

function hasSupplierOrders(worksheet, header) {
  for (let rowNumber = header.rowNumber + 1; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    const product = worksheetValueToString(row.getCell(1).value).trim();

    if (!product || isTitleCell(product)) {
      continue;
    }

    if (rowHasStoreQuantity(row, header.storeColumns)) {
      return true;
    }
  }

  return false;
}

function revealEntireWorkbook(workbook) {
  workbook.worksheets.forEach((worksheet) => {
    worksheet.state = "visible";

    for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber += 1) {
      worksheet.getRow(rowNumber).hidden = false;
    }

    for (let columnNumber = 1; columnNumber <= worksheet.columnCount; columnNumber += 1) {
      worksheet.getColumn(columnNumber).hidden = false;
    }
  });
}

function clearAndFillSupplierWorksheet(worksheet, quantities, fileName, consumedKeys) {
  const header = findHeaderRow(worksheet);
  ensureStoreColumnsForAlwaysSupplierProducts(worksheet, header, quantities, fileName);
  const usesOnlyExplicitSupplierRules = new Set(["DELORENZE", "LARANJA PERA"])
    .has(normalizeText(supplierNameFromFile(fileName)));
  const mappedCells = [];
  const existingProducts = new Set();
  const existingProductRows = new Map();

  for (let rowNumber = header.rowNumber + 1; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    const product = worksheetValueToString(row.getCell(1).value).trim();

    if (isTitleCell(product)) {
      continue;
    }

    const normalizedProduct = normalizeProduct(product);
    if (normalizedProduct === "OVOS BRANCOS DZ" && product !== "OVOS BRANCOS DZ") {
      row.getCell(1).value = "OVOS BRANCOS DZ";
    }
    existingProducts.add(normalizedProduct);
    if (!existingProductRows.has(normalizedProduct)) {
      existingProductRows.set(normalizedProduct, rowNumber);
    }

    for (const { store, columnNumber } of header.storeColumns) {
      const cell = row.getCell(columnNumber);
      if (!isForcedPendingAssociation(product, store)
        && !isAssignedToOtherSupplier(fileName, product, store)
        && (isAlwaysSupplierProduct(fileName, product, store)
          || (!usesOnlyExplicitSupplierRules && isFilledQuantity(cell.value)))) {
        mappedCells.push({ rowNumber, columnNumber, product, store });
      }
      cell.value = null;
    }
  }

  for (const rule of getAlwaysSupplierRules(fileName)) {
    for (const product of rule.produtos) {
      const storesWithQuantity = header.storeColumns.filter(({ store }) => {
        return !isForcedPendingAssociation(product, store)
          && (!rule.lojas || rule.lojas.has(store))
          && quantities.has(lookupKey(product, store));
      });

      if (!storesWithQuantity.length) {
        continue;
      }

      let rowNumber = existingProductRows.get(product);
      if (!rowNumber) {
        rowNumber = addSupplierProductRow(worksheet, header, product);
        existingProducts.add(product);
        existingProductRows.set(product, rowNumber);
      }

      for (const { store, columnNumber } of storesWithQuantity) {
        mappedCells.push({ rowNumber, columnNumber, product, store });
      }
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
  return hasSupplierOrders(worksheet, header);
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
  const formattedDate = formatDate(addDays(options.now ?? new Date(), 1));

  if (!fs.existsSync(templateDir)) {
    throw new Error("Pasta exemplo nao encontrada.");
  }

  const { quantities, entries } = await loadMapQuantities(mapDir);
  fs.rmSync(outputDir, { recursive: true, force: true });
  fs.mkdirSync(outputDir, { recursive: true });
  const consumedKeys = new Set();

  const supplierFiles = fs
    .readdirSync(templateDir)
    .filter((fileName) => !/^~\$/.test(fileName) && /\.xlsx$/i.test(fileName))
    .sort((a, b) => a.localeCompare(b, "pt-BR"));

  if (!supplierFiles.some((fileName) => /^Delorenze\.xlsx$/i.test(fileName))) {
    const adonaiTemplate = supplierFiles.find((fileName) => /^adonai\.xlsx$/i.test(fileName));
    if (!adonaiTemplate) {
      throw new Error("Modelo da Delorenze nao encontrado e modelo da Adonai indisponivel para copia.");
    }
    supplierFiles.push("Delorenze.xlsx");
  }

  if (!supplierFiles.some((fileName) => /^Laranja pera\.xlsx$/i.test(fileName))) {
    const milanesTemplate = supplierFiles.find((fileName) => /^Milanes\.xlsx$/i.test(fileName));
    if (!milanesTemplate) {
      throw new Error("Modelo da Laranja pera nao encontrado e modelo da Milanes indisponivel para copia.");
    }
    supplierFiles.push("Laranja pera.xlsx");
  }

  const generatedFiles = [];

  for (const fileName of supplierFiles) {
    let templateFileName = fileName;
    if (!fs.existsSync(path.join(templateDir, fileName))) {
      if (/^Delorenze\.xlsx$/i.test(fileName)) {
        templateFileName = supplierFiles.find((candidate) => /^adonai\.xlsx$/i.test(candidate));
      } else if (/^Laranja pera\.xlsx$/i.test(fileName)) {
        templateFileName = supplierFiles.find((candidate) => /^Milanes\.xlsx$/i.test(candidate));
      }
    }
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(path.join(templateDir, templateFileName));

    if (/^Delorenze\.xlsx$/i.test(fileName)) {
      workbook.worksheets.forEach((worksheet) => updateWorksheetSupplierName(worksheet, "ADONAI", "DELORENZE"));
    } else if (/^Laranja pera\.xlsx$/i.test(fileName)) {
      workbook.worksheets.forEach((worksheet) => updateWorksheetSupplierName(worksheet, "MILANES", "LARANJA PERA"));
    }

    workbook.worksheets.forEach((worksheet) => updateWorksheetDates(worksheet, formattedDate));

    let hasOrders = false;
    if (workbook.worksheets[0]) {
      hasOrders = clearAndFillSupplierWorksheet(workbook.worksheets[0], quantities, fileName, consumedKeys);
    }

    if (!hasOrders) {
      continue;
    }

    if (normalizeText(supplierNameFromFile(fileName)) === "CRT") {
      revealEntireWorkbook(workbook);
    }

    await workbook.xlsx.writeFile(path.join(outputDir, fileName));
    generatedFiles.push(fileName);
  }

  const pendingEntries = entries.filter((entry) => !consumedKeys.has(entry.key));
  const unmatchedFile = await writeUnmatchedWorkbook(outputDir, pendingEntries);

  return {
    outputDir,
    files: generatedFiles,
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
