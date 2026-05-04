const fs = require("fs");
const path = require("path");

const ROOT_DIR = path.resolve(__dirname, "..");
const DIST_DIR = path.join(ROOT_DIR, "dist");
const TEMPLATE_SOURCE_DIR = path.join(ROOT_DIR, "template");
const DATA_SOURCE_DIR = path.join(ROOT_DIR, "data");

function ensureCleanDist() {
  if (path.basename(DIST_DIR) !== "dist") {
    throw new Error("Diretorio de build inesperado.");
  }

  fs.mkdirSync(DIST_DIR, { recursive: true });
}

function copyFile(sourceRelativePath, targetRelativePath) {
  const sourcePath = path.join(ROOT_DIR, sourceRelativePath);
  const targetPath = path.join(DIST_DIR, targetRelativePath);

  fs.mkdirSync(path.dirname(targetPath), { recursive: true });
  fs.copyFileSync(sourcePath, targetPath);
}

function buildHtml() {
  const sourcePath = path.join(ROOT_DIR, "index.html");
  const targetPath = path.join(DIST_DIR, "index.html");
  const html = fs.readFileSync(sourcePath, "utf8")
    .replace('href="web/styles.css"', 'href="./styles.css"')
    .replace('<script src="web/app.js"></script>', '<script src="./index.js"></script>');

  fs.writeFileSync(targetPath, html, "utf8");
}

function copyTemplates() {
  fs.cpSync(TEMPLATE_SOURCE_DIR, path.join(DIST_DIR, "template"), { recursive: true });
}

function copyData() {
  fs.cpSync(DATA_SOURCE_DIR, path.join(DIST_DIR, "data"), { recursive: true });
}

function writeNoJekyllFile() {
  fs.writeFileSync(path.join(DIST_DIR, ".nojekyll"), "", "utf8");
}

function main() {
  ensureCleanDist();
  buildHtml();
  copyFile(path.join("web", "app.js"), "index.js");
  copyFile(path.join("web", "styles.css"), "styles.css");
  copyTemplates();
  copyData();
  writeNoJekyllFile();
  console.log("Build do GitHub Pages gerado em dist/");
}

main();
