const fs = require("fs");
const path = require("path");
const { spawnSync } = require("child_process");

const PROJECT_ROOT = path.resolve(__dirname, "..");
const INPUT_DIR = path.join(PROJECT_ROOT, "input");
const OUTPUT_DIR = path.join(PROJECT_ROOT, "output");
const EXTRACTOR_PATH = path.join(PROJECT_ROOT, "src", "extract-formulario08.js");

function fail(message) {
  console.error(`FAIL: ${message}`);
  process.exit(1);
}

function listOdsFiles() {
  if (!fs.existsSync(OUTPUT_DIR)) return new Map();

  return new Map(
    fs
      .readdirSync(OUTPUT_DIR, { withFileTypes: true })
      .filter(
        (entry) =>
          entry.isFile() &&
          path.extname(entry.name).toLowerCase() === ".ods"
      )
      .map((entry) => {
        const filePath = path.join(OUTPUT_DIR, entry.name);
        return [filePath, fs.statSync(filePath).mtimeMs];
      })
  );
}

if (!fs.existsSync(EXTRACTOR_PATH)) {
  fail(
    "no se encontro src/extract-formulario08.js. Ejecutar desde el proyecto digit actual."
  );
}

if (process.argv.length > 2) {
  fail("el uso simple no acepta argumentos. Ejecutar solamente: extraer");
}

if (!fs.existsSync(INPUT_DIR)) {
  fail("no existe la carpeta input. Crear digit/input y dejar un unico PDF.");
}

const pdfs = fs
  .readdirSync(INPUT_DIR, { withFileTypes: true })
  .filter(
    (entry) =>
      entry.isFile() &&
      path.extname(entry.name).toLowerCase() === ".pdf"
  )
  .map((entry) => path.join(INPUT_DIR, entry.name));

if (pdfs.length === 0) {
  fail("input no tiene ningun PDF. Dejar un unico PDF en input.");
}

if (pdfs.length > 1) {
  console.error("FAIL: input tiene mas de un PDF. Dejar uno solo:");

  for (const pdfPath of pdfs) {
    console.error(`- ${path.basename(pdfPath)}`);
  }

  process.exit(1);
}

const inputFile = pdfs[0];
const beforeOds = listOdsFiles();

console.log(`Input: ${path.relative(PROJECT_ROOT, inputFile)}`);

const result = spawnSync(process.execPath, [EXTRACTOR_PATH], {
  cwd: PROJECT_ROOT,
  stdio: "inherit",
  env: process.env,
});

if (result.error) {
  fail(result.error.message);
}

if (result.status !== 0) {
  fail(`extractor termino con codigo ${result.status}.`);
}

const afterOds = listOdsFiles();

const generatedOds = [...afterOds.entries()]
  .filter(
    ([filePath, mtimeMs]) =>
      !beforeOds.has(filePath) || beforeOds.get(filePath) !== mtimeMs
  )
  .sort((a, b) => b[1] - a[1])
  .map(([filePath]) => filePath);

if (generatedOds.length === 0) {
  fail("el extractor termino sin generar ODS en output.");
}

console.log(
  `PASS: ODS final generado en ${path.relative(
    PROJECT_ROOT,
    generatedOds[0]
  )}`
);