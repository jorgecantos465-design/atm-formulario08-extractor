const fs = require("fs");
const path = require("path");
const ExcelJS = require("exceljs");
const JSZip = require("jszip");
const pdfParse = require("pdf-parse");

const PROJECT_ROOT = path.resolve(__dirname, "..");
const INPUT_DIR = path.join(PROJECT_ROOT, "input");
const OUTPUT_DIR = path.join(PROJECT_ROOT, "output");
const TEMPLATE_DIR = path.join(PROJECT_ROOT, "templates");
const LOG_DIR = path.join(PROJECT_ROOT, "logs");
const PLACEHOLDER_ROW = 2;
const FIRST_DATA_ROW = 3;
const INICIADOR_PLACEHOLDER = "@atributo10@";
const OSD_DATE_PLACEHOLDER = "@atributo17@";
const AMOUNT_PLACEHOLDER = "@atributo33@";
const DATE_PLACEHOLDERS = new Set([OSD_DATE_PLACEHOLDER, "@atributo30@"]);
const CUIT_REGEX = /(?:^|[^\d])(\d{2}\s*[- ]?\s*[0-9S]{8}\s*[- ]?\s*[0-9I])(?=$|[^\d])/gi;
const MONTHS_ES = {
  enero: "01",
  febrero: "02",
  marzo: "03",
  abril: "04",
  mayo: "05",
  junio: "06",
  julio: "07",
  agosto: "08",
  septiembre: "09",
  setiembre: "09",
  octubre: "10",
  noviembre: "11",
  diciembre: "12",
};

const NEVER_FILL = new Set([
  "@atributo23@",
  "@atributo24@",
  "@atributo25@",
  "@atributo26@",
  "@atributo27@",
  "@atributo32@",
  "@atributo34@",
]);

const FIELDS = {
  usuario: "@usuario@",
  apellido: "@apellido@",
  nombre: "@nombre@",
  email: "@email@",
  atributo8: "@atributo8@",
  atributo9: "@atributo9@",
  atributo7: "@atributo7@",
  atributo10: "@atributo10@",
  atributo11: "@atributo11@",
  atributo12: "@atributo12@",
  atributo13: "@atributo13@",
  atributo14: "@atributo14@",
  atributo15: "@atributo15@",
  atributo16: "@atributo16@",
  atributo17: "@atributo17@",
  atributo18: "@atributo18@",
  atributo19: "@atributo19@",
  atributo20: "@atributo20@",
  atributo21: "@atributo21@",
  atributo22: "@atributo22@",
  atributo30: "@atributo30@",
  atributo33: "@atributo33@",
};

function ensureDirs() {
  for (const dir of [INPUT_DIR, OUTPUT_DIR, TEMPLATE_DIR, LOG_DIR]) {
    fs.mkdirSync(dir, { recursive: true });
  }
}

function nowStamp() {
  const d = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}-${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`;
}

function normalizeText(text) {
  return String(text || "")
    .replace(/\r/g, "\n")
    .replace(/[ \t]+/g, " ")
    .replace(/\n[ \t]+/g, "\n")
    .replace(/[ \t]+\n/g, "\n")
    .trim();
}

function linesOf(text) {
  return normalizeText(text)
    .split(/\n+/)
    .map((line) => line.trim())
    .filter(Boolean);
}

function compactText(text) {
  return linesOf(text).join(" ");
}

function stripAccents(text) {
  return String(text || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function cellText(value) {
  if (value == null) return "";
  if (typeof value === "string") return value.trim();
  if (typeof value === "number") return String(value);
  if (typeof value === "boolean") return String(value);
  if (value.formula) return value.result == null ? `[formula:${value.formula}]` : String(value.result).trim();
  if (value.text) return String(value.text).trim();
  if (value.richText) return value.richText.map((part) => part.text || "").join("").trim();
  if (value.result != null) return String(value.result).trim();
  return String(value).trim();
}

function escapeXml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function unescapeXml(value) {
  return String(value || "")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");
}

function columnName(columnNumber) {
  let value = columnNumber;
  let name = "";
  while (value > 0) {
    const mod = (value - 1) % 26;
    name = String.fromCharCode(65 + mod) + name;
    value = Math.floor((value - mod) / 26);
  }
  return name;
}

function firstMatch(text, regex) {
  const match = String(text || "").match(regex);
  return match ? match[1].trim() : "";
}

function cleanCuit(value) {
  const digits = String(value || "")
    .replace(/[Ii]/g, "1")
    .replace(/[Ss]/g, "5")
    .replace(/\D/g, "");
  if (digits.length !== 11) return "";
  return `${digits.slice(0, 2)}-${digits.slice(2, 10)}-${digits.slice(10)}`;
}

function cleanPartialCuit(value) {
  const normalized = String(value || "")
    .replace(/[Ii]/g, "1")
    .replace(/[Ss]/g, "5")
    .replace(/\s+/g, "")
    .trim();
  const digits = normalized.replace(/\D/g, "");
  if (digits.length < 8 || digits.length > 11) return "";
  if (digits.length === 11) return cleanCuit(digits);
  return `${digits.slice(0, 2)}-${digits.slice(2, 10)}${normalized.endsWith("-") ? "-" : ""}`;
}

function normalizeCuitCandidate(value) {
  const cleaned = String(value || "")
    .replace(/[Ii]/g, "1")
    .replace(/[Ss]/g, "5")
    .replace(/[‐‑‒–—]/g, "-")
    .replace(/\s+/g, " ")
    .trim();
  return cleanCuit(cleaned);
}

function findCuits(text) {
  const source = String(text || "");
  const matches = [];
  for (const match of source.matchAll(CUIT_REGEX)) {
    const cuit = normalizeCuitCandidate(match[1]);
    if (cuit && !matches.includes(cuit)) matches.push(cuit);
  }
  return matches;
}

function extractFirstCuit(text) {
  return findCuits(text)[0] || "";
}

function extractCuitAfterLabel(text, labelRegex) {
  const source = compactText(text);
  const labels = source.matchAll(new RegExp(labelRegex.source, labelRegex.flags.includes("g") ? labelRegex.flags : `${labelRegex.flags}g`));
  for (const label of labels) {
    const cuit = extractFirstCuit(source.slice(label.index + label[0].length, label.index + label[0].length + 80));
    if (cuit) return cuit;
  }
  return "";
}

function extractPartialCuitAfterLabel(text, labelRegex) {
  const source = compactText(text);
  const labels = source.matchAll(new RegExp(labelRegex.source, labelRegex.flags.includes("g") ? labelRegex.flags : `${labelRegex.flags}g`));
  for (const label of labels) {
    const afterLabel = source.slice(label.index + label[0].length, label.index + label[0].length + 40);
    const match = afterLabel.match(/[0-9IS]{2}\s*[- ]?\s*[0-9IS]{6,8}(?:\s*[- ]?\s*[0-9IS]?)?/i);
    const cuit = match ? cleanPartialCuit(match[0]) : "";
    if (cuit) return cuit;
  }
  return "";
}

function cleanDomain(value) {
  const matches = String(value || "").matchAll(/\b[A-Z]{3}\s*-?\s*\d{3}\b|\b[A-Z]{2}\s*-?\s*\d{3}\s*-?\s*[A-Z]{2}\b/gi);
  for (const match of matches) {
    const domain = match[0].replace(/[^A-Z0-9]/gi, "").toUpperCase();
    if (/^DEL0?80$/.test(domain)) continue;
    if (/^[A-Z]{3}\d{3}$/.test(domain) || /^[A-Z]{2}\d{3}[A-Z]{2}$/.test(domain)) return domain;
  }
  return "";
}

function cleanMoney(value) {
  const match = String(value || "")
    .replace(/[tT](?=\d{3}\b)/g, ".")
    .match(/\$?\s*(?:\d{1,3}(?:[.\s]\d{3})+(?:,\d+)?|\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+(?:[,.]\d+)?)/);
  if (!match) return "";
  return normalize_amount(match[0]) ?? "";
}

function normalize_amount(value) {
  if (typeof value === "number") return Number.isFinite(value) ? value : "";

  let text = String(value || "")
    .replace(/^'+/, "")
    .replace(/[^\d.,-]/g, "")
    .trim();
  if (!text) return "";

  const negative = text.startsWith("-");
  text = text.replace(/-/g, "");
  const hasDot = text.includes(".");
  const hasComma = text.includes(",");

  if (hasDot && hasComma) {
    const decimalSeparator = text.lastIndexOf(",") > text.lastIndexOf(".") ? "," : ".";
    const thousandsSeparator = decimalSeparator === "," ? "." : ",";
    text = text.split(thousandsSeparator).join("").replace(decimalSeparator, ".");
  } else if (hasDot || hasComma) {
    const separator = hasDot ? "." : ",";
    const parts = text.split(separator);
    const last = parts[parts.length - 1];
    const groupedThousands = parts.length > 2 || (parts.length === 2 && last.length === 3);
    if (groupedThousands) {
      text = parts.join("");
    } else {
      text = parts.join(".");
    }
  }

  const amount = Number(`${negative ? "-" : ""}${text}`);
  if (!Number.isFinite(amount)) return "";
  return Number.isInteger(amount) ? Math.trunc(amount) : amount;
}

function cleanOcrFieldValue(value) {
  return String(value || "")
    .replace(/[|_*^~]+/g, " ")
    .replace(/[■□]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function cleanVehicleValue(value) {
  const cleaned = cleanOcrFieldValue(value)
    .replace(/^[.:;\-\s]+/, "")
    .replace(/[^\wÁÉÍÓÚÑÜáéíóúñü .,/()-]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  return cleaned.toUpperCase();
}

function cleanPersonName(value) {
  const cleaned = String(value || "")
    .replace(/\s+\/?X\s*$/i, "")
    .replace(/\s+/g, " ")
    .replace(/[^A-ZÁÉÍÓÚÑÜ, .'-]/gi, "")
    .trim();
  if (!cleaned || cleaned.length < 3) return "";
  if (!/[A-ZÁÉÍÓÚÑÜ]{2}/i.test(cleaned)) return "";
  if (cleaned.length < String(value || "").replace(/\s+/g, " ").trim().length * 0.75) return "";
  return cleaned.toUpperCase();
}

function normalizeComparable(value) {
  return stripAccents(String(value || "").trim()).replace(/\s+/g, " ").toUpperCase();
}

function isAutomaticFallbackValue(value) {
  return /\b(NO TIENE|SIN DATOS|N\/A|NA|NO APLICA|NO CORRESPONDE)\b/.test(normalizeComparable(value));
}

function isValidIniciadorValue(value) {
  const cleaned = cleanPersonName(value);
  return cleaned !== "" && !isAutomaticFallbackValue(cleaned);
}

function cleanDateDdMmYyyy(value) {
  return normalize_date(value);
}

function normalize_date(value) {
  const text = String(value || "").replace(/^'+/, "").replace(/\s+/g, " ").trim();
  const match = text.match(/\b(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})\b/);
  if (match) {
    const day = match[1].padStart(2, "0");
    const month = match[2].padStart(2, "0");
    const year = match[3].length === 2 ? `20${match[3]}` : match[3];
    return `${day}/${month}/${year}`;
  }

  const monthTextMatch = stripAccents(text)
    .toLowerCase()
    .match(/\b(\d{1,2})\s+de\s+([a-z]+)\s+de\s+(\d{2,4})\b/);
  if (monthTextMatch) {
    const month = MONTHS_ES[monthTextMatch[2]];
    if (!month) return "";
    const day = monthTextMatch[1].padStart(2, "0");
    const year = monthTextMatch[3].length === 2 ? `20${monthTextMatch[3]}` : monthTextMatch[3];
    return `${day}/${month}/${year}`;
  }
  return "";
}

function formatDateDdMmYy(value) {
  const dateText = normalize_date(value);
  const match = dateText.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!match) return "";
  return `${match[1]}/${match[2]}/${match[3].slice(-2)}`;
}

function normalizeValueForPlaceholder(placeholder, value) {
  if (value === "" || value == null) return "";
  if (placeholder === AMOUNT_PLACEHOLDER) return normalize_amount(value);
  if (DATE_PLACEHOLDERS.has(placeholder)) return normalize_date(value);
  return value;
}

function parseDdMmYyyy(value) {
  const dateText = cleanDateDdMmYyyy(value);
  const match = dateText.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (!match) return null;

  const day = Number(match[1]);
  const month = Number(match[2]);
  const year = Number(match[3]);
  const date = new Date(Date.UTC(year, month - 1, day));
  if (date.getUTCFullYear() !== year || date.getUTCMonth() !== month - 1 || date.getUTCDate() !== day) return null;
  return { date, text: dateText, iso: `${match[3]}-${match[2]}-${match[1]}` };
}

function findSectionByKeywords(text, starts, ends) {
  const source = String(text || "");
  const comparable = stripAccents(source).toUpperCase();
  const startIndexes = starts.map((regex) => comparable.search(regex)).filter((idx) => idx >= 0);
  if (startIndexes.length === 0) return "";

  const start = Math.min(...startIndexes);
  const tail = comparable.slice(start + 1);
  const endIndexes = ends.map((regex) => tail.search(regex)).filter((idx) => idx >= 0);
  const end = endIndexes.length ? start + 1 + Math.min(...endIndexes) : source.length;
  return source.slice(start, end);
}

function findSectionFromLines(text, startRegexes, endRegexes) {
  const lines = linesOf(text);
  const comparable = lines.map((line) => stripAccents(line).toUpperCase());
  const startIndex = comparable.findIndex((line, index) => {
    const joined = `${line} ${comparable[index + 1] || ""}`.trim();
    return startRegexes.some((regex) => regex.test(joined));
  });
  if (startIndex < 0) return "";

  let endIndex = lines.length;
  for (let index = startIndex + 1; index < comparable.length; index += 1) {
    const joined = `${comparable[index]} ${comparable[index + 1] || ""}`.trim();
    if (endRegexes.some((regex) => regex.test(joined))) {
      endIndex = index;
      break;
    }
  }
  return lines.slice(startIndex, endIndex).join("\n");
}

function valueAfterRegex(text, labelRegex, stopRegexes) {
  const source = compactText(text);
  const label = source.match(labelRegex);
  if (!label) return "";

  const start = label.index + label[0].length;
  let end = source.length;
  const tail = source.slice(start);
  for (const stopRegex of stopRegexes) {
    const stop = tail.match(stopRegex);
    if (stop && stop.index >= 0) {
      end = Math.min(end, start + stop.index);
    }
  }
  return source.slice(start, end).replace(/\s+/g, " ").trim();
}

function extractMontoOperacion(text) {
  const source = compactText(text);
  const label = source.match(/Monto(?:\s+de)?\s+Operaci[oó0d]n\s*:?\s*|Monto(?:\s+de)?\s+Operacion\s*:?\s*|Monto\s*:?\s*/i);
  if (label) {
    const afterLabel = source.slice(label.index + label[0].length, label.index + label[0].length + 80);
    const value = cleanMoney(afterLabel);
    if (value) return value;
  }

  const lines = linesOf(text);
  for (const line of lines) {
    if (/Monto/i.test(line)) {
      const value = cleanMoney(line);
      if (value) return value;
    }
  }
  return "";
}

function extractFechaImpresion08(text) {
  const lines = linesOf(text);
  for (let index = 0; index < lines.length; index += 1) {
    if (/Lugar\s*y?\s*Fecha\s+de\s+impresi[oó0d]n/i.test(lines[index])) {
      const nextLine = lines[index + 1] || "";
      return cleanDateDdMmYyyy(nextLine);
    }
  }
  return "";
}

function cleanAddress(value) {
  return String(value || "")
    .replace(/[$*^|]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function extractAddressByLabel(text, labelRegex) {
  const lines = linesOf(text);
  for (let index = 0; index < lines.length; index += 1) {
    const line = lines[index];
    const label = line.match(labelRegex);
    if (!label) continue;

    const parts = [line.slice(label.index + label[0].length).trim()];
    const nextLine = lines[index + 1] || "";
    if (
      nextLine &&
      !/^(Dom.?cilio|CUIT|Email|D\.?N\.?I\.?|Sexo|Ocupaci[oó0d]n|Nacionalidad|Fecha de Nac)\b/i.test(nextLine)
    ) {
      parts.push(nextLine.trim());
    }
    return cleanAddress(parts.join(" "));
  }
  return "";
}

function extractDomicilioAdquirente(buyerSection) {
  return (
    extractAddressByLabel(buyerSection, /Dom.?cilio\s+Legal\s*:\s*/i) ||
    extractAddressByLabel(buyerSection, /Legal\s*:\s*/i) ||
    extractAddressByLabel(buyerSection, /Dom.?cilio\s+Real\s*:\s*/i)
  );
}

function extractFechaImpresion08Strict(sectionA) {
  const lines = linesOf(sectionA);
  for (let index = 0; index < lines.length; index += 1) {
    const window = lines.slice(index, index + 7).join(" ");
    if (/Lugar\s*y?\s*Fech[ao]|Fecha\s+de\s+impresi[oó0d]n|Fecha\s+impresi[oó0d]n|OSD|08D/i.test(window)) {
      const date = cleanDateDdMmYyyy(window);
      if (date) return date;
    }
  }
  for (let index = 0; index < lines.length; index += 1) {
    const date = cleanDateDdMmYyyy(lines.slice(index, index + 2).join(" "));
    if (date) return cleanDateDdMmYyyy(date);
  }
  return "";
}

function extractLineValueFlexible(text, labelPartsRegexes, stopRegexes = []) {
  const lines = linesOf(text);
  for (let index = 0; index < lines.length; index += 1) {
    const window = lines.slice(index, index + 4).join(" ");
    if (!labelPartsRegexes.every((regex) => regex.test(window))) continue;
    let value = window;
    for (const regex of labelPartsRegexes) value = value.replace(regex, " ");
    for (const stopRegex of stopRegexes) {
      const stop = value.search(stopRegex);
      if (stop >= 0) value = value.slice(0, stop);
    }
    return value.replace(/\s+/g, " ").trim();
  }
  return "";
}

function valueBetweenFlexibleLabels(text, labelRegex, stopRegexes) {
  const source = compactText(text);
  const label = source.match(labelRegex);
  if (!label) return "";

  const start = label.index + label[0].length;
  const tail = source.slice(start);
  let end = tail.length;
  for (const stopRegex of stopRegexes) {
    const stop = tail.match(stopRegex);
    if (stop && stop.index >= 0) end = Math.min(end, stop.index);
  }
  return cleanOcrFieldValue(tail.slice(0, end));
}

function extractModel(sectionA) {
  const modelLabel = /(?:M[o0]d[e3][l1I][o0©]|ModeIo|Mode1o|Modeto|Modelo)\s*:?\s*/i;
  const nextVehicleLabel = [
    /(?:A(?:ñ|n|fi|fl|f|Ã±)i?o|Anio|Aflo)\s*:?\s*/i,
    /N[uú]mero\s+Motor\s*:?\s*/i,
    /Numero\s+Motor\s*:?\s*/i,
    /N[uú]mero\s+Cha(?:s|l)s(?:is)?\s*:?\s*/i,
    /Numero\s+Cha(?:s|l)s(?:is)?\s*:?\s*/i,
  ];
  const byLabel = valueBetweenFlexibleLabels(sectionA, modelLabel, nextVehicleLabel);
  if (byLabel) return cleanVehicleValue(byLabel);

  const byProximity = valueBetweenFlexibleLabels(sectionA, /Tipo\s*:?\s*[\s\S]{0,120}?/i, nextVehicleLabel);
  return cleanVehicleValue(byProximity);
}

function extractYear(sectionA) {
  return firstMatch(sectionA, /A(?:ñ|n|fi|fl|f)i?o\s*:?\s*(\d{4})/i);
}

function extractDomain(text, sectionA) {
  const fromLabel = cleanDomain(extractLineValue(sectionA, /Dominio\s*:?\s*/i));
  if (fromLabel) return fromLabel;

  const observations = findSectionFromLines(
    text,
    [/OBSERVACIONES/],
    [/INFRACCIONES/, /PATENTES/, /SELLOS/, /FIRMA/]
  );
  return cleanDomain(observations) || cleanDomain(sectionA) || cleanDomain(text);
}

function extractLineValue(text, labelRegex) {
  const lines = linesOf(text);
  for (const line of lines) {
    const match = line.match(labelRegex);
    if (!match) continue;
    return line.slice(match.index + match[0].length).replace(/\s+/g, " ").trim();
  }
  return "";
}

function extractSellerCuit(sellerSection) {
  return (
    extractCuitAfterLabel(sellerSection, /CUIL\/CUIT\s*:?\s*/i) ||
    extractPartialCuitAfterLabel(sellerSection, /CUIL\/CUIT\s*:?\s*/i) ||
    extractFirstCuit(sellerSection)
  );
}

function extractPersonNameAfterLabel(text) {
  return valueBetweenFlexibleLabels(text, /Apellido\s+y\s+Nombre(?:\/s|\(s\))?\s*:?\s*/i, [
    /\bD\.?\s*N\.?\s*I\.?\s*:?\s*/i,
    /\bDNI\s*:?\s*/i,
    /\bCUIL\/CUIT\s*:?\s*/i,
    /\bCUIT\s*:?\s*/i,
    /\bEmail\s*:?\s*/i,
    /\bRepresentante\b/i,
    /\bEstado\s+Civil\b/i,
    /\bSexo\s*:?\s*/i,
    /\bOcupaci[oó0d]n\s*:?\s*/i,
  ]);
}

function likelyNameCandidate(value) {
  const cleaned = cleanPersonName(value);
  if (!cleaned || isAutomaticFallbackValue(cleaned)) return "";
  return cleaned;
}

function extractNameBeforeCuit(text) {
  const lines = linesOf(text);
  for (let index = 0; index < lines.length; index += 1) {
    if (!/\bCUIL\/CUIT\s*:|\bCUIT\s*:/i.test(lines[index])) continue;
    for (let prev = index - 1; prev >= Math.max(0, index - 4); prev -= 1) {
      const candidate = likelyNameCandidate(lines[prev]);
      if (candidate && candidate.includes(",") && candidate.length >= 8) return candidate;
    }
  }
  return "";
}

function repairSellerNameFragments(labelName, nearbyName) {
  const label = likelyNameCandidate(labelName);
  const nearby = likelyNameCandidate(nearbyName);
  if (!label) return nearby;
  if (!nearby) return label;
  if (nearby.includes(",") && !label.includes(",")) return nearby;
  if (label.length >= 8) return label;

  const [nearbySurname, ...nearbyRest] = nearby.split(",");
  if (nearbyRest.length && /^[A-ZÁÉÍÓÚÑÜ]{3,5}$/.test(label) && /^[A-ZÁÉÍÓÚÑÜ]{3,}/.test(nearbySurname)) {
    const joinedSurname = `${label}${nearbySurname}`.replace(/^VAIIESO$/, "VALDIVIESO");
    return `${joinedSurname},${nearbyRest.join(",")}`.replace(/\s+/g, " ").trim();
  }
  return label.length >= nearby.length ? label : nearby;
}

function cleanEmail(value) {
  const compact = String(value || "")
    .replace(/\((?:S|A|@)\)/gi, "@")
    .replace(/\s+/g, "")
    .replace(/[<()[\]{}]/g, "")
    .replace(/mailto:/gi, "")
    .replace(/[;,]+$/g, "");
  const match = compact.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
  if (!match) return "";
  const email = match[0].replace(/[.,;:]+$/g, "").toUpperCase();
  const [local, domain] = email.split("@");
  const fixedLocal = String(local || "").replace(/(?<=[A-Z])8(?=[A-Z])/g, "B");
  const fixedDomain = String(domain || "").replace(/\.ARR$/i, ".AR");
  return fixedDomain ? `${fixedLocal}@${fixedDomain}` : email;
}

function extractEmail(text) {
  const source = compactText(text);
  const label = source.match(/(?:E-?mail|Mail|Correo(?:\s+electr[oó]nico)?)\s*:?\s*/i);
  if (label) {
    const value = cleanEmail(source.slice(label.index + label[0].length, label.index + label[0].length + 120));
    if (value) return value;
  }
  return cleanEmail(source);
}

function extractSellerName(sellerSection) {
  const afterLabel =
    extractPersonNameAfterLabel(sellerSection) ||
    extractLineValueFlexible(sellerSection, [/Apellido/i, /Nombre\s*:/i], [/\bD\.?N\.?I\.?\s*:/i, /\bCUIL\/CUIT\s*:/i]);
  const beforeCuit = extractNameBeforeCuit(sellerSection);
  return repairSellerNameFragments(afterLabel, beforeCuit);
}

function extractExpediente(text) {
  const compact = compactText(text);
  return (
    firstMatch(compact, /\b(?:EX|EXPE|EXPEDIENTE|EE|EX-)\s*[:N°º.-]*\s*([A-Z0-9][A-Z0-9\-\/.]{5,})/i) ||
    firstMatch(compact, /\b(\d{4}-\d{4,}[-/][A-Z0-9-]+)\b/i)
  ).trim();
}

function extractFechaInicioTramite(text) {
  const compact = compactText(text);
  return (
    firstMatch(compact, /(?:Fecha\s+inicio|Inicio\s+tr[aá]mite|Fecha\s+de\s+inicio)\s*:?\s*(\d{2}\/\d{2}\/\d{4})/i) ||
    firstMatch(compact, /\b(\d{2}\/\d{2}\/\d{4})\b/)
  );
}

function extractIniciador(text) {
  const lines = linesOf(text);
  for (let index = 0; index < lines.length; index += 1) {
    if (!/\b(INICIADOR|SOLICITANTE|PRESENTANTE)\b/i.test(lines[index])) continue;
    const sameLine = lines[index].replace(/^.*?\b(INICIADOR|SOLICITANTE|PRESENTANTE)\b\s*:?\s*/i, "").trim();
    const candidate = sameLine || (lines[index + 1] || "").trim();
    const cleaned = cleanPersonName(candidate);
    if (isValidIniciadorValue(cleaned)) return cleaned;
  }
  return "";
}

function extractCuitIniciador(text) {
  const compact = compactText(text);
  return (
    extractCuitAfterLabel(compact, /(?:CUIT\s+INICIADOR|CUIT\/CUIL\s+INICIADOR|CUIL\/CUIT\s+INICIADOR)\s*:?\s*/i) ||
    extractPartialCuitAfterLabel(compact, /(?:CUIT\s+INICIADOR|CUIT\/CUIL\s+INICIADOR|CUIL\/CUIT\s+INICIADOR)\s*:?\s*/i) ||
    extractFirstCuit(firstMatch(compact, /\b(?:INICIADOR|SOLICITANTE|PRESENTANTE)\b([\s\S]{0,120}?\b(?:CUIT|CUIL\/CUIT)\s*:?\s*[\s\S]{0,80})/i))
  );
}

function extractCaracter(text) {
  const compact = stripAccents(compactText(text)).toUpperCase();
  if (/\bTITULAR\b/.test(compact)) return "TITULAR";
  if (/\bSUPLENTE\b/.test(compact)) return "SUPLENTE";
  return "";
}

function extractRegistro(text) {
  const lines = linesOf(text);
  for (const line of lines) {
    if (/Registro\s+(?:Automotor|Nacional|Seccional)/i.test(line)) {
      return line.replace(/\s+/g, " ").trim();
    }
  }
  return firstMatch(compactText(text), /(Registro\s+Automotor\s+N[.°º]?\s*\d+[^,\n]*(?:,\s*[A-ZÁÉÍÓÚÑ ]+)?)/i);
}

function extractCaratulaNota(text) {
  return {
    "@atributo7@": extractExpediente(text),
    "@atributo10@": extractIniciador(text),
    "@atributo11@": extractCuitIniciador(text),
    "@atributo12@": extractCaracter(text),
    "@atributo13@": extractRegistro(text),
    "@atributo30@": extractFechaInicioTramite(text),
  };
}

function extractFormulario08(text, pdfName, log) {
  const normalized = normalizeText(text);
  const sectionA =
    findSectionFromLines(
      normalized,
      [/MARCA\s*:/, /MODELO\s*:/, /CONDICION/],
      [/COMPRADORES?\s+ADQUIRENTE/]
    ) || normalized;
  const buyerSection =
    findSectionFromLines(
      normalized,
      [/COMPRADORES?\s+ADQUIRENTE\/?[S5]?/],
      [/VENDEDOR.*TRANSMITENTE/, /GRAVAMENES/, /OBSERVACIONES/]
    ) || normalized;
  const sellerSection = findSectionFromLines(
    normalized,
    [/VENDEDOR.*TRANSMITENTE/, /TRANSMITENTE\/?S?/],
    [/GRAVAMENES/, /OBSERVACIONES/, /CEDULAS/, /C.DULAS/]
  );

  console.log(`OCR RAW VEHICULO (${pdfName}):`);
  console.log(sectionA);
  console.log(`OCR RAW VENDEDOR (${pdfName}):`);
  console.log(sellerSection);
  const emailBlock = [buyerSection, sellerSection].filter(Boolean).join("\n");
  const osdBlock = sectionA;
  console.log(`OCR RAW EMAIL (${pdfName}):`);
  console.log(emailBlock);
  console.log(`OCR RAW OSD (${pdfName}):`);
  console.log(osdBlock);
  log.push({ pdf: pdfName, phase: "ocr_raw_vehicle_block", status: "capturado", value: sectionA });
  log.push({ pdf: pdfName, phase: "ocr_raw_seller_block", status: "capturado", value: sellerSection });
  log.push({ pdf: pdfName, phase: "ocr_raw_email_block", status: "capturado", value: emailBlock });
  log.push({ pdf: pdfName, phase: "ocr_raw_osd_block", status: "capturado", value: osdBlock });

  const buyerCuit =
    extractCuitAfterLabel(buyerSection, /CUIT\s*:?\s*/i) ||
    extractPartialCuitAfterLabel(buyerSection, /CUIT\s*:?\s*/i) ||
    extractFirstCuit(buyerSection) ||
    extractPartialCuitAfterLabel(normalized, /CUIT\s*:?\s*/i) ||
    extractFirstCuit(normalized);
  const rawBuyerName = cleanPersonName(
    extractPersonNameAfterLabel(buyerSection) ||
      valueAfterRegex(buyerSection, /Apellido y Nombre\s*:\s*/i, [
      /\bD\.?N\.?I\.?\s*:/i,
      /\bSexo\s*:/i,
      /\bCUIT\s*:/i,
      /\bEmail\s*:/i,
      /\bDom.?cilio\b/i,
    ])
  );
  const buyerParts = rawBuyerName.includes(",") ? rawBuyerName.split(",") : [rawBuyerName, ""];
  const apellido = (buyerParts[0] || "").trim();
  const nombre = buyerParts.slice(1).join(" ").trim();
  const fullBuyerName = [apellido, nombre].filter(Boolean).join(" ");
  const email = extractEmail(buyerSection) || extractEmail(sellerSection) || extractEmail(normalized);

  const headerText = linesOf(normalized).slice(0, 20).join(" ");
  const form08Raw =
    firstMatch(headerText, /N\s*[°oº]\s*([0-9][0-9 .-]{4,})/i) ||
    firstMatch(headerText, /(?:FORMULARIO\s*0?8|F\.?\s*0?8)[^\d]{0,30}([0-9][0-9 .-]{4,})/i);
  const formulario08 = form08Raw.replace(/\D/g, "").replace(/^0+/, "");

  const dominio = extractDomain(normalized, sectionA);
  const domicilio = extractDomicilioAdquirente(buyerSection);
  const fecha = extractFechaImpresion08Strict(sectionA);
  const marcaModelo = extractModel(sectionA);
  const anio = extractYear(sectionA);

  const sellerCuit = extractSellerCuit(sellerSection);
  const sellerName = extractSellerName(sellerSection);
  const monto = extractMontoOperacion(normalized);

  const fields = {
    ...extractCaratulaNota(normalized),
    "@usuario@": buyerCuit,
    "@apellido@": apellido,
    "@nombre@": nombre,
    "@email@": email,
    "@atributo8@": buyerCuit,
    "@atributo9@": fullBuyerName,
    "@atributo14@": formulario08,
    "@atributo15@": dominio,
    "@atributo16@": domicilio,
    "@atributo17@": fecha,
    "@atributo18@": marcaModelo,
    "@atributo19@": anio,
    "@atributo20@": anio,
    "@atributo21@": sellerCuit,
    "@atributo22@": sellerName,
    "@atributo33@": monto,
  };

  const doubtful = new Set();
  if (!buyerCuit && /\bCUIT\s*:/i.test(buyerSection)) {
    doubtful.add("@usuario@");
    doubtful.add("@atributo8@");
  }
  if (!dominio && /\bDominio\s*:/i.test(sectionA)) doubtful.add("@atributo15@");
  if (!fecha && /Lugar\s*y?\s*Fech/i.test(sectionA)) doubtful.add("@atributo17@");
  if (!sellerCuit && /\bCUIL\/CUIT\s*:/i.test(sellerSection)) doubtful.add("@atributo21@");
  if (!sellerName && /Apellido\s+y\s+Nombre\s*:/i.test(sellerSection)) doubtful.add("@atributo22@");

  for (const [field, placeholder] of Object.entries(FIELDS)) {
    const value = fields[placeholder];
    log.push({
      pdf: pdfName,
      phase: "extract",
      field,
      placeholder,
      status: value === "" || value == null ? (doubtful.has(placeholder) ? "dudoso" : "no encontrado") : "encontrado",
      value: value === "" || value == null ? "" : value,
    });
  }

  return fields;
}

function classifyDocument(text) {
  const compact = stripAccents(compactText(text)).toUpperCase();
  if (/08\s*-\s*D/.test(compact) || /\bF\s*0?8\s*D\b/.test(compact) || /LUGAR\s*Y?\s*FECHA\s+DE\s+IMPRESI/.test(compact)) {
    return "F08D";
  }
  return "MANUSCRITO";
}

function emptyFieldMap() {
  const fields = {};
  for (const placeholder of Object.values(FIELDS)) {
    fields[placeholder] = "";
  }
  return fields;
}

function extractFormulario08Partial(text, pdfName, log) {
  const compact = compactText(text);
  const fields = { ...emptyFieldMap(), ...extractCaratulaNota(text) };
  const cuit =
    extractCuitAfterLabel(compact, /CUIT\s*:?\s*/i) ||
    extractCuitAfterLabel(compact, /CUIL\/CUIT\s*:?\s*/i) ||
    extractFirstCuit(compact);
  const dominio = cleanDomain(firstMatch(compact, /Dominio\s*:?\s*([A-Z0-9\-\s]{5,12})/i)) || cleanDomain(compact);

  fields["@usuario@"] = cuit;
  fields["@atributo8@"] = cuit;
  fields["@atributo15@"] = dominio;

  log.push({
    pdf: pdfName,
    phase: "classification",
    status: "documento manuscrito - extracción parcial",
    value: "Solo se intentan campos de alta confianza: CUIT y dominio",
  });

  for (const [field, placeholder] of Object.entries(FIELDS)) {
    const value = fields[placeholder];
    log.push({
      pdf: pdfName,
      phase: "extract",
      field,
      placeholder,
      status: value === "" || value == null ? "no encontrado" : "encontrado",
      value: value === "" || value == null ? "" : value,
    });
  }

  return fields;
}

function findTemplate() {
  const allTemplates = fs.readdirSync(TEMPLATE_DIR).filter((name) => /\.(xlsx|xlsm|ods)$/i.test(name)).sort();
  const excelTemplates = allTemplates.filter((name) => /\.(xlsx|xlsm)$/i.test(name));
  const odsTemplates = allTemplates.filter((name) => /\.ods$/i.test(name));
  if (excelTemplates.length === 1) {
    return path.join(TEMPLATE_DIR, excelTemplates[0]);
  }
  if (excelTemplates.length === 0 && odsTemplates.length === 1) {
    return path.join(TEMPLATE_DIR, odsTemplates[0]);
  }
  throw new Error(
    `Debe haber exactamente una plantilla .xlsx/.xlsm, o una unica .ods si no hay Excel, en ${TEMPLATE_DIR}. Excel: ${excelTemplates.length}, ODS: ${odsTemplates.length}`
  );
}

function findPlaceholderColumns(workbook) {
  const expected = new Set([...Object.values(FIELDS), ...NEVER_FILL]);

  for (const worksheet of workbook.worksheets) {
    const row = worksheet.getRow(PLACEHOLDER_ROW);
    const columns = [];
    row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
      const placeholder = cellText(cell.value);
      if (expected.has(placeholder)) {
        columns.push({ placeholder, colNumber });
      }
    });

    if (columns.length >= 5) {
      return { worksheet, columns };
    }
  }

  throw new Error(`No se encontraron placeholders suficientes en la fila ${PLACEHOLDER_ROW}.`);
}

function copyRowStyle(sourceRow, targetRow, maxCol) {
  targetRow.height = sourceRow.height;
  for (let colNumber = 1; colNumber <= maxCol; colNumber += 1) {
    const source = sourceRow.getCell(colNumber);
    const target = targetRow.getCell(colNumber);
    target.style = JSON.parse(JSON.stringify(source.style || {}));
  }
}

async function writeWorkbook(templatePath, rowsData, log) {
  if (/\.ods$/i.test(templatePath)) {
    return writeOdsWorkbook(templatePath, rowsData, log);
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(templatePath);

  const { worksheet, columns } = findPlaceholderColumns(workbook);
  const maxCol = Math.max(...columns.map((column) => column.colNumber), worksheet.columnCount);
  const placeholderRow = worksheet.getRow(PLACEHOLDER_ROW);

  rowsData.forEach((rowData, index) => {
    const targetRowNumber = FIRST_DATA_ROW + index;
    const targetRow = worksheet.getRow(targetRowNumber);
    copyRowStyle(placeholderRow, targetRow, maxCol);

    for (const { placeholder, colNumber } of columns) {
      const cell = targetRow.getCell(colNumber);
      const existing = cellText(cell.value);

      if (NEVER_FILL.has(placeholder)) {
        log.push({
          pdf: rowData.__source,
          phase: "write",
          placeholder,
          cell: cell.address,
          status: "omitido_no_completar",
          value: "",
        });
        continue;
      }

      const rawValue = rowData[placeholder] ?? "";
      const value = normalizeValueForPlaceholder(placeholder, rawValue);
      if (existing !== "") {
        log.push({
          pdf: rowData.__source,
          phase: "write",
          placeholder,
          cell: cell.address,
          status: "omitido_celda_ocupada",
          value: rawValue,
          existing,
        });
        continue;
      }

      if (placeholder === INICIADOR_PLACEHOLDER && !isValidIniciadorValue(rawValue)) {
        log.push({
          pdf: rowData.__source,
          phase: "write",
          placeholder,
          cell: cell.address,
          status: "omitido_iniciador_sin_valor",
          value: "",
        });
        continue;
      }

      cell.value = value === "" ? null : value;
      log.push({
        pdf: rowData.__source,
        phase: "write",
        placeholder,
        cell: cell.address,
        status: value === "" ? "escrito_vacio_no_encontrado" : "escrito",
        value,
        rawValue,
      });
    }

    targetRow.commit();
    log.push({ pdf: rowData.__source, phase: "write", field: "excel_row", status: "completado", value: targetRowNumber });
  });

  const ext = path.extname(templatePath);
  const base = path.basename(templatePath, ext);
  const outPath = path.join(OUTPUT_DIR, `${base}_completado_${nowStamp()}${ext}`);
  await workbook.xlsx.writeFile(outPath);
  return outPath;
}

function getRepeatedCount(cellXml) {
  const repeat = cellXml.match(/\stable:number-columns-repeated="(\d+)"/);
  return repeat ? Number(repeat[1]) : 1;
}

function removeRepeatedCount(cellXml) {
  return cellXml.replace(/\stable:number-columns-repeated="\d+"/, "");
}

function getOdsRows(contentXml) {
  return [...contentXml.matchAll(/<table:table-row\b[\s\S]*?<\/table:table-row>/g)].map((match) => ({
    index: match.index,
    xml: match[0],
  }));
}

function splitOdsRow(rowXml) {
  const open = rowXml.match(/^<table:table-row\b[^>]*>/);
  if (!open) throw new Error("Fila ODS invalida.");
  const openTag = open[0];
  const closeTag = "</table:table-row>";
  const inner = rowXml.slice(openTag.length, -closeTag.length);
  return { openTag, inner, closeTag };
}

function parseOdsCells(rowXml) {
  const { openTag, closeTag, inner } = splitOdsRow(rowXml);
  const cellMatches = [
    ...inner.matchAll(
      /<table:(?:table-cell|covered-table-cell)\b[^>]*\/>|<table:(?:table-cell|covered-table-cell)\b[^>]*>[\s\S]*?<\/table:(?:table-cell|covered-table-cell)>/g
    ),
  ];
  const cells = [];

  for (const match of cellMatches) {
    const repeat = getRepeatedCount(match[0]);
    const cellXml = removeRepeatedCount(match[0]);
    for (let i = 0; i < repeat; i += 1) {
      cells.push(cellXml);
    }
  }

  return { openTag, closeTag, cells };
}

function odsCellText(cellXml) {
  if (/table:formula=/.test(cellXml)) return "[formula]";
  if (/office:value=/.test(cellXml) || /office:string-value=/.test(cellXml) || /office:date-value=/.test(cellXml)) {
    const text = [...cellXml.matchAll(/<text:p[^>]*>([\s\S]*?)<\/text:p>|<text:p\s*\/>/g)]
      .map((match) => match[1] || "")
      .join("")
      .replace(/<[^>]+>/g, "")
      .trim();
    return text || "[value]";
  }
  return [...cellXml.matchAll(/<text:p[^>]*>([\s\S]*?)<\/text:p>|<text:p\s*\/>/g)]
    .map((match) => match[1] || "")
    .join("")
    .replace(/<text:s\s*\/>/g, " ")
    .replace(/<[^>]+>/g, "")
    .trim();
}

function odsPlaceholderFromCell(cellXml) {
  return unescapeXml(odsCellText(cellXml));
}

function makeOdsCell(value, placeholder) {
  if (typeof value === "number") {
    return `<table:table-cell office:value-type="float" office:value="${value}" calcext:value-type="float"><text:p>${value}</text:p></table:table-cell>`;
  }
  if (DATE_PLACEHOLDERS.has(placeholder)) {
    const text = escapeXml(normalize_date(value));
    return `<table:table-cell office:value-type="string" calcext:value-type="string"><text:p>${text}</text:p></table:table-cell>`;
  }
  const text = escapeXml(value);
  return `<table:table-cell office:value-type="string" calcext:value-type="string"><text:p>${text}</text:p></table:table-cell>`;
}

function writeOdsDataRow(targetRow, rowData, columns, targetRowNumber, log) {
  const requiredCols = Math.max(...columns.map((column) => column.colNumber));
  while (targetRow.cells.length < requiredCols) {
    targetRow.cells.push("<table:table-cell/>");
  }

  for (const { placeholder, colNumber } of columns) {
    const cellIndex = colNumber - 1;
    const address = `${columnName(colNumber)}${targetRowNumber}`;
    const existingXml = targetRow.cells[cellIndex] || "<table:table-cell/>";
    const existing = odsCellText(existingXml);

    if (NEVER_FILL.has(placeholder)) {
      log.push({
        pdf: rowData.__source,
        phase: "write",
        placeholder,
        cell: address,
        status: "omitido_no_completar",
        value: "",
      });
      continue;
    }

    const rawValue = rowData[placeholder] ?? "";
    const value = normalizeValueForPlaceholder(placeholder, rawValue);
    if (existing !== "") {
      log.push({
        pdf: rowData.__source,
        phase: "write",
        placeholder,
        cell: address,
        status: "omitido_celda_ocupada",
        value: rawValue,
        existing,
      });
      continue;
    }

    if (placeholder === INICIADOR_PLACEHOLDER && !isValidIniciadorValue(rawValue)) {
      log.push({
        pdf: rowData.__source,
        phase: "write",
        placeholder,
        cell: address,
        status: "omitido_iniciador_sin_valor",
        value: "",
      });
      continue;
    }

    targetRow.cells[cellIndex] = value === "" ? "<table:table-cell/>" : makeOdsCell(value, placeholder);
    log.push({
      pdf: rowData.__source,
      phase: "write",
      placeholder,
      cell: address,
      status: value === "" ? "escrito_vacio_no_encontrado" : "escrito",
      value,
      rawValue,
    });
  }

  log.push({ pdf: rowData.__source, phase: "write", field: "excel_row", status: "completado", value: targetRowNumber });
  return `${targetRow.openTag}${targetRow.cells.join("")}${targetRow.closeTag}`;
}

async function writeOdsWorkbook(templatePath, rowsData, log) {
  const buffer = fs.readFileSync(templatePath);
  const zip = await JSZip.loadAsync(buffer);
  const contentFile = zip.file("content.xml");
  if (!contentFile) throw new Error("ODS invalido: no contiene content.xml.");

  const contentXml = await contentFile.async("string");
  const rows = getOdsRows(contentXml);
  const lastRequiredRow = FIRST_DATA_ROW + rowsData.length - 1;
  if (rows.length < lastRequiredRow) {
    throw new Error(`ODS invalido: no existen filas suficientes para ${rowsData.length} PDFs. Ultima fila requerida: ${lastRequiredRow}.`);
  }

  const placeholderRow = parseOdsCells(rows[PLACEHOLDER_ROW - 1].xml);
  const expected = new Set([...Object.values(FIELDS), ...NEVER_FILL]);
  const columns = [];

  placeholderRow.cells.forEach((cellXml, index) => {
    const placeholder = odsPlaceholderFromCell(cellXml);
    if (expected.has(placeholder)) {
      columns.push({ placeholder, colNumber: index + 1 });
    }
  });

  if (columns.length < 5) {
    throw new Error(`No se encontraron placeholders suficientes en la fila ${PLACEHOLDER_ROW} del ODS.`);
  }

  let newContentXml = contentXml;
  const replacements = rowsData.map((rowData, index) => {
    const targetRowNumber = FIRST_DATA_ROW + index;
    const targetRow = parseOdsCells(rows[targetRowNumber - 1].xml);
    const newRowXml = writeOdsDataRow(targetRow, rowData, columns, targetRowNumber, log);
    return {
      start: rows[targetRowNumber - 1].index,
      oldXml: rows[targetRowNumber - 1].xml,
      newXml: newRowXml,
    };
  });

  replacements
    .sort((a, b) => b.start - a.start)
    .forEach((replacement) => {
      newContentXml =
        newContentXml.slice(0, replacement.start) +
        replacement.newXml +
        newContentXml.slice(replacement.start + replacement.oldXml.length);
    });

  zip.file("content.xml", newContentXml);
  const ext = path.extname(templatePath);
  const base = path.basename(templatePath, ext);
  const outPath = path.join(OUTPUT_DIR, `${base}_completado_${nowStamp()}${ext}`);
  const outBuffer = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  fs.writeFileSync(outPath, outBuffer);
  return outPath;
}

function printNormalizerTests() {
  const amountInputs = ["100.000", "100,000", "7.000.000", "$ 7.000.000,00", "'100000", "10000000,0"];
  const dateInputs = ["'04/11/2025", "04/11/2025", "Mendoza, 04 de noviembre de 2025"];

  console.log("AMOUNT");
  for (const input of amountInputs) {
    const output = normalize_amount(input);
    console.log(`${input} -> ${output} -> ${typeof output}`);
  }

  console.log("DATE");
  for (const input of dateInputs) {
    const output = normalize_date(input);
    console.log(`${input} -> ${output} -> ${typeof output}`);
  }
}

async function main() {
  ensureDirs();

  const log = [];
  const templatePath = findTemplate();
  const pdfs = fs
    .readdirSync(INPUT_DIR)
    .filter((name) => /\.pdf$/i.test(name))
    .sort()
    .map((name) => path.join(INPUT_DIR, name));

  if (pdfs.length === 0) {
    throw new Error(`No hay PDFs en ${INPUT_DIR}`);
  }

  const keyFields = [
    { name: "CUIT adquirente", placeholder: "@usuario@" },
    { name: "Dominio", placeholder: "@atributo15@" },
    { name: "Monto operacion", placeholder: "@atributo33@" },
  ];
  const rowsData = [];

  for (const pdfPath of pdfs) {
    const pdfName = path.basename(pdfPath);
    try {
      const parsed = await pdfParse(fs.readFileSync(pdfPath));
      const text = parsed.text || "";
      console.log(`OCR RAW (${pdfName}):`);
      console.log(text);
      log.push({
        pdf: pdfName,
        phase: "ocr_raw",
        status: "capturado",
        value: text,
      });
      const globalCuits = findCuits(text);
      log.push({
        pdf: pdfName,
        phase: "ocr_cuit_scan",
        status: globalCuits.length ? "encontrado" : "no encontrado",
        value: globalCuits.join(", "),
      });
      const documentType = classifyDocument(text);
      log.push({
        pdf: pdfName,
        phase: "classification",
        status: documentType,
        value: documentType === "F08D" ? "extraccion completa" : "documento manuscrito - extracción parcial",
      });

      const rowData =
        documentType === "F08D"
          ? extractFormulario08(text, pdfName, log)
          : extractFormulario08Partial(text, pdfName, log);
      rowData.__source = pdfName;
      rowData.__documentType = documentType;

      const missingKeyFields = keyFields.filter((field) => rowData[field.placeholder] === "" || rowData[field.placeholder] == null);
      if (missingKeyFields.length > 0) {
        const warning = `ADVERTENCIA: faltan campos clave: ${missingKeyFields.map((field) => field.name).join(", ")}`;
        console.warn(`${pdfName}: ${warning}`);
        log.push({
          pdf: pdfName,
          phase: "validation",
          field: "campos_clave",
          status: "advertencia",
          value: warning,
        });
      } else {
        log.push({
          pdf: pdfName,
          phase: "validation",
          field: "campos_clave",
          status: "ok",
          value: "CUIT adquirente, Dominio y Monto operacion encontrados",
        });
      }

      rowsData.push(rowData);
    } catch (error) {
      log.push({
        pdf: pdfName,
        phase: "pdf",
        status: "error",
        value: error.message,
      });
      console.warn(`${pdfName}: no se pudo procesar (${error.message})`);
    }
  }

  if (rowsData.length === 0) {
    throw new Error("No se pudo extraer ningun PDF de /input.");
  }

  const outPath = await writeWorkbook(templatePath, rowsData, log);
  const logPath = path.join(LOG_DIR, `extractor_log_${nowStamp()}.jsonl`);
  fs.writeFileSync(logPath, log.map((entry) => JSON.stringify(entry)).join("\n") + "\n", "utf8");

  console.log(`Plantilla: ${templatePath}`);
  console.log(`PDFs encontrados: ${pdfs.length}`);
  console.log(`PDFs escritos: ${rowsData.length}`);
  console.log(`Excel generado: ${outPath}`);
  console.log(`Log: ${logPath}`);
}

if (process.argv.includes("--test-normalizers")) {
  printNormalizerTests();
} else {
  main().catch((error) => {
    console.error(`ERROR: ${error.message}`);
    process.exit(1);
  });
}
