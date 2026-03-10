const path = require("path");
const express = require("express");
const { google } = require("googleapis");

const app = express();
const PORT = process.env.PORT || 3000;

const SHEET_PUBLISH_ID =
  process.env.SHEET_PUBLISH_ID ||
  "2PACX-1vQL2uV2BS5DCGOlUQx4X2A7ABEWgC-c3CYA46B3S92pUG5H8VhFXta7qL00F3XjdqolkZ9jEPIqrp3Q";
const SHEET_TAB_NAME = process.env.SHEET_TAB_NAME || "Nomes";
const SHEET_NOMES_GID = process.env.SHEET_NOMES_GID || "1958765595";
const LISTA_TAB_NAME = process.env.SHEET_LISTA_TAB_NAME || "Lista";
const LISTA_COL_START = process.env.SHEET_LISTA_COL_START || "R";
const LISTA_COL_END = process.env.SHEET_LISTA_COL_END || "AZ";
const GOOGLE_SPREADSHEET_ID =
  process.env.GOOGLE_SPREADSHEET_ID || "1spXWVi4VD1wIkGVXMVdcLpm9dHAgCJ5CTCBgmpiUji8";
const GOOGLE_SERVICE_ACCOUNT_EMAIL = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL || "";
const GOOGLE_PRIVATE_KEY = process.env.GOOGLE_PRIVATE_KEY || "";

app.use(express.static("public"));
app.use(express.json());

function parseGoogleVizResponse(text) {
  const start = text.indexOf("{");
  const end = text.lastIndexOf("}");

  if (start === -1 || end === -1 || end <= start) {
    throw new Error("Resposta gviz em formato inesperado.");
  }

  return JSON.parse(text.slice(start, end + 1));
}

function normalizeCellValue(cell) {
  if (!cell) return "";

  if (typeof cell.f === "string" && cell.f.trim()) {
    return cell.f.trim();
  }

  if (cell.v == null) return "";
  return String(cell.v).trim();
}

function extractNamesFromTable(table) {
  if (!table || !Array.isArray(table.rows)) {
    return [];
  }

  const names = [];

  for (const row of table.rows) {
    if (!row || !Array.isArray(row.c) || row.c.length === 0) continue;

    const firstColumn = normalizeCellValue(row.c[0]);
    if (!firstColumn) continue;

    names.push(firstColumn);
  }

  return names;
}

function extractRowsFromGoogleVizTable(table) {
  if (!table || !Array.isArray(table.rows)) {
    return [];
  }

  return table.rows.map((row) => {
    const cells = Array.isArray(row?.c) ? row.c : [];
    return cells.map((cell) => normalizeCellValue(cell));
  });
}

function decodeHtml(text) {
  return text
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");
}

function escapeRegex(text) {
  return text.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function extractTabGidFromPubHtml(html, tabName) {
  const safeTab = escapeRegex(tabName.trim());
  const regex = new RegExp(
    `href="([^"]*?[?&]gid=(\\d+)[^"]*?)"[^>]*>\\s*${safeTab}\\s*<`,
    "i"
  );
  const match = html.match(regex);
  return match ? match[2] : null;
}

function parseCsvFirstColumn(csvText) {
  return parseCsv(csvText)
    .map((r) => (r[0] || "").trim())
    .filter(Boolean);
}

function parseCsv(csvText) {
  const rows = [];
  let row = [];
  let field = "";
  let i = 0;
  let inQuotes = false;

  while (i < csvText.length) {
    const char = csvText[i];
    const next = csvText[i + 1];

    if (inQuotes) {
      if (char === '"' && next === '"') {
        field += '"';
        i += 2;
        continue;
      }
      if (char === '"') {
        inQuotes = false;
        i += 1;
        continue;
      }
      field += char;
      i += 1;
      continue;
    }

    if (char === '"') {
      inQuotes = true;
      i += 1;
      continue;
    }

    if (char === ",") {
      row.push(field);
      field = "";
      i += 1;
      continue;
    }

    if (char === "\r") {
      i += 1;
      continue;
    }

    if (char === "\n") {
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
      i += 1;
      continue;
    }

    field += char;
    i += 1;
  }

  if (field.length > 0 || row.length > 0) {
    row.push(field);
    rows.push(row);
  }

  return rows;
}

function normalizeText(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function normalizeDateLabel(value) {
  const text = String(value || "").trim();
  if (!text) return "";

  const match = text.match(
    /^0*(\d{1,2})\s*[\/\-.]\s*0*(\d{1,2})(?:\s*[\/\-.]\s*\d{2,4})?$/
  );
  if (!match) return text;

  const day = Number(match[1]);
  const month = Number(match[2]);

  if (!day || !month) return text;
  return `${day}/${month}`;
}

function toA1Column(colIndexZeroBased) {
  let n = colIndexZeroBased + 1;
  let result = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
}

function fromA1Column(colLabel) {
  const value = String(colLabel || "").trim().toUpperCase();
  if (!/^[A-Z]+$/.test(value)) return -1;

  let total = 0;
  for (const char of value) {
    total = total * 26 + (char.charCodeAt(0) - 64);
  }

  return total - 1;
}

function getListaColumnBounds() {
  const startCol = fromA1Column(LISTA_COL_START);
  const endCol = fromA1Column(LISTA_COL_END);

  const defaultStart = fromA1Column("R");
  const defaultEnd = fromA1Column("AZ");

  if (startCol < 0 || endCol < 0) {
    return {
      startColIndex: defaultStart,
      endColIndex: defaultEnd,
      startColLabel: "R",
      endColLabel: "AZ"
    };
  }

  const minCol = Math.min(startCol, endCol);
  const maxCol = Math.max(startCol, endCol);

  return {
    startColIndex: minCol,
    endColIndex: maxCol,
    startColLabel: toA1Column(minCol),
    endColLabel: toA1Column(maxCol)
  };
}

function isValidA1Cell(cell) {
  return /^[A-Z]+[1-9]\d*$/.test(String(cell || "").trim());
}

function getSheetsClient() {
  if (!GOOGLE_SERVICE_ACCOUNT_EMAIL || !GOOGLE_PRIVATE_KEY || !GOOGLE_SPREADSHEET_ID) {
    throw new Error(
      "Credenciais do Google Sheets nao configuradas. Configure GOOGLE_SERVICE_ACCOUNT_EMAIL, GOOGLE_PRIVATE_KEY e GOOGLE_SPREADSHEET_ID."
    );
  }

  const auth = new google.auth.JWT(
    GOOGLE_SERVICE_ACCOUNT_EMAIL,
    null,
    GOOGLE_PRIVATE_KEY.replace(/\\n/g, "\n"),
    ["https://www.googleapis.com/auth/spreadsheets"]
  );

  return google.sheets({ version: "v4", auth });
}

function formatTodayPtBrShort() {
  const now = new Date();
  return `${now.getDate()}/${now.getMonth() + 1}`;
}

const TURMA_OPTIONS = [
  "8 Ano A",
  "8 Ano B",
  "9 Ano Anchieta",
  "1 Série Funcionários",
  "1 Série Anchieta",
  "2 Série Anchieta",
  "3 Série Anchieta"
];

const TURMA_ALIASES = {
  "8 ano a": ["8 a", "8 ano a"],
  "8 ano b": ["8 b", "8 ano b"],
  "9 ano anchieta": ["9 ano anchieta", "9 anchieta", "9 ano", "9 a", "9"],
  "1 serie funcionarios": ["1 serie funcionarios", "1 serie", "1 série"],
  "1 serie anchieta": ["1 serie anchieta", "1 serie", "1 série"],
  "2 serie anchieta": ["2 serie anchieta", "2 serie", "2 série"],
  "3 serie anchieta": ["3 serie anchieta", "3 serie", "3 série"]
};

function parseTurmaDescriptor(text) {
  const normalized = normalizeText(text);
  if (!normalized) return null;

  const tokens = normalized.split(" ").filter(Boolean);
  const numberMatch = normalized.match(/^(\d{1,2})/);

  return {
    normalized,
    tokens,
    number: numberMatch ? numberMatch[1] : null,
    letter: tokens.find((token) => /^[ab]$/.test(token)) || null,
    stage: tokens.includes("ano") ? "ano" : tokens.includes("serie") ? "serie" : null,
    campus: tokens.includes("anchieta")
      ? "anchieta"
      : tokens.includes("funcionarios")
        ? "funcionarios"
        : null
  };
}

function looksLikeTurmaHeader(cellA) {
  const descriptor = parseTurmaDescriptor(cellA);
  if (!descriptor?.number) return false;
  if (!/^\d{1,2}/.test(descriptor.normalized)) return false;

  // Turma labels are short markers, unlike student full names.
  if (descriptor.tokens.length > 4) return false;

  return Boolean(descriptor.stage || descriptor.letter || descriptor.campus);
}

function scoreTurmaCandidate(selectedDescriptor, candidateDescriptor) {
  if (!selectedDescriptor || !candidateDescriptor?.number) {
    return Number.NEGATIVE_INFINITY;
  }

  let score = 0;

  if (selectedDescriptor.number) {
    if (selectedDescriptor.number !== candidateDescriptor.number) {
      return Number.NEGATIVE_INFINITY;
    }
    score += 100;
  }

  if (selectedDescriptor.letter) {
    if (candidateDescriptor.letter && selectedDescriptor.letter !== candidateDescriptor.letter) {
      return Number.NEGATIVE_INFINITY;
    }
    score += candidateDescriptor.letter ? 20 : 4;
  }

  if (selectedDescriptor.stage) {
    if (candidateDescriptor.stage === selectedDescriptor.stage) {
      score += 12;
    } else if (!candidateDescriptor.stage) {
      score += 2;
    }
  }

  if (selectedDescriptor.campus) {
    if (candidateDescriptor.campus && candidateDescriptor.campus !== selectedDescriptor.campus) {
      return Number.NEGATIVE_INFINITY;
    }
    score += candidateDescriptor.campus ? 18 : 1;
  }

  if (candidateDescriptor.normalized === selectedDescriptor.normalized) {
    score += 30;
  }

  return score;
}

function findDateHeaderRow(rows) {
  return rows.findIndex((row) => normalizeText(row[0]) === "turma");
}

function extractDateColumnsFromRow(row, rowIndex) {
  const dates = [];

  for (let col = 0; col < row.length; col++) {
    const raw = normalizeDateLabel(row[col]);
    if (!/^\d{1,2}\/\d{1,2}$/.test(raw)) continue;
    dates.push({
      value: raw,
      colIndex: col,
      a1Column: toA1Column(col)
    });
  }

  return {
    dateRowIndex: rowIndex,
    dates
  };
}

function extractDateColumns(rows) {
  const headerRowIndex = findDateHeaderRow(rows);
  if (headerRowIndex === -1) return { dateRowIndex: -1, dates: [] };

  const row = rows[headerRowIndex] || [];
  return extractDateColumnsFromRow(row, headerRowIndex);
}

function extractDateColumnsNearTurma(rows, turmaRowIndex, selectedDate) {
  if (turmaRowIndex < 0) {
    return { dateRowIndex: -1, dates: [] };
  }

  const candidateIndexes = [
    turmaRowIndex,
    turmaRowIndex + 1,
    turmaRowIndex + 2,
    turmaRowIndex - 1
  ].filter((rowIndex, index, list) => {
    return rowIndex >= 0 && rowIndex < rows.length && list.indexOf(rowIndex) === index;
  });

  const candidates = candidateIndexes
    .map((rowIndex) => extractDateColumnsFromRow(rows[rowIndex] || [], rowIndex))
    .filter((candidate) => candidate.dates.length > 0);

  if (!candidates.length) {
    return { dateRowIndex: -1, dates: [] };
  }

  const withSelectedDate = candidates.filter((candidate) =>
    candidate.dates.some((date) => date.value === selectedDate)
  );
  const rankedCandidates = withSelectedDate.length ? withSelectedDate : candidates;

  rankedCandidates.sort((a, b) => {
    const aDistance = Math.abs(a.dateRowIndex - turmaRowIndex);
    const bDistance = Math.abs(b.dateRowIndex - turmaRowIndex);

    if (aDistance !== bDistance) return aDistance - bDistance;
    if (a.dates.length !== b.dates.length) return b.dates.length - a.dates.length;
    return a.dateRowIndex - b.dateRowIndex;
  });

  return rankedCandidates[0];
}

function findTurmaRow(rows, turmaSelection) {
  const normalizedSelection = normalizeText(turmaSelection);
  const aliases = TURMA_ALIASES[normalizedSelection] || [normalizedSelection];
  const aliasSet = new Set(aliases.map(normalizeText));
  const selectedDescriptor = parseTurmaDescriptor(turmaSelection);
  const candidates = [];

  for (let i = 0; i < rows.length; i++) {
    const rawCellA = String(rows[i]?.[0] || "").trim();
    const normalizedCellA = normalizeText(rawCellA);
    if (!normalizedCellA) continue;

    const aliasMatch = aliasSet.has(normalizedCellA);
    const headerLike = looksLikeTurmaHeader(rawCellA);
    if (!aliasMatch && !headerLike) continue;

    const candidateDescriptor = parseTurmaDescriptor(rawCellA);
    const descriptorScore = scoreTurmaCandidate(selectedDescriptor, candidateDescriptor);

    if (descriptorScore === Number.NEGATIVE_INFINITY && !aliasMatch) {
      continue;
    }

    let score = aliasMatch ? 200 : 0;
    if (descriptorScore !== Number.NEGATIVE_INFINITY) {
      score += descriptorScore;
    }

    candidates.push({
      rowIndex: i,
      score
    });
  }

  if (!candidates.length) return -1;

  candidates.sort((a, b) => b.score - a.score || a.rowIndex - b.rowIndex);

  const bestScore = candidates[0].score;
  const bestCandidates = candidates.filter((candidate) => candidate.score === bestScore);

  if (
    bestCandidates.length > 1 &&
    selectedDescriptor?.number === "1" &&
    selectedDescriptor?.stage === "serie" &&
    selectedDescriptor?.campus
  ) {
    if (selectedDescriptor.campus === "funcionarios") {
      return bestCandidates[0].rowIndex;
    }

    if (selectedDescriptor.campus === "anchieta") {
      return bestCandidates[bestCandidates.length - 1].rowIndex;
    }
  }

  return bestCandidates[0].rowIndex;
}

function isKnownTurmaLabel(cellA) {
  if (looksLikeTurmaHeader(cellA)) return true;

  const normalized = normalizeText(cellA);
  if (!normalized) return false;

  const allAliases = Object.values(TURMA_ALIASES).flat().map(normalizeText);
  return allAliases.includes(normalized);
}

function extractStudentsForTurma(rows, startRowIndex, dateColIndex) {
  if (startRowIndex < 0) return [];
  const students = [];

  for (let i = startRowIndex + 1; i < rows.length; i++) {
    const row = rows[i] || [];
    const name = String(row[0] || "").trim();
    const rowHasAnyValue = row.some((cell) => String(cell || "").trim() !== "");

    if (!rowHasAnyValue) {
      if (students.length) break;
      continue;
    }

    if (!name) {
      if (students.length) break;
      continue;
    }

    if (isKnownTurmaLabel(name)) break;

    students.push({
      name,
      rowIndex: i,
      rowNumber: i + 1,
      currentValue:
        typeof dateColIndex === "number"
          ? String(row[dateColIndex] || "").trim().toUpperCase()
          : ""
    });
  }

  return students;
}

function buildTurmaSelectionData(rows, sheetName, selectedDate, selectedTurma, options = {}) {
  const normalizedSelectedDate = normalizeDateLabel(selectedDate);
  const turmaRowIndex = findTurmaRow(rows, selectedTurma);
  const turmaRowNumber = turmaRowIndex >= 0 ? turmaRowIndex + 1 : null;
  const turmaCell = turmaRowNumber ? `A${turmaRowNumber}` : null;
  const useTurmaRowForDates = options.dateRowMode === "turma_row";
  const dateSource =
    useTurmaRowForDates && turmaRowIndex >= 0
      ? extractDateColumnsNearTurma(rows, turmaRowIndex, normalizedSelectedDate)
      : extractDateColumns(rows);
  const { dateRowIndex, dates } = dateSource;
  const dateMatches = dates.filter((d) => d.value === normalizedSelectedDate);
  const selectedDateColumn = dateMatches[0] || null;
  const studentStartRowIndex = Math.max(turmaRowIndex, dateRowIndex);
  const students = extractStudentsForTurma(
    rows,
    studentStartRowIndex,
    selectedDateColumn ? selectedDateColumn.colIndex : undefined
  ).map((student) => ({
    ...student,
    cell: selectedDateColumn ? `${selectedDateColumn.a1Column}${student.rowNumber}` : null
  }));

  return {
    sheet: sheetName,
    availableDates: dates.map((d) => d.value),
    selected: {
      date: selectedDate,
      normalizedDate: normalizedSelectedDate,
      turma: selectedTurma,
      dateColumn: selectedDateColumn,
      dateMatchesCount: dateMatches.length,
      dateHeaderRow: dateRowIndex >= 0 ? dateRowIndex + 1 : null,
      turmaRow: turmaRowNumber,
      turmaCell
    },
    students
  };
}

function buildStudentNameBuckets(students) {
  const buckets = new Map();

  for (const student of students || []) {
    const key = normalizeText(student.name);
    if (!key) continue;
    if (!buckets.has(key)) buckets.set(key, []);
    buckets.get(key).push(student);
  }

  return buckets;
}

function mergeStudentsWithNomes(chamadaStudents, nomesSelection) {
  const nomesStudents = nomesSelection?.students || [];
  const namesBuckets = buildStudentNameBuckets(nomesStudents);

  return (chamadaStudents || []).map((student, index) => {
    const nameKey = normalizeText(student.name);
    let nomesStudent = null;

    if (nameKey && namesBuckets.has(nameKey)) {
      const bucket = namesBuckets.get(nameKey);
      nomesStudent = bucket.shift() || null;
    }

    // Fallback por posição relativa quando os nomes divergem levemente entre abas.
    if (!nomesStudent && index < nomesStudents.length) {
      nomesStudent = nomesStudents[index];
    }

    const nomesDateColumn = nomesSelection?.selected?.dateColumn || null;
    const nomesCell =
      nomesStudent && nomesDateColumn
        ? `${nomesDateColumn.a1Column}${nomesStudent.rowNumber}`
        : null;

    return {
      ...student,
      nomesRowNumber: nomesStudent ? nomesStudent.rowNumber : null,
      nomesCurrentValue: nomesStudent ? nomesStudent.currentValue || "" : "",
      nomesCell,
      nomesMatched: Boolean(nomesStudent)
    };
  });
}

async function fetchPublishedTabRowsByGid(tabName, gid) {
  const csvUrl = `https://docs.google.com/spreadsheets/d/e/${SHEET_PUBLISH_ID}/pub?gid=${gid}&single=true&output=csv`;
  const csvResponse = await fetch(csvUrl, {
    headers: { "User-Agent": "Mozilla/5.0" }
  });

  if (!csvResponse.ok) {
    throw new Error(`Erro ao baixar CSV da aba ${tabName}: ${csvResponse.status}`);
  }

  const csvText = await csvResponse.text();
  return { rows: parseCsv(csvText), sourceUrl: csvUrl, gid };
}

async function fetchPublishedTabRowsByName(tabName) {
  const pubHtmlUrl = `https://docs.google.com/spreadsheets/d/e/${SHEET_PUBLISH_ID}/pubhtml`;
  const pubHtmlResponse = await fetch(pubHtmlUrl, {
    headers: { "User-Agent": "Mozilla/5.0" }
  });

  if (!pubHtmlResponse.ok) {
    throw new Error(
      `Erro ao consultar pagina publicada da planilha: ${pubHtmlResponse.status}`
    );
  }

  const pubHtml = decodeHtml(await pubHtmlResponse.text());
  const gid = extractTabGidFromPubHtml(pubHtml, tabName);

  if (!gid) {
    throw new Error(`Nao foi possivel localizar a aba "${tabName}".`);
  }

  return fetchPublishedTabRowsByGid(tabName, gid);
}

async function fetchTabRowsByGviz(tabName) {
  const url = `https://docs.google.com/spreadsheets/d/e/${SHEET_PUBLISH_ID}/gviz/tq?sheet=${encodeURIComponent(
    tabName
  )}&tqx=out:json`;

  const response = await fetch(url, {
    headers: {
      "User-Agent": "Mozilla/5.0"
    }
  });

  if (!response.ok) {
    throw new Error(`Erro ao consultar Google Sheets (${tabName}): ${response.status}`);
  }

  const text = await response.text();
  const payload = parseGoogleVizResponse(text);

  return {
    rows: extractRowsFromGoogleVizTable(payload.table),
    sourceUrl: url
  };
}

async function fetchNomesSheetRows() {
  const attempts = [];

  if (SHEET_NOMES_GID) {
    attempts.push(() => fetchPublishedTabRowsByGid(SHEET_TAB_NAME, SHEET_NOMES_GID));
  }

  attempts.push(() => fetchTabRowsByGviz(SHEET_TAB_NAME));
  attempts.push(() => fetchPublishedTabRowsByName(SHEET_TAB_NAME));

  let lastError = null;

  for (const attempt of attempts) {
    try {
      return await attempt();
    } catch (error) {
      lastError = error;
    }
  }

  throw lastError || new Error(`Nao foi possivel localizar a aba "${SHEET_TAB_NAME}".`);
}

async function fetchChamadaData(selectedDate, selectedTurma) {
  const todaySuggestion = formatTodayPtBrShort();
  const chosenDate = String(selectedDate || todaySuggestion).trim();
  const chosenTurma = String(selectedTurma || TURMA_OPTIONS[0]).trim();
  const nomesSource = await fetchNomesSheetRows();
  const nomesSelection = buildTurmaSelectionData(
    nomesSource.rows,
    SHEET_TAB_NAME,
    chosenDate,
    chosenTurma,
    { dateRowMode: "turma_row" }
  );

  return {
    sourceUrl: nomesSource.sourceUrl,
    sheet: SHEET_TAB_NAME,
    todaySuggestion,
    turmaOptions: TURMA_OPTIONS,
    availableDates: nomesSelection.availableDates,
    selected: nomesSelection.selected,
    students: nomesSelection.students
  };
}

async function fetchNamesFromPublishedCsv() {
  const pubHtmlUrl = `https://docs.google.com/spreadsheets/d/e/${SHEET_PUBLISH_ID}/pubhtml`;
  const pubHtmlResponse = await fetch(pubHtmlUrl, {
    headers: { "User-Agent": "Mozilla/5.0" }
  });

  if (!pubHtmlResponse.ok) {
    throw new Error(
      `Erro ao consultar página publicada da planilha: ${pubHtmlResponse.status}`
    );
  }

  const pubHtml = decodeHtml(await pubHtmlResponse.text());
  const gid = extractTabGidFromPubHtml(pubHtml, SHEET_TAB_NAME);

  if (!gid) {
    throw new Error(`Não foi possível localizar a aba "${SHEET_TAB_NAME}".`);
  }

  const csvUrl = `https://docs.google.com/spreadsheets/d/e/${SHEET_PUBLISH_ID}/pub?gid=${gid}&single=true&output=csv`;
  const csvResponse = await fetch(csvUrl, {
    headers: { "User-Agent": "Mozilla/5.0" }
  });

  if (!csvResponse.ok) {
    throw new Error(`Erro ao baixar CSV da aba ${SHEET_TAB_NAME}: ${csvResponse.status}`);
  }

  const csvText = await csvResponse.text();
  const names = parseCsvFirstColumn(csvText);

  return {
    names,
    sourceUrl: csvUrl,
    sheet: SHEET_TAB_NAME,
    updatedAt: new Date().toISOString()
  };
}

async function fetchNames() {
  const url = `https://docs.google.com/spreadsheets/d/e/${SHEET_PUBLISH_ID}/gviz/tq?sheet=${encodeURIComponent(
    SHEET_TAB_NAME
  )}&tqx=out:json`;

  try {
    const response = await fetch(url, {
      headers: {
        "User-Agent": "Mozilla/5.0"
      }
    });

    if (!response.ok) {
      throw new Error(`Erro ao consultar Google Sheets: ${response.status}`);
    }

    const text = await response.text();
    const payload = parseGoogleVizResponse(text);
    const names = extractNamesFromTable(payload.table);

    return {
      names,
      sourceUrl: url,
      sheet: SHEET_TAB_NAME,
      updatedAt: new Date().toISOString()
    };
  } catch (_error) {
    // Fallback para planilhas publicadas no formato /d/e/.../pubhtml
    return fetchNamesFromPublishedCsv();
  }
}

const LISTA_IGNORED_LABELS = new Set(["gabaritos"]);

function normalizeListaCellText(cell) {
  const formatted = String(cell?.formattedValue || "").trim();
  if (formatted) return formatted;

  if (cell?.effectiveValue?.stringValue != null) {
    return String(cell.effectiveValue.stringValue).trim();
  }

  if (cell?.effectiveValue?.numberValue != null) {
    return String(cell.effectiveValue.numberValue).trim();
  }

  if (cell?.effectiveValue?.boolValue != null) {
    return cell.effectiveValue.boolValue ? "TRUE" : "FALSE";
  }

  if (cell?.userEnteredValue?.stringValue != null) {
    return String(cell.userEnteredValue.stringValue).trim();
  }

  if (cell?.userEnteredValue?.numberValue != null) {
    return String(cell.userEnteredValue.numberValue).trim();
  }

  return "";
}

function extractHyperlinkFromFormula(formulaValue) {
  const formula = String(formulaValue || "").trim();
  if (!formula) return "";

  const directMatch = formula.match(/HYPERLINK\s*\(\s*"([^"]+)"/i);
  if (!directMatch) return "";

  return String(directMatch[1] || "").trim();
}

function normalizeListaHyperlink(link) {
  const text = String(link || "").trim();
  if (!text) return "";

  if (/^https?:\/\//i.test(text)) return text;

  if (text.startsWith("www.")) {
    return `https://${text}`;
  }

  return "";
}

function extractListaCellLink(cell) {
  const direct = normalizeListaHyperlink(cell?.hyperlink);
  if (direct) return direct;

  const richTextLink = normalizeListaHyperlink(
    cell?.userEnteredFormat?.textFormat?.link?.uri ||
      cell?.effectiveFormat?.textFormat?.link?.uri
  );
  if (richTextLink) return richTextLink;

  if (Array.isArray(cell?.textFormatRuns)) {
    for (const run of cell.textFormatRuns) {
      const link = normalizeListaHyperlink(run?.format?.link?.uri);
      if (link) return link;
    }
  }

  return normalizeListaHyperlink(
    extractHyperlinkFromFormula(cell?.userEnteredValue?.formulaValue)
  );
}

async function fetchListaCellsBySheetsApi() {
  const sheets = getSheetsClient();
  const bounds = getListaColumnBounds();
  const a1Range = `'${LISTA_TAB_NAME}'!${bounds.startColLabel}:${bounds.endColLabel}`;

  const response = await sheets.spreadsheets.get({
    spreadsheetId: GOOGLE_SPREADSHEET_ID,
    includeGridData: true,
    ranges: [a1Range]
  });

  const grid = response?.data?.sheets?.[0]?.data?.[0];
  const rowData = Array.isArray(grid?.rowData) ? grid.rowData : [];
  const startRow = Number(grid?.startRow || 0);
  const startCol = Number(grid?.startColumn ?? bounds.startColIndex);

  const cells = [];
  let linkCount = 0;

  for (let rowOffset = 0; rowOffset < rowData.length; rowOffset++) {
    const values = Array.isArray(rowData[rowOffset]?.values) ? rowData[rowOffset].values : [];

    for (let colOffset = 0; colOffset < values.length; colOffset++) {
      const cell = values[colOffset];
      const text = normalizeListaCellText(cell);
      const link = extractListaCellLink(cell);

      if (!text && !link) continue;

      const absoluteCol = startCol + colOffset;
      if (absoluteCol < bounds.startColIndex || absoluteCol > bounds.endColIndex) continue;

      const absoluteRow = startRow + rowOffset;
      if (link) linkCount += 1;

      cells.push({
        row: absoluteRow + 1,
        col: absoluteCol + 1,
        a1: `${toA1Column(absoluteCol)}${absoluteRow + 1}`,
        text,
        link
      });
    }
  }

  return {
    cells,
    source: "google_sheets_api",
    sourceUrl: a1Range,
    linkCount
  };
}

async function fetchListaCellsByGviz() {
  const bounds = getListaColumnBounds();
  const range = `${bounds.startColLabel}:${bounds.endColLabel}`;
  const url = `https://docs.google.com/spreadsheets/d/${GOOGLE_SPREADSHEET_ID}/gviz/tq?sheet=${encodeURIComponent(
    LISTA_TAB_NAME
  )}&range=${encodeURIComponent(range)}&tqx=out:csv`;

  const response = await fetch(url, {
    headers: { "User-Agent": "Mozilla/5.0" }
  });

  if (!response.ok) {
    throw new Error(`Erro ao carregar CSV da aba ${LISTA_TAB_NAME}: ${response.status}`);
  }

  const csvText = await response.text();
  const rows = parseCsv(csvText);
  const cells = [];

  for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
    const row = rows[rowIndex] || [];
    for (let colOffset = 0; colOffset < row.length; colOffset++) {
      const text = String(row[colOffset] || "").trim();
      if (!text) continue;

      const absoluteCol = bounds.startColIndex + colOffset;
      const absoluteRow = rowIndex;

      cells.push({
        row: absoluteRow + 1,
        col: absoluteCol + 1,
        a1: `${toA1Column(absoluteCol)}${absoluteRow + 1}`,
        text,
        link: ""
      });
    }
  }

  return {
    cells,
    source: "gviz_csv",
    sourceUrl: url,
    linkCount: 0
  };
}

function isExerciseNumber(text) {
  return /^\d+$/.test(String(text || "").trim());
}

function isListaYearLabel(text) {
  const normalized = normalizeText(text);
  return /^\d{1,2}\s*(ano|anos|serie|series)\b/.test(normalized);
}

function isListaTitle(text) {
  const raw = String(text || "").trim();
  if (!raw) return false;
  if (isExerciseNumber(raw)) return false;
  if (isListaYearLabel(raw)) return false;

  const normalized = normalizeText(raw);
  if (!/[a-z]/.test(normalized)) return false;
  if (LISTA_IGNORED_LABELS.has(normalized)) return false;

  return true;
}

function pickAnchor(anchors, target, options = {}) {
  const maxRowDistance = Number(options.maxRowDistance ?? 22);
  const maxColDistance = Number(options.maxColDistance ?? 20);
  const weightRow = Number(options.weightRow ?? 3);
  const weightCol = Number(options.weightCol ?? 1);
  const minCol = Number.isFinite(options.minCol) ? options.minCol : null;
  const maxCol = Number.isFinite(options.maxCol) ? options.maxCol : null;
  const minRow = Number.isFinite(options.minRow) ? options.minRow : null;

  let best = null;

  for (const anchor of anchors) {
    if (!anchor) continue;
    if (anchor.row > target.row || anchor.col > target.col) continue;

    if (minCol != null && anchor.col < minCol) continue;
    if (maxCol != null && anchor.col > maxCol) continue;
    if (minRow != null && anchor.row < minRow) continue;

    const rowDistance = target.row - anchor.row;
    const colDistance = target.col - anchor.col;

    if (rowDistance > maxRowDistance || colDistance > maxColDistance) continue;

    const score = rowDistance * weightRow + colDistance * weightCol;

    if (
      !best ||
      score < best.score ||
      (score === best.score && rowDistance < best.rowDistance) ||
      (score === best.score &&
        rowDistance === best.rowDistance &&
        colDistance < best.colDistance)
    ) {
      best = { anchor, score, rowDistance, colDistance };
    }
  }

  return best ? best.anchor : null;
}

function resolveYearAnchor(yearAnchors, listAnchor, target) {
  const preferred = pickAnchor(yearAnchors, target, {
    maxRowDistance: 28,
    maxColDistance: 8,
    weightRow: 4,
    weightCol: 1,
    minCol: listAnchor ? listAnchor.col - 1 : null,
    maxCol: listAnchor ? listAnchor.col + 4 : null,
    minRow: listAnchor ? listAnchor.row : null
  });

  if (preferred) return preferred;

  return pickAnchor(yearAnchors, target, {
    maxRowDistance: 28,
    maxColDistance: 20,
    weightRow: 4,
    weightCol: 1,
    minRow: listAnchor ? listAnchor.row : null
  });
}

function buildListaCatalog(cells) {
  const sortedCells = [...(cells || [])].sort(
    (a, b) => a.row - b.row || a.col - b.col
  );

  const listAnchors = sortedCells.filter((cell) => isListaTitle(cell.text));
  const yearAnchors = sortedCells.filter((cell) => isListaYearLabel(cell.text));
  const numberCells = sortedCells.filter((cell) => isExerciseNumber(cell.text));

  const groups = new Map();

  for (const cell of numberCells) {
    const listAnchor = pickAnchor(listAnchors, cell, {
      maxRowDistance: 30,
      maxColDistance: 20,
      weightRow: 3,
      weightCol: 1
    });

    if (!listAnchor) continue;

    const yearAnchor = resolveYearAnchor(yearAnchors, listAnchor, cell);
    const listName = String(listAnchor.text || "").trim();
    const yearName = String(yearAnchor?.text || "Sem ano").trim();
    const groupKey = `${normalizeText(listName)}||${normalizeText(yearName)}`;

    if (!groups.has(groupKey)) {
      groups.set(groupKey, {
        lista: listName,
        ano: yearName,
        firstRow: listAnchor.row,
        firstCol: listAnchor.col,
        items: []
      });
    }

    groups.get(groupKey).items.push({
      numero: String(cell.text || "").trim(),
      link: String(cell.link || "").trim(),
      hasLink: Boolean(String(cell.link || "").trim()),
      cell: cell.a1,
      row: cell.row,
      col: cell.col
    });
  }

  const combinations = Array.from(groups.values())
    .map((group) => {
      const uniqueByCell = new Map();
      for (const item of group.items) {
        uniqueByCell.set(item.cell, item);
      }

      const sortedItems = Array.from(uniqueByCell.values())
        .sort((a, b) => {
          const aNum = Number(a.numero);
          const bNum = Number(b.numero);

          if (Number.isFinite(aNum) && Number.isFinite(bNum) && aNum !== bNum) {
            return aNum - bNum;
          }

          return a.row - b.row || a.col - b.col;
        })
        .map(({ row, col, ...rest }) => rest);

      return {
        lista: group.lista,
        ano: group.ano,
        total: sortedItems.length,
        withLink: sortedItems.filter((item) => item.hasLink).length,
        firstRow: group.firstRow,
        firstCol: group.firstCol,
        items: sortedItems
      };
    })
    .sort(
      (a, b) =>
        a.firstRow - b.firstRow ||
        a.firstCol - b.firstCol ||
        a.lista.localeCompare(b.lista, "pt-BR")
    );

  const listas = [];
  const anosByLista = {};

  for (const combo of combinations) {
    if (!listas.includes(combo.lista)) {
      listas.push(combo.lista);
      anosByLista[combo.lista] = [];
    }

    if (!anosByLista[combo.lista].includes(combo.ano)) {
      anosByLista[combo.lista].push(combo.ano);
    }
  }

  return {
    listas,
    anosByLista,
    combinations
  };
}

function selectListaCombination(combinations, selectedList, selectedYear) {
  if (!Array.isArray(combinations) || combinations.length === 0) return null;

  const normalizedList = normalizeText(selectedList);
  const normalizedYear = normalizeText(selectedYear);

  if (normalizedList && normalizedYear) {
    const exact = combinations.find(
      (combo) =>
        normalizeText(combo.lista) === normalizedList &&
        normalizeText(combo.ano) === normalizedYear
    );
    if (exact) return exact;
  }

  if (normalizedList) {
    const listOnly = combinations.find(
      (combo) => normalizeText(combo.lista) === normalizedList
    );
    if (listOnly) return listOnly;
  }

  return combinations[0];
}

async function fetchListaData(selectedList, selectedYear) {
  const bounds = getListaColumnBounds();
  const range = `${bounds.startColLabel}:${bounds.endColLabel}`;

  let sourcePayload;
  let warning = "";

  try {
    sourcePayload = await fetchListaCellsBySheetsApi();
  } catch (apiError) {
    sourcePayload = await fetchListaCellsByGviz();
    warning =
      "Links indisponiveis sem credenciais do Google Sheets. Configure as variaveis de servico no Render.";

    if (apiError?.message) {
      warning += ` (${apiError.message})`;
    }
  }

  const catalog = buildListaCatalog(sourcePayload.cells);
  if (!catalog.combinations.length) {
    throw new Error(
      `Nenhuma lista encontrada no intervalo ${LISTA_TAB_NAME}!${range}.`
    );
  }

  const selected = selectListaCombination(
    catalog.combinations,
    selectedList,
    selectedYear
  );

  if (!selected) {
    throw new Error("Nao foi possivel determinar a lista selecionada.");
  }

  const combinationsMeta = catalog.combinations.map(
    ({ items, firstRow, firstCol, ...meta }) => meta
  );
  const hasAnyLink = combinationsMeta.some((combo) => combo.withLink > 0);

  if (!hasAnyLink && !warning) {
    warning = "Nenhum hyperlink foi encontrado nas celulas desse intervalo.";
  }

  return {
    sheet: LISTA_TAB_NAME,
    range,
    source: sourcePayload.source,
    sourceUrl: sourcePayload.sourceUrl,
    warning,
    options: {
      listas: catalog.listas,
      anosByLista: catalog.anosByLista,
      combinacoes: combinationsMeta
    },
    selected: {
      lista: selected.lista,
      ano: selected.ano
    },
    items: selected.items,
    updatedAt: new Date().toISOString()
  };
}

app.get("/api/nomes", async (_req, res) => {
  try {
    const data = await fetchNames();
    res.json(data);
  } catch (error) {
    res.status(500).json({
      error: "Falha ao carregar nomes da planilha.",
      details: error.message
    });
  }
});

app.get("/api/chamada", async (req, res) => {
  try {
    const data = await fetchChamadaData(req.query.date, req.query.turma);
    res.json(data);
  } catch (error) {
    res.status(500).json({
      error: "Falha ao carregar dados da aba Nomes.",
      details: error.message
    });
  }
});

app.get("/api/lista", async (req, res) => {
  try {
    const data = await fetchListaData(req.query.lista, req.query.ano);
    res.json(data);
  } catch (error) {
    res.status(500).json({
      error: "Falha ao carregar dados da aba Lista.",
      details: error.message
    });
  }
});

app.post("/api/chamada/marcar", async (req, res) => {
  try {
    const sheet = String(req.body?.sheet || SHEET_TAB_NAME).trim();
    const cell = String(req.body?.cell || "")
      .trim()
      .toUpperCase();
    const value = String(req.body?.value || "")
      .trim()
      .toUpperCase();

    if (!isValidA1Cell(cell)) {
      return res.status(400).json({
        error: "Celula invalida. Use formato A1 (ex.: F23)."
      });
    }

    if (sheet !== SHEET_TAB_NAME) {
      return res.status(400).json({
        error: `Apenas a aba ${SHEET_TAB_NAME} pode ser gravada.`
      });
    }

    if (!["", "F", "1", "2", "3"].includes(value)) {
      return res.status(400).json({
        error: "Valor invalido. Use 'F', '1', '2', '3' ou vazio para limpar."
      });
    }

    const sheets = getSheetsClient();
    const range = `'${sheet}'!${cell}`;

    await sheets.spreadsheets.values.update({
      spreadsheetId: GOOGLE_SPREADSHEET_ID,
      range,
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: [[value]]
      }
    });

    res.json({
      ok: true,
      sheet,
      cell,
      value
    });
  } catch (error) {
    res.status(500).json({
      error: "Falha ao gravar no Google Sheets.",
      details: error.message
    });
  }
});

app.get("/", (_req, res) => {
  res.sendFile(path.join(process.cwd(), "public", "index.html"));
});

app.listen(PORT, () => {
  console.log(`Servidor rodando na porta ${PORT}`);
});
