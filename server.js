const path = require('path');
const express = require('express');
const { google } = require('googleapis');
const XLSX = require('xlsx');

const app = express();
const PORT = process.env.PORT || 3000;

const DEFAULT_SPREADSHEET_ID = '1spXWVi4VD1wIkGVXMVdcLpm9dHAgCJ5CTCBgmpiUji8';
const SHEET_ID =
  process.env.SHEET_ID ||
  process.env.GOOGLE_SPREADSHEET_ID ||
  DEFAULT_SPREADSHEET_ID;
const SHEET_NAME = process.env.SHEET_NAME || 'DPC';
const SHEET_GID = process.env.SHEET_GID || '';
const LISTA_SHEET_NAME = process.env.SHEET_LISTA_NAME || 'Gabarito';
const LISTA_COL_START = process.env.SHEET_LISTA_COL_START || 'A';
const LISTA_COL_END = process.env.SHEET_LISTA_COL_END || 'ZZ';
const GABARITO_CB_SHEET_NAME = process.env.SHEET_GABARITO_CB_NAME || 'GabaritoCB';
const GABARITO_CB_COL_START = process.env.SHEET_GABARITO_CB_COL_START || 'A';
const GABARITO_CB_COL_END = process.env.SHEET_GABARITO_CB_COL_END || 'ZZ';
const GABARITO_CB_MAX_ROW = Number(process.env.SHEET_GABARITO_CB_MAX_ROW || 40);
const GOOGLE_SPREADSHEET_ID = process.env.GOOGLE_SPREADSHEET_ID || DEFAULT_SPREADSHEET_ID;
const GOOGLE_SERVICE_ACCOUNT_EMAIL = process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL || '';
const GOOGLE_PRIVATE_KEY = process.env.GOOGLE_PRIVATE_KEY || '';

const RANGE_DPC = 'B1:B7';
const RANGE_AGENDA = 'G7:K43';
const RANGE_LISTA = `${LISTA_COL_START}:${LISTA_COL_END}`;
const RANGE_GABARITO_CB = `${GABARITO_CB_COL_START}:${GABARITO_CB_COL_END}`;
const LISTA_LINK_CACHE_TTL_MS = Number(process.env.LISTA_LINK_CACHE_TTL_MS || 5 * 60 * 1000);

const listaLinkCache = {
  at: 0,
  map: null,
  sourceUrl: ''
};

const listaCellsCache = {
  at: 0,
  payload: null
};

// Uses native fetch on Node 18+ with a node-fetch fallback.
const fetchFn =
  typeof fetch === 'function'
    ? fetch
    : (...args) => import('node-fetch').then(({ default: fetch }) => fetch(...args));

function buildSheetBaseUrl(sheetId) {
  const id = String(sheetId || '').trim();
  if (!id) return '';

  if (id.startsWith('2PACX-')) {
    return `https://docs.google.com/spreadsheets/d/e/${id}`;
  }

  return `https://docs.google.com/spreadsheets/d/${id}`;
}

function getCandidateSheetIds() {
  const seen = new Set();
  return [SHEET_ID, process.env.GOOGLE_SPREADSHEET_ID || '', DEFAULT_SPREADSHEET_ID]
    .map((value) => String(value || '').trim())
    .filter((value) => value && !seen.has(value) && seen.add(value));
}

function getDirectSheetIds() {
  return getCandidateSheetIds().filter((value) => !value.startsWith('2PACX-'));
}

function csvUrl(range, options = {}, sheetId = SHEET_ID) {
  const sheet = options.sheetName || SHEET_NAME;
  const allowGid = options.allowGid !== false;
  const base = buildSheetBaseUrl(sheetId);
  if (!base) return '';

  if (allowGid && SHEET_GID && sheet === SHEET_NAME) {
    return (
      `${base}/pub` +
      `?gid=${encodeURIComponent(SHEET_GID)}` +
      `&single=true&output=csv&range=${encodeURIComponent(range)}`
    );
  }

  return (
    `${base}/gviz/tq` +
    `?tqx=out:csv&headers=0&sheet=${encodeURIComponent(sheet)}&range=${encodeURIComponent(range)}`
  );
}

function parseCsv(csvText) {
  const rows = [];
  let row = [];
  let field = '';
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

    if (char === ',') {
      row.push(field.trim());
      field = '';
      i += 1;
      continue;
    }

    if (char === '\r') {
      i += 1;
      continue;
    }

    if (char === '\n') {
      row.push(field.trim());
      rows.push(row);
      row = [];
      field = '';
      i += 1;
      continue;
    }

    field += char;
    i += 1;
  }

  if (field.length > 0 || row.length > 0) {
    row.push(field.trim());
    rows.push(row);
  }

  return rows;
}

async function fetchSheet(range, options = {}) {
  const candidateIds = getCandidateSheetIds();

  let lastError = new Error('No sheet id configured.');

  for (const candidateId of candidateIds) {
    const url = csvUrl(range, options, candidateId);
    if (!url) continue;

    const res = await fetchFn(url);
    const csv = await res.text();

    if (res.ok) {
      return { rows: parseCsv(csv), sourceUrl: url, sheetId: candidateId };
    }

    lastError = new Error(`Google Sheets returned ${res.status} - ${csv.slice(0, 300)}`);
  }

  throw lastError;
}

function normalizeText(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function toColIndex(colLabel) {
  const text = String(colLabel || '').trim().toUpperCase();
  if (!/^[A-Z]+$/.test(text)) return 0;

  let result = 0;
  for (const char of text) {
    result = result * 26 + (char.charCodeAt(0) - 64);
  }
  return result - 1;
}

function toA1Column(indexZeroBased) {
  let n = indexZeroBased + 1;
  let out = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    out = String.fromCharCode(65 + rem) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
}

function getRangeStartColumn(range) {
  const text = String(range || '').trim().toUpperCase();
  const match = text.match(/^([A-Z]+)(?:\d*)\s*:\s*[A-Z]+(?:\d*)$/);
  return match ? match[1] : 'A';
}

function getRangeColumns(range) {
  const text = String(range || '').trim().toUpperCase();
  const match = text.match(/^([A-Z]+)(?:\d*)\s*:\s*([A-Z]+)(?:\d*)$/);

  if (!match) {
    return { start: 0, end: Number.MAX_SAFE_INTEGER };
  }

  return {
    start: toColIndex(match[1]),
    end: toColIndex(match[2])
  };
}

function fromA1Column(colLabel) {
  const value = String(colLabel || '').trim().toUpperCase();
  if (!/^[A-Z]+$/.test(value)) return -1;

  let total = 0;
  for (const char of value) {
    total = total * 26 + (char.charCodeAt(0) - 64);
  }

  return total - 1;
}

function getSheetColumnBounds(startLabel, endLabel) {
  const startCol = fromA1Column(startLabel);
  const endCol = fromA1Column(endLabel);

  if (startCol < 0 || endCol < 0) {
    return {
      startColIndex: fromA1Column('A'),
      endColIndex: fromA1Column('ZZ'),
      startColLabel: 'A',
      endColLabel: 'ZZ'
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

function getListaColumnBounds() {
  return getSheetColumnBounds(LISTA_COL_START, LISTA_COL_END);
}

function getGabaritoCbColumnBounds() {
  return getSheetColumnBounds(GABARITO_CB_COL_START, GABARITO_CB_COL_END);
}

function getSheetsClient() {
  if (!GOOGLE_SERVICE_ACCOUNT_EMAIL || !GOOGLE_PRIVATE_KEY || !GOOGLE_SPREADSHEET_ID) {
    throw new Error(
      'Credenciais do Google Sheets nao configuradas. Configure GOOGLE_SERVICE_ACCOUNT_EMAIL, GOOGLE_PRIVATE_KEY e GOOGLE_SPREADSHEET_ID.'
    );
  }

  const auth = new google.auth.JWT(
    GOOGLE_SERVICE_ACCOUNT_EMAIL,
    null,
    GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n'),
    ['https://www.googleapis.com/auth/spreadsheets']
  );

  return google.sheets({ version: 'v4', auth });
}

function normalizeListaCellText(cell) {
  const formatted = String(cell?.formattedValue || '').trim();
  if (formatted) return formatted;

  if (cell?.effectiveValue?.stringValue != null) {
    return String(cell.effectiveValue.stringValue).trim();
  }

  if (cell?.effectiveValue?.numberValue != null) {
    return String(cell.effectiveValue.numberValue).trim();
  }

  if (cell?.effectiveValue?.boolValue != null) {
    return cell.effectiveValue.boolValue ? 'TRUE' : 'FALSE';
  }

  if (cell?.userEnteredValue?.stringValue != null) {
    return String(cell.userEnteredValue.stringValue).trim();
  }

  if (cell?.userEnteredValue?.numberValue != null) {
    return String(cell.userEnteredValue.numberValue).trim();
  }

  return '';
}

function extractHyperlinkFromFormula(formulaValue) {
  const formula = String(formulaValue || '').trim();
  if (!formula) return '';

  const directMatch = formula.match(/HYPERLINK\s*\(\s*"([^"]+)"/i);
  if (!directMatch) return '';

  return String(directMatch[1] || '').trim();
}

function normalizeListaHyperlink(link) {
  const text = String(link || '').trim();
  if (!text) return '';
  if (/^https?:\/\//i.test(text)) return text;
  if (/^www\./i.test(text)) return `https://${text}`;
  return '';
}

function extractListaCellLink(cell) {
  const direct = normalizeListaHyperlink(cell?.hyperlink);
  if (direct) return direct;

  const richTextLink = normalizeListaHyperlink(
    cell?.userEnteredFormat?.textFormat?.link?.uri || cell?.effectiveFormat?.textFormat?.link?.uri
  );
  if (richTextLink) return richTextLink;

  if (Array.isArray(cell?.textFormatRuns)) {
    for (const run of cell.textFormatRuns) {
      const link = normalizeListaHyperlink(run?.format?.link?.uri);
      if (link) return link;
    }
  }

  return '';
}

async function fetchCellsBySheetsApi(sheetName, bounds, rangeA1 = `${bounds.startColLabel}:${bounds.endColLabel}`) {
  const sheets = getSheetsClient();
  const a1Range = `'${sheetName}'!${rangeA1}`;

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

  for (let rowOffset = 0; rowOffset < rowData.length; rowOffset += 1) {
    const values = Array.isArray(rowData[rowOffset]?.values) ? rowData[rowOffset].values : [];

    for (let colOffset = 0; colOffset < values.length; colOffset += 1) {
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
    source: 'google_sheets_api',
    sourceUrl: a1Range,
    linkCount
  };
}

async function fetchListaCellsBySheetsApi() {
  return fetchCellsBySheetsApi(LISTA_SHEET_NAME, getListaColumnBounds());
}

async function fetchListaCellsByXlsx() {
  const now = Date.now();
  if (listaCellsCache.payload && now - listaCellsCache.at < LISTA_LINK_CACHE_TTL_MS) {
    return listaCellsCache.payload;
  }

  const directIds = getDirectSheetIds();
  if (!directIds.length) {
    throw new Error('Nenhum id direto da planilha foi configurado para exportar XLSX.');
  }

  const bounds = getListaColumnBounds();
  let lastError = new Error('Falha ao carregar dados do XLSX.');

  for (const sheetId of directIds) {
    const exportUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=xlsx`;

    try {
      const response = await fetchFn(exportUrl);
      if (!response.ok) {
        lastError = new Error(`XLSX export retornou ${response.status}.`);
        continue;
      }

      const buffer = Buffer.from(await response.arrayBuffer());
      const workbook = XLSX.read(buffer, {
        type: 'buffer',
        cellFormula: true,
        cellText: true
      });

      const targetSheetName =
        workbook.SheetNames.find((name) => normalizeText(name) === normalizeText(LISTA_SHEET_NAME)) ||
        LISTA_SHEET_NAME;
      const sheet = getWorkbookSheet(workbook, targetSheetName);

      if (!sheet) {
        lastError = new Error(`Aba ${LISTA_SHEET_NAME} nao encontrada no XLSX.`);
        continue;
      }

      const cells = [];
      let linkCount = 0;

      for (const [address, cell] of Object.entries(sheet)) {
        if (address.startsWith('!')) continue;

        const decoded = XLSX.utils.decode_cell(address);
        if (decoded.c < bounds.startColIndex || decoded.c > bounds.endColIndex) continue;

        const text = String(getWorkbookCellText(workbook, targetSheetName, address) || '').trim();
        let link = normalizeUrl(cell?.l?.Target);

        if (!link && cell?.f) {
          link = extractLinkFromFormula(cell.f, workbook, targetSheetName);
        }

        if (!text && !link) continue;
        if (link) linkCount += 1;

        cells.push({
          row: decoded.r + 1,
          col: decoded.c + 1,
          a1: address.toUpperCase(),
          text,
          link
        });
      }

      const payload = {
        cells,
        source: 'xlsx_export',
        sourceUrl: exportUrl,
        linkCount
      };

      listaCellsCache.at = now;
      listaCellsCache.payload = payload;
      return payload;
    } catch (error) {
      lastError = error;
    }
  }

  throw lastError;
}

async function fetchCellsByGviz(sheetName, bounds, rangeA1 = `${bounds.startColLabel}:${bounds.endColLabel}`) {
  const payload = await fetchSheet(rangeA1, {
    sheetName,
    allowGid: false
  });

  const rows = payload.rows || [];
  const cells = [];

  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    const row = Array.isArray(rows[rowIndex]) ? rows[rowIndex] : [];
    for (let colOffset = 0; colOffset < row.length; colOffset += 1) {
      const text = String(row[colOffset] || '').trim();
      if (!text) continue;

      const absoluteCol = bounds.startColIndex + colOffset;
      const absoluteRow = rowIndex;

      cells.push({
        row: absoluteRow + 1,
        col: absoluteCol + 1,
        a1: `${toA1Column(absoluteCol)}${absoluteRow + 1}`,
        text,
        link: ''
      });
    }
  }

  return {
    cells,
    source: 'gviz_csv',
    sourceUrl: payload.sourceUrl,
    linkCount: 0
  };
}

async function fetchListaCellsByGviz() {
  return fetchCellsByGviz(LISTA_SHEET_NAME, getListaColumnBounds());
}

async function fetchGabaritoCbCellsBySheetsApi() {
  const bounds = getGabaritoCbColumnBounds();
  const rangeA1 = `${bounds.startColLabel}1:${bounds.endColLabel}${GABARITO_CB_MAX_ROW}`;
  return fetchCellsBySheetsApi(GABARITO_CB_SHEET_NAME, bounds, rangeA1);
}

async function fetchGabaritoCbCellsByGviz() {
  const bounds = getGabaritoCbColumnBounds();
  const rangeA1 = `${bounds.startColLabel}1:${bounds.endColLabel}${GABARITO_CB_MAX_ROW}`;
  return fetchCellsByGviz(GABARITO_CB_SHEET_NAME, bounds, rangeA1);
}

function isExerciseNumber(text) {
  return /^\d+$/.test(String(text || '').trim());
}

function isListaYearLabel(text) {
  const normalized = normalizeText(text);
  return /^\d{1,2}\s*(ano|anos|serie|series)\b/.test(normalized);
}

function isListaTitle(text) {
  const raw = String(text || '').trim();
  if (!raw) return false;
  if (isExerciseNumber(raw)) return false;
  if (isListaYearLabel(raw)) return false;

  const normalized = normalizeText(raw);
  if (!/[a-z]/.test(normalized)) return false;
  if (normalized === 'gabaritos') return false;
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

function buildListaCatalogFromCells(cells) {
  const byColumn = new Map();

  for (const cell of cells || []) {
    if (!cell || !Number.isFinite(cell.col)) continue;
    if (!byColumn.has(cell.col)) byColumn.set(cell.col, []);
    byColumn.get(cell.col).push(cell);
  }

  const groups = new Map();

  for (const columnCells of byColumn.values()) {
    const sortedColumn = [...columnCells].sort((a, b) => a.row - b.row);
    const header = sortedColumn.find((cell) => cell.row === 1 && isListaTitle(cell.text));
    if (!header) continue;

    const listName = String(header.text || '').trim();
    const key = normalizeText(listName);
    if (!key) continue;

    if (!groups.has(key)) {
      groups.set(key, {
        lista: listName,
        firstRow: header.row,
        firstCol: header.col,
        items: []
      });
    }

    const target = groups.get(key);

    for (const cell of sortedColumn) {
      if (cell.row <= header.row) continue;

      const numero = String(cell.text || '').trim();
      if (!isExerciseNumber(numero)) continue;

      target.items.push({
        numero,
        link: String(cell.link || '').trim(),
        hasLink: Boolean(String(cell.link || '').trim()),
        cell: cell.a1,
        row: cell.row,
        col: cell.col
      });
    }
  }

  const combinations = Array.from(groups.values())
    .map((group) => {
      const uniqueByCell = new Map();
      for (const item of group.items) uniqueByCell.set(item.cell, item);

      const sortedItems = Array.from(uniqueByCell.values())
        .sort((a, b) => {
          const aNum = Number(a.numero);
          const bNum = Number(b.numero);
          if (Number.isFinite(aNum) && Number.isFinite(bNum) && aNum !== bNum) return aNum - bNum;
          return a.row - b.row || a.col - b.col;
        })
        .map(({ row, col, ...rest }) => rest);

      return {
        lista: group.lista,
        total: sortedItems.length,
        withLink: sortedItems.filter((item) => item.hasLink).length,
        firstRow: group.firstRow,
        firstCol: group.firstCol,
        items: sortedItems
      };
    })
    .filter((group) => group.items.length > 0)
    .sort(
      (a, b) =>
        a.firstRow - b.firstRow || a.firstCol - b.firstCol || a.lista.localeCompare(b.lista, 'pt-BR')
    );

  return {
    listas: combinations.map((group) => group.lista),
    combinations
  };
}

function selectListaCombination(combinations, selectedList) {
  if (!Array.isArray(combinations) || combinations.length === 0) return null;

  const listKey = normalizeText(selectedList);

  if (listKey) {
    const byList = combinations.find((combo) => normalizeText(combo.lista) === listKey);
    if (byList) return byList;
  }

  return combinations[0];
}

async function fetchListaData(selectedList) {
  const bounds = getListaColumnBounds();
  const range = `${bounds.startColLabel}:${bounds.endColLabel}`;

  let sourcePayload;
  let warning = '';

  try {
    sourcePayload = await fetchListaCellsBySheetsApi();
  } catch (apiError) {
    sourcePayload = await fetchListaCellsByGviz();
    warning = 'Nao foi possivel carregar os hyperlinks automaticamente a partir da aba Gabarito.';
    if (apiError?.message) {
      warning += ` (${apiError.message})`;
    }
  }

  const catalog = buildListaCatalogFromCells(sourcePayload.cells);
  if (!catalog.combinations.length) {
    throw new Error(`Nenhuma lista encontrada no intervalo ${LISTA_SHEET_NAME}!${range}.`);
  }

  const selected = selectListaCombination(catalog.combinations, selectedList);
  if (!selected) {
    throw new Error('Nao foi possivel determinar a lista selecionada.');
  }

  const combinationsMeta = catalog.combinations.map(({ items, firstRow, firstCol, ...meta }) => meta);
  const selectedWithLink = selected.items.filter((item) => item.hasLink).length;

  if (!warning && selectedWithLink === 0) {
    warning = 'Nenhum hyperlink foi encontrado para essa lista.';
  }

  return {
    sheet: LISTA_SHEET_NAME,
    range,
    source: sourcePayload.source,
    sourceUrl: sourcePayload.sourceUrl,
    warning,
    options: {
      listas: catalog.listas,
      combinacoes: combinationsMeta
    },
    selected: {
      lista: selected.lista
    },
    items: selected.items,
    updatedAt: new Date().toISOString()
  };
}

function buildGabaritoCbCatalogFromCells(cells) {
  const byColumn = new Map();

  for (const cell of cells || []) {
    if (!cell || !Number.isFinite(cell.col)) continue;
    if (!byColumn.has(cell.col)) byColumn.set(cell.col, []);
    byColumn.get(cell.col).push(cell);
  }

  const columns = [];

  for (const [col, columnCells] of Array.from(byColumn.entries()).sort((a, b) => a[0] - b[0])) {
    const sortedColumn = [...columnCells].sort((a, b) => a.row - b.row);
    const nomeCell = sortedColumn.find((cell) => cell.row === 1 && String(cell.text || '').trim());
    const turmaCell = sortedColumn.find((cell) => cell.row === 2 && String(cell.text || '').trim());

    if (!nomeCell || !turmaCell) continue;

    const nome = String(nomeCell.text || '').trim();
    const turma = String(turmaCell.text || '').trim();
    const uniqueByCell = new Map();

    for (const cell of sortedColumn) {
      if (cell.row <= 2) continue;

      const numero = String(cell.text || '').trim();
      if (!isExerciseNumber(numero)) continue;

      uniqueByCell.set(cell.a1, {
        numero,
        link: String(cell.link || '').trim(),
        hasLink: Boolean(String(cell.link || '').trim()),
        cell: cell.a1,
        row: cell.row,
        col: cell.col
      });
    }

    const items = Array.from(uniqueByCell.values())
      .sort((a, b) => {
        const aNum = Number(a.numero);
        const bNum = Number(b.numero);
        if (a.row !== b.row) return a.row - b.row;
        if (a.col !== b.col) return a.col - b.col;
        if (Number.isFinite(aNum) && Number.isFinite(bNum)) return aNum - bNum;
        return String(a.numero).localeCompare(String(b.numero), 'pt-BR');
      })
      .map(({ row, col: itemCol, ...rest }) => rest);

    columns.push({
      turma,
      nome,
      col,
      a1Column: toA1Column(col - 1),
      items,
      withLink: items.filter((item) => item.hasLink).length
    });
  }

  const turmas = [];
  const nomesByTurma = {};

  for (const column of columns) {
    if (!turmas.some((turma) => normalizeText(turma) === normalizeText(column.turma))) {
      turmas.push(column.turma);
    }

    if (!nomesByTurma[column.turma]) {
      nomesByTurma[column.turma] = [];
    }

    if (
      !nomesByTurma[column.turma].some((nome) => normalizeText(nome) === normalizeText(column.nome))
    ) {
      nomesByTurma[column.turma].push(column.nome);
    }
  }

  return {
    turmas,
    nomesByTurma,
    columns
  };
}

function selectGabaritoCbColumn(columns, selectedTurma, selectedNome) {
  if (!Array.isArray(columns) || !columns.length) return null;

  const turmaKey = normalizeText(selectedTurma);
  const nomeKey = normalizeText(selectedNome);

  const columnsByTurma = turmaKey
    ? columns.filter((column) => normalizeText(column.turma) === turmaKey)
    : columns;
  const scopedColumns = columnsByTurma.length ? columnsByTurma : columns;

  if (nomeKey) {
    const byNome = scopedColumns.find((column) => normalizeText(column.nome) === nomeKey);
    if (byNome) return byNome;
  }

  return scopedColumns[0];
}

async function fetchGabaritoCbData(selectedTurma, selectedNome) {
  const bounds = getGabaritoCbColumnBounds();
  const range = `${bounds.startColLabel}1:${bounds.endColLabel}${GABARITO_CB_MAX_ROW}`;

  let sourcePayload;
  let warning = '';

  try {
    sourcePayload = await fetchGabaritoCbCellsBySheetsApi();
  } catch (apiError) {
    sourcePayload = await fetchGabaritoCbCellsByGviz();
    warning = 'Nao foi possivel carregar os hyperlinks automaticamente a partir da aba GabaritoCB.';
    if (apiError?.message) {
      warning += ` (${apiError.message})`;
    }
  }

  const catalog = buildGabaritoCbCatalogFromCells(sourcePayload.cells);
  if (!catalog.columns.length) {
    throw new Error(`Nenhuma coluna com nome e turma foi encontrada em ${GABARITO_CB_SHEET_NAME}!${range}.`);
  }

  const selected = selectGabaritoCbColumn(catalog.columns, selectedTurma, selectedNome);
  if (!selected) {
    throw new Error('Nao foi possivel determinar a coluna selecionada na aba GabaritoCB.');
  }

  if (!warning && selected.withLink === 0) {
    warning = 'Nenhum hyperlink foi encontrado para o nome selecionado.';
  }

  return {
    sheet: GABARITO_CB_SHEET_NAME,
    range,
    source: sourcePayload.source,
    sourceUrl: sourcePayload.sourceUrl,
    warning,
    options: {
      turmas: catalog.turmas,
      nomesByTurma: catalog.nomesByTurma
    },
    selected: {
      turma: selected.turma,
      nome: selected.nome,
      coluna: selected.a1Column
    },
    items: selected.items,
    updatedAt: new Date().toISOString()
  };
}

function normalizeUrl(value) {
  const text = String(value || '').trim();
  if (!text) return '';
  if (/^https?:\/\//i.test(text)) return text;
  if (/^www\./i.test(text)) return `https://${text}`;
  return '';
}

function splitFormulaArgs(text) {
  const args = [];
  let current = '';
  let depth = 0;
  let inString = false;

  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i];
    const next = text[i + 1];

    if (inString) {
      current += ch;
      if (ch === '"' && next === '"') {
        current += next;
        i += 1;
      } else if (ch === '"') {
        inString = false;
      }
      continue;
    }

    if (ch === '"') {
      inString = true;
      current += ch;
      continue;
    }

    if (ch === '(') {
      depth += 1;
      current += ch;
      continue;
    }

    if (ch === ')') {
      depth = Math.max(0, depth - 1);
      current += ch;
      continue;
    }

    if ((ch === ',' || ch === ';') && depth === 0) {
      args.push(current.trim());
      current = '';
      continue;
    }

    current += ch;
  }

  if (current.trim() || args.length) {
    args.push(current.trim());
  }

  return args;
}

function splitFormulaConcat(text) {
  const out = [];
  let current = '';
  let depth = 0;
  let inString = false;

  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i];
    const next = text[i + 1];

    if (inString) {
      current += ch;
      if (ch === '"' && next === '"') {
        current += next;
        i += 1;
      } else if (ch === '"') {
        inString = false;
      }
      continue;
    }

    if (ch === '"') {
      inString = true;
      current += ch;
      continue;
    }

    if (ch === '(') {
      depth += 1;
      current += ch;
      continue;
    }

    if (ch === ')') {
      depth = Math.max(0, depth - 1);
      current += ch;
      continue;
    }

    if (ch === '&' && depth === 0) {
      out.push(current.trim());
      current = '';
      continue;
    }

    current += ch;
  }

  if (current.trim() || out.length) {
    out.push(current.trim());
  }

  return out;
}

function parseA1Reference(text, defaultSheetName) {
  const raw = String(text || '').trim();
  const match = raw.match(/^(?:(?:'([^']+)'|([^!]+))!)?(\$?[A-Z]{1,3}\$?\d+)$/i);
  if (!match) return null;

  const sheetName = String(match[1] || match[2] || defaultSheetName || '').trim();
  const cell = match[3].replace(/\$/g, '').toUpperCase();
  if (!sheetName || !cell) return null;

  return { sheetName, cell };
}

function getWorkbookSheet(workbook, sheetName) {
  if (!workbook || !Array.isArray(workbook.SheetNames)) return null;
  const exact = workbook.Sheets[sheetName];
  if (exact) return exact;

  const normalized = normalizeText(sheetName);
  const foundName = workbook.SheetNames.find((name) => normalizeText(name) === normalized);
  return foundName ? workbook.Sheets[foundName] : null;
}

function getWorkbookCellText(workbook, sheetName, cellA1) {
  const sheet = getWorkbookSheet(workbook, sheetName);
  if (!sheet) return '';

  const cell = sheet[cellA1];
  if (!cell) return '';

  if (typeof cell.w === 'string' && cell.w.trim()) return cell.w.trim();
  if (cell.v != null) return String(cell.v).trim();
  if (cell.f) return evalFormulaText(cell.f, workbook, sheetName, 1).trim();
  return '';
}

function evalFormulaNumber(expr, workbook, defaultSheetName) {
  const text = String(expr || '').trim();
  if (!text) return NaN;
  if (/^-?\d+(?:\.\d+)?$/.test(text)) return Number(text);

  const ref = parseA1Reference(text, defaultSheetName);
  if (ref) {
    const rawValue = getWorkbookCellText(workbook, ref.sheetName, ref.cell);
    const normalized = rawValue.replace(',', '.');
    const value = Number(normalized);
    return Number.isFinite(value) ? value : NaN;
  }

  return NaN;
}

function evalFormulaText(expr, workbook, defaultSheetName, depth = 0) {
  if (depth > 8) return '';
  const text = String(expr || '').trim().replace(/^=/, '');
  if (!text) return '';

  if (text.startsWith('"') && text.endsWith('"') && text.length >= 2) {
    return text.slice(1, -1).replace(/""/g, '"');
  }

  const concatParts = splitFormulaConcat(text);
  if (concatParts.length > 1) {
    return concatParts
      .map((part) => evalFormulaText(part, workbook, defaultSheetName, depth + 1))
      .join('');
  }

  const midMatch = text.match(/^(?:MID|EXT\.?TEXTO)\(([\s\S]*)\)$/i);
  if (midMatch) {
    const args = splitFormulaArgs(midMatch[1]);
    if (args.length >= 3) {
      const source = evalFormulaText(args[0], workbook, defaultSheetName, depth + 1);
      const start = Math.trunc(evalFormulaNumber(args[1], workbook, defaultSheetName));
      const len = Math.trunc(evalFormulaNumber(args[2], workbook, defaultSheetName));

      if (!source || !Number.isFinite(start) || !Number.isFinite(len) || start <= 0 || len <= 0) {
        return '';
      }

      return source.slice(start - 1, start - 1 + len);
    }
  }

  const ref = parseA1Reference(text, defaultSheetName);
  if (ref) {
    return getWorkbookCellText(workbook, ref.sheetName, ref.cell);
  }

  return '';
}

function extractLinkFromFormula(formula, workbook, sheetName) {
  const raw = String(formula || '').trim().replace(/^=/, '');
  if (!raw) return '';

  const upper = raw.toUpperCase();
  const fnNames = ['HYPERLINK', 'HIPERLINK'];
  let argsText = '';

  outer: for (let i = 0; i < raw.length; i += 1) {
    for (const fnName of fnNames) {
      if (!upper.startsWith(fnName, i)) continue;

      const prevChar = i > 0 ? upper[i - 1] : '';
      if (/[A-Z0-9_.]/.test(prevChar)) continue;

      let openAt = i + fnName.length;
      while (raw[openAt] === ' ' || raw[openAt] === '\t') openAt += 1;
      if (raw[openAt] !== '(') continue;

      let depth = 0;
      let inString = false;
      for (let j = openAt; j < raw.length; j += 1) {
        const ch = raw[j];
        const next = raw[j + 1];

        if (inString) {
          if (ch === '"' && next === '"') {
            j += 1;
          } else if (ch === '"') {
            inString = false;
          }
          continue;
        }

        if (ch === '"') {
          inString = true;
          continue;
        }

        if (ch === '(') {
          depth += 1;
          continue;
        }

        if (ch === ')') {
          depth -= 1;
          if (depth === 0) {
            argsText = raw.slice(openAt + 1, j);
            break outer;
          }
        }
      }
    }
  }

  if (!argsText) {
    const directInFormula = raw.match(/https?:\/\/[^\s)"',;]+/i);
    return directInFormula ? normalizeUrl(directInFormula[0]) : '';
  }

  const args = splitFormulaArgs(argsText);
  if (!args.length) return '';

  const resolved = evalFormulaText(args[0], workbook, sheetName);
  const normalized = normalizeUrl(resolved);
  if (normalized) return normalized;

  const directInArg = String(args[0] || '').match(/https?:\/\/[^\s)"',;]+/i);
  if (directInArg) {
    return normalizeUrl(directInArg[0]);
  }

  const directInFormula = raw.match(/https?:\/\/[^\s)"',;]+/i);
  return directInFormula ? normalizeUrl(directInFormula[0]) : '';
}

async function fetchListaLinkMapFromXlsx() {
  const now = Date.now();
  if (listaLinkCache.map && now - listaLinkCache.at < LISTA_LINK_CACHE_TTL_MS) {
    return { map: listaLinkCache.map, sourceUrl: listaLinkCache.sourceUrl };
  }

  const directIds = getDirectSheetIds();
  if (!directIds.length) {
    throw new Error('Nenhum id direto da planilha foi configurado para exportar XLSX.');
  }

  const { start, end } = getRangeColumns(RANGE_LISTA);
  let lastError = new Error('Falha ao carregar links do XLSX.');

  for (const sheetId of directIds) {
    const exportUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=xlsx`;
    try {
      const response = await fetchFn(exportUrl);
      if (!response.ok) {
        lastError = new Error(`XLSX export retornou ${response.status}.`);
        continue;
      }

      const buffer = Buffer.from(await response.arrayBuffer());
      const workbook = XLSX.read(buffer, {
        type: 'buffer',
        cellFormula: true,
        cellText: true
      });

      const targetSheetName =
        workbook.SheetNames.find((name) => normalizeText(name) === normalizeText(LISTA_SHEET_NAME)) ||
        LISTA_SHEET_NAME;
      const sheet = getWorkbookSheet(workbook, targetSheetName);

      if (!sheet) {
        lastError = new Error(`Aba ${LISTA_SHEET_NAME} nao encontrada no XLSX.`);
        continue;
      }

      const linkMap = {};
      for (const [address, cell] of Object.entries(sheet)) {
        if (address.startsWith('!')) continue;

        const decoded = XLSX.utils.decode_cell(address);
        if (decoded.c < start || decoded.c > end) continue;

        let link = normalizeUrl(cell?.l?.Target);
        if (!link && cell?.f) {
          link = extractLinkFromFormula(cell.f, workbook, targetSheetName);
        }

        if (!link) {
          link = normalizeUrl(cell?.w || cell?.v);
        }

        if (!link) continue;
        linkMap[address.toUpperCase()] = link;
      }

      listaLinkCache.at = now;
      listaLinkCache.map = linkMap;
      listaLinkCache.sourceUrl = exportUrl;

      return { map: linkMap, sourceUrl: exportUrl };
    } catch (error) {
      lastError = error;
    }
  }

  throw lastError;
}

function extractNumbersFromCell(text) {
  const raw = String(text || '').trim();
  if (!raw) return [];
  if (!/^[\d,\s]+$/.test(raw)) return [];
  return raw.match(/\d+/g) || [];
}

function isYearLabel(text) {
  return /^\d{1,2}\s*(ano|anos|serie|series)\b/.test(normalizeText(text));
}

function isTitle(text) {
  const raw = String(text || '').trim();
  if (!raw) return false;
  if (extractNumbersFromCell(raw).length > 0 && /^[\d,\s]+$/.test(raw)) return false;
  if (isYearLabel(raw)) return false;

  const norm = normalizeText(raw);
  if (norm === 'gabaritos') return false;
  return /[a-z]/.test(norm);
}

function toMatrix(rows) {
  const width = Math.max(1, ...(rows || []).map((row) => row.length));
  return (rows || []).map((row) =>
    Array.from({ length: width }, (_unused, idx) => String(row[idx] || '').trim())
  );
}

function buildListaCatalog(rows, range) {
  const matrix = toMatrix(rows);
  const height = matrix.length;
  const width = matrix[0] ? matrix[0].length : 0;
  const startCol = toColIndex(getRangeStartColumn(range));

  const titlesByRow = Array.from({ length: height }, () => []);

  for (let r = 0; r < height; r += 1) {
    for (let c = 0; c < width; c += 1) {
      const value = matrix[r][c];
      if (isTitle(value)) {
        titlesByRow[r].push({ row: r, col: c, text: value });
      }
    }
    titlesByRow[r].sort((a, b) => a.col - b.col);
  }

  const groups = new Map();

  for (let r = 0; r < height; r += 1) {
    const rowTitles = titlesByRow[r];
    if (!rowTitles.length) continue;

    for (let i = 0; i < rowTitles.length; i += 1) {
      const anchor = rowTitles[i];
      const start = anchor.col;
      const end = i + 1 < rowTitles.length ? rowTitles[i + 1].col : width;

      let endRow = height;
      for (let rr = r + 1; rr < height; rr += 1) {
        const hasNextTitleInRange = titlesByRow[rr].some(
          (title) => title.col >= start && title.col < end
        );
        if (hasNextTitleInRange) {
          endRow = rr;
          break;
        }
      }

      let yearName = 'Sem ano';
      let yearRow = -1;
      const yearSearchEnd = Math.min(endRow, r + 4);
      for (let rr = r + 1; rr < yearSearchEnd; rr += 1) {
        let found = false;
        for (let cc = start; cc < end; cc += 1) {
          const candidate = matrix[rr][cc];
          if (isYearLabel(candidate)) {
            yearName = candidate;
            yearRow = rr;
            found = true;
            break;
          }
        }
        if (found) break;
      }

      const fromRow = yearRow >= 0 ? yearRow + 1 : r + 1;
      const key = `${normalizeText(anchor.text)}||${normalizeText(yearName)}`;

      if (!groups.has(key)) {
        groups.set(key, {
          lista: anchor.text,
          ano: yearName,
          items: []
        });
      }

      const target = groups.get(key);

      for (let rr = fromRow; rr < endRow; rr += 1) {
        for (let cc = start; cc < end; cc += 1) {
          const cell = matrix[rr][cc];
          const tokens = extractNumbersFromCell(cell);
          if (!tokens.length) continue;

          const a1 = `${toA1Column(startCol + cc)}${rr + 1}`;
          for (const token of tokens) {
            target.items.push({
              numero: token,
              link: '',
              hasLink: false,
              cell: a1
            });
          }
        }
      }
    }
  }

  const combinations = Array.from(groups.values())
    .map((group) => {
      const seen = new Set();
      const deduped = [];

      for (const item of group.items) {
        const key = `${item.numero}|${item.cell}`;
        if (seen.has(key)) continue;
        seen.add(key);
        deduped.push(item);
      }

      deduped.sort((a, b) => Number(a.numero) - Number(b.numero));

      return {
        lista: group.lista,
        ano: group.ano,
        total: deduped.length,
        withLink: 0,
        items: deduped
      };
    })
    .filter((combo) => combo.items.length > 0)
    .sort((a, b) => a.lista.localeCompare(b.lista, 'pt-BR') || a.ano.localeCompare(b.ano, 'pt-BR'));

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

  return { listas, anosByLista, combinations };
}

function selectCombination(combinations, selectedList, selectedYear) {
  if (!Array.isArray(combinations) || !combinations.length) return null;

  const listKey = normalizeText(selectedList);
  const yearKey = normalizeText(selectedYear);

  if (listKey && yearKey) {
    const exact = combinations.find(
      (combo) => normalizeText(combo.lista) === listKey && normalizeText(combo.ano) === yearKey
    );
    if (exact) return exact;
  }

  if (listKey) {
    const byList = combinations.find((combo) => normalizeText(combo.lista) === listKey);
    if (byList) return byList;
  }

  return combinations[0];
}

app.use(express.static(path.join(__dirname, 'public')));

app.get('/api/data', async (_req, res) => {
  try {
    const payload = await fetchSheet(RANGE_DPC, { sheetName: SHEET_NAME, allowGid: true });
    res.json({ data: payload.rows, sourceUrl: payload.sourceUrl });
  } catch (err) {
    console.error('[api/data] error:', err);
    res.status(500).json({ error: 'Falha ao buscar DPC B1:B7', detail: err.message });
  }
});

app.get('/api/agenda', async (_req, res) => {
  try {
    const payload = await fetchSheet(RANGE_AGENDA, { sheetName: SHEET_NAME, allowGid: true });
    res.json({ data: payload.rows, sourceUrl: payload.sourceUrl });
  } catch (err) {
    console.error('[api/agenda] error:', err);
    res.status(500).json({ error: 'Falha ao buscar DPC G7:K43', detail: err.message });
  }
});

app.get('/api/lista', async (req, res) => {
  try {
    const data = await fetchListaData(req.query.lista);
    res.json(data);
  } catch (err) {
    console.error('[api/lista] error:', err);
    res.status(500).json({ error: 'Falha ao buscar dados da aba Gabarito', detail: err.message });
  }
});

app.get('/api/gabarito-cb', async (req, res) => {
  try {
    const data = await fetchGabaritoCbData(req.query.turma, req.query.nome);
    res.json(data);
  } catch (err) {
    console.error('[api/gabarito-cb] error:', err);
    res.status(500).json({ error: 'Falha ao buscar dados da aba GabaritoCB', detail: err.message });
  }
});

app.get('/', (_req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.listen(PORT, () => {
  console.log(`APP DPC rodando em http://localhost:${PORT}`);
});
