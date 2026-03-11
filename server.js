const path = require('path');
const express = require('express');
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
const LISTA_SHEET_NAME = process.env.SHEET_LISTA_NAME || 'Lista';

const RANGE_DPC = 'A1:B5';
const RANGE_AGENDA = 'F5:T43';
const RANGE_LISTA = 'R:AZ';
const LISTA_LINK_CACHE_TTL_MS = Number(process.env.LISTA_LINK_CACHE_TTL_MS || 5 * 60 * 1000);

const listaLinkCache = {
  at: 0,
  map: null,
  sourceUrl: ''
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
    res.status(500).json({ error: 'Falha ao buscar DPC A1:B5', detail: err.message });
  }
});

app.get('/api/agenda', async (_req, res) => {
  try {
    const payload = await fetchSheet(RANGE_AGENDA, { sheetName: SHEET_NAME, allowGid: true });
    res.json({ data: payload.rows, sourceUrl: payload.sourceUrl });
  } catch (err) {
    console.error('[api/agenda] error:', err);
    res.status(500).json({ error: 'Falha ao buscar DPC F5:T43', detail: err.message });
  }
});

app.get('/api/lista', async (req, res) => {
  try {
    const payload = await fetchSheet(RANGE_LISTA, {
      sheetName: LISTA_SHEET_NAME,
      allowGid: false
    });

    const catalog = buildListaCatalog(payload.rows, RANGE_LISTA);
    const selected = selectCombination(catalog.combinations, req.query.lista, req.query.ano);

    if (!selected) {
      return res.status(404).json({
        error: `Nenhuma lista encontrada em ${LISTA_SHEET_NAME}!${RANGE_LISTA}.`
      });
    }

    let linksWarning = '';
    let linkSourceUrl = '';
    let linkMap = {};

    try {
      const linkPayload = await fetchListaLinkMapFromXlsx();
      linkMap = linkPayload.map || {};
      linkSourceUrl = linkPayload.sourceUrl || '';
    } catch (error) {
      linksWarning = `Nao foi possivel carregar hyperlinks automaticamente (${error.message}).`;
    }

    const items = (selected.items || []).map((item) => {
      const link = String(linkMap[String(item.cell || '').toUpperCase()] || '').trim();
      return {
        ...item,
        link,
        hasLink: Boolean(link)
      };
    });

    const withLinkCount = items.filter((item) => item.hasLink).length;
    if (!linksWarning && withLinkCount === 0) {
      linksWarning = 'Nenhum hyperlink encontrado para as celulas selecionadas.';
    }

    res.json({
      sheet: LISTA_SHEET_NAME,
      range: RANGE_LISTA,
      source: 'gviz_csv + xlsx_export',
      sourceUrl: payload.sourceUrl,
      linkSourceUrl,
      warning: linksWarning,
      options: {
        listas: catalog.listas,
        anosByLista: catalog.anosByLista
      },
      selected: {
        lista: selected.lista,
        ano: selected.ano
      },
      items
    });
  } catch (err) {
    console.error('[api/lista] error:', err);
    res.status(500).json({ error: 'Falha ao buscar dados da aba Lista', detail: err.message });
  }
});

app.get('/', (_req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.listen(PORT, () => {
  console.log(`APP DPC rodando em http://localhost:${PORT}`);
});
