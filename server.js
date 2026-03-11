const path = require('path');
const express = require('express');

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
    `?tqx=out:csv&sheet=${encodeURIComponent(sheet)}&range=${encodeURIComponent(range)}`
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
  const tried = new Set();
  const candidateIds = [
    SHEET_ID,
    process.env.GOOGLE_SPREADSHEET_ID || '',
    DEFAULT_SPREADSHEET_ID
  ]
    .map((value) => String(value || '').trim())
    .filter((value) => value && !tried.has(value) && tried.add(value));

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

    res.json({
      sheet: LISTA_SHEET_NAME,
      range: RANGE_LISTA,
      source: 'gviz_csv',
      sourceUrl: payload.sourceUrl,
      warning: 'Os links so aparecem se existirem como URL publica no proprio intervalo.',
      options: {
        listas: catalog.listas,
        anosByLista: catalog.anosByLista
      },
      selected: {
        lista: selected.lista,
        ano: selected.ano
      },
      items: selected.items
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
