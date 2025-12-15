const path = require('path');
const express = require('express');

const app = express();
const PORT = process.env.PORT || 3000;
const SHEET_ID = '2PACX-1vQL2uV2BS5DCGOlUQx4X2A7ABEWgC-c3CYA46B3S92pUG5H8VhFXta7qL00F3XjdqolkZ9jEPIqrp3Q';
const SHEET_NAME = process.env.SHEET_NAME || 'DPC';
const SHEET_GID = process.env.SHEET_GID; // mais robusto quando o nome da aba muda
const RANGE = 'A1:B5';

// Usa fetch nativo no Node 18+ ou faz fallback para node-fetch.
const fetchFn = typeof fetch === 'function'
  ? fetch
  : (...args) => import('node-fetch').then(({ default: fetch }) => fetch(...args));

const sheetUrl = () => {
  if (SHEET_GID) {
    // Link publicado com gid; usa output=csv.
    return (
      `https://docs.google.com/spreadsheets/d/e/${SHEET_ID}/pub` +
      `?gid=${encodeURIComponent(SHEET_GID)}&single=true&output=csv&range=${encodeURIComponent(RANGE)}`
    );
  }

  // Padrão usando o nome da aba.
  return (
    `https://docs.google.com/spreadsheets/d/e/${SHEET_ID}/gviz/tq` +
    `?tqx=out:csv&sheet=${encodeURIComponent(SHEET_NAME)}&range=${encodeURIComponent(RANGE)}`
  );
};

async function fetchSheet() {
  const url = sheetUrl();
  const res = await fetchFn(url);
  const csv = await res.text();
  if (!res.ok) {
    throw new Error(`Google Sheets respondeu ${res.status} - body: ${csv.slice(0, 300)}`);
  }
  // Parse simples CSV (sem virgulas internas)
  return csv
    .trim()
    .split(/\r?\n/)
    .map((row) => row.split(','));
}

app.use(express.static(path.join(__dirname, 'public')));

app.get('/api/data', async (_req, res) => {
  try {
    const data = await fetchSheet();
    res.json({ data });
  } catch (err) {
    console.error('[fetchSheet] erro:', err);
    res.status(500).json({ error: 'Falha ao buscar a planilha', detail: err.message });
  }
});

app.get('/', (_req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.listen(PORT, () => {
  console.log(`APP DPC rodando em http://localhost:${PORT}`);
});
