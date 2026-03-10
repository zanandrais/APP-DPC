# Chamada (Google Sheets + Render)

App web para:
- selecionar `data` e `turma`
- localizar a coluna/linha na aba `Nomes`
- listar alunos
- marcar/desmarcar `F`, `1`, `2` e `3` diretamente no Google Sheets
- listar exercicios da aba `Lista` (colunas `R:AZ`) por `lista` e `ano`
- abrir links dos exercicios diretamente pela interface

## Rodar localmente

```bash
npm install
npm start
```

Acesse `http://localhost:3000`.

## Deploy no Render

Use o `render.yaml` ou configure manualmente:
- Build Command: `npm install`
- Start Command: `npm start`

## Variaveis de ambiente (leitura)

- `SHEET_PUBLISH_ID`
- `SHEET_TAB_NAME` (padrao: `Nomes`)
- `SHEET_NOMES_GID` (aba `Nomes`)
- `SHEET_LISTA_TAB_NAME` (padrao: `Lista`)
- `SHEET_LISTA_COL_START` (padrao: `R`)
- `SHEET_LISTA_COL_END` (padrao: `AZ`)

## Variaveis para gravacao no Google Sheets

- `GOOGLE_SPREADSHEET_ID`
- `GOOGLE_SERVICE_ACCOUNT_EMAIL`
- `GOOGLE_PRIVATE_KEY`

## Como habilitar a gravacao

1. No Google Cloud, crie um projeto.
2. Ative a `Google Sheets API`.
3. Crie uma `Service Account`.
4. Gere uma chave JSON.
5. Compartilhe a planilha com o e-mail da Service Account como `Editor`.
6. No Render, adicione:
   - `GOOGLE_SERVICE_ACCOUNT_EMAIL` (campo `client_email` do JSON)
   - `GOOGLE_PRIVATE_KEY` (campo `private_key` do JSON)

Observacao:
- Se colar a chave no Render com `\n`, o backend ja converte automaticamente.

## Endpoints

- `GET /` interface de chamada
- `GET /api/chamada` leitura de data/turma/alunos
- `GET /api/lista` leitura da aba `Lista` (filtros opcionais: `lista`, `ano`)
- `POST /api/chamada/marcar` grava `F`, `1`, `2`, `3` ou limpa uma celula (`{ "cell": "F23", "value": "F" }`)
