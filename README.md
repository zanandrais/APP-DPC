# APP DPC (Google Sheets + Render)

Aplicacao web para visualizar a planilha publicada em quatro abas:

- `DPC` (`B1:B7`)
- `Agenda` (`G7:K43`)
- `Gabarito` (`A:ZZ`, exibido na aba de listas)
- `GabaritoCB` (`A:ZZ`, filtrado por `turma` na linha 2 e `nome` na linha 1)

## Executar localmente

```bash
npm install
npm start
```

Acesse `http://localhost:3000`.

## Deploy no Render

- Build Command: `npm install`
- Start Command: `npm start`

## Variaveis de ambiente

- `SHEET_ID` (id publicado `/d/e/...`, opcional)
- `SHEET_NAME` (padrao: `DPC`)
- `SHEET_GID` (opcional, para a aba `DPC`)
- `SHEET_LISTA_NAME` (padrao: `Gabarito`)
- `SHEET_LISTA_COL_START` (padrao: `A`)
- `SHEET_LISTA_COL_END` (padrao: `ZZ`)
- `SHEET_GABARITO_CB_NAME` (padrao: `GabaritoCB`)
- `SHEET_GABARITO_CB_COL_START` (padrao: `A`)
- `SHEET_GABARITO_CB_COL_END` (padrao: `ZZ`)
- `SHEET_GABARITO_CB_MAX_ROW` (padrao: `40`)
- `GOOGLE_SPREADSHEET_ID` (id direto `/d/...`, padrao do projeto)
- `GOOGLE_SERVICE_ACCOUNT_EMAIL` (opcional, legado)
- `GOOGLE_PRIVATE_KEY` (opcional, legado)

## Endpoints

- `GET /api/data` retorna `DPC!B1:B7`
- `GET /api/agenda` retorna `DPC!G7:K43`
- `GET /api/lista` retorna leitura de `Gabarito!A:ZZ` com filtro opcional `lista`
- `GET /api/gabarito-cb` retorna leitura de `GabaritoCB!A:ZZ` com filtros opcionais `turma` e `nome`
