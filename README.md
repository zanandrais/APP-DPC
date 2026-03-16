# APP DPC (Google Sheets + Render)

Aplicacao web para visualizar a planilha publicada em tres abas:

- `DPC` (`A1:B5`)
- `Agenda` (`F5:T43`)
- `Gabarito` (`A:ZZ`, exibido na aba de listas)

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
- `GOOGLE_SPREADSHEET_ID` (id direto `/d/...`, padrao do projeto)
- `GOOGLE_SERVICE_ACCOUNT_EMAIL` (opcional, legado)
- `GOOGLE_PRIVATE_KEY` (opcional, legado)

## Endpoints

- `GET /api/data` retorna `DPC!A1:B5`
- `GET /api/agenda` retorna `DPC!F5:T43`
- `GET /api/lista` retorna leitura de `Gabarito!A:ZZ` com filtro opcional `lista`
