# APP DPC

Webservice em Node/Express para exibir:
- `DPC!A1:B5` (aba DPC)
- `DPC!F5:T43` (aba Agenda)

## Rodar localmente

```bash
npm install
npm start
# abre http://localhost:3000
```

- A porta pode ser definida via variavel `PORT`.
- Se a aba tiver outro nome, ajuste `SHEET_NAME` (padrao: `DPC`).
- Opcional: use `SHEET_GID` para leitura via `gid`.

## Deploy no Render

1. Crie um novo Web Service apontando para este repositorio.
2. Build Command: `npm install`
3. Start Command: `npm start`
4. Runtime Node 18+.
5. Opcional: ajuste `SHEET_NAME` e `SHEET_GID` nas variaveis de ambiente.

O arquivo `render.yaml` ja descreve esse servico.

## Endpoints

- `GET /api/data` retorna `DPC!A1:B5`
- `GET /api/agenda` retorna `DPC!F5:T43`
