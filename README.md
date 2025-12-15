# APP DPC

Webservice em Node/Express para exibir as células A1:B5 da aba **DPC** da planilha publicada no Google Sheets.

## Rodar localmente

```bash
npm install
npm start
# abre http://localhost:3000
```

- A porta pode ser definida via variável `PORT`.
- Se a aba tiver outro nome, ajuste `SHEET_NAME` (padrão: `DPC`).

## Deploy no Render

1. Crie um novo Web Service apontando para este repositório.
2. Build Command: `npm install`
3. Start Command: `npm start`
4. Deixe o runtime Node na versão padrão do Render (18+).
5. Opcional: defina `SHEET_NAME` nas variáveis de ambiente se o nome da aba mudar.

O arquivo `render.yaml` já descreve esse serviço.
