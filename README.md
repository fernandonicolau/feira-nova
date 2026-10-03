# Feira Nova API

API NestJS stateless para transformar pedidos em mapas e arquivos de fornecedores. O runtime não usa banco, `input/` ou `output/` globais.

## Requisitos e desenvolvimento

- Node.js 24
- npm 11

```bash
copy .env.example .env
npm ci
npm run start:dev
```

Por padrão a API abre em `http://localhost:3000`. Endpoints auxiliares:

- health check: `GET /health`;
- Swagger UI: `GET /docs`;
- documento OpenAPI JSON: `GET /docs-json`.

## Configuração

```dotenv
NODE_ENV=development
PORT=3000
CORS_ORIGINS=http://localhost:5173
```

`CORS_ORIGINS` aceita origens explícitas separadas por vírgula, sem caminhos ou curingas. Nenhuma variável contém segredo.

## Processar um lote

Os dois endpoints recebem `multipart/form-data`. O campo `batch` contém o contrato JSON; cada entrada de arquivo referencia pelo `fileRef` o nome da respectiva parte multipart.

### Arquivo único ou múltiplos arquivos

```bash
curl -X POST http://localhost:3000/api/v1/batches/process \
  -F 'batch={"name":"Pedido manhã","entries":[{"id":"ceramica","type":"file","store":"Cerâmica","fileRef":"sheet1"},{"id":"coelho","type":"file","store":"Coelho","fileRef":"sheet2"}]}' \
  -F 'sheet1=@ceramica.xlsx' \
  -F 'sheet2=@coelho.xlsx'
```

### Entrada textual

```bash
curl -X POST http://localhost:3000/api/v1/batches/process \
  -F 'batch={"name":"Pedido manual","entries":[{"id":"manual-1","type":"text","store":"Queimados","text":"ABACATE 5\nBANANA PRATA 2"}]}'
```

### Lote misto

```bash
curl -X POST http://localhost:3000/api/v1/batches/process \
  -F 'batch={"entries":[{"id":"arquivo-1","type":"file","store":"Cerâmica","fileRef":"sheet1"},{"id":"manual-1","type":"text","store":"Coelho","text":"ABACATE 3"}]}' \
  -F 'sheet1=@ceramica.xlsx'
```

`POST /api/v1/batches/process` retorna JSON com `requestId`, resumo, avisos e manifesto de mapas, fornecedores e pendências.

## Download ZIP

Envie o mesmo multipart para `POST /api/v1/batches/process/download`. A resposta `application/zip` contém:

```text
mapas/
  MAPA.xlsx
  MAPA2.xlsx
  MAPA3.xlsx
fornecedores/
  ...arquivos com pedidos...
  associacoes-pendentes.xlsx (quando necessário)
manifest.json
```

Exemplo:

```bash
curl -X POST http://localhost:3000/api/v1/batches/process/download \
  -F 'batch={"entries":[{"id":"manual-1","type":"text","store":"Cerâmica","text":"CEBOLA ROXA 6"}]}' \
  --output feira-nova.zip
```

Cada requisição usa buffers e, somente para adaptar o gerador legado de fornecedores, um diretório temporário exclusivo removido em `finally`.

## Erros e limites

Erros usam envelope JSON com `requestId`, código estável, mensagem segura e detalhes quando aplicável. Limites atuais:

- 20 entradas por lote;
- 10 arquivos;
- 10 MiB por arquivo e 50 MiB no total;
- 50 mil caracteres por entrada textual;
- formatos `.xlsx` e `.xlsm`.

## Validação

```bash
npm run typecheck
npm test
npm run test:e2e
npm run build
```

A suíte cobre normalização, extração, mapas, fornecedores, arquivo/texto/misto, erros, ZIP, cleanup e isolamento concorrente.

## Docker e Render

```bash
docker build -t feira-nova-api .
docker run --rm -p 3000:3000 --env-file .env feira-nova-api
```

No Render, use o `Dockerfile`, branch `master`, health check `/health`, `NODE_ENV=production` e `CORS_ORIGINS=https://feira-nova-web.onrender.com`. `PORT` é fornecida pelo Render. A imagem executa sem privilégios e não precisa de volume persistente.

O deploy é nativo do Render a partir do Git. O workflow `API CI` apenas valida typecheck, testes, build e imagem Docker; não requer `RENDER_API_KEY` nem publica o serviço.

## Estrutura principal

- `src/api/`: controllers, configuração e aplicação NestJS;
- `src/index.js`: core de planilhas preservado e chamável por buffers;
- `template/mapa/`: modelos dos mapas;
- `template/fornecedores/`: modelos de fornecedores;
- `data/`: catálogos de produtos;
- `test/`: regressão e E2E;
- `docs/batch-contract.md`: contrato detalhado do lote.

O runtime produtivo é exclusivamente a API. O antigo frontend estático, o deploy no GitHub Pages e os comandos baseados em pastas globais `input/`, `output/` e `exemplo/` foram removidos. Fixtures sintéticas vivem em `test/fixtures/`; modelos necessários ao processamento permanecem versionados em `template/`.
