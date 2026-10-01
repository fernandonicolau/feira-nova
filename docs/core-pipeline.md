# Pipeline do core

## Fluxo legado

1. `src/index.js` procura planilhas em `input/` a partir de `process.cwd()`.
2. Cada workbook identifica a loja pelo cabeçalho `FILIAL` ou pelo nome do arquivo.
3. Itens estruturados ou linhas livres são convertidos em produto e quantidade.
4. Nomes de produtos são normalizados e associados aos três templates em `template/mapa/`.
5. Os mapas são gravados em uma pasta global `output-*`.
6. `scripts/generate-fornecedores.js` relê os mapas e usa modelos de `exemplo/` para gerar arquivos por fornecedor.

As dependências de runtime do fluxo antigo são `input/`, `template/mapa/`, `data/`, `exemplo/` e `output-*`. A pasta `exemplo/` não está versionada no estado atual e, portanto, a geração legada de fornecedores não é reproduzível em um checkout limpo.

## Entrada programática

`processBatch({ entries, templateDir?, now? })` recebe planilhas como buffers, interpreta todas as entradas e devolve mapas também como buffers. A função não procura `input/`, não cria `output-*` e não conhece HTTP.

Os templates e catálogos permanecem assets versionados somente para leitura. A CLI continua disponível como adaptador temporário e chama as mesmas funções antes de gravar seus arquivos.
