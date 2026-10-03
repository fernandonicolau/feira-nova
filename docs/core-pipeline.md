# Pipeline do core

## Fluxo atual

1. A API recebe um lote HTTP com arquivos e/ou entradas textuais.
2. O controller valida o contrato e converte uploads em buffers.
3. `processBatch` normaliza todas as entradas no mesmo modelo interno.
4. O core gera mapas em memória usando `template/mapa/` e `data/`.
5. O serviço adapta os mapas a um workspace temporário exclusivo para gerar os arquivos de fornecedores com os modelos versionados em `template/fornecedores/`.
6. Os artefatos são devolvidos por HTTP e o workspace é removido em `finally`.

Nenhum fluxo produtivo procura pastas globais `input/`, `output/` ou `exemplo/`. Os casos de regressão usam buffers e fixtures sintéticas de `test/fixtures/`.

## Entrada programática

`processBatch({ entries, templateDir?, now? })` recebe planilhas ou textos normalizados, interpreta as entradas e devolve mapas como buffers. A função não conhece HTTP nem grava artefatos em diretórios globais.

Os templates e catálogos permanecem assets versionados somente para leitura. A escrita em disco fica restrita ao workspace temporário isolado de cada requisição.
