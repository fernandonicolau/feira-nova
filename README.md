# feira-nova

Automacao em Node.js para ler planilhas de entrada das filiais e preencher automaticamente os mapas `MAPA.xlsx`, `MAPA2.xlsx` e `MAPA3.xlsx` a partir dos templates da pasta `template/mapa`.

## Como funciona

O script principal:

- le todos os arquivos `.xlsx` e `.xlsm` da pasta `input/`
- identifica a filial pelo cabecalho da planilha ou, como fallback, pelo nome do arquivo
- extrai os itens e quantidades mesmo quando a planilha vem em formatos um pouco diferentes
- normaliza nomes de produtos para encaixar no padrao dos mapas
- preenche os templates em `template/mapa`
- cria uma pasta de saida no formato `output-AAAA-MM-DD-HH-mm`
- destaca em amarelo os produtos que nao encontrou no template para revisao manual

## Estrutura esperada

```text
.
|-- input/
|-- template/
|   `-- mapa/
|       |-- MAPA.xlsx
|       |-- MAPA2.xlsx
|       `-- MAPA3.xlsx
|-- src/
|   `-- index.js
`-- exemplo/
```

## Requisitos

- Node.js instalado
- dependencias instaladas com `npm install`

## Como usar

1. Coloque os arquivos de entrada das filiais dentro da pasta `input/`.
2. Rode o comando:

```bash
npm run generate
```

3. Abra a pasta `output-...` gerada na raiz do projeto.
4. Revise os arquivos `MAPA.xlsx`, `MAPA2.xlsx` e `MAPA3.xlsx`.
5. Se houver linhas destacadas em amarelo, ajuste manualmente ou atualize as regras de normalizacao no codigo.

## Validacao com exemplo

O projeto tem um conjunto de exemplo para comparacao.

```bash
npm run check:example
```

Esse comando compara a ultima pasta `output-*` gerada com os arquivos de referencia em `exemplo/output/output mapa`.

## Regras importantes

- Qualquer pasta na raiz que comece com `input` fica no `.gitignore`, como `input/`, `input-old/` e variacoes semelhantes.
- Qualquer pasta na raiz que comece com `output` tambem fica no `.gitignore`, incluindo as pastas geradas pelo processo.
- Se voce precisar manter exemplos versionados, vale usar nomes de pasta que nao comecem com `input` ou `output`.
