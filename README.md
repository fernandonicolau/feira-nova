# feira-nova

Automacao em Node.js para ler planilhas de entrada das filiais e preencher automaticamente os mapas `MAPA.xlsx`, `MAPA2.xlsx` e `MAPA3.xlsx` a partir dos templates da pasta `template/mapa`.

## Como funciona

O script principal:

- le todos os arquivos `.xlsx` e `.xlsm` da pasta `input/`
- identifica a filial pelo cabecalho da planilha ou, como fallback, pelo nome do arquivo
- extrai os itens e quantidades mesmo quando a planilha vem em formatos um pouco diferentes
- normaliza nomes de produtos para encaixar no padrao dos mapas
- aplica duas listas compartilhadas de produtos em `data/produtos-legumes.json` e `data/produtos-frutas.json`
- preenche os templates em `template/mapa`
- cria uma pasta de saida no formato `output-AAAA-MM-DD-HH-mm`
- destaca em amarelo os produtos que nao encontrou no template para revisao manual

## Estrutura esperada

```text
.
|-- input/
|-- data/
|   |-- produtos-legumes.json
|   `-- produtos-frutas.json
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

## Uso no navegador

O projeto agora tambem tem uma interface web estatica em `index.html`, pensada para GitHub Pages.

Fluxo da interface:

- aceita upload de arquivos `.xlsx` e `.xlsm`
- aceita upload de `.zip` com varias planilhas dentro
- gera os mapas diretamente no navegador
- libera um botao para baixar um `.zip` com `MAPA.xlsx`, `MAPA2.xlsx` e `MAPA3.xlsx`
- possui um botao `Limpar` para remover os arquivos carregados e resetar a tela

### Publicar no GitHub Pages

1. Suba o projeto para o GitHub.
2. Em `Settings > Pages`, selecione `GitHub Actions` como source.
3. A cada push na branch `master`, a workflow `.github/workflows/deploy-pages.yml` roda `npm ci` e `npm run build:web`.
4. O script `scripts/build-pages.js` gera a pasta `dist/` com a versao publicada do frontend.
5. O deploy envia apenas `dist/` para o GitHub Pages.

Importante:

- no GitHub Pages nao existe backend, entao o processamento acontece 100% no navegador
- a interface depende dos arquivos de template versionados no repositorio
- para abrir localmente com todos os recursos funcionando, prefira servir a pasta com um servidor estatico ou usar o proprio GitHub Pages

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
