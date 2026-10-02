# Contrato canônico de lote

Um lote possui uma ou mais entradas e não depende de persistência. Cada entrada tem `id`, `type` e `store`, garantindo a associação explícita com sua loja de origem.

Entradas `file` usam `fileRef` para apontar para uma parte do futuro `multipart/form-data`. Entradas `text` carregam o conteúdo em `text`. O mesmo lote pode conter qualquer combinação dos dois tipos.

```json
{
  "name": "Pedido da manhã",
  "entries": [
    { "id": "loja-1", "type": "file", "store": "Cerâmica", "fileRef": "sheet1" },
    { "id": "loja-2", "type": "text", "store": "Coelho", "text": "Banana 5" }
  ]
}
```

Os limites iniciais são 20 entradas, 10 arquivos, 10 MiB por arquivo, 50 MiB no total e 50 mil caracteres por texto. Respostas de processamento usam `requestId`, resumo, avisos e manifesto de artefatos. Falhas usam código estável, mensagem segura e detalhes de validação.

## Outputs e download

`POST /api/v1/batches/process` mantém a resposta JSON com resumo, avisos e manifesto dos mapas, arquivos de fornecedores e associações pendentes. Cada item aponta para o endpoint de download direto.

`POST /api/v1/batches/process/download` recebe o mesmo `multipart/form-data` e devolve um ZIP (`application/zip`) com:

- `mapas/`: os três mapas processados;
- `fornecedores/`: arquivos que possuem pedidos e, quando necessário, `associacoes-pendentes.xlsx`;
- `manifest.json`: `requestId`, resumo dos artefatos e avisos estruturados.

O download reprocessa o lote no contexto da própria requisição. Temporários têm diretório exclusivo e são removidos antes do encerramento da resposta; não existe pasta global de output nem token persistido no servidor.
