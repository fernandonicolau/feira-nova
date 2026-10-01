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
