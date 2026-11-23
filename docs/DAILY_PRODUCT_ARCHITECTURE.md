# Produto diário: contrato de persistência

## Escolha

Um projeto Supabase Free guarda metadados no Postgres e originais/resultados em bucket privado. A API NestJS é o único cliente com `SUPABASE_SERVICE_ROLE_KEY`; a Web nunca recebe essa chave. A migração SQL está em `supabase/migrations/202611230001_daily_product.sql` e deve ser executada antes de ativar o fluxo diário. O endpoint legado de lote permanece disponível durante a transição.

## Modelo

- `operation_days`: data civil `YYYY-MM-DD` escolhida pela pessoa em `America/Sao_Paulo`, única; `revision` aumenta quando entradas ou revisão mudam.
- `day_entries`: arquivo original ou texto manual; arquivo em Storage com nome opaco, hash e tamanho. O nome apresentado nunca é usado como chave do objeto. Correções preservam o original.
- `day_runs`: snapshot das entradas e revisão, estado e ZIP gerado. Cada processamento cria uma execução nova; baixar não chama o motor novamente.

## Rotas planejadas

- `POST /api/v1/access/session`, `GET /api/v1/access/session`, `DELETE /api/v1/access/session`.
- `GET /api/v1/days`, `PUT /api/v1/days/:day`, `GET /api/v1/days/:day`.
- `POST /api/v1/days/:day/files`, `POST /api/v1/days/:day/texts`, `PATCH/DELETE /api/v1/days/:day/entries/:id`.
- `POST /api/v1/days/:day/runs`, `GET /api/v1/days/:day/runs/:id/download`.

Todas as rotas de dias exigem sessão; `health` permanece público. Upload e execução são operações explícitas separadas. O bucket não tem policy de acesso público.

## Limites e recuperação

O plano Free tem 1 GB de Storage e 500 MB de Postgres. A API impõe limite por upload e total operacional menor que 1 GB; monitoramento no painel do Supabase e exportação periódica são necessários. O plano Free pode pausar após inatividade; o projeto precisa ser restaurado no painel antes de voltar a operar. Antes de usar dados reais, fazer exportação do banco e bucket, ensaiar restauração e documentar retenção. Metadados e objetos podem divergir em falhas parciais; cada operação compensa o objeto recém-criado se o registro no banco falhar.
