# Produto diário: contrato de persistência

## Escolha

Um projeto Supabase Free guarda metadados no Postgres e originais/resultados em bucket privado. A API NestJS é o único cliente com `SUPABASE_SERVICE_ROLE_KEY`; a Web nunca recebe essa chave. Execute as migrações em `supabase/migrations/` na ordem do nome antes de ativar o fluxo diário. O endpoint legado de lote permanece disponível durante a transição.

## Modelo

- `operation_days`: data civil `YYYY-MM-DD` escolhida pela pessoa em `America/Sao_Paulo`, única; `revision` aumenta quando entradas ou revisão mudam.
- `day_entries`: arquivo original ou texto manual; arquivo em Storage com nome opaco, hash e tamanho. O nome apresentado nunca é usado como chave do objeto. Correções preservam o original.
- `day_runs`: snapshot das entradas e revisão, estado e ZIP gerado. Cada processamento cria uma execução nova; baixar não chama o motor novamente.

## Rotas planejadas

- `POST /api/v1/access/session`, `GET /api/v1/access/session`, `DELETE /api/v1/access/session`.
- `GET /api/v1/days`, `PUT /api/v1/days/:day`, `GET /api/v1/days/:day`.
- `POST /api/v1/days/:day/files`, `POST /api/v1/days/:day/texts`, `POST /api/v1/days/:day/entries/:id/replace`, `DELETE /api/v1/days/:day/entries/:id`, `GET /api/v1/days/:day/uploads/:id/original`.
- `GET /api/v1/days/stores`, `PATCH /api/v1/days/:day/entries/:id/review`.
- `POST /api/v1/days/:day/runs`, `GET /api/v1/days/:day/runs/:id/download`.

As rotas de dias e entradas até o download do original são entregues na E04-T03. Execuções e seus downloads pertencem à E04-T06. Upload de vários arquivos retorna sucessos e falhas por arquivo; o mesmo `clientUploadId` em uma tentativa repetida recupera os sucessos já salvos sem duplicá-los. ZIP preserva o original no Storage e cria entradas ligadas a ele para cada planilha interna.

A E04-T04 identifica a loja pelo cabeçalho e nome do arquivo, registra fonte e confiança e sugere pares de colunas de produto e quantidade. Divergências ou estrutura desconhecida ficam em revisão no arquivo afetado. A correção manual da loja e das colunas é salva na entrada, mantém o objeto e o hash originais e deve ser passada ao motor pela E04-T06. A prévia mostra itens reconhecidos e avisos de linhas incompletas.

Todas as rotas de dias exigem sessão; `health` permanece público. Upload e execução são operações explícitas separadas. O bucket não tem policy de acesso público.

## Limites e recuperação

O plano Free tem 1 GB de Storage e 500 MB de Postgres. A API impõe limite por upload e total operacional menor que 1 GB; monitoramento no painel do Supabase e exportação periódica são necessários. O plano Free pode pausar após inatividade; o projeto precisa ser restaurado no painel antes de voltar a operar. Antes de usar dados reais, fazer exportação do banco e bucket, ensaiar restauração e documentar retenção. Metadados e objetos podem divergir em falhas parciais; cada operação compensa o objeto recém-criado se o registro no banco falhar.

Para validar localmente com Docker: `npx supabase start` aplica as migrações. O teste `test/day-flow.e2e-spec.ts` exige `RUN_DAY_INTEGRATION=1`, `SUPABASE_URL` local e `SUPABASE_SERVICE_ROLE_KEY` local; ele não executa contra um projeto remoto. O ambiente local pode ser parado com `npx supabase stop`.
