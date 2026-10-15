# Assistente O2C — chat com a base de premissas (Notion + Claude)

Tela **Operação → Assistente O2C** no app. O analista pergunta ("qual a data de corte
da NF da Três Corações?") e a função:

1. confere se a pessoa está logada no app (perfil comercial não usa);
2. deixa o Claude consultar o Notion (bases **Projetos** e **Clientes**, guia de
   preenchimento e processo de e-commerce) — **só leitura**;
3. devolve a resposta com os PEPs usados como fonte (link para a ficha no Notion);
4. grava pergunta/resposta em `assistant_logs` (só admin lê).

As chaves ficam **somente** nos segredos do Supabase. Nada sensível vai para o
código (o repositório é público) nem para o navegador.

## Ativar (uma vez)

### 1. Chave da API do Claude
1. Entre em <https://console.anthropic.com> (crie a conta da empresa, se ainda não houver).
2. **Billing**: cadastre um cartão e coloque crédito. Em **Limits**, defina um limite mensal de gasto.
3. **API Keys → Create Key** (nome: `assistente-o2c`). Copie a chave (`sk-ant-...`) — ela só aparece uma vez.

### 2. Token de integração do Notion
1. Logado com a conta **dona do workspace** (Daniela), abra <https://www.notion.so/my-integrations>.
2. **New integration** → nome `Assistente O2C`, tipo **Internal**, workspace "Daniela Feck".
3. Em **Capabilities**, deixe só **Read content** (desmarque Update/Insert).
4. Salve e copie o **Internal Integration Secret** (`ntn_...`).
5. No Notion, abra a página **O2C – Base de Conhecimento** → menu `•••` → **Connections**
   → **Add connection** → `Assistente O2C` → confirme. As subpáginas e bases herdam o acesso.

### 3. Segredos e deploy no Supabase
Pelo terminal (com a [Supabase CLI](https://supabase.com/docs/guides/cli) logada no projeto):

```bash
supabase secrets set ANTHROPIC_API_KEY=sk-ant-xxxx NOTION_TOKEN=ntn_xxxx
supabase functions deploy assistente-o2c
```

Ou pelo painel: **Edge Functions → Secrets** (adicione as duas chaves) e
**Edge Functions → Deploy a new function** colando o `index.ts` com o nome
`assistente-o2c`.

> `SUPABASE_URL`, `SUPABASE_ANON_KEY` e `SUPABASE_SERVICE_ROLE_KEY` já vêm do ambiente.

### 4. Tabela de log
No **SQL Editor**, rode `supabase/migrations/0035_assistant_logs.sql`.
(Sem ela o chat funciona, só não grava o histórico.)

### 5. Testar
Entre no app → **Operação → Assistente O2C** → clique numa das sugestões.

## Segredos opcionais

| Segredo | Para quê | Padrão |
|---|---|---|
| `ANTHROPIC_MODEL` | Trocar o modelo do Claude | `claude-sonnet-5-5` |
| `NOTION_PROJETOS_DS` / `NOTION_CLIENTES_DS` | IDs das bases (se migrar o Notion) | bases atuais |
| `NOTION_PAGE_GUIA` / `NOTION_PAGE_ECOMMERCE` / `NOTION_PAGE_PRINCIPAL` | Páginas de referência | páginas atuais |

## Ver o que o time perguntou

```sql
select created_at, usuario, pergunta, resposta, input_tokens, output_tokens
from public.assistant_logs order by created_at desc limit 100;
```
