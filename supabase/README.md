# Backend (Supabase) — Fase 1: modelagem

Esta pasta contém a estrutura do banco para migrar o app do `localStorage` para um
backend real (dados compartilhados, login seguro, controle de acesso por papel).

> **Status:** Fase 1 (modelagem). Ainda **não** está conectado ao app — a versão
> publicada continua usando `localStorage` e não é afetada por estes arquivos.

## Conteúdo

| Arquivo | O quê |
|---|---|
| `migrations/0001_initial_schema.sql` | Tabelas (`profiles`, `records`, `tasks`, `import_history`), funções auxiliares e políticas de segurança (RLS) |
| `seed.sql` | Dados de exemplo (opcional) para validar as telas |

## Modelo de dados (resumo)

- **profiles** — espelha os usuários do Supabase Auth. Guarda `name` e `is_admin`.
- **records** — registros de faturamento (campos atuais + `progress` em JSON).
- **tasks** — tarefas do kanban.
- **import_history** — log das importações.

### Regras de acesso (RLS)
- **Admin** (Daniela, Luana): acesso total; importa/exporta; gerencia acessos.
- **Analista**: vê e edita **somente os registros onde `responsavel` = seu nome**.
- **Tarefas**: kanban compartilhado entre todos os autenticados.
- **Histórico de importação**: apenas admin.

## Como aplicar (passo a passo)

1. Crie um projeto em <https://supabase.com> (plano Free para validar).
2. No painel, vá em **SQL Editor** e rode o conteúdo de
   `migrations/0001_initial_schema.sql`.
3. Crie o primeiro administrador em **Authentication → Users → Add user**
   (ex.: e-mail da Daniela). Depois, em **SQL Editor**, promova-o a admin:

   ```sql
   update public.profiles set is_admin = true
   where id = (select id from auth.users where email = 'daniela@grupofcamara.com');
   ```

   > Ao criar usuários pelo app (Fase 2), o nome e o papel virão automaticamente
   > pelos metadados — este passo manual é só para o primeiro admin.

4. (Opcional) Rode `seed.sql` para inserir dados de exemplo.
5. Copie as credenciais do projeto (**Project Settings → API**):
   - `Project URL`  → `VITE_SUPABASE_URL`
   - `anon public key` → `VITE_SUPABASE_ANON_KEY`

   Veja `.env.example` na raiz do projeto.

## Próximas fases

- **Fase 2** — Login real (Supabase Auth) substituindo o login local + "esqueci minha senha" por e-mail.
- **Fase 3** — Trocar leitura/escrita do `localStorage` por chamadas ao banco (registros, tarefas, histórico, gestão de acessos).
- **Fase 4** — Validar permissões por papel.
- **Fase 5** — Migração de dados + publicação.
