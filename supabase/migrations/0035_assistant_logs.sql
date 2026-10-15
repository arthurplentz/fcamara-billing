-- ════════════════════════════════════════════════════════════════════════════
-- ASSISTENTE O2C — log de perguntas e respostas.
-- Gravado pela Edge Function "assistente-o2c" (service role). Serve para:
--   • auditoria (quem perguntou o quê, o que a IA respondeu);
--   • ver o que o time mais pergunta → quais cadastros completar primeiro;
--   • acompanhar o consumo (tokens) da API.
-- Só ADMIN lê. Ninguém escreve pelo app (só a função, com service role).
-- ════════════════════════════════════════════════════════════════════════════

create table if not exists public.assistant_logs (
  id            uuid primary key default gen_random_uuid(),
  user_id       uuid references auth.users(id) on delete set null,
  usuario       text,
  pergunta      text,
  resposta      text,
  input_tokens  integer default 0,
  output_tokens integer default 0,
  created_at    timestamptz not null default now()
);
create index if not exists assistant_logs_created_idx on public.assistant_logs(created_at desc);

alter table public.assistant_logs enable row level security;
drop policy if exists al_read on public.assistant_logs;
create policy al_read on public.assistant_logs for select to authenticated using (public.is_admin());

notify pgrst, 'reload schema';
