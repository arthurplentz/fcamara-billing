-- ════════════════════════════════════════════════════════════════════════════
-- VARIAÇÃO DE RECEITA PÓS-FECHAMENTO (ex.: Casas Bahia).
-- Alguns clientes trazem um valor extra por consultor DEPOIS que a receita já
-- foi fechada/conciliada. Esse extra NÃO altera a receita original (que continua
-- sendo a verdade final) — é um lançamento à parte, faturável, com histórico.
--   clients.aceita_variacao  → liga a função só para os clientes marcados
--   receita_variacoes        → livro de lançamentos por registro (consultor)
-- ════════════════════════════════════════════════════════════════════════════

alter table public.clients add column if not exists aceita_variacao boolean not null default false;

create table if not exists public.receita_variacoes (
  id         uuid primary key default gen_random_uuid(),
  record_id  uuid not null references public.records(id) on delete cascade,
  valor      numeric not null default 0,
  motivo     text,
  criado_por text,
  created_at timestamptz not null default now()
);
create index if not exists rv_record_idx on public.receita_variacoes(record_id);

alter table public.receita_variacoes enable row level security;
drop policy if exists rv_read  on public.receita_variacoes;
drop policy if exists rv_write on public.receita_variacoes;
create policy rv_read  on public.receita_variacoes for select to authenticated using (true);
create policy rv_write on public.receita_variacoes for all    to authenticated using (true) with check (true);

notify pgrst, 'reload schema';
