-- ════════════════════════════════════════════════════════════════════════════
-- ROTINAS do time (tarefas recorrentes) + marcação mensal de "feito".
-- Cada pessoa tem suas rotinas (diária/semanal/mensal) com prazo e premissa.
-- No checklist do mês, a pessoa marca o que já fez; o sistema aponta se saiu
-- no prazo. Compartilhado: a gestão vê o status de todo mundo.
-- ════════════════════════════════════════════════════════════════════════════

create table if not exists public.routines (
  id          uuid primary key default gen_random_uuid(),
  pessoa      text,                 -- responsável (nome)
  area        text,                 -- "Contas a Receber" | "Financial Controls"
  titulo      text,
  cliente     text,                 -- opcional
  cadencia    text default 'mensal',-- diaria | semanal | mensal
  prazo_txt   text,                 -- descrição ("1º dia útil", "dia 15", "data de corte 20")
  prazo_tipo  text default 'livre', -- dia_mes | dia_util | livre
  prazo_valor int,                  -- dia do mês (dia_mes) ou N-ésimo dia útil (dia_util)
  premissa    text,                 -- particularidades
  ordem       int default 0,
  ativo       boolean default true,
  created_at  timestamptz not null default now()
);
create index if not exists routines_pessoa_idx on public.routines(pessoa);

alter table public.routines enable row level security;
drop policy if exists rt_read  on public.routines;
create policy rt_read  on public.routines for select to authenticated using (true);
drop policy if exists rt_write on public.routines;
create policy rt_write on public.routines for all to authenticated using (true) with check (true);

-- Marcação de conclusão por rotina + período ("AAAA-MM").
create table if not exists public.routine_done (
  id          uuid primary key default gen_random_uuid(),
  routine_id  uuid not null,
  periodo     text not null,        -- "AAAA-MM"
  feito       boolean default true,
  feito_em    date,
  feito_por   text,
  obs         text,
  updated_at  timestamptz not null default now(),
  unique (routine_id, periodo)
);
create index if not exists routine_done_routine_idx on public.routine_done(routine_id);

alter table public.routine_done enable row level security;
drop policy if exists rd_read  on public.routine_done;
create policy rd_read  on public.routine_done for select to authenticated using (true);
drop policy if exists rd_write on public.routine_done;
create policy rd_write on public.routine_done for all to authenticated using (true) with check (true);

notify pgrst, 'reload schema';
