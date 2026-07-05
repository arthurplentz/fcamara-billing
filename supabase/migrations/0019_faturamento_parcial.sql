-- ════════════════════════════════════════════════════════════════════════════
-- Faturamento PARCIAL — "livro de faturamento" por registro.
-- Cada registro (profissional) pode ser faturado em pedaços: cada conciliação
-- aloca um valor. Faturado do registro = soma das alocações; saldo = total − faturado.
-- Não altera os dados importados (o valor_total do registro permanece intacto).
-- ════════════════════════════════════════════════════════════════════════════

create table if not exists public.record_faturamentos (
  id             uuid primary key default gen_random_uuid(),
  record_id      uuid not null references public.records(id) on delete cascade,
  conciliacao_id uuid,            -- grupo de conciliação (mesma nota/lote)
  valor          numeric not null default 0,
  criado_por     text,
  created_at     timestamptz not null default now()
);
create index if not exists rf_record_idx on public.record_faturamentos(record_id);
create index if not exists rf_conc_idx   on public.record_faturamentos(conciliacao_id);

alter table public.record_faturamentos enable row level security;
drop policy if exists rf_read  on public.record_faturamentos;
drop policy if exists rf_write on public.record_faturamentos;
create policy rf_read  on public.record_faturamentos for select to authenticated using (true);
create policy rf_write on public.record_faturamentos for all    to authenticated using (true) with check (true);

notify pgrst, 'reload schema';
