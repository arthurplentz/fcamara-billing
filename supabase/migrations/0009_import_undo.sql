-- ════════════════════════════════════════════════════════════════════════════
-- Desfazer importação: marca cada registro com o id da importação (import_id),
-- para conseguir remover de uma vez todos os registros de uma carga específica.
-- ════════════════════════════════════════════════════════════════════════════

alter table public.records        add column if not exists import_id uuid;
alter table public.import_history add column if not exists import_id uuid;
create index if not exists records_import_id_idx on public.records(import_id);
