-- ════════════════════════════════════════════════════════════════════════════
-- Conciliação N:N — um "conjunto de conciliação" liga VÁRIAS notas a VÁRIAS
-- receitas (os totais devem bater). Ambos os lados guardam o mesmo conciliacao_id.
-- ════════════════════════════════════════════════════════════════════════════

alter table public.records         add column if not exists conciliacao_id uuid;
alter table public.municipal_notes add column if not exists conciliacao_id uuid;
create index if not exists records_conciliacao_idx         on public.records(conciliacao_id);
create index if not exists municipal_notes_conciliacao_idx on public.municipal_notes(conciliacao_id);

notify pgrst, 'reload schema';
