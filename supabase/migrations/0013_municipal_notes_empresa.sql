-- ════════════════════════════════════════════════════════════════════════════
-- Conciliação por empresa do grupo (BR02, BR04, ...).
-- A nota da prefeitura passa a guardar a empresa (escolhida na importação),
-- já que a razão social/CNPJ do tomador nem sempre batem com o nome na receita.
-- ════════════════════════════════════════════════════════════════════════════

alter table public.municipal_notes add column if not exists empresa text;
create index if not exists municipal_notes_empresa_idx on public.municipal_notes(empresa);

notify pgrst, 'reload schema';
