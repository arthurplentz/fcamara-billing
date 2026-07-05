-- ════════════════════════════════════════════════════════════════════════════
-- DESFAZER IMPORTAÇÃO (merge) — guarda o estado anterior dos registros.
-- Ao "Atualizar mês", salvamos em import_history.snapshot o retrato de cada
-- registro ANTES da atualização. Desfazer restaura esses valores e apaga os
-- registros incluídos na importação.
-- ════════════════════════════════════════════════════════════════════════════

alter table public.import_history add column if not exists snapshot jsonb;

notify pgrst, 'reload schema';
