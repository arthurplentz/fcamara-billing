-- ════════════════════════════════════════════════════════════════════════════
-- Ordem de venda por profissional (registro) — campo OPCIONAL, de organização
-- interna do analista. Não faz parte dos dados de reconhecimento importados, então
-- o analista pode preencher/editar livremente (fora do records_guard).
-- ════════════════════════════════════════════════════════════════════════════

alter table public.records add column if not exists ordem_venda text;

notify pgrst, 'reload schema';
