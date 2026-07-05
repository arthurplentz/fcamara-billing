-- ════════════════════════════════════════════════════════════════════════════
-- RECEITA CONCILIADA = VERDADE FINAL (imutável até reabrir).
-- Quando uma base nova traz um valor diferente para um item JÁ CONCILIADO, o
-- valor conciliado NÃO muda. Guardamos o valor que a base trouxe em
-- valor_base_divergente só para exibir o alerta "base R$ X ≠ conciliado R$ Y —
-- reabra para atualizar". Ao reabrir a conciliação, esse valor é aplicado e o
-- marcador é limpo.
-- ════════════════════════════════════════════════════════════════════════════

alter table public.records add column if not exists valor_base_divergente numeric;

notify pgrst, 'reload schema';
