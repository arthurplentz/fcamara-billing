-- ════════════════════════════════════════════════════════════════════════════
-- BU (unidade de negócio) por registro de receita.
-- Classificação comercial: BU Health, Multisector, Logistics, Others, Finance,
-- Retail. Normalmente 1 BU por cliente (com exceções raras por projeto).
-- Campo aditivo — não altera receita nem conciliação; setado só pelo admin
-- (via classificador no app ou lido da coluna "Vertical" da importação).
-- ════════════════════════════════════════════════════════════════════════════

alter table public.records add column if not exists bu text;

notify pgrst, 'reload schema';
