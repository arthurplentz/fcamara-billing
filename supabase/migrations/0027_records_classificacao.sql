-- ════════════════════════════════════════════════════════════════════════════
-- Classificação do não-faturado.
-- Para receitas que ainda NÃO foram faturadas, o analista registra o porquê:
--   • class_motivo — motivo padronizado (lista fechada na app + "Outro").
--   • class_obs    — texto livre para detalhar (ex.: motivo do represamento).
-- A CATEGORIA (dentro do ciclo × represado) NÃO é armazenada: é derivada da data
-- (mês atual × competência de faturamento do cliente), então muda sozinha quando
-- o mês vira. Vazio nos dois campos = ainda sem classificação.
--
-- Campos livres para qualquer analista editar — o records_guard (0004) só protege
-- os dados de reconhecimento (valores, cliente, PEP, competência…), não estes.
-- ════════════════════════════════════════════════════════════════════════════

alter table public.records add column if not exists class_motivo text;
alter table public.records add column if not exists class_obs   text;

notify pgrst, 'reload schema';
