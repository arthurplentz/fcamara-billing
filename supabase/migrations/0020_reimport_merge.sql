-- ════════════════════════════════════════════════════════════════════════════
-- RE-IMPORTAÇÃO INTELIGENTE (merge de fechamento mensal)
-- Ao subir uma base nova do mês, os valores são atualizados sem apagar o
-- progresso nem a conciliação. Estas colunas guardam o alerta de mudança:
--   valor_anterior     → valor total antes da última alteração (NULL = sem alerta)
--   valor_alterado_em  → quando o valor mudou
--   ausente_relatorio  → a linha sumiu do último relatório (sinalizada, não apagada)
-- ════════════════════════════════════════════════════════════════════════════

alter table public.records add column if not exists valor_anterior    numeric;
alter table public.records add column if not exists valor_alterado_em timestamptz;
alter table public.records add column if not exists ausente_relatorio boolean not null default false;

notify pgrst, 'reload schema';
