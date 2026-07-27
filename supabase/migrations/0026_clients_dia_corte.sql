-- ════════════════════════════════════════════════════════════════════════════
-- Período quebrado: dia de corte do faturamento por cliente.
-- A maioria fatura de 01 a 31 (mês-calendário). Alguns clientes faturam num
-- ciclo deslocado (ex.: 10 a 10, 20 a 20). O dia de corte marca esse ciclo — o
-- app deriva a "competência de faturamento" (visão cliente) a partir das datas
-- do registro + este dia. Vazio/null = 01 a 31 (comportamento normal).
-- ════════════════════════════════════════════════════════════════════════════

alter table public.clients add column if not exists dia_corte text;

notify pgrst, 'reload schema';
