-- ════════════════════════════════════════════════════════════════════════════
-- Cadastro do cliente: PROCESSO centralizado + PROJETOS/vencimentos.
-- Hoje cada analista guarda "o processo do seu cliente" na cabeça (ou numa
-- planilha própria). Estas colunas trazem esse conhecimento para dentro do
-- cadastro, onde todo mundo enxerga:
--   • processo  → texto livre: como este cliente funciona de ponta a ponta
--                 (particularidades, quem aprova, ordem de faturamento, etc.).
--   • projetos  → JSON (texto) montado pelo app: lista de projetos/contratos com
--                 datas de início e VENCIMENTO, valor e status. Serve para
--                 acompanhar projetos vencidos / a vencer, propostas e prazos.
-- Campos aditivos — não alteram receita nem conciliação.
-- ════════════════════════════════════════════════════════════════════════════

alter table public.clients add column if not exists processo text;
alter table public.clients add column if not exists projetos text;  -- [{ nome, pep, inicio, vencimento, valor, status, obs }]

notify pgrst, 'reload schema';
