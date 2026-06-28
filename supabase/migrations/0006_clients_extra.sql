-- ════════════════════════════════════════════════════════════════════════════
-- Clientes: PEPs por tipo de contrato, múltiplas propostas e classificação do portal.
-- Colunas guardam JSON (texto) montado pelo app.
-- ════════════════════════════════════════════════════════════════════════════

alter table public.clients add column if not exists tipos_peps  text;  -- { "Time & Expenses": ["PEP1","PEP2"], ... }
alter table public.clients add column if not exists propostas   text;  -- ["link1","link2"]
alter table public.clients add column if not exists portal_tipo text;  -- "Inclusão de notas, Medição de serviços"
