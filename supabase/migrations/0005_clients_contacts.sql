-- ════════════════════════════════════════════════════════════════════════════
-- Clientes: e-mail do contato financeiro e do account manager.
-- (nome continua nas colunas contato_financeiro / account_manager)
-- ════════════════════════════════════════════════════════════════════════════

alter table public.clients add column if not exists contato_financeiro_email text;
alter table public.clients add column if not exists account_manager_email   text;
