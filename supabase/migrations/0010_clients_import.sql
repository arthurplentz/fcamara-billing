-- ════════════════════════════════════════════════════════════════════════════
-- Clientes: CNPJ, flag de cadastro incompleto (carga em massa) e analista dono.
-- ════════════════════════════════════════════════════════════════════════════

alter table public.clients add column if not exists cnpj       text;
alter table public.clients add column if not exists incompleto boolean default false;
alter table public.clients add column if not exists owner      text;  -- analista responsável pelo cliente

create index if not exists clients_cod_sap_idx on public.clients(cod_sap);
