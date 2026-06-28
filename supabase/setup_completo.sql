-- ════════════════════════════════════════════════════════════════════════════
-- SETUP COMPLETO — cole tudo no SQL Editor do Supabase e clique em Run.
-- É SEGURO rodar mesmo que partes já tenham sido aplicadas (não duplica nada).
-- ════════════════════════════════════════════════════════════════════════════

-- ── 1) Tabela de CLIENTES (aba "Clientes") ──────────────────────────────────
create table if not exists public.clients (
  id         uuid primary key default gen_random_uuid(),
  nome       text not null,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);
alter table public.clients add column if not exists cod_sap                  text;
alter table public.clients add column if not exists grupo_empresa            text;
alter table public.clients add column if not exists tipos_contrato           text;
alter table public.clients add column if not exists tipos_peps               text;
alter table public.clients add column if not exists proposta_url             text;
alter table public.clients add column if not exists propostas                text;
alter table public.clients add column if not exists periodo_faturamento      text;
alter table public.clients add column if not exists calendario               text;
alter table public.clients add column if not exists tem_portal               boolean default false;
alter table public.clients add column if not exists portal_tipo              text;
alter table public.clients add column if not exists portal_link              text;
alter table public.clients add column if not exists portal_usuario           text;
alter table public.clients add column if not exists portal_senha             text;
alter table public.clients add column if not exists portal_passo_url         text;
alter table public.clients add column if not exists prazo_vencimento         text;
alter table public.clients add column if not exists forma_pagamento          text;
alter table public.clients add column if not exists contato_financeiro       text;
alter table public.clients add column if not exists contato_financeiro_email text;
alter table public.clients add column if not exists account_manager          text;
alter table public.clients add column if not exists account_manager_email    text;

alter table public.clients enable row level security;
drop policy if exists clients_all on public.clients;
create policy clients_all on public.clients for all to authenticated using (true) with check (true);

-- ── 2) Proteção dos dados de reconhecimento ─────────────────────────────────
-- Analistas não podem alterar dados importados (valores, cliente, PEP, etc.).
create or replace function public.records_guard()
returns trigger language plpgsql security definer set search_path = public as $$
begin
  if public.is_admin() then return new; end if;
  if (new.responsavel  is distinct from old.responsavel)
  or (new.empresa      is distinct from old.empresa)
  or (new.tipo         is distinct from old.tipo)
  or (new.cod_cliente  is distinct from old.cod_cliente)
  or (new.cliente      is distinct from old.cliente)
  or (new.pep          is distinct from old.pep)
  or (new.inicio       is distinct from old.inicio)
  or (new.fim          is distinct from old.fim)
  or (new.profissional is distinct from old.profissional)
  or (new.valor_venda  is distinct from old.valor_venda)
  or (new.hrs_aprovadas is distinct from old.hrs_aprovadas)
  or (new.valor_total  is distinct from old.valor_total)
  or (new.valor_liquido is distinct from old.valor_liquido)
  or (new.competencia  is distinct from old.competencia)
  then raise exception 'Apenas o administrador pode alterar os dados de reconhecimento (importados).';
  end if;
  return new;
end; $$;
drop trigger if exists records_guard_update on public.records;
create trigger records_guard_update before update on public.records
  for each row execute function public.records_guard();

-- ── 3) Vínculo do responsável (ignora maiúsculas/espaços) ───────────────────
drop policy if exists records_select on public.records;
create policy records_select on public.records for select to authenticated
  using (public.is_admin() or lower(btrim(responsavel)) = lower(btrim(public.current_name())));
drop policy if exists records_update on public.records;
create policy records_update on public.records for update to authenticated
  using (public.is_admin() or lower(btrim(responsavel)) = lower(btrim(public.current_name())));
