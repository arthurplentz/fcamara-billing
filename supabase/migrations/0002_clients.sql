-- ════════════════════════════════════════════════════════════════════════════
-- Perfil de faturamento dos clientes (aba "Clientes")
-- Cadastro manual das peculiaridades de faturamento de cada cliente.
-- ════════════════════════════════════════════════════════════════════════════

create table public.clients (
  id                 uuid primary key default gen_random_uuid(),
  nome               text not null,
  cod_sap            text,
  grupo_empresa      text,
  tipos_contrato     text,
  proposta_url       text,
  periodo_faturamento text,
  calendario         text,
  tem_portal         boolean default false,
  portal_link        text,
  portal_usuario     text,
  portal_senha       text,
  portal_passo_url   text,
  prazo_vencimento   text,
  forma_pagamento    text,
  contato_financeiro text,
  account_manager    text,
  created_at         timestamptz not null default now(),
  updated_at         timestamptz not null default now()
);

alter table public.clients enable row level security;

-- Analistas e admins podem consultar e manter o cadastro de clientes.
create policy clients_all on public.clients
  for all to authenticated using (true) with check (true);
