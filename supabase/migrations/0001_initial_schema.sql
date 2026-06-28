-- ════════════════════════════════════════════════════════════════════════════
-- Fcamara Billing — Esquema inicial (Fase 1 da migração para backend)
-- Banco: PostgreSQL (Supabase). Aplica tabelas, funções auxiliares e RLS.
-- ════════════════════════════════════════════════════════════════════════════

create extension if not exists "pgcrypto";

-- ─── PROFILES ────────────────────────────────────────────────────────────────
-- Estende auth.users. A autenticação (e-mail/senha, reset) é gerida pelo
-- Supabase Auth; aqui guardamos apenas nome de exibição e papel.
create table public.profiles (
  id         uuid primary key references auth.users(id) on delete cascade,
  name       text not null,
  is_admin   boolean not null default false,
  created_at timestamptz not null default now()
);

-- ─── FUNÇÕES AUXILIARES ──────────────────────────────────────────────────────
-- SECURITY DEFINER para evitar recursão de RLS ao consultar profiles.

create or replace function public.is_admin()
returns boolean
language sql stable security definer set search_path = public
as $$
  select coalesce((select p.is_admin from public.profiles p where p.id = auth.uid()), false);
$$;

create or replace function public.current_name()
returns text
language sql stable security definer set search_path = public
as $$
  select (select p.name from public.profiles p where p.id = auth.uid());
$$;

-- Cria o profile automaticamente quando um usuário é registrado no Auth.
-- O nome e o papel vêm dos metadados informados no cadastro.
create or replace function public.handle_new_user()
returns trigger
language plpgsql security definer set search_path = public
as $$
begin
  insert into public.profiles (id, name, is_admin)
  values (
    new.id,
    coalesce(new.raw_user_meta_data->>'name', new.email),
    coalesce((new.raw_user_meta_data->>'is_admin')::boolean, false)
  );
  return new;
end;
$$;

create trigger on_auth_user_created
  after insert on auth.users
  for each row execute function public.handle_new_user();

-- ─── RECORDS (registros de faturamento) ──────────────────────────────────────
create table public.records (
  id            uuid primary key default gen_random_uuid(),
  responsavel   text not null,
  empresa       text not null,
  tipo          text not null,
  cod_cliente   text,
  cliente       text not null,
  pep           text not null,
  inicio        text,
  fim           text,
  profissional  text,
  valor_venda   numeric default 0,
  hrs_aprovadas numeric default 0,
  valor_total   numeric default 0,
  valor_liquido numeric default 0,
  competencia   text not null,
  progress      jsonb not null default '{}'::jsonb,
  nf_numero     text default '',
  obs           text default '',
  created_at    timestamptz not null default now(),
  updated_at    timestamptz not null default now()
);
create index records_competencia_idx on public.records (competencia);
create index records_responsavel_idx on public.records (responsavel);
create index records_empresa_tipo_idx on public.records (empresa, tipo);

-- ─── TASKS (kanban do time) ──────────────────────────────────────────────────
create table public.tasks (
  id         uuid primary key default gen_random_uuid(),
  title      text not null,
  descricao  text default '',
  due_date   date,
  assignee   text,
  status     text not null default 'inbox',
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

-- ─── IMPORT HISTORY (log de importações) ─────────────────────────────────────
create table public.import_history (
  id          uuid primary key default gen_random_uuid(),
  date        timestamptz not null default now(),
  competencia text,
  empresa     text,
  tipo        text,
  mode        text,
  count       integer,
  user_name   text,
  note        text
);

-- ════════════════════════════════════════════════════════════════════════════
-- ROW LEVEL SECURITY
-- ════════════════════════════════════════════════════════════════════════════
alter table public.profiles       enable row level security;
alter table public.records        enable row level security;
alter table public.tasks          enable row level security;
alter table public.import_history enable row level security;

-- PROFILES: todos os autenticados leem (para listas de responsáveis);
-- só admin cria/remove; cada um pode atualizar o próprio, admin atualiza todos.
create policy profiles_select_all on public.profiles
  for select to authenticated using (true);
create policy profiles_admin_insert on public.profiles
  for insert to authenticated with check (public.is_admin());
create policy profiles_update on public.profiles
  for update to authenticated using (public.is_admin() or id = auth.uid());
create policy profiles_admin_delete on public.profiles
  for delete to authenticated using (public.is_admin() and id <> auth.uid());

-- RECORDS: admin vê/edita tudo; analista vê/edita apenas os seus
-- (responsavel = nome do profile). Importar/excluir é exclusivo de admin.
create policy records_select on public.records
  for select to authenticated
  using (public.is_admin() or responsavel = public.current_name());
create policy records_update on public.records
  for update to authenticated
  using (public.is_admin() or responsavel = public.current_name());
create policy records_admin_insert on public.records
  for insert to authenticated with check (public.is_admin());
create policy records_admin_delete on public.records
  for delete to authenticated using (public.is_admin());

-- TASKS: kanban compartilhado — qualquer autenticado opera.
create policy tasks_all on public.tasks
  for all to authenticated using (true) with check (true);

-- IMPORT HISTORY: somente admin.
create policy history_select on public.import_history
  for select to authenticated using (public.is_admin());
create policy history_insert on public.import_history
  for insert to authenticated with check (public.is_admin());
