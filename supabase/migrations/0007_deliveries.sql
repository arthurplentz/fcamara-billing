-- ════════════════════════════════════════════════════════════════════════════
-- Entregas recorrentes + tarefas (ordinárias x recorrentes)
-- - delivery_templates: MODELOS de entrega (recorrentes), com uma lista de tarefas
-- - deliveries: instâncias geradas por mês ("Gerar entrega do mês")
-- - tasks ganha vínculo com a entrega + flag de recorrente + competência
-- ════════════════════════════════════════════════════════════════════════════

create table if not exists public.delivery_templates (
  id         uuid primary key default gen_random_uuid(),
  title      text not null,
  items      jsonb not null default '[]'::jsonb,  -- [{title, desc, assignee}]
  created_at timestamptz not null default now()
);

create table if not exists public.deliveries (
  id          uuid primary key default gen_random_uuid(),
  template_id uuid references public.delivery_templates(id) on delete set null,
  title       text not null,
  competencia text,
  created_at  timestamptz not null default now()
);

alter table public.tasks add column if not exists delivery_id uuid references public.deliveries(id) on delete set null;
alter table public.tasks add column if not exists recorrente  boolean default false;
alter table public.tasks add column if not exists competencia text;

alter table public.delivery_templates enable row level security;
alter table public.deliveries         enable row level security;

-- Todos leem; apenas admin cria/edita/gera.
drop policy if exists dt_select on public.delivery_templates;
drop policy if exists dt_admin  on public.delivery_templates;
create policy dt_select on public.delivery_templates for select to authenticated using (true);
create policy dt_admin  on public.delivery_templates for all to authenticated using (public.is_admin()) with check (public.is_admin());

drop policy if exists dv_select on public.deliveries;
drop policy if exists dv_admin  on public.deliveries;
create policy dv_select on public.deliveries for select to authenticated using (true);
create policy dv_admin  on public.deliveries for all to authenticated using (public.is_admin()) with check (public.is_admin());
