-- ════════════════════════════════════════════════════════════════════════════
-- Base de PROJETOS ATIVOS (export do SAP) + marcas do analista — COMPARTILHADA.
-- Antes a lista de projetos e os flags (interno / checado) ficavam só no
-- navegador (localStorage), então cada pessoa via uma tela diferente. Agora a
-- base vive no banco: todo mundo (analista, admin, FP&A) vê a MESMA conferência,
-- e o "checado" aparece ao vivo para o time usar como verídico.
--
-- active_projects  → a lista importada (substituída a cada importação).
-- project_marks    → marcas por PEP que SOBREVIVEM a re-importações
--                    (interno = não espera receita; checado = analisado pelo FP&A).
-- ════════════════════════════════════════════════════════════════════════════

create table if not exists public.active_projects (
  id          uuid primary key default gen_random_uuid(),
  pep         text,              -- código do PEP extraído do nome (ex.: BR02CLP00041)
  pep_base    text,              -- base (antes do 1º ponto) — chave de cruzamento
  nome        text,              -- nome completo do projeto
  cliente     text,
  gerente     text,
  empresa     text,              -- BR02, BR04, ...
  bu          text,              -- "BU Finance", "BU Others", ...
  fase        text,              -- 10 execução · 40 concluído · 00/05 ...
  inicio_key  text,              -- "AAAA-MM" (janela de vigência)
  fim_key     text,              -- "AAAA-MM"
  inativ      boolean default false,  -- nome contém INATIV
  concluido   boolean default false,  -- fase 40
  interno_auto boolean default false, -- PEP INP (projeto interno)
  created_at  timestamptz not null default now()
);
create index if not exists active_projects_pepbase_idx on public.active_projects(pep_base);

alter table public.active_projects enable row level security;
drop policy if exists ap_read  on public.active_projects;
create policy ap_read  on public.active_projects for select to authenticated using (true);
drop policy if exists ap_write on public.active_projects;
create policy ap_write on public.active_projects for all to authenticated
  using (true) with check (true);

-- Marcas manuais por PEP (chave: PEP base, ou "np:empresa|cliente|nome" sem PEP).
create table if not exists public.project_marks (
  mark_key    text primary key,
  interno     boolean default false,
  checado     boolean default false,
  checado_por text,
  updated_at  timestamptz not null default now()
);

alter table public.project_marks enable row level security;
drop policy if exists pm_read  on public.project_marks;
create policy pm_read  on public.project_marks for select to authenticated using (true);
drop policy if exists pm_write on public.project_marks;
create policy pm_write on public.project_marks for all to authenticated
  using (true) with check (true);

notify pgrst, 'reload schema';
