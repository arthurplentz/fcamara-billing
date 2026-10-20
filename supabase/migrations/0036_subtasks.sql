-- ════════════════════════════════════════════════════════════════════════════
-- Detalhe da tarefa + subtarefas atribuíveis.
--  · routines.descricao  → descrição fixa da rotina (a gestão escreve o "como fazer").
--  · routine_done.obs    → observação que CADA pessoa escreve na sua execução do período
--                          (já existia a coluna; agora é usada pela tela de detalhe).
--  · routine_subtasks    → subtarefas criadas dentro de uma rotina. Quando o "responsavel"
--                          é outra pessoa, elas aparecem como um CARD novo no quadro dela,
--                          como uma tarefa repassada (a rotina original fica dependendo dela).
-- ════════════════════════════════════════════════════════════════════════════

alter table public.routines add column if not exists descricao text;

create table if not exists public.routine_subtasks (
  id          uuid primary key default gen_random_uuid(),
  routine_id  uuid not null,        -- rotina original (de onde a subtarefa nasceu)
  titulo      text,
  responsavel text,                 -- pessoa designada (nome)
  criado_por  text,                 -- quem repassou
  prazo       date,                 -- opcional
  feito       boolean default false,
  feito_em    date,
  feito_por   text,
  obs         text,
  ordem       int default 0,
  created_at  timestamptz not null default now()
);
create index if not exists routine_subtasks_routine_idx on public.routine_subtasks(routine_id);
create index if not exists routine_subtasks_resp_idx    on public.routine_subtasks(responsavel);

alter table public.routine_subtasks enable row level security;
drop policy if exists rs_read  on public.routine_subtasks;
create policy rs_read  on public.routine_subtasks for select to authenticated using (true);
drop policy if exists rs_write on public.routine_subtasks;
create policy rs_write on public.routine_subtasks for all to authenticated using (true) with check (true);

notify pgrst, 'reload schema';
