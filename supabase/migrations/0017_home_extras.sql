-- ════════════════════════════════════════════════════════════════════════════
-- Home: aniversário do analista (dd/mm) + o próprio analista poder editar o apelido.
-- ════════════════════════════════════════════════════════════════════════════

alter table public.profiles add column if not exists aniversario text; -- "dd/mm"

-- Deixa cada pessoa editar APENAS o próprio apelido (sem poder virar admin).
create or replace function public.set_my_apelido(p text)
returns void
language sql
security definer
set search_path = public
as $$
  update public.profiles set apelido = nullif(btrim(p), '') where id = auth.uid();
$$;

notify pgrst, 'reload schema';
