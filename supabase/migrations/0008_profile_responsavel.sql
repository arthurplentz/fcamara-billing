-- ════════════════════════════════════════════════════════════════════════════
-- Vínculo explícito do login com o "Responsável" da base de receitas.
-- profiles.responsavel = nome exatamente como aparece na coluna "Responsável"
-- do Excel. Se vazio, cai no nome de exibição (name). Assim cada analista
-- enxerga as receitas onde é o responsável.
-- ════════════════════════════════════════════════════════════════════════════

alter table public.profiles add column if not exists responsavel text;

create or replace function public.current_responsavel()
returns text
language sql stable security definer set search_path = public
as $$
  select coalesce(
    (select nullif(btrim(p.responsavel), '') from public.profiles p where p.id = auth.uid()),
    (select p.name from public.profiles p where p.id = auth.uid())
  );
$$;

drop policy if exists records_select on public.records;
create policy records_select on public.records for select to authenticated
  using (public.is_admin() or lower(btrim(responsavel)) = lower(btrim(public.current_responsavel())));

drop policy if exists records_update on public.records;
create policy records_update on public.records for update to authenticated
  using (public.is_admin() or lower(btrim(responsavel)) = lower(btrim(public.current_responsavel())));
