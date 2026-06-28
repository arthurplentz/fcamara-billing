-- ════════════════════════════════════════════════════════════════════════════
-- Vínculo robusto do analista com seus registros.
-- O analista vê os registros onde "responsavel" = o nome do seu perfil.
-- Aqui tornamos a comparação tolerante a maiúsculas/minúsculas e espaços.
-- (Mesmo assim, o NOME do perfil precisa bater com o "Responsável" da planilha.)
-- ════════════════════════════════════════════════════════════════════════════

drop policy if exists records_select on public.records;
create policy records_select on public.records
  for select to authenticated
  using (public.is_admin() or lower(btrim(responsavel)) = lower(btrim(public.current_name())));

drop policy if exists records_update on public.records;
create policy records_update on public.records
  for update to authenticated
  using (public.is_admin() or lower(btrim(responsavel)) = lower(btrim(public.current_name())));
