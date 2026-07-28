-- ─── PAPEL "SOMENTE VISUALIZAÇÃO" (viewer) ───────────────────────────────────
-- Um viewer LÊ tudo (como admin), mas NÃO escreve nada. A trava é aqui no RLS
-- (segurança real); a UI apenas esconde os botões. is_viewer é independente de
-- is_admin (um viewer nunca deve ter is_admin=true).

alter table public.profiles add column if not exists is_viewer boolean not null default false;

-- Helper SECURITY DEFINER (evita recursão de RLS ao consultar profiles).
create or replace function public.is_viewer()
returns boolean
language sql stable security definer set search_path = public
as $$
  select coalesce((select p.is_viewer from public.profiles p where p.id = auth.uid()), false);
$$;

-- ── RECORDS ──────────────────────────────────────────────────────────────────
-- Leitura: viewer enxerga TUDO (como admin). Escrita (update): nunca viewer.
drop policy if exists records_select on public.records;
create policy records_select on public.records for select to authenticated
  using (public.is_admin() or public.is_viewer()
         or lower(btrim(responsavel)) = lower(btrim(public.current_responsavel())));

drop policy if exists records_update on public.records;
create policy records_update on public.records for update to authenticated
  using (not public.is_viewer() and (public.is_admin()
         or lower(btrim(responsavel)) = lower(btrim(public.current_responsavel()))))
  with check (not public.is_viewer() and (public.is_admin()
         or lower(btrim(responsavel)) = lower(btrim(public.current_responsavel()))));
-- insert/delete já são admin-only → viewer (não-admin) já está bloqueado.

-- ── Tabelas "abertas": separa LEITURA (todos) da ESCRITA (não-viewer) ─────────
-- TASKS
drop policy if exists tasks_all on public.tasks;
create policy tasks_select on public.tasks for select to authenticated using (true);
create policy tasks_write  on public.tasks for all    to authenticated
  using (not public.is_viewer()) with check (not public.is_viewer());

-- CLIENTS
drop policy if exists clients_all on public.clients;
create policy clients_select on public.clients for select to authenticated using (true);
create policy clients_write  on public.clients for all    to authenticated
  using (not public.is_viewer()) with check (not public.is_viewer());

-- RECORD_FATURAMENTOS (rf_read continua liberado)
drop policy if exists rf_write on public.record_faturamentos;
create policy rf_write on public.record_faturamentos for all to authenticated
  using (not public.is_viewer()) with check (not public.is_viewer());

-- RECEITA_VARIACOES (rv_read continua liberado)
drop policy if exists rv_write on public.receita_variacoes;
create policy rv_write on public.receita_variacoes for all to authenticated
  using (not public.is_viewer()) with check (not public.is_viewer());

-- MUNICIPAL_NOTES: bloqueia insert/update do viewer (delete já é admin-only).
drop policy if exists mn_insert on public.municipal_notes;
create policy mn_insert on public.municipal_notes for insert to authenticated
  with check (not public.is_viewer());
drop policy if exists mn_update on public.municipal_notes;
create policy mn_update on public.municipal_notes for update to authenticated
  using (not public.is_viewer()) with check (not public.is_viewer());

-- deliveries / delivery_templates / mural / import_history / profiles:
-- a escrita dessas tabelas já exige public.is_admin() → viewer já bloqueado.
