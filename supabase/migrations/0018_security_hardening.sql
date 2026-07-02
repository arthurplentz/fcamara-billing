-- ════════════════════════════════════════════════════════════════════════════
-- ENDURECIMENTO DE SEGURANÇA (auditoria)
-- Fecha 3 brechas de escalonamento/adulteração. Seguro rodar mais de uma vez.
-- ════════════════════════════════════════════════════════════════════════════

-- ── 1) [CRÍTICO] Impedir auto-promoção a admin ──────────────────────────────
-- A política antiga permitia o usuário atualizar a PRÓPRIA linha em profiles
-- (using: is_admin() OR id = auth.uid()), sem restringir colunas — ou seja, um
-- analista podia chamar a API e fazer  UPDATE profiles SET is_admin = true.
-- Agora só o admin altera profiles pela tabela. A auto-edição do apelido
-- continua funcionando pela função segura set_my_apelido (SECURITY DEFINER).
drop policy if exists profiles_update on public.profiles;
create policy profiles_update on public.profiles
  for update to authenticated
  using (public.is_admin())
  with check (public.is_admin());

-- ── 2) [ALTO] Novo usuário nunca nasce admin ────────────────────────────────
-- handle_new_user lia is_admin dos metadados do cadastro. Se o cadastro público
-- estiver habilitado, alguém poderia se registrar já como admin. Agora ignora o
-- metadado e cria sempre como analista; a promoção é feita por um admin.
create or replace function public.handle_new_user()
returns trigger language plpgsql security definer set search_path = public as $$
begin
  insert into public.profiles (id, name, is_admin)
  values (new.id, coalesce(new.raw_user_meta_data->>'name', new.email), false);
  return new;
end; $$;

-- ── 3) [MÉDIO] Exclusão de notas só para admin (bate com a UI) ──────────────
drop policy if exists municipal_notes_all on public.municipal_notes;
create policy mn_read   on public.municipal_notes for select to authenticated using (true);
create policy mn_insert on public.municipal_notes for insert to authenticated with check (true);
create policy mn_update on public.municipal_notes for update to authenticated using (true) with check (true);
create policy mn_delete on public.municipal_notes for delete to authenticated using (public.is_admin());

notify pgrst, 'reload schema';
