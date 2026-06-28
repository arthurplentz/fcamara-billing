-- ════════════════════════════════════════════════════════════════════════════
-- Proteção dos dados de reconhecimento.
-- Analistas NÃO podem alterar os dados importados (valores, cliente, PEP, etc.).
-- Podem alterar apenas: progresso do funil, número da NF e observações.
-- Apenas o ADMIN pode editar esses campos ou excluir/recarregar registros.
-- ════════════════════════════════════════════════════════════════════════════

create or replace function public.records_guard()
returns trigger
language plpgsql security definer set search_path = public
as $$
begin
  if public.is_admin() then return new; end if;
  if (new.responsavel  is distinct from old.responsavel)
  or (new.empresa      is distinct from old.empresa)
  or (new.tipo         is distinct from old.tipo)
  or (new.cod_cliente  is distinct from old.cod_cliente)
  or (new.cliente      is distinct from old.cliente)
  or (new.pep          is distinct from old.pep)
  or (new.inicio       is distinct from old.inicio)
  or (new.fim          is distinct from old.fim)
  or (new.profissional is distinct from old.profissional)
  or (new.valor_venda  is distinct from old.valor_venda)
  or (new.hrs_aprovadas is distinct from old.hrs_aprovadas)
  or (new.valor_total  is distinct from old.valor_total)
  or (new.valor_liquido is distinct from old.valor_liquido)
  or (new.competencia  is distinct from old.competencia)
  then
    raise exception 'Apenas o administrador pode alterar os dados de reconhecimento (importados).';
  end if;
  return new;
end;
$$;

drop trigger if exists records_guard_update on public.records;
create trigger records_guard_update
  before update on public.records
  for each row execute function public.records_guard();
