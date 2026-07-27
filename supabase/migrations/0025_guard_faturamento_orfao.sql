-- ════════════════════════════════════════════════════════════════════════════
-- BARREIRA: receita faturada NÃO pode ficar sem nota (alocação órfã).
-- Uma conciliação amarra receitas (record_faturamentos.conciliacao_id) a notas
-- (municipal_notes.conciliacao_id) pelo MESMO uuid de lote. Se a última nota de
-- um lote for apagada — ou tiver o conciliacao_id limpo — enquanto ainda existir
-- receita alocada nesse lote, a receita fica "Faturado" sem nota. Foi o caso do
-- Erik (nota apagada por um "desfazer importação" sem reabrir o lote).
--
-- Este trigger RECUSA na raiz qualquer operação que deixaria um lote com receita
-- alocada e nenhuma nota. O caminho correto é REABRIR a conciliação (que remove
-- as alocações) ANTES de apagar/desvincular a nota — e é o que a app já faz nos
-- fluxos de reabertura. Vale para app, SQL manual, undo de importação, qualquer um.
-- ════════════════════════════════════════════════════════════════════════════

create or replace function public.guard_faturamento_orfao()
returns trigger
language plpgsql security definer set search_path = public
as $$
declare
  cid uuid := old.conciliacao_id;
begin
  -- Nota que não estava conciliada não pode orfanar nada.
  if cid is null then
    return case tg_op when 'DELETE' then old else new end;
  end if;

  -- Só age se a operação REMOVE o vínculo desta nota com o lote:
  --   DELETE da nota, ou UPDATE que muda/limpa o conciliacao_id.
  if tg_op = 'DELETE' or (new.conciliacao_id is distinct from old.conciliacao_id) then
    -- Bloqueia se o lote ainda tem receita alocada E esta era a ÚLTIMA nota dele.
    if exists (select 1 from public.record_faturamentos f where f.conciliacao_id = cid)
       and not exists (select 1 from public.municipal_notes n
                       where n.conciliacao_id = cid and n.id <> old.id)
    then
      raise exception
        'Nota % pertence a uma conciliação com receita alocada (lote %). Reabra a conciliação antes de remover/desvincular a nota.',
        old.numero, cid
        using errcode = 'raise_exception';
    end if;
  end if;

  return case tg_op when 'DELETE' then old else new end;
end;
$$;

drop trigger if exists municipal_notes_guard_orfao on public.municipal_notes;
create trigger municipal_notes_guard_orfao
  before delete or update of conciliacao_id on public.municipal_notes
  for each row execute function public.guard_faturamento_orfao();

notify pgrst, 'reload schema';
