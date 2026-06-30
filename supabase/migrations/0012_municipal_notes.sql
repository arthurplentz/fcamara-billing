-- ════════════════════════════════════════════════════════════════════════════
-- Conciliação de notas fiscais da prefeitura (NFS-e).
-- Importa o relatório da prefeitura e permite o analista AMARRAR (conciliar)
-- cada nota às receitas reconhecidas. Nada é conciliado automaticamente — o
-- vínculo só é criado por ação do analista (com registro de quem/quando).
-- ════════════════════════════════════════════════════════════════════════════

create table if not exists public.municipal_notes (
  id             uuid primary key default gen_random_uuid(),
  municipio      text,                 -- prefeitura (São Paulo, Maringá, ...)
  numero         text,                 -- Nº NFS-e
  emitida_em     timestamptz,          -- Data/Hora da emissão
  fato_gerador   date,                 -- Data do fato gerador
  prestador_cnpj text,                 -- empresa Fcamara que emitiu
  prestador_nome text,
  tomador_cnpj   text,                 -- CNPJ do cliente (liga ao cadastro)
  tomador_nome   text,
  valor_servicos numeric default 0,    -- valor bruto dos serviços (face da nota)
  valor_total    numeric default 0,    -- valor total recebido
  iss            numeric default 0,
  situacao       text,
  cancelada      boolean default false,
  pedidos        text,                 -- pedidos/OV extraídos da discriminação
  competencias   text,                 -- competências citadas na discriminação
  profissionais  text,                 -- nomes citados na discriminação
  discriminacao  text,
  import_id      uuid,
  created_at     timestamptz not null default now()
);

create index if not exists municipal_notes_tomador_idx on public.municipal_notes(tomador_cnpj);
create index if not exists municipal_notes_numero_idx  on public.municipal_notes(numero);
create index if not exists municipal_notes_import_idx   on public.municipal_notes(import_id);

alter table public.municipal_notes enable row level security;
drop policy if exists municipal_notes_all on public.municipal_notes;
create policy municipal_notes_all on public.municipal_notes
  for all to authenticated using (true) with check (true);

-- Vínculo da conciliação no próprio registro de receita (vários registros podem
-- apontar para a mesma nota). conciliado_em/por servem de auditoria.
alter table public.records add column if not exists municipal_note_id uuid;
alter table public.records add column if not exists conciliado_em      timestamptz;
alter table public.records add column if not exists conciliado_por     text;
create index if not exists records_municipal_note_idx on public.records(municipal_note_id);

notify pgrst, 'reload schema';
