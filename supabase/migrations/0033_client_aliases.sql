-- ════════════════════════════════════════════════════════════════════════════
-- DE → PARA de clientes (unificação de nomes).
-- Às vezes o MESMO cliente aparece com nomes totalmente diferentes nos registros
-- (ex.: "alocacao mandic" e "SOCIEDADE REGIONAL DE ENSINO E SAUDE LTDA"). A
-- canonicalização automática do app só junta variações parecidas (prefixos), não
-- nomes distintos. Esta tabela guarda o mapa manual: cada nome de ORIGEM (de)
-- passa a ser lido como o nome CANÔNICO (para).
--
-- Aplicado na LEITURA (o app reescreve o cliente ao carregar) — não altera o dado
-- bruto e por isso SOBREVIVE a re-importações. Só o admin edita.
-- Casamento por nome normalizado (minúsculas/aparado), então variação de caixa
-- não importa.
-- ════════════════════════════════════════════════════════════════════════════

create table if not exists public.client_aliases (
  de         text primary key,          -- nome como aparece nos registros (origem)
  para       text not null,             -- nome canônico (destino)
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

alter table public.client_aliases enable row level security;

drop policy if exists ca_read  on public.client_aliases;
create policy ca_read  on public.client_aliases for select to authenticated using (true);
drop policy if exists ca_write on public.client_aliases;
create policy ca_write on public.client_aliases for all to authenticated
  using (public.is_admin()) with check (public.is_admin());

notify pgrst, 'reload schema';
