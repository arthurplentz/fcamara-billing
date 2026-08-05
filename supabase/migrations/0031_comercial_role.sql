-- ════════════════════════════════════════════════════════════════════════════
-- PAPEL "COMERCIAL" (acesso do diretor/gestor comercial da BU).
-- Um comercial NÃO é analista nem admin: ele entra e vê apenas a RECEITA DA SUA
-- BU (uma BU por comercial) — Dashboard, Visão por projeto e o Report semanal.
-- Não altera nada. A trava real é aqui no RLS; a UI apenas mostra a navegação
-- restrita.
--   • is_comercial → marca o perfil como comercial (independente de admin/viewer).
--   • bu           → a única unidade de negócio que este comercial enxerga
--                    (ex.: "BU Retail"). Casada com records.bu.
-- ════════════════════════════════════════════════════════════════════════════

alter table public.profiles add column if not exists is_comercial boolean not null default false;
alter table public.profiles add column if not exists bu text;

-- Helpers SECURITY DEFINER (evitam recursão de RLS ao consultar profiles).
create or replace function public.is_comercial()
returns boolean
language sql stable security definer set search_path = public
as $$
  select coalesce((select p.is_comercial from public.profiles p where p.id = auth.uid()), false);
$$;

create or replace function public.my_bu()
returns text
language sql stable security definer set search_path = public
as $$
  select (select p.bu from public.profiles p where p.id = auth.uid());
$$;

-- ── RECORDS: leitura ─────────────────────────────────────────────────────────
-- Admin e viewer veem tudo; analista vê os seus (por responsável); comercial vê
-- SÓ os registros classificados na sua BU. Sem BU classificada → não vê nada
-- (comportamento seguro).
drop policy if exists records_select on public.records;
create policy records_select on public.records for select to authenticated
  using (
    public.is_admin()
    or public.is_viewer()
    or (public.is_comercial() and bu is not null and bu = public.my_bu())
    or lower(btrim(responsavel)) = lower(btrim(public.current_responsavel()))
  );

-- Escrita de records continua NÃO permitida ao comercial: as policies de
-- update/insert/delete exigem admin ou dono (responsável) — comercial não é
-- nenhum dos dois, então já está bloqueado. (Ver 0028_viewer_role.sql.)

notify pgrst, 'reload schema';
