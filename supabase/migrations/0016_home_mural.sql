-- ════════════════════════════════════════════════════════════════════════════
-- Tela inicial (Home): apelido do analista + mural da semana (frase + lembretes).
-- ════════════════════════════════════════════════════════════════════════════

alter table public.profiles add column if not exists apelido text;

create table if not exists public.mural (
  id         uuid primary key default gen_random_uuid(),
  frase      text,
  autor      text,
  lembretes  text,           -- JSON: ["lembrete 1", "lembrete 2", ...]
  updated_at timestamptz not null default now()
);

alter table public.mural enable row level security;
drop policy if exists mural_read  on public.mural;
drop policy if exists mural_write on public.mural;
create policy mural_read  on public.mural for select to authenticated using (true);
create policy mural_write on public.mural for all to authenticated using (public.is_admin()) with check (public.is_admin());

notify pgrst, 'reload schema';
