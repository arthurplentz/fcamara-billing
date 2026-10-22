-- Tarefa pontual (uma vez só): guarda a data específica da ocorrência.
-- cadencia = 'pontual' usa esta data; não recorre e não reinicia.
alter table public.routines add column if not exists data_unica date;

notify pgrst, 'reload schema';
