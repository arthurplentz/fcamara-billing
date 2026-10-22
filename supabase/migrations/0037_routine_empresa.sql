-- Empresa do grupo a que a rotina se refere (FC Holding, Fcamara, Omnik, ...).
-- "Todas" = vale para todas as empresas. Usado para classificar e filtrar o quadro.
alter table public.routines add column if not exists empresa text;

notify pgrst, 'reload schema';
