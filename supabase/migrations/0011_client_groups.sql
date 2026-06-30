-- ════════════════════════════════════════════════════════════════════════════
-- Grupos de clientes: um cadastro pode reunir vários CNPJs (empresas do grupo).
-- cnpjs guarda JSON: [{ razao, cnpj, codSap }, ...]. As colunas cnpj/cod_sap
-- continuam com o "principal" (primeiro da lista), para busca e exibição.
-- ════════════════════════════════════════════════════════════════════════════

alter table public.clients add column if not exists cnpjs text;
