-- ════════════════════════════════════════════════════════════════════════════
-- Anti-duplicação de notas da prefeitura.
-- A identidade real de uma NFS-e é (numero + prestador_cnpj) — campos que vêm do
-- próprio relatório. A `empresa` era escolhida na importação e variava (NULL vs
-- BR02), o que furava o dedup do app e duplicava a base inteira (355 cópias).
--
-- PRÉ-REQUISITO: rode a limpeza de duplicatas ANTES desta migration (o índice
-- único falha se ainda houver duplicata). Confira com:
--   select numero, prestador_cnpj, count(*) from public.municipal_notes
--   group by numero, prestador_cnpj having count(*) > 1;
-- Deve retornar 0 linhas.
-- ════════════════════════════════════════════════════════════════════════════

-- Índice único parcial: só vale quando prestador_cnpj está preenchido (nota sem
-- CNPJ do prestador não entra na regra — evita bloquear cargas incompletas).
create unique index if not exists municipal_notes_numero_prestador_uidx
  on public.municipal_notes (numero, prestador_cnpj)
  where prestador_cnpj is not null and prestador_cnpj <> '';

notify pgrst, 'reload schema';
