-- ════════════════════════════════════════════════════════════════════════════
-- Dados de exemplo (opcional). Rode DEPOIS de criar o primeiro usuário admin.
-- Os registros não dependem de profiles; servem para validar telas e relatórios.
-- ════════════════════════════════════════════════════════════════════════════

insert into public.records
  (responsavel, empresa, tipo, cod_cliente, cliente, pep, inicio, fim, profissional, valor_venda, hrs_aprovadas, valor_total, valor_liquido, competencia)
values
  ('Fernanda',    'BR02','Time & Expenses','1002351','Banco ABC Brasil S.A.','BR02CLP00005.1.1','01/05/2026','31/05/2026','Bruna Paz Amorim',     192.5, 160, 30800.00, 11774.70, '05/2026'),
  ('Fernanda',    'BR02','Time & Expenses','1002351','Banco ABC Brasil S.A.','BR02CLP00005.1.1','01/05/2026','31/05/2026','Gilliard Costa Santos',145.2, 156, 22651.20, 21167.55, '05/2026'),
  ('Fernanda',    'BR02','Time & Expenses','1002840','Banco BS2 S.A.',        'BR02CLP00100.1.1','01/05/2026','31/05/2026','Emerson França',       217.0, 160, 34720.00, 32445.84, '05/2026'),
  ('Layza Arruda','BR02','Time & Expenses','1002100','Diagnósticos da América S.A.','BR02CLP00041','01/05/2026','31/05/2026','Adriano Silva Gama', 135.0, 168, 22680.00, 21194.46, '05/2026'),
  ('Layza Arruda','BR02','Time & Expenses','1002418','Grupo Casas Bahia S.A.','BR02CLP00042.0.3','01/05/2026','31/05/2026','Abel de Meira Junior',172.29,168, 28944.72, 27048.84, '05/2026');

insert into public.tasks (title, descricao, due_date, assignee, status)
values
  ('Extrair base de T&E de junho','Baixar relatório no FC Team e validar colunas.','2026-07-03','Fernanda',    'todo'),
  ('Cobrar retorno do comercial — Casas Bahia','Acompanhar aprovação dos valores ajustados.','2026-07-01','Layza Arruda','doing'),
  ('Fechar NFs dentro do corte','Garantir emissão das notas aprovadas antes do corte.','2026-06-30','Daniela',     'inbox');
