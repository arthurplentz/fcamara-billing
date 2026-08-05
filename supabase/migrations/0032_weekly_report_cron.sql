-- ════════════════════════════════════════════════════════════════════════════
-- AGENDAMENTO do report semanal do comercial (pg_cron → Edge Function).
--
-- Esta migração deixa o agendamento PRONTO PARA LIGAR. O disparo real depende de:
--   1. A Edge Function publicada:  supabase functions deploy weekly-commercial-report
--   2. Os segredos definidos:      RESEND_API_KEY e REPORT_FROM
--      (ver supabase/functions/weekly-commercial-report/index.ts)
--
-- Como o comando de agendamento precisa da URL do seu projeto e da service-role
-- key (que NUNCA devem ficar versionadas no repositório), ele fica abaixo como
-- template comentado. Rode-o UMA VEZ no SQL Editor, trocando os dois <placeholders>.
-- A extensão é habilitada aqui (idempotente) para não travar quem não usa cron.
-- ════════════════════════════════════════════════════════════════════════════

create extension if not exists pg_cron;
create extension if not exists pg_net;

-- ── Rode este bloco no SQL Editor depois do deploy da função ──────────────────
-- (troque <PROJECT_REF> e <SERVICE_ROLE_KEY>). Envia toda segunda-feira 08:00 UTC.
--
--   select cron.schedule(
--     'weekly-commercial-report',
--     '0 8 * * 1',
--     $$
--       select net.http_post(
--         url     := 'https://<PROJECT_REF>.functions.supabase.co/weekly-commercial-report',
--         headers := jsonb_build_object(
--           'Content-Type',  'application/json',
--           'Authorization', 'Bearer <SERVICE_ROLE_KEY>'
--         ),
--         body    := '{}'::jsonb
--       );
--     $$
--   );
--
-- Para remover:   select cron.unschedule('weekly-commercial-report');
-- Para ver:       select * from cron.job;
