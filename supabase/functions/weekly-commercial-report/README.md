# Report semanal do comercial — automação (disparo)

Envia, por e-mail, o report semanal de cada BU para o respectivo acesso comercial
(`is_comercial = true` com `bu` e e-mail). O mesmo conteúdo existe dentro do app,
na tela **Report semanal** do acesso comercial — esta função é só o *disparo
automático*.

## Ativar (uma vez)

1. **Deploy da função**
   ```bash
   supabase functions deploy weekly-commercial-report
   ```
2. **Segredos** (provedor de e-mail — usamos [Resend](https://resend.com)):
   ```bash
   supabase secrets set RESEND_API_KEY=re_xxxxx \
                        REPORT_FROM="Faturamento <faturamento@grupofcamara.com>"
   ```
   > `SUPABASE_URL` e `SUPABASE_SERVICE_ROLE_KEY` já vêm do ambiente da função.
3. **Agendar** — rode o bloco comentado em
   `supabase/migrations/0032_weekly_report_cron.sql` no SQL Editor (troca de dois
   placeholders). Padrão: segunda-feira 08:00 UTC.

## Testar agora (sem esperar a segunda)

```bash
curl -X POST "$SUPABASE_URL/functions/v1/weekly-commercial-report" \
     -H "Authorization: Bearer $SUPABASE_SERVICE_ROLE_KEY"
# uma BU só:
curl -X POST "$SUPABASE_URL/functions/v1/weekly-commercial-report?bu=BU%20Retail" \
     -H "Authorization: Bearer $SUPABASE_SERVICE_ROLE_KEY"
```

Sem `RESEND_API_KEY`, a função **não quebra**: ela responde quantos e-mails
enviaria e por quê pulou (útil para validar antes de plugar o provedor).
