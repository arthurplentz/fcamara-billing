# Guia de transferência (handover) — mudar para as suas contas

Este projeto foi desenvolvido usando contas temporárias (Claude, GitHub e Supabase).
**Todo o código é portável.** Só duas coisas dependem de conta: o endereço do
Supabase e os dados/usuários que vivem dentro dele. Siga o checklist abaixo para
migrar tudo para as suas próprias contas.

---

## 1. Supabase (sua conta)

1. Crie uma conta em <https://supabase.com> e um **novo projeto** (plano Free serve).
2. No **SQL Editor**, rode o arquivo `supabase/migrations/0001_initial_schema.sql`
   deste repositório. Isso recria **todas as tabelas e regras de segurança**, idêntico.
3. Crie o primeiro usuário administrador:
   - **Authentication → Users → Add user** (marque *Auto Confirm User*).
   - No **SQL Editor**, promova-o a admin (troque o e-mail):
     ```sql
     update public.profiles set is_admin = true, name = 'Daniela'
     where id = (select id from auth.users where email = 'EMAIL_DO_ADMIN');
     ```
4. Em **Authentication → URL Configuration**, defina **Site URL** e **Redirect URLs**
   com o endereço final do site (ex.: `https://SEU-USUARIO.github.io/fcamara-billing/`).
   Isso faz o "Esqueci minha senha" funcionar.
5. (Opcional) Configure um **SMTP próprio** em Authentication → Emails para os e-mails
   de redefinição não caírem em spam (o e-mail padrão do Supabase tem limite baixo).
6. Copie de **Project Settings → API**:
   - **Project URL**
   - **publishable key** (ou *anon public*)

## 2. Apontar o app para o seu Supabase

Edite **um único arquivo**: `src/lib/supabase.js`. Troque os dois valores no topo
pela **URL** e **chave** do seu projeto (ou defina as variáveis de ambiente
`VITE_SUPABASE_URL` e `VITE_SUPABASE_ANON_KEY`). É só isso.

> A chave *publishable*/*anon* é pública por natureza (protegida pelas regras de RLS),
> então pode ficar no código. **Nunca** coloque a chave *service_role* aqui.

## 3. GitHub (sua conta)

Escolha um caminho:
- **Transferir** o repositório atual para a sua conta (Settings → Transfer ownership), **ou**
- Criar um **repositório novo** na sua conta e subir este mesmo código.

Depois, ative o **GitHub Pages**: Settings → Pages → Source = **GitHub Actions**.
O fluxo de publicação (`.github/workflows/deploy.yml`) já está pronto e publica a cada
push na `main`.

## 4. Claude (sua conta)

Abra uma sessão do Claude Code na sua conta apontando para o **seu** repositório.
Nada no código depende do Claude — é só ambiente de desenvolvimento.

## 5. Dados reais

Só **depois** de estar nas suas contas, importe os dados reais de faturamento pela
própria tela de **Importar** do app. Enquanto estiver nas contas temporárias, use
**apenas dados de exemplo/fictícios** (nunca dados sensíveis em conta de terceiros).

---

## O que fica salvo onde

| Item | Onde vive | Portável? |
|---|---|---|
| Código do app (`src/`) | Repositório Git | ✅ 100% |
| Esquema do banco | `supabase/migrations/` (código) | ✅ recria com 1 SQL |
| Endereço do Supabase | `src/lib/supabase.js` | ✅ troca em 1 lugar |
| Registros, tarefas, histórico | Banco do Supabase | ⤳ reimportar/migrar |
| Usuários e senhas | Supabase Auth | ⤳ recriar na sua conta |
| Publicação do site | GitHub Pages (Actions) | ✅ reativar nas suas contas |
