// ════════════════════════════════════════════════════════════════════════════
// Disparo do REPORT SEMANAL para o comercial (Supabase Edge Function).
//
// Para cada perfil comercial (is_comercial = true) com e-mail e BU definidos,
// monta o report da BU dele — reconhecido, faturado, represado e o detalhe por
// tipo de projeto (Time & Expenses com a linha de cada consultor: valor/hora,
// horas e total) — e envia por e-mail via Resend.
//
// Isto é a AUTOMAÇÃO ("disparo"). O mesmo conteúdo também vive dentro do app,
// na tela "Report semanal" do acesso comercial (para ver/baixar/enviar na hora).
//
// ── Como ativar (uma vez) ────────────────────────────────────────────────────
//   1. Deploy:   supabase functions deploy weekly-commercial-report
//   2. Segredos: supabase secrets set RESEND_API_KEY=re_xxx REPORT_FROM="Faturamento <faturamento@grupofcamara.com>"
//      (SUPABASE_URL e SUPABASE_SERVICE_ROLE_KEY já vêm do ambiente da função.)
//   3. Agende semanalmente — ver supabase/migrations/0032_weekly_report_cron.sql
//   Teste manual (envia agora):
//      curl -X POST "$SUPABASE_URL/functions/v1/weekly-commercial-report" \
//           -H "Authorization: Bearer $SUPABASE_SERVICE_ROLE_KEY"
// ════════════════════════════════════════════════════════════════════════════
import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

const BRL = (n: number) =>
  (n || 0).toLocaleString("pt-BR", { style: "currency", currency: "BRL" });

Deno.serve(async (req) => {
  try {
    const SUPABASE_URL = Deno.env.get("SUPABASE_URL")!;
    const SERVICE_KEY = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;
    const RESEND_API_KEY = Deno.env.get("RESEND_API_KEY");
    const FROM = Deno.env.get("REPORT_FROM") || "Faturamento <onboarding@resend.dev>";
    const db = createClient(SUPABASE_URL, SERVICE_KEY);

    // Permite disparar uma BU só (teste) via ?bu=BU%20Retail
    const url = new URL(req.url);
    const onlyBu = url.searchParams.get("bu") || "";

    // Comerciais ativos (com e-mail e BU). O e-mail vem do auth.users.
    let q = db.from("profiles").select("id,name,bu,is_comercial").eq("is_comercial", true);
    if (onlyBu) q = q.eq("bu", onlyBu);
    const { data: comerciais, error: pErr } = await q;
    if (pErr) throw pErr;
    if (!comerciais?.length) return json({ ok: true, sent: 0, msg: "Nenhum comercial ativo." });

    const { data: records, error: rErr } = await db.from("records").select("*");
    if (rErr) throw rErr;
    const { data: fats } = await db.from("record_faturamentos").select("record_id,valor");
    const { data: vars } = await db.from("receita_variacoes").select("record_id,valor");
    const fatBy: Record<string, number> = {};
    (fats || []).forEach((f) => (fatBy[f.record_id] = (fatBy[f.record_id] || 0) + Number(f.valor || 0)));
    const varBy: Record<string, number> = {};
    (vars || []).forEach((v) => (varBy[v.record_id] = (varBy[v.record_id] || 0) + Number(v.valor || 0)));

    const results: any[] = [];
    for (const c of comerciais) {
      if (!c.bu) continue;
      const { data: authUser } = await db.auth.admin.getUserById(c.id);
      const email = authUser?.user?.email;
      if (!email) { results.push({ bu: c.bu, skipped: "sem e-mail" }); continue; }

      const recs = (records || []).filter((r) => r.bu === c.bu);
      const html = buildHtml(c.bu, c.name || "", recs, fatBy, varBy);

      if (!RESEND_API_KEY) { results.push({ bu: c.bu, email, skipped: "RESEND_API_KEY ausente (deploy sem chave)" }); continue; }
      const resp = await fetch("https://api.resend.com/emails", {
        method: "POST",
        headers: { Authorization: `Bearer ${RESEND_API_KEY}`, "Content-Type": "application/json" },
        body: JSON.stringify({
          from: FROM,
          to: [email],
          subject: `Report semanal — ${c.bu}`,
          html,
        }),
      });
      results.push({ bu: c.bu, email, ok: resp.ok, status: resp.status });
    }
    return json({ ok: true, sent: results.filter((r) => r.ok).length, results });
  } catch (e) {
    return json({ ok: false, error: String(e?.message || e) }, 500);
  }
});

function json(body: unknown, status = 200) {
  return new Response(JSON.stringify(body), { status, headers: { "Content-Type": "application/json" } });
}

function buildHtml(
  bu: string,
  nome: string,
  recs: any[],
  fatBy: Record<string, number>,
  varBy: Record<string, number>,
): string {
  const bill = (r: any) => Number(r.valor_total || 0) + (varBy[r.id] || 0);
  const fat = (r: any) => fatBy[r.id] || 0;
  const rep = (r: any) => bill(r) - fat(r);
  const tot = recs.reduce(
    (a, r) => ({ rec: a.rec + bill(r), fat: a.fat + fat(r), rep: a.rep + Math.max(0, rep(r)) }),
    { rec: 0, fat: 0, rep: 0 },
  );

  // Detalhe por tipo de projeto.
  const tipos: Record<string, any[]> = {};
  recs.forEach((r) => { (tipos[r.tipo || "—"] = tipos[r.tipo || "—"] || []).push(r); });

  const secoes = Object.entries(tipos).map(([tipo, list]) => {
    const isTE = /time|expense|t&e/i.test(tipo);
    if (isTE) {
      // T&E: linha por consultor (valor/hora, horas, total).
      const rows = list
        .sort((a, b) => bill(b) - bill(a))
        .map((r) => `<tr>
          <td>${esc(r.cliente)}</td><td>${esc(r.profissional || "—")}</td>
          <td style="text-align:right">${BRL(Number(r.valor_venda || 0))}</td>
          <td style="text-align:right">${Number(r.hrs_aprovadas || 0).toLocaleString("pt-BR")}</td>
          <td style="text-align:right">${BRL(bill(r))}</td>
          <td style="text-align:right">${BRL(fat(r))}</td>
          <td style="text-align:right;color:${rep(r) > 0.01 ? "#c2410c" : "#94a3b8"}">${BRL(Math.max(0, rep(r)))}</td>
        </tr>`).join("");
      return `<h3>${esc(tipo)} — por consultor</h3>
        <table><thead><tr>
          <th>Cliente</th><th>Consultor</th><th style="text-align:right">Valor/hora</th>
          <th style="text-align:right">Horas</th><th style="text-align:right">Total</th>
          <th style="text-align:right">Faturado</th><th style="text-align:right">Represado</th>
        </tr></thead><tbody>${rows}</tbody></table>`;
    }
    // Demais tipos: consolidado por cliente.
    const byCli: Record<string, any> = {};
    list.forEach((r) => {
      const k = r.cliente || "—";
      const g = (byCli[k] = byCli[k] || { rec: 0, fat: 0, rep: 0 });
      g.rec += bill(r); g.fat += fat(r); g.rep += Math.max(0, rep(r));
    });
    const rows = Object.entries(byCli).map(([cli, g]: any) => `<tr>
      <td>${esc(cli)}</td>
      <td style="text-align:right">${BRL(g.rec)}</td>
      <td style="text-align:right">${BRL(g.fat)}</td>
      <td style="text-align:right;color:${g.rep > 0.01 ? "#c2410c" : "#94a3b8"}">${BRL(g.rep)}</td>
    </tr>`).join("");
    return `<h3>${esc(tipo)}</h3>
      <table><thead><tr><th>Cliente</th><th style="text-align:right">Reconhecido</th>
      <th style="text-align:right">Faturado</th><th style="text-align:right">Represado</th></tr></thead>
      <tbody>${rows}</tbody></table>`;
  }).join("");

  return `<!doctype html><html><body style="font-family:Arial,Helvetica,sans-serif;color:#1f2937;max-width:760px;margin:0 auto;padding:16px">
    <div style="border-bottom:3px solid #ef5a2d;padding-bottom:10px;margin-bottom:16px">
      <div style="font-size:12px;color:#94a3b8;text-transform:uppercase;letter-spacing:.1em">Report semanal · ${esc(bu)}</div>
      <div style="font-size:20px;font-weight:800">Olá${nome ? ", " + esc(nome.split(" ")[0]) : ""} — sua receita da semana</div>
    </div>
    <table style="width:100%;border-collapse:collapse;margin-bottom:18px">
      <tr>
        ${kpi("Reconhecido", BRL(tot.rec), "#1f2937")}
        ${kpi("Faturado", BRL(tot.fat), "#15803d")}
        ${kpi("A faturar", BRL(tot.rec - tot.fat), "#c2410c")}
        ${kpi("Represado", BRL(tot.rep), "#b91c1c")}
      </tr>
    </table>
    <style>table{width:100%;border-collapse:collapse;margin:6px 0 18px;font-size:13px}
      th{background:#f8fafc;text-align:left;padding:7px 9px;border-bottom:1px solid #e2e8f0;color:#64748b;font-size:11px}
      td{padding:7px 9px;border-bottom:1px solid #f1f5f9}h3{font-size:14px;margin:18px 0 4px}</style>
    ${secoes || "<p>Sem receitas classificadas nesta BU.</p>"}
    <p style="color:#94a3b8;font-size:11px;margin-top:24px">Gerado automaticamente pelo sistema de Order to Cash · Grupo Fcamara.</p>
  </body></html>`;
}

const kpi = (label: string, valor: string, cor: string) =>
  `<td style="padding:12px;border:1px solid #e2e8f0;border-radius:8px;text-align:center">
    <div style="font-size:11px;color:#64748b">${label}</div>
    <div style="font-size:16px;font-weight:800;color:${cor}">${valor}</div>
  </td>`;

const esc = (s: string) =>
  String(s || "").replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]!));
