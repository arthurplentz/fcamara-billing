// ════════════════════════════════════════════════════════════════════════════
// ASSISTENTE O2C (Supabase Edge Function)
//
// Chat de perguntas sobre as premissas de faturamento dos clientes/projetos.
// A base de conhecimento é o Notion "O2C – Base de Conhecimento" (bases Projetos
// e Clientes + guia de preenchimento). A IA é o Claude (API da Anthropic).
//
// Segurança:
//   • As chaves (ANTHROPIC_API_KEY, NOTION_TOKEN) ficam SÓ nos segredos do
//     Supabase — nunca no código nem no navegador.
//   • Só responde a usuário logado no app (JWT do Supabase Auth).
//   • Perfil COMERCIAL não usa o assistente (ele vê só a BU dele no app).
//   • A IA só LÊ o Notion. Não altera nada.
//
// ── Como ativar (uma vez) ────────────────────────────────────────────────────
//   Ver supabase/functions/assistente-o2c/README.md
// ════════════════════════════════════════════════════════════════════════════
import { createClient } from "https://esm.sh/@supabase/supabase-js@2";

const NOTION_VERSION = "2025-09-03";
const env = (k: string, d = "") => Deno.env.get(k) || d;

// IDs da base no Notion (workspace da Daniela). Podem ser trocados por segredo.
const DS_PROJETOS = () => env("NOTION_PROJETOS_DS", "e57ff558-cd1e-4926-8ad8-a1d904665bbf");
const DS_CLIENTES = () => env("NOTION_CLIENTES_DS", "dbd8ae2f-5d58-4aa9-b23d-59c7de2fabd1");
const PAGINAS: Record<string, string> = {
  guia: env("NOTION_PAGE_GUIA", "3fa17493-e195-815c-abaa-f02aaf87f0f7"),
  processo_ecommerce: env("NOTION_PAGE_ECOMMERCE", "3f817493-e195-818f-ae5b-efe80cca6f2a"),
  principal: env("NOTION_PAGE_PRINCIPAL", "3f417493-e195-81f3-97ab-e8b561c29c37"),
};

const CORS = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Headers": "authorization, x-client-info, apikey, content-type",
  "Access-Control-Allow-Methods": "POST, OPTIONS",
};
function json(body: unknown, status = 200) {
  return new Response(JSON.stringify(body), { status, headers: { ...CORS, "Content-Type": "application/json" } });
}

// ─── Notion ──────────────────────────────────────────────────────────────────
async function notion(path: string, body?: unknown) {
  const r = await fetch(`https://api.notion.com/v1/${path}`, {
    method: body ? "POST" : "GET",
    headers: {
      Authorization: `Bearer ${env("NOTION_TOKEN")}`,
      "Notion-Version": NOTION_VERSION,
      "Content-Type": "application/json",
    },
    body: body ? JSON.stringify(body) : undefined,
  });
  const data = await r.json();
  if (!r.ok) throw new Error(`Notion ${r.status}: ${data?.message || "erro"}`);
  return data;
}

// Converte uma propriedade do Notion em texto simples.
// deno-lint-ignore no-explicit-any
function valor(p: any): string | string[] {
  if (!p) return "";
  switch (p.type) {
    case "title":
    case "rich_text": return (p[p.type] || []).map((t: any) => t.plain_text).join("");
    case "select": return p.select?.name ?? "";
    case "status": return p.status?.name ?? "";
    case "multi_select": return (p.multi_select || []).map((x: any) => x.name).join(", ");
    case "checkbox": return p.checkbox ? "Sim" : "";
    case "number": return p.number == null ? "" : String(p.number);
    case "date": return p.date ? p.date.start + (p.date.end ? ` a ${p.date.end}` : "") : "";
    case "formula": {
      const f = p.formula || {};
      if (f.type === "string") return f.string ?? "";
      if (f.type === "number") return f.number == null ? "" : String(f.number);
      if (f.type === "boolean") return f.boolean ? "Sim" : "Não";
      if (f.type === "date") return f.date?.start ?? "";
      return "";
    }
    case "relation": return (p.relation || []).map((r: any) => r.id);
    case "people": return (p.people || []).map((x: any) => x.name).join(", ");
    case "url": case "email": case "phone_number": return p[p.type] ?? "";
    default: return "";
  }
}

// Campos que não ajudam a responder (relações internas, controle).
const IGNORAR = new Set(["Tarefas", "Alocações", "Exceções", "Fechamentos", "Analista responsável", "Projetos"]);

// Página do Notion → objeto só com campos preenchidos (vazio = não cadastrado).
// deno-lint-ignore no-explicit-any
function achatar(page: any) {
  const o: Record<string, unknown> = { _id: page.id, _url: page.url };
  for (const [nome, p] of Object.entries(page.properties || {})) {
    if (IGNORAR.has(nome)) continue;
    const v = valor(p);
    if (Array.isArray(v) ? v.length : v !== "") o[nome] = v;
  }
  return o;
}

// Cache de nomes de clientes (por requisição).
const nomesCliente = new Map<string, string>();
async function nomeCliente(id: string) {
  if (nomesCliente.has(id)) return nomesCliente.get(id)!;
  try {
    const pg = await notion(`pages/${id}`);
    const t = Object.values(pg.properties || {}).find((p: any) => p.type === "title");
    const n = String(valor(t) || "");
    nomesCliente.set(id, n);
    return n;
  } catch { return ""; }
}
async function comCliente(rows: Record<string, unknown>[]) {
  for (const r of rows) {
    const ids = r["Cliente"];
    if (Array.isArray(ids) && ids.length) r["Cliente"] = (await Promise.all(ids.map((i) => nomeCliente(String(i))))).join(", ");
  }
  return rows;
}

async function consultar(ds: string, filter: unknown, limite = 100) {
  const out: any[] = [];
  let cursor: string | undefined;
  do {
    const body: Record<string, unknown> = { page_size: 100 };
    if (filter) body.filter = filter;
    if (cursor) body.start_cursor = cursor;
    const r = await notion(`data_sources/${ds}/query`, body);
    out.push(...(r.results || []));
    cursor = r.has_more ? r.next_cursor : undefined;
  } while (cursor && out.length < limite);
  return out.slice(0, limite);
}

// ─── Ferramentas que a IA pode usar ──────────────────────────────────────────
const SELECTS = [
  "Tipo de contrato", "Formato de horas", "Ciclo de apuração", "Situação de faturamento", "Status do cadastro",
  "Racional: data fixa?", "NF: data fixa?", "PO no corpo da nota?", "Padrão no corpo da nota?", "Forma de pagamento",
  "Portal de medição?", "Portal de notas?", "Modelo de racional próprio?", "Controle de hora extra?", "Vertical",
];

const TOOLS = [
  {
    name: "buscar_projetos",
    description: "Busca projetos pelo nome do projeto, nome do cliente, código PEP (ex.: BR02CLP000120), gerente ou ordem do cliente. Devolve todos os campos cadastrados de cada projeto (premissas de faturamento). Use primeiro para qualquer pergunta sobre um cliente ou projeto específico.",
    input_schema: { type: "object", properties: { termo: { type: "string", description: "Parte do nome do cliente/projeto, PEP ou gerente. Use poucas palavras (ex.: 'Corações', 'Bunzl', 'BR02CLP00081')." } }, required: ["termo"] },
  },
  {
    name: "listar_projetos",
    description: "Lista projetos com filtros, para perguntas gerais (ex.: 'quais projetos da Viviane exigem PO na nota', 'quais clientes são hora fechada', 'quais cadastros estão pendentes'). Devolve campos resumidos.",
    input_schema: {
      type: "object",
      properties: {
        gerente: { type: "string", description: "Parte do nome do gerente/analista no SAP (ex.: 'Viviane')." },
        bu: { type: "string", enum: ["Retail", "Multisector", "Logistics", "Hyper", "Health", "Finance", "Others"] },
        campo: { type: "string", enum: SELECTS, description: "Campo de opção para filtrar." },
        valor: { type: "string", description: "Valor exato da opção (ex.: 'Sim', 'Hora fechada', 'T&M'). Use 'VAZIO' para listar quem ainda não respondeu." },
        incluir_inativos: { type: "boolean", description: "Padrão: só projetos ativos no SAP." },
      },
    },
  },
  {
    name: "ler_pagina",
    description: "Lê uma página de referência do Notion: 'guia' (o que significa cada campo e como preencher), 'processo_ecommerce' (passo a passo de faturamento da vertical e-commerce) ou 'principal' (visão geral da base).",
    input_schema: { type: "object", properties: { pagina: { type: "string", enum: Object.keys(PAGINAS) } }, required: ["pagina"] },
  },
];

async function buscarProjetos(termo: string) {
  const t = termo.trim();
  const porProjeto = await consultar(DS_PROJETOS(), {
    or: [
      { property: "Projeto", title: { contains: t } },
      { property: "Código do projeto", rich_text: { contains: t } },
      { property: "Gerente de projetos", rich_text: { contains: t } },
      { property: "Ordem do cliente", rich_text: { contains: t } },
    ],
  }, 25);
  // Também pelo nome do cliente (base Clientes → projetos ligados).
  const clientes = await consultar(DS_CLIENTES(), { property: "Cliente", title: { contains: t } }, 5);
  const ids = new Set(porProjeto.map((p) => p.id));
  const extras: any[] = [];
  for (const c of clientes) {
    for (const rel of c.properties?.Projetos?.relation || []) {
      if (!ids.has(rel.id) && extras.length < 15) { ids.add(rel.id); extras.push(await notion(`pages/${rel.id}`)); }
    }
  }
  const rows = await comCliente([...porProjeto, ...extras].slice(0, 25).map(achatar));
  return rows.length ? rows : { aviso: `Nenhum projeto encontrado para "${t}". Tente outra parte do nome.` };
}

async function listarProjetos(a: { gerente?: string; bu?: string; campo?: string; valor?: string; incluir_inativos?: boolean }) {
  const and: unknown[] = [];
  if (!a.incluir_inativos) and.push({ property: "Situação na lista", select: { equals: "Ativo" } });
  if (a.gerente) and.push({ property: "Gerente de projetos", rich_text: { contains: a.gerente } });
  if (a.bu) and.push({ property: "BU", select: { equals: a.bu } });
  if (a.campo && SELECTS.includes(a.campo)) {
    and.push(a.valor && a.valor.toUpperCase() !== "VAZIO"
      ? { property: a.campo, select: { equals: a.valor } }
      : { property: a.campo, select: { is_empty: true } });
  }
  const pages = await consultar(DS_PROJETOS(), and.length ? { and } : undefined, 300);
  const resumo = pages.map((p) => {
    const f = achatar(p);
    const r: Record<string, unknown> = {
      Projeto: f["Projeto"], PEP: f["Código do projeto"], Gerente: f["Gerente de projetos"], BU: f["BU"],
      Pendências: f["Pendências do projeto"],
    };
    if (a.campo) r[a.campo] = f[a.campo] ?? "(não cadastrado)";
    return r;
  });
  return { total: resumo.length, projetos: resumo };
}

async function lerPagina(pagina: string) {
  const id = PAGINAS[pagina];
  if (!id) return { erro: "Página desconhecida." };
  const linhas: string[] = [];
  let cursor: string | undefined;
  do {
    const r = await notion(`blocks/${id}/children?page_size=100${cursor ? `&start_cursor=${cursor}` : ""}`);
    for (const b of r.results || []) {
      const blk = b[b.type] || {};
      if (b.type === "table") {
        const rows = await notion(`blocks/${b.id}/children?page_size=100`);
        for (const tr of rows.results || []) {
          linhas.push((tr.table_row?.cells || []).map((c: any[]) => c.map((t) => t.plain_text).join("")).join(" | "));
        }
        continue;
      }
      const txt = (blk.rich_text || []).map((t: any) => t.plain_text).join("");
      if (txt) linhas.push(txt);
    }
    cursor = r.has_more ? r.next_cursor : undefined;
  } while (cursor);
  return linhas.join("\n").slice(0, 30000);
}

// deno-lint-ignore no-explicit-any
async function executar(nome: string, input: any) {
  try {
    if (nome === "buscar_projetos") return await buscarProjetos(String(input.termo || ""));
    if (nome === "listar_projetos") return await listarProjetos(input || {});
    if (nome === "ler_pagina") return await lerPagina(String(input.pagina || ""));
    return { erro: "Ferramenta desconhecida." };
  } catch (e) {
    return { erro: (e as Error).message };
  }
}

// ─── Claude ──────────────────────────────────────────────────────────────────
const SISTEMA = (usuario: string) => `Você é o Assistente O2C da FCamara: ajuda o time de faturamento (order to cash) a consultar as premissas de faturamento de cada cliente e projeto.

Regras:
- Responda SEMPRE com base no que as ferramentas devolverem da base do Notion. Nunca invente regras, datas, e-mails, valores ou taxas.
- Campo que não aparece no resultado = não cadastrado. Diga isso claramente ("Isso ainda não está cadastrado no projeto X") e sugira preencher na aba 📝 do analista na base Projetos.
- Cite sempre o projeto e o código PEP de onde tirou a informação.
- Se a busca trouxer mais de um projeto do mesmo cliente e as regras forem diferentes, mostre a diferença. Se for ambíguo, pergunte qual projeto.
- Projeto só é inativo quando sai da lista de ativos do SAP ("Situação na lista").
- Responda em português do Brasil, direto e curto, em linguagem de quem fatura. Use lista quando houver vários itens.
- Você só consulta. Não consegue alterar o Notion; se pedirem para mudar algo, explique onde a pessoa altera.
- Assuntos fora de faturamento/O2C da FCamara: diga educadamente que só ajuda com isso.

Quem está perguntando: ${usuario}. Data de hoje: ${new Date().toISOString().slice(0, 10)}.`;

async function claude(body: unknown) {
  const r = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "x-api-key": env("ANTHROPIC_API_KEY"),
      "anthropic-version": "2023-06-01",
      "content-type": "application/json",
    },
    body: JSON.stringify(body),
  });
  const data = await r.json();
  if (!r.ok) throw new Error(`Claude ${r.status}: ${data?.error?.message || "erro"}`);
  return data;
}

// ─── Handler ─────────────────────────────────────────────────────────────────
Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: CORS });
  try {
    if (!env("ANTHROPIC_API_KEY") || !env("NOTION_TOKEN")) {
      return json({ erro: "Assistente ainda não configurado: faltam os segredos ANTHROPIC_API_KEY e/ou NOTION_TOKEN no Supabase." }, 503);
    }

    // 1) Quem está pedindo? (precisa estar logado no app)
    const SUPABASE_URL = env("SUPABASE_URL");
    const authHeader = req.headers.get("Authorization") || "";
    const asUser = createClient(SUPABASE_URL, env("SUPABASE_ANON_KEY"), { global: { headers: { Authorization: authHeader } } });
    const { data: u } = await asUser.auth.getUser();
    if (!u?.user) return json({ erro: "Faça login no sistema para usar o assistente." }, 401);

    const admin = createClient(SUPABASE_URL, env("SUPABASE_SERVICE_ROLE_KEY"));
    const { data: prof } = await admin.from("profiles").select("*").eq("id", u.user.id).single();
    if (prof?.is_comercial) return json({ erro: "O assistente é de uso do time de faturamento." }, 403);
    const usuario = prof?.name || u.user.email || "usuário";

    // 2) Conversa (mantém só as últimas mensagens para controlar custo)
    const { messages } = await req.json();
    if (!Array.isArray(messages) || !messages.length) return json({ erro: "Mensagem vazia." }, 400);
    // deno-lint-ignore no-explicit-any
    const conversa: any[] = messages.slice(-12)
      .filter((m: any) => (m.role === "user" || m.role === "assistant") && typeof m.content === "string" && m.content.trim())
      .map((m: any) => ({ role: m.role, content: m.content.slice(0, 4000) }));
    if (!conversa.length || conversa[0].role !== "user") conversa.shift();
    const pergunta = conversa[conversa.length - 1]?.content || "";

    // 3) Loop de ferramentas (a IA consulta o Notion quantas vezes precisar, com limite)
    const fontes = new Map<string, { projeto: string; pep: string; url: string }>();
    let resposta = "";
    let uso = { input_tokens: 0, output_tokens: 0 };
    for (let i = 0; i < 6; i++) {
      const r = await claude({
        model: env("ANTHROPIC_MODEL", "claude-sonnet-5-5"),
        max_tokens: 1500,
        system: SISTEMA(usuario),
        tools: TOOLS,
        messages: conversa,
      });
      uso.input_tokens += r.usage?.input_tokens || 0;
      uso.output_tokens += r.usage?.output_tokens || 0;
      conversa.push({ role: "assistant", content: r.content });
      if (r.stop_reason !== "tool_use") {
        resposta = (r.content || []).filter((c: any) => c.type === "text").map((c: any) => c.text).join("\n").trim();
        break;
      }
      const resultados = [];
      for (const c of r.content) {
        if (c.type !== "tool_use") continue;
        const out = await executar(c.name, c.input);
        const lista = Array.isArray(out) ? out : (out as any)?.projetos;
        if (Array.isArray(lista)) for (const p of lista.slice(0, 30)) {
          const pep = String(p["Código do projeto"] || p["PEP"] || "");
          if (pep && p._url) fontes.set(pep, { projeto: String(p["Projeto"] || ""), pep, url: String(p._url) });
        }
        resultados.push({ type: "tool_result", tool_use_id: c.id, content: typeof out === "string" ? out : JSON.stringify(out) });
      }
      conversa.push({ role: "user", content: resultados });
    }
    if (!resposta) resposta = "Não consegui concluir a consulta. Tente reformular a pergunta, citando o cliente ou o PEP.";

    // 4) Log (auditoria e para ver o que o time mais pergunta) — não bloqueia a resposta
    try {
      await admin.from("assistant_logs").insert({
        user_id: u.user.id, usuario, pergunta, resposta,
        input_tokens: uso.input_tokens, output_tokens: uso.output_tokens,
      });
    } catch { /* tabela ainda não criada: segue sem log */ }

    return json({ resposta, fontes: [...fontes.values()].slice(0, 8) });
  } catch (e) {
    return json({ erro: (e as Error).message }, 500);
  }
});
