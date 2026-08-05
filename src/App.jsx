import { useState, useEffect, useRef, useCallback, createContext, useContext, Fragment } from "react";
import * as XLSX from "xlsx";
import { supabase, SITE_URL } from "./lib/supabase";
import * as db from "./lib/db";

const APP_VERSION = "v3.1";
const LS_KEY = "fcamara_billing_v3";
const ADMIN_NAME = "Daniela";

const EMPRESAS = [
  { cod: "BR02", nome: "Fcamara" },
  { cod: "BR04", nome: "Nação Digital" },
  { cod: "BR05", nome: "SGA" },
  { cod: "BR07", nome: "FC Hyperautomation" },
  { cod: "BR08", nome: "Dojo" },
  { cod: "BR09", nome: "Nextgeneration" },
];

const TIPOS_PROJETO = ["Time & Expenses", "Fee", "WIP", "Usage Based"];
// BUs (unidades de negócio) — classificação comercial. Normalmente 1 por cliente.
const BUS = ["BU Health", "BU Multisector", "BU Logistics", "BU Others", "BU Finance", "BU Retail"];

// PEP canônico para JUNÇÃO DE VALORES: o sufixo após o 1º ponto (".1.1", ".0.3"…)
// é variação sistêmica e conta como o MESMO PEP. Ex.: BR02CLP00046.1.1 →
// BR02CLP00046. Use em todo agrupamento/casamento por PEP (import, Visão por
// projeto, Minha visão, Dashboard, filtros). Não altera o dado gravado.
const pepBase = pep => String(pep||"").split(".")[0].trim();

const STEPS = [
  { id: "p1_extrair",      group: 1, label: "P1",  name: "Extrair dados (FC Team)",  type: "check" },
  { id: "p2_racional",     group: 2, label: "P2",  name: "Montar Racional",          type: "check" },
  { id: "p3_envio_com",    group: 3, label: "P3a", name: "Envio ao Comercial",       type: "check" },
  { id: "p3_retorno_com",  group: 3, label: "P3b", name: "Retorno do Comercial",     type: "check" },
  { id: "p3_data_retorno", group: 3, label: "P3c", name: "Data Retorno",             type: "date"  },
  { id: "p4_envio_cli",    group: 4, label: "P4a", name: "Envio ao Cliente",         type: "check" },
  { id: "p4_aprovacao",    group: 4, label: "P4b", name: "Aprovação do Cliente",     type: "check" },
  { id: "p4_data_aprov",   group: 4, label: "P4c", name: "Data Aprovação",           type: "date"  },
  { id: "p5_liberado",     group: 5, label: "P5",  name: "Liberado para faturamento", type: "check" },
];

const STEP_GROUPS = [
  { num: 1, title: "Extração de dados",    short: "Extração",  steps: ["p1_extrair"] },
  { num: 2, title: "Racional",             short: "Racional",  steps: ["p2_racional"] },
  { num: 3, title: "Validação comercial",  short: "Comercial", steps: ["p3_envio_com","p3_retorno_com","p3_data_retorno"] },
  { num: 4, title: "Aprovação do cliente", short: "Cliente",   steps: ["p4_envio_cli","p4_aprovacao","p4_data_aprov"] },
  { num: 5, title: "Liberação p/ faturamento", short: "Liberado", steps: ["p5_liberado"] },
];

// Funil por tipo de projeto. Fee e WIP não passam por comercial nem cliente:
// apenas Extração de dados → Racional → Faturamento (emissão).
const FUNNELS = {
  "Time & Expenses": [1, 2, 3, 4, 5],
  "Usage Based":     [1, 2, 3, 4, 5],
  "Fee":             [1, 2, 5],
  "WIP":             [1, 2, 5],
};
const funnelGroups = (tipo) => { const nums = FUNNELS[tipo] || [1, 2, 3, 4, 5]; return STEP_GROUPS.filter(g => nums.includes(g.num)); };

// "Etapa concluída" por grupo considera apenas os passos do tipo check obrigatórios.
const GROUP_DONE_KEYS = {
  1: ["p1_extrair"],
  2: ["p2_racional"],
  3: ["p3_envio_com", "p3_retorno_com"],
  4: ["p4_envio_cli", "p4_aprovacao"],
  5: ["p5_liberado"],
};

function groupState(prog, num) {
  if (!prog) return "todo";
  const keys = GROUP_DONE_KEYS[num];
  const done = keys.filter(k => prog[k]).length;
  if (done === keys.length) return "done";
  if (done > 0) return "partial";
  return "todo";
}

// Usuários padrão (semente). A lista real fica no estado (localStorage) e pode
// ser gerida pela administração. Senha inicial = nome em minúsculas.
const DEFAULT_USERS = [
  { name: "Daniela",     password: "daniela",      isAdmin: true  },
  { name: "Luana",       password: "luana",        isAdmin: true  },
  { name: "Fernanda",    password: "fernanda",     isAdmin: false },
  { name: "Layza Arruda",password: "layza arruda", isAdmin: false },
];

const SAMPLE_RECORDS = [
  { responsavel:"Fernanda",    empresa:"BR02", tipo:"Time & Expenses", codCliente:"1002351", cliente:"Banco ABC Brasil S.A.",         pep:"BR02CLP00005.1.1", inicio:"01/05/2026", fim:"31/05/2026", profissional:"Bruna Paz Amorim",         valorVenda:192.5,  hrsAprovadas:160, valorTotal:30800,    valorLiquido:11774.70, competencia:"05/2026" },
  { responsavel:"Fernanda",    empresa:"BR02", tipo:"Time & Expenses", codCliente:"1002351", cliente:"Banco ABC Brasil S.A.",         pep:"BR02CLP00005.1.1", inicio:"01/05/2026", fim:"31/05/2026", profissional:"Gilliard Costa Santos",    valorVenda:145.2,  hrsAprovadas:156, valorTotal:22651.20, valorLiquido:21167.55, competencia:"05/2026" },
  { responsavel:"Fernanda",    empresa:"BR02", tipo:"Time & Expenses", codCliente:"1002351", cliente:"Banco ABC Brasil S.A.",         pep:"BR02CLP00005.1.1", inicio:"01/05/2026", fim:"31/05/2026", profissional:"Joyce Graciete da Costa",  valorVenda:225,    hrsAprovadas:160, valorTotal:36000,    valorLiquido:33642.00, competencia:"05/2026" },
  { responsavel:"Fernanda",    empresa:"BR02", tipo:"Time & Expenses", codCliente:"1002840", cliente:"Banco BS2 S.A.",                pep:"BR02CLP00100.1.1", inicio:"01/05/2026", fim:"31/05/2026", profissional:"Emerson França",           valorVenda:217,    hrsAprovadas:160, valorTotal:34720,    valorLiquido:32445.84, competencia:"05/2026" },
  { responsavel:"Fernanda",    empresa:"BR02", tipo:"Time & Expenses", codCliente:"1002342", cliente:"Banco Digio S.A.",              pep:"BR02CLP00007.1.1", inicio:"01/05/2026", fim:"31/05/2026", profissional:"Tamiris Ferreira",         valorVenda:202.14, hrsAprovadas:168, valorTotal:33959.52, valorLiquido:31735.17, competencia:"05/2026" },
  { responsavel:"Layza Arruda",empresa:"BR02", tipo:"Time & Expenses", codCliente:"1002100", cliente:"Diagnósticos da América S.A.", pep:"BR02CLP00041",      inicio:"01/05/2026", fim:"31/05/2026", profissional:"Adriano Silva Gama",       valorVenda:135,    hrsAprovadas:168, valorTotal:22680,    valorLiquido:21194.46, competencia:"05/2026" },
  { responsavel:"Layza Arruda",empresa:"BR02", tipo:"Time & Expenses", codCliente:"1002100", cliente:"Diagnósticos da América S.A.", pep:"BR02CLP00041",      inicio:"01/05/2026", fim:"31/05/2026", profissional:"Bruno Eduardo Ferreira",   valorVenda:100,    hrsAprovadas:168, valorTotal:16800,    valorLiquido:15699.60, competencia:"05/2026" },
  { responsavel:"Layza Arruda",empresa:"BR02", tipo:"Time & Expenses", codCliente:"1002100", cliente:"Diagnósticos da América S.A.", pep:"BR02CLP00041",      inicio:"01/05/2026", fim:"31/05/2026", profissional:"Caio Enrique Marcelli",    valorVenda:146,    hrsAprovadas:168, valorTotal:24528,    valorLiquido:22921.42, competencia:"05/2026" },
  { responsavel:"Layza Arruda",empresa:"BR02", tipo:"Time & Expenses", codCliente:"1002214", cliente:"Dr. Consulta Centro Médico",   pep:"BR02CLP00022.1.1", inicio:"01/05/2026", fim:"31/05/2026", profissional:"Adriano Costa Andrade",    valorVenda:142,    hrsAprovadas:168, valorTotal:23856,    valorLiquido:22293.43, competencia:"05/2026" },
  { responsavel:"Layza Arruda",empresa:"BR02", tipo:"Time & Expenses", codCliente:"1002418", cliente:"Grupo Casas Bahia S.A.",       pep:"BR02CLP00042.0.3", inicio:"01/05/2026", fim:"31/05/2026", profissional:"Abel de Meira Junior",     valorVenda:172.29, hrsAprovadas:168, valorTotal:28944.72, valorLiquido:27048.84, competencia:"05/2026" },
  { responsavel:"Layza Arruda",empresa:"BR02", tipo:"Time & Expenses", codCliente:"1002418", cliente:"Grupo Casas Bahia S.A.",       pep:"BR02CLP00042.0.3", inicio:"01/05/2026", fim:"31/05/2026", profissional:"Amanda Penido",            valorVenda:127.68, hrsAprovadas:168, valorTotal:21450.24, valorLiquido:20045.25, competencia:"05/2026" },
];

// ─── KANBAN — colunas e tarefas de exemplo ───────────────────────────────────

const TASK_COLUMNS = [
  { id:"inbox", title:"Inbox",   color:"#4b5563", accent:"#9ca3af", hint:"Crie e organize novas tarefas" },
  { id:"todo",  title:"A fazer", color:"#b45309", accent:"#f59e0b" },
  { id:"doing", title:"Fazendo", color:"#1d4ed8", accent:"#3b82f6" },
  { id:"done",  title:"Feito",   color:"#15803d", accent:"#22c55e" },
];

const SAMPLE_TASKS = [
  { title:"Extrair base de T&E de junho",                desc:"Baixar relatório no FC Team e validar as colunas antes de montar o racional.", dueDate:"2026-07-03", assignee:"Fernanda",     status:"todo"  },
  { title:"Cobrar retorno do comercial — Casas Bahia",   desc:"Acompanhar aprovação dos valores ajustados com o time comercial.",             dueDate:"2026-07-01", assignee:"Layza Arruda", status:"doing" },
  { title:"Fechar NFs dentro do corte",                  desc:"Garantir a emissão das notas aprovadas antes da data de corte.",               dueDate:"2026-06-30", assignee:"Daniela",      status:"inbox" },
  { title:"Revisar dashboard de faturamento de maio",    desc:"Validar os números do mês com a Daniela.",                                     dueDate:"2026-06-26", assignee:"Daniela",      status:"done"  },
];

// ─── HELPERS ─────────────────────────────────────────────────────────────────

const fmtShort = (n) => n == null ? "—" : "R$ " + Math.round(n).toLocaleString("pt-BR");
const nowISO   = () => new Date().toISOString();
const fmtDT    = (iso) => { if (!iso) return "—"; const d = new Date(iso); return d.toLocaleDateString("pt-BR") + " " + d.toLocaleTimeString("pt-BR", { hour:"2-digit", minute:"2-digit" }); };
const genId    = () => "r" + Date.now() + Math.random().toString(36).slice(2,7);
const uuid     = () => (typeof crypto!=="undefined" && crypto.randomUUID) ? crypto.randomUUID()
  : "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g,c=>{const r=Math.random()*16|0;return (c==="x"?r:(r&0x3|0x8)).toString(16);});
const makeProgress = () => Object.fromEntries(STEPS.map(s => [s.id, s.type==="date" ? "" : false]));
const initials = (name="") => name.trim().split(/\s+/).slice(0,2).map(p=>p[0]||"").join("").toUpperCase();
const parseJSON = (str, fallback) => { try { const v = JSON.parse(str); return v ?? fallback; } catch { return fallback; } };

// ─── CSV helpers ──────────────────────────────────────────────────────────────
const csvEscape = (v) => `"${String(v ?? "").replace(/"/g, '""')}"`;
function downloadCSV(filename, headers, rows) {
  const csv = "﻿" + [headers.map(csvEscape).join(","), ...rows.map(r => r.map(csvEscape).join(","))].join("\n");
  const a = document.createElement("a");
  a.href = URL.createObjectURL(new Blob([csv], { type: "text/csv;charset=utf-8;" }));
  a.download = filename; a.click();
}
// Exporta .xlsx com tipos corretos: números viram célula numérica e datas (JS
// Date) recebem formato dd/mm/aaaa — assim o Excel já abre formatado.
function downloadXLSX(filename, headers, rows) {
  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows], { cellDates: true });
  const range = XLSX.utils.decode_range(ws["!ref"]);
  for (let R = 1; R <= range.e.r; R++) for (let Cc = range.s.c; Cc <= range.e.c; Cc++) {
    const cell = ws[XLSX.utils.encode_cell({ r: R, c: Cc })];
    if (cell && cell.t === "d") cell.z = "dd/mm/yyyy";
  }
  ws["!cols"] = headers.map(h => ({ wch: Math.min(40, Math.max(10, String(h).length + 2)) }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Relatório");
  XLSX.writeFile(wb, filename);
}
// dd/mm/aaaa ou aaaa-mm-dd (ou ISO) → Date (para célula de data no xlsx).
function toDate(v) {
  const s = String(v || "").trim(); if (!s) return "";
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);        if (m) return new Date(+m[1], +m[2]-1, +m[3]);
  m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})/);           if (m) return new Date(+m[3], +m[2]-1, +m[1]);
  return "";
}
function fmtPepsCSV(json)      { const m = parseJSON(json, {}); if (!m || typeof m !== "object") return ""; return Object.entries(m).map(([k, v]) => `${k}: ${(Array.isArray(v) ? v : []).join(", ")}`).join(" | "); }
function fmtPropostasCSV(json) { const a = parseJSON(json, []); return Array.isArray(a) ? a.join(" ; ") : ""; }
function fmtCalendarioCSV(json){ const a = parseJSON(json, []); return Array.isArray(a) ? a.map((p, i) => `${i+1}) ${p.quando || ""} - ${p.etapa || ""}: ${p.oQueFazer || ""}`).join(" | ") : ""; }
function exportClientsCSV(clients) {
  const headers = ["Nome","Cód SAP","Grupo de empresa","Tipos de contrato","PEPs","Propostas","Período faturamento","Calendário (passo a passo)","Tem portal","Classificação portal","Link portal","Usuário portal","Senha portal","Passo a passo portal","Prazo vencimento","Forma de pagamento","Contato financeiro (nome)","Contato financeiro (e-mail)","Account manager (nome)","Account manager (e-mail)"];
  const rows = clients.map(c => [c.nome, c.codSap, c.grupoEmpresa, c.tiposContrato, fmtPepsCSV(c.tiposPeps), fmtPropostasCSV(c.propostas), c.periodoFaturamento, fmtCalendarioCSV(c.calendario), c.temPortal ? "Sim" : "Não", c.portalTipo, c.portalLink, c.portalUsuario, c.portalSenha, c.portalPassoUrl, c.prazoVencimento, c.formaPagamento, c.contatoFinanceiro, c.contatoFinanceiroEmail, c.accountManager, c.accountManagerEmail]);
  downloadCSV(`FCamara_Clientes_${clients.length}.csv`, headers, rows);
}

const PORTAL_TIPOS = ["Inclusão de notas", "Medição de serviços"];

// "faturado" = conciliado (a NF é amarrada na tela de Conciliação, não no passo a passo).
const isFaturado = (prog) => !!(prog && (prog.p5_nf || prog.p5_no_corte));
function calcStatus(prog) {
  if (!prog) return "Não iniciado";
  if (isFaturado(prog))    return "Faturado";
  if (prog.p5_liberado)    return "Liberado para faturamento";
  if (prog.p4_aprovacao)   return "Cliente aprovou";
  if (prog.p4_envio_cli)   return "Aguard. aprovação cliente";
  if (prog.p3_retorno_com) return "Retorno comercial recebido";
  if (prog.p3_envio_com)   return "Aguard. retorno comercial";
  if (prog.p2_racional)    return "Racional montado";
  if (prog.p1_extrair)     return "Dados extraídos";
  return "Não iniciado";
}

function calcStatusColor(prog) {
  if (!prog) return "gray";
  if (isFaturado(prog))    return "green";
  if (prog.p5_liberado)    return "teal";
  if (prog.p4_aprovacao)   return "blue";
  if (prog.p4_envio_cli || prog.p3_retorno_com || prog.p3_envio_com) return "yellow";
  if (prog.p1_extrair || prog.p2_racional) return "orange";
  return "gray";
}

const STATUS_ORDER = ["Não iniciado","Dados extraídos","Racional montado","Aguard. retorno comercial","Retorno comercial recebido","Aguard. aprovação cliente","Cliente aprovou","Liberado para faturamento","Faturado parcial","Faturado"];
// Status considerando o faturamento parcial (allocated = quanto já foi faturado).
function recStatus(r, allocated, totalOverride) {
  const total = totalOverride==null ? (r.valorTotal||0) : totalOverride, a = allocated||0;
  if (Math.abs(total) > 0.01 && Math.abs(total - a) < 0.01) return "Faturado";
  if (Math.abs(a) > 0.001)                                  return "Faturado parcial";
  return calcStatus(r.progress);
}
function recStatusColor(r, allocated, totalOverride) {
  const total = totalOverride==null ? (r.valorTotal||0) : totalOverride, a = allocated||0;
  if (Math.abs(total) > 0.01 && Math.abs(total - a) < 0.01) return "green";
  if (Math.abs(a) > 0.001)                                  return "orange";
  return calcStatusColor(r.progress);
}

// Datas obrigatórias do funil (só as etapas que existem no funil do tipo).
function reqDateSteps(tipo) {
  const nums = funnelGroups(tipo).map(g => g.num);
  const reqs = [];
  if (nums.includes(3)) reqs.push("p3_data_retorno");
  if (nums.includes(4)) reqs.push("p4_data_aprov");
  return reqs;
}
// Registro faturado mas sem as datas obrigatórias preenchidas (caminho reverso).
function faltaDatas(r) {
  const p = r.progress || {};
  if (!p.p5_nf) return false;
  return reqDateSteps(r.tipo).some(k => !String(p[k] || "").trim());
}
// Conclui o funil do registro ao conciliar: marca os checks do funil + NF +
// faturado no corte e carimba a data de emissão da nota. NÃO inventa as datas
// obrigatórias (retorno/aprovação) — elas continuam pendentes para o time.
function faturarProgress(record, note) {
  const p = { ...(record.progress || {}) };
  const nums = funnelGroups(record.tipo).map(g => g.num);
  if (nums.includes(1)) p.p1_extrair = true;
  if (nums.includes(2)) p.p2_racional = true;
  if (nums.includes(3)) { p.p3_envio_com = true; p.p3_retorno_com = true; }
  if (nums.includes(4)) { p.p4_envio_cli = true; p.p4_aprovacao = true; }
  p.p5_liberado = true;
  p.p5_nf = true; // marcador interno de "faturado" (a NF vem da conciliação)
  p.p5_data_nf = String(note?.emitidaEm || "").slice(0, 10) || p.p5_data_nf || "";
  p.p5_no_corte = true;
  return p;
}

// ─── STORAGE ─────────────────────────────────────────────────────────────────

// Os DADOS (registros, tarefas, histórico) agora ficam no Supabase.
// O localStorage guarda apenas preferências de UI + a lista local de usuários
// (a gestão de acessos no banco é um marco posterior).
function initState() {
  return {
    competenciaAtual: "05/2026",
    users: DEFAULT_USERS.map(u => ({ ...u })),
  };
}

function loadState() {
  try {
    const r = localStorage.getItem(LS_KEY);
    if (r) {
      const p = JSON.parse(r);
      if (!p.competenciaAtual) p.competenciaAtual = "05/2026";
      // Migração: garante a lista de usuários e a presença dos admins-semente.
      if (!Array.isArray(p.users) || p.users.length === 0) p.users = DEFAULT_USERS.map(u => ({ ...u }));
      DEFAULT_USERS.filter(d => d.isAdmin).forEach(d => {
        if (!p.users.some(u => u.name.toLowerCase() === d.name.toLowerCase())) p.users.push({ ...d });
      });
      return p;
    }
  } catch {}
  return initState();
}

function saveState(s) { try { localStorage.setItem(LS_KEY, JSON.stringify({ competenciaAtual: s.competenciaAtual, users: s.users })); } catch {} }

// ─── DESIGN TOKENS ───────────────────────────────────────────────────────────
// Fonte única de cores, raios, sombras e tipografia. Mantém o azul da marca.

// Tipografia — identidade Fcamara: Poppins (títulos, arredondada) + Inter (corpo, legível).
const FONT = "'Inter', system-ui, -apple-system, 'Segoe UI', Roboto, sans-serif";
const FONT_DISPLAY = "'Poppins', 'Inter', system-ui, sans-serif";
// Paleta Fcamara: laranja #SangueLaranja sobre muito branco + charcoal (minimalista).
// Cores via CSS variables (definidas em GLOBAL_CSS p/ claro e escuro). Como os
// estilos inline usam var(--x), tudo troca de tema na hora que o data-theme muda,
// sem re-render — inclusive constantes derivadas (inp, Ty) que espelham estes valores.
const T = {
  font: FONT, fontDisplay: FONT_DISPLAY,
  brand:      "var(--brand)",
  brandDark:  "var(--brand-dark)",
  brandBg:    "var(--brand-bg)",
  brandSoft:  "var(--brand-soft)",
  accent:     "var(--brand)",
  accentDark: "var(--brand-dark)",
  accentBg:   "var(--brand-bg)",
  dark:       "var(--dark)",
  ink:     "var(--ink)",
  inkSoft: "var(--ink-soft)",
  muted:   "var(--muted)",
  faint:   "var(--faint)",
  surface: "var(--surface)",
  canvas:  "var(--canvas)",
  line:    "var(--line)",
  lineSoft:"var(--line-soft)",
  ok:      "var(--ok)", okBg:"var(--ok-bg)", okLine:"var(--ok-line)",
  warn:    "var(--warn)", warnBg:"var(--warn-bg)", warnLine:"var(--warn-line)",
  danger:  "var(--danger)", dangerBg:"var(--danger-bg)", dangerLine:"var(--danger-line)",
  rSm:8, rMd:10, rLg:14, rXl:18, rPill:999,
  shSm:"var(--sh-sm)",
  shCard:"var(--sh-card)",
  shMd:"var(--sh-md)",
  shLg:"var(--sh-lg)",
};

// Escala tipográfica
const Ty = {
  h1:    { fontSize:20, fontWeight:700, fontFamily:FONT_DISPLAY, color:T.ink, margin:0, letterSpacing:"-.01em" },
  h2:    { fontSize:15, fontWeight:600, fontFamily:FONT_DISPLAY, color:T.ink, margin:0 },
  body:  { fontSize:13, color:T.inkSoft },
  small: { fontSize:12, color:T.muted },
  label: { fontSize:12, fontWeight:600, color:T.inkSoft, display:"block", marginBottom:5 },
};

const C = {
  green:  { bg:"#dcfce7", text:"#14532d", border:"#86efac", solid:"#16a34a" },
  teal:   { bg:"#ccfbf1", text:"#134e4a", border:"#5eead4", solid:"#0d9488" },
  blue:   { bg:"#dbeafe", text:"#1e3a8a", border:"#93c5fd", solid:"#2563eb" },
  yellow: { bg:"#fef9c3", text:"#713f12", border:"#fde047", solid:"#ca8a04" },
  orange: { bg:"#ffedd5", text:"#7c2d12", border:"#fdba74", solid:"#ea580c" },
  gray:   { bg:"#f3f4f6", text:"#374151", border:"#d1d5db", solid:"#6b7280" },
  red:    { bg:"#fee2e2", text:"#7f1d1d", border:"#fca5a5", solid:"#dc2626" },
  purple: { bg:"#f3e8ff", text:"#581c87", border:"#d8b4fe", solid:"#9333ea" },
};

const inp = { padding:"8px 11px", borderRadius:T.rMd, border:`1px solid ${T.line}`, fontSize:13, fontFamily:"inherit", background:"var(--surface)", color:T.ink, width:"100%", boxSizing:"border-box", outline:"none" };

// ─── Período quebrado + classificação do não-faturado ────────────────────────
// Dia de corte do cliente → "competência de faturamento" (visão cliente). Peça
// cujo dia de início é >= dia de corte fatura no mês seguinte. Sem corte = mês
// de serviço (01–31). Helpers no módulo p/ reuso entre as telas.
const _normCliNome = s => (s||"").toString().trim().toLowerCase();
// Aplica o mapa DE→PARA (chave = nome normalizado) ao carregar os registros, para
// que o MESMO cliente com nomes diferentes vire um só em TODAS as telas. Guarda
// clienteOrig como rastro. Não muta o dado no banco (é só leitura).
function applyClientAliases(list, aliasMap) {
  if (!aliasMap || !Object.keys(aliasMap).length) return list;
  return list.map(r => { const para = aliasMap[_normCliNome(r.cliente)]; return (para && para!==r.cliente) ? { ...r, cliente: para, clienteOrig: r.cliente } : r; });
}
function diaCorteOf(record, clients=[]) {
  const rc = _normCliNome(record.cliente);
  const c = (clients||[]).find(cl => { const cn=_normCliNome(cl.nome); return cn && (rc===cn || (cn.length>4 && rc.includes(cn))); });
  const d = parseInt(c?.diaCorte, 10);
  return (d>=1 && d<=28) ? d : 0;
}
function compFatOf(record, clients=[]) {
  const D = diaCorteOf(record, clients);
  if (!D) return record.competencia || "";
  const p = String(record.inicio||"").split("/");   // dd/mm/yyyy
  let day=+p[0], m=+p[1], y=+p[2];
  if (!m||!y){ const cp=String(record.competencia||"").split("/"); m=+cp[0]; y=+cp[1]; day=1; }
  if (!m||!y) return record.competencia || "";
  if (day>=D){ m+=1; if(m>12){m=1;y+=1;} }
  return `${String(m).padStart(2,"0")}/${y}`;
}
const _compIdx = mmYYYY => { const [m,y]=String(mmYYYY||"").split("/").map(Number); return (y&&m)?(y*12+(m-1)):null; };
// Categoria automática do não-faturado: dentro do ciclo enquanto o mês atual for
// <= compFat+1 (1 mês de folga p/ fechar/conciliar); de compFat+2 em diante vira
// represado. hoje é injetável só para teste.
function categoriaOf(record, clients=[], hoje) {
  const cf = compFatOf(record, clients);
  const ci = _compIdx(cf);
  if (ci==null) return { cat:"ciclo", compFat:cf };
  const d = hoje || new Date();
  const now = d.getFullYear()*12 + d.getMonth();
  return { cat: now >= ci+2 ? "represado" : "ciclo", compFat:cf };
}
// Motivos padrão do não-faturado (dropdown). "Outro" → detalhar no texto livre.
const CLASS_MOTIVOS = [
  "Aguardando aprovação do cliente",
  "Aguardando nota/PO do cliente",
  "Pendência de medição/timesheet",
  "Divergência de valor",
  "Contrato/aditivo em andamento",
  "Bloqueio interno FCamara",
  "Faturamento postergado (acordo comercial)",
  "Glosa / não faturável",
  "Erro de lançamento",
  "Outro",
];

// ─── GLOBAL STYLES (foco, hover, animações, scrollbar) ───────────────────────

const GLOBAL_CSS = `
  :root{
    --brand:#f1572c; --brand-dark:#d8431b; --brand-bg:#fef1ec; --brand-soft:#f9d3c6;
    --dark:#26221f;
    --ink:#1c1917; --ink-soft:#44403c; --muted:#78716c; --faint:#a8a29e;
    --surface:#ffffff; --canvas:#fafaf9; --line:#e7e5e4; --line-soft:#f5f5f4;
    --ok:#15803d; --ok-bg:#f0fdf4; --ok-line:#86efac;
    --warn:#b45309; --warn-bg:#fffbeb; --warn-line:#fcd34d;
    --danger:#dc2626; --danger-bg:#fef2f2; --danger-line:#fca5a5;
    --sh-sm:0 1px 2px rgba(28,25,23,.05);
    --sh-card:0 1px 2px rgba(28,25,23,.04), 0 8px 22px rgba(28,25,23,.05);
    --sh-md:0 6px 20px rgba(28,25,23,.07);
    --sh-lg:0 20px 50px rgba(28,25,23,.16);
    --scrollbar:#cbd5e1; --row-hover:#fef1ec80;
  }
  :root[data-theme="dark"]{
    --brand:#ff6a41; --brand-dark:#ff8a68; --brand-bg:#3a2018; --brand-soft:#5c3326;
    --dark:#0c0a09;
    --ink:#f4efe9; --ink-soft:#cdc5be; --muted:#968c84; --faint:#6f665f;
    --surface:#221e1a; --canvas:#161311; --line:#38322d; --line-soft:#2a2521;
    --ok:#4ade80; --ok-bg:#14271b; --ok-line:#2f5b46;
    --warn:#fbbf24; --warn-bg:#2c2410; --warn-line:#5c4526;
    --danger:#f87171; --danger-bg:#2e1a1a; --danger-line:#5e3030;
    --sh-sm:0 1px 2px rgba(0,0,0,.4);
    --sh-card:0 1px 2px rgba(0,0,0,.3), 0 8px 22px rgba(0,0,0,.45);
    --sh-md:0 6px 20px rgba(0,0,0,.5);
    --sh-lg:0 20px 50px rgba(0,0,0,.6);
    --scrollbar:#4b4540; --row-hover:#ff6a4118;
  }
  *{box-sizing:border-box}
  body{margin:0;font-family:${FONT};-webkit-font-smoothing:antialiased;-moz-osx-font-smoothing:grayscale;color:${T.ink};background:${T.canvas};transition:background .25s ease,color .25s ease}
  button{font-family:inherit}
  :focus-visible{outline:2px solid ${T.brand};outline-offset:2px;border-radius:6px}
  input:focus,select:focus,textarea:focus{border-color:${T.brand};box-shadow:0 0 0 3px rgba(241,87,44,.15)}
  .fc-btn{transition:filter .12s,box-shadow .12s,background .12s,transform .06s}
  .fc-btn:hover:not(:disabled){filter:brightness(.97)}
  .fc-btn:active:not(:disabled){transform:translateY(1px)}
  .fc-row:hover{background:var(--row-hover)}
  .fc-card-int{transition:box-shadow .15s,border-color .15s,transform .12s}
  .fc-card-int:hover{box-shadow:${T.shMd};border-color:${T.brandSoft};transform:translateY(-1px)}
  .fc-scroll::-webkit-scrollbar{height:9px;width:9px}
  .fc-scroll::-webkit-scrollbar-thumb{background:var(--scrollbar);border-radius:9px}
  .fc-scroll::-webkit-scrollbar-track{background:transparent}
  @keyframes fcToastIn{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:none}}
  @keyframes fcOverlay{from{opacity:0}to{opacity:1}}
  @keyframes fcModalIn{from{opacity:0;transform:translateY(10px) scale(.98)}to{opacity:1;transform:none}}
  .fc-toast{animation:fcToastIn .18s ease}
`;

function GlobalStyles() {
  return <style dangerouslySetInnerHTML={{ __html: GLOBAL_CSS }} />;
}

// ─── HOOKS ───────────────────────────────────────────────────────────────────

function useIsMobile(maxWidth = 820) {
  const [m, setM] = useState(() => typeof window !== "undefined" && window.innerWidth <= maxWidth);
  useEffect(() => {
    const on = () => setM(window.innerWidth <= maxWidth);
    window.addEventListener("resize", on);
    return () => window.removeEventListener("resize", on);
  }, [maxWidth]);
  return m;
}

// ─── TOASTS ──────────────────────────────────────────────────────────────────

const ToastCtx = createContext(() => {});
const useToast = () => useContext(ToastCtx);

function ToastProvider({ children }) {
  const [toasts, setToasts] = useState([]);
  const push = useCallback((text, type = "ok") => {
    const id = genId();
    setToasts(t => [...t, { id, text, type }]);
    setTimeout(() => setToasts(t => t.filter(x => x.id !== id)), 3800);
  }, []);
  const tc = { ok:{bar:T.ok,ic:"✓"}, error:{bar:T.danger,ic:"✕"}, info:{bar:T.brand,ic:""} };
  return (
    <ToastCtx.Provider value={push}>
      {children}
      <div style={{ position:"fixed", bottom:18, right:18, zIndex:500, display:"flex", flexDirection:"column", gap:8, maxWidth:"calc(100vw - 36px)" }} role="status" aria-live="polite">
        {toasts.map(t => {
          const s = tc[t.type] || tc.info;
          return (
            <div key={t.id} className="fc-toast" style={{ display:"flex", alignItems:"center", gap:10, background:"var(--surface)", borderLeft:`4px solid ${s.bar}`, boxShadow:T.shMd, borderRadius:T.rMd, padding:"11px 16px", minWidth:240, fontSize:13, color:T.ink }}>
              <span style={{ width:20, height:20, borderRadius:"50%", background:s.bar, color:"#fff", display:"flex", alignItems:"center", justifyContent:"center", fontSize:12, flexShrink:0 }}>{s.ic}</span>
              <span style={{ fontWeight:500 }}>{t.text}</span>
            </div>
          );
        })}
      </div>
    </ToastCtx.Provider>
  );
}

// ─── ATOMS ───────────────────────────────────────────────────────────────────

function Badge({ label, color="gray", small, dot }) {
  const c = C[color]||C.gray;
  return (
    <span style={{ display:"inline-flex", alignItems:"center", gap:5, fontSize:small?10:11, padding:small?"2px 8px":"3px 10px", borderRadius:T.rPill, background:c.bg, color:c.text, border:`1px solid ${c.border}`, fontWeight:600, whiteSpace:"nowrap", lineHeight:1.4 }}>
      {dot && <span style={{ width:6, height:6, borderRadius:"50%", background:c.solid, flexShrink:0 }} />}
      {label}
    </span>
  );
}

// Ícones minimalistas de traço (sem emojis) — Feather-style, herdam a cor do texto.
function Icon({ name, size=16, style }) {
  const p = { width:size, height:size, viewBox:"0 0 24 24", fill:"none", stroke:"currentColor", strokeWidth:1.8, strokeLinecap:"round", strokeLinejoin:"round", style:{display:"inline-block",flexShrink:0,verticalAlign:"-.15em",...style}, "aria-hidden":true };
  switch (name) {
    case "home":     return <svg {...p}><path d="M3 10.5 12 3l9 7.5"/><path d="M5 9.5V21h5v-6h4v6h5V9.5"/></svg>;
    case "list":     return <svg {...p}><rect x="4" y="4" width="16" height="16" rx="2"/><path d="M8 9h8M8 13h8M8 17h5"/></svg>;
    case "chart":    return <svg {...p}><path d="M4 20V4"/><path d="M4 20h16"/><rect x="7" y="12" width="3" height="6"/><rect x="12" y="8" width="3" height="10"/><rect x="17" y="14" width="3" height="4"/></svg>;
    case "receipt":  return <svg {...p}><path d="M6 2h12v20l-2.5-1.7L13 22l-2.5-1.7L8 22l-2-1.7V2z"/><path d="M9 7h6M9 11h6"/></svg>;
    case "file":     return <svg {...p}><path d="M14 3H6v18h12V8z"/><path d="M14 3v5h5"/><path d="M9 13h6M9 17h4"/></svg>;
    case "building": return <svg {...p}><rect x="4" y="3" width="16" height="18" rx="1"/><path d="M9 7h2M13 7h2M9 11h2M13 11h2M9 15h2M13 15h2"/></svg>;
    case "task":     return <svg {...p}><circle cx="12" cy="12" r="9"/><path d="m8.5 12 2.5 2.5 4.5-5"/></svg>;
    case "import":   return <svg {...p}><path d="M12 3v12m0 0 4-4m-4 4-4-4"/><path d="M4 17v2a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-2"/></svg>;
    case "upload":   return <svg {...p}><path d="M12 15V3m0 0 4 4m-4-4L8 7"/><path d="M4 15v4a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-4"/></svg>;
    case "download": return <svg {...p}><path d="M12 3v12m0 0 4-4m-4 4-4-4"/><path d="M4 17v2a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-2"/></svg>;
    case "lock":     return <svg {...p}><rect x="5" y="11" width="14" height="10" rx="2"/><path d="M8 11V8a4 4 0 0 1 8 0v3"/></svg>;
    case "wallet":   return <svg {...p}><path d="M3 7a2 2 0 0 1 2-2h11v3"/><path d="M3 7v10a2 2 0 0 0 2 2h13a2 2 0 0 0 2-2V9a2 2 0 0 0-2-2H4"/><circle cx="17" cy="13" r="1.3"/></svg>;
    case "gift":     return <svg {...p}><rect x="3" y="8" width="18" height="4" rx="1"/><path d="M5 12v9h14v-9M12 8v13"/><path d="M12 8C11 4 7 4.5 7.5 6.5S12 8 12 8zM12 8c1-4 5-3.5 4.5-1.5S12 8 12 8z"/></svg>;
    case "pin":      return <svg {...p}><path d="M12 21s-6-5.5-6-10a6 6 0 1 1 12 0c0 4.5-6 10-6 10z"/><circle cx="12" cy="11" r="2.2"/></svg>;
    case "pencil":   return <svg {...p}><path d="M4 20h4L18.5 9.5a2.12 2.12 0 0 0-3-3L5 17v3z"/><path d="m13.5 6.5 3 3"/></svg>;
    case "trash":    return <svg {...p}><path d="M4 7h16"/><path d="M9 7V5h6v2"/><path d="m6 7 1 13h10l1-13"/></svg>;
    case "undo":     return <svg {...p}><path d="M9 14 4 9l5-5"/><path d="M4 9h11a5 5 0 0 1 5 5v5"/></svg>;
    case "search":   return <svg {...p}><circle cx="11" cy="11" r="7"/><path d="m21 21-4-4"/></svg>;
    case "star":     return <svg {...p}><path d="M12 3l2.2 5.8L20 11l-5.8 2.2L12 19l-2.2-5.8L4 11l5.8-2.2z"/></svg>;
    case "plus":     return <svg {...p}><path d="M12 5v14M5 12h14"/></svg>;
    case "x":        return <svg {...p}><path d="M6 6l12 12M18 6 6 18"/></svg>;
    case "check":    return <svg {...p}><path d="M5 12.5 10 17l9-10"/></svg>;
    case "link":     return <svg {...p}><path d="M9 15l6-6"/><path d="M11 6l1-1a4 4 0 0 1 6 6l-1 1"/><path d="M13 18l-1 1a4 4 0 0 1-6-6l1-1"/></svg>;
    case "alert":    return <svg {...p}><path d="M12 3 2 20h20z"/><path d="M12 10v4M12 17h.01"/></svg>;
    case "chevronDown":  return <svg {...p}><path d="m6 9 6 6 6-6"/></svg>;
    case "chevronUp":    return <svg {...p}><path d="m6 15 6-6 6 6"/></svg>;
    case "chevronLeft":  return <svg {...p}><path d="m15 6-6 6 6 6"/></svg>;
    case "chevronRight": return <svg {...p}><path d="m9 6 6 6-6 6"/></svg>;
    case "info":     return <svg {...p}><circle cx="12" cy="12" r="9"/><path d="M12 11v5M12 8h.01"/></svg>;
    case "refresh":  return <svg {...p}><path d="M20 11a8 8 0 1 0-2 5"/><path d="M20 5v6h-6"/></svg>;
    case "folder":   return <svg {...p}><path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v9a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z"/></svg>;
    case "calendar": return <svg {...p}><rect x="3" y="5" width="18" height="16" rx="2"/><path d="M3 9h18M8 3v4M16 3v4"/></svg>;
    case "menu":     return <svg {...p}><path d="M4 7h16M4 12h16M4 17h16"/></svg>;
    case "sun":      return <svg {...p}><circle cx="12" cy="12" r="4"/><path d="M12 2v2M12 20v2M4.9 4.9l1.4 1.4M17.7 17.7l1.4 1.4M2 12h2M20 12h2M4.9 19.1l1.4-1.4M17.7 6.3l1.4-1.4"/></svg>;
    case "moon":     return <svg {...p}><path d="M21 12.8A9 9 0 1 1 11.2 3a7 7 0 0 0 9.8 9.8z"/></svg>;
    default:         return null;
  }
}

function Btn({ children, onClick, primary, danger, ghost, small, disabled, title, icon, style:s={} }) {
  const base = { padding:small?"6px 12px":"9px 18px", borderRadius:T.rMd, fontSize:small?12:13, fontWeight:600, cursor:disabled?"not-allowed":"pointer", display:"inline-flex", alignItems:"center", justifyContent:"center", gap:6, border:"none", opacity:disabled?.5:1, ...s };
  const v = primary ? { background:T.brand, color:"#fff" }
          : danger  ? { background:T.danger, color:"#fff" }
          : ghost   ? { background:"transparent", color:T.inkSoft }
          :           { background:"var(--surface)", color:T.inkSoft, border:`1px solid ${T.line}` };
  return <button className="fc-btn" title={title} aria-label={title} onClick={disabled?undefined:onClick} disabled={disabled} style={{...base,...v}}>{icon && <Icon name={icon} size={small?13:15}/>}{children}</button>;
}

function Avatar({ name, size=30, admin }) {
  return (
    <span style={{ width:size, height:size, borderRadius:"50%", background:admin?T.brandDark:T.brand, color:"#fff", display:"inline-flex", alignItems:"center", justifyContent:"center", fontSize:size*0.38, fontWeight:700, flexShrink:0 }} aria-hidden="true">
      {initials(name)}
    </span>
  );
}

// Marca Fcamara — wordmark "FCamara": "FC" laranja + "amara" (escuro no claro,
// branco no escuro). Mesma leitura do logo do site.
function FcamaraLogo({ size = 22, onDark }) {
  return (
    <span aria-label="FCamara" style={{ fontFamily:FONT_DISPLAY, fontWeight:700, fontSize:size, letterSpacing:"-.03em", lineHeight:1, whiteSpace:"nowrap", display:"inline-flex", alignItems:"baseline" }}>
      <span style={{ color:T.brand }}>FC</span>
      <span style={{ color:onDark ? "var(--surface)" : T.ink }}>amara</span>
    </span>
  );
}

function Card({ children, style:s={}, interactive, ...rest }) {
  return <div className={interactive?"fc-card-int":undefined} style={{ background:"var(--surface)", border:`1px solid ${T.line}`, borderRadius:T.rXl, boxShadow:T.shCard, ...s }} {...rest}>{children}</div>;
}

// Chip laranja com ícone — usado no cabeçalho de página (mesmo capricho do Início).
function HeadChip({ icon }) {
  return <div style={{width:40,height:40,borderRadius:12,background:T.brandBg,color:T.brand,display:"grid",placeItems:"center",flexShrink:0}}><Icon name={icon} size={20}/></div>;
}

// Cabeçalho de página padrão — ícone (chip laranja) + título (display) + subtítulo.
// Dá o mesmo capricho do Início nas demais telas.
function PageHead({ icon, title, sub, right }) {
  return (
    <div style={{display:"flex",alignItems:"center",gap:13,marginBottom:18,flexWrap:"wrap"}}>
      {icon && <HeadChip icon={icon}/>}
      <div style={{flex:1,minWidth:180}}>
        <h1 style={{...Ty.h1,fontSize:22}}>{title}</h1>
        {sub && <div style={{...Ty.small,marginTop:2}}>{sub}</div>}
      </div>
      {right}
    </div>
  );
}

function Field({ label, hint, children }) {
  return (
    <div>
      <label style={Ty.label}>{label}{hint && <span style={{ color:T.danger, fontWeight:500 }}> {hint}</span>}</label>
      {children}
    </div>
  );
}

function SectionTitle({ children, count }) {
  return (
    <div style={{ display:"flex", alignItems:"center", gap:8, marginBottom:12 }}>
      <h2 style={Ty.h2}>{children}</h2>
      {count != null && <span style={{ fontSize:11, fontWeight:700, color:T.muted, background:T.lineSoft, borderRadius:T.rPill, padding:"1px 9px" }}>{count}</span>}
    </div>
  );
}

// Stepper visual do funil P1→P5
function PipelineStepper({ states, groups=STEP_GROUPS, size="md", showLabels }) {
  // states: array alinhado a `groups`, com 'done' | 'partial' | 'todo'
  const dot = size==="sm" ? 14 : 18;
  const colorFor  = (st) => st==="done" ? T.ok : st==="partial" ? C.blue.solid : "var(--surface)";
  const borderFor = (st) => st==="done" ? T.ok : st==="partial" ? C.blue.solid : "#cbd2dc";
  return (
    <div style={{ display:"flex", alignItems:"flex-start" }} role="img" aria-label={"Funil: " + groups.map((g,i)=>`${g.short} ${states[i]}`).join(", ")}>
      {groups.map((g, i) => {
        const st = states[i];
        return (
          <div key={g.num} style={{ display:"flex", flexDirection:"column", alignItems:"center", flex: showLabels ? 1 : "0 0 auto" }}>
            <div style={{ display:"flex", alignItems:"center", width:"100%" }}>
              {i>0 && <div style={{ flex:1, height:2, background: states[i-1]==="done" ? T.ok : "#e2e6ec", minWidth: showLabels?0:16 }} />}
              <div title={g.title} style={{ width:dot, height:dot, borderRadius:"50%", flexShrink:0, background:colorFor(st), border:`2px solid ${borderFor(st)}`, display:"flex", alignItems:"center", justifyContent:"center", fontSize:dot*0.5, color:st==="todo"?T.faint:"#fff", fontWeight:700 }}>
                {st==="done" ? "✓" : g.num}
              </div>
              {i<groups.length-1 && <div style={{ flex:1, height:2, background: st==="done" ? T.ok : "#e2e6ec", minWidth: showLabels?0:16 }} />}
            </div>
            {showLabels && <span style={{ fontSize:10, color: st==="todo"?T.faint:T.inkSoft, marginTop:4, fontWeight: st==="done"?700:500, textAlign:"center" }}>{g.short}</span>}
          </div>
        );
      })}
    </div>
  );
}

function recordStates(prog, tipo) { return funnelGroups(tipo).map(g => groupState(prog, g.num)); }
function aggregateStates(records, tipo) {
  return funnelGroups(tipo).map(g => {
    const sts = records.map(r => groupState(r.progress, g.num));
    if (sts.every(s => s==="done")) return "done";
    if (sts.some(s => s!=="todo")) return "partial";
    return "todo";
  });
}

// ─── MODAL ───────────────────────────────────────────────────────────────────

function Modal({ title, subtitle, onClose, children, footer, wide, extraWide }) {
  const w = extraWide ? 960 : wide ? 780 : 520;
  const ref = useRef();
  const closeRef = useRef(onClose);
  closeRef.current = onClose;
  // Foca o primeiro campo só uma vez (na montagem). Depender de onClose fazia o
  // efeito re-rodar a cada tecla quando o estado do modal mora no componente pai,
  // roubando o foco e "travando" a digitação.
  useEffect(() => {
    const onKey = (e) => { if (e.key === "Escape") closeRef.current?.(); };
    document.addEventListener("keydown", onKey);
    const t = setTimeout(() => { const el = ref.current?.querySelector("input,select,textarea,button"); el?.focus(); }, 30);
    return () => { document.removeEventListener("keydown", onKey); clearTimeout(t); };
  }, []);
  return (
    <div style={{ position:"fixed", inset:0, background:"rgba(15,23,42,.55)", display:"flex", alignItems:"center", justifyContent:"center", zIndex:300, padding:16, animation:"fcOverlay .15s ease" }} onClick={onClose}>
      <div ref={ref} role="dialog" aria-modal="true" aria-label={title} className="fc-scroll" style={{ background:"var(--surface)", borderRadius:T.rXl+2, padding:"22px 26px", width:w, maxWidth:"100%", maxHeight:"92vh", overflowY:"auto", boxShadow:T.shLg, animation:"fcModalIn .18s ease" }} onClick={e=>e.stopPropagation()}>
        <div style={{ display:"flex", alignItems:"flex-start", marginBottom:18, gap:12 }}>
          <div style={{ flex:1 }}>
            <h2 style={{ fontSize:17, fontWeight:800, color:T.ink, margin:0 }}>{title}</h2>
            {subtitle && <div style={{ ...Ty.small, marginTop:3 }}>{subtitle}</div>}
          </div>
          <button onClick={onClose} aria-label="Fechar" style={{ background:T.lineSoft, border:"none", width:30, height:30, borderRadius:8, fontSize:18, cursor:"pointer", color:T.muted, lineHeight:1, flexShrink:0 }}>×</button>
        </div>
        {children}
        {footer && <div style={{ display:"flex", gap:8, justifyContent:"flex-end", marginTop:18 }}>{footer}</div>}
      </div>
    </div>
  );
}

function ConfirmDialog({ title, message, confirmLabel="Confirmar", danger, onConfirm, onClose }) {
  return (
    <Modal title={title} onClose={onClose}>
      <p style={{ ...Ty.body, lineHeight:1.5, marginTop:0 }}>{message}</p>
      <div style={{ display:"flex", gap:8, justifyContent:"flex-end", marginTop:18 }}>
        <Btn onClick={onClose}>Cancelar</Btn>
        <Btn primary={!danger} danger={danger} onClick={()=>{ onConfirm(); onClose(); }}>{confirmLabel}</Btn>
      </div>
    </Modal>
  );
}

// ─── IMPORT (admin) ──────────────────────────────────────────────────────────

const TE_COL_MAP = {
  responsavel:  ["RESPONSÁVEL","RESPONSAVEL"],
  codCliente:   ["COD CLIENTE"],
  cliente:      ["NOME CLIENTE"],
  pep:          ["PEP"],
  inicio:       ["INICIO","INÍCIO"],
  fim:          ["FIM"],
  profissional: ["PROFISSIONAL"],
  valorVenda:   ["VALOR DE VENDA"],
  hrsAprovadas: ["HRS APROVADAS"],
  valorTotal:   ["VALOR TOTAL"],
  valorLiquido: ["Valor Liquido :)","VALOR LIQUIDO","Valor Liquido"],
};

function excelDateToStr(val) {
  if (typeof val==="number") { const d=XLSX.SSF.parse_date_code(val); if(d) return `${String(d.d).padStart(2,"0")}/${String(d.m).padStart(2,"0")}/${d.y}`; }
  return typeof val==="string"?val.trim():"";
}

function findCol(headers, candidates) {
  const n = h=>(h||"").toString().trim().toUpperCase().replace(/\s+/g," ");
  for (const c of candidates) { const i=headers.findIndex(h=>n(h)===n(c)); if(i!==-1) return i; }
  return -1;
}

function parseSheetRows(rows, empresa, tipo, competencia) {
  let hi=0;
  for (let i=0;i<Math.min(6,rows.length);i++) { if(rows[i].some(c=>(c||"").toString().toUpperCase().includes("RESPONSAV"))) { hi=i; break; } }
  const headers = rows[hi].map(h=>(h||"").toString());
  const colIdx={};
  for (const [key,cands] of Object.entries(TE_COL_MAP)) { const i=findCol(headers,cands); if(i!==-1) colIdx[key]=i; }
  // "Vertical" = BU (opcional; não conta na validação de cabeçalho).
  const buIdx = findCol(headers, ["VERTICAL","Vertical","BU","VERTICAL/BU"]);
  const missing = Object.keys(TE_COL_MAP).filter(k=>colIdx[k]==null);
  if (missing.length>4) return { records:[], errors:[`Cabeçalhos não encontrados: ${missing.join(", ")}. Use a aba "Time & Expenses".`] };
  const records=[]; const skipped=[];
  for (let i=hi+1;i<rows.length;i++) {
    const row=rows[i];
    if(!row||row.every(c=>c==null||c==="")) continue;
    const get=k=>colIdx[k]!=null?(row[colIdx[k]]??""):"";
    const getNum=k=>parseFloat(String(get(k)).replace(",","."))||0;
    const getStr=k=>String(get(k)).trim();
    const cliente=getStr("cliente"), pep=getStr("pep"), responsavel=getStr("responsavel");
    if(!cliente||!pep||!responsavel){skipped.push(i+1);continue;}
    records.push({ id:genId(), responsavel, empresa, tipo, codCliente:getStr("codCliente"), cliente, pep, inicio:excelDateToStr(get("inicio")), fim:excelDateToStr(get("fim")), profissional:getStr("profissional"), valorVenda:getNum("valorVenda"), hrsAprovadas:getNum("hrsAprovadas"), valorTotal:getNum("valorTotal"), valorLiquido:getNum("valorLiquido"), bu: buIdx!==-1 ? String(row[buIdx]??"").trim() : "", competencia, progress:makeProgress(), nfNumero:"", obs:"", updatedAt:nowISO() });
  }
  const errors=[];
  if(skipped.length) errors.push(`${skipped.length} linhas ignoradas por falta de dados (linhas: ${skipped.slice(0,5).join(", ")}${skipped.length>5?"...":""}).`);
  return { records, errors };
}

// Layout Fee/WIP: tipo e empresa vêm de cada linha; sem profissional/horas.
const FEEWIP_COL_MAP = {
  responsavel:  ["RESPONSÁVEL","RESPONSAVEL"],
  tipo:         ["TIPO"],
  empresa:      ["EMPRESA"],
  codCliente:   ["COD CLIENTE"],
  cliente:      ["NOME CLIENTE"],
  pep:          ["PEP"],
  inicio:       ["INICIO","INÍCIO"],
  fim:          ["FIM"],
  valorTotal:   ["RECEITA PLANEJADA","VALOR TOTAL"],
  valorLiquido: ["Formula Líquido","FORMULA LIQUIDO","Formula Liquido","RECEITA LIQUIDA","RECEITA LÍQUIDA","Valor Liquido :)","VALOR LIQUIDO","Valor Liquido"],
  obs:          ["OBSERVAÇÃO","OBSERVACAO","Projeto","PROJETO"],
};
// Aceita competência como "05/2026" ou "052026" → "05/2026" (vazio se inválida).
function normComp(v) {
  const d = String(v||"").replace(/\D/g, "");
  if (d.length === 6) return d.slice(0,2) + "/" + d.slice(2);
  return /^\d{2}\/\d{4}$/.test(String(v||"").trim()) ? String(v).trim() : "";
}
function normTipo(t) {
  const s = String(t||"").trim().toLowerCase();
  if (s==="fee") return "Fee";
  if (s==="wip") return "WIP";
  if (s.includes("time")) return "Time & Expenses";
  if (s.includes("usage")) return "Usage Based";
  return String(t||"").trim();
}
// Parser de planilhas "por PEP" (Fee, WIP, Usage Based). Sem profissional/horas.
// tipoFixo: se informado, força o tipo (não precisa coluna TIPO); senão lê de cada linha.
function parseRevenueRows(rows, competencia, { tipoFixo=null, tipoFallback=null, empresaFixa=null } = {}) {
  let hi=0;
  for (let i=0;i<Math.min(6,rows.length);i++) { if(rows[i].some(c=>(c||"").toString().toUpperCase().includes("RESPONSAV"))) { hi=i; break; } }
  const headers = rows[hi].map(h=>(h||"").toString());
  const colIdx={};
  for (const [key,cands] of Object.entries(FEEWIP_COL_MAP)) { const i=findCol(headers,cands); if(i!==-1) colIdx[key]=i; }
  // TIPO só é obrigatório se não houver tipo fixo nem padrão.
  const needTipo = !tipoFixo && !tipoFallback;
  const need=["responsavel","cliente","pep","valorTotal"].concat(needTipo ? ["tipo"] : []);
  const missing=need.filter(k=>colIdx[k]==null);
  if (missing.length) return { records:[], errors:[`Cabeçalhos não encontrados (${missing.join(", ")}). A planilha precisa ter: RESPONSÁVEL${needTipo?", TIPO":""}, NOME CLIENTE, PEP, RECEITA PLANEJADA/VALOR TOTAL.`] };
  const records=[]; const skipped=[];
  for (let i=hi+1;i<rows.length;i++) {
    const row=rows[i];
    if(!row||row.every(c=>c==null||c==="")) continue;
    const get=k=>colIdx[k]!=null?(row[colIdx[k]]??""):"";
    const getNum=k=>parseFloat(String(get(k)).replace(",","."))||0;
    const getStr=k=>String(get(k)).trim();
    const cliente=getStr("cliente"), pep=getStr("pep"), responsavel=getStr("responsavel");
    const tipo=tipoFixo || normTipo(get("tipo")) || tipoFallback;
    if(!cliente||!pep||!responsavel||!tipo){skipped.push(i+1);continue;}
    records.push({ id:genId(), responsavel, empresa:getStr("empresa")||empresaFixa||"BR02", tipo, codCliente:getStr("codCliente"), cliente, pep, inicio:excelDateToStr(get("inicio")), fim:excelDateToStr(get("fim")), profissional:"", valorVenda:0, hrsAprovadas:0, valorTotal:getNum("valorTotal"), valorLiquido:getNum("valorLiquido"), competencia, progress:makeProgress(), nfNumero:"", obs:getStr("obs"), updatedAt:nowISO() });
  }
  const errors=[];
  if(skipped.length) errors.push(`${skipped.length} linha(s) ignorada(s) por falta de dados.`);
  return { records, errors };
}

function ImportModal({ onImport, onClose }) {
  const [layout,setLayout]=useState("te"); // te | feewip | usage
  const [competencia,setComp]=useState("");
  const [empresa,setEmpresa]=useState("BR02");
  const [tipo,setTipo]=useState("Time & Expenses");
  const [tipoFb,setTipoFb]=useState("Fee"); // padrão p/ Fee/WIP quando a planilha não tem coluna TIPO
  const [mode,setMode]=useState("merge");
  const [note,setNote]=useState("");
  const [preview,setPreview]=useState(null);
  const [fileName,setFileName]=useState("");
  const [msgs,setMsgs]=useState([]);
  const [loading,setLoading]=useState(false);
  const [dragOver,setDragOver]=useState(false);
  const fileRef=useRef();

  const reset=()=>{setPreview(null);setFileName("");setMsgs([]);};
  const comp = normComp(competencia); // aceita "05/2026" ou "052026"

  function readFile(file) {
    if(!comp){setMsgs([{type:"error",text:"Informe a competência (ex.: 05/2026 ou 052026) antes de carregar o arquivo."}]);return;}
    setLoading(true);setFileName(file.name);setPreview(null);setMsgs([]);
    const reader=new FileReader();
    reader.onload=e=>{
      try {
        const wb=XLSX.read(new Uint8Array(e.target.result),{type:"array",cellDates:false});
        const sheetName=layout==="te" ? (wb.SheetNames.find(n=>n.toLowerCase().includes("time")&&n.toLowerCase().includes("expense"))||wb.SheetNames[0]) : wb.SheetNames[0];
        const rows=XLSX.utils.sheet_to_json(wb.Sheets[sheetName],{header:1,defval:""});
        const {records,errors}=
            layout==="feewip" ? parseRevenueRows(rows, comp, { empresaFixa:empresa, tipoFallback:tipoFb })
          : layout==="usage"  ? parseRevenueRows(rows, comp, { tipoFixo:"Usage Based", empresaFixa:empresa })
          :                     parseSheetRows(rows, empresa, "Time & Expenses", comp);
        const m=[];
        errors.forEach(e=>m.push({type:"warn",text:e}));
        if(records.length===0){m.push({type:"error",text:"Nenhum registro válido. Confira as colunas da planilha."});setMsgs(m);}
        else{m.push({type:"ok",text:`${records.length} registros encontrados na aba "${sheetName}".`});setMsgs(m);setPreview(records);}
      } catch(err){setMsgs([{type:"error",text:"Erro ao ler o arquivo: "+err.message}]);}
      setLoading(false);
    };
    reader.readAsArrayBuffer(file);
  }

  const onDrop=useCallback(e=>{e.preventDefault();setDragOver(false);const f=e.dataTransfer.files[0];if(f)readFile(f);},[layout,competencia,empresa,tipoFb]);
  const mc={ok:{bg:T.okBg,text:T.ok,border:T.okLine},warn:{bg:T.warnBg,text:T.warn,border:T.warnLine},error:{bg:T.dangerBg,text:T.danger,border:T.dangerLine}};

  return (
    <Modal title="Importar dados" subtitle="Carrega a planilha .xlsm/.xlsx de receitas" onClose={onClose} wide>
      <div style={{background:T.warnBg,border:`1px solid ${T.warnLine}`,borderRadius:T.rMd,padding:"10px 14px",marginBottom:16,fontSize:12,color:T.warn,display:"flex",gap:8}}>
        <span aria-hidden="true"></span><span>Apenas administradores podem importar dados.</span>
      </div>

      {/* Layout da planilha */}
      <div style={{marginBottom:14}}>
        <label style={Ty.label}>Layout da planilha</label>
        <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
          {[{v:"te",l:"Time & Expenses",d:"Por profissional, com horas e valores."},{v:"feewip",l:"Fee / WIP",d:"Por PEP, receita planejada. Sobe Fee e WIP juntos."},{v:"usage",l:"Usage Based",d:"Por PEP, receita planejada. Tipo Usage Based."}].map(opt=>(
            <label key={opt.v} style={{flex:"1 1 180px",display:"flex",gap:8,padding:"10px 12px",borderRadius:T.rMd,border:`2px solid ${layout===opt.v?T.brand:T.line}`,cursor:"pointer",background:layout===opt.v?T.brandBg:"var(--surface)"}}>
              <input type="radio" name="layout" checked={layout===opt.v} onChange={()=>{setLayout(opt.v);reset();}} style={{marginTop:2}}/>
              <div><div style={{fontSize:13,fontWeight:700,color:layout===opt.v?T.brand:T.inkSoft}}>{opt.l}</div><div style={{fontSize:11,color:T.muted,marginTop:2}}>{opt.d}</div></div>
            </label>
          ))}
        </div>
      </div>

      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))",gap:12,marginBottom:14}}>
        <Field label="Competência *" hint="(preencha primeiro)"><input style={inp} placeholder="05/2026" value={competencia} onChange={e=>{setComp(e.target.value);reset();}}/></Field>
        <Field label="Empresa"><select style={inp} value={empresa} onChange={e=>{setEmpresa(e.target.value);reset();}}>{EMPRESAS.map(e=><option key={e.cod} value={e.cod}>{e.cod} — {e.nome}</option>)}</select></Field>
        {layout==="te" && <Field label="Tipo de projeto"><input style={{...inp,color:T.muted}} value="Time & Expenses" disabled/></Field>}
        {layout==="feewip" && <Field label="Tipo padrão" hint="(se a planilha não tiver coluna TIPO)"><select style={inp} value={tipoFb} onChange={e=>{setTipoFb(e.target.value);reset();}}><option>Fee</option><option>WIP</option></select></Field>}
        {layout==="usage" && <Field label="Tipo de projeto"><input style={{...inp,color:T.muted}} value="Usage Based" disabled/></Field>}
      </div>
      <div style={{marginBottom:14}}>
        <label style={Ty.label}>Modo de importação</label>
        <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
          {[{v:"merge",l:"Atualizar mês",d:"Reprocessa o mês: casa por PEP+profissional, atualiza valores mantendo passo a passo e conciliação, e sinaliza mudanças. Use no fechamento diário."},{v:"add",l:"Incluir novos",d:"Adiciona sem apagar nem casar. Use só na 1ª carga do mês."},{v:"replace",l:"Substituir (apagar e refazer)",d:"Remove e reimporta a competência + empresa + tipo do zero. Perde o progresso desse recorte."}].map(opt=>(
            <label key={opt.v} style={{flex:"1 1 200px",display:"flex",gap:8,padding:"10px 12px",borderRadius:T.rMd,border:`2px solid ${mode===opt.v?T.brand:T.line}`,cursor:"pointer",background:mode===opt.v?T.brandBg:"var(--surface)"}}>
              <input type="radio" name="mode" value={opt.v} checked={mode===opt.v} onChange={()=>setMode(opt.v)} style={{marginTop:2}}/>
              <div><div style={{fontSize:13,fontWeight:700,color:mode===opt.v?T.brand:T.inkSoft}}>{opt.l}</div><div style={{fontSize:11,color:T.muted,marginTop:2}}>{opt.d}</div></div>
            </label>
          ))}
        </div>
        {mode==="replace"&&<div style={{marginTop:8,fontSize:12,color:T.danger,fontWeight:600}}>Apenas os registros desta competência (mês), empresa e tipo serão substituídos. O progresso já registrado para esse recorte será perdido; os demais meses permanecem intactos.</div>}
        {mode==="merge"&&<div style={{marginTop:8,fontSize:12,color:T.inkSoft}}>As linhas são casadas por <b>empresa + PEP + profissional + competência + período (início/fim)</b>. O período na chave garante o casamento certo dos clientes de faturamento quebrado (peças 01–10 e 11–31). Valores mudados são atualizados sem apagar o passo a passo; se a NF já foi emitida, a diferença vira <b>saldo a faturar</b> e um alerta vermelho aparece na Minha visão. Linhas que sumirem do relatório são sinalizadas, não apagadas.</div>}
      </div>
      <div style={{marginBottom:14}}><Field label="Nota da importação (opcional)"><input style={inp} placeholder="Ex: Ajuste de valores de maio" value={note} onChange={e=>setNote(e.target.value)}/></Field></div>
      <input type="file" ref={fileRef} style={{display:"none"}} accept=".xlsx,.xlsm,.xls" onChange={e=>{if(e.target.files[0])readFile(e.target.files[0]);e.target.value="";}}/>
      <div onDragOver={e=>{e.preventDefault();setDragOver(true);}} onDragLeave={()=>setDragOver(false)} onDrop={onDrop} onClick={()=>fileRef.current.click()} role="button" tabIndex={0} aria-label="Carregar arquivo"
        style={{border:`2px dashed ${dragOver?T.brand:fileName?T.okLine:"#cbd2dc"}`,borderRadius:T.rLg,padding:"28px 20px",textAlign:"center",cursor:"pointer",background:dragOver?T.brandBg:fileName?T.okBg:"var(--canvas)",marginBottom:14}}>
        {loading?<div style={{color:T.muted,fontSize:13}}>Lendo arquivo...</div>:fileName?<><div style={{marginBottom:6,color:T.ok}}><Icon name="check" size={26}/></div><div style={{fontSize:13,fontWeight:700,color:T.ok}}>{fileName}</div><div style={{fontSize:11,color:T.muted,marginTop:4}}>Clique para trocar</div></>:<><div style={{marginBottom:8,color:T.muted}}><Icon name="upload" size={26}/></div><div style={{fontSize:14,fontWeight:600,color:T.inkSoft}}>Clique ou arraste o arquivo aqui</div><div style={{fontSize:12,color:T.muted,marginTop:4}}>Aceita .xlsm e .xlsx</div></>}
      </div>
      {msgs.map((m,i)=><div key={i} style={{marginBottom:6,fontSize:12,padding:"8px 12px",borderRadius:T.rMd,background:mc[m.type].bg,color:mc[m.type].text,border:`1px solid ${mc[m.type].border}`}}>{m.text}</div>)}
      {preview&&<div style={{marginBottom:14,padding:"12px 14px",borderRadius:T.rMd,background:T.okBg,border:`1px solid ${T.okLine}`}}>
        <div style={{fontSize:13,fontWeight:700,color:T.ok,marginBottom:8}}>✓ {preview.length} registros prontos</div>
        <div className="fc-scroll" style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}><thead><tr style={{background:"#dcfce7"}}>{["Responsável","Cliente","PEP","Profissional","Val. Total"].map(h=><th key={h} style={{padding:"4px 8px",textAlign:"left",color:T.ok,fontWeight:700}}>{h}</th>)}</tr></thead><tbody>{preview.slice(0,5).map(r=><tr key={r.id}><td style={{padding:"4px 8px"}}>{r.responsavel}</td><td style={{padding:"4px 8px"}}>{r.cliente}</td><td style={{padding:"4px 8px",fontFamily:"monospace"}}>{r.pep}</td><td style={{padding:"4px 8px"}}>{r.profissional}</td><td style={{padding:"4px 8px"}}>{fmtShort(r.valorTotal)}</td></tr>)}{preview.length>5&&<tr><td colSpan={5} style={{padding:"4px 8px",color:T.muted,fontStyle:"italic"}}>... e mais {preview.length-5}</td></tr>}</tbody></table></div>
      </div>}
      <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
        <Btn onClick={onClose}>Cancelar</Btn>
        <Btn primary onClick={()=>{if(!preview)return;onImport({records:preview,competencia:comp,empresa,tipo,mode,note:note||(mode==="replace"?"Substituição":mode==="merge"?"Atualização do mês":"Adição")});onClose();}} disabled={!preview}>{mode==="replace"?"Confirmar substituição":mode==="merge"?"✓ Atualizar mês":"✓ Confirmar importação"}</Btn>
      </div>
    </Modal>
  );
}

// ─── EXPORT (admin) ──────────────────────────────────────────────────────────

function ExportModal({ records, onClose, onDone }) {
  const [empresa,setE]=useState("todas");
  const [analista,setA]=useState("todos");
  const [comp,setC]=useState("todas");
  const [soNaoFat,setSN]=useState(false);
  const analistas=[...new Set(records.map(r=>r.responsavel))].sort();
  const comps=[...new Set(records.map(r=>r.competencia))].sort();
  function doExport(){
    let f=records;
    if(empresa!=="todas") f=f.filter(r=>r.empresa===empresa);
    if(analista!=="todos") f=f.filter(r=>r.responsavel===analista);
    if(comp!=="todas") f=f.filter(r=>r.competencia===comp);
    if(soNaoFat) f=f.filter(r=>!r.progress?.p5_nf);
    const headers=["Analista","Empresa","Tipo","Competência","Cliente","PEP","Profissional","Val. Venda","Hrs","Val. Total","Val. Líquido","NF Número","Status","P1","P2","P3a","P3b","P3c Data","P4a","P4b","P4c Data","P5a NF","P5b Data NF","P5c Corte","Obs","Atualizado"];
    const rows=f.map(r=>{const p=r.progress||{};return[r.responsavel,r.empresa,r.tipo,r.competencia,r.cliente,r.pep,r.profissional,r.valorVenda,r.hrsAprovadas,r.valorTotal,r.valorLiquido,r.nfNumero||"",calcStatus(p),p.p1_extrair?"S":"N",p.p2_racional?"S":"N",p.p3_envio_com?"S":"N",p.p3_retorno_com?"S":"N",p.p3_data_retorno||"",p.p4_envio_cli?"S":"N",p.p4_aprovacao?"S":"N",p.p4_data_aprov||"",p.p5_nf?"S":"N",p.p5_data_nf||"",p.p5_no_corte?"S":"N",r.obs||"",fmtDT(r.updatedAt)].map(v=>`"${String(v).replace(/"/g,'""')}"`).join(",");});
    const csv="﻿"+[headers.join(","),...rows].join("\n");
    const a=document.createElement("a");a.href=URL.createObjectURL(new Blob([csv],{type:"text/csv;charset=utf-8;"}));a.download=`FCamara_Billing_${[empresa,analista,comp].filter(v=>v!=="todas"&&v!=="todos").join("_")||"Tudo"}.csv`;a.click();
    onDone?.(f.length); onClose();
  }
  return(
    <Modal title="Exportar CSV" subtitle="Gera um arquivo .csv com os filtros abaixo" onClose={onClose}>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:16}}>
        <Field label="Empresa"><select style={inp} value={empresa} onChange={e=>setE(e.target.value)}><option value="todas">Todas</option>{EMPRESAS.map(e=><option key={e.cod} value={e.cod}>{e.cod} — {e.nome}</option>)}</select></Field>
        <Field label="Analista"><select style={inp} value={analista} onChange={e=>setA(e.target.value)}><option value="todos">Todos</option>{analistas.map(a=><option key={a}>{a}</option>)}</select></Field>
        <Field label="Competência"><select style={inp} value={comp} onChange={e=>setC(e.target.value)}><option value="todas">Todas</option>{comps.map(c=><option key={c}>{c}</option>)}</select></Field>
        <div style={{display:"flex",alignItems:"flex-end",paddingBottom:2}}><label style={{display:"flex",alignItems:"center",gap:8,fontSize:13,cursor:"pointer",color:T.inkSoft}}><input type="checkbox" checked={soNaoFat} onChange={e=>setSN(e.target.checked)} style={{width:16,height:16}}/>Somente não faturado</label></div>
      </div>
      <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}><Btn onClick={onClose}>Cancelar</Btn><Btn primary onClick={doExport}>Exportar CSV</Btn></div>
    </Modal>
  );
}

// ─── HISTORY (admin) ─────────────────────────────────────────────────────────

function HistoryModal({ history, onClose, onUndo }) {
  const [confirm, setConfirm] = useState(null);
  return(
    <Modal title="Histórico de importações" onClose={onClose} wide>
      {confirm && <ConfirmDialog title="Desfazer importação" danger confirmLabel="Desfazer importação"
        message={`Remover os ${confirm.count} registro(s) importados em ${fmtDT(confirm.date)} (${confirm.competencia} · ${confirm.tipo})? Os registros dessa carga serão apagados de uma vez. Não afeta outras importações.`}
        onConfirm={()=>onUndo(confirm)} onClose={()=>setConfirm(null)}/>}
      {history.length===0?<p style={{fontSize:13,color:T.muted}}>Nenhuma importação.</p>:
      <div className="fc-scroll" style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
        <thead><tr style={{background:T.canvas}}>{["Data/Hora","Usuário","Competência","Empresa","Tipo","Modo","Registros","Nota",""].map(h=><th key={h} style={{padding:"8px 10px",textAlign:"left",borderBottom:`1px solid ${T.line}`,fontWeight:600,color:T.muted,whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
        <tbody>{[...history].reverse().map(h=><tr key={h.id} style={{borderBottom:`1px solid ${T.lineSoft}`}}>
          <td style={{padding:"8px 10px",whiteSpace:"nowrap"}}>{fmtDT(h.date)}</td>
          <td style={{padding:"8px 10px"}}><Badge label={h.user} color="purple" small/></td>
          <td style={{padding:"8px 10px"}}>{h.competencia}</td>
          <td style={{padding:"8px 10px"}}>{h.empresa}</td>
          <td style={{padding:"8px 10px"}}>{h.tipo}</td>
          <td style={{padding:"8px 10px"}}><Badge label={h.mode==="replace"?"Substituição":"Adição"} color={h.mode==="replace"?"red":"green"} small/></td>
          <td style={{padding:"8px 10px",fontWeight:700}}>{h.count}</td>
          <td style={{padding:"8px 10px",color:T.muted}}>{h.note}</td>
          <td style={{padding:"8px 10px",textAlign:"right",whiteSpace:"nowrap"}}>
            {h.importId
              ? <Btn small danger onClick={()=>setConfirm(h)}>↩ Desfazer</Btn>
              : <span style={{fontSize:11,color:T.faint}} title="Importações anteriores a este recurso não têm como ser desfeitas automaticamente">—</span>}
          </td>
        </tr>)}</tbody>
      </table></div>}
      <div style={{display:"flex",justifyContent:"flex-end",marginTop:16}}><Btn onClick={onClose}>Fechar</Btn></div>
    </Modal>
  );
}

// ─── BULK TIMELINE MODAL ─────────────────────────────────────────────────────
// Atualiza passos de múltiplos profissionais de um cliente de uma só vez

function BulkTimelineModal({ cliente, pep, records, onSave, onClose, onOpenNF }) {
  const tipo = records[0]?.tipo;
  const grupos = funnelGroups(tipo);
  const [selected, setSelected] = useState(new Set(records.map(r=>r.id)));
  const [sharedProg, setSharedProg] = useState(() => ({ ...records[0]?.progress } || makeProgress()));
  const [obs, setObs] = useState("");
  const [ovMap, setOvMap] = useState(() => Object.fromEntries(records.map(r=>[r.id, r.ordemVenda||""])));
  const setOV = (id,v) => setOvMap(m=>({...m,[id]:v}));
  const [error, setError] = useState("");
  const [pickOpen, setPickOpen] = useState(false);   // seletor de profissionais recolhido
  const [qProf, setQProf] = useState("");

  const toggleAll = () => setSelected(s => s.size === records.length ? new Set() : new Set(records.map(r=>r.id)));
  const toggle = (id) => setSelected(s => { const n=new Set(s); n.has(id)?n.delete(id):n.add(id); return n; });
  const setVal = (id, val) => setSharedProg(p=>({...p,[id]:val}));

  // Sequência obrigatória dos passos do tipo "check" — derivada do funil do tipo.
  const FLOW = grupos.flatMap(g => STEPS.filter(s=>g.steps.includes(s.id) && s.type==="check").map(s=>s.id));
  const stepName = (id) => STEPS.find(s=>s.id===id)?.name || id;
  function setCheck(id, val) {
    setError("");
    setSharedProg(p => {
      const idx = FLOW.indexOf(id);
      if (val) {
        for (let i=0;i<idx;i++) if (!p[FLOW[i]]) { setError(`Conclua antes: “${stepName(FLOW[i])}”.`); return p; }
        return { ...p, [id]: true };
      }
      // Ao desmarcar, desmarca também os passos posteriores (mantém a sequência coerente).
      const np = { ...p };
      for (let i=idx;i<FLOW.length;i++) np[FLOW[i]] = false;
      return np;
    });
  }

  function handleSave() {
    if (selected.size === 0) { setError("Selecione ao menos um profissional."); return; }
    if (sharedProg.p3_retorno_com && !(sharedProg.p3_data_retorno||"").trim()) { setError("Marcou “Retorno do Comercial”: informe a Data Retorno."); return; }
    if (sharedProg.p4_aprovacao   && !(sharedProg.p4_data_aprov||"").trim())   { setError("Marcou “Aprovação do Cliente”: informe a Data Aprovação."); return; }
    const now = nowISO();
    // A emissão da NF (p5_nf / data / número) é gerida na tela "Notas fiscais".
    // Aqui preservamos esses campos por profissional e atualizamos só o restante do funil.
    // A Ordem de venda é por profissional (independe da seleção de passos).
    const updated = records.map(r => {
      const ovChanged = (ovMap[r.id]||"") !== (r.ordemVenda||"");
      if (!selected.has(r.id) && !ovChanged) return null;
      const base = { ...r, ordemVenda: ovMap[r.id]||"", updatedAt: now };
      return selected.has(r.id)
        ? { ...base, progress: { ...sharedProg, p5_nf: r.progress?.p5_nf || false, p5_no_corte: r.progress?.p5_no_corte || false, p5_data_nf: r.progress?.p5_data_nf || "" }, obs: obs || r.obs }
        : base;
    }).filter(Boolean);
    onSave(updated);
    onClose();
  }

  return (
    <Modal title={`Atualizar passos — ${cliente}`} subtitle={`${pep} · ${records.length} profissionais`} onClose={onClose} extraWide>
      {/* Pré-visualização do funil compartilhado */}
      <div style={{ background:T.canvas, border:`1px solid ${T.line}`, borderRadius:T.rLg, padding:"14px 16px", marginBottom:18 }}>
        <PipelineStepper states={recordStates(sharedProg, tipo)} groups={grupos} showLabels/>
      </div>

      {/* Seleção de profissionais — recolhida por padrão, abre estilo filtro */}
      <div style={{marginBottom:20}}>
        <div style={{fontSize:13,fontWeight:700,color:T.ink,marginBottom:8}}>Aplicar a</div>
        <button type="button" onClick={()=>setPickOpen(o=>!o)} style={{width:"100%",display:"flex",alignItems:"center",gap:8,background:"var(--surface)",border:`1px solid ${T.line}`,borderRadius:T.rMd,padding:"9px 12px",cursor:"pointer",fontSize:13,color:T.ink,textAlign:"left"}}>
          <Icon name="task" size={15}/>
          <span style={{flex:1}}><b>{selected.size}</b> de {records.length} profissional(is) selecionado(s){selected.size>0?` · ${brl(records.filter(r=>selected.has(r.id)).reduce((s,r)=>s+(r.valorTotal||0),0))}`:""}</span>
          <Icon name={pickOpen?"chevronUp":"chevronDown"} size={16}/>
        </button>
        {pickOpen && (() => {
          const list = records.filter(r=>!qProf.trim() || (r.profissional||"").toLowerCase().includes(qProf.trim().toLowerCase()));
          return (
            <div style={{border:`1px solid ${T.line}`,borderTop:"none",borderRadius:`0 0 ${T.rMd}px ${T.rMd}px`,padding:"10px 12px",background:T.canvas}}>
              <div style={{display:"flex",gap:8,marginBottom:8,alignItems:"center"}}>
                <input autoFocus style={{...inp,flex:1,fontSize:12,padding:"6px 9px"}} placeholder="Buscar profissional…" value={qProf} onChange={e=>setQProf(e.target.value)}/>
                <Btn small onClick={toggleAll}>{selected.size===records.length?"Limpar":"Todos"}</Btn>
              </div>
              <div className="fc-scroll" style={{maxHeight:240,overflowY:"auto",display:"flex",flexDirection:"column",gap:2}}>
                {list.length===0 ? <div style={{fontSize:12,color:T.muted,padding:"8px 4px"}}>Nenhum profissional.</div>
                  : list.map(r=>(
                    <label key={r.id} style={{display:"flex",alignItems:"center",gap:9,padding:"7px 8px",borderRadius:T.rSm,cursor:"pointer",background:selected.has(r.id)?T.brandBg:"transparent",fontSize:13}}>
                      <input type="checkbox" checked={selected.has(r.id)} onChange={()=>toggle(r.id)} style={{width:15,height:15}}/>
                      <span style={{flex:1,fontWeight:selected.has(r.id)?600:400,color:selected.has(r.id)?T.brand:T.inkSoft}}>{r.profissional}</span>
                      <span style={{fontSize:12,fontWeight:600,color:T.ink,whiteSpace:"nowrap",fontVariantNumeric:"tabular-nums"}}>{brl(r.valorTotal||0)}</span>
                      <Badge label={calcStatus(r.progress)} color={calcStatusColor(r.progress)} small dot/>
                    </label>
                  ))}
              </div>
            </div>
          );
        })()}
        {selected.size>0&&<div style={{fontSize:11,color:T.muted,marginTop:6}}>Os passos abaixo serão aplicados aos <b>{selected.size}</b> selecionado(s).</div>}
      </div>

      {/* Passos */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(260px,1fr))",gap:12,marginBottom:16}}>
        {grupos.map(g=>{
          const gs = groupState(sharedProg, g.num);
          return (
          <div key={g.num} style={{background:T.canvas,borderRadius:T.rLg,padding:"12px 14px",border:`1px solid ${gs==="done"?T.okLine:T.line}`}}>
            <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:10}}>
              <div style={{width:22,height:22,borderRadius:"50%",background:gs==="done"?T.ok:T.brand,color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:700,flexShrink:0}}>{gs==="done"?"✓":g.num}</div>
              <span style={{fontWeight:700,fontSize:13,color:T.ink}}>{g.title}</span>
            </div>
            {STEPS.filter(s=>g.steps.includes(s.id)).map(s=>{
              return (
                <div key={s.id} style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:8,gap:8}}>
                  <span style={{fontSize:12,color:T.inkSoft,flex:1}}>{s.name}</span>
                  {s.type==="check"
                    ? <label style={{display:"flex",alignItems:"center",gap:5,cursor:"pointer",fontSize:12,whiteSpace:"nowrap"}}>
                        <input type="checkbox" checked={!!sharedProg[s.id]} onChange={e=>setCheck(s.id,e.target.checked)} style={{width:15,height:15}}/>
                        <span style={{color:sharedProg[s.id]?T.ok:T.muted}}>{sharedProg[s.id]?"✓ Feito":"Pendente"}</span>
                      </label>
                    : <input type="date" value={sharedProg[s.id]||""} onChange={e=>setVal(s.id,e.target.value)} style={{...inp,width:150}}/>
                  }
                </div>
              );
            })}
          </div>
        );})}
      </div>

      {/* Ordem de venda — por profissional (organização interna do analista) */}
      <div style={{marginBottom:18}}>
        <div style={{fontSize:13,fontWeight:700,color:T.ink,marginBottom:8}}>Ordem de venda <span style={{fontWeight:400,color:T.muted,fontSize:11}}>(opcional · por profissional · uso interno)</span></div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(240px,1fr))",gap:8}}>
          {records.map(r=>(
            <div key={r.id} style={{display:"flex",alignItems:"center",gap:8}}>
              <span style={{fontSize:12,color:T.inkSoft,flex:1,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}} title={r.profissional}>{r.profissional}</span>
              <input style={{...inp,width:140,fontSize:12,padding:"5px 8px"}} placeholder="OV…" value={ovMap[r.id]||""} onChange={e=>setOV(r.id,e.target.value)}/>
            </div>
          ))}
        </div>
      </div>

      {/* Obs */}
      <div style={{marginBottom:16}}>
        <Field label="Observações (opcional)"><textarea value={obs} onChange={e=>setObs(e.target.value)} placeholder="Observações para todos os selecionados..." style={{...inp,minHeight:60,resize:"vertical"}}/></Field>
      </div>

      {error&&<div style={{marginBottom:12,fontSize:13,padding:"8px 12px",borderRadius:T.rMd,background:T.dangerBg,color:T.danger,border:`1px solid ${T.dangerLine}`}}>{error}</div>}

      <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
        <Btn onClick={onClose}>Cancelar</Btn>
        <Btn primary onClick={handleSave} disabled={selected.size===0}>Salvar — {selected.size} profissional(is)</Btn>
      </div>
    </Modal>
  );
}

// ─── NOTAS FISCAIS POR CLIENTE ───────────────────────────────────────────────
// Agrupa profissionais de um cliente em NFs (sem regra fixa: 1, 2, 3...).
// Permite NF X para um grupo e NF Y para outro, mostrando o valor somado de cada.

function NFGroupModal({ cliente, pep, records, onSave, onClose }) {
  const tipo = records[0]?.tipo;
  // Pré-requisito para emitir a NF: a última etapa "check" antes da emissão no
  // funil do tipo (T&E: aprovação do cliente; Fee/WIP: racional).
  const preReqId = (() => { const flow = funnelGroups(tipo).flatMap(g => STEPS.filter(s=>g.steps.includes(s.id) && s.type==="check").map(s=>s.id)); const i = flow.indexOf("p5_nf"); return i>0 ? flow[i-1] : null; })();
  const preReqName = STEPS.find(s=>s.id===preReqId)?.name || "etapa anterior";
  const [localRecs, setLocalRecs] = useState(records);
  const [dirty, setDirty]   = useState(new Set());
  const [selected, setSel]  = useState(new Set());
  const [nf, setNf]         = useState("");
  const [dataNf, setDataNf] = useState("");
  const [emitida, setEmit]  = useState(true);
  const [error, setError]   = useState("");

  const toggle = (id) => setSel(s => { const n=new Set(s); n.has(id)?n.delete(id):n.add(id); return n; });
  const selRecs = localRecs.filter(r => selected.has(r.id));
  const valorFinal = selRecs.reduce((a,r)=>a+(r.valorTotal||0), 0);

  // Agrupa por número de NF já atribuído
  const groupsMap = {};
  localRecs.forEach(r => { const k=(r.nfNumero||"").trim(); if(!k) return; (groupsMap[k]=groupsMap[k]||{nf:k,recs:[],total:0}); groupsMap[k].recs.push(r); groupsMap[k].total+=(r.valorTotal||0); });
  const groups = Object.values(groupsMap);
  const semNf = localRecs.filter(r => !(r.nfNumero||"").trim());

  function assign() {
    if (selected.size === 0) { setError("Selecione ao menos um profissional."); return; }
    const num = nf.trim();
    if (emitida && !num) { setError("Informe o número da NF para marcar como emitida."); return; }
    // Não é possível incluir número de NF sem marcar como emitida.
    if (num && !emitida) { setError("Para incluir o número da NF, marque “Marcar como emitida (P5)”."); return; }
    // Não é possível emitir a NF sem concluir a etapa anterior do funil do tipo.
    if (num && emitida && preReqId) {
      const semOk = selRecs.filter(r => !r.progress?.[preReqId]);
      if (semOk.length) { setError(`Conclua "${preReqName}" antes de emitir a NF para: ${semOk.map(r=>r.profissional||r.pep).join(", ")}.`); return; }
    }
    const now = nowISO();
    setLocalRecs(list => list.map(r => {
      if (!selected.has(r.id)) return r;
      const prog = { ...(r.progress||{}) };
      if (num) { if (emitida) { prog.p5_nf = true; if (dataNf) prog.p5_data_nf = dataNf; } }
      else { prog.p5_nf = false; prog.p5_data_nf = ""; }
      return { ...r, nfNumero: num, progress: prog, updatedAt: now };
    }));
    setDirty(d => { const n=new Set(d); selected.forEach(id=>n.add(id)); return n; });
    setSel(new Set()); setNf(""); setDataNf(""); setError("");
  }

  function clearNF(num) {
    const now = nowISO();
    const ids = [];
    setLocalRecs(list => list.map(r => {
      if ((r.nfNumero||"").trim() !== num) return r;
      ids.push(r.id);
      const prog = { ...(r.progress||{}), p5_nf: false, p5_data_nf: "" };
      return { ...r, nfNumero: "", progress: prog, updatedAt: now };
    }));
    setDirty(d => { const n=new Set(d); ids.forEach(id=>n.add(id)); return n; });
  }

  function handleSave() {
    const changed = localRecs.filter(r => dirty.has(r.id));
    if (changed.length) onSave(changed);
    onClose();
  }

  const palette = ["blue","green","teal","purple","orange","yellow"];

  return (
    <Modal title={`Notas fiscais — ${cliente}`} subtitle={`${pep} · ${localRecs.length} profissionais`} onClose={onClose} extraWide>
      {/* NFs já montadas */}
      <div style={{marginBottom:18}}>
        <div style={{fontSize:13,fontWeight:700,color:T.ink,marginBottom:8}}>NFs montadas neste cliente</div>
        {groups.length===0
          ? <div style={{fontSize:12,color:T.muted,padding:"10px 12px",background:T.canvas,borderRadius:T.rMd,border:`1px dashed ${T.line}`}}>Nenhuma NF montada ainda. Selecione os profissionais abaixo, informe o número e clique em “Atribuir NF”.</div>
          : <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(220px,1fr))",gap:10}}>
              {groups.map((g,i)=>(
                <div key={g.nf} style={{border:`1px solid ${T.line}`,borderRadius:T.rLg,padding:"12px 14px",background:"var(--surface)"}}>
                  <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:6,gap:8}}>
                    <Badge label={`NF ${g.nf}`} color={palette[i%palette.length]} small/>
                    <span style={{fontSize:15,fontWeight:800,color:T.ink}}>{fmtShort(g.total)}</span>
                  </div>
                  <div style={{fontSize:11,color:T.muted,lineHeight:1.5}}>{g.recs.map(r=>r.profissional).join(", ")}</div>
                  <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginTop:6,gap:8}}>
                    <span style={{fontSize:11,color:T.faint}}>{g.recs.length} profissional(is)</span>
                    <button onClick={()=>clearNF(g.nf)} title="Excluir esta NF (limpa o número dos profissionais)"
                      style={{display:"inline-flex",alignItems:"center",gap:4,background:"none",border:"none",cursor:"pointer",color:T.danger,fontSize:11,fontWeight:600,padding:0}}>✕ Excluir NF</button>
                  </div>
                </div>
              ))}
            </div>}
        {groups.length>0 && semNf.length>0 && <div style={{fontSize:11,color:T.warn,marginTop:8}}>{semNf.length} profissional(is) ainda sem NF.</div>}
      </div>

      {/* Montar nova NF */}
      <div style={{background:T.canvas,border:`1px solid ${T.line}`,borderRadius:T.rLg,padding:"14px 16px"}}>
        <div style={{fontSize:13,fontWeight:700,color:T.ink,marginBottom:10}}>Atribuir NF a profissionais</div>

        <div style={{display:"flex",flexWrap:"wrap",gap:8,marginBottom:12}}>
          {localRecs.map(r=>{
            const on = selected.has(r.id);
            const cur = (r.nfNumero||"").trim();
            return (
              <label key={r.id} style={{display:"flex",alignItems:"center",gap:8,padding:"8px 12px",borderRadius:T.rMd,border:`1.5px solid ${on?T.brand:T.line}`,background:on?T.brandBg:"var(--surface)",cursor:"pointer",fontSize:13}}>
                <input type="checkbox" checked={on} onChange={()=>toggle(r.id)} style={{width:15,height:15}}/>
                <span style={{fontWeight:on?600:400,color:on?T.brand:T.inkSoft}}>{r.profissional}</span>
                <span style={{fontSize:12,color:T.muted}}>{fmtShort(r.valorTotal)}</span>
                {cur && <Badge label={`NF ${cur}`} color="gray" small/>}
              </label>
            );
          })}
        </div>

        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(160px,1fr))",gap:12,alignItems:"end"}}>
          <Field label="Número da NF"><input style={inp} value={nf} onChange={e=>{setNf(e.target.value);setError("");}} placeholder="Ex: 123456 (vazio = limpar)"/></Field>
          <Field label="Data de emissão"><input type="date" style={inp} value={dataNf} onChange={e=>setDataNf(e.target.value)}/></Field>
          <label style={{display:"flex",alignItems:"center",gap:8,fontSize:13,color:T.inkSoft,cursor:"pointer",paddingBottom:8}}>
            <input type="checkbox" checked={emitida} onChange={e=>setEmit(e.target.checked)} style={{width:16,height:16}}/>
            Marcar como emitida (P5)
          </label>
        </div>

        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",gap:12,marginTop:12,flexWrap:"wrap"}}>
          <div style={{fontSize:13,color:T.inkSoft}}>
            Valor final da NF: <b style={{fontSize:16,color:T.ink}}>{fmtShort(valorFinal)}</b>
            <span style={{fontSize:12,color:T.muted}}> · {selected.size} selecionado(s)</span>
          </div>
          <Btn primary onClick={assign} disabled={selected.size===0}>Atribuir NF aos selecionados</Btn>
        </div>
      </div>

      {error&&<div style={{marginTop:12,fontSize:13,padding:"8px 12px",borderRadius:T.rMd,background:T.dangerBg,color:T.danger,border:`1px solid ${T.dangerLine}`}}>{error}</div>}

      <div style={{display:"flex",gap:8,justifyContent:"flex-end",marginTop:18}}>
        <Btn onClick={onClose}>Cancelar</Btn>
        <Btn primary onClick={handleSave} disabled={dirty.size===0}>Salvar alterações</Btn>
      </div>
    </Modal>
  );
}

// Edição de um registro importado (somente admin).
function RecordEditModal({ record, conciliado, novo=false, onSave, onClose }) {
  const [f, setF] = useState(record);
  const set = (k,v) => setF(p=>({...p,[k]:v}));
  // Campos de valor: guardam o TEXTO enquanto edita (aceita vírgula nas casas
  // decimais) e só viram número ao salvar — sem "comer" a vírgula a cada tecla.
  const numStr = (k,v) => setF(p=>({...p,[k]: String(v).replace(/[^\d.,-]/g,"")}));
  const toNum = (v) => { const s=String(v==null?"":v).trim(); if(!s) return 0; return s.includes(",") ? (parseFloat(s.replace(/\./g,"").replace(",","."))||0) : (parseFloat(s)||0); };
  const lockVal = { ...inp, background:"#f1f5f9", color:T.muted, cursor:"not-allowed" };
  // No modo "novo" exige os campos de identidade antes de deixar salvar.
  const faltando = [ !(f.cliente||"").trim()&&"Cliente", !(f.pep||"").trim()&&"PEP", !(f.competencia||"").trim()&&"Competência" ].filter(Boolean);
  const podeSalvar = !novo || faltando.length===0;
  function save() { if(!podeSalvar) return; onSave({ ...record, ...f,
    valorVenda:toNum(f.valorVenda), hrsAprovadas:toNum(f.hrsAprovadas), valorTotal:toNum(f.valorTotal), valorLiquido:toNum(f.valorLiquido),
    updatedAt: nowISO() }); onClose(); }
  return (
    <Modal title={novo?"Incluir registro":"Editar registro"} subtitle={novo?"Lançamento que o time esqueceu — preencha os dados":`${record.cliente} · ${record.profissional}`} onClose={onClose} wide>
      {conciliado && <div style={{display:"flex",alignItems:"center",gap:8,fontSize:12,padding:"9px 12px",borderRadius:T.rMd,background:"#fef2f2",border:"1px solid #fecaca",color:"#991b1b",fontWeight:600,marginBottom:14}}>
        <Icon name="lock" size={14}/> Registro <b>conciliado</b> — a NF é a verdade final. Os valores estão bloqueados. Reabra a conciliação para alterar.
      </div>}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(180px,1fr))",gap:12,marginBottom:16}}>
        <Field label="Responsável"><input style={inp} value={f.responsavel||""} onChange={e=>set("responsavel",e.target.value)}/></Field>
        <Field label="Empresa"><select style={inp} value={f.empresa||""} onChange={e=>set("empresa",e.target.value)}>{EMPRESAS.map(e=><option key={e.cod} value={e.cod}>{e.cod} — {e.nome}</option>)}</select></Field>
        <Field label="Tipo"><select style={inp} value={f.tipo||""} onChange={e=>set("tipo",e.target.value)}>{TIPOS_PROJETO.map(t=><option key={t}>{t}</option>)}</select></Field>
        <Field label="BU (comercial)"><select style={inp} value={f.bu||""} onChange={e=>set("bu",e.target.value)}><option value="">— sem BU —</option>{BUS.map(b=><option key={b}>{b}</option>)}</select></Field>
        <Field label="Competência"><input style={inp} placeholder="MM/AAAA" value={f.competencia||""} onChange={e=>set("competencia",e.target.value)}/></Field>
        <Field label="Cód. Cliente"><input style={inp} value={f.codCliente||""} onChange={e=>set("codCliente",e.target.value)}/></Field>
        <Field label="Cliente"><input style={inp} value={f.cliente||""} onChange={e=>set("cliente",e.target.value)}/></Field>
        <Field label="PEP"><input style={inp} value={f.pep||""} onChange={e=>set("pep",e.target.value)}/></Field>
        <Field label="Profissional"><input style={inp} value={f.profissional||""} onChange={e=>set("profissional",e.target.value)}/></Field>
        <Field label="Início"><input style={inp} value={f.inicio||""} onChange={e=>set("inicio",e.target.value)}/></Field>
        <Field label="Fim"><input style={inp} value={f.fim||""} onChange={e=>set("fim",e.target.value)}/></Field>
        <Field label="Valor de venda"><input style={conciliado?lockVal:inp} inputMode="decimal" disabled={conciliado} value={f.valorVenda} onChange={e=>numStr("valorVenda",e.target.value)}/></Field>
        <Field label="Hrs aprovadas"><input style={conciliado?lockVal:inp} inputMode="decimal" disabled={conciliado} value={f.hrsAprovadas} onChange={e=>numStr("hrsAprovadas",e.target.value)}/></Field>
        <Field label="Valor total"><input style={conciliado?lockVal:inp} inputMode="decimal" disabled={conciliado} value={f.valorTotal} onChange={e=>numStr("valorTotal",e.target.value)}/></Field>
        <Field label="Valor líquido"><input style={conciliado?lockVal:inp} inputMode="decimal" disabled={conciliado} value={f.valorLiquido} onChange={e=>numStr("valorLiquido",e.target.value)}/></Field>
      </div>
      <div style={{marginBottom:16}}><Field label="Observações"><textarea style={{...inp,minHeight:54,resize:"vertical"}} value={f.obs||""} onChange={e=>set("obs",e.target.value)}/></Field></div>
      {novo && faltando.length>0 && <div style={{fontSize:12,color:T.warn,marginBottom:10}}>Preencha para incluir: <b>{faltando.join(", ")}</b>.</div>}
      <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
        <Btn onClick={onClose}>Cancelar</Btn>
        <Btn primary disabled={!podeSalvar} onClick={save}>{novo?"Incluir registro":"Salvar registro"}</Btn>
      </div>
    </Modal>
  );
}

// Lançar variação de receita pós-fechamento (ex.: Casas Bahia). Não altera a
// receita original — cria um lançamento faturável à parte, com histórico.
function VariacaoModal({ record, lancamentos, onAdd, onDelete, onClose }) {
  const [valor, setValor] = useState("");
  const [motivo, setMotivo] = useState("");
  const total = lancamentos.reduce((s,v)=>s+(v.valor||0),0);
  function add() {
    const v = parseBR(valor);
    if (!v) return;
    onAdd(v, motivo.trim());
    setValor(""); setMotivo("");
  }
  return (
    <Modal title="Variação de receita" subtitle={`${record.cliente} · ${record.profissional||record.pep}`} onClose={onClose}
      footer={<Btn onClick={onClose}>Fechar</Btn>}>
      <div style={{fontSize:12.5,color:T.inkSoft,background:C.purple.bg,border:`1px solid ${C.purple.border}`,borderRadius:T.rMd,padding:"10px 12px",marginBottom:16}}>
        Ajuste <b>pós-fechamento</b>: não altera a receita original (que continua a verdade final). O valor entra como <b>saldo a faturar</b> — emita uma NF própria — e aparece nos relatórios como variação.
      </div>
      <div style={{display:"flex",gap:8,alignItems:"flex-end",marginBottom:6,flexWrap:"wrap"}}>
        <Field label="Valor da variação (R$)"><input style={{...inp,width:150}} placeholder="0,00" value={valor} onChange={e=>setValor(e.target.value)} onKeyDown={e=>{if(e.key==="Enter")add();}}/></Field>
        <Field label="Motivo (opcional)"><input style={{...inp,minWidth:200}} placeholder="Ex.: valor extra do cliente" value={motivo} onChange={e=>setMotivo(e.target.value)} onKeyDown={e=>{if(e.key==="Enter")add();}}/></Field>
        <Btn primary onClick={add} disabled={!parseBR(valor)}>Lançar</Btn>
      </div>
      <div style={{marginTop:14}}>
        <div style={{display:"flex",justifyContent:"space-between",fontSize:12,color:T.muted,marginBottom:6}}>
          <span>Lançamentos ({lancamentos.length})</span>
          <span>Total: <b style={{color:C.purple.solid}}>{brl(total)}</b></span>
        </div>
        {lancamentos.length===0 ? <div style={{fontSize:12,color:T.muted,padding:"10px 0"}}>Nenhuma variação lançada ainda.</div>
          : <div style={{display:"flex",flexDirection:"column",gap:6}}>
            {lancamentos.slice().sort((a,b)=>String(b.createdAt).localeCompare(String(a.createdAt))).map(v=>(
              <div key={v.id} style={{display:"flex",alignItems:"center",gap:10,padding:"8px 11px",border:`1px solid ${T.line}`,borderRadius:T.rSm,fontSize:12.5}}>
                <span style={{fontWeight:700,color:C.purple.solid,whiteSpace:"nowrap"}}>{brl(v.valor)}</span>
                <span style={{flex:1,color:T.inkSoft,overflow:"hidden",textOverflow:"ellipsis"}}>{v.motivo||"—"}</span>
                <span style={{fontSize:11,color:T.faint,whiteSpace:"nowrap"}}>{v.criadoPor}{v.createdAt?` · ${new Date(v.createdAt).toLocaleDateString("pt-BR")}`:""}</span>
                <button title="Remover" onClick={()=>onDelete(v.id)} style={{border:"none",background:"none",cursor:"pointer",color:T.danger}}><Icon name="trash" size={13}/></button>
              </div>
            ))}
          </div>}
      </div>
    </Modal>
  );
}

// Classificar não-faturado — categoria automática (dentro do ciclo × represado,
// derivada da data) + motivo padrão + observação livre. Reutilizada na Minha
// visão, na Visão por projeto e na Conciliação. onSave(id,{motivo,obs}).
function ClassifyModal({ record:r, clients=[], fatByRec={}, varByRec={}, onSave, onClose }) {
  const [motivo, setMotivo] = useState(r.classMotivo||"");
  const [obs, setObs]       = useState(r.classObs||"");
  const { cat, compFat }    = categoriaOf(r, clients);
  const repres = cat==="represado";
  const bill = (r.valorTotal||0) + (varByRec[r.id]||0);
  const fat  = fatByRec[r.id]||0;
  const save = () => { onSave(r.id, { motivo, obs }); onClose(); };
  return (
    <Modal title="Classificar não-faturado" subtitle={`${r.profissional||r.pep||r.cliente} · ${r.cliente}`} onClose={onClose}
      footer={<><Btn onClick={onClose}>Cancelar</Btn><Btn primary onClick={save}>Salvar</Btn></>}>
      <div style={{display:"flex",gap:9,flexWrap:"wrap",alignItems:"center",marginBottom:10}}>
        <Badge label={repres?"Represado":"Dentro do ciclo"} color={repres?"red":"teal"} dot/>
        <span style={{fontSize:12,color:T.muted}}>competência de faturamento <b style={{color:T.inkSoft}}>{compFat||"—"}</b> · categoria automática pela data</span>
      </div>
      <div style={{fontSize:11.5,color:T.muted,marginBottom:16}}>{r.pep} · {r.tipo} · {r.inicio||"—"}–{r.fim||"—"} · faturável {brl(bill)}{fat>0.01?` · faturado ${brl(fat)}`:""}</div>
      <div style={{marginBottom:16}}>
        <div style={Ty.label}>Status por etapa</div>
        <PipelineStepper states={recordStates(r.progress, r.tipo)} groups={funnelGroups(r.tipo)} showLabels/>
      </div>
      <div style={{marginBottom:14}}><Field label="Motivo do não-faturamento">
        <select style={inp} value={motivo} onChange={e=>setMotivo(e.target.value)}>
          <option value="">Selecione um motivo…</option>
          {CLASS_MOTIVOS.map(m=><option key={m}>{m}</option>)}
        </select>
      </Field></div>
      <Field label={repres?"Motivo do represamento (texto livre)":"Observação (texto livre)"}>
        <textarea style={{...inp,minHeight:70,resize:"vertical"}} value={obs} onChange={e=>setObs(e.target.value)} placeholder={repres?"Explique por que represou…":"Detalhe, se necessário…"}/>
      </Field>
    </Modal>
  );
}

// Chip clicável de classificação — mostra o motivo/categoria de um não-faturado.
// Vermelho quando represado; tracejado "+ classificar" quando ainda sem motivo.
function ClassifyChip({ record:r, clients=[], onClick, readOnly, style:s={} }) {
  const { cat } = categoriaOf(r, clients);
  const repres = cat==="represado";
  const has = !!(r.classMotivo || r.classObs);
  const base = { marginLeft:6, cursor:readOnly?"default":"pointer", borderRadius:T.rPill, padding:"2px 8px", fontSize:10, fontWeight:600, verticalAlign:"middle", ...s };
  // Só leitura (Minha visão / Visão por projeto): mostra o estado, não edita.
  // A edição vive só na aba Represados. Sem motivo e dentro do ciclo → nada.
  if (readOnly) {
    if (!has && !repres) return null;
    const label = has ? (repres?"⚠ ":"")+(r.classMotivo||"obs") : "⚠ represado";
    return (
      <span title={`${r.classMotivo||(repres?"represado":"—")}${r.classObs?` — ${r.classObs}`:""}${has?"":" · classifique na aba Represados"}`}
        style={{...base, cursor:"default", border:`1px solid ${repres?C.red.border:C.gray.border}`, background:repres?C.red.bg:C.gray.bg, color:repres?C.red.text:C.gray.text}}>
        {label}
      </span>
    );
  }
  if (has) return (
    <button onClick={onClick} title={`${r.classMotivo||"—"}${r.classObs?` — ${r.classObs}`:""} (clique p/ editar)`}
      style={{...base, border:`1px solid ${repres?C.red.border:C.gray.border}`, background:repres?C.red.bg:C.gray.bg, color:repres?C.red.text:C.gray.text}}>
      {repres?"⚠ ":""}{r.classMotivo||"obs"}
    </button>
  );
  return (
    <button onClick={onClick} title="Classificar não-faturado (motivo / observação)"
      style={{...base, border:`1px dashed ${repres?C.red.border:T.line}`, background:"var(--surface)", color:repres?C.red.solid:T.muted}}>
      {repres?"⚠ classificar":"+ classificar"}
    </button>
  );
}

// ─── MY VIEW (team) ──────────────────────────────────────────────────────────

function MyView({ records, clients=[], analista, isAdmin, isViewer=false, fatByRec={}, varByRec={}, varsByRec={}, aceitaVar=()=>false, onAddVariacao, onDelVariacao, onUpdateBulk, onDeleteRecord, onClearAlert, onSaveClass, competenciaAtual, onCompetenciaChange }) {
  const isMobile = useIsMobile();
  const seeAll = isAdmin || isViewer;   // enxerga todos os analistas (admin ou viewer)
  const [varTarget, setVarTarget] = useState(null);   // registro para lançar variação
  const bill = (r) => (r.valorTotal||0) + (varByRec[r.id]||0);   // faturável (receita + variação)
  const [recordEdit, setRecEdit] = useState(null);
  const [recordDel, setRecDel]   = useState(null);
  const [empresa, setEmpresa]       = useState("");
  const [tipo, setTipo]             = useState("");
  const [filterComp, setFilterComp] = useState(competenciaAtual);
  const [filterAnalista, setFA]     = useState("todos");
  const [filterEtapa, setFEt]       = useState("todas");
  // Filtro composável: escolhe a dimensão (Cliente/Profissional/PEP) e o valor;
  // "+" adiciona outra dimensão. Combina em E (todas precisam bater).
  const [filtros, setFiltros]       = useState([{ dim:"cliente", val:"" }]);
  const FDIMS = { cliente:{label:"Cliente", get:r=>r.cliente}, profissional:{label:"Profissional", get:r=>r.profissional}, pep:{label:"PEP", get:r=>pepBase(r.pep)} };
  const updF = (i,patch)=> setFiltros(fs=>fs.map((f,j)=>j===i?{...f,...patch}:f));
  const rmF  = (i)=> setFiltros(fs=>fs.length>1?fs.filter((_,j)=>j!==i):fs);
  const addF = ()=> setFiltros(fs=>{ const used=new Set(fs.map(f=>f.dim)); const next=Object.keys(FDIMS).find(k=>!used.has(k)); return next?[...fs,{dim:next,val:""}]:fs; });
  const [expandedCliente, setExp]   = useState(null);
  const [bulkTarget, setBulk]       = useState(null);
  const [nfTarget, setNf]           = useState(null);

  // O banco (RLS) já entrega apenas os registros do analista, vinculados pelo
  // "Responsável na base". Não refiltramos pelo nome de exibição.
  const myRecords = records;
  const competencias = [...new Set(records.map(r=>r.competencia))].sort();
  const analistas = [...new Set(records.map(r=>r.responsavel))].sort();
  const empresasUsed = [...new Set(myRecords.map(r=>r.empresa))];
  const tiposUsed = [...new Set((empresa?myRecords.filter(r=>r.empresa===empresa):myRecords).map(r=>r.tipo))].filter(Boolean).sort();

  let filtered = myRecords;
  if (empresa) filtered = filtered.filter(r=>r.empresa===empresa);
  if (tipo)    filtered = filtered.filter(r=>r.tipo===tipo);
  if (filterComp!=="todas") filtered = filtered.filter(r=>r.competencia===filterComp);
  if (seeAll && filterAnalista!=="todos") filtered = filtered.filter(r=>r.responsavel===filterAnalista);
  if (filterEtapa==="_faltam_datas") filtered = filtered.filter(faltaDatas);
  else if (filterEtapa!=="todas") filtered = filtered.filter(r=>recStatus(r, fatByRec[r.id], bill(r))===filterEtapa);
  filtros.forEach(f=>{ const v=(f.val||"").trim().toLowerCase(); const g=FDIMS[f.dim]?.get; if(v&&g) filtered = filtered.filter(r=>String(g(r)||"").toLowerCase().includes(v)); });

  // Resumo por tipo de contrato
  const porTipo = {};
  // "Fora do relatório" (ausente) não conta nos totais — só aparece sinalizado.
  filtered.forEach(r=>{ if(r.ausenteRelatorio) return; const t=r.tipo||"—"; if(!porTipo[t]) porTipo[t]={count:0,total:0,fat:0}; porTipo[t].count++; porTipo[t].total+=bill(r); porTipo[t].fat+=(fatByRec[r.id]||0); });
  const tipoColors = { "Time & Expenses":"blue", "Fee":"purple", "WIP":"teal", "Usage Based":"orange" };

  const grouped = {};
  filtered.forEach(r=>{
    const key = r.cliente+"|"+pepBase(r.pep);
    if(!grouped[key]) grouped[key]={ cliente:r.cliente, pep:pepBase(r.pep), records:[] };
    grouped[key].records.push(r);
  });
  const groups = Object.values(grouped);
  const selW = isMobile ? "100%" : "auto";

  return (
    <div>
      {bulkTarget&&<BulkTimelineModal {...bulkTarget} onClose={()=>setBulk(null)} onSave={updated=>{onUpdateBulk(updated);setBulk(null);}}/>}
      {recordEdit&&<RecordEditModal record={recordEdit} conciliado={(fatByRec[recordEdit.id]||0)>0.001} onClose={()=>setRecEdit(null)} onSave={r=>{onUpdateBulk([r]);setRecEdit(null);}}/>}
      {recordDel&&<ConfirmDialog title="Excluir registro" danger confirmLabel="Excluir"
        message={`Excluir o registro de "${recordDel.profissional}" (${recordDel.cliente})? Esta ação não pode ser desfeita.${(fatByRec[recordDel.id]||0)>0.001?" ⚠️ Este registro está CONCILIADO — a(s) nota(s) do lote serão reabertas na conciliação." : ""}`}
        onConfirm={()=>onDeleteRecord(recordDel.id)} onClose={()=>setRecDel(null)}/>}
      {varTarget&&<VariacaoModal record={varTarget} lancamentos={(varsByRec[varTarget.id]||[])} onAdd={(valor,motivo)=>onAddVariacao(varTarget.id,valor,motivo)} onDelete={onDelVariacao} onClose={()=>setVarTarget(null)}/>}

      <PageHead icon="list" title="Minha visão" sub={`${groups.length} cliente(s) · ${filtered.length} registro(s)`}/>

      {/* Resumo por tipo de contrato */}
      {Object.keys(porTipo).length>0 && <div style={{display:"flex",gap:10,flexWrap:"wrap",marginBottom:14}}>
        {Object.entries(porTipo).sort((a,b)=>b[1].total-a[1].total).map(([t,d])=>{
          const c=C[tipoColors[t]||"gray"]||C.gray;
          return (
            <div key={t} style={{flex:"1 1 170px",border:`1px solid ${c.border}`,background:c.bg,borderRadius:T.rLg,padding:"10px 14px"}}>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",gap:8,marginBottom:4}}>
                <span style={{fontSize:12,fontWeight:700,color:c.text}}>{t}</span>
                <span style={{fontSize:11,color:c.text,opacity:.8}}>{d.count} reg.</span>
              </div>
              <div style={{fontSize:16,fontWeight:800,color:c.text}}>{fmtShort(d.total)}</div>
              <div style={{fontSize:11,color:c.text,opacity:.85}}>Faturado: {fmtShort(d.fat)}</div>
            </div>
          );
        })}
      </div>}

      {/* Filtros */}
      <Card style={{ padding:"12px 14px", marginBottom:16 }}>
        <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
          <select style={{...inp,width:selW,flex:isMobile?"none":1,minWidth:150}} value={filterComp} onChange={e=>{setFilterComp(e.target.value);onCompetenciaChange(e.target.value);}} aria-label="Competência">
            {competencias.map(c=><option key={c} value={c}>{c}</option>)}
            <option value="todas">Todas as competências</option>
          </select>
          <select style={{...inp,width:selW,flex:isMobile?"none":1,minWidth:150}} value={empresa} onChange={e=>{setEmpresa(e.target.value);setTipo("");}} aria-label="Empresa">
            <option value="">Todas as empresas</option>
            {(isAdmin?EMPRESAS:EMPRESAS.filter(e=>empresasUsed.includes(e.cod))).map(e=><option key={e.cod} value={e.cod}>{e.cod} — {e.nome}</option>)}
          </select>
          <select style={{...inp,width:selW,minWidth:150}} value={tipo} onChange={e=>setTipo(e.target.value)} aria-label="Tipo de contrato">
            <option value="">Todos os contratos</option>
            {tiposUsed.map(t=><option key={t}>{t}</option>)}
          </select>
          {seeAll&&<select style={{...inp,width:selW}} value={filterAnalista} onChange={e=>setFA(e.target.value)} aria-label="Analista">
            <option value="todos">Todos os analistas</option>
            {analistas.map(a=><option key={a}>{a}</option>)}
          </select>}
          <select style={{...inp,width:selW,minWidth:160}} value={filterEtapa} onChange={e=>setFEt(e.target.value)} aria-label="Etapa do funil">
            <option value="todas">Todas as etapas</option>
            {STATUS_ORDER.map(s=><option key={s}>{s}</option>)}
            <option value="_faltam_datas">Faltam datas</option>
          </select>
          {filtros.map((f,i)=>(
            <div key={i} style={{display:"flex",alignItems:"stretch"}}>
              <select value={f.dim} onChange={e=>updF(i,{dim:e.target.value})} aria-label="Tipo de filtro"
                style={{...inp,width:"auto",borderRadius:`${T.rMd} 0 0 ${T.rMd}`,borderRight:"none",background:T.canvas,fontWeight:600,color:T.inkSoft,paddingRight:6}}>
                {Object.entries(FDIMS).map(([k,d])=><option key={k} value={k}>{d.label}</option>)}
              </select>
              <input style={{...inp,width:isMobile?130:150,borderRadius:filtros.length>1?0:`0 ${T.rMd} ${T.rMd} 0`}} placeholder={`Filtrar por ${FDIMS[f.dim].label.toLowerCase()}…`} value={f.val} onChange={e=>updF(i,{val:e.target.value})}/>
              {filtros.length>1 && <button onClick={()=>rmF(i)} title="Remover filtro" style={{border:`1px solid ${T.line}`,borderLeft:"none",borderRadius:`0 ${T.rMd} ${T.rMd} 0`,background:"var(--surface)",color:T.muted,cursor:"pointer",padding:"0 9px",fontSize:15,lineHeight:1}}>×</button>}
            </div>
          ))}
          {filtros.length < Object.keys(FDIMS).length && <button onClick={addF} title="Adicionar filtro" style={{border:`1px dashed ${T.line}`,borderRadius:T.rMd,background:T.canvas,color:T.brand,cursor:"pointer",padding:"0 12px",fontSize:13,fontWeight:700}}>+ filtro</button>}
        </div>
      </Card>

      {/* Legenda do funil */}
      {groups.length>0 && <div style={{display:"flex",alignItems:"center",gap:14,flexWrap:"wrap",margin:"0 2px 12px",fontSize:11,color:T.muted}}>
        <span style={{fontWeight:600}}>Funil:</span>
        <span style={{display:"flex",alignItems:"center",gap:5}}><i style={{width:11,height:11,borderRadius:"50%",background:T.ok}}/> concluído</span>
        <span style={{display:"flex",alignItems:"center",gap:5}}><i style={{width:11,height:11,borderRadius:"50%",background:C.blue.solid}}/> parcial</span>
        <span style={{display:"flex",alignItems:"center",gap:5}}><i style={{width:11,height:11,borderRadius:"50%",background:"var(--surface)",border:"2px solid #cbd2dc"}}/> pendente</span>
        <span style={{color:T.faint}}>· {STEP_GROUPS.map(g=>g.num+" "+g.short).join("  ·  ")}</span>
      </div>}

      {groups.length===0&&<Card style={{textAlign:"center",padding:"3rem"}}>
        <div style={{fontSize:32,marginBottom:10}}></div>
        <div style={{fontSize:14,color:T.muted}}>Nenhum registro encontrado para os filtros selecionados.</div>
      </Card>}

      {/* Cards de clientes */}
      {groups.map(g=>{
        const ativos  = g.records.filter(r=>!r.ausenteRelatorio);   // fora do relatório não soma
        const total   = ativos.reduce((a,r)=>a+(r.valorTotal||0),0);
        const varG    = ativos.reduce((a,r)=>a+(varByRec[r.id]||0),0);
        const totalG  = ativos.reduce((a,r)=>a+bill(r),0);   // faturável (com variação)
        const fatG    = ativos.reduce((a,r)=>a+(fatByRec[r.id]||0),0);
        const pct     = totalG>0 ? Math.round(fatG/totalG*100) : 0;
        const isOpen  = expandedCliente===(g.cliente+g.pep);
        const fullG = g.records.filter(r=>{const t=bill(r);return Math.abs(t)>0.01 && Math.abs(t-(fatByRec[r.id]||0))<0.01;}).length;
        const anyG  = g.records.filter(r=>Math.abs(fatByRec[r.id]||0)>0.001).length;
        const overallStatus = fullG===g.records.length?"Faturado":anyG>0?"Faturado parcial":g.records.every(r=>r.progress?.p5_liberado)?"Liberado p/ faturamento":"Em andamento";
        const overallColor  = fullG===g.records.length?"green":anyG>0?"orange":g.records.every(r=>r.progress?.p5_liberado)?"teal":"yellow";
        const gtipo = g.records[0]?.tipo;
        const ggrupos = funnelGroups(gtipo);
        const agg = aggregateStates(g.records, gtipo);
        const temProf = g.records.some(r=>r.profissional);
        const faltam = g.records.filter(faltaDatas).length;
        const alertRed = g.records.filter(r=>r.valorAnterior!=null && (fatByRec[r.id]||0)>0.001).length;
        const alertYel = g.records.filter(r=>r.valorAnterior!=null && !((fatByRec[r.id]||0)>0.001)).length;
        const foraRel  = g.records.filter(r=>r.ausenteRelatorio).length;
        const divConc  = g.records.filter(r=>r.valorBaseDivergente!=null).length;

        return (
          <Card key={g.cliente+g.pep} interactive style={{marginBottom:10,overflow:"hidden"}}>
            {/* Cabeçalho do cliente — clicável */}
            <div onClick={()=>setExp(isOpen?null:(g.cliente+g.pep))} style={{display:"flex",alignItems:"center",gap:14,padding:"14px 18px",cursor:"pointer",userSelect:"none",flexWrap:"wrap"}}>
              <div style={{flex:1,minWidth:200}}>
                <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:5}}>
                  <span style={{fontSize:14,fontWeight:700,color:T.ink}}>{g.cliente}</span>
                  <Badge label={overallStatus} color={overallColor} small dot/>
                  {faltam>0 && <Badge label={`faltam datas (${faltam})`} color="yellow" small/>}
                  {divConc>0 && <Badge label={`⚠ base diverge do conciliado (${divConc}) — reabrir`} color="red" small/>}
                  {alertRed>0 && <Badge label={`⚠ valor mudou p/ faturado (${alertRed})`} color="red" small/>}
                  {alertYel>0 && <Badge label={`valor alterado (${alertYel})`} color="yellow" small/>}
                  {foraRel>0 && <Badge label={`fora do relatório (${foraRel})`} color="gray" small/>}
                </div>
                <div style={{fontSize:11,color:T.muted}}>{g.pep} · {gtipo} · {g.records.length} {temProf?"profissionais":"registro(s)"} · {fmtShort(total)}</div>
              </div>
              {!isMobile && <div style={{ width:230 }}><PipelineStepper states={agg} groups={ggrupos} showLabels/></div>}
              <div style={{textAlign:"center"}}>
                <div style={{fontSize:18,fontWeight:800,color:pct===100?T.ok:pct>50?T.brand:C.orange.solid}}>{pct}%</div>
                <div style={{fontSize:10,color:T.muted}}>faturado</div>
              </div>
              {!isViewer && <Btn small icon="pencil" onClick={e=>{e.stopPropagation();setBulk({cliente:g.cliente,pep:g.pep,records:g.records});}}>Atualizar passos</Btn>}
              <span style={{fontSize:16,color:T.faint}} aria-hidden="true">{isOpen?"▲":"▼"}</span>
            </div>

            {isMobile && <div style={{ padding:"0 18px 12px" }}><PipelineStepper states={agg} groups={ggrupos} showLabels/></div>}

            {/* Detalhe dos profissionais */}
            {isOpen&&<div style={{borderTop:`1px solid ${T.lineSoft}`,padding:"0 18px 14px"}}>
              <div className="fc-scroll" style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,marginTop:10}}>
                <thead><tr style={{background:T.canvas}}>
                  {[seeAll&&"Analista","Profissional","OV","Funil","Período","Val. Total","NF","Status",isAdmin&&"Ações"].filter(Boolean).map(h=>
                    <th key={h} style={{padding:"7px 10px",textAlign:h==="Ações"?"right":"left",borderBottom:`1px solid ${T.line}`,fontWeight:600,color:T.muted,whiteSpace:"nowrap"}}>{h}</th>
                  )}
                </tr></thead>
                <tbody>
                  {g.records.map(r=>{
                    const colCount = (seeAll?1:0)+7+(isAdmin?1:0);
                    const fatR = fatByRec[r.id]||0;
                    const saldoR = Math.max(0,(r.valorTotal||0)-fatR);
                    const saldoBill = bill(r) - fatR;                                   // saldo faturável (inclui variação)
                    const parcial = Math.abs(fatR)>0.01 && Math.abs(saldoBill)>0.01;     // faturado em parte
                    const temAlerta = r.valorAnterior!=null || r.ausenteRelatorio || r.valorBaseDivergente!=null;
                    const subRow = temAlerta || parcial;
                    const alertaFat = r.valorAnterior!=null && fatR>0.001;
                    return (
                    <Fragment key={r.id}>
                    <tr className="fc-row" style={{borderBottom:subRow?"none":`1px solid ${T.lineSoft}`}}>
                      {seeAll&&<td style={{padding:"7px 10px"}}><Badge label={r.responsavel} color="purple" small/></td>}
                      <td style={{padding:"7px 10px",fontWeight:500,color:T.ink}}>{r.profissional}</td>
                      <td style={{padding:"7px 10px",fontFamily:"monospace",fontSize:11,color:r.ordemVenda?T.inkSoft:T.faint}}>{r.ordemVenda||"—"}</td>
                      <td style={{padding:"7px 10px"}}><PipelineStepper states={recordStates(r.progress, r.tipo)} groups={funnelGroups(r.tipo)} size="sm"/></td>
                      <td style={{padding:"7px 10px",color:T.muted,whiteSpace:"nowrap"}}>{r.inicio} → {r.fim}</td>
                      <td style={{padding:"7px 10px",fontWeight:500,whiteSpace:"nowrap"}}>{fmtShort(r.valorTotal)}{(varByRec[r.id]||0)>0.001&&<div style={{fontSize:10.5,color:C.purple.solid,fontWeight:700}}>+ {fmtShort(varByRec[r.id])} variação</div>}</td>
                      <td style={{padding:"7px 10px",fontFamily:"monospace",fontSize:11}}>{r.nfNumero||"—"}</td>
                      <td style={{padding:"7px 10px",whiteSpace:"nowrap"}}>
                        <Badge label={recStatus(r, fatByRec[r.id], bill(r))} color={recStatusColor(r, fatByRec[r.id], bill(r))} small dot/>
                        {aceitaVar(r) && !isViewer && <button title="Lançar/ver variação de receita (pós-fechamento)" onClick={()=>setVarTarget(r)} style={{marginLeft:6,border:`1px solid ${C.purple.border}`,background:(varByRec[r.id]||0)>0.001?C.purple.bg:"var(--surface)",color:C.purple.solid,borderRadius:T.rSm,padding:"2px 7px",cursor:"pointer",fontSize:10.5,fontWeight:700,verticalAlign:"middle"}}>± variação</button>}
                        {fatR<0.01 && <ClassifyChip record={r} clients={clients} readOnly/>}
                      </td>
                      {isAdmin&&<td style={{padding:"7px 10px",textAlign:"right",whiteSpace:"nowrap"}}>
                        <button title="Editar registro" onClick={()=>setRecEdit(r)} style={{border:"none",background:"none",cursor:"pointer",color:T.muted,fontSize:14,padding:"0 4px"}}><Icon name="pencil" size={14}/></button>
                        <button title="Excluir registro" onClick={()=>setRecDel(r)} style={{border:"none",background:"none",cursor:"pointer",color:T.danger,fontSize:14,padding:"0 4px"}}><Icon name="trash" size={14}/></button>
                      </td>}
                    </tr>
                    {subRow&&<tr style={{borderBottom:`1px solid ${T.lineSoft}`}}><td colSpan={colCount} style={{padding:"0 10px 8px"}}>
                      {parcial && <div style={{display:"flex",alignItems:"center",gap:9,flexWrap:"wrap",fontSize:12,padding:"7px 11px",borderRadius:T.rMd,background:"#fff7ed",border:"1px solid #fed7aa",color:"#9a3412",fontWeight:600}}>
                        <Icon name="wallet" size={14}/>
                        <span><b>Faturamento parcial</b> — ✓ faturado {brl(fatR)} · <b>saldo a faturar {brl(saldoBill)}</b> (pendente de nova NF na conciliação)</span>
                      </div>}
                      {r.valorAnterior!=null && <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap",fontSize:12,padding:"7px 11px",borderRadius:T.rMd,background:alertaFat?"#fef2f2":"#fffbeb",border:`1px solid ${alertaFat?"#fecaca":"#fde68a"}`,color:alertaFat?"#991b1b":"#92400e",fontWeight:600,marginTop:parcial?6:0}}>
                        <Icon name={alertaFat?"alert":"info"} size={14}/>
                        {alertaFat
                          ? <span>Valor mudou <b>após faturamento</b>: {brl(r.valorAnterior)} → {brl(r.valorTotal)}. Já faturado {brl(fatR)}{saldoR>0.01?<> · <b>saldo a faturar {brl(saldoR)}</b></>:<> · <b>faturado a maior {brl(fatR-(r.valorTotal||0))} — NF a corrigir</b></>}. {r.nfNumero?`NF ${r.nfNumero} a revisar/cancelar.`:""}</span>
                          : <span>Valor alterado no relatório: <b>{brl(r.valorAnterior)} → {brl(r.valorTotal)}</b>.</span>}
                        {!isViewer && <button onClick={()=>onClearAlert&&onClearAlert(r.id)} style={{marginLeft:"auto",border:`1px solid ${alertaFat?"#fca5a5":"#fcd34d"}`,background:"var(--surface)",borderRadius:T.rSm,padding:"3px 10px",cursor:"pointer",fontSize:11,fontWeight:700,color:alertaFat?"#991b1b":"#92400e"}}>Ciente</button>}
                      </div>}
                      {r.valorBaseDivergente!=null && <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap",fontSize:12,padding:"7px 11px",borderRadius:T.rMd,background:"#fef2f2",border:`1px solid #fecaca`,color:"#991b1b",fontWeight:600,marginTop:r.valorAnterior!=null?6:0}}>
                        <Icon name="alert" size={14}/>
                        <span>A base mudou para <b>{brl(r.valorBaseDivergente)}</b>, mas está <b>conciliado</b> em <b>{brl(r.valorTotal)}</b>. A NF é a verdade final — <b>reabra a conciliação</b> para atualizar o valor.</span>
                      </div>}
                      {r.ausenteRelatorio && <div style={{display:"flex",alignItems:"center",gap:8,fontSize:12,padding:"7px 11px",borderRadius:T.rMd,background:T.canvas,border:`1px solid ${T.line}`,color:T.muted,fontWeight:600,marginTop:(r.valorAnterior!=null||r.valorBaseDivergente!=null)?6:0}}>
                        <Icon name="info" size={14}/> Fora do último relatório importado — verifique se saiu do projeto.
                        {!isViewer && <button onClick={()=>onClearAlert&&onClearAlert(r.id)} style={{marginLeft:"auto",border:`1px solid ${T.line}`,background:"var(--surface)",borderRadius:T.rSm,padding:"3px 10px",cursor:"pointer",fontSize:11,fontWeight:700,color:T.inkSoft}}>Ciente</button>}
                      </div>}
                    </td></tr>}
                    </Fragment>
                  );})}
                </tbody>
              </table>
              </div>
            </div>}
          </Card>
        );
      })}
    </div>
  );
}

// ─── DASHBOARD ───────────────────────────────────────────────────────────────

function Donut({ pct, size=120, label, sub }) {
  const r = (size-18)/2, cx=size/2, cy=size/2, circ=2*Math.PI*r;
  const off = circ*(1-pct/100);
  return (
    <div style={{ display:"flex", alignItems:"center", gap:16, flexWrap:"wrap" }}>
      <div style={{ position:"relative", width:size, height:size, flexShrink:0 }}>
        <svg width={size} height={size} style={{ transform:"rotate(-90deg)" }} aria-hidden="true">
          <circle cx={cx} cy={cy} r={r} fill="none" stroke="#fee2e2" strokeWidth={14}/>
          <circle cx={cx} cy={cy} r={r} fill="none" stroke={T.ok} strokeWidth={14} strokeDasharray={circ} strokeDashoffset={off} strokeLinecap="round" style={{ transition:"stroke-dashoffset .5s ease" }}/>
        </svg>
        <div style={{ position:"absolute", inset:0, display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center" }}>
          <span style={{ fontSize:24, fontWeight:800, color:T.ink }}>{pct}%</span>
          <span style={{ fontSize:10, color:T.muted }}>{label}</span>
        </div>
      </div>
      {sub}
    </div>
  );
}

function Dashboard({ records, analista, isAdmin, fatByRec={}, varByRec={} }) {
  const [filterEmpresa, setFE] = useState("todas");
  const [filterComp,    setFC] = useState("todas");
  const [filterAnalista,setFA] = useState(isAdmin?"todos":analista);
  const [filterEtapa,   setFEt]= useState("todas");
  const [filterTipo,    setFTi]= useState("todas");

  const comps     = [...new Set(records.map(r=>r.competencia))].sort();
  const analistas = [...new Set(records.map(r=>r.responsavel))].sort();
  const tipos     = [...new Set(records.map(r=>r.tipo))].filter(Boolean).sort();

  // RLS já restringe os registros do analista; usamos o que veio do banco.
  let base = records;
  let f = base;
  if (filterEmpresa!=="todas") f=f.filter(r=>r.empresa===filterEmpresa);
  if (filterTipo!=="todas")    f=f.filter(r=>r.tipo===filterTipo);
  if (filterComp!=="todas")    f=f.filter(r=>r.competencia===filterComp);
  if (isAdmin&&filterAnalista!=="todos") f=f.filter(r=>r.responsavel===filterAnalista);
  const fat = (r) => fatByRec[r.id]||0;                     // já faturado (alocado)
  const billOf = (r) => (r.valorTotal||0) + (varByRec[r.id]||0);   // faturável (receita + variação)
  const saldo = (r) => billOf(r) - fat(r);                  // a faturar (com sinal)
  const isFat = (r) => Math.abs(billOf(r))>0.01 && Math.abs(billOf(r)-fat(r))<0.01;  // faturado por completo (faturável)
  if (filterEtapa!=="todas")   f=f.filter(r=>recStatus(r, fatByRec[r.id], billOf(r))===filterEtapa);

  const totalValor = f.reduce((a,r)=>a+billOf(r),0);   // faturável (receita + variação)
  const naoFat     = f.filter(r=>Math.abs(saldo(r)) > 0.01); // tem saldo a faturar (inclui descontos)
  const valorFat   = f.reduce((a,r)=>a+fat(r),0);           // faturado = só o emitido
  const valorRep   = f.reduce((a,r)=>a+saldo(r),0);         // represado = saldo
  const pctFat     = totalValor>0 ? Math.round((valorFat/totalValor)*100) : 0;
  const faturados  = f.filter(isFat);

  const byEtapa = {};
  STATUS_ORDER.forEach(s=>{ byEtapa[s]={ count:0, valor:0 }; });
  f.forEach(r=>{ const s=recStatus(r, fatByRec[r.id], billOf(r)); if(byEtapa[s]){byEtapa[s].count++;byEtapa[s].valor+=billOf(r);} });

  const byAnalista = {};
  f.forEach(r=>{
    if(!byAnalista[r.responsavel]) byAnalista[r.responsavel]={ total:0, fat:0, rep:0, cnt:0, fatCnt:0 };
    byAnalista[r.responsavel].total+=billOf(r); byAnalista[r.responsavel].cnt++;
    byAnalista[r.responsavel].fat+=fat(r); byAnalista[r.responsavel].rep+=saldo(r);
    if(isFat(r)) byAnalista[r.responsavel].fatCnt++;
  });

  const byEmpresa = {};
  f.forEach(r=>{ if(!byEmpresa[r.empresa])byEmpresa[r.empresa]={total:0,fat:0}; byEmpresa[r.empresa].total+=billOf(r); byEmpresa[r.empresa].fat+=fat(r); });

  const byTipo = {};
  f.forEach(r=>{ const t=r.tipo||"—"; if(!byTipo[t])byTipo[t]={total:0,fat:0,cnt:0}; byTipo[t].total+=billOf(r); byTipo[t].cnt++; byTipo[t].fat+=fat(r); });

  const naoFatByCliente = {};
  naoFat.forEach(r=>{
    const key=r.cliente+"|"+pepBase(r.pep);
    if(!naoFatByCliente[key]) naoFatByCliente[key]={ cliente:r.cliente, pep:pepBase(r.pep), responsavel:r.responsavel, count:0, valor:0, status:recStatus(r, fatByRec[r.id], billOf(r)), color:recStatusColor(r, fatByRec[r.id], billOf(r)) };
    naoFatByCliente[key].count++; naoFatByCliente[key].valor+=saldo(r);
  });

  const etapaColors = { "Faturado":"green","Liberado para faturamento":"teal","Cliente aprovou":"blue","Aguard. aprovação cliente":"yellow","Retorno comercial recebido":"yellow","Aguard. retorno comercial":"orange","Racional montado":"orange","Dados extraídos":"orange","Não iniciado":"gray" };
  const MetCard=({label,value,color,sub,highlight})=>(
    <Card style={{padding:"14px 16px", ...(highlight?{borderColor:C.orange.border,background:"#fffaf3"}:{})}}>
      <div style={{fontSize:11,color:T.muted,marginBottom:4,fontWeight:600,textTransform:"uppercase",letterSpacing:".3px"}}>{label}</div>
      <div style={{fontSize:20,fontWeight:800,color:color||T.ink}}>{value}</div>
      {sub&&<div style={{fontSize:11,color:T.muted,marginTop:2}}>{sub}</div>}
    </Card>
  );

  return (
    <div>
      <PageHead icon="chart" title="Dashboard" sub="Visão geral do reconhecimento e do faturamento"/>

      {/* Filtros */}
      <Card style={{ padding:"12px 14px", marginBottom:18 }}>
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          <select style={{...inp,width:"auto",flex:"1 1 150px"}} value={filterComp} onChange={e=>setFC(e.target.value)} aria-label="Competência"><option value="todas">Todas as competências</option>{comps.map(c=><option key={c}>{c}</option>)}</select>
          <select style={{...inp,width:"auto",flex:"1 1 150px"}} value={filterEmpresa} onChange={e=>setFE(e.target.value)} aria-label="Empresa"><option value="todas">Todas as empresas</option>{EMPRESAS.map(e=><option key={e.cod} value={e.cod}>{e.cod} — {e.nome}</option>)}</select>
          <select style={{...inp,width:"auto",flex:"1 1 150px"}} value={filterTipo} onChange={e=>setFTi(e.target.value)} aria-label="Tipo de contrato"><option value="todas">Todos os contratos</option>{tipos.map(t=><option key={t}>{t}</option>)}</select>
          {isAdmin&&<select style={{...inp,width:"auto",flex:"1 1 150px"}} value={filterAnalista} onChange={e=>setFA(e.target.value)} aria-label="Analista"><option value="todos">Todos os analistas</option>{analistas.map(a=><option key={a}>{a}</option>)}</select>}
          <select style={{...inp,width:"auto",flex:"1 1 150px"}} value={filterEtapa} onChange={e=>setFEt(e.target.value)} aria-label="Etapa"><option value="todas">Todas as etapas</option>{STATUS_ORDER.map(s=><option key={s}>{s}</option>)}</select>
        </div>
      </Card>

      {/* Herói: donut + KPIs */}
      <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fit,minmax(260px,1fr))", gap:14, marginBottom:16, alignItems:"stretch" }}>
        <Card style={{ padding:"18px 20px", display:"flex", alignItems:"center" }}>
          <Donut pct={pctFat} label="faturado" sub={
            <div>
              <div style={{ fontSize:12, color:T.muted }}>do valor do período</div>
              <div style={{ fontSize:15, fontWeight:800, color:T.ink, margin:"2px 0 8px" }}>{fmtShort(totalValor)} <span style={{fontWeight:500,fontSize:11,color:T.muted}}>total</span></div>
              <div style={{ display:"flex", alignItems:"center", gap:6, fontSize:12 }}><i style={{width:9,height:9,borderRadius:2,background:T.ok}}/> Faturado {fmtShort(valorFat)}</div>
              <div style={{ display:"flex", alignItems:"center", gap:6, fontSize:12, marginTop:3 }}><i style={{width:9,height:9,borderRadius:2,background:"#fca5a5"}}/> Represado {fmtShort(valorRep)}</div>
            </div>
          }/>
        </Card>
        <div style={{ display:"grid", gridTemplateColumns:"1fr 1fr", gap:14 }}>
          <MetCard label="Registros" value={f.length}/>
          <MetCard label="Faturado" value={fmtShort(valorFat)} color={T.ok} sub={`${faturados.length} registros`}/>
          <MetCard label="Represado" value={fmtShort(valorRep)} color={C.orange.solid} sub={`${naoFat.length} registros · atenção`} highlight/>
          <MetCard label="Valor total" value={fmtShort(totalValor)} color={T.brand}/>
        </div>
      </div>

      {/* Valores por etapa */}
      <Card style={{padding:16,marginBottom:16}}>
        <SectionTitle>Valor por etapa do funil</SectionTitle>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(180px,1fr))",gap:8}}>
          {STATUS_ORDER.filter(s=>byEtapa[s]?.count>0).map(s=>{
            const d=byEtapa[s]; const c=C[etapaColors[s]]||C.gray;
            return <div key={s} style={{padding:"10px 12px",borderRadius:T.rMd,background:c.bg,border:`1px solid ${c.border}`}}>
              <div style={{fontSize:11,color:c.text,fontWeight:700,marginBottom:4}}>{s}</div>
              <div style={{fontSize:16,fontWeight:800,color:c.text}}>{fmtShort(d.valor)}</div>
              <div style={{fontSize:11,color:c.text,opacity:.75}}>{d.count} registro(s)</div>
            </div>;
          })}
        </div>
      </Card>

      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(280px,1fr))",gap:16,marginBottom:16}}>
        {isAdmin&&<Card style={{padding:16}}>
          <SectionTitle>Por analista</SectionTitle>
          {Object.entries(byAnalista).map(([a,d])=>{
            const pct=d.total>0?Math.round((d.fat/d.total)*100):0;
            const bar=pct===100?T.ok:pct>50?T.brand:C.orange.solid;
            return <div key={a} style={{marginBottom:14}}>
              <div style={{display:"flex",justifyContent:"space-between",fontSize:12,marginBottom:3}}><span style={{fontWeight:600,color:T.ink,display:"flex",alignItems:"center",gap:7}}><Avatar name={a} size={22}/>{a}</span><span style={{color:T.muted}}>{pct}% · {d.cnt} reg.</span></div>
              <div style={{height:7,background:T.lineSoft,borderRadius:4}}><div style={{height:7,borderRadius:4,width:`${pct}%`,background:bar,transition:"width .4s"}}/></div>
              <div style={{display:"flex",justifyContent:"space-between",fontSize:11,color:T.muted,marginTop:3}}><span>Fat: {fmtShort(d.fat)}</span><span>Rep: {fmtShort(d.rep)}</span></div>
            </div>;
          })}
        </Card>}

        <Card style={{padding:16}}>
          <SectionTitle>Por empresa</SectionTitle>
          {Object.entries(byEmpresa).map(([cod,d])=>{
            const emp=EMPRESAS.find(e=>e.cod===cod);
            const pct=d.total>0?Math.round((d.fat/d.total)*100):0;
            return <div key={cod} style={{marginBottom:12}}>
              <div style={{display:"flex",justifyContent:"space-between",fontSize:12,marginBottom:3}}><span style={{fontWeight:600,color:T.ink}}>{cod} — {emp?.nome}</span><span style={{color:T.muted}}>{fmtShort(d.total)} · {pct}%</span></div>
              <div style={{height:7,background:T.lineSoft,borderRadius:4}}><div style={{height:7,borderRadius:4,width:`${pct}%`,background:T.brand,transition:"width .4s"}}/></div>
            </div>;
          })}
        </Card>

        <Card style={{padding:16}}>
          <SectionTitle>Por tipo de contrato</SectionTitle>
          {Object.keys(byTipo).length===0 && <div style={{fontSize:13,color:T.muted}}>Sem dados.</div>}
          {Object.entries(byTipo).sort((a,b)=>b[1].total-a[1].total).map(([t,d])=>{
            const pct=d.total>0?Math.round((d.fat/d.total)*100):0;
            const c=C[({ "Time & Expenses":"blue","Fee":"purple","WIP":"teal","Usage Based":"orange" })[t]||"gray"]||C.gray;
            return <div key={t} style={{marginBottom:12}}>
              <div style={{display:"flex",justifyContent:"space-between",fontSize:12,marginBottom:3}}><span style={{fontWeight:600,color:T.ink}}>{t} <span style={{fontWeight:400,color:T.muted}}>· {d.cnt} reg.</span></span><span style={{color:T.muted}}>{fmtShort(d.total)} · {pct}%</span></div>
              <div style={{height:7,background:T.lineSoft,borderRadius:4}}><div style={{height:7,borderRadius:4,width:`${pct}%`,background:c.solid,transition:"width .4s"}}/></div>
            </div>;
          })}
        </Card>
      </div>

      {/* Não faturados por cliente */}
      <Card style={{padding:16}}>
        <SectionTitle count={Object.keys(naoFatByCliente).length}>Não faturados — resumo por cliente</SectionTitle>
        {Object.keys(naoFatByCliente).length===0
          ?<div style={{textAlign:"center",padding:"1rem",color:T.muted,fontSize:13}}>Tudo faturado!</div>
          :<div className="fc-scroll" style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
            <thead><tr style={{background:T.canvas}}>{[isAdmin&&"Analista","Cliente","PEP","Profissionais","Val. Total","Etapa atual"].filter(Boolean).map(h=><th key={h} style={{padding:"8px 10px",textAlign:"left",borderBottom:`1px solid ${T.line}`,fontWeight:600,color:T.muted,whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
            <tbody>{Object.values(naoFatByCliente).sort((a,b)=>b.valor-a.valor).map((d,i)=><tr key={i} className="fc-row" style={{borderBottom:`1px solid ${T.lineSoft}`}}>
              {isAdmin&&<td style={{padding:"7px 10px"}}>{d.responsavel}</td>}
              <td style={{padding:"7px 10px",fontWeight:500,color:T.ink}}>{d.cliente}</td>
              <td style={{padding:"7px 10px",fontFamily:"monospace",fontSize:11}}>{d.pep}</td>
              <td style={{padding:"7px 10px",textAlign:"center"}}>{d.count}</td>
              <td style={{padding:"7px 10px",fontWeight:700}}>{fmtShort(d.valor)}</td>
              <td style={{padding:"7px 10px"}}><Badge label={d.status} color={d.color} small dot/></td>
            </tr>)}</tbody>
          </table></div>
        }
      </Card>
    </div>
  );
}

// ─── SIDEBAR / NAV ───────────────────────────────────────────────────────────

const NAV_SECTIONS = [
  { group:"", links:[ {id:"home",icon:"home",label:"Início"} ] },
  { group:"Reconhecimento & Faturamento Receita", links:[ {id:"time",icon:"list",label:"Minha visão"}, {id:"dash",icon:"chart",label:"Dashboard"}, {id:"concil",icon:"receipt",label:"Conciliação de notas"}, {id:"reports",icon:"file",label:"Relatórios"}, {id:"projeto",icon:"chart",label:"Visão por projeto"}, {id:"represados",icon:"alert",label:"Represados"}, {id:"valida",icon:"check",label:"Validações"} ] },
  { group:"Cadastros", links:[ {id:"clients",icon:"building",label:"Clientes"} ] },
  { group:"Operação",    links:[ {id:"tasks",icon:"task",label:"Tarefas"} ] },
];

const ADMIN_NAV_SECTION = { group:"Administração", links:[ {id:"dados",icon:"import",label:"Importar documentos"}, {id:"correcoes",icon:"pencil",label:"Correções"}, {id:"bu",icon:"building",label:"Classificar BU"}, {id:"aliases",icon:"building",label:"Unificar clientes"}, {id:"comercial",icon:"chart",label:"Visão comercial"}, {id:"report",icon:"file",label:"Report semanal (comercial)"}, {id:"previsao",icon:"chart",label:"Previsão & Saúde"}, {id:"access",icon:"lock",label:"Gestão de acessos"} ] };

// Navegação do acesso COMERCIAL — enxuta: só a receita da BU dele.
const COMERCIAL_NAV_SECTIONS = [
  { group:"", links:[ {id:"dash",icon:"chart",label:"Dashboard"} ] },
  { group:"Minha BU", links:[ {id:"projeto",icon:"chart",label:"Visão por projeto"}, {id:"represados",icon:"alert",label:"Represados"}, {id:"report",icon:"file",label:"Report semanal"} ] },
];
// Páginas que o comercial pode abrir (trava de UI; o RLS trava o dado).
const COMERCIAL_PAGES = new Set(["dash","projeto","represados","report"]);

function NavLinks({ page, setPage, isAdmin, isComercial, onNavigate }) {
  const sections = isComercial ? COMERCIAL_NAV_SECTIONS : isAdmin ? [...NAV_SECTIONS, ADMIN_NAV_SECTION] : NAV_SECTIONS;
  return (
    <>
      {sections.map(sec=>(
        <div key={sec.group||"__home"} style={{marginBottom:18}}>
          {sec.group && <div style={{fontSize:11,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:".5px",padding:"0 8px",marginBottom:6}}>{sec.group}</div>}
          {sec.links.map(l=>{
            const active = page===l.id;
            return (
              <button key={l.id} onClick={()=>{setPage(l.id);onNavigate?.();}} aria-current={active?"page":undefined} style={{
                width:"100%",display:"flex",alignItems:"center",gap:10,padding:"9px 10px",marginBottom:2,
                border:"none",borderRadius:T.rMd,cursor:"pointer",textAlign:"left",fontSize:13,
                fontWeight:active?700:500,
                background:active?T.brandBg:"transparent",
                color:active?T.brand:T.inkSoft,
                borderLeft:active?`3px solid ${T.brand}`:"3px solid transparent",
              }}>
                <Icon name={l.icon} size={17}/>{l.label}
              </button>
            );
          })}
        </div>
      ))}
    </>
  );
}

function UserChip({ user, isAdmin, isComercial }) {
  const papel = isComercial ? `Comercial${user.bu?` · ${user.bu}`:""}` : isAdmin ? "Administrador" : "Analista";
  return (
    <div style={{ display:"flex", alignItems:"center", gap:9, padding:"4px 8px 14px", marginBottom:6, borderBottom:`1px solid ${T.lineSoft}` }}>
      <Avatar name={user.name} admin={isAdmin}/>
      <div style={{ minWidth:0 }}>
        <div style={{ fontSize:13, fontWeight:700, color:T.ink, whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis" }}>{user.name}</div>
        <div style={{ fontSize:11, color:(isAdmin||isComercial)?T.brand:T.muted, fontWeight:600 }}>{papel}</div>
      </div>
    </div>
  );
}

function Sidebar({ page, setPage, user, isAdmin, isComercial }) {
  return (
    <aside style={{width:212,flexShrink:0,background:"var(--surface)",borderRight:`1px solid ${T.line}`,padding:"18px 12px",display:"flex",flexDirection:"column"}}>
      <UserChip user={user} isAdmin={isAdmin} isComercial={isComercial}/>
      <NavLinks page={page} setPage={setPage} isAdmin={isAdmin} isComercial={isComercial}/>
    </aside>
  );
}

function MobileDrawer({ open, onClose, page, setPage, user, isAdmin, isComercial }) {
  if (!open) return null;
  return (
    <div style={{ position:"fixed", inset:0, zIndex:250 }}>
      <div style={{ position:"absolute", inset:0, background:"rgba(15,23,42,.5)", animation:"fcOverlay .15s ease" }} onClick={onClose}/>
      <aside style={{ position:"absolute", top:0, left:0, bottom:0, width:240, background:"var(--surface)", padding:"18px 12px", boxShadow:T.shLg, display:"flex", flexDirection:"column", overflowY:"auto" }}>
        <UserChip user={user} isAdmin={isAdmin} isComercial={isComercial}/>
        <NavLinks page={page} setPage={setPage} isAdmin={isAdmin} isComercial={isComercial} onNavigate={onClose}/>
      </aside>
    </div>
  );
}

// ─── TASK MODAL (criar / editar) ─────────────────────────────────────────────

function TaskModal({ task, responsaveis, onSave, onDelete, onClose }) {
  const isNew = !task.id;
  const [title,setTitle]       = useState(task.title || "");
  const [desc,setDesc]         = useState(task.desc || "");
  const [dueDate,setDueDate]   = useState(task.dueDate || "");
  const [assignee,setAssignee] = useState(task.assignee || "");
  const [status,setStatus]     = useState(task.status || "inbox");
  const [err,setErr]           = useState("");

  function save() {
    if (!title.trim()) { setErr("Informe um título para a tarefa."); return; }
    const now = nowISO();
    if (isNew) onSave({ id:genId(), title:title.trim(), desc:desc.trim(), dueDate, assignee, status, createdAt:now, updatedAt:now });
    else       onSave({ ...task, title:title.trim(), desc:desc.trim(), dueDate, assignee, status, updatedAt:now });
    onClose();
  }

  return (
    <Modal title={isNew?"Nova tarefa":"Editar tarefa"} onClose={onClose}>
      <div style={{marginBottom:14}}><Field label="Título *"><input style={inp} placeholder="Ex: Extrair base de T&E de junho" value={title} onChange={e=>{setTitle(e.target.value);setErr("");}} autoFocus/></Field></div>
      <div style={{marginBottom:14}}><Field label="Descrição"><textarea style={{...inp,minHeight:70,resize:"vertical"}} placeholder="Breve descrição da tarefa..." value={desc} onChange={e=>setDesc(e.target.value)}/></Field></div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:14}}>
        <Field label="Data de entrega"><input type="date" style={inp} value={dueDate} onChange={e=>setDueDate(e.target.value)}/></Field>
        <Field label="Responsável"><select style={inp} value={assignee} onChange={e=>setAssignee(e.target.value)}><option value="">Não atribuído</option>{responsaveis.map(r=><option key={r}>{r}</option>)}</select></Field>
      </div>
      <div style={{marginBottom:18}}><Field label="Coluna"><select style={inp} value={status} onChange={e=>setStatus(e.target.value)}>{TASK_COLUMNS.map(c=><option key={c.id} value={c.id}>{c.title}</option>)}</select></Field></div>
      {err&&<div style={{marginBottom:12,fontSize:12,padding:"8px 12px",borderRadius:T.rMd,background:T.dangerBg,color:T.danger,border:`1px solid ${T.dangerLine}`}}>{err}</div>}
      <div style={{display:"flex",gap:8,justifyContent:"space-between",alignItems:"center"}}>
        <div>{!isNew&&<Btn danger small onClick={()=>{onDelete(task.id);onClose();}}>Excluir</Btn>}</div>
        <div style={{display:"flex",gap:8}}>
          <Btn onClick={onClose}>Cancelar</Btn>
          <Btn primary onClick={save}>{isNew?"Criar tarefa":"Salvar"}</Btn>
        </div>
      </div>
    </Modal>
  );
}

// ─── TASK CARD ───────────────────────────────────────────────────────────────

function dueInfo(dueDate, status) {
  if (!dueDate) return null;
  const [y,m,d] = dueDate.split("-");
  const label = `${d}/${m}`;
  const today = new Date(); today.setHours(0,0,0,0);
  const due = new Date(+y, +m-1, +d);
  const diff = Math.round((due - today) / 86400000);
  if (status==="done") return { label, color:"gray" };
  if (diff < 0)  return { label:`${label} · atrasada`, color:"red" };
  if (diff === 0) return { label:`${label} · hoje`,    color:"orange" };
  if (diff <= 2) return { label:`${label} · ${diff}d`, color:"yellow" };
  return { label, color:"gray" };
}

function TaskCard({ task, onOpen, onMove, onDragStart, onDragEnd }) {
  const idx = TASK_COLUMNS.findIndex(c=>c.id===task.status);
  const di = dueInfo(task.dueDate, task.status);
  return (
    <div
      draggable
      onDragStart={()=>onDragStart(task.id)}
      onDragEnd={onDragEnd}
      onClick={()=>onOpen(task)}
      className="fc-card-int"
      style={{background:"var(--surface)",border:`1px solid ${T.line}`,borderRadius:T.rLg,padding:"10px 12px",marginBottom:8,cursor:"pointer",boxShadow:T.shSm}}
    >
      <div style={{fontSize:13,fontWeight:600,color:T.ink,marginBottom:task.desc?4:8,lineHeight:1.35}}>{task.title}</div>
      {task.desc&&<div style={{fontSize:11.5,color:T.muted,marginBottom:8,lineHeight:1.4,display:"-webkit-box",WebkitLineClamp:2,WebkitBoxOrient:"vertical",overflow:"hidden"}}>{task.desc}</div>}
      <div style={{display:"flex",alignItems:"center",gap:6,flexWrap:"wrap"}}>
        {task.recorrente&&<Badge label="Recorrente" color="teal" small/>}
        {di&&<Badge label={""+di.label} color={di.color} small/>}
        <Badge label={task.assignee||"Não atribuído"} color={task.assignee?"purple":"gray"} small/>
        <div style={{flex:1}}/>
        <button title="Mover para a esquerda" aria-label="Mover para a esquerda" disabled={idx<=0} onClick={e=>{e.stopPropagation();onMove(task,TASK_COLUMNS[idx-1].id);}}
          style={{border:"none",background:"none",cursor:idx<=0?"default":"pointer",color:idx<=0?T.line:T.muted,fontSize:14,padding:"0 2px"}}>◀</button>
        <button title="Mover para a direita" aria-label="Mover para a direita" disabled={idx>=TASK_COLUMNS.length-1} onClick={e=>{e.stopPropagation();onMove(task,TASK_COLUMNS[idx+1].id);}}
          style={{border:"none",background:"none",cursor:idx>=TASK_COLUMNS.length-1?"default":"pointer",color:idx>=TASK_COLUMNS.length-1?T.line:T.muted,fontSize:14,padding:"0 2px"}}>▶</button>
      </div>
    </div>
  );
}

// ─── ENTREGAS (modelos recorrentes) ──────────────────────────────────────────

function DeliveryTemplateModal({ template, responsaveis, onSave, onDelete, onClose }) {
  const isNew = !template?.id;
  const [title, setTitle] = useState(template?.title || "");
  const [items, setItems] = useState(() => {
    const src = template?.items?.length ? template.items : [{title:"",desc:"",dia:"",assignees:[]}];
    return src.map(it => ({ title:it.title||"", desc:it.desc||"", dia:it.dia||"", assignees: Array.isArray(it.assignees)?it.assignees:(it.assignee?[it.assignee]:[]) }));
  });
  const [err, setErr] = useState("");
  const setItem = (i,k,v) => setItems(a=>a.map((it,j)=>j===i?{...it,[k]:v}:it));
  const addItem = () => setItems(a=>[...a,{title:"",desc:"",dia:"",assignees:[]}]);
  const delItem = (i) => setItems(a=>a.filter((_,j)=>j!==i));
  const toggleAssignee = (i,r) => setItems(a=>a.map((it,j)=>{ if(j!==i) return it; const has=it.assignees.includes(r); return {...it, assignees: has?it.assignees.filter(x=>x!==r):[...it.assignees,r]}; }));
  function save() {
    if (!title.trim()) { setErr("Informe o nome da entrega."); return; }
    const clean = items.filter(it=>(it.title||"").trim()).map(it=>({title:it.title.trim(),desc:(it.desc||"").trim(),dia:it.dia?Number(it.dia):"",assignees:it.assignees}));
    if (!clean.length) { setErr("Inclua ao menos uma tarefa na entrega."); return; }
    onSave({ ...(template||{}), title:title.trim(), items:clean });
    onClose();
  }
  return (
    <Modal title={isNew?"Novo modelo de entrega":`Editar — ${template.title}`} subtitle="Recorrente: as tarefas serão geradas a cada mês para os analistas" onClose={onClose} wide>
      <div style={{marginBottom:16}}><Field label="Nome da entrega *"><input style={inp} placeholder="Ex: Fechamento mensal de faturamento" value={title} onChange={e=>{setTitle(e.target.value);setErr("");}}/></Field></div>
      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:8}}>
        <div style={{fontSize:12,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:".4px"}}>Tarefas da entrega</div>
        <Btn small onClick={addItem}>+ Tarefa</Btn>
      </div>
      {items.map((it,i)=>(
        <div key={i} style={{border:`1px solid ${T.line}`,borderRadius:T.rLg,padding:"12px 14px",marginBottom:8,background:"var(--surface)"}}>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:8}}>
            <span style={{fontSize:12,fontWeight:700,color:T.brand}}>Tarefa {i+1}</span>
            {items.length>1 && <button onClick={()=>delItem(i)} title="Remover" style={{border:"none",background:"none",cursor:"pointer",color:T.danger,fontSize:12,fontWeight:600}}>✕ Remover</button>}
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 130px",gap:10,marginBottom:8}}>
            <Field label="Título"><input style={inp} placeholder="Ex: Extrair base de T&E" value={it.title} onChange={e=>setItem(i,"title",e.target.value)}/></Field>
            <Field label="Dia do mês" hint="(vira a data)"><input style={inp} inputMode="numeric" placeholder="Ex: 25" value={it.dia} onChange={e=>setItem(i,"dia", e.target.value.replace(/\D/g,"").slice(0,2))}/></Field>
          </div>
          <Field label="Descrição (opcional)"><input style={inp} placeholder="Detalhe da tarefa..." value={it.desc} onChange={e=>setItem(i,"desc",e.target.value)}/></Field>
          <div style={{marginTop:10}}>
            <label style={{...Ty.label}}>Enviar para quais analistas? <span style={{fontWeight:500,color:T.muted}}>(nenhum marcado = todos)</span></label>
            <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
              {responsaveis.map(r=>{ const on=it.assignees.includes(r); return (
                <label key={r} style={{display:"flex",alignItems:"center",gap:6,padding:"5px 10px",borderRadius:T.rPill,border:`1.5px solid ${on?T.brand:T.line}`,background:on?T.brandBg:"var(--surface)",cursor:"pointer",fontSize:12,fontWeight:on?600:400,color:on?T.brand:T.inkSoft}}>
                  <input type="checkbox" checked={on} onChange={()=>toggleAssignee(i,r)} style={{width:13,height:13}}/>{r}
                </label>
              );})}
              {responsaveis.length===0 && <span style={{fontSize:12,color:T.muted}}>Nenhum analista cadastrado ainda.</span>}
            </div>
          </div>
        </div>
      ))}
      {err&&<div style={{margin:"12px 0",fontSize:12,padding:"8px 12px",borderRadius:T.rMd,background:T.dangerBg,color:T.danger,border:`1px solid ${T.dangerLine}`}}>{err}</div>}
      <div style={{display:"flex",gap:8,justifyContent:"space-between",alignItems:"center",marginTop:14}}>
        <div>{!isNew && <Btn danger small onClick={()=>{onDelete(template.id);onClose();}}>Excluir modelo</Btn>}</div>
        <div style={{display:"flex",gap:8}}><Btn onClick={onClose}>Cancelar</Btn><Btn primary onClick={save}>{isNew?"Criar modelo":"Salvar"}</Btn></div>
      </div>
    </Modal>
  );
}

function DeliveryManager({ templates, responsaveis, competenciaAtual, onTemplateSave, onTemplateDelete, onGenerate, onClose }) {
  const [editing, setEditing] = useState(null);
  const [comp, setComp] = useState(competenciaAtual || "");
  const [confirmGen, setConfirmGen] = useState(null);
  return (
    <Modal title="Entregas recorrentes" subtitle="Modelos de entrega e geração mensal das tarefas" onClose={onClose} wide>
      {editing && <DeliveryTemplateModal template={editing.id?editing:null} responsaveis={responsaveis} onSave={onTemplateSave} onDelete={onTemplateDelete} onClose={()=>setEditing(null)}/>}
      {confirmGen && <ConfirmDialog title="Gerar entrega do mês" confirmLabel="Gerar"
        message={`Gerar as tarefas de "${confirmGen.title}" para a competência ${comp}? Serão criadas tarefas para os analistas.`}
        onConfirm={()=>onGenerate(confirmGen, comp)} onClose={()=>setConfirmGen(null)}/>}

      <div style={{display:"flex",alignItems:"center",gap:10,flexWrap:"wrap",marginBottom:16,padding:"12px 14px",background:T.canvas,borderRadius:T.rLg,border:`1px solid ${T.line}`}}>
        <span style={{fontSize:13,fontWeight:600,color:T.ink}}>Gerar para a competência:</span>
        <input style={{...inp,width:120}} placeholder="MM/AAAA" value={comp} onChange={e=>setComp(e.target.value)}/>
        <span style={{fontSize:11,color:T.muted}}>Use o botão "Gerar do mês" em cada modelo abaixo.</span>
      </div>

      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:10}}>
        <span style={{fontSize:13,fontWeight:700,color:T.ink}}>Modelos de entrega ({templates.length})</span>
        <Btn primary small onClick={()=>setEditing({})}>+ Novo modelo</Btn>
      </div>

      {templates.length===0
        ? <div style={{fontSize:13,color:T.muted,textAlign:"center",padding:"24px",background:T.canvas,borderRadius:T.rLg,border:`1px dashed ${T.line}`}}>Nenhum modelo ainda. Crie um modelo de entrega (ex.: "Fechamento mensal") com as tarefas que se repetem todo mês.</div>
        : templates.map(t=>(
          <div key={t.id} style={{border:`1px solid ${T.line}`,borderRadius:T.rLg,padding:"12px 14px",marginBottom:8,background:"var(--surface)",display:"flex",alignItems:"center",gap:12,flexWrap:"wrap"}}>
            <div style={{flex:1,minWidth:160}}>
              <div style={{fontSize:14,fontWeight:700,color:T.ink}}>{t.title}</div>
              <div style={{fontSize:11,color:T.muted,marginTop:2}}>{t.items.length} tarefa(s): {t.items.map(i=>i.title).join(", ").slice(0,80)}{t.items.map(i=>i.title).join(", ").length>80?"…":""}</div>
            </div>
            <Btn small onClick={()=>setEditing(t)}>Editar</Btn>
            <Btn primary small onClick={()=>{ if(!/^\d{2}\/\d{4}$/.test(comp)){alert("Informe a competência no formato MM/AAAA");return;} setConfirmGen(t); }}>Gerar do mês</Btn>
          </div>
        ))}

      <div style={{display:"flex",justifyContent:"flex-end",marginTop:16}}><Btn onClick={onClose}>Fechar</Btn></div>
    </Modal>
  );
}

// ─── KANBAN ──────────────────────────────────────────────────────────────────

function Kanban({ tasks, responsaveis, isAdmin, isViewer=false, competenciaAtual, templates, deliveries, onAdd, onUpdate, onDelete, onTemplateSave, onTemplateDelete, onGenerate }) {
  const isMobile = useIsMobile();
  const [editing, setEditing]       = useState(null);
  const [dragId, setDragId]         = useState(null);
  const [dragOverCol, setDragOver]  = useState(null);
  const [filterResp, setFilterResp] = useState("todos");
  const [filterTipo, setFilterTipo] = useState("todas"); // todas | recorrentes | ordinarias
  const [byAnalyst, setByAnalyst]   = useState(false);
  const [showDeliv, setShowDeliv]   = useState(false);

  let visible = filterResp==="todos" ? tasks : tasks.filter(t=>t.assignee===filterResp);
  if (filterTipo==="recorrentes") visible = visible.filter(t=>t.recorrente);
  if (filterTipo==="ordinarias")  visible = visible.filter(t=>!t.recorrente);

  const moveTo = (task, status) => { if (task.status!==status) onUpdate({ ...task, status, updatedAt:nowISO() }); };
  const onDropCol = (status) => { if (dragId) { const t=tasks.find(x=>x.id===dragId); if (t) moveTo(t,status); } setDragId(null); setDragOver(null); };
  const saveTask = (t) => { if (t.id && tasks.some(x=>x.id===t.id)) onUpdate(t); else onAdd(t); };

  function Column({ col, list }) {
    const colTasks = list.filter(t=>t.status===col.id);
    const isOver = dragOverCol===col.id;
    return (
      <div
        onDragOver={e=>{e.preventDefault();setDragOver(col.id);}}
        onDragLeave={()=>setDragOver(o=>o===col.id?null:o)}
        onDrop={()=>onDropCol(col.id)}
        style={{background:isOver?T.brandBg:"#f3f4f6",border:`1px solid ${isOver?col.accent:T.line}`,borderRadius:T.rXl,padding:"10px 10px 4px",minHeight:100,transition:"background .12s"}}>
        <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:10,padding:"2px 4px"}}>
          <span style={{width:9,height:9,borderRadius:"50%",background:col.accent}}/>
          <span style={{fontSize:13,fontWeight:700,color:col.color}}>{col.title}</span>
          <span style={{fontSize:11,color:T.muted,background:"var(--surface)",borderRadius:T.rPill,padding:"1px 8px",border:`1px solid ${T.line}`}}>{colTasks.length}</span>
        </div>
        {col.id==="inbox"&&<button onClick={()=>setEditing({ status:"inbox" })}
          style={{width:"100%",marginBottom:10,padding:"8px",border:`1.5px dashed #c7cdd6`,borderRadius:T.rMd,background:"var(--surface)",color:T.muted,fontSize:12,fontWeight:600,cursor:"pointer"}}>
          + Criar tarefa aqui
        </button>}
        {colTasks.map(t=>(
          <TaskCard key={t.id} task={t} onOpen={setEditing} onMove={moveTo} onDragStart={setDragId} onDragEnd={()=>{setDragId(null);setDragOver(null);}}/>
        ))}
        {colTasks.length===0&&col.id!=="inbox"&&<div style={{textAlign:"center",color:"#c7cdd6",fontSize:11.5,padding:"14px 4px"}}>Solte tarefas aqui</div>}
        {colTasks.length===0&&col.id==="inbox"&&<div style={{textAlign:"center",color:"#c7cdd6",fontSize:11.5,padding:"4px"}}>{col.hint}</div>}
      </div>
    );
  }
  const boardStyle = {display:"grid",gridTemplateColumns:`repeat(${TASK_COLUMNS.length},minmax(240px,1fr))`,gap:14,alignItems:"start",overflowX:"auto",paddingBottom:6};

  // Agrupamento por analista (visão do admin)
  const groups = byAnalyst
    ? [...responsaveis.map(r=>({key:r,label:r,list:visible.filter(t=>t.assignee===r)})), {key:"__none__",label:"Sem responsável",list:visible.filter(t=>!t.assignee)}].filter(g=>g.list.length>0)
    : null;

  return (
    <div>
      {editing && <TaskModal task={editing} responsaveis={responsaveis} onSave={saveTask} onDelete={onDelete} onClose={()=>setEditing(null)}/>}
      {showDeliv && <DeliveryManager templates={templates} responsaveis={responsaveis} competenciaAtual={competenciaAtual}
        onTemplateSave={onTemplateSave} onTemplateDelete={onTemplateDelete} onGenerate={onGenerate} onClose={()=>setShowDeliv(false)}/>}

      <div style={{display:"flex",alignItems:"center",gap:13,marginBottom:18,flexWrap:"wrap"}}>
        <HeadChip icon="task"/>
        <div style={{flex:1,minWidth:180}}>
          <h1 style={{...Ty.h1,fontSize:22}}>Tarefas do time</h1>
          <div style={{...Ty.small, marginTop:3}}>{tasks.length} tarefa(s) · arraste os cards entre as colunas ou use as setas</div>
        </div>
        <select style={{...inp,width:"auto"}} value={filterTipo} onChange={e=>setFilterTipo(e.target.value)} aria-label="Tipo de tarefa">
          <option value="todas">Todas</option>
          <option value="recorrentes">Recorrentes</option>
          <option value="ordinarias">Ordinárias</option>
        </select>
        <select style={{...inp,width:"auto"}} value={filterResp} onChange={e=>setFilterResp(e.target.value)} aria-label="Filtrar responsável">
          <option value="todos">Todos os responsáveis</option>
          {responsaveis.map(r=><option key={r}>{r}</option>)}
        </select>
        {isAdmin&&<Btn small onClick={()=>setByAnalyst(v=>!v)} style={byAnalyst?{borderColor:T.brand,color:T.brand}:{}}>Por analista</Btn>}
        {isAdmin&&<Btn small onClick={()=>setShowDeliv(true)}>Entregas</Btn>}
        {!isViewer && <Btn primary onClick={()=>setEditing({ status:"inbox" })}>+ Nova tarefa</Btn>}
      </div>

      {!byAnalyst && <div className="fc-scroll" style={boardStyle}>
        {TASK_COLUMNS.map(col=><Column key={col.id} col={col} list={visible}/>)}
      </div>}

      {byAnalyst && (groups.length===0
        ? <Card style={{textAlign:"center",padding:"2rem",color:T.muted,fontSize:13}}>Nenhuma tarefa para exibir.</Card>
        : groups.map(g=>(
            <div key={g.key} style={{marginBottom:22}}>
              <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8}}>
                {g.key!=="__none__" && <Avatar name={g.label} size={24}/>}
                <span style={{fontSize:14,fontWeight:700,color:T.ink}}>{g.label}</span>
                <span style={{fontSize:11,color:T.muted,background:T.lineSoft,borderRadius:T.rPill,padding:"1px 9px"}}>{g.list.length}</span>
              </div>
              <div className="fc-scroll" style={boardStyle}>
                {TASK_COLUMNS.map(col=><Column key={col.id} col={col} list={g.list}/>)}
              </div>
            </div>
          )))}
    </div>
  );
}

// ─── GESTÃO DE ACESSOS (admin) ───────────────────────────────────────────────

function NewAccessInfoModal({ onClose }) {
  return (
    <Modal title="Adicionar novo acesso" subtitle="Como criar um login novo (analista ou admin)" onClose={onClose}>
      <p style={{...Ty.body, lineHeight:1.55, marginTop:0}}>
        Por segurança, criar um <b>login novo</b> (e-mail + senha) é feito no painel do Supabase.
        Depois que a pessoa existir, você ajusta o papel (admin/analista) e remove o acesso aqui mesmo nesta tela.
      </p>
      <ol style={{...Ty.body, lineHeight:1.7, paddingLeft:18}}>
        <li>No Supabase, abra <b>Authentication → Users → Add user</b>.</li>
        <li>Informe o <b>e-mail</b> e uma <b>senha</b> inicial e marque <b>Auto Confirm User</b>.</li>
        <li>Volte aqui e clique em <b>Atualizar lista</b>: a pessoa aparece como <b>Analista</b>.</li>
        <li>Ajuste o <b>nome</b> (para casar com o "Responsável" das planilhas) e, se precisar, marque como <b>Administrador</b>.</li>
      </ol>
      <div style={{fontSize:12,color:T.muted,marginTop:6}}>No próximo marco dá para trazer essa criação para dentro do app (função no servidor).</div>
      <div style={{display:"flex",justifyContent:"flex-end",marginTop:18}}><Btn primary onClick={onClose}>Entendi</Btn></div>
    </Modal>
  );
}

function AccessEditModal({ profile, onSave, onClose }) {
  const [name, setName]       = useState(profile.name || "");
  const [apelido, setApelido] = useState(profile.apelido || "");
  const [aniversario, setAniv]= useState(profile.aniversario || "");
  const [responsavel, setResp]= useState(profile.responsavel || "");
  const [papel, setPapel]     = useState(profile.isComercial ? "comercial" : profile.isViewer ? "viewer" : profile.isAdmin ? "admin" : "analista");
  const [bu, setBu]           = useState(profile.bu || BUS[0]);
  const [err, setErr]         = useState("");
  function save() {
    const nm = name.trim();
    if (!nm) { setErr("Informe o nome de exibição."); return; }
    if (papel==="comercial" && !bu) { setErr("Escolha a BU deste comercial."); return; }
    onSave({ id: profile.id, name: nm, isAdmin: papel==="admin", isViewer: papel==="viewer", isComercial: papel==="comercial",
      bu: papel==="comercial" ? bu : "", responsavel: responsavel.trim(), apelido: apelido.trim(), aniversario: aniversario.trim() });
    onClose();
  }
  const PAPEIS = [
    { v:"analista", l:"Analista", d:"Vê e edita as próprias receitas (vínculo abaixo)." },
    { v:"comercial",l:"Comercial (BU)", d:"Vê só a receita da sua BU — Dashboard, Visão por projeto e Report semanal. Não altera dados." },
    { v:"viewer",   l:"Somente visualização", d:"Vê todas as telas e extrai relatórios, mas não altera nenhum dado." },
    { v:"admin",    l:"Administrador", d:"Acesso completo: importar, exportar, todos os analistas e gestão de acessos." },
  ];
  return (
    <Modal title={`Editar acesso — ${profile.name}`} subtitle="Ajuste o nome, o apelido, o vínculo de receitas e o papel" onClose={onClose}>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:12,marginBottom:14}}>
        <Field label="Nome de exibição *"><input style={inp} value={name} onChange={e=>{setName(e.target.value);setErr("");}} placeholder="Ex: Fernanda"/></Field>
        <Field label="Apelido" hint="(saudação)"><input style={inp} value={apelido} onChange={e=>setApelido(e.target.value)} placeholder="Ex: Fê"/></Field>
        <Field label="Aniversário" hint="(dd/mm)"><input style={inp} value={aniversario} onChange={e=>setAniv(e.target.value)} placeholder="Ex: 07/03"/></Field>
      </div>
      <div style={{marginBottom:14}}>
        <Field label="Responsável na base de receitas"><input style={inp} value={responsavel} onChange={e=>setResp(e.target.value)} placeholder="Ex: Juliana Teles"/></Field>
        <div style={{fontSize:11,color:T.muted,marginTop:4}}>Cole aqui o nome <b>exatamente</b> como aparece na coluna "Responsável" do Excel. É isso que faz a pessoa ver as receitas dela. (Se ficar vazio, usamos o nome de exibição.)</div>
      </div>
      <div style={{marginBottom:16}}>
        <label style={Ty.label}>Papel</label>
        <div style={{display:"grid",gap:8}}>
          {PAPEIS.map(op=>{ const on=papel===op.v; return (
            <label key={op.v} style={{display:"flex",alignItems:"flex-start",gap:8,fontSize:13,color:T.inkSoft,cursor:"pointer",padding:"10px 12px",borderRadius:T.rMd,border:`1px solid ${on?T.brand:T.line}`,background:on?T.brandBg:"var(--surface)"}}>
              <input type="radio" name="papel" checked={on} onChange={()=>setPapel(op.v)} style={{width:16,height:16,marginTop:1}}/>
              <span style={{flex:1}}><b style={{color:on?T.brand:T.inkSoft}}>{op.l}</b><br/><span style={{fontSize:11,color:T.muted}}>{op.d}</span>
                {op.v==="comercial" && on && <div style={{marginTop:10}}>
                  <label style={{fontSize:11,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:".4px",display:"block",marginBottom:5}}>Unidade de negócio (BU)</label>
                  <select style={{...inp,width:"auto",minWidth:200,fontWeight:700,color:T.brand,borderColor:T.brand}} value={bu} onChange={e=>setBu(e.target.value)}>{BUS.map(b=><option key={b}>{b}</option>)}</select>
                  <div style={{fontSize:11,color:T.muted,marginTop:5}}>Este comercial enxerga apenas os projetos/receita classificados nesta BU. Classifique os clientes em <b>Administração → Classificar BU</b>.</div>
                </div>}
              </span>
            </label>
          );})}
        </div>
      </div>
      {err&&<div style={{marginBottom:12,fontSize:12,padding:"8px 12px",borderRadius:T.rMd,background:T.dangerBg,color:T.danger,border:`1px solid ${T.dangerLine}`}}>{err}</div>}
      <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
        <Btn onClick={onClose}>Cancelar</Btn>
        <Btn primary onClick={save}>Salvar</Btn>
      </div>
    </Modal>
  );
}

function AccessManagement({ profiles, currentUserId, onUpdate, onRemove, onRefresh }) {
  const [editing, setEditing]    = useState(null);
  const [confirmDel, setConfirm] = useState(null);
  const [showNew, setShowNew]    = useState(false);
  const adminCount = profiles.filter(u => u.isAdmin).length;
  const sorted = [...profiles].sort((a,b) => (Number(b.isAdmin) - Number(a.isAdmin)) || (a.name||"").localeCompare(b.name||""));

  return (
    <div>
      {showNew && <NewAccessInfoModal onClose={()=>setShowNew(false)}/>}
      {editing && <AccessEditModal profile={editing} onSave={onUpdate} onClose={()=>setEditing(null)}/>}
      {confirmDel && <ConfirmDialog title="Remover acesso" danger confirmLabel="Remover"
        message={`Remover o acesso de "${confirmDel.name}"? A pessoa deixa de ver os dados na plataforma. Os registros já lançados são mantidos. (Para apagar o login por completo, use o painel do Supabase.)`}
        onConfirm={()=>onRemove(confirmDel)} onClose={()=>setConfirm(null)}/>}

      <div style={{display:"flex",alignItems:"center",gap:13,marginBottom:18,flexWrap:"wrap"}}>
        <HeadChip icon="lock"/>
        <div style={{flex:1,minWidth:200}}>
          <h1 style={{...Ty.h1,fontSize:22}}>Gestão de acessos</h1>
          <div style={{...Ty.small, marginTop:3}}>{profiles.length} usuário(s) · {adminCount} administrador(es). Ajuste papéis e remova acessos.</div>
        </div>
        <Btn onClick={onRefresh}>Atualizar lista</Btn>
        <Btn primary onClick={()=>setShowNew(true)}>+ Novo acesso</Btn>
      </div>

      <Card style={{padding:0,overflow:"hidden"}}>
        <div className="fc-scroll" style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
            <thead><tr style={{background:T.canvas}}>
              {["Usuário","Papel",""].map((h,i)=>
                <th key={i} style={{padding:"10px 14px",textAlign:i===2?"right":"left",borderBottom:`1px solid ${T.line}`,fontWeight:600,color:T.muted,whiteSpace:"nowrap"}}>{h}</th>
              )}
            </tr></thead>
            <tbody>
              {sorted.map(u=>{
                const isSelf = u.id === currentUserId;
                const lastAdmin = u.isAdmin && adminCount <= 1;
                return (
                  <tr key={u.id} className="fc-row" style={{borderBottom:`1px solid ${T.lineSoft}`}}>
                    <td style={{padding:"10px 14px"}}>
                      <span style={{display:"inline-flex",alignItems:"center",gap:9}}>
                        <Avatar name={u.name} size={28} admin={u.isAdmin}/>
                        <span style={{fontWeight:600,color:T.ink}}>{u.name}{isSelf&&<span style={{fontSize:11,color:T.muted,fontWeight:500}}> (você)</span>}</span>
                      </span>
                    </td>
                    <td style={{padding:"10px 14px"}}><Badge label={u.isComercial?`Comercial${u.bu?` · ${u.bu}`:""}`:u.isViewer?"Somente visualização":u.isAdmin?"Administrador":"Analista"} color={u.isComercial?"orange":u.isViewer?"teal":u.isAdmin?"blue":"gray"} small dot/></td>
                    <td style={{padding:"10px 14px",textAlign:"right",whiteSpace:"nowrap"}}>
                      <Btn small onClick={()=>setEditing(u)} style={{marginRight:6}}>Editar</Btn>
                      <Btn small danger disabled={isSelf||lastAdmin} onClick={()=>setConfirm(u)}
                        title={isSelf?"Você não pode remover o próprio acesso":lastAdmin?"É preciso ao menos um administrador":"Remover acesso"}>Remover</Btn>
                    </td>
                  </tr>
                );
              })}
              {sorted.length===0 && <tr><td colSpan={3} style={{padding:"18px 14px",textAlign:"center",color:T.muted}}>Nenhum usuário encontrado. Clique em "Atualizar lista".</td></tr>}
            </tbody>
          </table>
        </div>
      </Card>

      <div style={{marginTop:14,fontSize:12,color:T.muted,display:"flex",gap:8,alignItems:"flex-start"}}>
        <span aria-hidden="true"></span>
        <span>Os acessos ficam no banco (Supabase) e valem para todos. Sempre deve existir ao menos um administrador, e você não pode remover o seu próprio acesso. Para criar um login novo, use "+ Novo acesso".</span>
      </div>
    </div>
  );
}

// ─── CLIENTES (perfil de faturamento) ────────────────────────────────────────

// Seção de formulário — definida FORA do modal (evita perda de foco a cada tecla).
function CSec({ title, children, grid=true }) {
  return (
    <div style={{marginBottom:16}}>
      <div style={{fontSize:12,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:".4px",marginBottom:8}}>{title}</div>
      {grid ? <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(200px,1fr))",gap:12}}>{children}</div> : children}
    </div>
  );
}

function ClientModal({ client, onSave, onDelete, onClose }) {
  const isNew = !client?.id;
  const [f, setF] = useState(client || { temPortal:false });
  const [passos, setPassos] = useState(() => { const v=parseJSON(client?.calendario, []); return Array.isArray(v)?v:[]; });
  const [peps, setPeps] = useState(() => { const m=parseJSON(client?.tiposPeps, {}); return (m && typeof m==="object" && !Array.isArray(m))?m:{}; });
  const [propostas, setPropostas] = useState(() => { const a=parseJSON(client?.propostas, null); if(Array.isArray(a)) return a.length?a:[""]; return client?.propostaUrl?[client.propostaUrl]:[""]; });
  const [entidades, setEnt] = useState(() => { const a=parseJSON(client?.cnpjs, null); if(Array.isArray(a)&&a.length) return a; return [{ razao:client?.nome||"", cnpj:client?.cnpj||"", codSap:client?.codSap||"" }]; });
  const [projetos, setProjetos] = useState(() => { const a=parseJSON(client?.projetos, []); return Array.isArray(a)?a:[]; });
  const [tab, setTab] = useState("dados");
  const [err, setErr] = useState("");
  const set = (k,v) => setF(p=>({...p,[k]:v}));

  const selTipos = (f.tiposContrato||"").split(",").map(s=>s.trim()).filter(Boolean);
  function toggleTipo(t) {
    const has = selTipos.includes(t);
    set("tiposContrato", (has ? selTipos.filter(x=>x!==t) : [...selTipos,t]).join(", "));
    setPeps(m => { const n={...m}; if(has) delete n[t]; else if(!Array.isArray(n[t])||!n[t].length) n[t]=[""]; return n; });
  }
  const setPep = (t,i,v) => setPeps(m=>({...m,[t]:(m[t]||[]).map((p,j)=>j===i?v:p)}));
  const addPep = (t) => setPeps(m=>({...m,[t]:[...(m[t]||[]),""]}));
  const delPep = (t,i) => setPeps(m=>({...m,[t]:(m[t]||[]).filter((_,j)=>j!==i)}));

  const selPortalTipos = (f.portalTipo||"").split(",").map(s=>s.trim()).filter(Boolean);
  const togglePortalTipo = (t) => set("portalTipo", (selPortalTipos.includes(t)?selPortalTipos.filter(x=>x!==t):[...selPortalTipos,t]).join(", "));

  const setProp = (i,v) => setPropostas(a=>a.map((p,j)=>j===i?v:p));
  const addProp = () => setPropostas(a=>[...a,""]);
  const delProp = (i) => setPropostas(a=>a.filter((_,j)=>j!==i));

  const addPasso = () => setPassos(a=>[...a,{quando:"",etapa:"",oQueFazer:""}]);
  const setPasso = (i,k,v) => setPassos(a=>a.map((p,j)=>j===i?{...p,[k]:v}:p));
  const delPasso = (i) => setPassos(a=>a.filter((_,j)=>j!==i));

  const setEntidade = (i,k,v) => setEnt(a=>a.map((e,j)=>j===i?{...e,[k]:v}:e));
  const addEntidade = () => setEnt(a=>[...a,{razao:"",cnpj:"",codSap:""}]);
  const delEntidade = (i) => setEnt(a=>a.filter((_,j)=>j!==i));

  const addProjeto = () => setProjetos(a=>[...a,{nome:"",pep:"",inicio:"",vencimento:"",valor:"",status:"Ativo",obs:""}]);
  const setProjeto = (i,k,v) => setProjetos(a=>a.map((p,j)=>j===i?{...p,[k]:v}:p));
  const delProjeto = (i) => setProjetos(a=>a.filter((_,j)=>j!==i));
  const hojeISO = new Date().toISOString().slice(0,10);
  const projVencido = (p) => p.vencimento && String(p.vencimento).slice(0,10) < hojeISO && p.status!=="Encerrado" && p.status!=="Renovado";

  function save() {
    if (!(f.nome||"").trim()) { setTab("dados"); setErr("Informe o nome do cliente/grupo."); return; }
    if (f.temPortal && (!(f.portalLink||"").trim() || !(f.portalUsuario||"").trim() || !(f.portalSenha||"").trim())) {
      setTab("dados"); setErr("Como o cliente tem portal, preencha o Link, o Usuário e a Senha do portal."); return;
    }
    const cleanPeps = {}; selTipos.forEach(t=>{ const arr=(peps[t]||[]).map(s=>s.trim()).filter(Boolean); if(arr.length) cleanPeps[t]=arr; });
    const cleanProp = propostas.map(s=>s.trim()).filter(Boolean);
    const cleanEnt = entidades.map(e=>({razao:(e.razao||"").trim(),cnpj:(e.cnpj||"").replace(/\D/g,"").slice(0,14),codSap:(e.codSap||"").trim()})).filter(e=>e.cnpj||e.razao||e.codSap);
    const cleanProj = projetos.map(p=>({ nome:(p.nome||"").trim(), pep:(p.pep||"").trim(), inicio:p.inicio||"", vencimento:p.vencimento||"", valor:(p.valor||"").toString().trim(), status:p.status||"Ativo", obs:(p.obs||"").trim() })).filter(p=>p.nome||p.pep||p.vencimento);
    onSave({ ...f, nome:f.nome.trim(), calendario: JSON.stringify(passos), tiposPeps: JSON.stringify(cleanPeps), propostas: JSON.stringify(cleanProp),
      projetos: JSON.stringify(cleanProj), processo: (f.processo||"").trim(),
      cnpjs: JSON.stringify(cleanEnt), cnpj: cleanEnt[0]?.cnpj || "", codSap: cleanEnt[0]?.codSap || "" });
    onClose();
  }

  return (
    <Modal title={isNew?"Novo cliente":`Cliente — ${client.nome}`} subtitle="Perfil de faturamento do cliente" onClose={onClose} wide>
      {/* Abas: Dados x Calendário (passo a passo) */}
      <div style={{display:"flex",gap:6,borderBottom:`1px solid ${T.line}`,marginBottom:18}}>
        {[["dados","Dados do cliente"],["processo","Processo & projetos"],["calendario","Calendário (passo a passo)"]].map(([id,label])=>(
          <button key={id} onClick={()=>setTab(id)} style={{border:"none",background:"none",cursor:"pointer",padding:"8px 12px",fontSize:13,fontWeight:tab===id?700:500,color:tab===id?T.brand:T.muted,borderBottom:`2px solid ${tab===id?T.brand:"transparent"}`,marginBottom:-1}}>{label}</button>
        ))}
      </div>

      {tab==="dados" && <>
      {f.incompleto && <div style={{marginBottom:16,padding:"11px 14px",borderRadius:T.rMd,background:T.warnBg,border:`1px solid ${T.warnLine}`,fontSize:12.5,color:T.warn,display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
        <span style={{flex:1,minWidth:200}}><b>Cadastro incompleto</b> (veio da carga em massa). Complete os dados e marque como concluído.</span>
        <Btn small primary onClick={()=>set("incompleto",false)}>✓ Marcar como completo</Btn>
      </div>}
      <CSec title="Identificação">
        <Field label="Nome do cliente / grupo *"><input style={inp} placeholder="Ex: Klabin" value={f.nome||""} onChange={e=>{set("nome",e.target.value);setErr("");}}/></Field>
        <Field label="Grupo de empresa"><input style={inp} value={f.grupoEmpresa||""} onChange={e=>set("grupoEmpresa",e.target.value)}/></Field>
        <Field label="Analista responsável"><input style={inp} placeholder="Nome do analista dono" value={f.owner||""} onChange={e=>set("owner",e.target.value)}/></Field>
      </CSec>

      {/* Empresas do grupo — vários CNPJs num só cadastro */}
      <div style={{marginBottom:16}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:8,gap:8}}>
          <div>
            <div style={{fontSize:12,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:".4px"}}>Empresas do grupo ({entidades.length} CNPJ{entidades.length!==1?"s":""})</div>
            <div style={{fontSize:11,color:T.muted,marginTop:2}}>As regras de faturamento abaixo valem para todos estes CNPJs.</div>
          </div>
          <Btn small onClick={addEntidade}>+ Adicionar CNPJ</Btn>
        </div>
        {entidades.map((e,i)=>(
          <div key={i} style={{border:`1px solid ${T.line}`,borderRadius:T.rLg,padding:"10px 12px",marginBottom:8,background:"var(--surface)"}}>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:6}}>
              <span style={{fontSize:12,fontWeight:700,color:T.brand}}>{i===0?"Empresa principal":`Empresa ${i+1}`}</span>
              {entidades.length>1 && <button onClick={()=>delEntidade(i)} title="Remover" style={{border:"none",background:"none",cursor:"pointer",color:T.danger,fontSize:12,fontWeight:600}}>✕ Remover</button>}
            </div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 160px 130px",gap:10}}>
              <Field label="Razão social"><input style={inp} placeholder="Razão social desta empresa" value={e.razao||""} onChange={ev=>setEntidade(i,"razao",ev.target.value)}/></Field>
              <Field label="CNPJ"><input style={inp} inputMode="numeric" placeholder="00000000000000" value={e.cnpj||""} onChange={ev=>setEntidade(i,"cnpj",ev.target.value.replace(/\D/g,"").slice(0,14))}/></Field>
              <Field label="Cód. SAP"><input style={inp} inputMode="numeric" placeholder="código" value={e.codSap||""} onChange={ev=>setEntidade(i,"codSap",ev.target.value.replace(/\D/g,"").slice(0,10))}/></Field>
            </div>
          </div>
        ))}
      </div>

      <CSec title="Tipos de contrato" grid={false}>
        <div style={{display:"flex",flexWrap:"wrap",gap:8,marginBottom:selTipos.length?12:0}}>
          {TIPOS_PROJETO.map(t=>{ const on=selTipos.includes(t); return (
            <label key={t} style={{display:"flex",alignItems:"center",gap:7,padding:"8px 12px",borderRadius:T.rMd,border:`1.5px solid ${on?T.brand:T.line}`,background:on?T.brandBg:"var(--surface)",cursor:"pointer",fontSize:13,fontWeight:on?600:400,color:on?T.brand:T.inkSoft}}>
              <input type="checkbox" checked={on} onChange={()=>toggleTipo(t)} style={{width:15,height:15}}/>{t}
            </label>
          );})}
        </div>
        {selTipos.map(t=>(
          <div key={t} style={{border:`1px solid ${T.line}`,borderRadius:T.rLg,padding:"10px 12px",marginBottom:8,background:"var(--surface)"}}>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:8,gap:8}}>
              <span style={{fontSize:12,fontWeight:700,color:T.brand}}>PEPs · {t}</span>
              <Btn small onClick={()=>addPep(t)}>+ PEP</Btn>
            </div>
            {(peps[t]||[""]).map((pep,i)=>(
              <div key={i} style={{display:"flex",alignItems:"center",gap:8,marginBottom:6}}>
                <input style={inp} placeholder="Ex: BR02CLP00005.1.1" value={pep} onChange={e=>setPep(t,i,e.target.value)}/>
                {(peps[t]||[]).length>1 && <button onClick={()=>delPep(t,i)} title="Remover PEP" style={{border:"none",background:"none",cursor:"pointer",color:T.danger,fontSize:14,padding:"0 4px",flexShrink:0}}>✕</button>}
              </div>
            ))}
          </div>
        ))}
      </CSec>

      <CSec title="Propostas" grid={false}>
        {propostas.map((p,i)=>(
          <div key={i} style={{display:"flex",alignItems:"center",gap:8,marginBottom:6}}>
            <input style={inp} placeholder="Cole o link da proposta (Drive, SharePoint...)" value={p} onChange={e=>setProp(i,e.target.value)}/>
            {propostas.length>1 && <button onClick={()=>delProp(i)} title="Remover" style={{border:"none",background:"none",cursor:"pointer",color:T.danger,fontSize:14,padding:"0 4px",flexShrink:0}}>✕</button>}
          </div>
        ))}
        <Btn small onClick={addProp}>+ Adicionar proposta</Btn>
      </CSec>

      <CSec title="Faturamento">
        <Field label="Período de faturamento"><input style={inp} placeholder="Ex: 01 a 31 (ou outro)" value={f.periodoFaturamento||""} onChange={e=>set("periodoFaturamento",e.target.value)}/></Field>
        <Field label="Dia de corte" hint="(período quebrado — vazio = 01 a 31)"><input type="number" min="1" max="28" style={inp} placeholder="Ex: 10, 20…" value={f.diaCorte||""} onChange={e=>set("diaCorte",e.target.value)}/></Field>
        <Field label="Prazo de vencimento acordado"><input style={inp} placeholder="Ex: 30 dias" value={f.prazoVencimento||""} onChange={e=>set("prazoVencimento",e.target.value)}/></Field>
        <Field label="Forma de pagamento"><input style={inp} placeholder="Ex: Boleto, transferência" value={f.formaPagamento||""} onChange={e=>set("formaPagamento",e.target.value)}/></Field>
      </CSec>

      <div style={{marginBottom:16,padding:"12px 14px",borderRadius:T.rLg,background:T.canvas,border:`1px solid ${T.line}`}}>
        <label style={{display:"flex",alignItems:"center",gap:8,fontSize:13,fontWeight:700,color:T.ink,cursor:"pointer",marginBottom:f.temPortal?12:0}}>
          <input type="checkbox" checked={!!f.temPortal} onChange={e=>set("temPortal",e.target.checked)} style={{width:16,height:16}}/>
          Tem portal?
        </label>
        {f.temPortal && <div>
          <div style={{marginBottom:12}}>
            <label style={{fontSize:12,fontWeight:600,color:T.inkSoft,display:"block",marginBottom:6}}>Classificação do portal</label>
            <div style={{display:"flex",flexWrap:"wrap",gap:8}}>
              {PORTAL_TIPOS.map(t=>{ const on=selPortalTipos.includes(t); return (
                <label key={t} style={{display:"flex",alignItems:"center",gap:7,padding:"7px 11px",borderRadius:T.rMd,border:`1.5px solid ${on?T.brand:T.line}`,background:on?T.brandBg:"var(--surface)",cursor:"pointer",fontSize:13,fontWeight:on?600:400,color:on?T.brand:T.inkSoft}}>
                  <input type="checkbox" checked={on} onChange={()=>togglePortalTipo(t)} style={{width:15,height:15}}/>{t}
                </label>
              );})}
            </div>
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(200px,1fr))",gap:12}}>
            <Field label="Link do portal *"><input style={inp} value={f.portalLink||""} onChange={e=>set("portalLink",e.target.value)}/></Field>
            <Field label="Usuário *"><input style={inp} value={f.portalUsuario||""} onChange={e=>set("portalUsuario",e.target.value)}/></Field>
            <Field label="Senha *"><input style={inp} value={f.portalSenha||""} onChange={e=>set("portalSenha",e.target.value)}/></Field>
            <Field label="Link do passo a passo do portal"><input style={inp} placeholder="Cole o link" value={f.portalPassoUrl||""} onChange={e=>set("portalPassoUrl",e.target.value)}/></Field>
          </div>
        </div>}
      </div>

      <div style={{marginBottom:16,padding:"12px 14px",borderRadius:T.rLg,background:T.canvas,border:`1px solid ${T.line}`}}>
        <label style={{display:"flex",alignItems:"center",gap:8,fontSize:13,fontWeight:700,color:T.ink,cursor:"pointer"}}>
          <input type="checkbox" checked={!!f.aceitaVariacao} onChange={e=>set("aceitaVariacao",e.target.checked)} style={{width:16,height:16}}/>
          Aceita variação de receita pós-fechamento?
        </label>
        <div style={{fontSize:11.5,color:T.muted,marginTop:6}}>Habilita, na Minha visão, o lançamento de valores extras por consultor <b>depois</b> que a receita foi fechada (ex.: Casas Bahia). O extra vira saldo a faturar e aparece nos relatórios como variação.</div>
      </div>

      <CSec title="Contato financeiro">
        <Field label="Nome"><input style={inp} value={f.contatoFinanceiro||""} onChange={e=>set("contatoFinanceiro",e.target.value)}/></Field>
        <Field label="E-mail"><input style={inp} type="email" placeholder="financeiro@cliente.com" value={f.contatoFinanceiroEmail||""} onChange={e=>set("contatoFinanceiroEmail",e.target.value)}/></Field>
      </CSec>

      <CSec title="Account manager (comercial)">
        <Field label="Nome"><input style={inp} value={f.accountManager||""} onChange={e=>set("accountManager",e.target.value)}/></Field>
        <Field label="E-mail"><input style={inp} type="email" placeholder="am@grupofcamara.com" value={f.accountManagerEmail||""} onChange={e=>set("accountManagerEmail",e.target.value)}/></Field>
      </CSec>
      </>}

      {tab==="processo" && <>
        <CSec title="Processo do cliente" grid={false}>
          <div style={{fontSize:12,color:T.muted,marginBottom:8}}>O jeito que este cliente funciona de ponta a ponta — particularidades, quem aprova, ordem de faturamento, o que costuma travar. O conhecimento que hoje fica na cabeça do analista, centralizado aqui.</div>
          <textarea style={{...inp,minHeight:150,resize:"vertical",lineHeight:1.55}} placeholder="Ex: Fatura só depois da medição no portal. O gerente X aprova as horas até dia 5. NF sempre por CNPJ da filial. Costuma atrasar aprovação no fim do trimestre…" value={f.processo||""} onChange={e=>set("processo",e.target.value)}/>
        </CSec>

        <div style={{marginBottom:8,marginTop:6,padding:"10px 12px",borderRadius:T.rMd,background:T.canvas,border:`1px solid ${T.line}`,fontSize:12,color:T.inkSoft}}>
          <b>Prazos & datas deste cliente</b> — dia de corte, prazo de vencimento e período de faturamento ficam na aba <b>Dados do cliente → Faturamento</b>. As propostas, na aba <b>Dados do cliente → Propostas</b>.
        </div>

        <CSec title="Projetos & vencimentos" grid={false}>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:10,gap:8,flexWrap:"wrap"}}>
            <div style={{fontSize:12,color:T.muted}}>Contratos/projetos do cliente com data de vencimento — para enxergar o que está <b style={{color:C.red.solid}}>vencido</b> ou a vencer.</div>
            <Btn small primary onClick={addProjeto}>+ Projeto</Btn>
          </div>
          {projetos.length===0 && <div style={{fontSize:13,color:T.muted,padding:"20px 16px",textAlign:"center",background:T.canvas,borderRadius:T.rLg,border:`1px dashed ${T.line}`}}>Nenhum projeto cadastrado.<br/>Clique em “+ Projeto” para incluir contratos e datas de vencimento.</div>}
          {projetos.map((p,i)=>(
            <div key={i} style={{border:`1px solid ${projVencido(p)?"#fca5a5":T.line}`,borderRadius:T.rLg,padding:"12px",marginBottom:8,background:projVencido(p)?T.dangerBg:"var(--surface)"}}>
              <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:8,gap:8}}>
                <span style={{display:"inline-flex",alignItems:"center",gap:8,fontSize:12,fontWeight:700,color:T.brand}}>Projeto {i+1}{projVencido(p) && <Badge label="VENCIDO" color="red" small/>}</span>
                <button onClick={()=>delProjeto(i)} title="Remover" style={{border:"none",background:"none",cursor:"pointer",color:T.danger,fontSize:12,fontWeight:600}}>✕ Remover</button>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
                <Field label="Nome do projeto / contrato"><input style={inp} placeholder="Ex: Sustentação SAP 2026" value={p.nome||""} onChange={e=>setProjeto(i,"nome",e.target.value)}/></Field>
                <Field label="PEP"><input style={inp} placeholder="Ex: BR02CLP00046" value={p.pep||""} onChange={e=>setProjeto(i,"pep",e.target.value)}/></Field>
              </div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1.2fr",gap:10}}>
                <Field label="Início"><input type="date" style={inp} value={p.inicio||""} onChange={e=>setProjeto(i,"inicio",e.target.value)}/></Field>
                <Field label="Vencimento"><input type="date" style={inp} value={p.vencimento||""} onChange={e=>setProjeto(i,"vencimento",e.target.value)}/></Field>
                <Field label="Valor do contrato"><input style={inp} placeholder="Ex: 480.000" value={p.valor||""} onChange={e=>setProjeto(i,"valor",e.target.value)}/></Field>
                <Field label="Status"><select style={inp} value={p.status||"Ativo"} onChange={e=>setProjeto(i,"status",e.target.value)}>{["Ativo","A vencer","Vencido","Renovado","Encerrado"].map(s=><option key={s}>{s}</option>)}</select></Field>
              </div>
              <Field label="Observação"><input style={inp} placeholder="Renovação em negociação, aditivo pendente…" value={p.obs||""} onChange={e=>setProjeto(i,"obs",e.target.value)}/></Field>
            </div>
          ))}
        </CSec>
      </>}

      {tab==="calendario" && <>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12,gap:8,flexWrap:"wrap"}}>
          <div>
            <div style={{fontSize:14,fontWeight:700,color:T.ink}}>Calendário de faturamento</div>
            <div style={{fontSize:12,color:T.muted}}>Monte o passo a passo deste cliente. Cada etapa tem quando acontece, o nome da etapa e o que fazer.</div>
          </div>
          <Btn primary small onClick={addPasso}>+ Incluir etapa</Btn>
        </div>
        {passos.length===0 && <div style={{fontSize:13,color:T.muted,padding:"24px 16px",textAlign:"center",background:T.canvas,borderRadius:T.rLg,border:`1px dashed ${T.line}`}}>Nenhuma etapa adicionada ainda.<br/>Clique em “+ Incluir etapa” para começar.</div>}
        {passos.map((p,i)=>(
          <div key={i} style={{border:`1px solid ${T.line}`,borderRadius:T.rLg,marginBottom:10,background:"var(--surface)",overflow:"hidden"}}>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"10px 14px",background:T.brandBg,borderBottom:`1px solid ${T.line}`}}>
              <span style={{display:"inline-flex",alignItems:"center",gap:8,fontSize:13,fontWeight:700,color:T.brand}}>
                <span style={{width:22,height:22,borderRadius:"50%",background:T.brand,color:"#fff",display:"inline-flex",alignItems:"center",justifyContent:"center",fontSize:12}}>{i+1}</span>
                Passo {i+1}
              </span>
              <button onClick={()=>delPasso(i)} title="Remover etapa" style={{border:"none",background:"none",cursor:"pointer",color:T.danger,fontSize:12,fontWeight:600}}>✕ Remover</button>
            </div>
            <div style={{padding:"14px"}}>
              <div style={{display:"grid",gridTemplateColumns:"160px 1fr",gap:12,marginBottom:12}}>
                <Field label="Quando"><input style={inp} placeholder="Ex: Dia 25" value={p.quando||""} onChange={e=>setPasso(i,"quando",e.target.value)}/></Field>
                <Field label="Etapa"><input style={inp} placeholder="Ex: Extração de dados no FC Team" value={p.etapa||""} onChange={e=>setPasso(i,"etapa",e.target.value)}/></Field>
              </div>
              <Field label="O que fazer"><textarea style={{...inp,minHeight:64,resize:"vertical",lineHeight:1.5}} placeholder="Descreva a ação desta etapa em detalhe..." value={p.oQueFazer||""} onChange={e=>setPasso(i,"oQueFazer",e.target.value)}/></Field>
            </div>
          </div>
        ))}
        {passos.length>0 && <div style={{textAlign:"center",marginTop:6}}><Btn small onClick={addPasso}>+ Incluir etapa</Btn></div>}
      </>}

      {err&&<div style={{marginBottom:12,fontSize:12,padding:"8px 12px",borderRadius:T.rMd,background:T.dangerBg,color:T.danger,border:`1px solid ${T.dangerLine}`}}>{err}</div>}

      <div style={{display:"flex",gap:8,justifyContent:"space-between",alignItems:"center"}}>
        <div>{!isNew && <Btn danger small onClick={()=>{onDelete(client.id);onClose();}}>Excluir</Btn>}</div>
        <div style={{display:"flex",gap:8}}>
          <Btn onClick={onClose}>Cancelar</Btn>
          <Btn primary onClick={save}>{isNew?"Criar cliente":"Salvar"}</Btn>
        </div>
      </div>
    </Modal>
  );
}

// Lê a exportação SAP de parceiros de negócios → cadastros de clientes (incompletos).
function parseClientsSheet(rows) {
  const h = (rows[0]||[]).map(x=>String(x).trim());
  const col = (name) => h.findIndex(x=>x.toLowerCase()===name.toLowerCase());
  const iPN=col("Parceiro de negócios"), iNome=col("Nome 1"), iFiscal=col("Número de identificação fiscal");
  if (iNome===-1 && iPN===-1) return { clients:[], errors:["Não encontrei a coluna de nome (\"Nome 1\" ou \"Parceiro de negócios\")."] };
  const out=[]; const seen=new Set(); let dup=0;
  for (let i=1;i<rows.length;i++) {
    const r=rows[i]; if(!r) continue;
    const pn=String(iPN>=0?r[iPN]:"").trim();
    const nome=(String(iNome>=0?r[iNome]:"").trim()) || pn.replace(/\s*\(\d+\)\s*$/,"").trim();
    if(!nome) continue;
    const m=pn.match(/\((\d{4,})\)\s*$/); const codSap=m?m[1]:"";
    const cnpj=String(iFiscal>=0?r[iFiscal]:"").split(/[,;\n]/)[0].replace(/\D/g,"").slice(0,14);
    const key=codSap||nome.toLowerCase();
    if(seen.has(key)){dup++;continue;} seen.add(key);
    out.push({ nome, codSap, cnpj, incompleto:true });
  }
  const errors=[]; if(dup) errors.push(`${dup} duplicado(s) na planilha foram ignorados.`);
  return { clients:out, errors };
}

function ClientImportModal({ existing, onImport, onClose }) {
  const [preview,setPreview]=useState(null);
  const [fileName,setFileName]=useState("");
  const [msgs,setMsgs]=useState([]);
  const [loading,setLoading]=useState(false);
  const fileRef=useRef();
  const existingCods = new Set(existing.map(c=>(c.codSap||"").trim()).filter(Boolean));
  const existingNames = new Set(existing.map(c=>(c.nome||"").trim().toLowerCase()));

  function readFile(file){
    setLoading(true);setFileName(file.name);setPreview(null);setMsgs([]);
    const reader=new FileReader();
    reader.onload=e=>{
      try{
        const wb=XLSX.read(new Uint8Array(e.target.result),{type:"array"});
        const rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1,defval:""});
        const {clients,errors}=parseClientsSheet(rows);
        const novos=clients.filter(c=>!(c.codSap&&existingCods.has(c.codSap)) && !existingNames.has(c.nome.toLowerCase()));
        const jaExistem=clients.length-novos.length;
        const m=errors.map(x=>({type:"warn",text:x}));
        if(jaExistem) m.push({type:"warn",text:`${jaExistem} cliente(s) já cadastrados foram ignorados.`});
        if(novos.length===0){m.push({type:"error",text:"Nenhum cliente novo para importar."});setMsgs(m);}
        else{m.push({type:"ok",text:`${novos.length} cliente(s) novo(s) prontos para importar.`});setMsgs(m);setPreview(novos);}
      }catch(err){setMsgs([{type:"error",text:"Erro ao ler o arquivo: "+err.message}]);}
      setLoading(false);
    };
    reader.readAsArrayBuffer(file);
  }
  const mc={ok:{bg:T.okBg,text:T.ok,border:T.okLine},warn:{bg:T.warnBg,text:T.warn,border:T.warnLine},error:{bg:T.dangerBg,text:T.danger,border:T.dangerLine}};

  return (
    <Modal title="Importar clientes" subtitle="Carga em massa (exportação do SAP). Entram como cadastro incompleto." onClose={onClose} wide>
      <input type="file" ref={fileRef} style={{display:"none"}} accept=".xlsx,.xls,.csv" onChange={e=>{if(e.target.files[0])readFile(e.target.files[0]);e.target.value="";}}/>
      <div onClick={()=>fileRef.current.click()} role="button" tabIndex={0}
        style={{border:`2px dashed ${fileName?T.okLine:"#cbd2dc"}`,borderRadius:T.rLg,padding:"28px 20px",textAlign:"center",cursor:"pointer",background:fileName?T.okBg:"var(--canvas)",marginBottom:14}}>
        {loading?<div style={{color:T.muted,fontSize:13}}>Lendo...</div>:fileName?<><div style={{color:T.ok,marginBottom:6}}><Icon name="check" size={26}/></div><div style={{fontSize:13,fontWeight:700,color:T.ok}}>{fileName}</div><div style={{fontSize:11,color:T.muted}}>Clique para trocar</div></>:<><div style={{color:T.muted,marginBottom:8}}><Icon name="upload" size={26}/></div><div style={{fontSize:14,fontWeight:600,color:T.inkSoft}}>Clique para selecionar a planilha</div></>}
      </div>
      {msgs.map((m,i)=><div key={i} style={{marginBottom:6,fontSize:12,padding:"8px 12px",borderRadius:T.rMd,background:mc[m.type].bg,color:mc[m.type].text,border:`1px solid ${mc[m.type].border}`}}>{m.text}</div>)}
      {preview&&<div style={{marginBottom:14,padding:"10px 14px",borderRadius:T.rMd,background:T.okBg,border:`1px solid ${T.okLine}`,fontSize:12,color:T.ok}}>
        Prévia: {preview.slice(0,3).map(c=>c.nome).join(", ")}{preview.length>3?` … e mais ${preview.length-3}`:""}
      </div>}
      <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
        <Btn onClick={onClose}>Cancelar</Btn>
        <Btn primary disabled={!preview} onClick={()=>{ if(preview){onImport(preview);} onClose(); }}>Importar {preview?`(${preview.length})`:""}</Btn>
      </div>
    </Modal>
  );
}

// ─── CONCILIAÇÃO DE NOTAS (NFS-e da prefeitura) ──────────────────────────────
const PREFEITURAS = ["São Paulo", "Maringá", "Florianópolis", "Belo Horizonte"];
const brl = (n) => "R$ " + (Number(n)||0).toLocaleString("pt-BR", { minimumFractionDigits:2, maximumFractionDigits:2 });
const onlyDigits = (s) => String(s||"").replace(/\D/g, "");
const stripAcc = (s) => String(s||"").toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g, "");
const normHdr  = (s) => stripAcc(s).replace(/[^a-z0-9]+/g, " ").trim();
const colByExact = (headers, exact) => { const t = normHdr(exact); return headers.findIndex(h => normHdr(h) === t); };

// Número em formato BR ("80.412,95") ou US ("80,412.95") ou simples ("69"/"2.07").
function parseBR(v) {
  let s = String(v==null?"":v).trim().replace(/[^\d.,-]/g, "");
  if (!s) return 0;
  const hasDot = s.includes("."), hasComma = s.includes(",");
  if (hasDot && hasComma) {
    if (s.lastIndexOf(",") > s.lastIndexOf(".")) s = s.replace(/\./g, "").replace(",", "."); // BR
    else s = s.replace(/,/g, "");                                                            // US
  } else if (hasComma) {
    s = s.replace(/\./g, "").replace(",", ".");
  } else if (hasDot) {
    const dots = (s.match(/\./g) || []).length;
    if (dots > 1) s = s.replace(/\./g, "");                          // 1.234.567 → milhar
    else if ((s.split(".")[1] || "").length === 3) s = s.replace(/\./g, ""); // 1.234 → milhar
  }
  return parseFloat(s) || 0;
}
// "29/06/2026 17:24:52" ou "29/06/2026" → ISO
function brToISO(v) {
  const m = String(v||"").trim().match(/(\d{2})\/(\d{2})\/(\d{4})(?:\s+(\d{2}):(\d{2})(?::(\d{2}))?)?/);
  if (!m) return null;
  const [, d, mo, y, hh, mi, ss] = m;
  return hh ? `${y}-${mo}-${d}T${hh}:${mi}:${ss||"00"}` : `${y}-${mo}-${d}`;
}
const MESES = { janeiro:"01", fevereiro:"02", marco:"03", abril:"04", maio:"05", junho:"06", julho:"07", agosto:"08", setembro:"09", outubro:"10", novembro:"11", dezembro:"12" };
// Extrai PEDIDO (OV), competências e nomes da "Discriminação dos Serviços".
function parseDiscriminacao(txt) {
  const t = String(txt||"");
  const pedidos = [...new Set((t.match(/PEDIDO\s*(\d+)/gi) || []).map(x => x.replace(/\D/g, "")))];
  const comps = [...new Set((t.match(/(JANEIRO|FEVEREIRO|MAR[ÇC]O|ABRIL|MAIO|JUNHO|JULHO|AGOSTO|SETEMBRO|OUTUBRO|NOVEMBRO|DEZEMBRO)\s*\/?\s*(\d{4})/gi) || []).map(x => {
    const mm = stripAcc((x.match(/[A-Za-zÇç]+/) || [""])[0]);
    const yy = (x.match(/\d{4}/) || ["----"])[0];
    return (MESES[mm] || "??") + "/" + yy;
  }))];
  const nomes = [...new Set((t.match(/\b\d{5}\s+([A-Za-zÀ-ÿ][A-Za-zÀ-ÿ.\s]+?)(?=\s+\d{5}|\s+PEDIDO|\s+ITEM|$)/gi) || [])
    .map(x => x.replace(/^\d{5}\s+/, "").replace(/\s+ITEM.*$/i, "").trim()).filter(n => n.length > 3))];
  return { pedidos: pedidos.join(", "), competencias: comps.join(", "), profissionais: nomes.join(", ") };
}
// Candidatos de cabeçalho por campo — cobre SP, Maringá e variações.
const NOTE_COLS = {
  numero:      ["n nfs e", "numero", "numero da nota", "nota"],
  emitida:     ["data hora nfe", "emitido em", "data emissao", "data de emissao", "data da emissao"],
  fato:        ["data do fato gerador", "fato gerador"],
  prestCnpj:   ["cpf cnpj do prestador", "cpf cnpj prestador"],
  prestNome:   ["razao social do prestador", "razao social prestador"],
  situacao:    ["situacao da nota fiscal", "situacao"],
  cancel:      ["data de cancelamento"],
  valorServ:   ["valor dos servicos", "valor servicos", "valor do servico", "valor servico"],
  valorTotal:  ["valor total recebido", "valor total"],
  iss:         ["iss devido", "valor iss", "iss"],
  tomadorCnpj: ["cpf cnpj do tomador", "cpf cnpj tomador"],
  tomadorNome: ["razao social do tomador", "razao social tomador"],
  discrim:     ["discriminacao dos servicos", "descriminacao servico", "discriminacao servico", "descriminacao dos servicos", "discriminacao", "descriminacao"],
};
const colByCandidates = (headers, cands) => { for (const c of cands) { const i = colByExact(headers, c); if (i !== -1) return i; } return -1; };
// Acha a linha de cabeçalho (alguns relatórios têm título/linhas em branco antes).
function findHeaderRow(rows) {
  for (let i = 0; i < Math.min(15, rows.length); i++) {
    const h = (rows[i] || []).map(c => String(c || ""));
    if (colByCandidates(h, NOTE_COLS.numero) !== -1 && colByCandidates(h, NOTE_COLS.tomadorCnpj) !== -1 && colByCandidates(h, NOTE_COLS.valorServ) !== -1) return i;
  }
  return -1;
}
// Parser genérico de NFS-e (cabeçalhos por nome, mapeamento flexível).
function parseMunicipalSheet(rows, municipio) {
  if (!rows.length) return { notes: [], errors: ["Arquivo vazio."] };
  const hi = findHeaderRow(rows);
  if (hi === -1) return { notes: [], errors: ["Layout não reconhecido. Esperado um relatório de NFS-e com colunas de número, CNPJ do tomador e valor dos serviços."] };
  const headers = rows[hi].map(h => String(h || ""));
  const idx = {};
  for (const [k, cands] of Object.entries(NOTE_COLS)) idx[k] = colByCandidates(headers, cands);
  const get = (row, k) => idx[k] >= 0 ? (row[idx[k]] ?? "") : "";
  const notes = []; let ignoradas = 0;
  for (let i = hi + 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.every(c => c == null || c === "")) continue;
    const numero = String(get(row, "numero")).trim();
    if (!numero) { ignoradas++; continue; }
    // Pula linhas de total/rodapé (ex.: "Total;355;...") — não têm tomador válido.
    const tomadorCnpj = onlyDigits(get(row, "tomadorCnpj"));
    if (tomadorCnpj.length < 11 || /^total/i.test(String(get(row, "numero")).trim()) || /^total/i.test(String(row[0]||"").trim())) { ignoradas++; continue; }
    // Rodapé de soma costuma vir SEM CNPJ do prestador (toda NFS-e real tem).
    // Só aplica quando o arquivo tem a coluna de prestador (senão dropava tudo).
    if (idx.prestCnpj >= 0 && onlyDigits(get(row, "prestCnpj")).length < 11) { ignoradas++; continue; }
    const disc = String(get(row, "discrim") || "");
    const meta = parseDiscriminacao(disc);
    const situacao = String(get(row, "situacao")).trim();
    const cancelada = !!String(get(row, "cancel")).trim() || /^c/i.test(situacao);
    notes.push({
      municipio, numero,
      emitidaEm: brToISO(get(row, "emitida")), fatoGerador: brToISO(get(row, "fato")),
      prestadorCnpj: onlyDigits(get(row, "prestCnpj")), prestadorNome: String(get(row, "prestNome")).trim(),
      tomadorCnpj, tomadorNome: String(get(row, "tomadorNome")).trim(),
      valorServicos: parseBR(get(row, "valorServ")), valorTotal: parseBR(get(row, "valorTotal")) || parseBR(get(row, "valorServ")),
      iss: parseBR(get(row, "iss")), situacao, cancelada,
      pedidos: meta.pedidos, competencias: meta.competencias, profissionais: meta.profissionais,
      discriminacao: disc, importId: null,
    });
  }
  const errors = []; if (ignoradas) errors.push(`${ignoradas} linha(s) sem número de nota ignoradas.`);
  return { notes, errors };
}

function NotesImportModal({ onImport, onClose }) {
  const [municipio, setMun] = useState("São Paulo");
  const [empresa, setEmpresa] = useState("BR02");
  const [preview, setPreview] = useState(null);
  const [fileName, setFileName] = useState("");
  const [msgs, setMsgs] = useState([]);
  const [loading, setLoading] = useState(false);
  const [dragOver, setDragOver] = useState(false);
  const fileRef = useRef();
  const reset = () => { setPreview(null); setFileName(""); setMsgs([]); };

  function readFile(file) {
    setLoading(true); setFileName(file.name); setPreview(null); setMsgs([]);
    const reader = new FileReader();
    reader.onload = e => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result), { type:"array", codepage:1252 });
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header:1, defval:"", blankrows:false, raw:false });
        const { notes, errors } = parseMunicipalSheet(rows, municipio);
        const m = []; errors.forEach(x => m.push({ type:"warn", text:x }));
        if (!notes.length) { m.push({ type:"error", text:"Nenhuma nota válida encontrada." }); setMsgs(m); }
        else { const can = notes.filter(n=>n.cancelada).length; m.push({ type:"ok", text:`${notes.length} nota(s) lidas${can?` · ${can} cancelada(s)`:""}.` }); setMsgs(m); setPreview(notes); }
      } catch (err) { setMsgs([{ type:"error", text:"Erro ao ler o arquivo: " + err.message }]); }
      setLoading(false);
    };
    reader.readAsArrayBuffer(file);
  }
  const mc = { ok:{bg:T.okBg,text:T.ok,border:T.okLine}, warn:{bg:T.warnBg,text:T.warn,border:T.warnLine}, error:{bg:T.dangerBg,text:T.danger,border:T.dangerLine} };

  function doImport() {
    if (!preview?.length) return;
    const importId = uuid();
    onImport(preview.map(n => ({ ...n, empresa, importId })));
    onClose();
  }

  return (
    <Modal title="Importar notas da prefeitura" subtitle="Relatório de NFS-e (.csv)" onClose={onClose} wide
      footer={<><Btn onClick={onClose}>Cancelar</Btn><Btn primary disabled={!preview?.length} onClick={doImport}>Importar {preview?.length||0} nota(s)</Btn></>}>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(160px,1fr))",gap:12,marginBottom:14}}>
        <Field label="Empresa do grupo *" hint="(emissora da nota)"><select style={inp} value={empresa} onChange={e=>setEmpresa(e.target.value)}>{EMPRESAS.map(e=><option key={e.cod} value={e.cod}>{e.cod} — {e.nome}</option>)}</select></Field>
        <Field label="Prefeitura"><select style={inp} value={municipio} onChange={e=>{setMun(e.target.value);reset();}}>{PREFEITURAS.map(p=><option key={p}>{p}</option>)}</select></Field>
      </div>
      {!["São Paulo","Maringá"].includes(municipio) && <div style={{marginBottom:12,padding:"10px 12px",background:T.warnBg,border:`1px solid ${T.warnLine}`,borderRadius:T.rMd,fontSize:12,color:T.warn}}>
        Por enquanto os layouts de <b>São Paulo</b> e <b>Maringá</b> estão mapeados. Me mande uma amostra do relatório de {municipio} para eu adicionar o layout.
      </div>}
      <div onDragOver={e=>{e.preventDefault();setDragOver(true);}} onDragLeave={()=>setDragOver(false)}
        onDrop={e=>{e.preventDefault();setDragOver(false);const f=e.dataTransfer.files[0];if(f)readFile(f);}}
        onClick={()=>fileRef.current?.click()}
        style={{border:`2px dashed ${dragOver?T.brand:T.line}`,borderRadius:T.rLg,padding:"26px",textAlign:"center",cursor:"pointer",background:dragOver?T.brandBg:T.canvas}}>
        <input ref={fileRef} type="file" accept=".csv,.xlsx,.xls,text/csv" style={{display:"none"}} onChange={e=>{const f=e.target.files[0];if(f)readFile(f);}}/>
        <div style={{marginBottom:6,color:T.muted}}><Icon name="receipt" size={26}/></div>
        <div style={{fontSize:13,fontWeight:600,color:T.ink}}>{fileName||"Arraste o .csv/.xlsx aqui ou clique para escolher"}</div>
        <div style={{fontSize:11,color:T.muted,marginTop:3}}>{loading?"Lendo…":"Relatório de NFS-e (SP em .csv · Maringá em .xlsx)"}</div>
      </div>
      {msgs.map((m,i)=>(<div key={i} style={{marginTop:10,padding:"9px 12px",borderRadius:T.rMd,fontSize:12.5,background:mc[m.type].bg,color:mc[m.type].text,border:`1px solid ${mc[m.type].border}`}}>{m.text}</div>))}
      {preview?.length>0 && <div style={{marginTop:12,fontSize:12,color:T.muted}}>Pré-visualização: {preview.slice(0,3).map(n=>`NF ${n.numero} · ${n.tomadorNome||"—"} · ${brl(n.valorServicos)}`).join("  |  ")}{preview.length>3?"  …":""}</div>}
    </Modal>
  );
}

// Validações do sistema — só leitura. Confere a integridade do faturamento:
// (1) conciliado da prefeitura × receita conciliada (regra de R$ 1,00 por lote),
// (2) receitas possivelmente duplicadas, (3) faturado sem nota amarrada.
function ValidatorsView({ records, notes, faturamentos=[], fatByRec={}, varByRec={} }) {
  const [open, setOpen] = useState("");
  const bill = (r) => (r.valorTotal||0) + (varByRec[r.id]||0);

  // 1) Conciliado prefeitura vs receita conciliada — regra de R$ 1,00 por lote.
  const cids = [...new Set(faturamentos.filter(a=>a.conciliacaoId).map(a=>a.conciliacaoId))];
  const lotes = cids.map(cid => {
    const sn = notes.filter(n=>n.conciliacaoId===cid && !n.cancelada).reduce((s,n)=>s+(n.valorServicos||0),0);
    const sr = faturamentos.filter(a=>a.conciliacaoId===cid).reduce((s,a)=>s+(a.valor||0),0);
    const r0 = records.find(r=>faturamentos.some(a=>a.conciliacaoId===cid && a.recordId===r.id));
    const nfs = notes.filter(n=>n.conciliacaoId===cid).map(n=>n.numero).filter(Boolean).join(", ");
    return { cid, sn, sr, dif: sn-sr, cli: r0?.cliente||"—", nfs };
  });
  const foraRegra = lotes.filter(l => Math.abs(l.dif) > 1.005).sort((a,b)=>Math.abs(b.dif)-Math.abs(a.dif));

  // 2) Receitas duplicadas — mesmo racional (empresa+tipo+PEP+profissional+competência+valor).
  const grp = {};
  records.forEach(r => { const k=[r.empresa,r.tipo,r.pep,r.profissional,r.competencia,r.valorTotal].join("|"); (grp[k]=grp[k]||[]).push(r); });
  const dups = Object.values(grp).filter(g => g.length>1);

  // 3) Faturado (total ou parcial) sem nota amarrada — o caso do Erik.
  const cidsComNota = new Set(notes.filter(n=>n.conciliacaoId).map(n=>n.conciliacaoId));
  const cidsByRec = {}; faturamentos.forEach(a=>{ if(a.conciliacaoId){ (cidsByRec[a.recordId]=cidsByRec[a.recordId]||new Set()).add(a.conciliacaoId); } });
  const semNota = records.filter(r => {
    const st = recStatus(r, fatByRec[r.id], bill(r));
    if (st!=="Faturado" && st!=="Faturado parcial") return false;
    const set = cidsByRec[r.id];
    const viaAloc = set && [...set].some(c=>cidsComNota.has(c));
    const viaLegado = r.municipalNoteId && notes.some(n=>n.id===r.municipalNoteId);
    return !viaAloc && !viaLegado;
  });

  // Conferência de mão dupla (o problema #5): faturado ↔ reconhecido.
  const totExc = foraRegra.length + semNota.length;
  const exportar = () => {
    const rows = [
      ...foraRegra.map(l=>["Divergência de conciliação", l.cli, `NF ${l.nfs||"—"}`, "", brl(l.sn), brl(l.sr), brl(l.dif)]),
      ...semNota.map(r=>["Faturado sem nota", r.cliente, `${r.profissional||r.pep} · ${r.competencia} · ${r.empresa}`, recStatus(r,fatByRec[r.id],bill(r)), "", "", brl(bill(r))]),
      ...dups.map(g=>["Receita duplicada", g[0].cliente, `${g[0].profissional||g[0].pep} · ${g[0].competencia} · ${g[0].tipo} · ${g[0].empresa}`, `${g.length}×`, "", "", brl(g[0].valorTotal)]),
    ];
    downloadCSV("FCamara_Conferencia.csv", ["Exceção","Cliente","Detalhe","Status/Qtd","Notas","Receitas","Valor/Diferença"], rows);
  };

  const Result = ({ id, titulo, desc, problemas, children }) => {
    const ok = problemas===0; const isOpen = open===id;
    return (
      <Card style={{padding:0,overflow:"hidden",marginBottom:14,border:`1px solid ${ok?T.okLine:T.dangerLine}`}}>
        <div onClick={()=>!ok&&setOpen(isOpen?"":id)} style={{display:"flex",alignItems:"center",gap:12,padding:"14px 16px",cursor:ok?"default":"pointer",background:ok?"#f0fdf4":"#fef2f2"}}>
          <div style={{fontSize:22}}>{ok?"✅":"⚠️"}</div>
          <div style={{flex:1}}>
            <div style={{fontWeight:700,fontSize:14,color:T.ink}}>{titulo}</div>
            <div style={{fontSize:12,color:T.muted,marginTop:2}}>{desc}</div>
          </div>
          <div style={{fontSize:14,fontWeight:800,color:ok?T.ok:T.danger}}>{ok?"Tudo certo":`${problemas} a revisar`}</div>
          {!ok && <div style={{fontSize:12,color:T.muted,width:14,textAlign:"center"}}>{isOpen?"▲":"▼"}</div>}
        </div>
        {!ok && isOpen && <div style={{borderTop:`1px solid ${T.line}`,maxHeight:420,overflowY:"auto"}}>{children}</div>}
      </Card>
    );
  };
  const Row = ({ children }) => <div style={{padding:"8px 16px",borderBottom:`1px solid ${T.lineSoft}`,fontSize:12.5,display:"flex",gap:12,flexWrap:"wrap",alignItems:"center"}}>{children}</div>;

  return (
    <div>
      <PageHead icon="check" title="Validações do sistema" sub="confere se o faturamento está batendo"/>
      <Card style={{padding:16,marginBottom:16,border:`1px solid ${totExc?T.dangerLine:T.okLine}`,background:totExc?T.dangerBg:T.okBg}}>
        <div style={{display:"flex",alignItems:"center",gap:14,flexWrap:"wrap"}}>
          <div style={{fontSize:26}}>{totExc?"⚠️":"🛡️"}</div>
          <div style={{flex:1,minWidth:220}}>
            <div style={{fontWeight:800,fontSize:15,color:T.ink}}>Conferência faturado ↔ reconhecido</div>
            <div style={{fontSize:12.5,color:T.inkSoft,marginTop:2}}>Garante os dois lados: toda receita faturada tem nota amarrada, e nota × receita batem por conciliação. {totExc?`${totExc} exceção(ões) a revisar.`:"Sem exceções — 100% reconciliado."}</div>
          </div>
          <div style={{fontSize:26,fontWeight:800,color:totExc?T.danger:T.ok,minWidth:32,textAlign:"center"}}>{totExc||"✓"}</div>
          <Btn icon="download" onClick={exportar}>Exportar exceções</Btn>
        </div>
      </Card>
      <div style={{fontSize:12.5,color:T.muted,marginBottom:16}}>Rodam sobre todos os dados carregados. Verde = ok; laranja = clique para ver o que revisar.</div>

      <Result id="lotes" titulo="Conciliado da prefeitura × receita conciliada (regra de R$ 1,00)"
        desc="Em cada conciliação, a soma das notas deve bater com a soma das receitas — diferença máxima de R$ 1,00."
        problemas={foraRegra.length}>
        {foraRegra.map(l=>(
          <Row key={l.cid}>
            <Badge label={`dif. ${brl(l.dif)}`} color="red" small/>
            <b style={{color:T.ink}}>{l.cli}</b>
            <span style={{color:T.muted}}>NF {l.nfs||"—"}</span>
            <div style={{flex:1}}/>
            <span>Notas <b style={{color:T.ink}}>{brl(l.sn)}</b> ↔ Receitas <b style={{color:T.ink}}>{brl(l.sr)}</b></span>
          </Row>
        ))}
      </Result>

      <Result id="dups" titulo="Receitas possivelmente duplicadas"
        desc="Registros com mesmo tipo, PEP, profissional, competência e valor — pode ser duplicidade de importação/racional."
        problemas={dups.length}>
        {dups.map((g,i)=>(
          <Row key={i}>
            <Badge label={`${g.length}×`} color="orange" small/>
            <b style={{color:T.ink}}>{g[0].cliente}</b>
            <span style={{color:T.muted}}>{g[0].profissional||g[0].pep} · {g[0].competencia} · {g[0].tipo} · {g[0].empresa}</span>
            <div style={{flex:1}}/>
            <span><b style={{color:T.ink}}>{brl(g[0].valorTotal)}</b> cada</span>
          </Row>
        ))}
      </Result>

      <Result id="semnota" titulo="Faturado sem nota amarrada"
        desc="Receitas com status Faturado/parcial mas sem nota conciliada — o tipo de caso que gerou o problema do Erik."
        problemas={semNota.length}>
        {semNota.map(r=>(
          <Row key={r.id}>
            <Badge label={recStatus(r, fatByRec[r.id], bill(r))} color="orange" small/>
            <b style={{color:T.ink}}>{r.cliente}</b>
            <span style={{color:T.muted}}>{r.profissional||r.pep} · {r.competencia} · {r.empresa}</span>
            <div style={{flex:1}}/>
            <span><b style={{color:T.ink}}>{brl(bill(r))}</b></span>
          </Row>
        ))}
      </Result>
    </div>
  );
}

// Visão por projeto — linha do tempo. Escolhe um cliente e vê cada PEP mês a mês,
// com o faturável e a distribuição por etapa (faturado / liberado / em andamento /
// não iniciado). Só renderiza ao escolher o cliente — mapa focado.
function ProjectTimelineView({ records, clients, fatByRec={}, varByRec={} }) {
  const [cliente, setCliente] = useState("");
  const [qProf, setQProf] = useState("");
  const [tipoF, setTipoF] = useState("todos");
  const [empF, setEmpF] = useState("todas");
  const [statusF, setStatusF] = useState("todos");  // faturado | represado | ciclo
  const [perSel, setPerSel] = useState([]);         // competências selecionadas (vazio = todas)
  const [detail, setDetail] = useState(null);       // {label, ids:Set} da célula clicada
  const bill = (r) => (r.valorTotal||0) + (varByRec[r.id]||0);
  const fat  = (r) => fatByRec[r.id]||0;
  // Status do registro: faturado (nada em aberto), ou (em aberto) represado / dentro do ciclo.
  const statusOf = (r) => { const saldo=bill(r)-fat(r); if (saldo<=0.01) return "faturado"; return categoriaOf(r,clients).cat==="represado" ? "represado" : "ciclo"; };
  // Valor represado de uma lista: saldo em aberto dos registros já represados.
  const represSaldo = (list) => list.reduce((s,r)=>{ const saldo=bill(r)-fat(r); return s + ((saldo>0.01 && categoriaOf(r,clients).cat==="represado") ? saldo : 0); }, 0);
  const empNome = (cod) => EMPRESAS.find(e=>e.cod===cod)?.nome || "";

  // Cliente canônico por NOME (o cod_cliente é manual e não confiável — o mesmo
  // código aparece em clientes diferentes e o mesmo cliente tem dezenas de
  // códigos). Normaliza (maiúsculas/espaços/pontuação final) e FUNDE truncamentos:
  // quando um nome (≥12 chars) é prefixo de outro, é o mesmo cliente. Ex.:
  // "JOHNSON & JOHNSON" ⊂ "JOHNSON & JOHNSON DO BRASIL INDUSTRIA E" → um cliente.
  // Mantém nomes distintos separados (Anglo × MRV). Não altera o dado gravado.
  const _nn = s => String(s||"").toUpperCase().replace(/\s+/g," ").trim().replace(/[^A-Z0-9)]+$/,"");
  const _freq = {};   // nome normalizado -> { nome original -> contagem }
  records.forEach(r => { const k=_nn(r.cliente); if(!k) return; const m=(_freq[k]=_freq[k]||{}); const o=(r.cliente||"").trim(); if(o) m[o]=(m[o]||0)+1; });
  const _names = Object.keys(_freq).sort((a,b)=>a.length-b.length);
  const _parent = {}; _names.forEach(n=>_parent[n]=n);
  const _find = x => { while(_parent[x]!==x){ _parent[x]=_parent[_parent[x]]; x=_parent[x]; } return x; };
  for (let i=0;i<_names.length;i++){ const a=_names[i]; if(a.length<12) continue;
    for (let j=i+1;j<_names.length;j++){ const b=_names[j]; if(b.startsWith(a)){ const ra=_find(a), rb=_find(b); if(ra!==rb) _parent[ra]=rb; } } }
  // nome canônico do grupo: o original mais frequente (desempate: mais longo).
  const _tally = {};
  _names.forEach(n=>{ const root=_find(n); const t=(_tally[root]=_tally[root]||{}); Object.entries(_freq[n]).forEach(([o,c])=>t[o]=(t[o]||0)+c); });
  const _canonByRoot = {};
  Object.entries(_tally).forEach(([root,t])=>{ _canonByRoot[root]=Object.entries(t).sort((a,b)=>b[1]-a[1]||b[0].length-a[0].length)[0][0]; });
  const cliNome = r => { const k=_nn(r.cliente); return (k && _parent[k]!=null) ? _canonByRoot[_find(k)] : (r.cliente||""); };

  const empresasAll = [...new Set(records.map(r=>r.empresa).filter(Boolean))].sort();
  // Empresa é um filtro de topo (sempre visível): restringe a lista de clientes
  // e o mapa. "todas" = grupo inteiro.
  const recsEmp = empF==="todas" ? records : records.filter(r=>r.empresa===empF);
  const clientesList = [...new Set(recsEmp.map(r=>cliNome(r)).filter(Boolean))].sort((a,b)=>a.localeCompare(b));

  let recs = cliente ? recsEmp.filter(r=>cliNome(r)===cliente) : [];
  const tiposCli = [...new Set(recs.map(r=>r.tipo).filter(Boolean))].sort();
  const mesesOpts = [...new Set(recs.map(r=>r.competencia).filter(Boolean))].sort((a,b)=>compRank(a).localeCompare(compRank(b)));
  if (tipoF!=="todos") recs = recs.filter(r=>r.tipo===tipoF);
  if (qProf.trim()) { const s=qProf.trim().toLowerCase(); recs = recs.filter(r=>(r.profissional||"").toLowerCase().includes(s)); }
  if (statusF!=="todos") recs = recs.filter(r=>statusOf(r)===statusF);
  if (perSel.length) recs = recs.filter(r=>perSel.includes(r.competencia));

  const meses = [...new Set(recs.map(r=>r.competencia).filter(Boolean))].sort((a,b)=>compRank(a).localeCompare(compRank(b)));
  const projMap = {};
  recs.forEach(r=>{ const pb=pepBase(r.pep); const k=`${r.tipo}||${pb}`; (projMap[k]=projMap[k]||{tipo:r.tipo,pep:pb,recs:[]}).recs.push(r); });
  const projetos = Object.values(projMap).map(p=>({...p, tot:p.recs.reduce((s,r)=>s+bill(r),0)})).sort((a,b)=>b.tot-a.tot);

  const SEG = [
    { key:"faturado",  label:"Faturado",     color:C.green.solid },
    { key:"liberado",  label:"Liberado",     color:C.teal.solid },
    { key:"andamento", label:"Em andamento", color:C.orange.solid },
    { key:"pendente",  label:"Não iniciado", color:C.gray.solid },
  ];
  const segsOf = (list) => {
    let faturado=0, liberado=0, andamento=0, pendente=0;
    list.forEach(r=>{
      const f=fat(r), saldo=bill(r)-f;
      faturado += Math.max(0,f);
      if (saldo>0.01){ const p=r.progress||{};
        if (p.p5_liberado) liberado+=saldo;
        else if (calcStatus(p)==="Não iniciado") pendente+=saldo;
        else andamento+=saldo;
      }
    });
    return { faturado, liberado, andamento, pendente, total:faturado+liberado+andamento+pendente };
  };

  const Cell = ({ list, label }) => {
    const s = list.length ? segsOf(list) : null;
    if (!s || s.total<=0.01) return <div style={{color:T.faint,fontSize:11,textAlign:"center"}}>—</div>;
    const pctFat = Math.round(s.faturado/s.total*100);
    const repres = list.filter(r=>bill(r)-fat(r)>0.01 && categoriaOf(r,clients).cat==="represado").length;
    return (
      <div onClick={()=>setDetail({label, ids:new Set(list.map(r=>r.id))})} style={{cursor:"pointer"}} title="Ver detalhe / classificar">
        <div style={{display:"flex",alignItems:"center",gap:5,lineHeight:1}}>
          <div style={{fontSize:11.5,fontWeight:700,color:T.ink,fontVariantNumeric:"tabular-nums"}}>{fmtShort(s.total)}</div>
          {repres>0 && <span title={`${repres} represado(s)`} style={{fontSize:9,fontWeight:700,color:C.red.solid}}>⚠{repres}</span>}
        </div>
        <div style={{display:"flex",height:8,borderRadius:4,overflow:"hidden",margin:"5px 0 3px",background:T.lineSoft}}>
          {SEG.map(g=> s[g.key]>0.01 ? <div key={g.key} title={`${g.label}: ${brl(s[g.key])}`} style={{width:`${s[g.key]/s.total*100}%`,background:g.color}}/> : null)}
        </div>
        <div style={{fontSize:10,fontWeight:600}}><span style={{color:pctFat>=100?C.green.solid:T.muted}}>{pctFat}% fat</span> · <span style={{color:C.orange.solid}}>{100-pctFat}% aberto</span></div>
      </div>
    );
  };

  const detailList = detail ? recs.filter(r=>detail.ids.has(r.id)).sort((a,b)=>bill(b)-bill(a)) : [];
  // Totalizadores da célula clicada: faturado vs. em aberto (represado / dentro do ciclo).
  const dFat   = detailList.reduce((s,r)=>s+Math.max(0,fat(r)),0);
  const dRep   = detailList.reduce((s,r)=>{ const a=bill(r)-fat(r); return s + (a>0.01 && categoriaOf(r,clients).cat==="represado" ? a : 0); },0);
  const dCiclo = detailList.reduce((s,r)=>{ const a=bill(r)-fat(r); return s + (a>0.01 && categoriaOf(r,clients).cat!=="represado" ? a : 0); },0);
  const dTot   = dFat + dRep + dCiclo;
  const totBox = (label,value,color,strong)=>(
    <div style={{textAlign:"right",minWidth:96}}>
      <div style={{fontSize:10,color:T.muted,textTransform:"uppercase",letterSpacing:".04em",fontWeight:600}}>{label}</div>
      <div style={{fontSize:strong?16:14,fontWeight:strong?800:700,color,fontVariantNumeric:"tabular-nums"}}>{brl(value)}</div>
    </div>
  );

  const thProj = { position:"sticky", left:0, zIndex:2, background:T.canvas, textAlign:"left", padding:"10px 12px", fontSize:11, textTransform:"uppercase", letterSpacing:".05em", color:T.muted, borderBottom:`1px solid ${T.line}`, minWidth:190 };
  const thMes  = { padding:"10px 8px", fontSize:12, fontWeight:700, color:T.ink, borderBottom:`1px solid ${T.line}`, borderLeft:`1px solid ${T.lineSoft}`, whiteSpace:"nowrap", textAlign:"center", minWidth:100 };
  const tdProj = { position:"sticky", left:0, zIndex:1, background:"var(--surface)", padding:"8px 12px", borderBottom:`1px solid ${T.lineSoft}`, minWidth:190 };
  const tdCell = { padding:"8px 10px", borderBottom:`1px solid ${T.lineSoft}`, borderLeft:`1px solid ${T.lineSoft}`, verticalAlign:"top" };

  return (
    <div>
      <PageHead icon="chart" title="Visão por projeto" sub="linha do tempo — faturável por PEP, mês a mês"/>
      <div style={{fontSize:12.5,color:T.muted,marginBottom:16}}>Escolha um cliente para desenhar o mapa. Cada célula mostra o faturável do mês e a distribuição por etapa.</div>

      <Card style={{padding:"12px 14px",marginBottom:16}}>
        <div style={{display:"flex",gap:10,flexWrap:"wrap",alignItems:"center"}}>
          {empresasAll.length>1 && <select style={{...inp,width:"auto",minWidth:170}} value={empF} onChange={e=>{setEmpF(e.target.value);setCliente("");setTipoF("todos");setPerSel([]);}}>
            <option value="todas">Todas as empresas</option>
            {empresasAll.map(cod=><option key={cod} value={cod}>{cod}{empNome(cod)?` — ${empNome(cod)}`:""}</option>)}
          </select>}
          <select style={{...inp,width:"auto",minWidth:240,flex:"1 1 240px"}} value={cliente} onChange={e=>{setCliente(e.target.value);setTipoF("todos");setPerSel([]);}}>
            <option value="">— Selecione um cliente —</option>
            {clientesList.map(c=><option key={c} value={c}>{c}</option>)}
          </select>
          {cliente && <>
            <select style={{...inp,width:"auto"}} value={statusF} onChange={e=>setStatusF(e.target.value)}><option value="todos">Todos os status</option><option value="faturado">Faturado</option><option value="represado">Represado</option><option value="ciclo">Dentro do ciclo</option></select>
            <select style={{...inp,width:"auto"}} value={tipoF} onChange={e=>setTipoF(e.target.value)}><option value="todos">Todos os contratos</option>{tiposCli.map(t=><option key={t} value={t}>{t}</option>)}</select>
            <input style={{...inp,width:"auto",minWidth:180,flex:"1 1 180px"}} placeholder="Profissional dentro do cliente…" value={qProf} onChange={e=>setQProf(e.target.value)}/>
          </>}
        </div>
        {cliente && mesesOpts.length>0 && <div style={{display:"flex",gap:6,alignItems:"center",flexWrap:"wrap",marginTop:11,paddingTop:11,borderTop:`1px solid ${T.lineSoft}`}}>
          <span style={{fontSize:11,color:T.muted,fontWeight:700,textTransform:"uppercase",letterSpacing:".04em",marginRight:2}}>Período</span>
          {mesesOpts.map(m=>{ const on=perSel.includes(m); return (
            <button key={m} onClick={()=>setPerSel(on?perSel.filter(x=>x!==m):[...perSel,m])}
              style={{padding:"5px 11px",borderRadius:T.rPill,fontSize:12,fontWeight:600,cursor:"pointer",border:`1px solid ${on?T.brand:T.line}`,background:on?"#fff5f1":"var(--surface)",color:on?T.brand:T.inkSoft}}>{m}</button>
          );})}
          {perSel.length>0
            ? <button onClick={()=>setPerSel([])} style={{padding:"5px 10px",borderRadius:T.rPill,fontSize:12,fontWeight:600,cursor:"pointer",border:`1px solid ${T.line}`,background:"var(--surface)",color:T.muted}}>limpar</button>
            : <span style={{fontSize:11,color:T.faint}}>todos os meses</span>}
        </div>}
      </Card>

      {!cliente ? (
        <Card style={{textAlign:"center",padding:"3rem"}}>
          <div style={{fontSize:14,color:T.muted}}>Escolha um cliente acima para ver o mapa de faturamento por projeto.</div>
        </Card>
      ) : projetos.length===0 ? (
        <Card style={{textAlign:"center",padding:"2rem"}}><div style={{fontSize:13,color:T.muted}}>Nenhum registro para este filtro.</div></Card>
      ) : (
        <>
          <div style={{display:"flex",gap:16,flexWrap:"wrap",marginBottom:12,fontSize:11.5,color:T.muted}}>
            {SEG.map(g=><span key={g.key} style={{display:"flex",alignItems:"center",gap:6}}><span style={{width:13,height:13,borderRadius:3,background:g.color}}/>{g.label}</span>)}
          </div>
          <div style={{overflowX:"auto",border:`1px solid ${T.line}`,borderRadius:T.rLg,background:"var(--surface)"}}>
            <table style={{borderCollapse:"collapse",width:"100%"}}>
              <thead><tr>
                <th style={thProj}>Projeto (tipo · PEP)</th>
                {meses.map(m=><th key={m} style={thMes}>{m}</th>)}
                <th style={{...thMes,background:T.canvas,borderLeft:`2px solid ${T.line}`}}>Total projeto</th>
                <th style={{...thMes,background:T.canvas,color:C.red.solid}}>Represado</th>
              </tr></thead>
              <tbody>
                {projetos.map(p=>{ const rep=represSaldo(p.recs); return (
                  <tr key={p.tipo+"|"+p.pep}>
                    <td style={tdProj}>
                      <div style={{fontWeight:700,fontSize:12,color:T.ink}}>{p.pep||"—"}</div>
                      <div style={{fontSize:10.5,color:T.muted}}>{p.tipo}</div>
                    </td>
                    {meses.map(m=><td key={m} style={tdCell}><Cell list={p.recs.filter(r=>r.competencia===m)} label={`${p.pep||p.tipo} · ${m}`}/></td>)}
                    <td style={{...tdCell,background:T.canvas,borderLeft:`2px solid ${T.line}`}}><Cell list={p.recs} label={`${p.pep||p.tipo} · Total`}/></td>
                    <td style={{...tdCell,background:T.canvas,textAlign:"right",fontWeight:700,color:rep>0.01?C.red.solid:T.faint,whiteSpace:"nowrap"}}>{rep>0.01?fmtShort(rep):"—"}</td>
                  </tr>
                );})}
                <tr>
                  <td style={{...tdProj,background:T.canvas,fontWeight:800,color:T.ink,fontSize:12}}>TOTAL · {cliente}</td>
                  {meses.map(m=><td key={m} style={{...tdCell,background:T.canvas}}><Cell list={recs.filter(r=>r.competencia===m)} label={`${cliente} · ${m}`}/></td>)}
                  <td style={{...tdCell,background:T.canvas,borderLeft:`2px solid ${T.line}`}}><Cell list={recs} label={`${cliente} · Total`}/></td>
                  <td style={{...tdCell,background:T.canvas,textAlign:"right",fontWeight:800,color:represSaldo(recs)>0.01?C.red.solid:T.faint,whiteSpace:"nowrap"}}>{represSaldo(recs)>0.01?fmtShort(represSaldo(recs)):"—"}</td>
                </tr>
              </tbody>
            </table>
          </div>
        </>
      )}

      {detail && <Modal title={`Detalhe — ${detail.label}`} subtitle={`${detailList.length} registro(s) · status e classificação (edição na aba Represados)`} wide onClose={()=>setDetail(null)}>
        <div className="fc-scroll" style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
            <thead><tr style={{background:T.canvas}}>
              {["Profissional","Período","Faturável","Faturado","Status","Classificação"].map(h=>
                <th key={h} style={{padding:"7px 10px",textAlign:"left",borderBottom:`1px solid ${T.line}`,fontWeight:600,color:T.muted,whiteSpace:"nowrap"}}>{h}</th>)}
            </tr></thead>
            <tbody>
              {detailList.map(r=>{
                const f=fat(r), saldo=bill(r)-f, aberto=saldo>0.01;
                return (
                  <tr key={r.id} style={{borderBottom:`1px solid ${T.lineSoft}`}}>
                    <td style={{padding:"7px 10px",fontWeight:500,color:T.ink}}>{r.profissional||r.pep||"—"}</td>
                    <td style={{padding:"7px 10px",color:T.muted,whiteSpace:"nowrap"}}>{r.inicio||"—"} → {r.fim||"—"}</td>
                    <td style={{padding:"7px 10px",fontWeight:500,whiteSpace:"nowrap"}}>{fmtShort(bill(r))}</td>
                    <td style={{padding:"7px 10px",whiteSpace:"nowrap",color:f>0.01?C.green.solid:T.faint}}>{f>0.01?fmtShort(f):"—"}</td>
                    <td style={{padding:"7px 10px"}}><PipelineStepper states={recordStates(r.progress, r.tipo)} groups={funnelGroups(r.tipo)} size="sm"/></td>
                    <td style={{padding:"7px 10px",whiteSpace:"nowrap"}}>
                      {aberto ? (<ClassifyChip record={r} clients={clients} readOnly style={{marginLeft:0}}/> || <span style={{fontSize:11,color:T.faint}}>—</span>)
                              : <span style={{fontSize:11,color:T.faint}}>faturado</span>}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
        <div style={{display:"flex",gap:20,flexWrap:"wrap",justifyContent:"flex-end",alignItems:"flex-end",marginTop:14,padding:"12px 16px",background:T.canvas,borderRadius:T.rLg,border:`1px solid ${T.line}`}}>
          {totBox("Faturado", dFat, dFat>0.01?C.green.solid:T.faint)}
          {totBox("Represado", dRep, dRep>0.01?C.red.solid:T.faint)}
          {dCiclo>0.01 && totBox("Dentro do ciclo", dCiclo, C.orange.solid)}
          <div style={{width:1,alignSelf:"stretch",background:T.line}}/>
          {totBox("Total", dTot, T.ink, true)}
        </div>
        <div style={{fontSize:11,color:T.muted,marginTop:12}}>Para classificar ou editar o motivo do represamento, use a aba <b>Represados</b>.</div>
      </Modal>}
    </div>
  );
}

// Represados — receitas ainda em aberto que passaram da folga de faturamento.
// É o ÚNICO lugar onde se edita a classificação (motivo) e a observação do
// represamento. Filtra por empresa e cliente; agrupa por cliente.
function RepresadosView({ records, clients, fatByRec={}, varByRec={}, onSaveClass, isViewer=false }) {
  const [empF, setEmpF] = useState("todas");
  const [cliF, setCliF] = useState("todos");
  const [q, setQ] = useState("");
  const [onlyPend, setOnlyPend] = useState(false);   // só sem classificação
  const [classTarget, setClass] = useState(null);
  const bill = (r) => (r.valorTotal||0)+(varByRec[r.id]||0);
  const fat  = (r) => fatByRec[r.id]||0;
  const rep  = (r) => bill(r)-fat(r);
  const empNome = cod => EMPRESAS.find(e=>e.cod===cod)?.nome||"";

  // Base: saldo em aberto E categoria represado (pela data).
  const base = records.filter(r => rep(r)>0.01 && categoriaOf(r,clients).cat==="represado");
  const empresas = [...new Set(base.map(r=>r.empresa).filter(Boolean))].sort();
  const clientes = [...new Set(base.filter(r=>empF==="todas"||r.empresa===empF).map(r=>r.cliente).filter(Boolean))].sort();

  let list = base;
  if (empF!=="todas") list = list.filter(r=>r.empresa===empF);
  if (cliF!=="todos") list = list.filter(r=>r.cliente===cliF);
  if (q.trim()){ const s=q.trim().toLowerCase(); list=list.filter(r=>[r.cliente,r.pep,r.profissional,r.classMotivo,r.classObs].some(v=>String(v||"").toLowerCase().includes(s))); }
  if (onlyPend) list = list.filter(r=>!(r.classMotivo||r.classObs));

  const totalBase = base.reduce((s,r)=>s+rep(r),0);
  const semClass  = base.filter(r=>!(r.classMotivo||r.classObs)).length;
  const totalList = list.reduce((s,r)=>s+rep(r),0);

  const grouped = {};
  list.forEach(r=>{ (grouped[r.cliente]=grouped[r.cliente]||[]).push(r); });
  const groups = Object.entries(grouped).map(([cli,recs])=>({cli, recs, tot:recs.reduce((s,r)=>s+rep(r),0)})).sort((a,b)=>b.tot-a.tot);

  // Exporta o que está filtrado (mesma lista da tela), já formatado em .xlsx.
  function exportXLSX() {
    const headers = ["Empresa","Cliente","PEP","Profissional","Tipo","Compet. faturamento","Represado (R$)","Motivo","Observação"];
    const rows = list
      .slice().sort((a,b)=>rep(b)-rep(a))
      .map(r=>{ const { compFat } = categoriaOf(r,clients); return [
        r.empresa||"", r.cliente||"", r.pep||"", r.profissional||"", r.tipo||"",
        compFat||"", Number(rep(r).toFixed(2)), r.classMotivo||"", r.classObs||"",
      ]; });
    const stamp = new Date().toISOString().slice(0,10);
    downloadXLSX(`Represados_${rows.length}_${stamp}.xlsx`, headers, rows);
  }

  return (
    <div>
      {classTarget && <ClassifyModal record={classTarget} clients={clients} fatByRec={fatByRec} varByRec={varByRec} onSave={onSaveClass} onClose={()=>setClass(null)}/>}
      <PageHead icon="alert" title="Represados" sub={`${base.length} receita(s) · ${brl(totalBase)} represado · ${semClass} sem classificação`}
        right={<Btn primary icon="download" disabled={!list.length} onClick={exportXLSX}>Exportar Excel</Btn>}/>
      <div style={{fontSize:12.5,color:T.muted,marginBottom:16}}>Único lugar para registrar/editar o motivo do represamento e a observação. Represado = receita ainda em aberto que passou da folga de 1 mês sobre a competência de faturamento (respeita período quebrado).</div>

      <Card style={{padding:"12px 14px",marginBottom:16}}>
        <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
          <select style={{...inp,width:"auto",minWidth:170}} value={empF} onChange={e=>{setEmpF(e.target.value);setCliF("todos");}}>
            <option value="todas">Todas as empresas</option>
            {empresas.map(cod=><option key={cod} value={cod}>{cod}{empNome(cod)?` — ${empNome(cod)}`:""}</option>)}
          </select>
          <select style={{...inp,width:"auto",minWidth:200,flex:"1 1 200px"}} value={cliF} onChange={e=>setCliF(e.target.value)}>
            <option value="todos">Todos os clientes</option>
            {clientes.map(c=><option key={c} value={c}>{c}</option>)}
          </select>
          <input style={{...inp,width:"auto",minWidth:160,flex:"1 1 160px"}} placeholder="Buscar PEP, profissional, motivo…" value={q} onChange={e=>setQ(e.target.value)}/>
          <label style={{display:"flex",alignItems:"center",gap:6,fontSize:12,color:T.inkSoft,fontWeight:600,cursor:"pointer",whiteSpace:"nowrap"}}>
            <input type="checkbox" checked={onlyPend} onChange={e=>setOnlyPend(e.target.checked)}/> só sem classificação
          </label>
        </div>
      </Card>

      {groups.length===0
        ? <Card style={{textAlign:"center",padding:"3rem"}}><div style={{fontSize:14,color:T.muted}}>Nenhum represado para os filtros selecionados.</div></Card>
        : <>
          <div style={{fontSize:13,marginBottom:12,color:T.inkSoft}}>Mostrando <b>{list.length}</b> receita(s) · total represado <b style={{color:C.red.solid}}>{brl(totalList)}</b></div>
          {groups.map(g=>(
            <Card key={g.cli} style={{marginBottom:10,overflow:"hidden"}}>
              <div style={{display:"flex",alignItems:"center",gap:10,padding:"11px 16px",background:T.canvas,borderBottom:`1px solid ${T.lineSoft}`,flexWrap:"wrap"}}>
                <span style={{fontWeight:700,fontSize:13,color:T.ink}}>{g.cli}</span>
                <span style={{fontSize:11,color:T.muted}}>{g.recs.length} receita(s)</span>
                <div style={{flex:1}}/>
                <span style={{fontSize:13,fontWeight:800,color:C.red.solid}}>{brl(g.tot)}</span>
              </div>
              <div className="fc-scroll" style={{overflowX:"auto"}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                  <thead><tr>
                    {["Empresa","PEP · Profissional","Compet. fat.","Represado","Motivo","Observação",""].map(h=>
                      <th key={h} style={{padding:"7px 10px",textAlign:h==="Represado"?"right":"left",borderBottom:`1px solid ${T.line}`,fontWeight:600,color:T.muted,whiteSpace:"nowrap"}}>{h}</th>)}
                  </tr></thead>
                  <tbody>
                    {g.recs.sort((a,b)=>rep(b)-rep(a)).map(r=>{
                      const { compFat } = categoriaOf(r,clients);
                      return (
                      <tr key={r.id} style={{borderBottom:`1px solid ${T.lineSoft}`}}>
                        <td style={{padding:"7px 10px"}}><Badge label={r.empresa||"—"} color="gray" small/></td>
                        <td style={{padding:"7px 10px"}}><div style={{fontWeight:600,color:T.ink}}>{r.profissional||"—"}</div><div style={{fontSize:10.5,color:T.muted}}>{r.pep||"—"} · {r.tipo}</div></td>
                        <td style={{padding:"7px 10px",color:T.inkSoft,whiteSpace:"nowrap"}}>{compFat||"—"}</td>
                        <td style={{padding:"7px 10px",textAlign:"right",fontWeight:700,color:C.red.solid,whiteSpace:"nowrap"}}>{brl(rep(r))}</td>
                        <td style={{padding:"7px 10px",color:r.classMotivo?T.ink:T.faint,maxWidth:180,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{r.classMotivo||"— a classificar"}</td>
                        <td style={{padding:"7px 10px",color:r.classObs?T.inkSoft:T.faint,maxWidth:220,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={r.classObs||""}>{r.classObs||"—"}</td>
                        <td style={{padding:"7px 10px",textAlign:"right",whiteSpace:"nowrap"}}>
                          {!isViewer && <Btn small icon="pencil" onClick={()=>setClass(r)}>{(r.classMotivo||r.classObs)?"Editar":"Classificar"}</Btn>}
                        </td>
                      </tr>
                    );})}
                  </tbody>
                </table>
              </div>
            </Card>
          ))}
        </>}
    </div>
  );
}

// Conciliação (estilo conciliação bancária), por empresa do grupo (BR02, BR04…).
// De um lado as notas da prefeitura a conciliar; do outro as receitas. Filtros e
// ordenação independentes nos dois lados. Nada é conciliado automaticamente.
function ConciliationView({ records, clients, notes, isAdmin, isViewer=false, fatByRec={}, varByRec={}, faturamentos=[], orfas=0, onReopenOrphans, onImport, onUndoImport, onDeleteNote, onConciliate, onReopen }) {
  const [importing, setImporting] = useState(false);
  const [manage, setManage] = useState(false);
  const [noteDel, setNoteDel] = useState(null);
  const [empresa, setEmpresa] = useState("");
  const [selNotes, setSelNotes] = useState(() => new Set());
  const [selRecs, setSelRecs] = useState(() => new Set());
  const [valores, setValores] = useState({});   // valor a faturar por registro (parcial)
  const [expNote, setExpNote] = useState("");
  const [showConf, setShowConf] = useState(false);   // painel de conferência de lotes
  const [confSoDif, setConfSoDif] = useState(true);  // só lotes que não batem
  const toast = useToast();
  // filtros e ordenação — lado esquerdo (notas)
  const [qNote, setQNote] = useState(""); const [noteStat, setNoteStat] = useState("pendentes"); const [noteSort, setNoteSort] = useState("valor_desc"); const [noteDe, setNoteDe] = useState(""); const [noteAte, setNoteAte] = useState(""); const [noteCli, setNoteCli] = useState("todos");
  // filtros e ordenação — lado direito (receitas)
  const [qRec, setQRec] = useState(""); const [recStat, setRecStat] = useState("pendentes"); const [recSort, setRecSort] = useState("valor_desc"); const [recComp, setRecComp] = useState("todas");
  const [recDim, setRecDim] = useState("servico");   // competência: mês de serviço × ciclo de faturamento
  // Incluir receitas ainda não "Liberadas para faturamento" (ex.: conciliação
  // adiantada antes do time preencher o passo a passo). Conciliar já libera o funil.
  const [incluirNaoLib, setIncluirNaoLib] = useState(true);
  const podeFaturar = (r) => incluirNaoLib || r.progress?.p5_liberado;

  const empresasComDados = EMPRESAS.filter(e => notes.some(n=>n.empresa===e.cod) || records.some(r=>r.empresa===e.cod));
  const empNotes = notes.filter(n => n.empresa===empresa && !n.cancelada);
  const empRecs  = records.filter(r => r.empresa===empresa);

  // Nota já conciliada? (novo modelo: conciliacao_id; compat: municipal_note_id antigo)
  const notaConc = (n) => !!n.conciliacaoId || records.some(r => r.municipalNoteId === n.id);
  // Faturamento parcial: quanto já foi faturado e quanto falta (saldo).
  const fat    = (r) => fatByRec[r.id] || 0;
  // Faturável = receita + variação pós-fechamento (ex.: Casas Bahia). A variação
  // vira saldo a conciliar (NF própria) sem alterar a receita original.
  const bill   = (r) => (r.valorTotal||0) + (varByRec[r.id]||0);
  // Saldo COM sinal: descontos (valor negativo) também precisam ser conciliados
  // para fechar o total da nota. Ex.: +10.500 serviço − 500 desconto = 10.000.
  const saldoR = (r) => bill(r) - fat(r);
  const hasSaldo = (r) => Math.abs(saldoR(r)) > 0.01;
  const hasFat   = (r) => Math.abs(fat(r)) > 0.001;
  // valor a faturar do registro (padrão = saldo; digitável; nunca passa do saldo,
  // respeitando o sinal — negativo não pode ficar mais negativo que o saldo).
  const valorDe = (r) => {
    const v = valores[r.id]; const s = saldoR(r);
    if (v===undefined || v==="") return s;
    let x = parseBR(v);
    return s >= 0 ? Math.max(0, Math.min(x, s)) : Math.min(0, Math.max(x, s));
  };

  const compKey = (c) => { const [m,y]=String(c||"").split("/"); return (y||"0000")+(m||"00"); };
  const sortNotes = (a,b) => ({ valor_desc:b.valorServicos-a.valorServicos, valor_asc:a.valorServicos-b.valorServicos,
    data_desc:String(b.emitidaEm||"").localeCompare(String(a.emitidaEm||"")), data_asc:String(a.emitidaEm||"").localeCompare(String(b.emitidaEm||"")),
    tomador_az:(a.tomadorNome||"").localeCompare(b.tomadorNome||"") }[noteSort] || 0);
  const sortRecs = (a,b) => ({ valor_desc:(b.valorTotal||0)-(a.valorTotal||0), valor_asc:(a.valorTotal||0)-(b.valorTotal||0),
    cliente_az:(a.cliente||"").localeCompare(b.cliente||""), comp:compKey(a.competencia).localeCompare(compKey(b.competencia)) }[recSort] || 0);

  let leftNotes = empNotes.slice();
  if (noteStat==="pendentes") leftNotes = leftNotes.filter(n=>!notaConc(n));
  if (noteStat==="conciliadas") leftNotes = leftNotes.filter(n=>notaConc(n));
  // Intervalo de data de emissão (de/até) — emitidaEm é ISO (aaaa-mm-dd), então
  // a comparação de string já ordena por data. Cada limite é opcional.
  if (noteDe)  leftNotes = leftNotes.filter(n=>{ const d=String(n.emitidaEm||"").slice(0,10); return d && d>=noteDe; });
  if (noteAte) leftNotes = leftNotes.filter(n=>{ const d=String(n.emitidaEm||"").slice(0,10); return d && d<=noteAte; });
  // Opções do seletor de cliente = tomadores das notas já filtradas por status/dia
  // (pendentes mostra só clientes com nota pendente; conciliadas idem).
  const tomadoresUsados = [...new Set(leftNotes.map(n=>n.tomadorNome).filter(Boolean))].sort((a,b)=>a.localeCompare(b));
  if (noteCli!=="todos") leftNotes = leftNotes.filter(n=>(n.tomadorNome||"")===noteCli);
  if (qNote.trim()) { const s=qNote.trim().toLowerCase(); const dig=s.replace(/\D/g,""); leftNotes = leftNotes.filter(n=>(n.numero||"").toLowerCase().includes(s)||(n.tomadorNome||"").toLowerCase().includes(s)||(!!dig&&(n.pedidos||"").includes(dig))); }
  leftNotes = leftNotes.sort(sortNotes);

  // Período quebrado: dia de corte do cliente → "competência de faturamento" (ciclo).
  // Para cliente com corte D, uma peça cujo dia de início é >= D fatura no mês
  // seguinte (a nota fecha em D do próximo mês); antes de D, fatura no mês atual.
  const diaCorteDe = (r) => diaCorteOf(r, clients);
  const compFat = (r) => compFatOf(r, clients);
  const compValue = (r) => recDim==="ciclo" ? compFat(r) : (r.competencia || "");

  const compsUsadas = [...new Set(empRecs.map(compValue).filter(Boolean))].sort((a,b)=>compKey(b).localeCompare(compKey(a)));
  let rightRecs = empRecs.slice();
  // Só entra na conciliação quem está "Liberado para faturamento" e tem saldo (gate).
  if (recStat==="pendentes") rightRecs = rightRecs.filter(r=>hasSaldo(r) && podeFaturar(r));
  if (recStat==="faturados") rightRecs = rightRecs.filter(r=>hasFat(r));
  if (recComp!=="todas") rightRecs = rightRecs.filter(r=>compValue(r)===recComp);
  if (qRec.trim()) { const s=qRec.trim().toLowerCase(); rightRecs = rightRecs.filter(r=>(r.cliente||"").toLowerCase().includes(s)||(r.profissional||"").toLowerCase().includes(s)||(r.pep||"").toLowerCase().includes(s)); }
  rightRecs = rightRecs.sort(sortRecs);
  const LIMIT = 400; const rightShown = rightRecs.slice(0,LIMIT); const leftShown = leftNotes.slice(0,LIMIT);

  // Totais do que está filtrado (mostrados no topo de cada quadro).
  // Coluna de receitas: soma o MESMO que cada linha mostra — saldo a faturar
  // (faturados = o valor já faturado). Assim bate com a seleção e com as notas.
  const leftTotVal = leftNotes.reduce((s,n)=>s+(n.valorServicos||0),0);
  const rightTotVal = rightRecs.reduce((s,r)=>s+(recStat==="faturados" ? fat(r) : saldoR(r)),0);
  const rightPend = rightRecs.filter(r=>hasSaldo(r) && podeFaturar(r));
  const leftPendFiltered = leftNotes.filter(n=>!notaConc(n));
  const allRightSel = rightPend.length>0 && rightPend.every(r=>selRecs.has(r.id));
  const allLeftSel = leftPendFiltered.length>0 && leftPendFiltered.every(n=>selNotes.has(n.id));
  const toggleAllRecs = () => setSelRecs(s => { const n=new Set(s); if(allRightSel) rightPend.forEach(r=>n.delete(r.id)); else rightPend.forEach(r=>n.add(r.id)); return n; });
  const toggleAllNotes = () => setSelNotes(s => { const n=new Set(s); if(allLeftSel) leftPendFiltered.forEach(x=>n.delete(x.id)); else leftPendFiltered.forEach(x=>n.add(x.id)); return n; });

  // Pendências (sempre sobre o total da empresa, não o filtrado)
  const notasPend = empNotes.filter(n=>!notaConc(n)); const notasPendVal = notasPend.reduce((s,n)=>s+(n.valorServicos||0),0);
  const recsPend = empRecs.filter(r=>hasSaldo(r) && podeFaturar(r)); const recsPendVal = recsPend.reduce((s,r)=>s+saldoR(r),0);

  const selectedNotes = empNotes.filter(n=>selNotes.has(n.id));
  const somaNotes = selectedNotes.reduce((s,n)=>s+(n.valorServicos||0),0);
  const selRecList = empRecs.filter(r=>selRecs.has(r.id));
  const somaSel = selRecList.reduce((s,r)=>s+valorDe(r),0);   // soma dos valores a faturar
  const diff = Math.abs(somaSel-somaNotes);
  // Tolerância dura: só concilia se a diferença for de no máximo R$ 1,00.
  const bate = selectedNotes.length>0 && selRecs.size>0 && diff <= 1.005;

  const toggleRec = (id) => setSelRecs(s=>{ const n=new Set(s); n.has(id)?n.delete(id):n.add(id); return n; });
  // N:N — pode selecionar várias notas e várias receitas; a trava de R$ 1,00
  // garante que o total bata antes de conciliar.
  const toggleNote = (id) => setSelNotes(s=>{ const n=new Set(s); n.has(id)?n.delete(id):n.add(id); return n; });
  const resetSel = () => { setSelRecs(new Set()); setSelNotes(new Set()); setValores({}); };
  const setValor = (id,v) => setValores(m=>({...m,[id]:v}));
  const pickEmpresa = (cod) => { setEmpresa(cod); resetSel(); };

  // Sugestão (não automática): profissional/competência citados em ALGUMA nota selecionada.
  const selNomes = stripAcc(selectedNotes.map(n=>n.profissionais).join(" "));
  const selComps = selectedNotes.map(n=>n.competencias).join(" ");
  const isSug = (r) => { if(!selectedNotes.length) return false; const nm=stripAcc(r.profissional); const byN = nm && nm.length>3 && selNomes.includes((nm.split(" ")[0]||"")) && selNomes.includes((nm.split(" ").slice(-1)[0]||"")); const byC = r.competencia && selComps.includes(r.competencia); return !!(byN||byC); };
  const selSugeridos = () => { if(!selectedNotes.length) return; setSelRecs(new Set(rightRecs.filter(r=>hasSaldo(r) && isSug(r)).map(r=>r.id))); };

  function confirmar() {
    if(!selectedNotes.length||!selRecs.size) return;
    if(diff > 1.005) { toast(`Diferença de ${brl(diff)} — só é possível conciliar com diferença de até R$ 1,00.`, "error"); return; }
    const valoresMap = {}; selRecList.forEach(r=>{ valoresMap[r.id] = valorDe(r); });
    onConciliate([...selRecs], selectedNotes, valoresMap); resetSel();
  }

  const importBatches = [...new Set(notes.map(n=>n.importId).filter(Boolean))];
  const SortSel = ({value,onChange,opts,active}) => <select style={{...inp,width:"auto",fontSize:12,padding:"5px 8px",...(active?{borderColor:T.brand,background:T.brandBg,color:T.brand,fontWeight:700}:{})}} value={value} onChange={e=>onChange(e.target.value)}>{opts.map(([v,l])=><option key={v} value={v}>{l}</option>)}</select>;

  // ── Conferência: lotes de conciliação (nota × receita) da empresa selecionada.
  // Agrupa por conciliacao_id e compara Σnotas vs Σreceitas alocadas → acha o
  // par errado mesmo quando o total do cliente "bate".
  const lotesMap = {};
  notes.forEach(n => { if (n.conciliacaoId && n.empresa===empresa && !n.cancelada) { (lotesMap[n.conciliacaoId] = lotesMap[n.conciliacaoId] || { cid:n.conciliacaoId, notas:[], recs:[] }).notas.push(n); } });
  faturamentos.forEach(a => { const L = a.conciliacaoId && lotesMap[a.conciliacaoId]; if (L) { const r = records.find(x=>x.id===a.recordId); L.recs.push({ r, valor:a.valor||0 }); } });
  const lotes = Object.values(lotesMap).map(L => {
    const sn = L.notas.reduce((s,n)=>s+(n.valorServicos||0),0);
    const sr = L.recs.reduce((s,x)=>s+(x.valor||0),0);
    return { ...L, sn, sr, dif: sn-sr, bate: Math.abs(sn-sr) < 0.01 };
  }).sort((a,b)=>Math.abs(b.dif)-Math.abs(a.dif));
  const lotesDiverg = lotes.filter(L=>!L.bate).length;
  const lotesShown = confSoDif ? lotes.filter(L=>!L.bate) : lotes;

  // Lotes órfãos: têm receita alocada mas NENHUMA nota amarrada (ex.: a nota foi
  // apagada/recriada numa re-importação e perdeu o conciliacao_id). É o "Erik":
  // faturado de um lado, nota some do outro. Escopo: lotes com receita da empresa.
  const cidsComNota = new Set(notes.filter(n=>n.conciliacaoId).map(n=>n.conciliacaoId));
  const lotesOrfaos = [...new Set(faturamentos
    .filter(a => a.conciliacaoId && !cidsComNota.has(a.conciliacaoId))
    .map(a => a.conciliacaoId))]
    .filter(cid => faturamentos.some(a => a.conciliacaoId===cid && (records.find(x=>x.id===a.recordId)?.empresa===empresa)));

  return (
    <div>
      {importing && <NotesImportModal onImport={onImport} onClose={()=>setImporting(false)}/>}
      {noteDel && <ConfirmDialog title="Excluir nota da base" danger confirmLabel="Excluir"
        message={`Excluir a NF ${noteDel.numero} (${noteDel.tomadorNome||"—"} · ${brl(noteDel.valorServicos)})?${records.some(r=>r.municipalNoteId===noteDel.id)?" Os registros conciliados com ela serão reabertos." : ""}`}
        onConfirm={()=>onDeleteNote(noteDel)} onClose={()=>setNoteDel(null)}/>}
      {manage && (
        <Modal title="Importações de notas" onClose={()=>setManage(false)} footer={<Btn onClick={()=>setManage(false)}>Fechar</Btn>}>
          {importBatches.length===0 ? <div style={{fontSize:13,color:T.muted}}>Nenhuma importação ainda.</div>
            : importBatches.map(bid=>{ const lote=notes.filter(n=>n.importId===bid); return (
                <div key={bid} style={{display:"flex",alignItems:"center",gap:10,padding:"9px 0",borderBottom:`1px solid ${T.lineSoft}`,fontSize:13}}>
                  <span style={{flex:1}}>{lote[0]?.empresa||"—"} · {lote[0]?.municipio||"—"} · <b>{lote.length}</b> nota(s)</span>
                  <Btn small danger onClick={()=>onUndoImport(bid)}>↩ Desfazer</Btn>
                </div>); })}
        </Modal>
      )}

      {/* Cabeçalho congelado: título + empresa + cards ficam fixos ao rolar. */}
      <div style={{position:"sticky",top:0,zIndex:15,background:T.canvas,paddingTop:8,marginTop:-8}}>
      <div style={{display:"flex",alignItems:"center",gap:13,marginBottom:14,flexWrap:"wrap"}}>
        <HeadChip icon="receipt"/>
        <div style={{flex:1,minWidth:200}}>
          <h1 style={{...Ty.h1,fontSize:22}}>Conciliação de notas</h1>
          <div style={{...Ty.small,marginTop:3}}>{notes.length} nota(s) importada(s) · conciliação por empresa do grupo</div>
        </div>
        {isAdmin && orfas>0 && <Btn icon="undo" onClick={onReopenOrphans}>Reabrir notas órfãs ({orfas})</Btn>}
        {isAdmin && <Btn icon="folder" onClick={()=>setManage(true)}>Importações</Btn>}
        {isAdmin && <Btn primary icon="upload" onClick={()=>setImporting(true)}>Importar notas</Btn>}
      </div>

      <Card style={{padding:"10px 12px",marginBottom:14}}>
        <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
          <span style={{fontSize:13,fontWeight:700,color:T.ink}}>Empresa:</span>
          {empresasComDados.length===0 ? <span style={{fontSize:12,color:T.muted}}>Importe notas para começar.</span>
            : empresasComDados.map(e=>(
              <button key={e.cod} onClick={()=>pickEmpresa(e.cod)} style={{padding:"6px 12px",borderRadius:T.rMd,border:`1.5px solid ${empresa===e.cod?T.brand:T.line}`,background:empresa===e.cod?T.brandBg:"var(--surface)",color:empresa===e.cod?T.brand:T.inkSoft,fontWeight:empresa===e.cod?700:500,fontSize:12.5,cursor:"pointer"}}>{e.cod} — {e.nome}</button>
            ))}
        </div>
      </Card>

      {empresa && (
          <div style={{display:"flex",gap:14,flexWrap:"wrap",marginBottom:14}}>
            <Card style={{flex:1,minWidth:200,padding:"12px 14px",borderLeft:`3px solid ${T.warn}`}}>
              <div style={Ty.small}>Notas a conciliar (represadas)</div>
              <div style={{fontSize:18,fontWeight:800,color:T.ink}}>{notasPend.length}</div>
              <div style={{fontSize:11,color:T.muted}}>{brl(notasPendVal)} em aberto</div>
            </Card>
            <Card style={{flex:1,minWidth:200,padding:"12px 14px",borderLeft:`3px solid ${C.blue.solid}`}}>
              <div style={Ty.small}>Receitas sem nota (represadas)</div>
              <div style={{fontSize:18,fontWeight:800,color:T.ink}}>{recsPend.length}</div>
              <div style={{fontSize:11,color:T.muted}}>{brl(recsPendVal)} a faturar</div>
            </Card>
            <Card style={{flex:1,minWidth:200,padding:"12px 14px",borderLeft:`3px solid ${lotesDiverg>0?T.danger:T.ok}`,cursor:"pointer"}} onClick={()=>setShowConf(v=>!v)}>
              <div style={Ty.small}>Conferência de conciliações</div>
              <div style={{fontSize:18,fontWeight:800,color:lotesDiverg>0?T.danger:T.ok}}>{lotesDiverg>0?`${lotesDiverg} divergente(s)`:`✓ ${lotes.length} ok`}</div>
              <div style={{fontSize:11,color:T.muted}}>{lotes.length} lote(s) · clique para {showConf?"ocultar":"conferir"}</div>
            </Card>
            <Card style={{flex:1,minWidth:200,padding:"12px 14px",borderLeft:`3px solid ${lotesOrfaos.length>0?T.danger:T.ok}`}}>
              <div style={Ty.small}>Faturado sem nota (órfãos)</div>
              <div style={{fontSize:18,fontWeight:800,color:lotesOrfaos.length>0?T.danger:T.ok}}>{lotesOrfaos.length>0?`${lotesOrfaos.length} lote(s)`:`✓ 0`}</div>
              <div style={{fontSize:11,color:T.muted}}>receita conciliada, nota ausente</div>
            </Card>
          </div>
      )}
      </div>

      {!empresa ? (
        <Card style={{textAlign:"center",padding:"3rem"}}>
          <div style={{fontSize:32,marginBottom:10}}></div>
          <div style={{fontSize:14,color:T.muted}}>Escolha uma empresa do grupo acima para conciliar as notas com as receitas.</div>
        </Card>
      ) : (
        <>
          {showConf && (
            <Card style={{padding:0,overflow:"hidden",marginBottom:14,border:`1px solid ${lotesDiverg>0?T.dangerLine:T.line}`}}>
              <div style={{display:"flex",alignItems:"center",gap:10,padding:"10px 14px",borderBottom:`1px solid ${T.line}`,flexWrap:"wrap"}}>
                <span style={{fontWeight:700,fontSize:13,color:T.ink}}>Conferência nota × receita por lote</span>
                <span style={{fontSize:12,color:T.muted}}>{lotesShown.length} de {lotes.length}</span>
                <div style={{flex:1}}/>
                <label style={{display:"flex",alignItems:"center",gap:6,fontSize:12,color:T.inkSoft,cursor:"pointer"}}>
                  <input type="checkbox" checked={confSoDif} onChange={e=>setConfSoDif(e.target.checked)} style={{width:14,height:14}}/> só divergentes
                </label>
              </div>
              <div className="fc-scroll" style={{maxHeight:360,overflowY:"auto"}}>
                {lotesShown.length===0 ? <div style={{padding:"1.4rem",textAlign:"center",fontSize:13,color:T.muted}}>{lotes.length===0?"Nenhuma conciliação nesta empresa.":"Todos os lotes batem. ✓"}</div>
                  : lotesShown.map(L=>(
                    <div key={L.cid} style={{padding:"10px 14px",borderBottom:`1px solid ${T.lineSoft}`,background:L.bate?"var(--surface)":"#fef2f2"}}>
                      <div style={{display:"flex",alignItems:"center",gap:10,flexWrap:"wrap",marginBottom:6}}>
                        <Badge label={L.bate?"✓ bate":`≠ dif. ${brl(L.dif)}`} color={L.bate?"green":"red"} small/>
                        <span style={{fontSize:12,color:T.ink,fontWeight:600}}>{L.recs[0]?.r?.cliente || L.notas[0]?.tomadorNome || "—"}</span>
                        <div style={{flex:1}}/>
                        <span style={{fontSize:12,color:T.muted}}>Notas <b style={{color:T.ink}}>{brl(L.sn)}</b> ↔ Receitas <b style={{color:T.ink}}>{brl(L.sr)}</b></span>
                        <Btn small icon="undo" onClick={()=>onReopen({ conciliacaoId:L.cid })}>Reabrir</Btn>
                      </div>
                      <div style={{fontSize:11,color:T.muted,display:"flex",gap:16,flexWrap:"wrap"}}>
                        <span><b>Notas:</b> {L.notas.map(n=>`${n.numero} (${brl(n.valorServicos)})`).join(", ")||"—"}</span>
                        <span><b>Receitas:</b> {L.recs.map(x=>`${x.r?.profissional||x.r?.pep||"?"} (${brl(x.valor)})`).join(", ")||"—"}</span>
                      </div>
                    </div>
                  ))}
              </div>
            </Card>
          )}

          {(selRecs.size>0 || selectedNotes.length>0) && (
            <Card style={{padding:"12px 16px",marginBottom:14,display:"flex",alignItems:"center",gap:14,flexWrap:"wrap",position:"sticky",top:12,zIndex:20,boxShadow:T.shMd,border:`1px solid ${bate?T.okLine:T.line}`}}>
              <div style={{fontSize:13}}>
                <b>{selectedNotes.length}</b> nota(s) = <b>{brl(somaNotes)}</b> &nbsp;↔&nbsp; <b>{selRecs.size}</b> receita(s) = <b>{brl(somaSel)}</b>
                {selectedNotes.length>0 && selRecs.size>0 && <> &nbsp; {bate ? <span style={{color:T.ok,fontWeight:700}}>✓ bate</span> : <span style={{color:T.warn,fontWeight:700}}>≠ dif. {brl(diff)}</span>}</>}
              </div>
              <div style={{flex:1}}/>
              <Btn onClick={resetSel}>Limpar</Btn>
              {!isViewer && <Btn primary disabled={!selectedNotes.length||!selRecs.size||!bate} onClick={confirmar}>Conciliar {selectedNotes.length}×{selRecs.size}</Btn>}
            </Card>
          )}

          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14,alignItems:"start"}}>
            {/* ESQUERDA — NOTAS DA PREFEITURA */}
            <Card style={{padding:0,overflow:"hidden"}}>
              <div style={{padding:"10px 12px",borderBottom:`1px solid ${T.line}`}}>
                <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8,flexWrap:"wrap"}}>
                  <span style={{fontWeight:700,fontSize:13}}>Notas da prefeitura</span>
                  <span style={{fontSize:12,color:T.muted}}><b style={{color:T.ink}}>{leftNotes.length}</b> · {brl(leftTotVal)}</span>
                  <div style={{flex:1}}/>
                  {leftPendFiltered.length>0 && <Btn small onClick={toggleAllNotes}>{allLeftSel?"Limpar":"Tudo filtrado"}</Btn>}
                </div>
                <div style={{display:"flex",gap:6,flexWrap:"wrap",alignItems:"center"}}>
                  <input style={{...inp,flex:1,minWidth:120,fontSize:12,padding:"6px 9px"}} placeholder="nº, tomador, pedido" value={qNote} onChange={e=>setQNote(e.target.value)}/>
                  <SortSel value={noteCli} onChange={setNoteCli} opts={[["todos","Todos os clientes"],...(tomadoresUsados.includes(noteCli)||noteCli==="todos"?[]:[[noteCli,noteCli]]),...tomadoresUsados.map(c=>[c,c])]}/>
                  <div style={{display:"flex",alignItems:"center",gap:4}}>
                    <input type="date" title="Emissão — de" style={{...inp,width:"auto",fontSize:12,padding:"5px 8px"}} value={noteDe} onChange={e=>setNoteDe(e.target.value)}/>
                    <span style={{fontSize:12,color:T.muted}}>até</span>
                    <input type="date" title="Emissão — até" style={{...inp,width:"auto",fontSize:12,padding:"5px 8px"}} value={noteAte} onChange={e=>setNoteAte(e.target.value)}/>
                  </div>
                  {(noteDe||noteAte) && <Btn small onClick={()=>{setNoteDe("");setNoteAte("");}}>limpar</Btn>}
                  <SortSel value={noteStat} onChange={setNoteStat} opts={[["pendentes","Pendentes"],["conciliadas","Conciliadas"],["todas","Todas"]]}/>
                  <SortSel value={noteSort} onChange={setNoteSort} opts={[["valor_desc","↓ Valor"],["valor_asc","↑ Valor"],["data_desc","↓ Data"],["data_asc","↑ Data"],["tomador_az","A–Z"]]}/>
                </div>
              </div>
              <div className="fc-scroll" key={`ln-${noteStat}-${noteCli}-${noteDe}-${noteAte}-${noteSort}-${qNote}`} style={{maxHeight:480,overflowY:"auto"}}>
                {leftShown.length===0 ? <div style={{padding:"1.4rem",textAlign:"center",fontSize:13,color:T.muted}}>Nenhuma nota.</div>
                  : leftShown.map(n=>{
                      const conc=notaConc(n), on=selNotes.has(n.id), exp=expNote===n.id;
                      return (
                        <div key={n.id} style={{borderBottom:`1px solid ${T.lineSoft}`,background:on?T.brandBg:(conc?"#f6fdf9":"var(--surface)")}}>
                          <div style={{display:"flex",alignItems:"flex-start",gap:9,padding:"9px 12px"}}>
                            {conc
                              ? <span title="Conciliada" style={{width:15,textAlign:"center",color:T.ok,marginTop:2}}>✓</span>
                              : <input type="checkbox" checked={on} onChange={()=>toggleNote(n.id)} style={{marginTop:3}}/>}
                            <div style={{flex:1,minWidth:0,cursor:"pointer"}} onClick={()=>setExpNote(exp?"":n.id)}>
                              <div style={{fontSize:12.5,fontWeight:700,color:T.ink}}>NF {n.numero} · {brl(n.valorServicos)} {conc&&<Badge label="conciliada" color="green" small/>}</div>
                              <div style={{fontSize:11,color:T.muted,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{n.tomadorNome||"—"}</div>
                              <div style={{fontSize:11,color:T.muted}}>{fmtDT(n.emitidaEm)}{n.pedidos?` · pedido ${n.pedidos}`:""}{n.competencias?` · ${n.competencias}`:""}</div>
                            </div>
                            {conc && <button onClick={()=>onReopen({ conciliacaoId:n.conciliacaoId, noteId:n.id })} title="Desfazer conciliação" style={{background:"none",border:"none",cursor:"pointer",color:T.warn,fontSize:14}}>↩</button>}
                            <button onClick={()=>setExpNote(exp?"":n.id)} title="Detalhes" style={{background:"none",border:"none",cursor:"pointer",color:T.muted,fontSize:14}}>{exp?"▲":"ⓘ"}</button>
                            {isAdmin && <button onClick={()=>setNoteDel(n)} title="Excluir nota da base" style={{background:"none",border:"none",cursor:"pointer",color:T.danger,fontSize:14}}><Icon name="trash" size={14}/></button>}
                          </div>
                          {exp && <div style={{padding:"0 12px 11px 33px",fontSize:11.5,color:T.inkSoft,lineHeight:1.5}}>
                            <div><b>Tomador:</b> {n.tomadorNome||"—"} · {n.tomadorCnpj||"—"}</div>
                            {n.profissionais && <div><b>Profissionais:</b> {n.profissionais}</div>}
                            <div><b>Valor serviços:</b> {brl(n.valorServicos)} · <b>ISS:</b> {brl(n.iss)}</div>
                            <div style={{marginTop:5,padding:"7px 9px",background:T.canvas,borderRadius:T.rMd,whiteSpace:"pre-wrap"}}>{n.discriminacao||"(sem discriminação)"}</div>
                          </div>}
                        </div>
                      );
                    })}
                {leftNotes.length>LIMIT && <div style={{padding:"8px 12px",fontSize:11,color:T.muted,textAlign:"center"}}>Mostrando {LIMIT} de {leftNotes.length} — refine a busca.</div>}
              </div>
            </Card>

            {/* DIREITA — RECEITAS RECONHECIDAS */}
            <Card style={{padding:0,overflow:"hidden"}}>
              <div style={{padding:"10px 12px",borderBottom:`1px solid ${T.line}`}}>
                <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:8,flexWrap:"wrap"}}>
                  <span style={{fontWeight:700,fontSize:13}}>Receitas reconhecidas</span>
                  <span style={{fontSize:12,color:T.muted}}><b style={{color:T.ink}}>{rightRecs.length}</b> · {brl(rightTotVal)} <span style={{fontSize:10.5}}>{recStat==="faturados"?"faturado":"a faturar"}</span></span>
                  <div style={{flex:1}}/>
                  {rightPend.length>0 && <Btn small onClick={toggleAllRecs}>{allRightSel?"Limpar":"Tudo filtrado"}</Btn>}
                  {selectedNotes.length>0 && <Btn small onClick={selSugeridos}>Sugeridos</Btn>}
                </div>
                <div style={{display:"flex",gap:6,flexWrap:"wrap",alignItems:"center"}}>
                  <input style={{...inp,flex:1,minWidth:120,fontSize:12,padding:"6px 9px"}} placeholder="cliente, profissional, PEP" value={qRec} onChange={e=>setQRec(e.target.value)}/>
                  <SortSel value={recDim} onChange={v=>{setRecDim(v);setRecComp("todas");}} opts={[["servico","Ver por: Serviço"],["ciclo","Ver por: Ciclo"]]}/>
                  <SortSel value={recComp} onChange={setRecComp} active={recComp!=="todas"} opts={[["todas",recDim==="ciclo"?"Todos ciclos":"Todos os meses"],...compsUsadas.map(c=>[c,c])]}/>
                  <SortSel value={recStat} onChange={setRecStat} opts={[["pendentes","Sem nota"],["faturados","Faturados"],["todas","Todas"]]}/>
                  <SortSel value={recSort} onChange={setRecSort} opts={[["valor_desc","↓ Valor"],["valor_asc","↑ Valor"],["cliente_az","A–Z"],["comp","Competência"]]}/>
                </div>
                <label style={{display:"flex",alignItems:"center",gap:6,marginTop:8,fontSize:12,color:T.inkSoft,cursor:"pointer"}}>
                  <input type="checkbox" checked={incluirNaoLib} onChange={e=>setIncluirNaoLib(e.target.checked)} style={{width:14,height:14}}/>
                  Incluir receitas ainda não liberadas no passo a passo <span style={{color:T.muted,fontSize:11}}>(conciliar já libera o funil)</span>
                </label>
              </div>
              <div className="fc-scroll" key={`rr-${recDim}-${recComp}-${recStat}-${recSort}-${qRec}`} style={{maxHeight:480,overflowY:"auto"}}>
                {rightShown.length===0 ? <div style={{padding:"1.4rem",textAlign:"center",fontSize:13,color:T.muted}}>Nenhuma receita.</div>
                  : rightShown.map(r=>{
                      const full=!hasSaldo(r), parcial=hasFat(r)&&hasSaldo(r), on=selRecs.has(r.id), sug=hasSaldo(r)&&isSug(r), falta=faltaDatas(r), dc=diaCorteDe(r);
                      const temCls=hasSaldo(r)&&(r.classMotivo||r.classObs), clsRepres=temCls&&categoriaOf(r,clients).cat==="represado";
                      return (
                        <div key={r.id} style={{display:"flex",alignItems:"center",gap:9,padding:"8px 12px",borderBottom:`1px solid ${T.lineSoft}`,background:on?T.brandBg:(sug?"#f0fdf4":"var(--surface)")}}>
                          {full
                            ? <span title="Faturada" style={{width:15,textAlign:"center",color:T.ok}}>✓</span>
                            : <input type="checkbox" checked={on} onChange={()=>toggleRec(r.id)}/>}
                          <div style={{flex:1,minWidth:0}}>
                            <div style={{fontSize:12.5,fontWeight:600,color:T.ink,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{r.cliente||"—"} {parcial&&<Badge label="parcial" color="orange" small/>} {(varByRec[r.id]||0)>0.001&&<Badge label="variação" color="purple" small/>} {r.valorAnterior!=null&&<Badge label="valor mudou" color="red" small/>} {sug&&<Badge label="sugerido" color="green" small/>} {falta&&<Badge label="faltam datas" color="yellow" small/>} {dc>0&&<Badge label={`ciclo ${compFat(r)}`} color="teal" small/>} {temCls&&<Badge label={(clsRepres?"⚠ ":"")+(r.classMotivo||"obs")} color={clsRepres?"red":"gray"} small/>}</div>
                            <div style={{fontSize:11,color:T.muted}}>{r.competencia} · {r.tipo} · {r.profissional||r.pep||"—"}{dc>0?` · fatura ${compFat(r)} · ${r.inicio}–${r.fim}`:""}{hasFat(r)?` · faturado ${brl(fat(r))} de ${brl(bill(r))}`:""}{(varByRec[r.id]||0)>0.001?` · inclui variação ${brl(varByRec[r.id])}`:""}</div>
                            {temCls&&r.classObs&&<div style={{fontSize:11,color:clsRepres?C.red.solid:T.inkSoft,marginTop:2,fontStyle:"italic",whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>📝 {r.classObs}</div>}
                          </div>
                          {on && hasSaldo(r)
                            ? <input style={{...inp,width:110,fontSize:12,padding:"4px 7px",textAlign:"right"}} title="Valor a faturar (parcial)" value={valores[r.id] ?? saldoR(r).toFixed(2)} onChange={e=>setValor(r.id,e.target.value)} onClick={e=>e.stopPropagation()}/>
                            : <div style={{fontSize:12.5,fontWeight:700,color:full?T.muted:T.ink,whiteSpace:"nowrap"}}>{brl(full ? fat(r) : saldoR(r))}</div>}
                        </div>
                      );
                    })}
                {rightRecs.length>LIMIT && <div style={{padding:"8px 12px",fontSize:11,color:T.muted,textAlign:"center"}}>Mostrando {LIMIT} de {rightRecs.length} — refine a busca.</div>}
              </div>
            </Card>
          </div>
        </>
      )}
    </div>
  );
}

// Conta quantos CNPJs um cadastro reúne (grupo de empresas).
function clientCnpjs(c) {
  const a = parseJSON(c.cnpjs, null);
  if (Array.isArray(a) && a.length) return a.map(e=>(e.cnpj||"")).filter(Boolean);
  return c.cnpj ? [c.cnpj] : [];
}

function ClientsView({ clients, isAdmin, isViewer=false, onSave, onDelete, onBulkImport, onMerge }) {
  const [editing, setEditing] = useState(null);
  const [importing, setImporting] = useState(false);
  const [q, setQ] = useState("");
  const [status, setStatus] = useState("todos");
  const [grupo, setGrupo] = useState("");
  const [page, setPage] = useState(0);
  const [sel, setSel] = useState(() => new Set());   // ids selecionados para agrupar
  const [merging, setMerging] = useState(null);       // { ids, nome } — criar grupo novo
  const [adding, setAdding] = useState(null);         // { ids, pick } — incluir em grupo existente
  const PAGE = 50;

  const grupos = [...new Set(clients.map(c=>(c.grupoEmpresa||"").trim()).filter(Boolean))].sort();

  let filtered = clients;
  if (status==="incompletos") filtered = filtered.filter(c=>c.incompleto);
  if (status==="completos")   filtered = filtered.filter(c=>!c.incompleto);
  if (status==="grupos")      filtered = filtered.filter(c=>clientCnpjs(c).length>1);
  if (grupo.trim())           { const g=grupo.trim().toLowerCase(); filtered = filtered.filter(c=>(c.grupoEmpresa||"").toLowerCase().includes(g)); }
  if (q.trim()) { const s=q.trim().toLowerCase(); const dig=s.replace(/\D/g,""); filtered = filtered.filter(c => (c.nome||"").toLowerCase().includes(s) || (c.codSap||"").toLowerCase().includes(s) || (!!dig && clientCnpjs(c).some(x=>x.includes(dig)))); }

  const incompletos = clients.filter(c=>c.incompleto).length;
  const totalPages = Math.max(1, Math.ceil(filtered.length/PAGE));
  const pg = Math.min(page, totalPages-1);
  const pageItems = filtered.slice(pg*PAGE, pg*PAGE+PAGE);
  const resetPage = (fn)=>(v)=>{ fn(v); setPage(0); };

  const toggle = (id) => setSel(s => { const n=new Set(s); n.has(id)?n.delete(id):n.add(id); return n; });
  const allSel = filtered.length>0 && filtered.every(c=>sel.has(c.id));
  const someSel = filtered.some(c=>sel.has(c.id));
  const toggleAll = () => setSel(s => { const n=new Set(s); if(allSel) filtered.forEach(c=>n.delete(c.id)); else filtered.forEach(c=>n.add(c.id)); return n; });
  const selList = clients.filter(c=>sel.has(c.id));
  const startMerge = () => setMerging({ ids:[...sel], nome: selList[0]?.nome || "" });
  const confirmMerge = async () => {
    await onMerge(merging.ids, merging.nome.trim() || selList[0]?.nome || "Grupo");
    setSel(new Set()); setMerging(null);
  };
  // Incluir os selecionados num grupo já existente (o grupo é o cadastro-base).
  const startAdd = () => setAdding({ ids:[...sel], pick:"", target:null });
  const grupoCandidatos = clients.filter(c => !adding?.ids.includes(c.id));   // não pode escolher os próprios selecionados
  const addMatches = adding ? grupoCandidatos.filter(c => !adding.pick.trim() || (c.nome||"").toLowerCase().includes(adding.pick.trim().toLowerCase())).slice(0,8) : [];
  const confirmAdd = async (target) => {
    await onMerge([target.id, ...adding.ids], target.nome, target.id);
    setSel(new Set()); setAdding(null);
  };

  return (
    <div>
      {editing && <ClientModal client={editing.id?editing:null} onSave={onSave} onDelete={onDelete} onClose={()=>setEditing(null)}/>}
      {importing && <ClientImportModal existing={clients} onImport={onBulkImport} onClose={()=>setImporting(false)}/>}
      {merging && (
        <Modal title="Agrupar clientes" onClose={()=>setMerging(null)} footer={<>
          <Btn onClick={()=>setMerging(null)}>Cancelar</Btn>
          <Btn primary onClick={confirmMerge}>Agrupar {merging.ids.length} cadastros</Btn>
        </>}>
          <div style={{fontSize:13,color:T.inkSoft,marginBottom:12,lineHeight:1.6}}>
            Os <b>{merging.ids.length}</b> cadastros selecionados viram <b>um único cadastro de grupo</b>, reunindo todos os CNPJs.
            Os cadastros individuais são <b>removidos</b> e passam a aparecer apenas dentro do grupo.
          </div>
          <Field label="Nome do grupo">
            <input style={inp} value={merging.nome} onChange={e=>setMerging(m=>({...m,nome:e.target.value}))} placeholder="Ex.: Klabin"/>
          </Field>
          <div style={{marginTop:12,padding:"10px 12px",background:T.canvas,borderRadius:T.rMd,border:`1px solid ${T.line}`}}>
            <div style={{...Ty.small,marginBottom:6,fontWeight:600}}>Empresas que serão reunidas:</div>
            {selList.map(c=>(
              <div key={c.id} style={{fontSize:12,color:T.inkSoft,padding:"3px 0"}}>
                • {c.nome} {clientCnpjs(c).length>0 && <span style={{fontFamily:"monospace",color:T.muted}}>({clientCnpjs(c).length} CNPJ)</span>}
              </div>
            ))}
          </div>
        </Modal>
      )}
      {adding && (
        <Modal title="Incluir em grupo existente" onClose={()=>setAdding(null)} footer={<>
          <Btn onClick={()=>setAdding(null)}>Cancelar</Btn>
          <Btn primary disabled={!adding.target} onClick={()=>confirmAdd(adding.target)}>
            {adding.target ? `Incluir em "${adding.target.nome}"` : "Escolha o destino"}
          </Btn>
        </>}>
          <div style={{fontSize:13,color:T.inkSoft,marginBottom:12,lineHeight:1.6}}>
            Escolha o <b>cadastro/grupo de destino</b>. Os <b>{adding.ids.length}</b> cadastros selecionados serão reunidos nele (CNPJs somados) e <b>removidos</b> como cadastros separados.
          </div>
          <Field label="Buscar grupo de destino">
            <input style={inp} autoFocus value={adding.pick} onChange={e=>setAdding(a=>({...a,pick:e.target.value,target:null}))} placeholder="Digite o nome do grupo… ex.: Elfa"/>
          </Field>
          <div style={{marginTop:10,border:`1px solid ${T.line}`,borderRadius:T.rMd,overflow:"hidden",maxHeight:260,overflowY:"auto"}}>
            {addMatches.length===0
              ? <div style={{padding:"14px",fontSize:12.5,color:T.muted,textAlign:"center"}}>Nenhum cadastro encontrado.</div>
              : addMatches.map(c=>{
                  const n=clientCnpjs(c).length;
                  const on=adding.target?.id===c.id;
                  return (
                    <button key={c.id} onClick={()=>setAdding(a=>({...a,target:c}))} className="fc-row" style={{display:"flex",alignItems:"center",gap:8,width:"100%",textAlign:"left",border:"none",borderBottom:`1px solid ${T.lineSoft}`,background:on?T.brandSoft||"#eef2ff":"var(--surface)",cursor:"pointer",padding:"10px 12px",fontSize:13,color:T.ink}}>
                      <span style={{width:16,color:T.brand,fontWeight:800}}>{on?"✓":""}</span>
                      <span style={{flex:1,fontWeight:600}}>{c.nome}</span>
                      {n>1 && <Badge label={`${n} CNPJs`} color="blue" small/>}
                    </button>
                  );
                })}
          </div>
        </Modal>
      )}

      <div style={{display:"flex",alignItems:"center",gap:13,marginBottom:14,flexWrap:"wrap"}}>
        <HeadChip icon="building"/>
        <div style={{flex:1,minWidth:200}}>
          <h1 style={{...Ty.h1,fontSize:22}}>Clientes</h1>
          <div style={{...Ty.small, marginTop:3}}>{clients.length} cliente(s){incompletos>0 && <> · <b style={{color:T.warn}}>{incompletos} incompleto(s)</b></>}</div>
        </div>
        {isAdmin && <Btn icon="upload" onClick={()=>setImporting(true)}>Importar clientes</Btn>}
        {!isViewer && <Btn primary icon="plus" onClick={()=>setEditing({ temPortal:false })}>Novo cliente</Btn>}
      </div>

      <Card style={{padding:"10px 12px",marginBottom:14}}>
        <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
          <select style={{...inp,width:"auto"}} value={status} onChange={e=>resetPage(setStatus)(e.target.value)} aria-label="Status do cadastro">
            <option value="todos">Todos os cadastros</option>
            <option value="incompletos">Incompletos</option>
            <option value="completos">✓ Completos</option>
            <option value="grupos">Grupos (vários CNPJs)</option>
          </select>
          {grupos.length>0 && (
            <>
              <input style={{...inp,width:160}} list="fc-grupos" placeholder="Grupo de empresa…" value={grupo} onChange={e=>resetPage(setGrupo)(e.target.value)} aria-label="Filtrar por grupo de empresa"/>
              <datalist id="fc-grupos">{grupos.map(g=><option key={g} value={g}/>)}</datalist>
            </>
          )}
          <input style={{...inp,flex:1,minWidth:200}} placeholder="Nome, Cód. SAP ou CNPJ..." value={q} onChange={e=>resetPage(setQ)(e.target.value)}/>
        </div>
      </Card>

      {sel.size>0 && (
        <Card style={{padding:"10px 14px",marginBottom:14,display:"flex",alignItems:"center",gap:12,flexWrap:"wrap",background:T.canvas}}>
          <b style={{fontSize:13,color:T.ink}}>{sel.size} selecionado(s)</b>
          <Btn small onClick={()=>setSel(new Set())}>Limpar</Btn>
          <div style={{flex:1}}/>
          {!isViewer && <Btn small icon="plus" onClick={startAdd}>Incluir em grupo existente</Btn>}
          {!isViewer && <Btn primary small icon="link" disabled={sel.size<2} onClick={startMerge}>Agrupar (novo grupo)</Btn>}
        </Card>
      )}

      {filtered.length===0
        ? <Card style={{textAlign:"center",padding:"3rem"}}>
            <div style={{fontSize:32,marginBottom:10}}></div>
            <div style={{fontSize:14,color:T.muted}}>{clients.length===0?"Nenhum cliente cadastrado. Importe a base ou clique em “+ Novo cliente”.":"Nenhum cliente encontrado."}</div>
          </Card>
        : <Card style={{padding:0,overflow:"hidden"}}>
            <div className="fc-scroll" style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                <thead><tr style={{background:T.canvas}}>
                  <th style={{padding:"10px 14px",textAlign:"left",borderBottom:`1px solid ${T.line}`,width:34}}>
                    <input type="checkbox" checked={allSel} ref={el=>{ if(el) el.indeterminate = !allSel && someSel; }} onChange={toggleAll} aria-label="Selecionar todos"/>
                  </th>
                  {["Cliente","Cód. SAP","CNPJ","Responsável","Account manager","Status"].map((h,i)=>
                    <th key={i} style={{padding:"10px 14px",textAlign:"left",borderBottom:`1px solid ${T.line}`,fontWeight:600,color:T.muted,whiteSpace:"nowrap"}}>{h}</th>
                  )}
                </tr></thead>
                <tbody>
                  {pageItems.map(c=>{
                    const cs = clientCnpjs(c);
                    return (
                    <tr key={c.id} className="fc-row" style={{borderBottom:`1px solid ${T.lineSoft}`,cursor:"pointer",background:sel.has(c.id)?T.canvas:undefined}}>
                      <td style={{padding:"10px 14px"}} onClick={e=>e.stopPropagation()}>
                        <input type="checkbox" checked={sel.has(c.id)} onChange={()=>toggle(c.id)} aria-label={`Selecionar ${c.nome}`}/>
                      </td>
                      <td style={{padding:"10px 14px",fontWeight:600,color:T.ink}} onClick={()=>setEditing(c)}>
                        {c.nome} {cs.length>1 && <Badge label={`${cs.length} CNPJs`} color="blue" small/>}
                      </td>
                      <td style={{padding:"10px 14px",color:T.inkSoft,fontFamily:"monospace"}} onClick={()=>setEditing(c)}>{c.codSap||"—"}</td>
                      <td style={{padding:"10px 14px",color:T.inkSoft,fontFamily:"monospace",fontSize:11}} onClick={()=>setEditing(c)}>{cs[0]||"—"}{cs.length>1 && <span style={{color:T.muted}}> +{cs.length-1}</span>}</td>
                      <td style={{padding:"10px 14px",color:T.inkSoft}} onClick={()=>setEditing(c)}>{c.owner||"—"}</td>
                      <td style={{padding:"10px 14px",color:T.inkSoft}} onClick={()=>setEditing(c)}>{c.accountManager||"—"}</td>
                      <td style={{padding:"10px 14px"}} onClick={()=>setEditing(c)}>{c.incompleto ? <Badge label="Incompleto" color="yellow" small dot/> : <Badge label="Completo" color="green" small dot/>}</td>
                    </tr>
                  );})}
                </tbody>
              </table>
            </div>
            {totalPages>1 && <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",padding:"10px 14px",borderTop:`1px solid ${T.line}`,fontSize:12,color:T.muted}}>
              <span>Mostrando {pg*PAGE+1}–{Math.min((pg+1)*PAGE,filtered.length)} de {filtered.length}</span>
              <div style={{display:"flex",gap:8,alignItems:"center"}}>
                <Btn small disabled={pg<=0} onClick={()=>setPage(pg-1)}>◀ Anterior</Btn>
                <span>Página {pg+1} de {totalPages}</span>
                <Btn small disabled={pg>=totalPages-1} onClick={()=>setPage(pg+1)}>Próxima ▶</Btn>
              </div>
            </div>}
          </Card>}
    </div>
  );
}

// ─── HOME (tela inicial) ─────────────────────────────────────────────────────
function MuralEditModal({ mural, onSave, onClose }) {
  const [frase, setFrase] = useState(mural.frase||"");
  const [autor, setAutor] = useState(mural.autor||"");
  const [lembretes, setLem] = useState(() => mural.lembretes?.length ? [...mural.lembretes] : [""]);
  const setL = (i,v) => setLem(a=>a.map((x,j)=>j===i?v:x));
  const addL = () => setLem(a=>[...a,""]);
  const delL = (i) => setLem(a=>a.filter((_,j)=>j!==i));
  function save() { onSave({ id:mural.id, frase:frase.trim(), autor:autor.trim(), lembretes:lembretes.map(s=>s.trim()).filter(Boolean) }); onClose(); }
  return (
    <Modal title="Editar mural da semana" subtitle="Aparece na tela inicial de todo mundo" onClose={onClose}
      footer={<><Btn onClick={onClose}>Cancelar</Btn><Btn primary onClick={save}>Salvar mural</Btn></>}>
      <Field label="Frase da semana"><textarea style={{...inp,minHeight:70,resize:"vertical"}} value={frase} onChange={e=>setFrase(e.target.value)} placeholder="Ex.: Feito é melhor que perfeito. Bora fechar o mês!"/></Field>
      <Field label="Autor (opcional)"><input style={inp} value={autor} onChange={e=>setAutor(e.target.value)} placeholder="Ex.: Daniela"/></Field>
      <div style={{marginTop:6}}>
        <label style={Ty.label}>Lembretes da semana</label>
        {lembretes.map((l,i)=>(
          <div key={i} style={{display:"flex",gap:8,marginBottom:6}}>
            <input style={inp} value={l} onChange={e=>setL(i,e.target.value)} placeholder="Ex.: Corte dia 25 · fechar conciliação de SP"/>
            <Btn small onClick={()=>delL(i)}>✕</Btn>
          </div>
        ))}
        <Btn small onClick={addL}>+ Lembrete</Btn>
      </div>
    </Modal>
  );
}

function ApelidoModal({ atual, onSave, onClose }) {
  const [v, setV] = useState(atual||"");
  return (
    <Modal title="Como querem te chamar?" subtitle="Seu apelido aparece na saudação da tela inicial" onClose={onClose}
      footer={<><Btn onClick={onClose}>Cancelar</Btn><Btn primary onClick={()=>{ onSave(v.trim()); onClose(); }}>Salvar</Btn></>}>
      <Field label="Apelido"><input style={inp} autoFocus value={v} onChange={e=>setV(e.target.value)} placeholder="Ex.: Fê, Lay, Dani…"/></Field>
    </Modal>
  );
}

function HomeView({ user, isAdmin, records, notes, tasks, profiles, fatByRec={}, varByRec={}, mural, onSaveMural, onSaveApelido, onNavigate }) {
  const [editing, setEditing] = useState(false);
  const [editApelido, setEditApelido] = useState(false);
  // Relógio leve: mantém data/saudação sempre corretas (vira o dia / muda o turno)
  // sem depender de refresh. Atualiza a cada minuto e ao voltar o foco pra aba.
  const [agora, setAgora] = useState(() => new Date());
  useEffect(() => {
    const tick = () => setAgora(new Date());
    const id = setInterval(tick, 60000);
    window.addEventListener("focus", tick);
    document.addEventListener("visibilitychange", tick);
    return () => { clearInterval(id); window.removeEventListener("focus", tick); document.removeEventListener("visibilitychange", tick); };
  }, []);
  const nome = user.apelido || (user.name||"").split(" ")[0];
  const hora = agora.getHours();
  const saud = hora<12 ? "Bom dia" : hora<18 ? "Boa tarde" : "Boa noite";
  const hoje = agora.toLocaleDateString("pt-BR", { weekday:"long", day:"2-digit", month:"long" });

  const semNota   = records.filter(r=>!(r.conciliacaoId||r.municipalNoteId) && r.progress?.p5_liberado).length;
  const notasPendList = notes.filter(n=>!n.cancelada && !n.conciliacaoId && !records.some(r=>r.municipalNoteId===n.id));
  const notasPend = notasPendList.length;
  // Conciliação atrasada: nota emitida em dia(s) ANTERIORES ainda sem conciliar.
  const hojeStr = `${agora.getFullYear()}-${String(agora.getMonth()+1).padStart(2,"0")}-${String(agora.getDate()).padStart(2,"0")}`;
  const atrasadas = notasPendList.filter(n=>{ const d=String(n.emitidaEm||"").slice(0,10); return d && d < hojeStr; }).length;
  const faltamDatas = records.filter(faltaDatas).length;
  const minhasTarefas = tasks.filter(t=>t.status!=="done" && (t.assignee===user.name || !t.assignee)).length;
  const valMudouFat = records.filter(r=>r.valorAnterior!=null && (fatByRec[r.id]||0)>0.001).length;
  const valMudou    = records.filter(r=>r.valorAnterior!=null && !((fatByRec[r.id]||0)>0.001)).length;

  // Faturamento do mês (competência mais recente presente na base)
  const comps = [...new Set(records.map(r=>r.competencia).filter(Boolean))].sort((a,b)=>compRank(a).localeCompare(compRank(b)));
  const compAtual = comps[comps.length-1] || "";
  const doMes = records.filter(r=>r.competencia===compAtual);
  const totalMes = doMes.reduce((s,r)=>s+(r.valorTotal||0)+(varByRec[r.id]||0),0);   // faturável (receita + variação)
  const fatMes = doMes.reduce((s,r)=>s+(fatByRec[r.id]||0),0);
  const pctMes = totalMes ? Math.round(fatMes/totalMes*100) : 0;

  // Aniversariantes do mês
  const mesAtual = String(agora.getMonth()+1).padStart(2,"0");
  const diaAtual = String(agora.getDate()).padStart(2,"0");
  const aniversariantes = (profiles||[]).filter(p=>{ const [,mm]=String(p.aniversario||"").split("/"); return mm===mesAtual; })
    .sort((a,b)=>String(a.aniversario).localeCompare(String(b.aniversario)));

  const Pend = ({ icon, n, label, color, to }) => (
    <button onClick={()=>onNavigate(to)} className="fc-btn fc-card-int" style={{textAlign:"left",border:`1px solid ${T.line}`,background:"var(--surface)",borderRadius:T.rLg,padding:"14px 16px",cursor:"pointer",display:"flex",flexDirection:"column",gap:6,borderLeft:`4px solid ${color}`}}>
      <span style={{color}}><Icon name={icon} size={22}/></span>
      <div style={{fontSize:24,fontWeight:800,color:T.ink,fontFamily:T.fontDisplay}}>{n}</div>
      <div style={{fontSize:12,color:T.muted}}>{label}</div>
    </button>
  );

  return (
    <div>
      {editing && <MuralEditModal mural={mural} onSave={onSaveMural} onClose={()=>setEditing(false)}/>}
      {editApelido && <ApelidoModal atual={user.apelido} onSave={onSaveApelido} onClose={()=>setEditApelido(false)}/>}

      {/* Hero — charcoal (neutro), com laranja só de acento */}
      <div style={{background:`linear-gradient(120deg, #201b18, ${T.dark})`,borderRadius:T.rXl,padding:"26px 28px",color:"#fff",marginBottom:18,boxShadow:T.shMd,borderLeft:`4px solid ${T.brand}`}}>
        <div style={{fontSize:12,color:"rgba(255,255,255,.6)",textTransform:"capitalize"}}>{hoje}</div>
        <div style={{fontSize:26,fontWeight:700,fontFamily:T.fontDisplay,marginTop:4,display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
          {saud}, {nome}!
          <button onClick={()=>setEditApelido(true)} title="Editar meu apelido" style={{background:T.brand,border:"none",color:"#fff",borderRadius:T.rPill,padding:"4px 11px",fontSize:11,fontWeight:600,cursor:"pointer",display:"inline-flex",alignItems:"center",gap:5,fontFamily:T.font}}><Icon name="pencil" size={12}/>apelido</button>
        </div>
        <div style={{fontSize:13,color:"rgba(255,255,255,.8)",marginTop:4}}>Bem-vindo(a) ao <b style={{color:"#fff"}}>Order to Cash</b> — o painel do time O2C.</div>
      </div>

      {/* Alerta de conciliação atrasada (notas de dias anteriores sem conciliar) */}
      {atrasadas>0 && (
        <button onClick={()=>onNavigate("concil")} className="fc-btn" style={{width:"100%",textAlign:"left",display:"flex",alignItems:"center",gap:12,background:T.dangerBg,border:`1px solid ${T.dangerLine}`,borderRadius:T.rLg,padding:"14px 16px",marginBottom:18,cursor:"pointer",color:T.danger}}>
          <Icon name="alert" size={22}/>
          <div style={{flex:1}}>
            <div style={{fontSize:14,fontWeight:700,fontFamily:T.fontDisplay}}>Conciliação atrasada</div>
            <div style={{fontSize:12,color:T.inkSoft}}><b>{atrasadas}</b> nota(s) da prefeitura de dias anteriores ainda sem conciliar. A conciliação deve ser feita diariamente.</div>
          </div>
          <Icon name="chevronRight" size={18}/>
        </button>
      )}

      {/* Alerta vermelho: receita mudou DEPOIS da NF emitida (nota a cancelar/corrigir) */}
      {valMudouFat>0 && (
        <button onClick={()=>onNavigate("time")} className="fc-btn" style={{width:"100%",textAlign:"left",display:"flex",alignItems:"center",gap:12,background:T.dangerBg,border:`1px solid ${T.dangerLine}`,borderRadius:T.rLg,padding:"14px 16px",marginBottom:18,cursor:"pointer",color:T.danger}}>
          <Icon name="alert" size={22}/>
          <div style={{flex:1}}>
            <div style={{fontSize:14,fontWeight:700,fontFamily:T.fontDisplay}}>Valor mudou após faturamento</div>
            <div style={{fontSize:12,color:T.inkSoft}}><b>{valMudouFat}</b> receita(s) tiveram o valor alterado na base <b>depois</b> da NF emitida. A diferença virou saldo a faturar e a nota pode precisar de ajuste/cancelamento.</div>
          </div>
          <Icon name="chevronRight" size={18}/>
        </button>
      )}

      {/* Aviso amarelo: valor alterado em receitas ainda não faturadas */}
      {valMudou>0 && (
        <button onClick={()=>onNavigate("time")} className="fc-btn" style={{width:"100%",textAlign:"left",display:"flex",alignItems:"center",gap:12,background:"#fffbeb",border:"1px solid #fde68a",borderRadius:T.rLg,padding:"12px 16px",marginBottom:18,cursor:"pointer",color:"#92400e"}}>
          <Icon name="info" size={20}/>
          <div style={{flex:1}}>
            <div style={{fontSize:13,fontWeight:700,fontFamily:T.fontDisplay}}>Valores atualizados no fechamento</div>
            <div style={{fontSize:12,color:T.inkSoft}}><b>{valMudou}</b> receita(s) em andamento tiveram o valor alterado na última importação. Confira na Minha visão.</div>
          </div>
          <Icon name="chevronRight" size={18}/>
        </button>
      )}

      {/* Faturamento do mês + Aniversariantes */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(280px,1fr))",gap:14,marginBottom:18}}>
        <Card style={{padding:"16px 18px"}}>
          <div style={{display:"flex",alignItems:"baseline",gap:8,marginBottom:10}}>
            <span style={{fontSize:14,fontWeight:700,fontFamily:T.fontDisplay,color:T.ink,flex:1,display:"inline-flex",alignItems:"center",gap:7}}><span style={{color:T.brand}}><Icon name="wallet" size={18}/></span>Faturamento do mês</span>
            <span style={{fontSize:12,color:T.muted}}>{compAtual||"—"}</span>
          </div>
          <div style={{display:"flex",alignItems:"baseline",gap:8}}>
            <span style={{fontSize:24,fontWeight:800,color:T.brand}}>{fmtShort(fatMes)}</span>
            <span style={{fontSize:12,color:T.muted}}>de {fmtShort(totalMes)} reconhecido</span>
          </div>
          <div style={{height:10,background:T.lineSoft,borderRadius:6,marginTop:12,overflow:"hidden"}}>
            <div style={{height:10,width:`${pctMes}%`,borderRadius:6,background:`linear-gradient(90deg,${T.brand},${T.accent})`,transition:"width .5s"}}/>
          </div>
          <div style={{fontSize:12,color:T.muted,marginTop:6}}><b style={{color:T.ink}}>{pctMes}%</b> faturado · {doMes.length} registro(s)</div>
        </Card>

        <Card style={{padding:"16px 18px"}}>
          <div style={{fontSize:14,fontWeight:700,fontFamily:T.fontDisplay,color:T.ink,marginBottom:10,display:"flex",alignItems:"center",gap:7}}><span style={{color:T.brand}}><Icon name="gift" size={18}/></span>Aniversariantes de {agora.toLocaleDateString("pt-BR",{month:"long"})}</div>
          {aniversariantes.length===0
            ? <div style={{fontSize:13,color:T.muted}}>Ninguém faz aniversário este mês. {isAdmin?"(Cadastre em Gestão de acessos)":""}</div>
            : <div style={{display:"flex",flexDirection:"column",gap:8}}>
                {aniversariantes.map(p=>{ const hoje=(p.aniversario||"").split("/")[0]===diaAtual; return (
                  <div key={p.id} style={{display:"flex",alignItems:"center",gap:10,fontSize:13}}>
                    <Avatar name={p.name} size={26}/>
                    <span style={{fontWeight:600,color:T.ink,flex:1}}>{p.apelido||p.name}</span>
                    <span style={{fontSize:12,color:hoje?T.brand:T.muted,fontWeight:hoje?700:500}}>{p.aniversario}{hoje?" · hoje!":""}</span>
                  </div>
                );})}
              </div>}
        </Card>
      </div>

      {/* Mural da semana */}
      <Card style={{padding:0,overflow:"hidden",marginBottom:18}}>
        <div style={{display:"flex",alignItems:"center",gap:8,padding:"12px 16px",borderBottom:`1px solid ${T.line}`}}>
          <span style={{fontSize:14,fontWeight:700,fontFamily:T.fontDisplay,color:T.ink,flex:1,display:"inline-flex",alignItems:"center",gap:7}}><span style={{color:T.brand}}><Icon name="pin" size={18}/></span>Mural da semana</span>
          {isAdmin && <Btn small icon="pencil" onClick={()=>setEditing(true)}>Editar</Btn>}
        </div>
        <div style={{padding:"18px 20px"}}>
          {mural.frase
            ? <div style={{fontSize:18,fontWeight:700,color:T.ink,lineHeight:1.5,fontStyle:"italic"}}>“{mural.frase}”{mural.autor && <span style={{display:"block",fontSize:12,fontWeight:500,color:T.muted,fontStyle:"normal",marginTop:6}}>— {mural.autor}</span>}</div>
            : <div style={{fontSize:14,color:T.muted}}>{isAdmin?"Nenhuma frase ainda — clique em Editar para escrever a frase da semana.":"Bora fazer um mês incrível!"}</div>}
          {mural.lembretes?.length>0 && (
            <div style={{marginTop:16}}>
              <div style={{fontSize:12,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:".5px",marginBottom:8}}>Lembretes</div>
              <div style={{display:"flex",flexDirection:"column",gap:6}}>
                {mural.lembretes.map((l,i)=>(
                  <div key={i} style={{display:"flex",alignItems:"center",gap:9,fontSize:13,color:T.inkSoft,background:T.canvas,borderRadius:T.rMd,padding:"9px 12px"}}>
                    <span style={{color:T.brand}}>▸</span>{l}
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>
      </Card>

      {/* Pendências */}
      <div style={{fontSize:13,fontWeight:700,color:T.ink,margin:"0 2px 10px"}}>Seus atalhos de hoje</div>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(150px,1fr))",gap:12}}>
        <Pend icon="receipt" n={notasPend} label="notas a conciliar" color={T.brand} to="concil"/>
        <Pend icon="wallet" n={semNota} label="receitas sem nota" color={C.blue.solid} to="concil"/>
        <Pend icon="alert" n={faltamDatas} label="faturados sem datas" color={T.warn} to="time"/>
        <Pend icon="task" n={minhasTarefas} label="tarefas em aberto" color={T.ok} to="tasks"/>
      </div>
    </div>
  );
}

// ─── RELATÓRIOS ──────────────────────────────────────────────────────────────
const compRank = (c) => { const [m,y]=String(c||"").split("/"); return (y||"0000")+String(m||"").padStart(2,"0"); };
// Notas ligadas a um registro. cidsByRec: record_id → Set de conciliacao_id (livro
// de faturamento). Fallback: vínculo antigo por conciliacao_id/municipal_note_id.
function notesForRecord(r, notes, cidsByRec) {
  const cids = cidsByRec && cidsByRec[r.id];
  if (cids && cids.size) return notes.filter(n => n.conciliacaoId && cids.has(n.conciliacaoId));
  if (r.conciliacaoId) return notes.filter(n=>n.conciliacaoId===r.conciliacaoId);
  if (r.municipalNoteId) return notes.filter(n=>n.id===r.municipalNoteId);
  return [];
}

function ReportsView({ records, clients, notes, faturamentos=[], variacoes=[], varByRec={}, fatByRec={}, isAdmin, analistas }) {
  const cidsByRec = {}; faturamentos.forEach(a=>{ if(a.conciliacaoId){ (cidsByRec[a.recordId]=cidsByRec[a.recordId]||new Set()).add(a.conciliacaoId); } });
  const [tab, setTab] = useState("receitas");

  // ── Filtros de receitas ──
  const [empresa, setEmpresa] = useState("todas");
  const [analista, setAnalista] = useState("todos");
  const [tipo, setTipo] = useState("todos");
  const [status, setStatus] = useState("todos");
  const [concil, setConcil] = useState("todas");
  const [compDe, setCompDe] = useState("todas");
  const [compAte, setCompAte] = useState("todas");
  const [dim, setDim] = useState("servico");   // competência: mês de serviço × ciclo de faturamento
  const [qCli, setQCli] = useState("");
  const [qProf, setQProf] = useState("");

  // Período quebrado: deriva a competência de faturamento (ciclo) do dia de corte
  // do cliente + a data de início da receita (peça >= dia de corte fatura no mês
  // seguinte). Cliente sem corte = mês de serviço.
  const diaCorteDe = (r) => diaCorteOf(r, clients);
  const compFat = (r) => compFatOf(r, clients);
  const compValue = (r) => dim==="ciclo" ? compFat(r) : (r.competencia || "");

  const comps = [...new Set(records.map(compValue).filter(Boolean))].sort((a,b)=>compRank(a).localeCompare(compRank(b)));
  const tipos = [...new Set(records.map(r=>r.tipo).filter(Boolean))].sort();

  let recFiltered = records;
  if (empresa!=="todas") recFiltered = recFiltered.filter(r=>r.empresa===empresa);
  if (analista!=="todos") recFiltered = recFiltered.filter(r=>r.responsavel===analista);
  if (tipo!=="todos") recFiltered = recFiltered.filter(r=>r.tipo===tipo);
  if (compDe!=="todas") recFiltered = recFiltered.filter(r=>compRank(compValue(r))>=compRank(compDe));
  if (compAte!=="todas") recFiltered = recFiltered.filter(r=>compRank(compValue(r))<=compRank(compAte));
  if (status==="_faltam_datas") recFiltered = recFiltered.filter(faltaDatas);
  else if (status!=="todos") recFiltered = recFiltered.filter(r=>recStatus(r, fatByRec[r.id], (r.valorTotal||0)+(varByRec[r.id]||0))===status);
  const temNota = r => (cidsByRec[r.id] && cidsByRec[r.id].size) || r.conciliacaoId || r.municipalNoteId;
  if (concil==="conciliado") recFiltered = recFiltered.filter(temNota);
  if (concil==="sem_nota") recFiltered = recFiltered.filter(r=>!temNota(r));
  if (qCli.trim()) { const s=qCli.trim().toLowerCase(); recFiltered = recFiltered.filter(r=>(r.cliente||"").toLowerCase().includes(s)); }
  if (qProf.trim()) { const s=qProf.trim().toLowerCase(); recFiltered = recFiltered.filter(r=>(r.profissional||"").toLowerCase().includes(s)); }

  // Faturamento parcial → DUAS linhas do consultor: o que já foi faturado
  // (status "Faturado", com a NF) e o saldo pendente (status a faturar, sem NF).
  // Assim o contábil vê exatamente quanto está faturado e quanto falta.
  const linhasDe = (r) => {
    const p=r.progress||{};
    const fat = fatByRec[r.id]||0;
    const bill = (r.valorTotal||0) + (varByRec[r.id]||0);
    const saldo = bill - fat;
    const ns=notesForRecord(r, notes, cidsByRec);
    const baseA=[r.responsavel,r.empresa,r.tipo,r.competencia,compFat(r),r.codCliente,r.cliente,r.pep,r.profissional,r.ordemVenda||""];
    let nfNum="", nfEm="", nfVal="", nfMun="";
    if (ns.length===1) { nfNum=ns[0].numero; nfEm=toDate(ns[0].emitidaEm); nfVal=ns[0].valorServicos||0; nfMun=ns[0].municipio||""; }
    else if (ns.length>1) { nfNum=ns.map(n=>n.numero).join(", "); }
    const funilBase=[p.p1_extrair?"S":"N",p.p2_racional?"S":"N",p.p3_retorno_com?"S":"N",toDate(p.p3_data_retorno),p.p4_aprovacao?"S":"N",toDate(p.p4_data_aprov)];
    const funilFat=[...funilBase,"S","S",r.obs||"",r.conciliadoPor||"",toDate(r.conciliadoEm)];       // P5 NF=S
    const funilSaldo=[...funilBase,"N","N",r.obs||"","",""];                                            // saldo: NF pendente
    // linha: [...baseA, vVenda, hrs, valTotal, variação, status, NF..., funil]
    const linha = (vVenda,hrs,valTot,varc,status,nf,funil) => [...baseA,vVenda,hrs,valTot,varc,status,nf.num,nf.em,nf.val,nf.mun,...funil];
    const nfCheia={num:nfNum,em:nfEm,val:nfVal,mun:nfMun}, nfVazia={num:"",em:"",val:"",mun:""};
    const out=[];
    const temFat = Math.abs(fat)>0.01, temSaldo = Math.abs(saldo)>0.01;
    const stSaldo = p.p5_liberado ? "Liberado para faturamento" : calcStatus(p);
    if (temFat)   out.push(linha(r.valorVenda||0,r.hrsAprovadas||0,fat,varByRec[r.id]||0,"Faturado",nfCheia,funilFat));
    if (temSaldo) out.push(linha(temFat?"":(r.valorVenda||0),temFat?"":(r.hrsAprovadas||0),saldo,temFat?"":(varByRec[r.id]||0),stSaldo,nfVazia,funilSaldo));
    if (!temFat && !temSaldo) out.push(linha(r.valorVenda||0,r.hrsAprovadas||0,r.valorTotal||0,varByRec[r.id]||0,calcStatus(p),nfVazia,funilSaldo));
    return out;
  };
  function buildRecRows() { const rows=[]; recFiltered.forEach(r=>rows.push(...linhasDe(r))); return rows; }
  const previewLines = recFiltered.reduce((s,r)=>s+linhasDe(r).length,0);

  function exportReceitas() {
    const headers=["Analista","Empresa","Tipo","Competência (serviço)","Compet. faturamento (cliente)","Cód Cliente","Cliente","PEP","Profissional","Ordem de venda","Val. Venda","Hrs","Val. Total","Variação pós-fecham.","Status","NF Número","NF Emissão","NF Valor","NF Município","P1 Extração","P2 Racional","P3 Retorno com.","Data Retorno","P4 Aprov. cliente","Data Aprovação","P5 NF","Faturado corte","Obs","Conciliado por","Conciliado em"];
    downloadXLSX(`Relatorio_Receitas_${previewLines}linhas.xlsx`, headers, buildRecRows());
  }

  // ── Filtros de clientes ──
  const [cq, setCq] = useState("");
  const [cGrupo, setCGrupo] = useState("");
  const [cStatus, setCStatus] = useState("todos");
  const [cOwner, setCOwner] = useState("todos");
  const [cPortal, setCPortal] = useState("todos");
  const gruposEmp = [...new Set(clients.map(c=>(c.grupoEmpresa||"").trim()).filter(Boolean))].sort();
  const owners = [...new Set(clients.map(c=>(c.owner||"").trim()).filter(Boolean))].sort();

  let cliFiltered = clients;
  if (cStatus==="incompletos") cliFiltered = cliFiltered.filter(c=>c.incompleto);
  if (cStatus==="completos")   cliFiltered = cliFiltered.filter(c=>!c.incompleto);
  if (cGrupo.trim()) { const g=cGrupo.trim().toLowerCase(); cliFiltered = cliFiltered.filter(c=>(c.grupoEmpresa||"").toLowerCase().includes(g)); }
  if (cOwner!=="todos") cliFiltered = cliFiltered.filter(c=>(c.owner||"")===cOwner);
  if (cPortal==="com") cliFiltered = cliFiltered.filter(c=>c.temPortal);
  if (cPortal==="sem") cliFiltered = cliFiltered.filter(c=>!c.temPortal);
  if (cq.trim()) { const s=cq.trim().toLowerCase(); const dig=s.replace(/\D/g,""); cliFiltered = cliFiltered.filter(c=>(c.nome||"").toLowerCase().includes(s)||(c.codSap||"").toLowerCase().includes(s)||(!!dig&&clientCnpjs(c).some(x=>x.includes(dig)))); }

  function exportClientes() {
    const headers=["Nome/Grupo","Cód SAP","CNPJs","Grupo de empresa","Analista responsável","Status cadastro","Tipos de contrato","PEPs","Propostas","Período faturamento","Tem portal","Classificação portal","Prazo vencimento","Forma de pagamento","Contato financeiro","E-mail financeiro","Account manager","E-mail AM"];
    // CNPJ/Cód SAP como texto (preserva zeros à esquerda no Excel).
    const rows = cliFiltered.map(c=>[c.nome,c.codSap,clientCnpjs(c).join(" ; "),c.grupoEmpresa,c.owner,c.incompleto?"Incompleto":"Completo",c.tiposContrato,fmtPepsCSV(c.tiposPeps),fmtPropostasCSV(c.propostas),c.periodoFaturamento,c.temPortal?"Sim":"Não",c.portalTipo,c.prazoVencimento,c.formaPagamento,c.contatoFinanceiro,c.contatoFinanceiroEmail,c.accountManager,c.accountManagerEmail]);
    downloadXLSX(`Relatorio_Clientes_${rows.length}.xlsx`, headers, rows);
  }

  const Sel = ({label,value,onChange,children}) => <Field label={label}><select style={inp} value={value} onChange={e=>onChange(e.target.value)}>{children}</select></Field>;

  return (
    <div>
      <PageHead icon="file" title="Relatórios" sub="Extraia relatórios em Excel (.xlsx) já formatados — números como número e datas em dd/mm/aaaa."/>

      <div style={{display:"flex",gap:6,borderBottom:`1px solid ${T.line}`,marginBottom:18}}>
        {[["receitas","Receitas"],["clientes","Clientes"]].map(([id,label])=>(
          <button key={id} onClick={()=>setTab(id)} style={{border:"none",background:"none",cursor:"pointer",padding:"8px 14px",fontSize:13,fontWeight:tab===id?700:500,color:tab===id?T.brand:T.muted,borderBottom:`2px solid ${tab===id?T.brand:"transparent"}`,marginBottom:-1}}>{label}</button>
        ))}
      </div>

      {tab==="receitas" && <>
        <Card style={{padding:"14px 16px",marginBottom:14}}>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(160px,1fr))",gap:12}}>
            <Sel label="Empresa" value={empresa} onChange={setEmpresa}><option value="todas">Todas</option>{EMPRESAS.map(e=><option key={e.cod} value={e.cod}>{e.cod} — {e.nome}</option>)}</Sel>
            {isAdmin && <Sel label="Analista" value={analista} onChange={setAnalista}><option value="todos">Todos</option>{analistas.map(a=><option key={a}>{a}</option>)}</Sel>}
            <Sel label="Tipo de contrato" value={tipo} onChange={setTipo}><option value="todos">Todos</option>{tipos.map(t=><option key={t}>{t}</option>)}</Sel>
            <Sel label="Visão da competência" value={dim} onChange={v=>{setDim(v);setCompDe("todas");setCompAte("todas");}}><option value="servico">Mês de serviço (Fcamara)</option><option value="ciclo">Ciclo de faturamento (cliente)</option></Sel>
            <Sel label="Competência (de)" value={compDe} onChange={setCompDe}><option value="todas">Início</option>{comps.map(c=><option key={c}>{c}</option>)}</Sel>
            <Sel label="Competência (até)" value={compAte} onChange={setCompAte}><option value="todas">Fim</option>{comps.map(c=><option key={c}>{c}</option>)}</Sel>
            <Sel label="Status" value={status} onChange={setStatus}><option value="todos">Todos</option>{STATUS_ORDER.map(s=><option key={s}>{s}</option>)}<option value="_faltam_datas">Faltam datas</option></Sel>
            <Sel label="Conciliação" value={concil} onChange={setConcil}><option value="todas">Todas</option><option value="conciliado">Conciliadas</option><option value="sem_nota">Sem nota</option></Sel>
            <Field label="Cliente"><input style={inp} placeholder="nome" value={qCli} onChange={e=>setQCli(e.target.value)}/></Field>
            <Field label="Profissional"><input style={inp} placeholder="nome" value={qProf} onChange={e=>setQProf(e.target.value)}/></Field>
          </div>
        </Card>
        <Card style={{padding:"14px 16px",display:"flex",alignItems:"center",gap:14,flexWrap:"wrap"}}>
          <div style={{fontSize:13}}><b>{recFiltered.length}</b> receita(s) · <b>{previewLines}</b> linha(s) no relatório <span style={{color:T.muted,fontSize:11}}>(faturamento parcial gera 2 linhas: valor faturado + saldo a faturar)</span></div>
          <div style={{flex:1}}/>
          <Btn primary icon="download" disabled={!recFiltered.length} onClick={exportReceitas}>Exportar receitas (.xlsx)</Btn>
        </Card>
      </>}

      {tab==="clientes" && <>
        <Card style={{padding:"14px 16px",marginBottom:14}}>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(160px,1fr))",gap:12}}>
            <Field label="Cliente / Cód SAP / CNPJ"><input style={inp} placeholder="buscar" value={cq} onChange={e=>setCq(e.target.value)}/></Field>
            <Field label="Grupo de empresa"><input style={inp} list="rep-grupos" placeholder="Todos" value={cGrupo} onChange={e=>setCGrupo(e.target.value)}/><datalist id="rep-grupos">{gruposEmp.map(g=><option key={g} value={g}/>)}</datalist></Field>
            <Sel label="Status do cadastro" value={cStatus} onChange={setCStatus}><option value="todos">Todos</option><option value="completos">Completos</option><option value="incompletos">Incompletos</option></Sel>
            {isAdmin && <Sel label="Analista responsável" value={cOwner} onChange={setCOwner}><option value="todos">Todos</option>{owners.map(o=><option key={o}>{o}</option>)}</Sel>}
            <Sel label="Portal" value={cPortal} onChange={setCPortal}><option value="todos">Todos</option><option value="com">Com portal</option><option value="sem">Sem portal</option></Sel>
          </div>
        </Card>
        <Card style={{padding:"14px 16px",display:"flex",alignItems:"center",gap:14,flexWrap:"wrap"}}>
          <div style={{fontSize:13}}><b>{cliFiltered.length}</b> cliente(s) no relatório</div>
          <div style={{flex:1}}/>
          <Btn primary icon="download" disabled={!cliFiltered.length} onClick={exportClientes}>Exportar clientes (.xlsx)</Btn>
        </Card>
      </>}
    </div>
  );
}

// ─── IMPORTAR DOCUMENTOS (admin) ─────────────────────────────────────────────

function DataIOView({ recordsCount, clientsCount, onImport, onExport, onHistory, onExportClients }) {
  const Tile = ({ icon, title, desc, btn, onClick, primary }) => (
    <Card style={{ padding:"18px 20px", display:"flex", flexDirection:"column" }}>
      <div style={{ fontSize:26, marginBottom:8 }} aria-hidden="true">{icon}</div>
      <div style={{ fontSize:15, fontWeight:700, color:T.ink }}>{title}</div>
      <div style={{ fontSize:12, color:T.muted, margin:"4px 0 14px", lineHeight:1.5, flex:1 }}>{desc}</div>
      <div><Btn primary={primary} onClick={onClick}>{btn}</Btn></div>
    </Card>
  );
  return (
    <div>
      <PageHead icon="import" title="Importar documentos" sub={<>Carga de dados de reconhecimento. As exportações agora ficam na aba <b>Relatórios</b>.</>}/>
      <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(240px,1fr))", gap:14 }}>
        <Tile icon="" title="Importar receitas" primary desc={`Carregar a planilha de T&E / Fee-WIP / Usage (.xlsm/.xlsx). ${recordsCount} registro(s) hoje.`} btn="Importar planilha" onClick={onImport}/>
        <Tile icon="" title="Histórico de importações" desc="Ver o log de todas as importações realizadas." btn="Ver histórico" onClick={onHistory}/>
      </div>
    </div>
  );
}

// ─── TOPBAR ──────────────────────────────────────────────────────────────────

function Topbar({ user, isAdmin, isMobile, onMenu, onLogout, theme, onToggleTheme }) {
  const dark = theme==="dark";
  return (
    <div style={{background:T.dark,color:"#fff",padding:"0 16px",display:"flex",alignItems:"center",gap:12,height:56}}>
      {isMobile && <button onClick={onMenu} aria-label="Abrir menu" style={{ background:"none", border:"none", color:"#fff", cursor:"pointer", lineHeight:1, padding:4, display:"inline-flex" }}><Icon name="menu" size={22}/></button>}
      <span style={{flex:1,display:"flex",alignItems:"center",gap:11}}>
        <FcamaraLogo size={20} onDark/>
        <span style={{width:1,height:20,background:"rgba(255,255,255,.22)"}}/>
        <span style={{fontSize:14,fontWeight:600,fontFamily:T.fontDisplay,color:"#fff"}}>Order to Cash</span>
        {!isMobile && <span style={{fontSize:12,fontWeight:500,color:"rgba(255,255,255,.6)"}}>Grupo Fcamara</span>}
      </span>

      {!isMobile && <span style={{display:"flex",alignItems:"center",gap:8,fontSize:12,color:"rgba(255,255,255,.9)",paddingLeft:4}}>
        <Avatar name={user.name} size={28} admin={isAdmin}/>{user.name}{isAdmin?" · Admin":""}
      </span>}
      <button className="fc-btn" onClick={onToggleTheme} title={dark?"Mudar para modo claro":"Mudar para modo escuro"} aria-label={dark?"Modo claro":"Modo escuro"} style={{ background:"rgba(255,255,255,.1)", border:"1px solid rgba(255,255,255,.18)", color:"#fff", borderRadius:T.rPill, width:32, height:32, display:"inline-flex", alignItems:"center", justifyContent:"center", cursor:"pointer" }}><Icon name={dark?"sun":"moon"} size={16}/></button>
      <button className="fc-btn" onClick={onLogout} style={{ background:"rgba(255,255,255,.1)", border:"1px solid rgba(255,255,255,.18)", color:"#fff", borderRadius:T.rMd, padding:"6px 14px", fontSize:12, fontWeight:600, cursor:"pointer" }}>Sair</button>
    </div>
  );
}

// ─── LOGIN ───────────────────────────────────────────────────────────────────

function ForgotPasswordModal({ onClose }) {
  const [email, setEmail] = useState("");
  const [sent, setSent]   = useState(false);
  const [err, setErr]     = useState("");
  const [busy, setBusy]   = useState(false);

  async function submit(e) {
    e?.preventDefault();
    if (!email.trim()) { setErr("Informe o e-mail do seu acesso."); return; }
    setBusy(true); setErr("");
    const { error } = await supabase.auth.resetPasswordForEmail(email.trim(), { redirectTo: SITE_URL });
    setBusy(false);
    if (error) setErr("Não foi possível enviar agora. Tente novamente em instantes."); else setSent(true);
  }

  return (
    <Modal title="Redefinir senha" subtitle="Enviaremos um link de redefinição para o seu e-mail" onClose={onClose}>
      {sent ? (
        <div>
          <div style={{padding:"12px 14px",borderRadius:T.rMd,background:T.okBg,border:`1px solid ${T.okLine}`,color:T.ok,fontSize:13,marginBottom:16}}>
            ✓ Link enviado para <b>{email}</b>. Abra o e-mail e clique no link para escolher uma nova senha.
          </div>
          <div style={{display:"flex",justifyContent:"flex-end"}}><Btn primary onClick={onClose}>Entendi</Btn></div>
        </div>
      ) : (
        <form onSubmit={submit}>
          <div style={{marginBottom:14}}><Field label="E-mail *"><input style={inp} type="email" value={email} onChange={e=>{setEmail(e.target.value);setErr("");}} placeholder="seu.email@empresa.com" autoFocus/></Field></div>
          {err&&<div style={{marginBottom:12,fontSize:12,padding:"8px 12px",borderRadius:T.rMd,background:T.dangerBg,color:T.danger,border:`1px solid ${T.dangerLine}`}}>{err}</div>}
          <div style={{fontSize:11,color:T.muted,marginBottom:16}}>Se você não reconhece nenhum acesso, fale com a administração (Daniela ou Luana).</div>
          <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
            <Btn onClick={onClose}>Cancelar</Btn>
            <Btn primary onClick={submit} disabled={busy}>{busy ? "Enviando..." : "Enviar link"}</Btn>
          </div>
        </form>
      )}
    </Modal>
  );
}

// Tela exibida quando o usuário chega pelo link de redefinição de senha.
function RecoveryModal({ onClose }) {
  const [pass, setPass]   = useState("");
  const [conf, setConf]   = useState("");
  const [err, setErr]     = useState("");
  const [busy, setBusy]   = useState(false);

  async function save(e) {
    e?.preventDefault();
    if (pass.length < 6) { setErr("A nova senha precisa ter ao menos 6 caracteres."); return; }
    if (pass !== conf)   { setErr("A confirmação não confere."); return; }
    setBusy(true); setErr("");
    const { error } = await supabase.auth.updateUser({ password: pass });
    setBusy(false);
    if (error) setErr(error.message); else onClose();
  }

  return (
    <Modal title="Escolher nova senha" subtitle="Defina a senha que você usará para entrar" onClose={onClose}>
      <form onSubmit={save}>
        <div style={{marginBottom:14}}><Field label="Nova senha *"><input style={inp} type="password" value={pass} onChange={e=>{setPass(e.target.value);setErr("");}} placeholder="Mínimo 6 caracteres" autoFocus/></Field></div>
        <div style={{marginBottom:16}}><Field label="Confirmar nova senha *"><input style={inp} type="password" value={conf} onChange={e=>{setConf(e.target.value);setErr("");}} placeholder="Repita a nova senha"/></Field></div>
        {err&&<div style={{marginBottom:12,fontSize:12,padding:"8px 12px",borderRadius:T.rMd,background:T.dangerBg,color:T.danger,border:`1px solid ${T.dangerLine}`}}>{err}</div>}
        <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
          <Btn primary onClick={save} disabled={busy}>{busy ? "Salvando..." : "Salvar senha"}</Btn>
        </div>
      </form>
    </Modal>
  );
}

function Login() {
  const [email, setEmail]  = useState("");
  const [pass, setPass]    = useState("");
  const [loginErr, setLE]  = useState("");
  const [busy, setBusy]    = useState(false);
  const [showForgot, setSF]= useState(false);

  async function submit(e) {
    e.preventDefault();
    setBusy(true); setLE("");
    const { error } = await supabase.auth.signInWithPassword({ email: email.trim(), password: pass });
    setBusy(false);
    if (error) setLE("E-mail ou senha incorretos.");
    // Em caso de sucesso, o onAuthStateChange no AppInner cuida de entrar.
  }

  return (
    <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:`linear-gradient(135deg,#201b18,${T.dark})`,fontFamily:T.font,padding:16}}>
      {showForgot && <ForgotPasswordModal onClose={()=>setSF(false)}/>}
      <div style={{background:"var(--surface)",borderRadius:18,padding:"34px 38px",width:400,maxWidth:"100%",boxShadow:T.shLg}}>
        <div style={{textAlign:"center",marginBottom:24}}>
          <div style={{ display:"inline-flex", marginBottom:14 }}><FcamaraLogo size={34}/></div>
          <h1 style={{fontSize:22,fontWeight:700,fontFamily:T.fontDisplay,color:T.ink,lineHeight:1.3,margin:0}}>Order to Cash</h1>
          <p style={{fontSize:13,color:T.muted,fontWeight:500,marginTop:4,marginBottom:0}}>Grupo Fcamara · time O2C</p>
          <p style={{fontSize:11,color:T.faint,marginTop:6}}>{APP_VERSION}</p>
        </div>
        <form onSubmit={submit}>
          <div style={{marginBottom:12}}><Field label="E-mail"><input style={inp} type="email" placeholder="seu.email@empresa.com" value={email} onChange={e=>{setEmail(e.target.value);setLE("");}} autoFocus/></Field></div>
          <div style={{marginBottom:8}}><Field label="Senha"><input style={inp} type="password" placeholder="Sua senha" value={pass} onChange={e=>{setPass(e.target.value);setLE("");}}/></Field></div>
          <div style={{textAlign:"right",marginBottom:16}}>
            <button type="button" onClick={()=>setSF(true)} style={{background:"none",border:"none",padding:0,fontSize:12,color:T.brand,fontWeight:600,cursor:"pointer"}}>Esqueci minha senha</button>
          </div>
          {loginErr&&<div style={{marginBottom:12,fontSize:12,padding:"8px 12px",borderRadius:T.rMd,background:T.dangerBg,color:T.danger,border:`1px solid ${T.dangerLine}`}}>{loginErr}</div>}
          <button type="submit" disabled={busy} className="fc-btn" style={{width:"100%",padding:"11px",borderRadius:T.rMd,border:"none",background:T.brand,color:"#fff",fontSize:14,fontWeight:700,cursor:busy?"wait":"pointer",opacity:busy?.7:1}}>{busy ? "Entrando..." : "Entrar"}</button>
        </form>
      </div>
    </div>
  );
}

// ─── CORREÇÕES (admin) ───────────────────────────────────────────────────────
// Painel único para ajustar a base sem SQL/Excel: busca qualquer registro (todos
// os meses), move PEP / corrige empresa (edição in loco — muda o próprio registro,
// sem gerar órfão), apaga lixo/fantasma e mescla duplicatas. Toda ação destrutiva
// mostra a prévia antes→depois do total do recorte e permite desfazer.

// Mesclar duplicata: escolhe qual linha MANTER; a outra (a de origem) é removida.
function MergeModal({ source, records, fatByRec={}, onConfirm, onClose }) {
  const [q, setQ] = useState("");
  const [targetId, setTargetId] = useState(null);
  const nrm = s => (s||"").toString().toLowerCase();
  // Candidatos naturais: mesmo mês+empresa+tipo+cliente, id diferente.
  const cands = records.filter(r => r.id!==source.id
    && r.competencia===source.competencia && r.empresa===source.empresa
    && r.tipo===source.tipo && nrm(r.cliente)===nrm(source.cliente));
  const base = q.trim()
    ? records.filter(r => r.id!==source.id && [r.cliente,r.pep,r.profissional,r.responsavel].some(x=>nrm(x).includes(nrm(q.trim()))))
    : cands;
  const lista = base.slice(0,40);
  const target = records.find(r=>r.id===targetId);
  const diff = target ? Math.abs((source.valorTotal||0)-(target.valorTotal||0)) : 0;
  const linha = r => `${r.pep||"—"} · ${r.profissional||"(sem profissional)"} · ${brl(r.valorTotal)}`;
  return (
    <Modal title="Mesclar duplicata" subtitle="Escolha a linha que fica; a de origem é removida." onClose={onClose} wide>
      <div style={{fontSize:12.5,color:T.inkSoft,background:C.orange.bg,border:`1px solid ${C.orange.border}`,borderRadius:T.rMd,padding:"9px 12px",marginBottom:14}}>
        <b>Vai remover</b> (origem): {linha(source)} · {source.competencia} · {source.empresa}
      </div>
      <Field label="Buscar a linha que MANTÉM (cliente, PEP, profissional)">
        <input style={inp} autoFocus placeholder={cands.length?`${cands.length} duplicata(s) provável(is) — ou busque`:"Digite para buscar"} value={q} onChange={e=>setQ(e.target.value)}/>
      </Field>
      <div style={{maxHeight:230,overflowY:"auto",border:`1px solid ${T.line}`,borderRadius:T.rMd,marginTop:10}}>
        {lista.length===0 && <div style={{padding:"14px",fontSize:12.5,color:T.muted,textAlign:"center"}}>Nenhum candidato. Ajuste a busca.</div>}
        {lista.map(r=>(
          <label key={r.id} style={{display:"flex",alignItems:"center",gap:10,padding:"9px 12px",borderBottom:`1px solid ${T.line}`,cursor:"pointer",background:targetId===r.id?T.brandBg:"var(--surface)"}}>
            <input type="radio" name="mergeTarget" checked={targetId===r.id} onChange={()=>setTargetId(r.id)}/>
            <span style={{flex:1,fontSize:12.5}}><b>{r.pep||"—"}</b> · {r.profissional||"(sem profissional)"} · {r.tipo} · {r.competencia}</span>
            <span style={{fontSize:12,fontWeight:700,whiteSpace:"nowrap"}}>{brl(r.valorTotal)}</span>
          </label>
        ))}
      </div>
      {target && diff>0.01 && <div style={{marginTop:12,fontSize:12,color:C.red.text,background:C.red.bg,border:`1px solid ${C.red.border}`,borderRadius:T.rMd,padding:"9px 12px"}}>
        ⚠️ Os valores diferem ({brl(source.valorTotal)} × {brl(target.valorTotal)}). Ao mesclar, o total do recorte cai <b>{brl(source.valorTotal)}</b> (o valor da origem removida). Confirme que é mesmo duplicata.
      </div>}
      {target && diff<=0.01 && <div style={{marginTop:12,fontSize:12,color:C.green.text,background:C.green.bg,border:`1px solid ${C.green.border}`,borderRadius:T.rMd,padding:"9px 12px"}}>
        ✓ Duplicata exata — mantém {linha(target)} e remove a origem. Total do recorte cai {brl(source.valorTotal)}.
      </div>}
      <div style={{display:"flex",gap:8,justifyContent:"flex-end",marginTop:16}}>
        <Btn onClick={onClose}>Cancelar</Btn>
        <Btn danger disabled={!target} onClick={()=>{ onConfirm(source.id); onClose(); }}>Mesclar (remover origem)</Btn>
      </div>
    </Modal>
  );
}

function CorrectionsView({ records, fatByRec={}, onEdit, onDelete, onMerge, onInsert, lastCorr, onUndo }) {
  const [q,setQ]=useState("");
  const [comp,setComp]=useState("todas");
  const [emp,setEmp]=useState("todas");
  const [tipo,setTipo]=useState("todos");
  const [status,setStatus]=useState("todos");
  const [editT,setEditT]=useState(null);
  const [delT,setDelT]=useState(null);
  const [mergeT,setMergeT]=useState(null);
  const [incT,setIncT]=useState(null);

  const nrm = s => (s||"").toString().toLowerCase();
  const isConc = r => (fatByRec[r.id]||0)>0.001;
  const comps = [...new Set(records.map(r=>r.competencia).filter(Boolean))]
    .sort((a,b)=>{ const [ma,ya]=String(a).split("/"), [mb,yb]=String(b).split("/"); return (Number(yb)-Number(ya))||(Number(mb)-Number(ma)); });
  // Registro em branco para "Incluir" — prefila competência/empresa pelos filtros ativos.
  const blankRec = () => ({ responsavel:"", empresa: emp!=="todas"?emp:"BR02", tipo: tipo!=="todos"?tipo:"Time & Expenses",
    competencia: comp!=="todas"?comp:(comps[0]||""), codCliente:"", cliente:"", pep:"", profissional:"",
    inicio:"", fim:"", valorVenda:0, hrsAprovadas:0, valorTotal:0, valorLiquido:0, obs:"", progress:{}, ausenteRelatorio:false });
  // Duplicar: cópia dos dados de uma linha SEM id (vira um novo registro no
  // "Incluir"). Não copia conciliação/progresso — nasce limpo para editar.
  const dupRec = (r) => ({ responsavel:r.responsavel||"", empresa:r.empresa, tipo:r.tipo, bu:r.bu||"",
    competencia:r.competencia, codCliente:r.codCliente||"", cliente:r.cliente||"", pep:r.pep||"", profissional:r.profissional||"",
    inicio:r.inicio||"", fim:r.fim||"", valorVenda:r.valorVenda||0, hrsAprovadas:r.hrsAprovadas||0,
    valorTotal:r.valorTotal||0, valorLiquido:r.valorLiquido||0, obs:r.obs||"", progress:{}, ausenteRelatorio:false });

  let list = records;
  if (comp!=="todas") list=list.filter(r=>r.competencia===comp);
  if (emp!=="todas")  list=list.filter(r=>r.empresa===emp);
  if (tipo!=="todos") list=list.filter(r=>r.tipo===tipo);
  if (status==="conciliado") list=list.filter(isConc);
  else if (status==="ausente") list=list.filter(r=>r.ausenteRelatorio);
  else if (status==="normal")  list=list.filter(r=>!isConc(r)&&!r.ausenteRelatorio);
  const term = q.trim();
  if (term) list=list.filter(r=>[r.cliente,r.pep,r.profissional,r.responsavel,r.codCliente].some(x=>nrm(x).includes(nrm(term))));
  const total = list.length;
  const shown = list.slice(0,200);
  const somaShown = shown.reduce((s,r)=>s+(r.valorTotal||0),0);

  // Total do recorte competência+empresa (espelha o soma_total do fingerprint).
  const compTotal = (c,e) => records.filter(r=>r.competencia===c&&r.empresa===e).reduce((s,r)=>s+(r.valorTotal||0),0);
  const delMsg = (r) => {
    const antes = compTotal(r.competencia,r.empresa), depois = antes-(r.valorTotal||0);
    return `Apagar "${r.profissional||r.pep||"registro"}" (${r.cliente}) — ${brl(r.valorTotal)}.\n\n`
      + `Total de ${r.competencia} · ${r.empresa}: ${brl(antes)} → ${brl(depois)}.`
      + (r.ausenteRelatorio ? "\n\n(Está marcado como 'fora do relatório'.)" : "");
  };

  const th = { padding:"7px 10px", textAlign:"left", fontSize:11, fontWeight:700, color:T.muted, textTransform:"uppercase", letterSpacing:".3px", whiteSpace:"nowrap", borderBottom:`1px solid ${T.line}` };
  const td = { padding:"7px 10px", fontSize:12.5, borderBottom:`1px solid ${T.line}`, verticalAlign:"middle" };

  return (
    <div>
      {editT && <RecordEditModal record={editT} conciliado={isConc(editT)} onClose={()=>setEditT(null)} onSave={r=>{onEdit(r);setEditT(null);}}/>}
      {incT && <RecordEditModal record={incT} novo conciliado={false} onClose={()=>setIncT(null)} onSave={r=>{onInsert(r);setIncT(null);}}/>}
      {mergeT && <MergeModal source={mergeT} records={records} fatByRec={fatByRec} onClose={()=>setMergeT(null)} onConfirm={id=>{onMerge(id);setMergeT(null);}}/>}
      {delT && <ConfirmDialog title="Apagar registro" danger confirmLabel="Apagar" message={delMsg(delT)} onConfirm={()=>onDelete(delT.id)} onClose={()=>setDelT(null)}/>}

      <PageHead icon="pencil" title="Correções" sub="Ajuste a base direto no app — incluir, mover PEP, corrigir empresa, apagar e mesclar. Com prévia e desfazer."
        right={<Btn primary icon="plus" onClick={()=>setIncT(blankRec())}>Incluir registro</Btn>}/>

      {lastCorr && <div style={{display:"flex",alignItems:"center",gap:10,padding:"9px 13px",borderRadius:T.rMd,background:C.green.bg,border:`1px solid ${C.green.border}`,marginBottom:14}}>
        <Icon name="check" size={15}/>
        <span style={{flex:1,fontSize:12.5,color:C.green.text}}>Correção aplicada — <b>{lastCorr.label}</b>.</span>
        <Btn small onClick={onUndo}>↶ Desfazer</Btn>
      </div>}

      <Card style={{padding:14,marginBottom:14}}>
        <div style={{display:"flex",gap:10,flexWrap:"wrap",alignItems:"flex-end"}}>
          <Field label="Buscar (cliente, PEP, profissional, responsável, cód.)">
            <input style={{...inp,minWidth:280}} placeholder="Ex.: RUMO, BR02CLP00128, valor fixo…" value={q} onChange={e=>setQ(e.target.value)}/>
          </Field>
          <Field label="Competência"><select style={{...inp,width:"auto"}} value={comp} onChange={e=>setComp(e.target.value)}><option value="todas">Todas</option>{comps.map(c=><option key={c}>{c}</option>)}</select></Field>
          <Field label="Empresa"><select style={{...inp,width:"auto"}} value={emp} onChange={e=>setEmp(e.target.value)}><option value="todas">Todas</option>{EMPRESAS.map(e=><option key={e.cod} value={e.cod}>{e.cod}</option>)}</select></Field>
          <Field label="Tipo"><select style={{...inp,width:"auto"}} value={tipo} onChange={e=>setTipo(e.target.value)}><option value="todos">Todos</option>{TIPOS_PROJETO.map(t=><option key={t}>{t}</option>)}</select></Field>
          <Field label="Status"><select style={{...inp,width:"auto"}} value={status} onChange={e=>setStatus(e.target.value)}><option value="todos">Todos</option><option value="normal">Normal</option><option value="conciliado">Conciliado</option><option value="ausente">Fora do relatório</option></select></Field>
        </div>
      </Card>

      <Card style={{padding:0,overflow:"hidden"}}>
        <div style={{padding:"9px 12px",fontSize:12,color:T.muted,borderBottom:`1px solid ${T.line}`,display:"flex",justifyContent:"space-between",flexWrap:"wrap",gap:8}}>
          <span><b style={{color:T.ink}}>{total}</b> registro(s){total>200?" · mostrando os 200 primeiros":""}</span>
          <span>Soma exibida: <b style={{color:T.ink}}>{brl(somaShown)}</b></span>
        </div>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse"}}>
            <thead><tr>{["Comp.","Empr.","Cliente","Tipo","PEP","Profissional","Valor","Status","Ações"].map(h=><th key={h} style={th}>{h}</th>)}</tr></thead>
            <tbody>
              {shown.length===0 && <tr><td colSpan={9} style={{padding:"22px 12px",textAlign:"center",color:T.muted,fontSize:13}}>Nenhum registro. Ajuste a busca ou os filtros.</td></tr>}
              {shown.map(r=>{
                const conc = isConc(r);
                return (
                  <tr key={r.id}>
                    <td style={td}>{r.competencia}</td>
                    <td style={td}>{r.empresa}</td>
                    <td style={{...td,maxWidth:200,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={r.cliente}>{r.cliente}</td>
                    <td style={td}>{r.tipo}</td>
                    <td style={{...td,whiteSpace:"nowrap",fontWeight:600}}>{r.pep||"—"}</td>
                    <td style={{...td,maxWidth:180,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={r.profissional}>{r.profissional||"—"}</td>
                    <td style={{...td,whiteSpace:"nowrap",fontWeight:600}}>{brl(r.valorTotal)}</td>
                    <td style={td}>
                      {conc ? <Badge label="Conciliado" color="blue" small/>
                        : r.ausenteRelatorio ? <Badge label="Fora do relat." color="orange" small/>
                        : <Badge label="Normal" color="gray" small/>}
                    </td>
                    <td style={{...td,whiteSpace:"nowrap"}}>
                      <div style={{display:"flex",gap:5}}>
                        <Btn small icon="pencil" onClick={()=>setEditT(r)}>Editar</Btn>
                        <Btn small icon="plus" title="Duplicar esta linha para editar só o que muda" onClick={()=>setIncT(dupRec(r))}>Duplicar</Btn>
                        <Btn small icon="link" disabled={conc} title={conc?"Reabra a conciliação primeiro":"Mesclar com outra linha"} onClick={()=>!conc&&setMergeT(r)}>Mesclar</Btn>
                        <Btn small danger icon="trash" disabled={conc} title={conc?"Reabra a conciliação primeiro":"Apagar"} onClick={()=>!conc&&setDelT(r)}>Apagar</Btn>
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      </Card>
      <div style={{fontSize:11.5,color:T.faint,marginTop:10,lineHeight:1.5}}>
        <b>Editar</b> muda o próprio registro (mover PEP, corrigir empresa) sem criar fantasma. <b>Mesclar</b> remove uma duplicata mantendo a outra. Registros <b>conciliados</b> ficam travados para apagar/mesclar — reabra a conciliação antes. Toda ação pode ser desfeita pelo botão acima.
      </div>
    </div>
  );
}

// ─── CLASSIFICAR BU (admin) ──────────────────────────────────────────────────
// Etiqueta cada cliente com sua unidade de negócio. Aplica em massa em todas as
// receitas do cliente (seguro para conciliados — só grava o campo bu). Casos
// raros de cliente com >1 BU: ajuste fino por registro em Correções.
function BuClassifierView({ records, onSetBu }) {
  const [q, setQ] = useState("");
  const [soSem, setSoSem] = useState(false);
  const [emp, setEmp] = useState("todas");
  const key = s => (s||"").toString().toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g,"").replace(/[^a-z0-9]/g,"");
  const empresasComDados = [...new Set(records.map(r=>r.empresa).filter(Boolean))].sort();
  const grupos = {};
  records.forEach(r => {
    if (emp!=="todas" && r.empresa!==emp) return;   // escopo por empresa
    const k = key(r.cliente); if(!k) return;
    (grupos[k] = grupos[k] || { nome:r.cliente, ids:[], total:0, bus:new Set(), emps:new Set() });
    grupos[k].ids.push(r.id); grupos[k].total += (r.valorTotal||0); grupos[k].bus.add(r.bu||""); grupos[k].emps.add(r.empresa||"");
  });
  const all = Object.values(grupos).map(g => {
    const arr=[...g.bus];
    const buAtual = arr.filter(Boolean).length===0 ? "" : (arr.length===1 ? arr[0] : "__MISTO__");
    return { ...g, buAtual, empLabel:[...g.emps].filter(Boolean).sort().join(", ") };
  }).sort((a,b)=>b.total-a.total);
  const totalCli = all.length;
  const feitos = all.filter(g=>g.buAtual && g.buAtual!=="__MISTO__").length;
  let lista = all;
  if (q.trim()) { const s=key(q); lista=lista.filter(g=>key(g.nome).includes(s)); }
  if (soSem) lista=lista.filter(g=>!g.buAtual || g.buAtual==="__MISTO__");
  const shown = lista.slice(0,400);
  const pct = totalCli ? Math.round(feitos/totalCli*100) : 0;
  const th={padding:"7px 10px",textAlign:"left",fontSize:11,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:".3px",borderBottom:`1px solid ${T.line}`};
  const td={padding:"7px 10px",fontSize:12.5,borderBottom:`1px solid ${T.line}`};
  return (
    <div>
      <PageHead icon="building" title="Classificar BU" sub="Etiquete cada cliente com sua unidade de negócio — aplica em todas as receitas do cliente de uma vez. Seguro para conciliados."/>
      <Card style={{padding:14,marginBottom:14}}>
        <div style={{display:"flex",alignItems:"center",gap:10,flexWrap:"wrap",marginBottom:12}}>
          <span style={{fontSize:13,fontWeight:700,whiteSpace:"nowrap"}}>{feitos} de {totalCli} clientes classificados</span>
          <div style={{flex:1,minWidth:120,height:8,background:T.lineSoft,borderRadius:999,overflow:"hidden"}}><div style={{width:`${pct}%`,height:"100%",background:T.brand,transition:"width .2s"}}/></div>
          <span style={{fontSize:12,color:T.muted}}>{pct}%</span>
        </div>
        <div style={{display:"flex",gap:10,flexWrap:"wrap",alignItems:"center"}}>
          <input style={{...inp,flex:1,minWidth:200}} placeholder="Buscar cliente…" value={q} onChange={e=>setQ(e.target.value)}/>
          <select style={{...inp,width:"auto"}} value={emp} onChange={e=>setEmp(e.target.value)}>
            <option value="todas">Todas as empresas</option>
            {empresasComDados.map(c=>{ const e=EMPRESAS.find(x=>x.cod===c); return <option key={c} value={c}>{c}{e?` — ${e.nome}`:""}</option>; })}
          </select>
          <label style={{display:"flex",alignItems:"center",gap:6,fontSize:12.5,color:T.inkSoft,cursor:"pointer",whiteSpace:"nowrap"}}><input type="checkbox" checked={soSem} onChange={e=>setSoSem(e.target.checked)}/> Só não classificados</label>
        </div>
      </Card>
      <Card style={{padding:0,overflow:"hidden"}}>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse"}}>
            <thead><tr>{["Cliente","Empresa","Receitas","Valor total","BU"].map(h=><th key={h} style={th}>{h}</th>)}</tr></thead>
            <tbody>
              {shown.length===0 && <tr><td colSpan={5} style={{padding:"22px",textAlign:"center",color:T.muted,fontSize:13}}>Nenhum cliente.</td></tr>}
              {shown.map(g=>(
                <tr key={g.nome+"|"+g.ids.length}>
                  <td style={{...td,maxWidth:300,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={g.nome}>{g.nome}</td>
                  <td style={{...td,whiteSpace:"nowrap",color:T.inkSoft}}>{g.empLabel||"—"}</td>
                  <td style={td}>{g.ids.length}</td>
                  <td style={{...td,whiteSpace:"nowrap",fontWeight:600}}>{brl(g.total)}</td>
                  <td style={{...td,whiteSpace:"nowrap"}}>
                    <select value={g.buAtual==="__MISTO__"?"":g.buAtual} onChange={e=>onSetBu(g.ids, e.target.value)}
                      style={{...inp,width:"auto",fontSize:12,padding:"5px 8px",...(g.buAtual&&g.buAtual!=="__MISTO__"?{borderColor:T.brand,color:T.brand,fontWeight:700}:{})}}>
                      <option value="">{g.buAtual==="__MISTO__"?"— misto (redefinir) —":"— sem BU —"}</option>
                      {BUS.map(b=><option key={b}>{b}</option>)}
                    </select>
                    {g.buAtual==="__MISTO__" && <span style={{marginLeft:8,fontSize:11,color:T.warn}}>múltiplas BUs · ajuste fino em Correções</span>}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>
      {lista.length>400 && <div style={{fontSize:11.5,color:T.faint,marginTop:8}}>Mostrando 400 de {lista.length} — use a busca.</div>}
    </div>
  );
}

// ─── VISÃO COMERCIAL (por BU) ────────────────────────────────────────────────
// O que um diretor comercial vê da sua unidade de negócio: reconhecido, faturado,
// a faturar e represado por cliente. Só leitura. (Acesso por diretor/BU vem depois.)
// ─── UNIFICAR CLIENTES (DE → PARA) ───────────────────────────────────────────
// Mapa manual para quando o MESMO cliente aparece com nomes totalmente diferentes
// nos registros. O nome DE passa a ser lido como o nome PARA em todas as telas.
function AliasesView({ aliases=[], records=[], onSave, onDelete, isViewer=false }) {
  const [de, setDe] = useState("");
  const [para, setPara] = useState("");
  // Nomes distintos que aparecem hoje nos registros (com contagem) — para escolher.
  const cont = {};
  records.forEach(r=>{ const n=(r.cliente||"").trim(); if(n) cont[n]=(cont[n]||0)+1; });
  const nomes = Object.entries(cont).sort((a,b)=>a[0].localeCompare(b[0]));
  const paraNomes = [...new Set([...nomes.map(n=>n[0]), ...aliases.map(a=>a.para)])].sort((a,b)=>a.localeCompare(b));

  function submit(){ onSave({ de, para }); setDe(""); setPara(""); }

  return (
    <div>
      <PageHead icon="building" title="Unificar clientes (DE → PARA)" sub="Quando o mesmo cliente aparece com nomes diferentes, aponte o nome de origem para o nome final. Vale em todas as telas e sobrevive a re-importações."/>
      <div style={{fontSize:12.5,color:T.muted,marginBottom:16}}>Ex.: <b>alocacao mandic</b> → <b>SOCIEDADE REGIONAL DE ENSINO E SAUDE LTDA</b>. O app não altera o dado bruto — só a leitura. Não remove receitas duplicadas: se a mesma receita foi importada duas vezes, use <b>Validações → possivelmente duplicadas</b> e <b>Correções</b>.</div>

      {!isViewer && <Card style={{padding:16,marginBottom:16}}>
        <div style={{display:"grid",gridTemplateColumns:"1fr auto 1fr auto",gap:12,alignItems:"end"}}>
          <Field label="DE (nome de origem)">
            <input style={inp} list="alias-de" placeholder="Nome como aparece hoje" value={de} onChange={e=>setDe(e.target.value)}/>
            <datalist id="alias-de">{nomes.map(([n,c])=><option key={n} value={n}>{n} ({c})</option>)}</datalist>
          </Field>
          <div style={{fontSize:20,color:T.brand,fontWeight:800,paddingBottom:8}}>→</div>
          <Field label="PARA (nome final / canônico)">
            <input style={inp} list="alias-para" placeholder="Nome que deve valer" value={para} onChange={e=>setPara(e.target.value)}/>
            <datalist id="alias-para">{paraNomes.map(n=><option key={n} value={n}/>)}</datalist>
          </Field>
          <Btn primary onClick={submit} disabled={!de.trim()||!para.trim()}>Unificar</Btn>
        </div>
      </Card>}

      <SectionTitle count={aliases.length}>Unificações ativas</SectionTitle>
      {aliases.length===0
        ? <Card style={{padding:24,textAlign:"center",color:T.muted,fontSize:13}}>Nenhuma unificação cadastrada ainda.</Card>
        : <Card style={{padding:0,overflow:"hidden"}}>
            <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
              <thead><tr>
                <th style={{padding:"9px 14px",textAlign:"left",fontSize:11,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:".3px",borderBottom:`1px solid ${T.line}`}}>DE (origem)</th>
                <th style={{padding:"9px 14px",textAlign:"left",fontSize:11,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:".3px",borderBottom:`1px solid ${T.line}`}}>PARA (final)</th>
                <th style={{borderBottom:`1px solid ${T.line}`}}/>
              </tr></thead>
              <tbody>
                {aliases.slice().sort((a,b)=>a.para.localeCompare(b.para)||a.de.localeCompare(b.de)).map(a=>(
                  <tr key={a.de} style={{borderBottom:`1px solid ${T.lineSoft}`}}>
                    <td style={{padding:"9px 14px",color:T.inkSoft}}>{a.de}</td>
                    <td style={{padding:"9px 14px",fontWeight:600,color:T.ink}}><span style={{color:T.brand,marginRight:6,fontWeight:800}}>→</span>{a.para}</td>
                    <td style={{padding:"9px 14px",textAlign:"right"}}>{!isViewer && <Btn small danger onClick={()=>onDelete(a.de)}>Remover</Btn>}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </Card>}
    </div>
  );
}

function ComercialView({ records, clients=[], fatByRec={}, varByRec={} }) {
  const key = s => (s||"").toString().toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g,"").replace(/[^a-z0-9]/g,"");
  const bus = [...new Set(records.map(r=>r.bu).filter(Boolean))].sort();
  const [buSel, setBuSel] = useState("");
  const [emp, setEmp] = useState("todas");
  const [comp, setComp] = useState("todas");
  const bu = buSel || bus[0] || "";
  const bill = r => (r.valorTotal||0)+(varByRec[r.id]||0);
  const fat  = r => fatByRec[r.id]||0;
  const daBu = records.filter(r=>r.bu===bu);
  const empresasComDados = [...new Set(daBu.map(r=>r.empresa).filter(Boolean))].sort();
  const comps = [...new Set(daBu.map(r=>r.competencia).filter(Boolean))].sort((a,b)=>{ const [ma,ya]=String(a).split("/"),[mb,yb]=String(b).split("/"); return (Number(yb)-Number(ya))||(Number(mb)-Number(ma)); });

  let recs = daBu;
  if (emp!=="todas")  recs = recs.filter(r=>r.empresa===emp);
  if (comp!=="todas") recs = recs.filter(r=>r.competencia===comp);

  const grupos = {};
  recs.forEach(r => {
    const k = key(r.cliente); if(!k) return;
    const g = (grupos[k] = grupos[k] || { nome:r.cliente, emps:new Set(), reconhecido:0, faturado:0, represado:0, ciclo:0 });
    const b=bill(r), f=fat(r), s=b-f;
    g.reconhecido+=b; g.faturado+=f; g.emps.add(r.empresa);
    if (s>0.01) { if (categoriaOf(r,clients).cat==="represado") g.represado+=s; else g.ciclo+=s; }
  });
  const lista = Object.values(grupos).map(g=>({ ...g, aberto:g.represado+g.ciclo, empLabel:[...g.emps].filter(Boolean).sort().join(", ") }))
    .sort((a,b)=>b.reconhecido-a.reconhecido);
  const tot = lista.reduce((t,g)=>({ rec:t.rec+g.reconhecido, fat:t.fat+g.faturado, rep:t.rep+g.represado, cic:t.cic+g.ciclo }),{rec:0,fat:0,rep:0,cic:0});
  const pctFat = tot.rec>0.01 ? Math.round(tot.fat/tot.rec*100) : 0;

  const th={padding:"7px 10px",textAlign:"left",fontSize:11,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:".3px",borderBottom:`1px solid ${T.line}`,whiteSpace:"nowrap"};
  const thR={...th,textAlign:"right"};
  const td={padding:"7px 10px",fontSize:12.5,borderBottom:`1px solid ${T.line}`};
  const tdR={...td,textAlign:"right",whiteSpace:"nowrap",fontVariantNumeric:"tabular-nums"};
  const kpi=(label,valor,cor,sub)=>(
    <Card style={{padding:16,flex:"1 1 180px"}}>
      <div style={{fontSize:12,color:T.muted,fontWeight:600}}>{label}</div>
      <div style={{fontSize:22,fontWeight:800,color:cor,marginTop:4,fontVariantNumeric:"tabular-nums"}}>{brl(valor)}</div>
      {sub && <div style={{fontSize:11.5,color:T.muted,marginTop:2}}>{sub}</div>}
    </Card>
  );

  return (
    <div>
      <PageHead icon="chart" title="Visão comercial" sub="Reconhecido, faturado e represado por cliente — na ótica da unidade de negócio."/>
      {bus.length===0
        ? <Card style={{padding:22,textAlign:"center",color:T.muted}}>Nenhuma BU classificada ainda. Classifique os clientes em <b>Administração → Classificar BU</b> e esta visão se preenche.</Card>
        : <>
          <Card style={{padding:14,marginBottom:14}}>
            <div style={{display:"flex",gap:10,flexWrap:"wrap",alignItems:"flex-end"}}>
              <Field label="Unidade de negócio (BU)"><select style={{...inp,width:"auto",minWidth:180,fontWeight:700,color:T.brand,borderColor:T.brand}} value={bu} onChange={e=>{setBuSel(e.target.value);setEmp("todas");setComp("todas");}}>{bus.map(b=><option key={b}>{b}</option>)}</select></Field>
              <Field label="Empresa"><select style={{...inp,width:"auto"}} value={emp} onChange={e=>setEmp(e.target.value)}><option value="todas">Todas</option>{empresasComDados.map(c=>{const e=EMPRESAS.find(x=>x.cod===c);return <option key={c} value={c}>{c}{e?` — ${e.nome}`:""}</option>;})}</select></Field>
              <Field label="Competência"><select style={{...inp,width:"auto"}} value={comp} onChange={e=>setComp(e.target.value)}><option value="todas">Todas</option>{comps.map(c=><option key={c}>{c}</option>)}</select></Field>
            </div>
          </Card>
          <div style={{display:"flex",gap:12,flexWrap:"wrap",marginBottom:14}}>
            {kpi("Reconhecido", tot.rec, T.ink, `${lista.length} cliente(s)`)}
            {kpi("Faturado", tot.fat, C.green.solid, `${pctFat}% do reconhecido`)}
            {kpi("A faturar", tot.rec-tot.fat, C.orange.solid, "reconhecido ainda em aberto")}
            {kpi("Represado", tot.rep, C.red.solid, "em aberto e fora do ciclo")}
          </div>
          <Card style={{padding:0,overflow:"hidden"}}>
            <div style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse"}}>
                <thead><tr><th style={th}>Cliente</th><th style={th}>Empresa</th><th style={thR}>Reconhecido</th><th style={thR}>Faturado</th><th style={thR}>A faturar</th><th style={thR}>Represado</th><th style={thR}>% fat.</th></tr></thead>
                <tbody>
                  {lista.length===0 && <tr><td colSpan={7} style={{padding:"22px",textAlign:"center",color:T.muted,fontSize:13}}>Sem receitas nesse recorte.</td></tr>}
                  {lista.map(g=>{ const pf=g.reconhecido>0.01?Math.round(g.faturado/g.reconhecido*100):0; return (
                    <tr key={g.nome+g.empLabel}>
                      <td style={{...td,maxWidth:280,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={g.nome}>{g.nome}</td>
                      <td style={{...td,color:T.inkSoft,whiteSpace:"nowrap"}}>{g.empLabel||"—"}</td>
                      <td style={{...tdR,fontWeight:600}}>{brl(g.reconhecido)}</td>
                      <td style={{...tdR,color:C.green.solid}}>{brl(g.faturado)}</td>
                      <td style={tdR}>{brl(g.aberto)}</td>
                      <td style={{...tdR,color:g.represado>0.01?C.red.solid:T.faint,fontWeight:g.represado>0.01?700:400}}>{brl(g.represado)}</td>
                      <td style={tdR}>{pf}%</td>
                    </tr>
                  );})}
                </tbody>
              </table>
            </div>
          </Card>
        </>}
    </div>
  );
}

// ─── REPORT SEMANAL DO COMERCIAL ─────────────────────────────────────────────
// A mesma foto que o "disparo" (Edge Function) manda por e-mail — aqui dentro do
// app, para o comercial ver, baixar em Excel e enviar na hora. Detalha por tipo
// de projeto; em Time & Expenses, mostra a linha de cada consultor (valor/hora,
// horas e total).
function WeeklyReportView({ records, clients=[], bu:buFixed, nome, fatByRec={}, varByRec={}, canPickBu=false }) {
  const toast = useToast();
  const [sending, setSending] = useState(false);
  const bill = r => (r.valorTotal||0)+(varByRec[r.id]||0);
  const fat  = r => fatByRec[r.id]||0;
  const rep  = r => bill(r)-fat(r);
  const empNome = cod => EMPRESAS.find(e=>e.cod===cod)?.nome||"";

  // Comercial: BU fixa (records já vem filtrado). Admin: escolhe a BU aqui.
  const busAll = [...new Set(records.map(r=>r.bu).filter(Boolean))].sort();
  const [buSel, setBuSel] = useState(buFixed || busAll[0] || "");
  const bu = canPickBu ? buSel : (buFixed||"");
  const scoped = canPickBu ? records.filter(r=>r.bu===bu) : records;

  const comps = [...new Set(scoped.map(r=>r.competencia).filter(Boolean))].sort((a,b)=>{ const [ma,ya]=String(a).split("/"),[mb,yb]=String(b).split("/"); return (Number(yb)-Number(ya))||(Number(mb)-Number(ma)); });
  const [comp, setComp] = useState("todas");
  const recs = comp==="todas" ? scoped : scoped.filter(r=>r.competencia===comp);

  const tot = recs.reduce((a,r)=>({ rec:a.rec+bill(r), fat:a.fat+fat(r), rep:a.rep+Math.max(0,rep(r)) }),{rec:0,fat:0,rep:0});
  const pctFat = tot.rec>0.01 ? Math.round(tot.fat/tot.rec*100) : 0;

  // Agrupa por tipo de projeto.
  const porTipo = {};
  recs.forEach(r=>{ (porTipo[r.tipo||"—"]=porTipo[r.tipo||"—"]||[]).push(r); });
  const isTE = t => /time|expense|t&e/i.test(t);

  const th={padding:"7px 10px",textAlign:"left",fontSize:11,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:".3px",borderBottom:`1px solid ${T.line}`,whiteSpace:"nowrap"};
  const thR={...th,textAlign:"right"};
  const td={padding:"7px 10px",fontSize:12.5,borderBottom:`1px solid ${T.lineSoft}`};
  const tdR={...td,textAlign:"right",whiteSpace:"nowrap",fontVariantNumeric:"tabular-nums"};
  const kpi=(label,valor,cor,sub)=>(
    <Card style={{padding:16,flex:"1 1 170px"}}>
      <div style={{fontSize:12,color:T.muted,fontWeight:600}}>{label}</div>
      <div style={{fontSize:22,fontWeight:800,color:cor,marginTop:4,fontVariantNumeric:"tabular-nums"}}>{brl(valor)}</div>
      {sub && <div style={{fontSize:11.5,color:T.muted,marginTop:2}}>{sub}</div>}
    </Card>
  );

  function exportXLSX() {
    const headers = ["Tipo","Empresa","Cliente","Consultor / PEP","Valor/hora","Horas","Reconhecido","Faturado","Represado"];
    const rows = [];
    Object.entries(porTipo).forEach(([tipo,list])=>{
      list.slice().sort((a,b)=>bill(b)-bill(a)).forEach(r=>{
        rows.push([ tipo, r.empresa||"", r.cliente||"", isTE(tipo)?(r.profissional||"—"):(r.pep||"—"),
          isTE(tipo)?Number((r.valorVenda||0).toFixed(2)):"", isTE(tipo)?(r.hrsAprovadas||0):"",
          Number(bill(r).toFixed(2)), Number(fat(r).toFixed(2)), Number(Math.max(0,rep(r)).toFixed(2)) ]);
      });
    });
    const stamp = new Date().toISOString().slice(0,10);
    downloadXLSX(`Report_${(bu||"BU").replace(/\s+/g,"_")}_${stamp}.xlsx`, headers, rows);
  }

  async function enviarAgora() {
    setSending(true);
    try {
      const { data, error } = await supabase.functions.invoke("weekly-commercial-report", { body:{ bu } });
      if (error) throw error;
      const sent = data?.sent ?? 0;
      toast(sent>0 ? `Report enviado por e-mail (${sent}).` : "Disparo executado — verifique a configuração de e-mail (RESEND_API_KEY).", sent>0?"success":"info");
    } catch(e) {
      toast("Envio automático ainda não ativado neste ambiente. Baixe o Excel para encaminhar, ou peça ao admin para publicar a função de e-mail.", "info");
    } finally { setSending(false); }
  }

  return (
    <div>
      <PageHead icon="file" title="Report semanal" sub={`${bu||"Sua BU"} — reconhecido, faturado e represado, detalhado por tipo de projeto.`}
        right={<div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          <Btn icon="download" disabled={!recs.length} onClick={exportXLSX}>Baixar Excel</Btn>
          <Btn primary icon="check" disabled={!recs.length||sending} onClick={enviarAgora}>{sending?"Enviando…":"Enviar por e-mail"}</Btn>
        </div>}/>

      <Card style={{padding:14,marginBottom:14}}>
        <div style={{display:"flex",gap:10,flexWrap:"wrap",alignItems:"flex-end"}}>
          {canPickBu && <Field label="Unidade de negócio (BU)"><select style={{...inp,width:"auto",minWidth:180,fontWeight:700,color:T.brand,borderColor:T.brand}} value={buSel} onChange={e=>{setBuSel(e.target.value);setComp("todas");}}>{busAll.length===0 && <option value="">— sem BU classificada —</option>}{busAll.map(b=><option key={b}>{b}</option>)}</select></Field>}
          <Field label="Competência"><select style={{...inp,width:"auto",minWidth:150}} value={comp} onChange={e=>setComp(e.target.value)}><option value="todas">Todas</option>{comps.map(c=><option key={c}>{c}</option>)}</select></Field>
          <div style={{fontSize:12,color:T.muted,flex:1,minWidth:200}}>Esta é exatamente a foto que o disparo automático semanal envia por e-mail para o comercial da BU{canPickBu?" selecionada":""}.</div>
        </div>
      </Card>

      <div style={{display:"flex",gap:12,flexWrap:"wrap",marginBottom:16}}>
        {kpi("Reconhecido", tot.rec, T.ink, `${recs.length} lançamento(s)`)}
        {kpi("Faturado", tot.fat, C.green.solid, `${pctFat}% do reconhecido`)}
        {kpi("A faturar", tot.rec-tot.fat, C.orange.solid, "reconhecido em aberto")}
        {kpi("Represado", tot.rep, C.red.solid, "em aberto e fora do ciclo")}
      </div>

      {recs.length===0
        ? <Card style={{padding:28,textAlign:"center",color:T.muted}}>Sem receitas classificadas na sua BU para este recorte.</Card>
        : Object.entries(porTipo).sort((a,b)=>b[1].reduce((s,r)=>s+bill(r),0)-a[1].reduce((s,r)=>s+bill(r),0)).map(([tipo,list])=>{
          const subtot = list.reduce((a,r)=>({rec:a.rec+bill(r),fat:a.fat+fat(r),rep:a.rep+Math.max(0,rep(r))}),{rec:0,fat:0,rep:0});
          if (isTE(tipo)) {
            return (
              <Card key={tipo} style={{padding:0,overflow:"hidden",marginBottom:14}}>
                <div style={{display:"flex",alignItems:"center",gap:10,padding:"11px 16px",background:T.canvas,borderBottom:`1px solid ${T.lineSoft}`,flexWrap:"wrap"}}>
                  <Badge label={tipo} color="blue" small/><span style={{fontSize:11.5,color:T.muted}}>por consultor</span><div style={{flex:1}}/>
                  <span style={{fontSize:12.5,fontWeight:700}}>{brl(subtot.rec)}</span>
                </div>
                <div style={{overflowX:"auto"}}>
                  <table style={{width:"100%",borderCollapse:"collapse"}}>
                    <thead><tr><th style={th}>Cliente</th><th style={th}>Consultor</th><th style={th}>PEP</th><th style={thR}>Valor/hora</th><th style={thR}>Horas</th><th style={thR}>Total</th><th style={thR}>Faturado</th><th style={thR}>Represado</th></tr></thead>
                    <tbody>
                      {list.slice().sort((a,b)=>bill(b)-bill(a)).map(r=>(
                        <tr key={r.id}>
                          <td style={{...td,maxWidth:200,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={r.cliente}>{r.cliente}</td>
                          <td style={{...td,whiteSpace:"nowrap"}}>{r.profissional||"—"}</td>
                          <td style={{...td,color:T.muted,fontSize:11,whiteSpace:"nowrap"}}>{r.pep||"—"}</td>
                          <td style={tdR}>{brl(r.valorVenda||0)}</td>
                          <td style={tdR}>{(r.hrsAprovadas||0).toLocaleString("pt-BR")}</td>
                          <td style={{...tdR,fontWeight:600}}>{brl(bill(r))}</td>
                          <td style={{...tdR,color:C.green.solid}}>{brl(fat(r))}</td>
                          <td style={{...tdR,color:Math.max(0,rep(r))>0.01?C.red.solid:T.faint}}>{brl(Math.max(0,rep(r)))}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </Card>
            );
          }
          // Demais tipos: consolidado por cliente.
          const byCli={};
          list.forEach(r=>{ const k=r.cliente||"—"; const g=(byCli[k]=byCli[k]||{rec:0,fat:0,rep:0,emp:new Set()}); g.rec+=bill(r); g.fat+=fat(r); g.rep+=Math.max(0,rep(r)); g.emp.add(r.empresa); });
          return (
            <Card key={tipo} style={{padding:0,overflow:"hidden",marginBottom:14}}>
              <div style={{display:"flex",alignItems:"center",gap:10,padding:"11px 16px",background:T.canvas,borderBottom:`1px solid ${T.lineSoft}`,flexWrap:"wrap"}}>
                <Badge label={tipo} color="gray" small/><div style={{flex:1}}/><span style={{fontSize:12.5,fontWeight:700}}>{brl(subtot.rec)}</span>
              </div>
              <div style={{overflowX:"auto"}}>
                <table style={{width:"100%",borderCollapse:"collapse"}}>
                  <thead><tr><th style={th}>Cliente</th><th style={th}>Empresa</th><th style={thR}>Reconhecido</th><th style={thR}>Faturado</th><th style={thR}>Represado</th></tr></thead>
                  <tbody>
                    {Object.entries(byCli).sort((a,b)=>b[1].rec-a[1].rec).map(([cli,g])=>(
                      <tr key={cli}>
                        <td style={{...td,maxWidth:280,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={cli}>{cli}</td>
                        <td style={{...td,color:T.inkSoft,whiteSpace:"nowrap"}}>{[...g.emp].filter(Boolean).join(", ")||"—"}</td>
                        <td style={{...tdR,fontWeight:600}}>{brl(g.rec)}</td>
                        <td style={{...tdR,color:C.green.solid}}>{brl(g.fat)}</td>
                        <td style={{...tdR,color:g.rep>0.01?C.red.solid:T.faint}}>{brl(g.rep)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </Card>
          );
        })}
    </div>
  );
}

// ─── PREVISÃO & SAÚDE DA RECEITA ─────────────────────────────────────────────
// Esperado por projeto (PEP): T&E/Usage = média dos últimos 6 meses; Fee = valor
// recorrente; WIP = média (sinalizado — o certo é a proposta). Somado = previsão
// do próximo mês. Comparado com o realizado = saúde do projeto (subiu/caiu).
function ForecastView({ records, varByRec={} }) {
  const key = s => (s||"").toString().toLowerCase().normalize("NFD").replace(/[̀-ͯ]/g,"").replace(/[^a-z0-9]/g,"");
  const rank = c => { const [m,y]=String(c||"").split("/").map(Number); return (y&&m)?y*12+m:0; };
  const bill = r => (r.valorTotal||0)+(varByRec[r.id]||0);
  const bus = [...new Set(records.map(r=>r.bu).filter(Boolean))].sort();
  const empresasComDados = [...new Set(records.map(r=>r.empresa).filter(Boolean))].sort();
  const [buF,setBuF]=useState("todas");
  const [empF,setEmpF]=useState("todas");
  const [tipoF,setTipoF]=useState("todos");
  const [q,setQ]=useState("");

  let recs = records;
  if (buF!=="todas")  recs = recs.filter(r=>r.bu===buF);
  if (empF!=="todas") recs = recs.filter(r=>r.empresa===empF);
  if (tipoF!=="todos") recs = recs.filter(r=>r.tipo===tipoF);

  const proj = {};
  recs.forEach(r => {
    const pk = key(pepBase(r.pep)); if(!pk) return;
    const gk = (r.empresa||"")+"|"+pk;
    const g = (proj[gk] = proj[gk] || { pep:pepBase(r.pep), cliente:r.cliente, tipo:r.tipo, bu:r.bu, empresa:r.empresa, byComp:{} });
    g.byComp[r.competencia] = (g.byComp[r.competencia]||0) + bill(r);
  });
  let lista = Object.values(proj).map(g => {
    const comps = Object.keys(g.byComp).sort((a,b)=>rank(a)-rank(b));
    const vals = comps.map(c=>g.byComp[c]);
    const ult6 = vals.slice(-6);
    const media6 = ult6.length ? ult6.reduce((a,b)=>a+b,0)/ult6.length : 0;
    const ultimo = vals.length ? vals[vals.length-1] : 0;
    const ant = vals.slice(-6,-1);
    const mediaAnt = ant.length ? ant.reduce((a,b)=>a+b,0)/ant.length : null;
    const esperado = g.tipo==="Fee" ? ultimo : media6;
    let saude="novo", varPct=null;
    if (mediaAnt!=null && mediaAnt>0.01) { const r=ultimo/mediaAnt; varPct=(ultimo-mediaAnt)/mediaAnt; saude = r>=0.85?"ok":(r>=0.5?"queda":"critico"); }
    else if (ultimo>0.01) saude="ok";
    return { ...g, meses:comps.length, media6, ultimo, esperado, saude, varPct };
  });
  if (q.trim()) lista = lista.filter(g=>key(g.cliente).includes(key(q))||key(g.pep).includes(key(q)));
  lista.sort((a,b)=>b.media6-a.media6);
  const shown = lista.slice(0,400);
  const previsao = lista.reduce((s,g)=>s+g.esperado,0);
  const porTipo = {}; lista.forEach(g=>{ porTipo[g.tipo]=(porTipo[g.tipo]||0)+g.esperado; });
  const emQueda = lista.filter(g=>g.saude==="queda"||g.saude==="critico");

  const SAUDE = { ok:{l:"saudável",c:"green"}, queda:{l:"caiu",c:"yellow"}, critico:{l:"crítico",c:"red"}, novo:{l:"novo",c:"gray"} };
  const th={padding:"7px 10px",textAlign:"left",fontSize:11,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:".3px",borderBottom:`1px solid ${T.line}`,whiteSpace:"nowrap"};
  const thR={...th,textAlign:"right"};
  const td={padding:"7px 10px",fontSize:12.5,borderBottom:`1px solid ${T.line}`};
  const tdR={...td,textAlign:"right",whiteSpace:"nowrap",fontVariantNumeric:"tabular-nums"};

  return (
    <div>
      <PageHead icon="chart" title="Previsão & Saúde da receita" sub="Receita esperada por projeto (média 6m · Fee recorrente) e se o projeto subiu ou caiu."/>
      <Card style={{padding:14,marginBottom:14}}>
        <div style={{display:"flex",gap:10,flexWrap:"wrap",alignItems:"flex-end"}}>
          <Field label="Buscar"><input style={{...inp,minWidth:200}} placeholder="cliente ou PEP" value={q} onChange={e=>setQ(e.target.value)}/></Field>
          {bus.length>0 && <Field label="BU"><select style={{...inp,width:"auto"}} value={buF} onChange={e=>setBuF(e.target.value)}><option value="todas">Todas</option>{bus.map(b=><option key={b}>{b}</option>)}</select></Field>}
          <Field label="Empresa"><select style={{...inp,width:"auto"}} value={empF} onChange={e=>setEmpF(e.target.value)}><option value="todas">Todas</option>{empresasComDados.map(c=><option key={c}>{c}</option>)}</select></Field>
          <Field label="Tipo"><select style={{...inp,width:"auto"}} value={tipoF} onChange={e=>setTipoF(e.target.value)}><option value="todos">Todos</option>{TIPOS_PROJETO.map(t=><option key={t}>{t}</option>)}</select></Field>
        </div>
      </Card>
      <div style={{display:"flex",gap:12,flexWrap:"wrap",marginBottom:14}}>
        <Card style={{padding:16,flex:"2 1 300px"}}>
          <div style={{fontSize:12,color:T.muted,fontWeight:600}}>Previsão do próximo mês</div>
          <div style={{fontSize:24,fontWeight:800,color:T.brand,marginTop:4,fontVariantNumeric:"tabular-nums"}}>{brl(previsao)}</div>
          <div style={{display:"flex",gap:6,flexWrap:"wrap",marginTop:8}}>
            {TIPOS_PROJETO.filter(t=>porTipo[t]>0.01).map(t=><Badge key={t} label={`${t}: ${fmtShort(porTipo[t])}`} color={({["Time & Expenses"]:"blue",Fee:"purple",WIP:"teal","Usage Based":"orange"})[t]||"gray"} small/>)}
          </div>
        </Card>
        <Card style={{padding:16,flex:"1 1 180px"}}>
          <div style={{fontSize:12,color:T.muted,fontWeight:600}}>Projetos em atenção</div>
          <div style={{fontSize:24,fontWeight:800,color:emQueda.length?C.red.solid:C.green.solid,marginTop:4}}>{emQueda.length}</div>
          <div style={{fontSize:11.5,color:T.muted,marginTop:2}}>caíram vs. a própria média</div>
        </Card>
      </div>
      <Card style={{padding:0,overflow:"hidden"}}>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse"}}>
            <thead><tr><th style={th}>Cliente</th><th style={th}>PEP</th><th style={th}>Tipo</th><th style={thR}>Meses</th><th style={thR}>Média 6m</th><th style={thR}>Último mês</th><th style={th}>Saúde</th><th style={thR}>Previsão</th></tr></thead>
            <tbody>
              {shown.length===0 && <tr><td colSpan={8} style={{padding:"22px",textAlign:"center",color:T.muted,fontSize:13}}>Sem projetos nesse recorte.</td></tr>}
              {shown.map(g=>{ const s=SAUDE[g.saude]; return (
                <tr key={g.empresa+g.pep+g.cliente}>
                  <td style={{...td,maxWidth:220,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}} title={g.cliente}>{g.cliente}</td>
                  <td style={{...td,whiteSpace:"nowrap",fontWeight:600}}>{g.pep}</td>
                  <td style={{...td,whiteSpace:"nowrap",color:T.inkSoft}}>{g.tipo}{g.tipo==="WIP"?" ⚠":""}</td>
                  <td style={tdR}>{g.meses}</td>
                  <td style={{...tdR,fontWeight:600}}>{brl(g.media6)}</td>
                  <td style={tdR}>{brl(g.ultimo)}{g.varPct!=null&&<span style={{marginLeft:6,fontSize:11,fontWeight:700,color:g.varPct>=-0.05?C.green.solid:(g.varPct>=-0.4?C.yellow.solid:C.red.solid)}}>{g.varPct>=0?"▲":"▼"}{Math.abs(Math.round(g.varPct*100))}%</span>}</td>
                  <td style={td}><Badge label={s.l} color={s.c} small dot/></td>
                  <td style={{...tdR,fontWeight:700,color:T.brand}}>{brl(g.esperado)}</td>
                </tr>
              );})}
            </tbody>
          </table>
        </div>
      </Card>
      <div style={{fontSize:11.5,color:T.faint,marginTop:10,lineHeight:1.5}}>
        Previsão: <b>T&E e Usage Based</b> = média dos últimos 6 meses · <b>Fee</b> = valor recorrente · <b>WIP</b> (⚠) = média como aproximação — o correto é a proposta/etapa. Saúde compara o último mês com a média dos meses anteriores.
      </div>
    </div>
  );
}

// ─── APP ROOT ─────────────────────────────────────────────────────────────────

function AppInner() {
  const toast = useToast();
  const isMobile = useIsMobile();
  // Tema claro/escuro — o data-theme no <html> troca as CSS variables (aplicado
  // em main.jsx antes do render). O botão só alterna o atributo e salva a escolha.
  const [theme, setTheme] = useState(() => document.documentElement.getAttribute("data-theme") === "dark" ? "dark" : "light");
  const toggleTheme = () => setTheme(t => {
    const next = t==="dark" ? "light" : "dark";
    document.documentElement.setAttribute("data-theme", next);
    try { localStorage.setItem("fc-theme", next); } catch {}
    return next;
  });
  const [state, setState]       = useState(()=>loadState());
  const [user, setUser]         = useState(null);
  const [authReady, setAuthRdy] = useState(false);
  const [recovery, setRecovery] = useState(false);
  const [page, setPage]         = useState("home");
  const [showImport, setImp]    = useState(false);
  const [showExport, setExp]    = useState(false);
  const [showHistory, setHist]  = useState(false);
  const [confirmLogout, setCL]  = useState(false);
  const [drawer, setDrawer]     = useState(false);
  const [records, setRecords]   = useState([]);
  const [tasks, setTasks]       = useState([]);
  const [history, setHistory]   = useState([]);
  const [profiles, setProfiles] = useState([]);
  const [clients, setClients]   = useState([]);
  const [templates, setTemplates] = useState([]);
  const [deliveries, setDeliveries] = useState([]);
  const [notes, setNotes] = useState([]);   // notas da prefeitura (NFS-e)
  const [faturamentos, setFaturamentos] = useState([]);  // livro de faturamento (alocações parciais)
  const [variacoes, setVariacoes] = useState([]);        // variações de receita pós-fechamento
  const [mural, setMural] = useState({ id:null, frase:"", autor:"", lembretes:[] });
  const [aliases, setAliases] = useState([]);        // DE→PARA de clientes
  const aliasMapRef = useRef({});                    // { nomeNormalizado(de) -> para }
  const buildAliasMap = (list) => { const m={}; (list||[]).forEach(a=>{ if(a.de&&a.para) m[_normCliNome(a.de)]=a.para; }); return m; };
  const [dataReady, setDataRdy] = useState(false);

  useEffect(()=>saveState(state),[state]);

  // ─ Carrega os dados do banco quando o usuário entra ─
  const reloadRecords = useCallback(async () => { try { setRecords(applyClientAliases(await db.fetchRecords(), aliasMapRef.current)); } catch(e){ toast("Erro ao carregar registros: "+e.message, "error"); } }, [toast]);
  const reloadAliases = useCallback(async () => { try { const al = await db.fetchClientAliases(); aliasMapRef.current = buildAliasMap(al); setAliases(al); } catch(e){ /* tabela pode não existir ainda */ } }, []);
  const reloadTasks   = useCallback(async () => { try { setTasks(await db.fetchTasks()); }     catch(e){ toast("Erro ao carregar tarefas: "+e.message, "error"); } }, [toast]);
  const reloadHistory  = useCallback(async () => { try { setHistory(await db.fetchHistory()); } catch(e){ /* histórico é só p/ admin */ } }, []);
  const reloadProfiles = useCallback(async () => { try { setProfiles(await db.fetchProfiles()); } catch(e){ toast("Erro ao carregar acessos: "+e.message, "error"); } }, [toast]);
  const reloadClients  = useCallback(async () => { try { setClients(await db.fetchClients()); } catch(e){ toast("Erro ao carregar clientes: "+e.message, "error"); } }, [toast]);
  const reloadTemplates = useCallback(async () => { try { setTemplates(await db.fetchTemplates()); } catch(e){ /* entregas: só admin */ } }, []);
  const reloadDeliveries = useCallback(async () => { try { setDeliveries(await db.fetchDeliveries()); } catch(e){ /* idem */ } }, []);
  const reloadNotes = useCallback(async () => { try { setNotes(await db.fetchMunicipalNotes()); } catch(e){ /* notas: tabela pode não existir ainda */ } }, []);
  const reloadFaturamentos = useCallback(async () => { try { setFaturamentos(await db.fetchFaturamentos()); } catch(e){ /* tabela pode não existir ainda */ } }, []);
  const reloadVariacoes = useCallback(async () => { try { setVariacoes(await db.fetchVariacoes()); } catch(e){ /* tabela pode não existir ainda */ } }, []);
  const reloadMural = useCallback(async () => { try { setMural(await db.fetchMural()); } catch(e){ /* mural: tabela pode não existir ainda */ } }, []);

  useEffect(() => {
    if (!user) { setRecords([]); setTasks([]); setHistory([]); setProfiles([]); setClients([]); setTemplates([]); setDeliveries([]); setNotes([]); setFaturamentos([]); setVariacoes([]); setMural({ id:null, frase:"", autor:"", lembretes:[] }); return; }
    let active = true;
    // NÃO voltamos para a tela de "Carregando" em recargas — isso desmontaria
    // formulários/modais abertos. A tela de carregamento só aparece na 1ª vez.
    Promise.all([db.fetchRecords(), db.fetchTasks(), db.fetchHistory().catch(()=>[]), db.fetchProfiles().catch(()=>[]), db.fetchClients().catch(()=>[]), db.fetchTemplates().catch(()=>[]), db.fetchDeliveries().catch(()=>[]), db.fetchMunicipalNotes().catch(()=>[]), db.fetchMural().catch(()=>({ id:null, frase:"", autor:"", lembretes:[] })), db.fetchFaturamentos().catch(()=>[]), db.fetchVariacoes().catch(()=>[]), db.fetchClientAliases().catch(()=>[])])
      .then(([r, t, h, p, c, tm, dv, nt, mu, fa, vr, al]) => { if (!active) return; aliasMapRef.current = buildAliasMap(al); setAliases(al); setRecords(applyClientAliases(r, aliasMapRef.current)); setTasks(t); setHistory(h); setProfiles(p); setClients(c); setTemplates(tm); setDeliveries(dv); setNotes(nt); setMural(mu); setFaturamentos(fa); setVariacoes(vr); })
      .catch(e => { if (active) toast("Erro ao carregar dados: "+e.message, "error"); })
      .finally(() => { if (active) setDataRdy(true); });
    return () => { active = false; };
  }, [user?.id, toast]); // só recarrega quando troca o usuário (login/logout), não em foco/refresh

  // ─ Autenticação (Supabase) ─
  const userIdRef = useRef(null);
  useEffect(() => {
    let mounted = true;
    async function applySession(session, greet) {
      if (session?.user) {
        const { data: prof } = await supabase.from("profiles").select("name,is_admin,apelido").eq("id", session.user.id).single();
        if (!mounted) return;
        // is_viewer é buscado à parte e tolerante a falha: se a coluna ainda não
        // existir (migração não aplicada), o login segue normal (viewer=false),
        // sem quebrar a detecção de admin.
        let isViewerFlag = false, isComercialFlag = false, buVal = "";
        const vres = await supabase.from("profiles").select("is_viewer").eq("id", session.user.id).single();
        if (!vres.error) isViewerFlag = !!vres.data?.is_viewer;
        // is_comercial + bu: idem — tolerante à migração 0031 ainda não aplicada.
        const cres = await supabase.from("profiles").select("is_comercial,bu").eq("id", session.user.id).single();
        if (!cres.error) { isComercialFlag = !!cres.data?.is_comercial; buVal = cres.data?.bu || ""; }
        if (!mounted) return;
        const next = { id: session.user.id, name: prof?.name || session.user.email, isAdmin: !!prof?.is_admin, isViewer: isViewerFlag, isComercial: isComercialFlag, bu: buVal, apelido: prof?.apelido || "", email: session.user.email };
        const isNewLogin = userIdRef.current !== next.id;
        userIdRef.current = next.id;
        // Mantém a MESMA referência do usuário se nada mudou — evita recarregar
        // a tela (e fechar modais) quando o app volta do foco / renova o token.
        setUser(prev => (prev && prev.id===next.id && prev.name===next.name && prev.isAdmin===next.isAdmin && prev.isViewer===next.isViewer && prev.isComercial===next.isComercial && prev.bu===next.bu && prev.apelido===next.apelido) ? prev : next);
        if (greet && isNewLogin) { setPage("home"); toast(`Bem-vindo(a), ${next.apelido || (next.name||"").split(" ")[0]}!`, "info"); }
      } else if (mounted) {
        userIdRef.current = null;
        setUser(null);
      }
    }
    supabase.auth.getSession().then(({ data }) => { applySession(data.session, false).finally(()=>{ if (mounted) setAuthRdy(true); }); });
    const { data: sub } = supabase.auth.onAuthStateChange((event, session) => {
      if (event === "PASSWORD_RECOVERY") setRecovery(true);
      applySession(session, event === "SIGNED_IN");
    });
    return () => { mounted = false; sub.subscription.unsubscribe(); };
  }, []);

  const isAdmin = user?.isAdmin || false;
  const isViewer = user?.isViewer || false;   // somente-visualização: lê tudo, não escreve
  const isComercial = user?.isComercial || false;  // comercial: só a receita da sua BU
  const userBu = user?.bu || "";
  // Receitas "ativas" = fora as sinalizadas "fora do relatório" (ausenteRelatorio).
  // Telas de RESUMO de receita usam esta lista para não somar o que saiu do
  // relatório. Conciliação e Minha visão (revisão) continuam vendo tudo.
  const recordsAtivos = records.filter(r => !r.ausenteRelatorio);
  // Comercial só enxerga a receita da sua BU. O RLS já limita o fetch; este
  // filtro é a visão do app (e defesa extra caso a migração não esteja aplicada).
  const recordsView = isComercial ? recordsAtivos.filter(r => r.bu === userBu) : recordsAtivos;
  // Comercial: trava a navegação nas telas permitidas (o RLS trava o dado).
  useEffect(() => { if (isComercial && !COMERCIAL_PAGES.has(page)) setPage("dash"); }, [isComercial, page]);
  // Trava de UX para o viewer (a trava real é o RLS). Retorna true se bloqueou.
  const blockIfViewer = () => { if (isViewer) { toast("Acesso somente visualização — ação não permitida.", "info"); return true; } return false; };

  async function handleUpdateBulk(updatedList) {
    if (blockIfViewer()) return;
    try {
      await db.upsertRecords(updatedList);
      await reloadRecords();
      toast(`Passos atualizados — ${updatedList.length} profissional(is)`);
    } catch(e) { toast("Erro ao salvar os passos: "+e.message, "error"); }
  }

  async function handleSaveClass(id, { motivo, obs }) {
    if (blockIfViewer()) return;
    try {
      await db.updateRecordClass(id, { motivo, obs });
      setRecords(rs => rs.map(r => r.id===id ? { ...r, classMotivo:motivo||"", classObs:obs||"" } : r));
      toast("Classificação salva");
    } catch(e) { toast("Erro ao salvar classificação: "+e.message, "error"); }
  }

  async function handleImport({ records:newRecs, competencia, empresa, tipo, mode, note }) {
    if (blockIfViewer()) return;
    try {
      if (mode==="merge") {
        // Re-importação: casa por empresa+tipo+PEP+profissional+competência+PERÍODO.
        // O período (início/fim) na chave é o que distingue clientes de faturamento
        // quebrado (10 a 10, 20 a 20): a mesma pessoa vem em 2 peças (ex.: 01–10 e
        // 11–31) e cada peça casa 1:1 com a sua, sem trocar valores. Cliente normal
        // (mês cheio) tem período fixo, então nada muda pra ele.
        const norm = s => (s||"").toString().trim().toLowerCase();
        const dnorm = s => String(s||"").replace(/\D/g,"");   // data: só dígitos (tolera separador)
        const keyOf = r => `${norm(r.empresa)}|${norm(r.tipo)}|${norm(pepBase(r.pep))}|${norm(r.profissional)}|${(r.competencia||"").trim()}|${dnorm(r.inicio)}|${dnorm(r.fim)}`;
        // Chave SEM empresa: identifica a mesma receita quando só o código da
        // empresa mudou (correção de digitação). O PEP já é específico de empresa
        // (BR02CLP…), então colisão entre empresas diferentes não acontece.
        const keyNoEmp = r => `${norm(r.tipo)}|${norm(pepBase(r.pep))}|${norm(r.profissional)}|${(r.competencia||"").trim()}|${dnorm(r.inicio)}|${dnorm(r.fim)}`;
        // Chave SEM PEP (mantém empresa+cliente): identifica a mesma receita quando
        // só o PEP mudou. O cliente entra pra nunca confundir clientes diferentes.
        const keyNoPep = r => `${norm(r.empresa)}|${norm(r.tipo)}|${norm(r.cliente)}|${norm(r.profissional)}|${(r.competencia||"").trim()}|${dnorm(r.inicio)}|${dnorm(r.fim)}`;
        const combos = new Set(newRecs.map(r=>`${r.competencia}|${r.empresa}|${r.tipo}`));
        const escopo = records.filter(r => combos.has(`${r.competencia}|${r.empresa}|${r.tipo}`));
        // A chave pode repetir (mesmo profissional/PEP com 2+ linhas). Guardamos
        // uma FILA por chave e casamos por posição: 1ª linha ↔ 1º registro, etc.
        // A chave pode repetir (mesmo profissional/PEP com 2+ linhas). Guardamos
        // uma FILA por chave e casamos POR VALOR: as linhas sem mudança batem
        // exatamente com a sua gêmea; só a que realmente mudou preenche o resto.
        const byKey = {}; escopo.forEach(r => { const k=keyOf(r); (byKey[k]=byKey[k]||[]).push(r); });
        const newByKey = {}; newRecs.forEach(nr => { const k=keyOf(nr); (newByKey[k]=newByKey[k]||[]).push(nr); });
        const consumidos = new Set();
        const upserts = []; const inserts = []; const snapshot = [];
        const importId = uuid();
        let novos=0, mudados=0, congelados=0, movidos=0, movidosPep=0;
        const casar = (nr, ex) => {
          consumidos.add(ex.id);
          snapshot.push(ex);   // guarda o estado ANTERIOR para permitir desfazer
          const antigo = ex.valorTotal||0, novo = nr.valorTotal||0;
          const mudou = Math.abs(novo-antigo) > 0.01;
          const fat = fatByRec[ex.id]||0;
          // CONCILIADO = verdade final: não altera o valor. Só sinaliza a
          // divergência para o usuário reabrir a conciliação e atualizar.
          if (fat > 0.001) {
            const merged = { ...ex, ausenteRelatorio:false, updatedAt: nowISO(),
              valorBaseDivergente: mudou ? novo : null };
            if (mudou) congelados++;
            upserts.push(merged);
            return;
          }
          // Não conciliado: atualiza normalmente. `empresa` é refrescada do
          // relatório — em casamento normal é igual (no-op); num MOVE por correção
          // de empresa, é o que de fato transfere a receita para o código certo.
          const merged = { ...ex,
            empresa: nr.empresa||ex.empresa,
            pep: nr.pep||ex.pep,   // sincroniza o PEP do relatório; num MOVE por correção de PEP, é o que transfere a receita
            cliente: nr.cliente||ex.cliente, codCliente: nr.codCliente||ex.codCliente,
            inicio: nr.inicio||ex.inicio, fim: nr.fim||ex.fim,
            valorVenda: nr.valorVenda, hrsAprovadas: nr.hrsAprovadas,
            valorTotal: nr.valorTotal, valorLiquido: nr.valorLiquido,
            ausenteRelatorio: false, updatedAt: nowISO(),
          };
          if (mudou && calcStatus(ex.progress) !== "Não iniciado") {  // cenário 1 = silêncio
            merged.valorAnterior = antigo;
            merged.valorAlteradoEm = nowISO();
            mudados++;
          }
          upserts.push(merged);
        };
        // Índice por identidade SEM empresa (mesma competência, todas as empresas)
        // para detectar receita que só teve o código de empresa corrigido.
        const compsNew = new Set(newRecs.map(r=>(r.competencia||"").trim()));
        const byNoEmp = {};
        records.forEach(r => { if (!compsNew.has((r.competencia||"").trim())) return; const k=keyNoEmp(r); (byNoEmp[k]=byNoEmp[k]||[]).push(r); });
        // Tenta mover um registro de OUTRA empresa (mesmo PEP/profissional/período)
        // em vez de criar um novo. Não mexe em conciliado. Retorna true se moveu.
        const tentarMover = (nr) => {
          const q = byNoEmp[keyNoEmp(nr)];
          if (!q || !q.length) return false;
          const idx = q.findIndex(ex => !consumidos.has(ex.id) && norm(ex.empresa)!==norm(nr.empresa) && (fatByRec[ex.id]||0)<=0.001);
          if (idx<0) return false;
          casar(nr, q.splice(idx,1)[0]);   // casar já seta empresa ← nr.empresa
          return true;
        };
        // Índice por identidade SEM PEP (mesma empresa+CLIENTE+tipo+profissional+
        // período) para detectar receita que só mudou de PEP (ex.: fee que trocou
        // de código). Inclui o CLIENTE para nunca colidir entre clientes.
        const byNoPep = {};
        records.forEach(r => { if (!compsNew.has((r.competencia||"").trim())) return; const k=keyNoPep(r); (byNoPep[k]=byNoPep[k]||[]).push(r); });
        // Move um registro de OUTRO PEP para o PEP novo — só quando é INEQUÍVOCO:
        // candidato ÚNICO, PEP diferente, MESMO valor e NÃO conciliado. Se houver
        // 0 ou vários candidatos, não move (cria novo), pra não juntar errado.
        const tentarMoverPep = (nr) => {
          const q = byNoPep[keyNoPep(nr)];
          if (!q || !q.length) return false;
          const cand = q.filter(ex => !consumidos.has(ex.id)
            && norm(pepBase(ex.pep)) !== norm(pepBase(nr.pep))
            && (fatByRec[ex.id]||0) <= 0.001
            && Math.abs((ex.valorTotal||0)-(nr.valorTotal||0)) < 0.01);
          if (cand.length !== 1) return false;
          const ex = cand[0]; const i = q.indexOf(ex); if (i>=0) q.splice(i,1);
          casar(nr, ex);   // casar seta pep ← nr.pep (move o PEP)
          return true;
        };
        Object.entries(newByKey).forEach(([k, news]) => {
          const bucket = (byKey[k]||[]).slice();
          const pend = [];
          // 1) casa exato por valor (linhas inalteradas travam na gêmea certa)
          news.forEach(nr => {
            const i = bucket.findIndex(ex => Math.abs((ex.valorTotal||0)-(nr.valorTotal||0)) < 0.01);
            if (i>=0) casar(nr, bucket.splice(i,1)[0]);
            else pend.push(nr);
          });
          // 2) as que sobraram: casa com o restante mais próximo no mesmo escopo;
          //    se o escopo estiver vazio, tenta MOVER (correção de empresa) antes
          //    de criar registro novo.
          pend.forEach(nr => {
            if (bucket.length) {
              let bi=0, bd=Infinity;
              bucket.forEach((ex,idx)=>{ const d=Math.abs((ex.valorTotal||0)-(nr.valorTotal||0)); if(d<bd){bd=d;bi=idx;} });
              casar(nr, bucket.splice(bi,1)[0]);
            } else if (tentarMover(nr)) {
              movidos++;
            } else if (tentarMoverPep(nr)) {
              movidosPep++;
            } else {
              inserts.push({ ...nr, importId }); novos++;
            }
          });
        });
        const absentIds = escopo.filter(r => !consumidos.has(r.id) && !r.ausenteRelatorio).map(r=>r.id);
        await db.mergeImport({ upserts, inserts, absentIds });
        try { await db.insertHistory({ competencia, empresa, tipo:[...new Set(newRecs.map(r=>r.tipo))].join("/")||tipo, mode, count:newRecs.length, user:user.name, note, importId, snapshot }); } catch {}
        await Promise.all([reloadRecords(), reloadFaturamentos(), reloadHistory()]);
        setState(s=>({...s, competenciaAtual:competencia}));
        toast(`Mês atualizado — ${novos} novo(s), ${mudados} com valor alterado${movidos?` · ${movidos} com empresa corrigida`:""}${movidosPep?` · ${movidosPep} com PEP corrigido`:""}${congelados?` · ${congelados} conciliado(s) divergindo (reabra p/ atualizar)`:""}, ${absentIds.length} fora do relatório`);
        return;
      }
      if (mode==="replace") {
        // Verdade final: não apaga o que já está conciliado. Bloqueia se houver.
        const combosR = new Set(newRecs.map(r=>`${r.competencia}|${r.empresa}|${r.tipo}`));
        const conc = records.filter(r => combosR.has(`${r.competencia}|${r.empresa}|${r.tipo}`) && (fatByRec[r.id]||0)>0.001);
        if (conc.length) { toast(`Há ${conc.length} registro(s) conciliado(s) neste recorte. Reabra a conciliação antes de substituir.`, "error"); return; }
        // Substitui apenas os recortes (competência+empresa+tipo) presentes na carga.
        const combos = [...new Set(newRecs.map(r=>`${r.competencia}|${r.empresa}|${r.tipo}`))];
        for (const c of combos) { const [cmp,emp,tp]=c.split("|"); await db.deleteRecordsBy({ competencia:cmp, empresa:emp, tipo:tp }); }
      }
      const importId = uuid();
      const tagged = newRecs.map(r=>({ ...r, importId }));
      await db.insertRecords(tagged);
      const tipoLog = [...new Set(newRecs.map(r=>r.tipo))].join("/") || tipo;
      const empLog  = [...new Set(newRecs.map(r=>r.empresa))].join("/") || empresa;
      try { await db.insertHistory({ competencia, empresa:empLog, tipo:tipoLog, mode, count:newRecs.length, user:user.name, note, importId }); } catch {}
      await Promise.all([reloadRecords(), reloadHistory()]);
      setState(s=>({...s, competenciaAtual:competencia}));
      toast(`${newRecs.length} registros importados (${mode==="replace"?"substituição":"adição"})`);
    } catch(e) { toast("Erro na importação: "+e.message, "error"); }
  }

  async function handleUndoImport(entry) {
    if (blockIfViewer()) return;
    if (!entry?.importId) { toast("Esta importação é antiga e não pode ser desfeita automaticamente.", "error"); return; }
    try {
      // 1) apaga os registros incluídos por esta importação
      const removed = await db.deleteRecordsByImport(entry.importId);
      // 2) restaura os registros que esta importação atualizou (merge) ao estado anterior
      let restored = 0;
      if (Array.isArray(entry.snapshot) && entry.snapshot.length) { await db.restoreRecords(entry.snapshot); restored = entry.snapshot.length; }
      await db.deleteHistory(entry.id);
      await Promise.all([reloadRecords(), reloadFaturamentos(), reloadHistory()]);
      toast(`Importação desfeita — ${removed} incluído(s) removido(s)${restored?` · ${restored} restaurado(s)`:""}`, "info");
    } catch(e) { toast("Erro ao desfazer importação: "+e.message, "error"); }
  }

  function handleCompetencia(val) { setState(s=>({...s, competenciaAtual:val})); }

  async function handleRecordDelete(id) {
    try {
      const rec = records.find(r => r.id === id);
      // Lotes de conciliação aos quais este registro está alocado.
      const cids = new Set(faturamentos.filter(a => a.recordId === id && a.conciliacaoId).map(a => a.conciliacaoId));
      if (rec?.conciliacaoId) cids.add(rec.conciliacaoId);   // conciliação antiga
      await db.deleteRecord(id);   // cascade remove as alocações deste registro
      // Reabre as notas de cada lote que ficou sem nenhuma outra receita alocada.
      const cidsToFree = [...cids].filter(cid => !faturamentos.some(a => a.conciliacaoId === cid && a.recordId !== id));
      if (cidsToFree.length) await db.freeNotes(cidsToFree);
      await Promise.all([reloadRecords(), reloadNotes(), reloadFaturamentos()]);
      toast(`Registro excluído${cidsToFree.length ? " · nota(s) reaberta(s) na conciliação" : ""}`, "info");
    } catch(e) { toast("Erro ao excluir registro: "+e.message, "error"); }
  }
  async function handleClearAlert(id) {
    if (blockIfViewer()) return;
    try { await db.clearRecordAlert(id); await reloadRecords(); toast("Alerta baixado", "info"); }
    catch(e) { toast("Erro ao baixar alerta: "+e.message, "error"); }
  }

  // ── Correções (admin): edita/apaga/mescla um registro com desfazer (in-memory).
  // O "before" guardado permite reverter via db.restoreRecords (upsert com id).
  const [lastCorr, setLastCorr] = useState(null);
  async function handleCorrEdit(updated) {
    if (blockIfViewer()) return;
    const before = records.find(r => r.id === updated.id);
    try {
      await db.upsertRecords([{ ...updated, updatedAt: nowISO() }]);
      await reloadRecords();
      if (before) setLastCorr({ kind:"edit", before, label: `edição de ${before.cliente}` });
      toast("Registro corrigido");
    } catch(e) { toast("Erro ao corrigir: "+e.message, "error"); }
  }
  // Classificar BU de todas as receitas de um cliente (bulk). Atualização otimista.
  async function handleSetBu(ids, bu) {
    if (blockIfViewer() || !ids || !ids.length) return;
    try {
      await db.setRecordsBu(ids, bu);
      const idSet = new Set(ids);
      setRecords(rs => rs.map(r => idSet.has(r.id) ? { ...r, bu } : r));
      toast(`BU aplicada a ${ids.length} receita(s)`);
    } catch(e) { toast("Erro ao classificar BU: "+e.message, "error"); }
  }
  async function handleCorrDelete(id) {
    if (blockIfViewer()) return;
    const before = records.find(r => r.id === id);
    try {
      // Só chega aqui para registros NÃO conciliados (o painel trava conciliados),
      // então não há lote/nota a reabrir — o desfazer via restoreRecords é fiel.
      await db.deleteRecord(id);
      await Promise.all([reloadRecords(), reloadFaturamentos()]);
      if (before) setLastCorr({ kind:"delete", before, label: `exclusão de ${before.profissional||before.pep} (${before.cliente})` });
      toast("Registro apagado", "info");
    } catch(e) { toast("Erro ao apagar: "+e.message, "error"); }
  }
  async function handleCorrInsert(rec) {
    if (blockIfViewer()) return;
    // id gerado no cliente → upsert insere (id inédito). Guarda o id p/ desfazer.
    const novo = { ...rec, id: uuid(), progress: rec.progress || {}, updatedAt: nowISO() };
    try {
      await db.upsertRecords([novo]);
      await reloadRecords();
      setLastCorr({ kind:"insert", id: novo.id, label: `inclusão de ${rec.profissional||rec.pep} (${rec.cliente})` });
      toast("Registro incluído");
    } catch(e) { toast("Erro ao incluir: "+e.message, "error"); }
  }
  async function handleCorrUndo() {
    if (blockIfViewer() || !lastCorr) return;
    try {
      if (lastCorr.kind==="insert") await db.deleteRecord(lastCorr.id);   // desfazer inclusão = apagar o que foi criado
      else await db.restoreRecords([lastCorr.before]);                     // edição/exclusão = restaura o estado anterior
      await Promise.all([reloadRecords(), reloadFaturamentos()]);
      setLastCorr(null);
      toast("Correção desfeita");
    } catch(e) { toast("Erro ao desfazer: "+e.message, "error"); }
  }

  async function handleTaskAdd(t)    { if(blockIfViewer())return; try { await db.insertTask(t); await reloadTasks(); toast("Tarefa criada"); } catch(e){ toast("Erro ao criar tarefa: "+e.message,"error"); } }
  async function handleTaskUpdate(u) { if(blockIfViewer())return; try { await db.updateTask(u); await reloadTasks(); } catch(e){ toast("Erro ao atualizar tarefa: "+e.message,"error"); } }
  async function handleTaskDelete(id){ if(blockIfViewer())return; try { await db.deleteTask(id); await reloadTasks(); toast("Tarefa excluída","info"); } catch(e){ toast("Erro ao excluir tarefa: "+e.message,"error"); } }

  // ─ Entregas recorrentes ─
  async function handleTemplateSave(t)   { try { t.id ? await db.updateTemplate(t) : await db.insertTemplate(t); await reloadTemplates(); toast(t.id?"Modelo atualizado":"Modelo de entrega criado"); } catch(e){ toast("Erro ao salvar modelo: "+e.message,"error"); } }
  async function handleTemplateDelete(id){ try { await db.deleteTemplate(id); await reloadTemplates(); toast("Modelo excluído","info"); } catch(e){ toast("Erro ao excluir modelo: "+e.message,"error"); } }
  async function handleGenerateDelivery(template, competencia) {
    try {
      const analistas = [...new Set([...profiles.filter(p=>!p.isAdmin).map(p=>p.name), ...records.map(r=>r.responsavel)].filter(Boolean))];
      const { count } = await db.generateDelivery(template, competencia, analistas);
      await Promise.all([reloadTasks(), reloadDeliveries()]);
      toast(`Entrega gerada — ${count} tarefa(s) para ${competencia}`);
    } catch(e){ toast("Erro ao gerar entrega: "+e.message,"error"); }
  }

  // ─ Gestão de acessos (Supabase) ─
  async function handleProfileUpdate(data) {
    if (blockIfViewer()) return;
    // Impede rebaixar o último administrador.
    if (data.isAdmin === false) {
      const target = profiles.find(p => p.id === data.id);
      if (target?.isAdmin && profiles.filter(p => p.isAdmin).length <= 1) { toast("É preciso manter ao menos um administrador.", "error"); return; }
    }
    try {
      await db.updateProfile(data);
      await reloadProfiles();
      // Se o admin alterou o próprio papel/nome, reflete na sessão atual.
      if (data.id === user.id) setUser(u => ({ ...u, name: data.name, isAdmin: data.isAdmin, isViewer: data.isViewer, isComercial: data.isComercial, bu: data.bu ?? u.bu, apelido: data.apelido ?? u.apelido }));
      toast(`Acesso de ${(data.name||"").split(" ")[0]} atualizado`);
    } catch(e) { toast("Erro ao atualizar acesso: "+e.message, "error"); }
  }
  async function handleProfileRemove(profile) {
    if (profile.id === user.id) { toast("Você não pode remover o próprio acesso.", "error"); return; }
    if (profile.isAdmin && profiles.filter(p => p.isAdmin).length <= 1) { toast("É preciso manter ao menos um administrador.", "error"); return; }
    try {
      await db.deleteProfile(profile.id);
      await reloadProfiles();
      toast("Acesso removido", "info");
    } catch(e) { toast("Erro ao remover acesso: "+e.message, "error"); }
  }
  // ─ Clientes ─
  async function handleClientSave(c) {
    if (blockIfViewer()) return;
    try {
      if (c.id) { await db.updateClient(c); toast("Cliente atualizado"); }
      else { await db.insertClient(c); toast("Cliente cadastrado"); }
      await reloadClients();
    } catch(e) { toast("Erro ao salvar cliente: "+e.message, "error"); }
  }
  async function handleClientDelete(id) {
    if (blockIfViewer()) return;
    try { await db.deleteClient(id); await reloadClients(); toast("Cliente excluído", "info"); }
    catch(e) { toast("Erro ao excluir cliente: "+e.message, "error"); }
  }
  async function handleClientsImport(list) {
    try { const n = await db.bulkInsertClients(list); await reloadClients(); toast(`${n} cliente(s) importados (cadastro incompleto)`); }
    catch(e) { toast("Erro ao importar clientes: "+e.message, "error"); }
  }
  // DE→PARA de clientes (unificação de nomes). Recarrega os registros para
  // reaplicar o mapa em todas as telas.
  async function handleAliasSave({ de, para }) {
    if (blockIfViewer()) return;
    const d=(de||"").trim(), p=(para||"").trim();
    if (!d || !p) { toast("Preencha os dois nomes (DE e PARA).", "error"); return; }
    if (_normCliNome(d)===_normCliNome(p)) { toast("DE e PARA são o mesmo nome.", "error"); return; }
    try { await db.saveClientAlias({ de:d, para:p }); await reloadAliases(); await reloadRecords(); toast(`Unificado: "${d}" → "${p}"`); }
    catch(e){ toast("Erro ao salvar unificação: "+e.message, "error"); }
  }
  async function handleAliasDelete(de) {
    if (blockIfViewer()) return;
    try { await db.deleteClientAlias(de); await reloadAliases(); await reloadRecords(); toast("Unificação removida", "info"); }
    catch(e){ toast("Erro ao remover: "+e.message, "error"); }
  }
  // Agrupa N cadastros num só. baseId define qual cadastro é o destino (grupo);
  // sem baseId, usa o primeiro da lista. Reúne os CNPJs no destino e remove os demais.
  async function handleClientsMerge(ids, nome, baseId) {
    if (blockIfViewer()) return;
    try {
      const sel = clients.filter(c=>ids.includes(c.id));
      if (sel.length<2) return;
      const base = (baseId && sel.find(c=>c.id===baseId)) || sel[0];
      const ordered = [base, ...sel.filter(c=>c.id!==base.id)];
      // Junta as entidades (empresas) de todos os cadastros, sem CNPJ duplicado.
      const ents = []; const seen = new Set();
      ordered.forEach(c=>{
        const arr = parseJSON(c.cnpjs, null);
        const list = (Array.isArray(arr)&&arr.length) ? arr : [{ razao:c.nome||"", cnpj:c.cnpj||"", codSap:c.codSap||"" }];
        list.forEach(e=>{
          const key = (e.cnpj||"").replace(/\D/g,"") || (e.razao||"").toLowerCase();
          if (key && seen.has(key)) return;
          if (key) seen.add(key);
          ents.push({ razao:(e.razao||"").trim(), cnpj:(e.cnpj||"").replace(/\D/g,"").slice(0,14), codSap:(e.codSap||"").trim() });
        });
      });
      const merged = { ...base, nome, cnpjs: JSON.stringify(ents), cnpj: ents[0]?.cnpj || "", codSap: ents[0]?.codSap || base.codSap || "" };
      await db.updateClient(merged);
      for (const c of ordered.slice(1)) await db.deleteClient(c.id);
      await reloadClients();
      toast(`Grupo "${nome}" com ${ents.length} CNPJ(s)`);
    } catch(e) { toast("Erro ao agrupar clientes: "+e.message, "error"); }
  }
  // ─ Conciliação de notas (prefeitura) ─
  async function handleNotesImport(list) {
    try {
      // Dedup por prestador+número (campos INTRÍNSECOS da nota, não a empresa
      // escolhida no modal): nota já existente é ATUALIZADA (não duplica). A
      // empresa era escolhida na importação e variava (NULL vs BR02), o que
      // furava o dedup e duplicava toda a base.
      const digits = s => String(s||"").replace(/\D/g,"");
      const key = n => `${digits(n.prestadorCnpj)}|${String(n.numero||"").trim()}`;
      const existing = {}; notes.forEach(n => { existing[key(n)] = n; });
      const toInsert = [], toUpdate = [];
      list.forEach(n => { const ex = existing[key(n)]; ex ? toUpdate.push({ ex, novo: n }) : toInsert.push(n); });
      if (toInsert.length) await db.insertMunicipalNotes(toInsert);
      // Notas já existentes NÃO são reimportadas — só agimos se houve NOVO
      // cancelamento (coluna de cancelamento da prefeitura). Importação diária.
      const reabrir = []; let canceladas = 0;
      for (const { ex, novo } of toUpdate) {
        if (novo.cancelada && !ex.cancelada) {
          canceladas++;
          await db.updateMunicipalNote(ex.id, { ...ex, cancelada: true, situacao: novo.situacao || ex.situacao });
          // Reabre o lote inteiro (reabre o saldo das receitas e libera as notas).
          if (ex.conciliacaoId) { await reopenCid(ex.conciliacaoId); canceladas += 0; }
          else records.filter(r => r.municipalNoteId === ex.id).forEach(r => reabrir.push({ id: r.id, progress: { ...(r.progress||{}), p5_nf:false, p5_data_nf:"", p5_no_corte:false } }));
        }
      }
      if (reabrir.length) await db.reopenRecords(reabrir);
      await Promise.all([reloadNotes(), reloadRecords(), reloadFaturamentos()]);
      const base = `${toInsert.length} nova(s) · ${toUpdate.length} já existiam${canceladas?` · ${canceladas} cancelada(s)`:""}`;
      if (reabrir.length) toast(`${base} · ${reabrir.length} registro(s) reabertos: NF cancelada`, "error");
      else toast(`Importação: ${base}`);
    } catch(e) { toast("Erro ao importar notas: "+e.message, "error"); }
  }
  async function handleNotesUndo(importId) {
    try {
      // Reabre os lotes conciliados dessas notas ANTES de apagar — senão a receita
      // ficaria faturada sem nota (alocação órfã, o caso do Erik). O trigger
      // guard_faturamento_orfao no banco barra o resto.
      const cids = [...new Set(notes.filter(n => n.importId === importId && n.conciliacaoId).map(n => n.conciliacaoId))];
      for (const cid of cids) await reopenCid(cid);
      const n = await db.deleteMunicipalNotesByImport(importId);
      await Promise.all([reloadNotes(), reloadRecords(), reloadFaturamentos()]);
      toast(`${n} nota(s) removida(s)${cids.length?` · ${cids.length} lote(s) reaberto(s)`:""}`, "info"); }
    catch(e) { toast("Erro ao desfazer importação: "+e.message, "error"); }
  }
  // Conciliação N:N com faturamento PARCIAL. valoresMap: {recordId: valor a faturar}.
  async function handleConciliate(recordIds, notesArr, valoresMap) {
    if (blockIfViewer()) return;
    try {
      const cid = uuid();
      const numero = notesArr.map(n => n.numero).join(", ");
      const dataNf = notesArr.map(n => n.emitidaEm).filter(Boolean).sort()[0] || "";
      const allocations = []; const recordItems = [];
      recordIds.forEach(id => {
        const r = records.find(x => x.id === id); if (!r) return;
        const jaFat = fatByRec[r.id] || 0;
        const faturavel = (r.valorTotal||0) + (varByRec[r.id]||0);   // inclui variação pós-fechamento
        const saldo = faturavel - jaFat;   // com sinal (aceita desconto negativo)
        const bruto = valoresMap?.[id] ?? saldo;
        // Clampa ao saldo respeitando o sinal (não ultrapassa nem inverte).
        const valor = saldo >= 0 ? Math.max(0, Math.min(bruto, saldo)) : Math.min(0, Math.max(bruto, saldo));
        if (Math.abs(valor) <= 0.001) return;
        allocations.push({ recordId: id, valor });
        const full = Math.abs(faturavel - (jaFat + valor)) <= 0.01;
        const progress = full ? faturarProgress(r, { emitidaEm: dataNf }) : { ...(r.progress||{}), p5_liberado: true };
        const nfNumero = [...new Set([r.nfNumero, numero].join(", ").split(", ").map(s=>s.trim()).filter(Boolean))].join(", ");
        recordItems.push({ id, progress, nfNumero });
      });
      if (!allocations.length) { toast("Informe um valor para faturar.", "error"); return; }
      await db.conciliateSet({ cid, allocations, recordItems, noteIds: notesArr.map(n => n.id), userName: user.name });
      await Promise.all([reloadRecords(), reloadNotes(), reloadFaturamentos()]);
      const parciais = allocations.filter(a => { const r = records.find(x=>x.id===a.recordId); return r && (fatByRec[r.id]||0)+a.valor < ((r.valorTotal||0)+(varByRec[r.id]||0))-0.01; }).length;   // faturável (receita + variação)
      toast(`${allocations.length} receita(s) × ${notesArr.length} nota(s) conciliadas${parciais?` · ${parciais} parcial(is)`:""}`);
    } catch(e) { toast("Erro ao conciliar: "+e.message, "error"); }
  }
  // Reabre um LOTE de conciliação: remove suas alocações, recalcula o saldo dos
  // registros afetados e libera as notas do lote.
  async function reopenCid(cid) {
    const allocs = faturamentos.filter(a => a.conciliacaoId === cid);
    const recIds = [...new Set(allocs.map(a => a.recordId))];
    const recordItems = recIds.map(id => {
      const r = records.find(x => x.id === id) || {};
      const removido = allocs.filter(a => a.recordId === id).reduce((s,a)=>s+(a.valor||0),0);
      const restante = (fatByRec[id]||0) - removido;
      const faturavel = (r.valorTotal||0) + (varByRec[r.id]||0);   // receita + variação
      const full = restante >= faturavel - 0.01 && faturavel > 0;
      const progress = { ...(r.progress||{}), p5_liberado: r.progress?.p5_liberado ?? true, p5_nf: full, p5_no_corte: full, p5_data_nf: full ? (r.progress?.p5_data_nf||"") : "" };
      const item = { id, progress, nfNumero: restante>0.001 ? (r.nfNumero||"") : "", conciliadoEm: restante>0.001 ? r.conciliadoEm : null, conciliadoPor: restante>0.001 ? r.conciliadoPor : null };
      // Se a base divergiu enquanto estava conciliado, ao reabrir aplicamos o
      // valor novo e limpamos o marcador — a receita volta com o valor correto.
      if (r.valorBaseDivergente != null) { item.valorTotal = r.valorBaseDivergente; item.valorBaseDivergente = null; }
      return item;
    });
    await db.reopenConciliacao(cid, recordItems);
  }
  async function handleReopenGroup({ conciliacaoId, noteId }) {
    if (blockIfViewer()) return;
    try {
      if (conciliacaoId) await reopenCid(conciliacaoId);
      else { // conciliações antigas (1 nota por registro, sem alocação)
        const recs = records.filter(r => r.municipalNoteId === noteId);
        if (recs.length) await db.reopenRecords(recs.map(r => ({ id: r.id, progress: { ...(r.progress||{}), p5_nf:false, p5_data_nf:"", p5_no_corte:false } })));
      }
      await Promise.all([reloadRecords(), reloadNotes(), reloadFaturamentos()]);
      toast("Conciliação desfeita — saldo reaberto", "info");
    } catch(e) { toast("Erro ao desfazer conciliação: "+e.message, "error"); }
  }
  // Notas órfãs: têm conciliacao_id mas nenhuma receita continua alocada ao lote
  // (ex.: receita excluída antes do fix de reabertura). Podem ser reabertas.
  const cidsAtivos = new Set(faturamentos.filter(a => a.conciliacaoId).map(a => a.conciliacaoId));
  const notasOrfas = notes.filter(n => n.conciliacaoId && !cidsAtivos.has(n.conciliacaoId) && !records.some(r => r.municipalNoteId === n.id));
  async function handleReopenOrphans() {
    try {
      const cids = [...new Set(notasOrfas.map(n => n.conciliacaoId))];
      if (!cids.length) { toast("Nenhuma nota órfã encontrada.", "info"); return; }
      await db.freeNotes(cids);
      await Promise.all([reloadNotes(), reloadRecords()]);
      toast(`${notasOrfas.length} nota(s) reaberta(s) — voltaram para pendentes`, "info");
    } catch(e) { toast("Erro ao reabrir notas órfãs: "+e.message, "error"); }
  }
  async function handleNoteDelete(note) {
    try {
      if (note.conciliacaoId) await reopenCid(note.conciliacaoId);
      else { const recs = records.filter(r => r.municipalNoteId === note.id); if (recs.length) await db.reopenRecords(recs.map(r => ({ id: r.id, progress: { ...(r.progress||{}), p5_nf:false, p5_data_nf:"", p5_no_corte:false } }))); }
      await db.deleteMunicipalNote(note.id);
      await Promise.all([reloadNotes(), reloadRecords(), reloadFaturamentos()]);
      toast(`Nota ${note.numero} removida`, "info");
    } catch(e) { toast("Erro ao remover nota: "+e.message, "error"); }
  }

  async function handleMuralSave(m) {
    if (blockIfViewer()) return;
    try { await db.saveMural(m); await reloadMural(); toast("Mural atualizado"); }
    catch(e) { toast("Erro ao salvar mural: "+e.message, "error"); }
  }
  async function handleSaveApelido(apelido) {
    try { await db.setMyApelido(apelido); setUser(u=>({...u, apelido})); await reloadProfiles(); toast("Apelido atualizado"); }
    catch(e) { toast("Erro ao salvar apelido: "+e.message, "error"); }
  }

  // Quanto de cada registro já foi faturado (soma das alocações). Legado: registros
  // marcados como faturados (p5_nf) sem alocação contam pelo valor total.
  const fatByRec = {};
  faturamentos.forEach(a => { fatByRec[a.recordId] = (fatByRec[a.recordId]||0) + (a.valor||0); });
  records.forEach(r => { if (!(r.id in fatByRec) && isFaturado(r.progress)) fatByRec[r.id] = r.valorTotal||0; });

  // Variação de receita pós-fechamento (faturável) — soma por registro + lista.
  const varByRec = {}; const varsByRec = {};
  variacoes.forEach(v => { varByRec[v.recordId] = (varByRec[v.recordId]||0) + (v.valor||0); (varsByRec[v.recordId] = varsByRec[v.recordId] || []).push(v); });
  // Clientes que aceitam variação (marcados no cadastro) — casa pelo nome.
  const normCli = s => (s||"").toString().trim().toLowerCase();
  const varClientes = clients.filter(c => c.aceitaVariacao).map(c => normCli(c.nome)).filter(Boolean);
  const aceitaVar = (r) => { const rc = normCli(r.cliente); return varClientes.some(cn => rc===cn || (cn.length>4 && rc.includes(cn))); };

  async function handleAddVariacao(recordId, valor, motivo) {
    if (blockIfViewer()) return;
    try { await db.insertVariacao({ recordId, valor, motivo, criadoPor: user.name }); await Promise.all([reloadVariacoes(), reloadRecords()]); toast("Variação lançada — vai aparecer como saldo a faturar"); }
    catch(e){ toast("Erro ao lançar variação: "+e.message, "error"); }
  }
  async function handleDelVariacao(id) {
    if (blockIfViewer()) return;
    try { await db.deleteVariacao(id); await reloadVariacoes(); toast("Variação removida", "info"); }
    catch(e){ toast("Erro ao remover variação: "+e.message, "error"); }
  }

  const responsaveis = [...new Set([...profiles.map(p=>p.name), ...records.map(r=>r.responsavel)].filter(Boolean))].sort();

  if (recovery) return (
    <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:`linear-gradient(135deg,#201b18,${T.dark})`,padding:16}}>
      <RecoveryModal onClose={()=>{ setRecovery(false); }}/>
    </div>
  );

  if (!authReady) return (
    <div style={{minHeight:"100vh",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:14,background:T.canvas,color:T.muted,fontFamily:T.font}}>
      <FcamaraLogo size={30}/>
      <div style={{fontSize:13}}>Carregando…</div>
    </div>
  );

  if (!user) return <Login/>;

  if (!dataReady) return (
    <div style={{minHeight:"100vh",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:14,background:T.canvas,color:T.muted,fontFamily:T.font}}>
      <FcamaraLogo size={30}/>
      <div style={{fontSize:13}}>Carregando dados…</div>
    </div>
  );

  return (
    <div style={{fontFamily:T.font,color:T.ink,minHeight:"100vh",background:T.canvas,display:"flex",flexDirection:"column"}}>
      {showImport  && <ImportModal onImport={handleImport} onClose={()=>setImp(false)}/>}
      {showExport  && <ExportModal records={records} onClose={()=>setExp(false)} onDone={(n)=>toast(`CSV exportado — ${n} registros`)}/>}
      {showHistory && <HistoryModal history={history} onClose={()=>setHist(false)} onUndo={(entry)=>handleUndoImport(entry)}/>}
      {confirmLogout && <ConfirmDialog title="Sair da plataforma" message="Deseja realmente encerrar a sessão?" confirmLabel="Sair" onConfirm={()=>{ supabase.auth.signOut(); setUser(null); }} onClose={()=>setCL(false)}/>}

      <Topbar user={user} isAdmin={isAdmin} isMobile={isMobile} onMenu={()=>setDrawer(true)} onLogout={()=>setCL(true)} theme={theme} onToggleTheme={toggleTheme}/>

      {isAdmin&&<div style={{background:T.warnBg,borderBottom:`1px solid ${T.warnLine}`,padding:"7px 20px",fontSize:12,color:T.warn,display:"flex",alignItems:"center",gap:8}}>
        <Badge label="Admin" color="blue" small/> Acesso completo a todos os analistas, empresas e competências.
      </div>}
      {isViewer&&<div style={{background:C.teal.bg,borderBottom:`1px solid ${C.teal.border}`,padding:"7px 20px",fontSize:12,color:C.teal.text,display:"flex",alignItems:"center",gap:8}}>
        <Badge label="Somente visualização" color="teal" small/> Você vê todas as telas e pode extrair relatórios, mas não pode alterar dados.
      </div>}

      {isMobile && <MobileDrawer open={drawer} onClose={()=>setDrawer(false)} page={page} setPage={setPage} user={user} isAdmin={isAdmin} isComercial={isComercial}/>}

      <div style={{display:"flex",flex:1,minHeight:0}}>
        {!isMobile && <Sidebar page={page} setPage={setPage} user={user} isAdmin={isAdmin} isComercial={isComercial}/>}
        <main style={{flex:1,overflowX:"auto",minWidth:0}}>
          {page==="home"&&(
            <div style={{maxWidth:1000,margin:"0 auto",padding:isMobile?"18px 14px":"24px 22px"}}>
              <HomeView user={user} isAdmin={isAdmin} records={records} notes={notes} tasks={tasks} profiles={profiles} fatByRec={fatByRec} varByRec={varByRec} mural={mural} onSaveMural={handleMuralSave} onSaveApelido={handleSaveApelido} onNavigate={setPage}/>
            </div>
          )}
          {(page==="time"||page==="dash")&&(
            <div style={{maxWidth:1140,margin:"0 auto",padding:isMobile?"18px 14px":"24px 22px"}}>
              {page==="time"&&<MyView records={records} clients={clients} analista={user.name} isAdmin={isAdmin} isViewer={isViewer} fatByRec={fatByRec} varByRec={varByRec} varsByRec={varsByRec} aceitaVar={aceitaVar} onAddVariacao={handleAddVariacao} onDelVariacao={handleDelVariacao} onUpdateBulk={handleUpdateBulk} onDeleteRecord={handleRecordDelete} onClearAlert={handleClearAlert} onSaveClass={handleSaveClass} competenciaAtual={state.competenciaAtual} onCompetenciaChange={handleCompetencia}/>}
              {page==="dash"&&<Dashboard records={recordsView} analista={user.name} isAdmin={isAdmin||isComercial} fatByRec={fatByRec} varByRec={varByRec}/>}
            </div>
          )}
          {page==="dados"&&isAdmin&&(
            <div style={{maxWidth:1140,margin:"0 auto",padding:isMobile?"18px 14px":"24px 22px"}}>
              <DataIOView recordsCount={records.length} clientsCount={clients.length}
                onImport={()=>setImp(true)} onExport={()=>setExp(true)} onHistory={()=>setHist(true)}
                onExportClients={()=>{ exportClientsCSV(clients); toast(`Clientes exportados — ${clients.length} cadastro(s)`); }}/>
            </div>
          )}
          {page==="concil"&&(
            <div style={{maxWidth:1280,margin:"0 auto",padding:isMobile?"18px 14px":"24px 22px"}}>
              <ConciliationView records={records} clients={clients} notes={notes} isAdmin={isAdmin} isViewer={isViewer} fatByRec={fatByRec} varByRec={varByRec} faturamentos={faturamentos}
                orfas={notasOrfas.length} onReopenOrphans={handleReopenOrphans}
                onImport={handleNotesImport} onUndoImport={handleNotesUndo} onDeleteNote={handleNoteDelete}
                onConciliate={handleConciliate} onReopen={handleReopenGroup}/>
            </div>
          )}
          {page==="reports"&&(
            <div style={{maxWidth:1140,margin:"0 auto",padding:isMobile?"18px 14px":"24px 22px"}}>
              <ReportsView records={records} clients={clients} notes={notes} faturamentos={faturamentos} variacoes={variacoes} varByRec={varByRec} fatByRec={fatByRec} isAdmin={isAdmin} analistas={responsaveis}/>
            </div>
          )}
          {page==="valida"&&(
            <div style={{maxWidth:1000,margin:"0 auto",padding:isMobile?"18px 14px":"24px 22px"}}>
              <ValidatorsView records={records} notes={notes} faturamentos={faturamentos} fatByRec={fatByRec} varByRec={varByRec}/>
            </div>
          )}
          {page==="projeto"&&(
            <div style={{maxWidth:1240,margin:"0 auto",padding:isMobile?"18px 14px":"24px 22px"}}>
              <ProjectTimelineView records={recordsView} clients={clients} fatByRec={fatByRec} varByRec={varByRec}/>
            </div>
          )}
          {page==="report"&&isComercial&&(
            <div style={{maxWidth:1180,margin:"0 auto",padding:isMobile?"18px 14px":"24px 22px"}}>
              <WeeklyReportView records={recordsView} clients={clients} bu={userBu} nome={user.name} fatByRec={fatByRec} varByRec={varByRec}/>
            </div>
          )}
          {page==="report"&&isAdmin&&!isComercial&&(
            <div style={{maxWidth:1180,margin:"0 auto",padding:isMobile?"18px 14px":"24px 22px"}}>
              <WeeklyReportView records={recordsAtivos} clients={clients} nome={user.name} fatByRec={fatByRec} varByRec={varByRec} canPickBu/>
            </div>
          )}
          {page==="represados"&&(
            <div style={{maxWidth:1240,margin:"0 auto",padding:isMobile?"18px 14px":"24px 22px"}}>
              <RepresadosView records={recordsView} clients={clients} fatByRec={fatByRec} varByRec={varByRec} onSaveClass={handleSaveClass} isViewer={isViewer||isComercial}/>
            </div>
          )}
          {page==="clients"&&(
            <div style={{maxWidth:1140,margin:"0 auto",padding:isMobile?"18px 14px":"24px 22px"}}>
              <ClientsView clients={clients} isAdmin={isAdmin} isViewer={isViewer} onSave={handleClientSave} onDelete={handleClientDelete} onBulkImport={handleClientsImport} onMerge={handleClientsMerge}/>
            </div>
          )}
          {page==="tasks"&&(
            <div style={{padding:isMobile?"18px 14px":"24px 22px"}}>
              <Kanban tasks={tasks} responsaveis={responsaveis} isAdmin={isAdmin} isViewer={isViewer} competenciaAtual={state.competenciaAtual}
                templates={templates} deliveries={deliveries}
                onAdd={handleTaskAdd} onUpdate={handleTaskUpdate} onDelete={handleTaskDelete}
                onTemplateSave={handleTemplateSave} onTemplateDelete={handleTemplateDelete} onGenerate={handleGenerateDelivery}/>
            </div>
          )}
          {page==="bu"&&isAdmin&&(
            <div style={{maxWidth:1100,margin:"0 auto",padding:isMobile?"18px 14px":"24px 22px"}}>
              <BuClassifierView records={records} onSetBu={handleSetBu}/>
            </div>
          )}
          {page==="aliases"&&isAdmin&&(
            <div style={{maxWidth:1000,margin:"0 auto",padding:isMobile?"18px 14px":"24px 22px"}}>
              <AliasesView aliases={aliases} records={records} onSave={handleAliasSave} onDelete={handleAliasDelete} isViewer={isViewer}/>
            </div>
          )}
          {page==="comercial"&&isAdmin&&(
            <div style={{maxWidth:1240,margin:"0 auto",padding:isMobile?"18px 14px":"24px 22px"}}>
              <ComercialView records={recordsAtivos} clients={clients} fatByRec={fatByRec} varByRec={varByRec}/>
            </div>
          )}
          {page==="previsao"&&isAdmin&&(
            <div style={{maxWidth:1240,margin:"0 auto",padding:isMobile?"18px 14px":"24px 22px"}}>
              <ForecastView records={recordsAtivos} varByRec={varByRec}/>
            </div>
          )}
          {page==="correcoes"&&isAdmin&&(
            <div style={{maxWidth:1240,margin:"0 auto",padding:isMobile?"18px 14px":"24px 22px"}}>
              <CorrectionsView records={records} fatByRec={fatByRec}
                onEdit={handleCorrEdit} onDelete={handleCorrDelete} onMerge={handleCorrDelete} onInsert={handleCorrInsert}
                lastCorr={lastCorr} onUndo={handleCorrUndo}/>
            </div>
          )}
          {page==="access"&&isAdmin&&(
            <div style={{maxWidth:1140,margin:"0 auto",padding:isMobile?"18px 14px":"24px 22px"}}>
              <AccessManagement profiles={profiles} currentUserId={user.id} onUpdate={handleProfileUpdate} onRemove={handleProfileRemove} onRefresh={reloadProfiles}/>
            </div>
          )}
        </main>
      </div>
    </div>
  );
}

export default function App() {
  return (
    <>
      <GlobalStyles/>
      <ToastProvider>
        <AppInner/>
      </ToastProvider>
    </>
  );
}

