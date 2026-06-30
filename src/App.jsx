import { useState, useEffect, useRef, useCallback, createContext, useContext } from "react";
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

const STEPS = [
  { id: "p1_extrair",      group: 1, label: "P1",  name: "Extrair dados (FC Team)",  type: "check" },
  { id: "p2_racional",     group: 2, label: "P2",  name: "Montar Racional",          type: "check" },
  { id: "p3_envio_com",    group: 3, label: "P3a", name: "Envio ao Comercial",       type: "check" },
  { id: "p3_retorno_com",  group: 3, label: "P3b", name: "Retorno do Comercial",     type: "check" },
  { id: "p3_data_retorno", group: 3, label: "P3c", name: "Data Retorno",             type: "date"  },
  { id: "p4_envio_cli",    group: 4, label: "P4a", name: "Envio ao Cliente",         type: "check" },
  { id: "p4_aprovacao",    group: 4, label: "P4b", name: "Aprovação do Cliente",     type: "check" },
  { id: "p4_data_aprov",   group: 4, label: "P4c", name: "Data Aprovação",           type: "date"  },
  { id: "p5_nf",           group: 5, label: "P5a", name: "NF Emitida?",              type: "check" },
  { id: "p5_data_nf",      group: 5, label: "P5b", name: "Data Emissão NF",          type: "date"  },
  { id: "p5_no_corte",     group: 5, label: "P5c", name: "Dentro do Corte?",         type: "check" },
];

const STEP_GROUPS = [
  { num: 1, title: "Extração de dados",    short: "Extração",  steps: ["p1_extrair"] },
  { num: 2, title: "Racional",             short: "Racional",  steps: ["p2_racional"] },
  { num: 3, title: "Validação comercial",  short: "Comercial", steps: ["p3_envio_com","p3_retorno_com","p3_data_retorno"] },
  { num: 4, title: "Aprovação do cliente", short: "Cliente",   steps: ["p4_envio_cli","p4_aprovacao","p4_data_aprov"] },
  { num: 5, title: "Faturamento",          short: "NF",        steps: ["p5_nf","p5_data_nf","p5_no_corte"] },
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
  5: ["p5_nf"],
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
function fmtPepsCSV(json)      { const m = parseJSON(json, {}); if (!m || typeof m !== "object") return ""; return Object.entries(m).map(([k, v]) => `${k}: ${(Array.isArray(v) ? v : []).join(", ")}`).join(" | "); }
function fmtPropostasCSV(json) { const a = parseJSON(json, []); return Array.isArray(a) ? a.join(" ; ") : ""; }
function fmtCalendarioCSV(json){ const a = parseJSON(json, []); return Array.isArray(a) ? a.map((p, i) => `${i+1}) ${p.quando || ""} - ${p.etapa || ""}: ${p.oQueFazer || ""}`).join(" | ") : ""; }
function exportClientsCSV(clients) {
  const headers = ["Nome","Cód SAP","Grupo de empresa","Tipos de contrato","PEPs","Propostas","Período faturamento","Calendário (passo a passo)","Tem portal","Classificação portal","Link portal","Usuário portal","Senha portal","Passo a passo portal","Prazo vencimento","Forma de pagamento","Contato financeiro (nome)","Contato financeiro (e-mail)","Account manager (nome)","Account manager (e-mail)"];
  const rows = clients.map(c => [c.nome, c.codSap, c.grupoEmpresa, c.tiposContrato, fmtPepsCSV(c.tiposPeps), fmtPropostasCSV(c.propostas), c.periodoFaturamento, fmtCalendarioCSV(c.calendario), c.temPortal ? "Sim" : "Não", c.portalTipo, c.portalLink, c.portalUsuario, c.portalSenha, c.portalPassoUrl, c.prazoVencimento, c.formaPagamento, c.contatoFinanceiro, c.contatoFinanceiroEmail, c.accountManager, c.accountManagerEmail]);
  downloadCSV(`FCamara_Clientes_${clients.length}.csv`, headers, rows);
}

const PORTAL_TIPOS = ["Inclusão de notas", "Medição de serviços"];

function calcStatus(prog) {
  if (!prog) return "Não iniciado";
  if (prog.p5_no_corte)    return "Faturado no corte";
  if (prog.p5_nf)          return "NF emitida";
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
  if (prog.p5_no_corte)    return "green";
  if (prog.p5_nf)          return "teal";
  if (prog.p4_aprovacao)   return "blue";
  if (prog.p4_envio_cli || prog.p3_retorno_com || prog.p3_envio_com) return "yellow";
  if (prog.p1_extrair || prog.p2_racional) return "orange";
  return "gray";
}

const STATUS_ORDER = ["Não iniciado","Dados extraídos","Racional montado","Aguard. retorno comercial","Retorno comercial recebido","Aguard. aprovação cliente","Cliente aprovou","NF emitida","Faturado no corte"];

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

const T = {
  brand:      "#1d4ed8",
  brandDark:  "#1e3a8a",
  brandBg:    "#eff6ff",
  ink:     "#111827", // títulos / texto principal
  inkSoft: "#374151", // texto secundário
  muted:   "#5b6472", // texto terciário (era #9ca3af — não passava contraste AA)
  faint:   "#9ca3af", // apenas decorativo (ícones, divisores)
  surface: "#ffffff",
  canvas:  "#f6f8fb",
  line:    "#e5e7eb",
  lineSoft:"#f1f3f5",
  ok:      "#15803d", okBg:"#f0fdf4", okLine:"#86efac",
  warn:    "#92400e", warnBg:"#fffbeb", warnLine:"#fde68a",
  danger:  "#b91c1c", dangerBg:"#fef2f2", dangerLine:"#fca5a5",
  rSm:6, rMd:8, rLg:10, rXl:12, rPill:999,
  shSm:"0 1px 2px rgba(16,24,40,.05)",
  shMd:"0 4px 14px rgba(16,24,40,.08)",
  shLg:"0 16px 48px rgba(16,24,40,.20)",
};

// Escala tipográfica
const Ty = {
  h1:    { fontSize:18, fontWeight:800, color:T.ink, margin:0, letterSpacing:"-.01em" },
  h2:    { fontSize:15, fontWeight:700, color:T.ink, margin:0 },
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

const inp = { padding:"8px 11px", borderRadius:T.rMd, border:`1px solid ${T.line}`, fontSize:13, fontFamily:"inherit", background:"#fff", color:T.ink, width:"100%", boxSizing:"border-box", outline:"none" };

// ─── GLOBAL STYLES (foco, hover, animações, scrollbar) ───────────────────────

const GLOBAL_CSS = `
  *{box-sizing:border-box}
  body{margin:0}
  button{font-family:inherit}
  :focus-visible{outline:2px solid ${T.brand};outline-offset:2px;border-radius:6px}
  input:focus,select:focus,textarea:focus{border-color:${T.brand};box-shadow:0 0 0 3px rgba(29,78,216,.12)}
  .fc-btn{transition:filter .12s,box-shadow .12s,background .12s}
  .fc-btn:hover:not(:disabled){filter:brightness(.96)}
  .fc-row:hover{background:#eff6ff80}
  .fc-card-int{transition:box-shadow .15s,border-color .15s}
  .fc-card-int:hover{box-shadow:${T.shMd};border-color:#cdd5e0}
  .fc-scroll::-webkit-scrollbar{height:9px;width:9px}
  .fc-scroll::-webkit-scrollbar-thumb{background:#cbd2dc;border-radius:9px}
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
  const tc = { ok:{bar:T.ok,ic:"✓"}, error:{bar:T.danger,ic:"✕"}, info:{bar:T.brand,ic:"ℹ"} };
  return (
    <ToastCtx.Provider value={push}>
      {children}
      <div style={{ position:"fixed", bottom:18, right:18, zIndex:500, display:"flex", flexDirection:"column", gap:8, maxWidth:"calc(100vw - 36px)" }} role="status" aria-live="polite">
        {toasts.map(t => {
          const s = tc[t.type] || tc.info;
          return (
            <div key={t.id} className="fc-toast" style={{ display:"flex", alignItems:"center", gap:10, background:"#fff", borderLeft:`4px solid ${s.bar}`, boxShadow:T.shMd, borderRadius:T.rMd, padding:"11px 16px", minWidth:240, fontSize:13, color:T.ink }}>
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

function Btn({ children, onClick, primary, danger, ghost, small, disabled, title, style:s={} }) {
  const base = { padding:small?"6px 12px":"9px 18px", borderRadius:T.rMd, fontSize:small?12:13, fontWeight:600, cursor:disabled?"not-allowed":"pointer", display:"inline-flex", alignItems:"center", justifyContent:"center", gap:6, border:"none", opacity:disabled?.5:1, ...s };
  const v = primary ? { background:T.brand, color:"#fff" }
          : danger  ? { background:T.danger, color:"#fff" }
          : ghost   ? { background:"transparent", color:T.inkSoft }
          :           { background:"#fff", color:T.inkSoft, border:`1px solid ${T.line}` };
  return <button className="fc-btn" title={title} aria-label={title} onClick={disabled?undefined:onClick} disabled={disabled} style={{...base,...v}}>{children}</button>;
}

function Avatar({ name, size=30, admin }) {
  return (
    <span style={{ width:size, height:size, borderRadius:"50%", background:admin?T.brandDark:T.brand, color:"#fff", display:"inline-flex", alignItems:"center", justifyContent:"center", fontSize:size*0.38, fontWeight:700, flexShrink:0 }} aria-hidden="true">
      {initials(name)}
    </span>
  );
}

// Marca Fcamara — "F" branca sobre quadrado laranja arredondado (SVG, sem assets externos)
const FC_ORANGE = "#F1572C";
function FcamaraLogo({ size = 40 }) {
  return (
    <svg width={size} height={size} viewBox="0 0 64 64" role="img" aria-label="Logo Fcamara" style={{ display:"block", flexShrink:0 }}>
      <rect width="64" height="64" rx="16" fill={FC_ORANGE} />
      {/* haste vertical */}
      <rect x="19" y="15" width="11" height="35" rx="5.5" fill="#fff" />
      {/* braço superior, canto direito arredondado (folha) */}
      <path d="M25 15 h12 a9 9 0 0 1 0 18 H25 Z" fill="#fff" />
      {/* braço central */}
      <rect x="25" y="33" width="18" height="10.5" rx="5.25" fill="#fff" />
    </svg>
  );
}

function Card({ children, style:s={}, interactive, ...rest }) {
  return <div className={interactive?"fc-card-int":undefined} style={{ background:"#fff", border:`1px solid ${T.line}`, borderRadius:T.rXl, ...s }} {...rest}>{children}</div>;
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
  const colorFor  = (st) => st==="done" ? T.ok : st==="partial" ? C.blue.solid : "#fff";
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
      <div ref={ref} role="dialog" aria-modal="true" aria-label={title} className="fc-scroll" style={{ background:"#fff", borderRadius:T.rXl+2, padding:"22px 26px", width:w, maxWidth:"100%", maxHeight:"92vh", overflowY:"auto", boxShadow:T.shLg, animation:"fcModalIn .18s ease" }} onClick={e=>e.stopPropagation()}>
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
  const missing = Object.keys(TE_COL_MAP).filter(k=>colIdx[k]==null);
  if (missing.length>4) return { records:[], errors:[`Cabeçalhos não encontrados: ${missing.join(", ")}. Use a aba "📥 Time & Expenses".`] };
  const records=[]; const skipped=[];
  for (let i=hi+1;i<rows.length;i++) {
    const row=rows[i];
    if(!row||row.every(c=>c==null||c==="")) continue;
    const get=k=>colIdx[k]!=null?(row[colIdx[k]]??""):"";
    const getNum=k=>parseFloat(String(get(k)).replace(",","."))||0;
    const getStr=k=>String(get(k)).trim();
    const cliente=getStr("cliente"), pep=getStr("pep"), responsavel=getStr("responsavel");
    if(!cliente||!pep||!responsavel){skipped.push(i+1);continue;}
    records.push({ id:genId(), responsavel, empresa, tipo, codCliente:getStr("codCliente"), cliente, pep, inicio:excelDateToStr(get("inicio")), fim:excelDateToStr(get("fim")), profissional:getStr("profissional"), valorVenda:getNum("valorVenda"), hrsAprovadas:getNum("hrsAprovadas"), valorTotal:getNum("valorTotal"), valorLiquido:getNum("valorLiquido"), competencia, progress:makeProgress(), nfNumero:"", obs:"", updatedAt:nowISO() });
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
  const [mode,setMode]=useState("add");
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
        <span aria-hidden="true">🔒</span><span>Apenas administradores podem importar dados.</span>
      </div>

      {/* Layout da planilha */}
      <div style={{marginBottom:14}}>
        <label style={Ty.label}>Layout da planilha</label>
        <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
          {[{v:"te",l:"Time & Expenses",d:"Por profissional, com horas e valores."},{v:"feewip",l:"Fee / WIP",d:"Por PEP, receita planejada. Sobe Fee e WIP juntos."},{v:"usage",l:"Usage Based",d:"Por PEP, receita planejada. Tipo Usage Based."}].map(opt=>(
            <label key={opt.v} style={{flex:"1 1 180px",display:"flex",gap:8,padding:"10px 12px",borderRadius:T.rMd,border:`2px solid ${layout===opt.v?T.brand:T.line}`,cursor:"pointer",background:layout===opt.v?T.brandBg:"#fff"}}>
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
          {[{v:"add",l:"➕ Incluir novos",d:"Adiciona sem apagar registros existentes."},{v:"replace",l:"🔄 Substituir",d:"Remove e reimporta SOMENTE a competência (mês) + empresa + tipo informados. Outros meses não são afetados."}].map(opt=>(
            <label key={opt.v} style={{flex:"1 1 200px",display:"flex",gap:8,padding:"10px 12px",borderRadius:T.rMd,border:`2px solid ${mode===opt.v?T.brand:T.line}`,cursor:"pointer",background:mode===opt.v?T.brandBg:"#fff"}}>
              <input type="radio" name="mode" value={opt.v} checked={mode===opt.v} onChange={()=>setMode(opt.v)} style={{marginTop:2}}/>
              <div><div style={{fontSize:13,fontWeight:700,color:mode===opt.v?T.brand:T.inkSoft}}>{opt.l}</div><div style={{fontSize:11,color:T.muted,marginTop:2}}>{opt.d}</div></div>
            </label>
          ))}
        </div>
        {mode==="replace"&&<div style={{marginTop:8,fontSize:12,color:T.danger,fontWeight:600}}>⚠ Apenas os registros desta competência (mês), empresa e tipo serão substituídos. O progresso já registrado para esse recorte será perdido; os demais meses permanecem intactos.</div>}
      </div>
      <div style={{marginBottom:14}}><Field label="Nota da importação (opcional)"><input style={inp} placeholder="Ex: Ajuste de valores de maio" value={note} onChange={e=>setNote(e.target.value)}/></Field></div>
      <input type="file" ref={fileRef} style={{display:"none"}} accept=".xlsx,.xlsm,.xls" onChange={e=>{if(e.target.files[0])readFile(e.target.files[0]);e.target.value="";}}/>
      <div onDragOver={e=>{e.preventDefault();setDragOver(true);}} onDragLeave={()=>setDragOver(false)} onDrop={onDrop} onClick={()=>fileRef.current.click()} role="button" tabIndex={0} aria-label="Carregar arquivo"
        style={{border:`2px dashed ${dragOver?T.brand:fileName?T.okLine:"#cbd2dc"}`,borderRadius:T.rLg,padding:"28px 20px",textAlign:"center",cursor:"pointer",background:dragOver?T.brandBg:fileName?T.okBg:"#fafbfc",marginBottom:14}}>
        {loading?<div style={{color:T.muted,fontSize:13}}>⏳ Lendo arquivo...</div>:fileName?<><div style={{fontSize:24,marginBottom:6}}>✅</div><div style={{fontSize:13,fontWeight:700,color:T.ok}}>{fileName}</div><div style={{fontSize:11,color:T.muted,marginTop:4}}>Clique para trocar</div></>:<><div style={{fontSize:28,marginBottom:8}}>📂</div><div style={{fontSize:14,fontWeight:600,color:T.inkSoft}}>Clique ou arraste o arquivo aqui</div><div style={{fontSize:12,color:T.muted,marginTop:4}}>Aceita .xlsm e .xlsx</div></>}
      </div>
      {msgs.map((m,i)=><div key={i} style={{marginBottom:6,fontSize:12,padding:"8px 12px",borderRadius:T.rMd,background:mc[m.type].bg,color:mc[m.type].text,border:`1px solid ${mc[m.type].border}`}}>{m.text}</div>)}
      {preview&&<div style={{marginBottom:14,padding:"12px 14px",borderRadius:T.rMd,background:T.okBg,border:`1px solid ${T.okLine}`}}>
        <div style={{fontSize:13,fontWeight:700,color:T.ok,marginBottom:8}}>✓ {preview.length} registros prontos</div>
        <div className="fc-scroll" style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}><thead><tr style={{background:"#dcfce7"}}>{["Responsável","Cliente","PEP","Profissional","Val. Total"].map(h=><th key={h} style={{padding:"4px 8px",textAlign:"left",color:T.ok,fontWeight:700}}>{h}</th>)}</tr></thead><tbody>{preview.slice(0,5).map(r=><tr key={r.id}><td style={{padding:"4px 8px"}}>{r.responsavel}</td><td style={{padding:"4px 8px"}}>{r.cliente}</td><td style={{padding:"4px 8px",fontFamily:"monospace"}}>{r.pep}</td><td style={{padding:"4px 8px"}}>{r.profissional}</td><td style={{padding:"4px 8px"}}>{fmtShort(r.valorTotal)}</td></tr>)}{preview.length>5&&<tr><td colSpan={5} style={{padding:"4px 8px",color:T.muted,fontStyle:"italic"}}>... e mais {preview.length-5}</td></tr>}</tbody></table></div>
      </div>}
      <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
        <Btn onClick={onClose}>Cancelar</Btn>
        <Btn primary onClick={()=>{if(!preview)return;onImport({records:preview,competencia:comp,empresa,tipo,mode,note:note||(mode==="replace"?"Substituição":"Adição")});onClose();}} disabled={!preview}>{mode==="replace"?"⚠ Confirmar substituição":"✓ Confirmar importação"}</Btn>
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
      <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}><Btn onClick={onClose}>Cancelar</Btn><Btn primary onClick={doExport}>⬇ Exportar CSV</Btn></div>
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
  const [error, setError] = useState("");

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
    const updated = records.map(r => selected.has(r.id)
      ? { ...r, progress: { ...sharedProg, p5_nf: r.progress?.p5_nf || false, p5_data_nf: r.progress?.p5_data_nf || "" }, obs: obs || r.obs, updatedAt: now }
      : r
    );
    onSave(updated);
    onClose();
  }

  return (
    <Modal title={`Atualizar passos — ${cliente}`} subtitle={`${pep} · ${records.length} profissionais`} onClose={onClose} extraWide>
      {/* Pré-visualização do funil compartilhado */}
      <div style={{ background:T.canvas, border:`1px solid ${T.line}`, borderRadius:T.rLg, padding:"14px 16px", marginBottom:18 }}>
        <PipelineStepper states={recordStates(sharedProg, tipo)} groups={grupos} showLabels/>
      </div>

      {/* Seleção de profissionais */}
      <div style={{marginBottom:20}}>
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:8,flexWrap:"wrap"}}>
          <span style={{fontSize:13,fontWeight:700,color:T.ink}}>Aplicar a:</span>
          <button onClick={toggleAll} style={{fontSize:11,color:T.brand,background:"none",border:"none",cursor:"pointer",textDecoration:"underline"}}>
            {selected.size===records.length?"Desmarcar todos":"Selecionar todos"}
          </button>
        </div>
        <div style={{display:"flex",flexWrap:"wrap",gap:8}}>
          {records.map(r=>(
            <label key={r.id} style={{display:"flex",alignItems:"center",gap:6,padding:"6px 12px",borderRadius:T.rMd,border:`1.5px solid ${selected.has(r.id)?T.brand:T.line}`,background:selected.has(r.id)?T.brandBg:"#fafbfc",cursor:"pointer",fontSize:13}}>
              <input type="checkbox" checked={selected.has(r.id)} onChange={()=>toggle(r.id)} style={{width:14,height:14}}/>
              <span style={{fontWeight:selected.has(r.id)?600:400,color:selected.has(r.id)?T.brand:T.inkSoft}}>{r.profissional}</span>
              <Badge label={calcStatus(r.progress)} color={calcStatusColor(r.progress)} small dot/>
            </label>
          ))}
        </div>
        {selected.size>0&&<div style={{fontSize:11,color:T.muted,marginTop:6}}>{selected.size} profissional(is) selecionado(s) — os passos abaixo serão aplicados a todos eles.</div>}
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
              // A NF (emitida e data de emissão) é tratada na tela "Notas fiscais".
              if (s.id === "p5_data_nf") return null;
              if (s.id === "p5_nf") return (
                <div key={s.id} style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:8,gap:8}}>
                  <span style={{fontSize:12,color:T.inkSoft,flex:1}}>{s.name}</span>
                  <button type="button" onClick={onOpenNF}
                    style={{display:"inline-flex",alignItems:"center",gap:5,background:T.brandBg,border:`1px solid ${C.blue.border}`,color:T.brand,borderRadius:T.rMd,padding:"5px 10px",fontSize:12,fontWeight:600,cursor:"pointer",whiteSpace:"nowrap"}}>🧾 Atribuir nota</button>
                </div>
              );
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
                <div key={g.nf} style={{border:`1px solid ${T.line}`,borderRadius:T.rLg,padding:"12px 14px",background:"#fff"}}>
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
        {groups.length>0 && semNf.length>0 && <div style={{fontSize:11,color:T.warn,marginTop:8}}>⚠ {semNf.length} profissional(is) ainda sem NF.</div>}
      </div>

      {/* Montar nova NF */}
      <div style={{background:T.canvas,border:`1px solid ${T.line}`,borderRadius:T.rLg,padding:"14px 16px"}}>
        <div style={{fontSize:13,fontWeight:700,color:T.ink,marginBottom:10}}>Atribuir NF a profissionais</div>

        <div style={{display:"flex",flexWrap:"wrap",gap:8,marginBottom:12}}>
          {localRecs.map(r=>{
            const on = selected.has(r.id);
            const cur = (r.nfNumero||"").trim();
            return (
              <label key={r.id} style={{display:"flex",alignItems:"center",gap:8,padding:"8px 12px",borderRadius:T.rMd,border:`1.5px solid ${on?T.brand:T.line}`,background:on?T.brandBg:"#fff",cursor:"pointer",fontSize:13}}>
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
          <Btn primary onClick={assign} disabled={selected.size===0}>🧾 Atribuir NF aos selecionados</Btn>
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
function RecordEditModal({ record, onSave, onClose }) {
  const [f, setF] = useState(record);
  const set = (k,v) => setF(p=>({...p,[k]:v}));
  const num = (k,v) => setF(p=>({...p,[k]: v===""?0:parseFloat(String(v).replace(",","."))||0}));
  function save() { onSave({ ...record, ...f, updatedAt: nowISO() }); onClose(); }
  return (
    <Modal title="Editar registro" subtitle={`${record.cliente} · ${record.profissional}`} onClose={onClose} wide>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(180px,1fr))",gap:12,marginBottom:16}}>
        <Field label="Responsável"><input style={inp} value={f.responsavel||""} onChange={e=>set("responsavel",e.target.value)}/></Field>
        <Field label="Empresa"><select style={inp} value={f.empresa||""} onChange={e=>set("empresa",e.target.value)}>{EMPRESAS.map(e=><option key={e.cod} value={e.cod}>{e.cod} — {e.nome}</option>)}</select></Field>
        <Field label="Tipo"><select style={inp} value={f.tipo||""} onChange={e=>set("tipo",e.target.value)}>{TIPOS_PROJETO.map(t=><option key={t}>{t}</option>)}</select></Field>
        <Field label="Competência"><input style={inp} value={f.competencia||""} onChange={e=>set("competencia",e.target.value)}/></Field>
        <Field label="Cód. Cliente"><input style={inp} value={f.codCliente||""} onChange={e=>set("codCliente",e.target.value)}/></Field>
        <Field label="Cliente"><input style={inp} value={f.cliente||""} onChange={e=>set("cliente",e.target.value)}/></Field>
        <Field label="PEP"><input style={inp} value={f.pep||""} onChange={e=>set("pep",e.target.value)}/></Field>
        <Field label="Profissional"><input style={inp} value={f.profissional||""} onChange={e=>set("profissional",e.target.value)}/></Field>
        <Field label="Início"><input style={inp} value={f.inicio||""} onChange={e=>set("inicio",e.target.value)}/></Field>
        <Field label="Fim"><input style={inp} value={f.fim||""} onChange={e=>set("fim",e.target.value)}/></Field>
        <Field label="Valor de venda"><input style={inp} value={f.valorVenda} onChange={e=>num("valorVenda",e.target.value)}/></Field>
        <Field label="Hrs aprovadas"><input style={inp} value={f.hrsAprovadas} onChange={e=>num("hrsAprovadas",e.target.value)}/></Field>
        <Field label="Valor total"><input style={inp} value={f.valorTotal} onChange={e=>num("valorTotal",e.target.value)}/></Field>
        <Field label="Valor líquido"><input style={inp} value={f.valorLiquido} onChange={e=>num("valorLiquido",e.target.value)}/></Field>
      </div>
      <div style={{marginBottom:16}}><Field label="Observações"><textarea style={{...inp,minHeight:54,resize:"vertical"}} value={f.obs||""} onChange={e=>set("obs",e.target.value)}/></Field></div>
      <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
        <Btn onClick={onClose}>Cancelar</Btn>
        <Btn primary onClick={save}>Salvar registro</Btn>
      </div>
    </Modal>
  );
}

// ─── MY VIEW (team) ──────────────────────────────────────────────────────────

function MyView({ records, analista, isAdmin, onUpdateBulk, onDeleteRecord, competenciaAtual, onCompetenciaChange }) {
  const isMobile = useIsMobile();
  const [recordEdit, setRecEdit] = useState(null);
  const [recordDel, setRecDel]   = useState(null);
  const [empresa, setEmpresa]       = useState("");
  const [tipo, setTipo]             = useState("");
  const [filterComp, setFilterComp] = useState(competenciaAtual);
  const [filterAnalista, setFA]     = useState("todos");
  const [filterEtapa, setFEt]       = useState("todas");
  const [searchCliente, setSC]      = useState("");
  const [searchProf, setSP]         = useState("");
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
  if (isAdmin && filterAnalista!=="todos") filtered = filtered.filter(r=>r.responsavel===filterAnalista);
  if (filterEtapa!=="todas") filtered = filtered.filter(r=>calcStatus(r.progress)===filterEtapa);
  if (searchCliente) filtered = filtered.filter(r=>r.cliente.toLowerCase().includes(searchCliente.toLowerCase()));
  if (searchProf)    filtered = filtered.filter(r=>r.profissional.toLowerCase().includes(searchProf.toLowerCase()));

  // Resumo por tipo de contrato
  const porTipo = {};
  filtered.forEach(r=>{ const t=r.tipo||"—"; if(!porTipo[t]) porTipo[t]={count:0,total:0,fat:0}; porTipo[t].count++; porTipo[t].total+=(r.valorTotal||0); if(r.progress?.p5_nf) porTipo[t].fat+=(r.valorTotal||0); });
  const tipoColors = { "Time & Expenses":"blue", "Fee":"purple", "WIP":"teal", "Usage Based":"orange" };

  const grouped = {};
  filtered.forEach(r=>{
    const key = r.cliente+"|"+r.pep;
    if(!grouped[key]) grouped[key]={ cliente:r.cliente, pep:r.pep, records:[] };
    grouped[key].records.push(r);
  });
  const groups = Object.values(grouped);
  const selW = isMobile ? "100%" : "auto";

  return (
    <div>
      {bulkTarget&&<BulkTimelineModal {...bulkTarget} onClose={()=>setBulk(null)} onOpenNF={()=>{ const t=bulkTarget; setBulk(null); setNf(t); }} onSave={updated=>{onUpdateBulk(updated);setBulk(null);}}/>}
      {nfTarget&&<NFGroupModal {...nfTarget} onClose={()=>setNf(null)} onSave={updated=>{onUpdateBulk(updated);setNf(null);}}/>}
      {recordEdit&&<RecordEditModal record={recordEdit} onClose={()=>setRecEdit(null)} onSave={r=>{onUpdateBulk([r]);setRecEdit(null);}}/>}
      {recordDel&&<ConfirmDialog title="Excluir registro" danger confirmLabel="Excluir"
        message={`Excluir o registro de "${recordDel.profissional}" (${recordDel.cliente})? Esta ação não pode ser desfeita.`}
        onConfirm={()=>onDeleteRecord(recordDel.id)} onClose={()=>setRecDel(null)}/>}

      <div style={{display:"flex",alignItems:"baseline",gap:10,marginBottom:14,flexWrap:"wrap"}}>
        <h1 style={Ty.h1}>📋 Minha visão</h1>
        <span style={Ty.small}>{groups.length} cliente(s) · {filtered.length} registro(s)</span>
      </div>

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
          {isAdmin&&<select style={{...inp,width:selW}} value={filterAnalista} onChange={e=>setFA(e.target.value)} aria-label="Analista">
            <option value="todos">Todos os analistas</option>
            {analistas.map(a=><option key={a}>{a}</option>)}
          </select>}
          <select style={{...inp,width:selW,minWidth:160}} value={filterEtapa} onChange={e=>setFEt(e.target.value)} aria-label="Etapa do funil">
            <option value="todas">Todas as etapas</option>
            {STATUS_ORDER.map(s=><option key={s}>{s}</option>)}
          </select>
          <input style={{...inp,width:isMobile?"100%":160}} placeholder="🔎 Cliente..." value={searchCliente} onChange={e=>setSC(e.target.value)}/>
          <input style={{...inp,width:isMobile?"100%":160}} placeholder="🔎 Profissional..." value={searchProf} onChange={e=>setSP(e.target.value)}/>
        </div>
      </Card>

      {/* Legenda do funil */}
      {groups.length>0 && <div style={{display:"flex",alignItems:"center",gap:14,flexWrap:"wrap",margin:"0 2px 12px",fontSize:11,color:T.muted}}>
        <span style={{fontWeight:600}}>Funil:</span>
        <span style={{display:"flex",alignItems:"center",gap:5}}><i style={{width:11,height:11,borderRadius:"50%",background:T.ok}}/> concluído</span>
        <span style={{display:"flex",alignItems:"center",gap:5}}><i style={{width:11,height:11,borderRadius:"50%",background:C.blue.solid}}/> parcial</span>
        <span style={{display:"flex",alignItems:"center",gap:5}}><i style={{width:11,height:11,borderRadius:"50%",background:"#fff",border:"2px solid #cbd2dc"}}/> pendente</span>
        <span style={{color:T.faint}}>· {STEP_GROUPS.map(g=>g.num+" "+g.short).join("  ·  ")}</span>
      </div>}

      {groups.length===0&&<Card style={{textAlign:"center",padding:"3rem"}}>
        <div style={{fontSize:32,marginBottom:10}}>📭</div>
        <div style={{fontSize:14,color:T.muted}}>Nenhum registro encontrado para os filtros selecionados.</div>
      </Card>}

      {/* Cards de clientes */}
      {groups.map(g=>{
        const total   = g.records.reduce((a,r)=>a+(r.valorTotal||0),0);
        const faturados = g.records.filter(r=>r.progress?.p5_nf).length;
        const pct     = Math.round((faturados/g.records.length)*100)||0;
        const isOpen  = expandedCliente===(g.cliente+g.pep);
        const overallStatus = g.records.every(r=>r.progress?.p5_no_corte)?"Faturado no corte":g.records.every(r=>r.progress?.p5_nf)?"NF emitida":g.records.some(r=>r.progress?.p5_nf)?"Faturado parcialmente":"Em andamento";
        const overallColor  = g.records.every(r=>r.progress?.p5_no_corte)?"green":g.records.every(r=>r.progress?.p5_nf)?"teal":g.records.some(r=>r.progress?.p5_nf)?"orange":"yellow";
        const gtipo = g.records[0]?.tipo;
        const ggrupos = funnelGroups(gtipo);
        const agg = aggregateStates(g.records, gtipo);
        const temProf = g.records.some(r=>r.profissional);

        return (
          <Card key={g.cliente+g.pep} interactive style={{marginBottom:10,overflow:"hidden"}}>
            {/* Cabeçalho do cliente — clicável */}
            <div onClick={()=>setExp(isOpen?null:(g.cliente+g.pep))} style={{display:"flex",alignItems:"center",gap:14,padding:"14px 18px",cursor:"pointer",userSelect:"none",flexWrap:"wrap"}}>
              <div style={{flex:1,minWidth:200}}>
                <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:5}}>
                  <span style={{fontSize:14,fontWeight:700,color:T.ink}}>🏦 {g.cliente}</span>
                  <Badge label={overallStatus} color={overallColor} small dot/>
                </div>
                <div style={{fontSize:11,color:T.muted}}>{g.pep} · {gtipo} · {g.records.length} {temProf?"profissionais":"registro(s)"} · {fmtShort(total)}</div>
              </div>
              {!isMobile && <div style={{ width:230 }}><PipelineStepper states={agg} groups={ggrupos} showLabels/></div>}
              <div style={{textAlign:"center"}}>
                <div style={{fontSize:18,fontWeight:800,color:pct===100?T.ok:pct>50?T.brand:C.orange.solid}}>{pct}%</div>
                <div style={{fontSize:10,color:T.muted}}>faturado</div>
              </div>
              <Btn small onClick={e=>{e.stopPropagation();setNf({cliente:g.cliente,pep:g.pep,records:g.records});}}>🧾 Notas fiscais</Btn>
              <Btn small onClick={e=>{e.stopPropagation();setBulk({cliente:g.cliente,pep:g.pep,records:g.records});}}>✎ Atualizar passos</Btn>
              <span style={{fontSize:16,color:T.faint}} aria-hidden="true">{isOpen?"▲":"▼"}</span>
            </div>

            {isMobile && <div style={{ padding:"0 18px 12px" }}><PipelineStepper states={agg} groups={ggrupos} showLabels/></div>}

            {/* Detalhe dos profissionais */}
            {isOpen&&<div style={{borderTop:`1px solid ${T.lineSoft}`,padding:"0 18px 14px"}}>
              <div className="fc-scroll" style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,marginTop:10}}>
                <thead><tr style={{background:T.canvas}}>
                  {[isAdmin&&"Analista","Profissional","Funil","Período","Val. Total","NF","Status",isAdmin&&"Ações"].filter(Boolean).map(h=>
                    <th key={h} style={{padding:"7px 10px",textAlign:h==="Ações"?"right":"left",borderBottom:`1px solid ${T.line}`,fontWeight:600,color:T.muted,whiteSpace:"nowrap"}}>{h}</th>
                  )}
                </tr></thead>
                <tbody>
                  {g.records.map(r=>(
                    <tr key={r.id} className="fc-row" style={{borderBottom:`1px solid ${T.lineSoft}`}}>
                      {isAdmin&&<td style={{padding:"7px 10px"}}><Badge label={r.responsavel} color="purple" small/></td>}
                      <td style={{padding:"7px 10px",fontWeight:500,color:T.ink}}>{r.profissional}</td>
                      <td style={{padding:"7px 10px"}}><PipelineStepper states={recordStates(r.progress, r.tipo)} groups={funnelGroups(r.tipo)} size="sm"/></td>
                      <td style={{padding:"7px 10px",color:T.muted,whiteSpace:"nowrap"}}>{r.inicio} → {r.fim}</td>
                      <td style={{padding:"7px 10px",fontWeight:500}}>{fmtShort(r.valorTotal)}</td>
                      <td style={{padding:"7px 10px",fontFamily:"monospace",fontSize:11}}>{r.nfNumero||"—"}</td>
                      <td style={{padding:"7px 10px"}}><Badge label={calcStatus(r.progress)} color={calcStatusColor(r.progress)} small dot/></td>
                      {isAdmin&&<td style={{padding:"7px 10px",textAlign:"right",whiteSpace:"nowrap"}}>
                        <button title="Editar registro" onClick={()=>setRecEdit(r)} style={{border:"none",background:"none",cursor:"pointer",color:T.muted,fontSize:14,padding:"0 4px"}}>✎</button>
                        <button title="Excluir registro" onClick={()=>setRecDel(r)} style={{border:"none",background:"none",cursor:"pointer",color:T.danger,fontSize:14,padding:"0 4px"}}>🗑</button>
                      </td>}
                    </tr>
                  ))}
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

function Dashboard({ records, analista, isAdmin }) {
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
  if (filterEtapa!=="todas")   f=f.filter(r=>calcStatus(r.progress)===filterEtapa);

  const totalValor = f.reduce((a,r)=>a+(r.valorTotal||0),0);
  const faturados  = f.filter(r=>r.progress?.p5_nf);
  const naoFat     = f.filter(r=>!r.progress?.p5_nf);
  const valorFat   = faturados.reduce((a,r)=>a+(r.valorTotal||0),0);
  const valorRep   = naoFat.reduce((a,r)=>a+(r.valorTotal||0),0);
  const pctFat     = totalValor>0 ? Math.round((valorFat/totalValor)*100) : 0;

  const byEtapa = {};
  STATUS_ORDER.forEach(s=>{ byEtapa[s]={ count:0, valor:0 }; });
  f.forEach(r=>{ const s=calcStatus(r.progress); if(byEtapa[s]){byEtapa[s].count++;byEtapa[s].valor+=(r.valorTotal||0);} });

  const byAnalista = {};
  f.forEach(r=>{
    if(!byAnalista[r.responsavel]) byAnalista[r.responsavel]={ total:0, fat:0, rep:0, cnt:0, fatCnt:0 };
    byAnalista[r.responsavel].total+=(r.valorTotal||0); byAnalista[r.responsavel].cnt++;
    if(r.progress?.p5_nf){byAnalista[r.responsavel].fat+=(r.valorTotal||0);byAnalista[r.responsavel].fatCnt++;}
    else byAnalista[r.responsavel].rep+=(r.valorTotal||0);
  });

  const byEmpresa = {};
  f.forEach(r=>{ if(!byEmpresa[r.empresa])byEmpresa[r.empresa]={total:0,fat:0}; byEmpresa[r.empresa].total+=(r.valorTotal||0); if(r.progress?.p5_nf)byEmpresa[r.empresa].fat+=(r.valorTotal||0); });

  const byTipo = {};
  f.forEach(r=>{ const t=r.tipo||"—"; if(!byTipo[t])byTipo[t]={total:0,fat:0,cnt:0}; byTipo[t].total+=(r.valorTotal||0); byTipo[t].cnt++; if(r.progress?.p5_nf)byTipo[t].fat+=(r.valorTotal||0); });

  const naoFatByCliente = {};
  naoFat.forEach(r=>{
    const key=r.cliente+"|"+r.pep;
    if(!naoFatByCliente[key]) naoFatByCliente[key]={ cliente:r.cliente, pep:r.pep, responsavel:r.responsavel, count:0, valor:0, status:calcStatus(r.progress), color:calcStatusColor(r.progress) };
    naoFatByCliente[key].count++; naoFatByCliente[key].valor+=(r.valorTotal||0);
  });

  const etapaColors = { "Faturado no corte":"green","NF emitida":"teal","Cliente aprovou":"blue","Aguard. aprovação cliente":"yellow","Retorno comercial recebido":"yellow","Aguard. retorno comercial":"orange","Racional montado":"orange","Dados extraídos":"orange","Não iniciado":"gray" };
  const MetCard=({label,value,color,sub,highlight})=>(
    <Card style={{padding:"14px 16px", ...(highlight?{borderColor:C.orange.border,background:"#fffaf3"}:{})}}>
      <div style={{fontSize:11,color:T.muted,marginBottom:4,fontWeight:600,textTransform:"uppercase",letterSpacing:".3px"}}>{label}</div>
      <div style={{fontSize:20,fontWeight:800,color:color||T.ink}}>{value}</div>
      {sub&&<div style={{fontSize:11,color:T.muted,marginTop:2}}>{sub}</div>}
    </Card>
  );

  return (
    <div>
      <h1 style={{...Ty.h1, marginBottom:14}}>📊 Dashboard</h1>

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
          ?<div style={{textAlign:"center",padding:"1rem",color:T.muted,fontSize:13}}>🎉 Tudo faturado!</div>
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
  { group:"Reconhecimento & Faturamento Receita", links:[ {id:"time",icon:"📋",label:"Minha visão"}, {id:"dash",icon:"📊",label:"Dashboard"} ] },
  { group:"Cadastros", links:[ {id:"clients",icon:"🏢",label:"Clientes"} ] },
  { group:"Operação",    links:[ {id:"tasks",icon:"✅",label:"Tarefas"} ] },
];

const ADMIN_NAV_SECTION = { group:"Administração", links:[ {id:"dados",icon:"📥",label:"Importar / Exportar"}, {id:"access",icon:"🔐",label:"Gestão de acessos"} ] };

function NavLinks({ page, setPage, isAdmin, onNavigate }) {
  const sections = isAdmin ? [...NAV_SECTIONS, ADMIN_NAV_SECTION] : NAV_SECTIONS;
  return (
    <>
      {sections.map(sec=>(
        <div key={sec.group} style={{marginBottom:18}}>
          <div style={{fontSize:11,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:".5px",padding:"0 8px",marginBottom:6}}>{sec.group}</div>
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
                <span style={{fontSize:15}} aria-hidden="true">{l.icon}</span>{l.label}
              </button>
            );
          })}
        </div>
      ))}
    </>
  );
}

function UserChip({ user, isAdmin }) {
  return (
    <div style={{ display:"flex", alignItems:"center", gap:9, padding:"4px 8px 14px", marginBottom:6, borderBottom:`1px solid ${T.lineSoft}` }}>
      <Avatar name={user.name} admin={isAdmin}/>
      <div style={{ minWidth:0 }}>
        <div style={{ fontSize:13, fontWeight:700, color:T.ink, whiteSpace:"nowrap", overflow:"hidden", textOverflow:"ellipsis" }}>{user.name}</div>
        <div style={{ fontSize:11, color:isAdmin?T.brand:T.muted, fontWeight:600 }}>{isAdmin?"Administrador":"Analista"}</div>
      </div>
    </div>
  );
}

function Sidebar({ page, setPage, user, isAdmin }) {
  return (
    <aside style={{width:212,flexShrink:0,background:"#fff",borderRight:`1px solid ${T.line}`,padding:"18px 12px",display:"flex",flexDirection:"column"}}>
      <UserChip user={user} isAdmin={isAdmin}/>
      <NavLinks page={page} setPage={setPage} isAdmin={isAdmin}/>
    </aside>
  );
}

function MobileDrawer({ open, onClose, page, setPage, user, isAdmin }) {
  if (!open) return null;
  return (
    <div style={{ position:"fixed", inset:0, zIndex:250 }}>
      <div style={{ position:"absolute", inset:0, background:"rgba(15,23,42,.5)", animation:"fcOverlay .15s ease" }} onClick={onClose}/>
      <aside style={{ position:"absolute", top:0, left:0, bottom:0, width:240, background:"#fff", padding:"18px 12px", boxShadow:T.shLg, display:"flex", flexDirection:"column", overflowY:"auto" }}>
        <UserChip user={user} isAdmin={isAdmin}/>
        <NavLinks page={page} setPage={setPage} isAdmin={isAdmin} onNavigate={onClose}/>
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
        <div>{!isNew&&<Btn danger small onClick={()=>{onDelete(task.id);onClose();}}>🗑 Excluir</Btn>}</div>
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
      style={{background:"#fff",border:`1px solid ${T.line}`,borderRadius:T.rLg,padding:"10px 12px",marginBottom:8,cursor:"pointer",boxShadow:T.shSm}}
    >
      <div style={{fontSize:13,fontWeight:600,color:T.ink,marginBottom:task.desc?4:8,lineHeight:1.35}}>{task.title}</div>
      {task.desc&&<div style={{fontSize:11.5,color:T.muted,marginBottom:8,lineHeight:1.4,display:"-webkit-box",WebkitLineClamp:2,WebkitBoxOrient:"vertical",overflow:"hidden"}}>{task.desc}</div>}
      <div style={{display:"flex",alignItems:"center",gap:6,flexWrap:"wrap"}}>
        {task.recorrente&&<Badge label="🔁 Recorrente" color="teal" small/>}
        {di&&<Badge label={"📅 "+di.label} color={di.color} small/>}
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
        <div key={i} style={{border:`1px solid ${T.line}`,borderRadius:T.rLg,padding:"12px 14px",marginBottom:8,background:"#fff"}}>
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
                <label key={r} style={{display:"flex",alignItems:"center",gap:6,padding:"5px 10px",borderRadius:T.rPill,border:`1.5px solid ${on?T.brand:T.line}`,background:on?T.brandBg:"#fff",cursor:"pointer",fontSize:12,fontWeight:on?600:400,color:on?T.brand:T.inkSoft}}>
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
        <div>{!isNew && <Btn danger small onClick={()=>{onDelete(template.id);onClose();}}>🗑 Excluir modelo</Btn>}</div>
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
        <span style={{fontSize:11,color:T.muted}}>Use o botão "⚡ Gerar do mês" em cada modelo abaixo.</span>
      </div>

      <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:10}}>
        <span style={{fontSize:13,fontWeight:700,color:T.ink}}>Modelos de entrega ({templates.length})</span>
        <Btn primary small onClick={()=>setEditing({})}>+ Novo modelo</Btn>
      </div>

      {templates.length===0
        ? <div style={{fontSize:13,color:T.muted,textAlign:"center",padding:"24px",background:T.canvas,borderRadius:T.rLg,border:`1px dashed ${T.line}`}}>Nenhum modelo ainda. Crie um modelo de entrega (ex.: "Fechamento mensal") com as tarefas que se repetem todo mês.</div>
        : templates.map(t=>(
          <div key={t.id} style={{border:`1px solid ${T.line}`,borderRadius:T.rLg,padding:"12px 14px",marginBottom:8,background:"#fff",display:"flex",alignItems:"center",gap:12,flexWrap:"wrap"}}>
            <div style={{flex:1,minWidth:160}}>
              <div style={{fontSize:14,fontWeight:700,color:T.ink}}>🔁 {t.title}</div>
              <div style={{fontSize:11,color:T.muted,marginTop:2}}>{t.items.length} tarefa(s): {t.items.map(i=>i.title).join(", ").slice(0,80)}{t.items.map(i=>i.title).join(", ").length>80?"…":""}</div>
            </div>
            <Btn small onClick={()=>setEditing(t)}>✎ Editar</Btn>
            <Btn primary small onClick={()=>{ if(!/^\d{2}\/\d{4}$/.test(comp)){alert("Informe a competência no formato MM/AAAA");return;} setConfirmGen(t); }}>⚡ Gerar do mês</Btn>
          </div>
        ))}

      <div style={{display:"flex",justifyContent:"flex-end",marginTop:16}}><Btn onClick={onClose}>Fechar</Btn></div>
    </Modal>
  );
}

// ─── KANBAN ──────────────────────────────────────────────────────────────────

function Kanban({ tasks, responsaveis, isAdmin, competenciaAtual, templates, deliveries, onAdd, onUpdate, onDelete, onTemplateSave, onTemplateDelete, onGenerate }) {
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
          <span style={{fontSize:11,color:T.muted,background:"#fff",borderRadius:T.rPill,padding:"1px 8px",border:`1px solid ${T.line}`}}>{colTasks.length}</span>
        </div>
        {col.id==="inbox"&&<button onClick={()=>setEditing({ status:"inbox" })}
          style={{width:"100%",marginBottom:10,padding:"8px",border:`1.5px dashed #c7cdd6`,borderRadius:T.rMd,background:"#fff",color:T.muted,fontSize:12,fontWeight:600,cursor:"pointer"}}>
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

      <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:18,flexWrap:"wrap"}}>
        <div style={{flex:1,minWidth:180}}>
          <h1 style={Ty.h1}>✅ Tarefas do time</h1>
          <div style={{...Ty.small, marginTop:3}}>{tasks.length} tarefa(s) · arraste os cards entre as colunas ou use as setas</div>
        </div>
        <select style={{...inp,width:"auto"}} value={filterTipo} onChange={e=>setFilterTipo(e.target.value)} aria-label="Tipo de tarefa">
          <option value="todas">Todas</option>
          <option value="recorrentes">🔁 Recorrentes</option>
          <option value="ordinarias">Ordinárias</option>
        </select>
        <select style={{...inp,width:"auto"}} value={filterResp} onChange={e=>setFilterResp(e.target.value)} aria-label="Filtrar responsável">
          <option value="todos">Todos os responsáveis</option>
          {responsaveis.map(r=><option key={r}>{r}</option>)}
        </select>
        {isAdmin&&<Btn small onClick={()=>setByAnalyst(v=>!v)} style={byAnalyst?{borderColor:T.brand,color:T.brand}:{}}>👥 Por analista</Btn>}
        {isAdmin&&<Btn small onClick={()=>setShowDeliv(true)}>🗂 Entregas</Btn>}
        <Btn primary onClick={()=>setEditing({ status:"inbox" })}>+ Nova tarefa</Btn>
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
      <div style={{fontSize:12,color:T.muted,marginTop:6}}>💡 No próximo marco dá para trazer essa criação para dentro do app (função no servidor).</div>
      <div style={{display:"flex",justifyContent:"flex-end",marginTop:18}}><Btn primary onClick={onClose}>Entendi</Btn></div>
    </Modal>
  );
}

function AccessEditModal({ profile, onSave, onClose }) {
  const [name, setName]       = useState(profile.name || "");
  const [responsavel, setResp]= useState(profile.responsavel || "");
  const [isAdmin, setIsAdmin] = useState(!!profile.isAdmin);
  const [err, setErr]         = useState("");
  function save() {
    const nm = name.trim();
    if (!nm) { setErr("Informe o nome de exibição."); return; }
    onSave({ id: profile.id, name: nm, isAdmin, responsavel: responsavel.trim() });
    onClose();
  }
  return (
    <Modal title={`Editar acesso — ${profile.name}`} subtitle="Ajuste o nome, o vínculo de receitas e o papel" onClose={onClose}>
      <div style={{marginBottom:14}}>
        <Field label="Nome de exibição *"><input style={inp} value={name} onChange={e=>{setName(e.target.value);setErr("");}} placeholder="Ex: Fernanda"/></Field>
      </div>
      <div style={{marginBottom:14}}>
        <Field label="Responsável na base de receitas"><input style={inp} value={responsavel} onChange={e=>setResp(e.target.value)} placeholder="Ex: Juliana Teles"/></Field>
        <div style={{fontSize:11,color:T.muted,marginTop:4}}>Cole aqui o nome <b>exatamente</b> como aparece na coluna "Responsável" do Excel. É isso que faz a pessoa ver as receitas dela. (Se ficar vazio, usamos o nome de exibição.)</div>
      </div>
      <label style={{display:"flex",alignItems:"flex-start",gap:8,fontSize:13,color:T.inkSoft,cursor:"pointer",marginBottom:16,padding:"10px 12px",borderRadius:T.rMd,border:`1px solid ${isAdmin?T.brand:T.line}`,background:isAdmin?T.brandBg:"#fff"}}>
        <input type="checkbox" checked={isAdmin} onChange={e=>setIsAdmin(e.target.checked)} style={{width:16,height:16,marginTop:1}}/>
        <span><b style={{color:isAdmin?T.brand:T.inkSoft}}>Administrador</b><br/><span style={{fontSize:11,color:T.muted}}>Acesso completo: importar, exportar, todos os analistas e gestão de acessos.</span></span>
      </label>
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

      <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:18,flexWrap:"wrap"}}>
        <div style={{flex:1,minWidth:200}}>
          <h1 style={Ty.h1}>🔐 Gestão de acessos</h1>
          <div style={{...Ty.small, marginTop:3}}>{profiles.length} usuário(s) · {adminCount} administrador(es). Ajuste papéis e remova acessos.</div>
        </div>
        <Btn onClick={onRefresh}>↻ Atualizar lista</Btn>
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
                    <td style={{padding:"10px 14px"}}><Badge label={u.isAdmin?"Administrador":"Analista"} color={u.isAdmin?"blue":"gray"} small dot/></td>
                    <td style={{padding:"10px 14px",textAlign:"right",whiteSpace:"nowrap"}}>
                      <Btn small onClick={()=>setEditing(u)} style={{marginRight:6}}>✎ Editar</Btn>
                      <Btn small danger disabled={isSelf||lastAdmin} onClick={()=>setConfirm(u)}
                        title={isSelf?"Você não pode remover o próprio acesso":lastAdmin?"É preciso ao menos um administrador":"Remover acesso"}>🗑 Remover</Btn>
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
        <span aria-hidden="true">ℹ️</span>
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

  function save() {
    if (!(f.nome||"").trim()) { setTab("dados"); setErr("Informe o nome do cliente/grupo."); return; }
    if (f.temPortal && (!(f.portalLink||"").trim() || !(f.portalUsuario||"").trim() || !(f.portalSenha||"").trim())) {
      setTab("dados"); setErr("Como o cliente tem portal, preencha o Link, o Usuário e a Senha do portal."); return;
    }
    const cleanPeps = {}; selTipos.forEach(t=>{ const arr=(peps[t]||[]).map(s=>s.trim()).filter(Boolean); if(arr.length) cleanPeps[t]=arr; });
    const cleanProp = propostas.map(s=>s.trim()).filter(Boolean);
    const cleanEnt = entidades.map(e=>({razao:(e.razao||"").trim(),cnpj:(e.cnpj||"").replace(/\D/g,"").slice(0,14),codSap:(e.codSap||"").trim()})).filter(e=>e.cnpj||e.razao||e.codSap);
    onSave({ ...f, nome:f.nome.trim(), calendario: JSON.stringify(passos), tiposPeps: JSON.stringify(cleanPeps), propostas: JSON.stringify(cleanProp),
      cnpjs: JSON.stringify(cleanEnt), cnpj: cleanEnt[0]?.cnpj || "", codSap: cleanEnt[0]?.codSap || "" });
    onClose();
  }

  return (
    <Modal title={isNew?"Novo cliente":`Cliente — ${client.nome}`} subtitle="Perfil de faturamento do cliente" onClose={onClose} wide>
      {/* Abas: Dados x Calendário (passo a passo) */}
      <div style={{display:"flex",gap:6,borderBottom:`1px solid ${T.line}`,marginBottom:18}}>
        {[["dados","📋 Dados do cliente"],["calendario","🗓️ Calendário (passo a passo)"]].map(([id,label])=>(
          <button key={id} onClick={()=>setTab(id)} style={{border:"none",background:"none",cursor:"pointer",padding:"8px 12px",fontSize:13,fontWeight:tab===id?700:500,color:tab===id?T.brand:T.muted,borderBottom:`2px solid ${tab===id?T.brand:"transparent"}`,marginBottom:-1}}>{label}</button>
        ))}
      </div>

      {tab==="dados" && <>
      {f.incompleto && <div style={{marginBottom:16,padding:"11px 14px",borderRadius:T.rMd,background:T.warnBg,border:`1px solid ${T.warnLine}`,fontSize:12.5,color:T.warn,display:"flex",alignItems:"center",gap:10,flexWrap:"wrap"}}>
        <span style={{flex:1,minWidth:200}}>⚠ <b>Cadastro incompleto</b> (veio da carga em massa). Complete os dados e marque como concluído.</span>
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
          <div key={i} style={{border:`1px solid ${T.line}`,borderRadius:T.rLg,padding:"10px 12px",marginBottom:8,background:"#fff"}}>
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
            <label key={t} style={{display:"flex",alignItems:"center",gap:7,padding:"8px 12px",borderRadius:T.rMd,border:`1.5px solid ${on?T.brand:T.line}`,background:on?T.brandBg:"#fff",cursor:"pointer",fontSize:13,fontWeight:on?600:400,color:on?T.brand:T.inkSoft}}>
              <input type="checkbox" checked={on} onChange={()=>toggleTipo(t)} style={{width:15,height:15}}/>{t}
            </label>
          );})}
        </div>
        {selTipos.map(t=>(
          <div key={t} style={{border:`1px solid ${T.line}`,borderRadius:T.rLg,padding:"10px 12px",marginBottom:8,background:"#fff"}}>
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
                <label key={t} style={{display:"flex",alignItems:"center",gap:7,padding:"7px 11px",borderRadius:T.rMd,border:`1.5px solid ${on?T.brand:T.line}`,background:on?T.brandBg:"#fff",cursor:"pointer",fontSize:13,fontWeight:on?600:400,color:on?T.brand:T.inkSoft}}>
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

      <CSec title="Contato financeiro">
        <Field label="Nome"><input style={inp} value={f.contatoFinanceiro||""} onChange={e=>set("contatoFinanceiro",e.target.value)}/></Field>
        <Field label="E-mail"><input style={inp} type="email" placeholder="financeiro@cliente.com" value={f.contatoFinanceiroEmail||""} onChange={e=>set("contatoFinanceiroEmail",e.target.value)}/></Field>
      </CSec>

      <CSec title="Account manager (comercial)">
        <Field label="Nome"><input style={inp} value={f.accountManager||""} onChange={e=>set("accountManager",e.target.value)}/></Field>
        <Field label="E-mail"><input style={inp} type="email" placeholder="am@grupofcamara.com" value={f.accountManagerEmail||""} onChange={e=>set("accountManagerEmail",e.target.value)}/></Field>
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
          <div key={i} style={{border:`1px solid ${T.line}`,borderRadius:T.rLg,marginBottom:10,background:"#fff",overflow:"hidden"}}>
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
        <div>{!isNew && <Btn danger small onClick={()=>{onDelete(client.id);onClose();}}>🗑 Excluir</Btn>}</div>
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
        style={{border:`2px dashed ${fileName?T.okLine:"#cbd2dc"}`,borderRadius:T.rLg,padding:"28px 20px",textAlign:"center",cursor:"pointer",background:fileName?T.okBg:"#fafbfc",marginBottom:14}}>
        {loading?<div style={{color:T.muted,fontSize:13}}>⏳ Lendo...</div>:fileName?<><div style={{fontSize:24}}>✅</div><div style={{fontSize:13,fontWeight:700,color:T.ok}}>{fileName}</div><div style={{fontSize:11,color:T.muted}}>Clique para trocar</div></>:<><div style={{fontSize:28}}>📂</div><div style={{fontSize:14,fontWeight:600,color:T.inkSoft}}>Clique para selecionar a planilha</div></>}
      </div>
      {msgs.map((m,i)=><div key={i} style={{marginBottom:6,fontSize:12,padding:"8px 12px",borderRadius:T.rMd,background:mc[m.type].bg,color:mc[m.type].text,border:`1px solid ${mc[m.type].border}`}}>{m.text}</div>)}
      {preview&&<div style={{marginBottom:14,padding:"10px 14px",borderRadius:T.rMd,background:T.okBg,border:`1px solid ${T.okLine}`,fontSize:12,color:T.ok}}>
        Prévia: {preview.slice(0,3).map(c=>c.nome).join(", ")}{preview.length>3?` … e mais ${preview.length-3}`:""}
      </div>}
      <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
        <Btn onClick={onClose}>Cancelar</Btn>
        <Btn primary disabled={!preview} onClick={()=>{ if(preview){onImport(preview);} onClose(); }}>⬆ Importar {preview?`(${preview.length})`:""}</Btn>
      </div>
    </Modal>
  );
}

// Conta quantos CNPJs um cadastro reúne (grupo de empresas).
function clientCnpjs(c) {
  const a = parseJSON(c.cnpjs, null);
  if (Array.isArray(a) && a.length) return a.map(e=>(e.cnpj||"")).filter(Boolean);
  return c.cnpj ? [c.cnpj] : [];
}

function ClientsView({ clients, isAdmin, onSave, onDelete, onBulkImport, onMerge }) {
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
  const selList = clients.filter(c=>sel.has(c.id));
  const startMerge = () => setMerging({ ids:[...sel], nome: selList[0]?.nome || "" });
  const confirmMerge = async () => {
    await onMerge(merging.ids, merging.nome.trim() || selList[0]?.nome || "Grupo");
    setSel(new Set()); setMerging(null);
  };
  // Incluir os selecionados num grupo já existente (o grupo é o cadastro-base).
  const startAdd = () => setAdding({ ids:[...sel], pick:"" });
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
        <Modal title="Incluir em grupo existente" onClose={()=>setAdding(null)} footer={<Btn onClick={()=>setAdding(null)}>Cancelar</Btn>}>
          <div style={{fontSize:13,color:T.inkSoft,marginBottom:12,lineHeight:1.6}}>
            Escolha o <b>cadastro/grupo de destino</b>. Os <b>{adding.ids.length}</b> cadastros selecionados serão reunidos nele (CNPJs somados) e <b>removidos</b> como cadastros separados.
          </div>
          <Field label="Buscar grupo de destino">
            <input style={inp} autoFocus value={adding.pick} onChange={e=>setAdding(a=>({...a,pick:e.target.value}))} placeholder="Digite o nome do grupo… ex.: Elfa"/>
          </Field>
          <div style={{marginTop:10,border:`1px solid ${T.line}`,borderRadius:T.rMd,overflow:"hidden"}}>
            {addMatches.length===0
              ? <div style={{padding:"14px",fontSize:12.5,color:T.muted,textAlign:"center"}}>Nenhum cadastro encontrado.</div>
              : addMatches.map(c=>{
                  const n=clientCnpjs(c).length;
                  return (
                    <button key={c.id} onClick={()=>confirmAdd(c)} className="fc-row" style={{display:"flex",alignItems:"center",gap:8,width:"100%",textAlign:"left",border:"none",borderBottom:`1px solid ${T.lineSoft}`,background:"#fff",cursor:"pointer",padding:"10px 12px",fontSize:13,color:T.ink}}>
                      <span style={{flex:1,fontWeight:600}}>{c.nome}</span>
                      {n>1 && <Badge label={`🔗 ${n} CNPJs`} color="blue" small/>}
                      <span style={{fontSize:18,color:T.brand}}>＋</span>
                    </button>
                  );
                })}
          </div>
        </Modal>
      )}

      <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:14,flexWrap:"wrap"}}>
        <div style={{flex:1,minWidth:200}}>
          <h1 style={Ty.h1}>🏢 Clientes</h1>
          <div style={{...Ty.small, marginTop:3}}>{clients.length} cliente(s){incompletos>0 && <> · <b style={{color:T.warn}}>{incompletos} incompleto(s)</b></>}</div>
        </div>
        {isAdmin && <Btn onClick={()=>setImporting(true)}>⬆ Importar clientes</Btn>}
        <Btn primary onClick={()=>setEditing({ temPortal:false })}>+ Novo cliente</Btn>
      </div>

      <Card style={{padding:"10px 12px",marginBottom:14}}>
        <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
          <select style={{...inp,width:"auto"}} value={status} onChange={e=>resetPage(setStatus)(e.target.value)} aria-label="Status do cadastro">
            <option value="todos">Todos os cadastros</option>
            <option value="incompletos">⚠ Incompletos</option>
            <option value="completos">✓ Completos</option>
            <option value="grupos">🔗 Grupos (vários CNPJs)</option>
          </select>
          {grupos.length>0 && (
            <>
              <input style={{...inp,width:160}} list="fc-grupos" placeholder="Grupo de empresa…" value={grupo} onChange={e=>resetPage(setGrupo)(e.target.value)} aria-label="Filtrar por grupo de empresa"/>
              <datalist id="fc-grupos">{grupos.map(g=><option key={g} value={g}/>)}</datalist>
            </>
          )}
          <input style={{...inp,flex:1,minWidth:200}} placeholder="🔎 Nome, Cód. SAP ou CNPJ..." value={q} onChange={e=>resetPage(setQ)(e.target.value)}/>
        </div>
      </Card>

      {sel.size>0 && (
        <Card style={{padding:"10px 14px",marginBottom:14,display:"flex",alignItems:"center",gap:12,flexWrap:"wrap",background:T.canvas}}>
          <b style={{fontSize:13,color:T.ink}}>{sel.size} selecionado(s)</b>
          <Btn small onClick={()=>setSel(new Set())}>Limpar</Btn>
          <div style={{flex:1}}/>
          <Btn small onClick={startAdd}>➕ Incluir em grupo existente</Btn>
          <Btn primary small disabled={sel.size<2} onClick={startMerge}>🔗 Agrupar (novo grupo)</Btn>
        </Card>
      )}

      {filtered.length===0
        ? <Card style={{textAlign:"center",padding:"3rem"}}>
            <div style={{fontSize:32,marginBottom:10}}>🏢</div>
            <div style={{fontSize:14,color:T.muted}}>{clients.length===0?"Nenhum cliente cadastrado. Importe a base ou clique em “+ Novo cliente”.":"Nenhum cliente encontrado."}</div>
          </Card>
        : <Card style={{padding:0,overflow:"hidden"}}>
            <div className="fc-scroll" style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                <thead><tr style={{background:T.canvas}}>
                  <th style={{padding:"10px 14px",textAlign:"left",borderBottom:`1px solid ${T.line}`,width:34}}></th>
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
                        {c.nome} {cs.length>1 && <Badge label={`🔗 ${cs.length} CNPJs`} color="blue" small/>}
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

// ─── IMPORTAR / EXPORTAR (admin) ─────────────────────────────────────────────

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
      <h1 style={{ ...Ty.h1, marginBottom:6 }}>📥 Importar / Exportar</h1>
      <div style={{ ...Ty.small, marginBottom:18 }}>Carga de dados de reconhecimento e exportações</div>
      <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill,minmax(240px,1fr))", gap:14 }}>
        <Tile icon="⬆️" title="Importar receitas" primary desc={`Carregar a planilha de T&E (.xlsm/.xlsx). ${recordsCount} registro(s) hoje.`} btn="Importar planilha" onClick={onImport}/>
        <Tile icon="⬇️" title="Exportar receitas" desc="Baixar os registros de faturamento em CSV, com filtros." btn="Exportar receitas" onClick={onExport}/>
        <Tile icon="🏢" title="Exportar clientes" desc={`Baixar os cadastros completos de clientes em CSV. ${clientsCount} cliente(s).`} btn="Exportar clientes" onClick={onExportClients}/>
        <Tile icon="🕐" title="Histórico de importações" desc="Ver o log de todas as importações realizadas." btn="Ver histórico" onClick={onHistory}/>
      </div>
    </div>
  );
}

// ─── TOPBAR ──────────────────────────────────────────────────────────────────

function Topbar({ user, isAdmin, isMobile, onMenu, onLogout }) {
  const ghostBtn = { background:"rgba(255,255,255,.16)", border:"none", color:"#fff", borderRadius:T.rMd, padding:"6px 12px", fontSize:12, fontWeight:600, cursor:"pointer", display:"inline-flex", alignItems:"center", gap:5 };
  return (
    <div style={{background:T.brand,color:"#fff",padding:"0 16px",display:"flex",alignItems:"center",gap:10,height:54,boxShadow:T.shSm}}>
      {isMobile && <button onClick={onMenu} aria-label="Abrir menu" style={{ background:"none", border:"none", color:"#fff", fontSize:22, cursor:"pointer", lineHeight:1, padding:4 }}>☰</button>}
      <span style={{fontSize:14,fontWeight:800,flex:1,display:"flex",alignItems:"center",gap:9}}>
        <FcamaraLogo size={30}/>{isMobile ? "Faturamento" : "Faturamento Grupo Fcamara"}
      </span>

      {!isMobile && <span style={{display:"flex",alignItems:"center",gap:7,fontSize:12,opacity:.95,paddingLeft:4}}>
        <Avatar name={user.name} size={26} admin={isAdmin}/>{user.name}{isAdmin?" · Admin":""}
      </span>}
      <button className="fc-btn" onClick={onLogout} style={{...ghostBtn, background:"rgba(255,255,255,.12)"}}>Sair</button>
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
    <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:`linear-gradient(135deg,${T.brandDark},${T.brand})`,fontFamily:"system-ui,sans-serif",padding:16}}>
      {showForgot && <ForgotPasswordModal onClose={()=>setSF(false)}/>}
      <div style={{background:"#fff",borderRadius:18,padding:"34px 38px",width:400,maxWidth:"100%",boxShadow:T.shLg}}>
        <div style={{textAlign:"center",marginBottom:24}}>
          <div style={{ display:"inline-flex", marginBottom:12 }}><FcamaraLogo size={62}/></div>
          <h1 style={{fontSize:19,fontWeight:800,color:T.ink,lineHeight:1.3,margin:0}}>Controle de Faturamento</h1>
          <p style={{fontSize:13,color:T.brand,fontWeight:600,marginTop:4,marginBottom:0}}>Grupo Fcamara</p>
          <p style={{fontSize:11,color:T.muted,marginTop:6}}>{APP_VERSION}</p>
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

// ─── APP ROOT ─────────────────────────────────────────────────────────────────

function AppInner() {
  const toast = useToast();
  const isMobile = useIsMobile();
  const [state, setState]       = useState(()=>loadState());
  const [user, setUser]         = useState(null);
  const [authReady, setAuthRdy] = useState(false);
  const [recovery, setRecovery] = useState(false);
  const [page, setPage]         = useState("dash");
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
  const [dataReady, setDataRdy] = useState(false);

  useEffect(()=>saveState(state),[state]);

  // ─ Carrega os dados do banco quando o usuário entra ─
  const reloadRecords = useCallback(async () => { try { setRecords(await db.fetchRecords()); } catch(e){ toast("Erro ao carregar registros: "+e.message, "error"); } }, [toast]);
  const reloadTasks   = useCallback(async () => { try { setTasks(await db.fetchTasks()); }     catch(e){ toast("Erro ao carregar tarefas: "+e.message, "error"); } }, [toast]);
  const reloadHistory  = useCallback(async () => { try { setHistory(await db.fetchHistory()); } catch(e){ /* histórico é só p/ admin */ } }, []);
  const reloadProfiles = useCallback(async () => { try { setProfiles(await db.fetchProfiles()); } catch(e){ toast("Erro ao carregar acessos: "+e.message, "error"); } }, [toast]);
  const reloadClients  = useCallback(async () => { try { setClients(await db.fetchClients()); } catch(e){ toast("Erro ao carregar clientes: "+e.message, "error"); } }, [toast]);
  const reloadTemplates = useCallback(async () => { try { setTemplates(await db.fetchTemplates()); } catch(e){ /* entregas: só admin */ } }, []);
  const reloadDeliveries = useCallback(async () => { try { setDeliveries(await db.fetchDeliveries()); } catch(e){ /* idem */ } }, []);

  useEffect(() => {
    if (!user) { setRecords([]); setTasks([]); setHistory([]); setProfiles([]); setClients([]); setTemplates([]); setDeliveries([]); return; }
    let active = true;
    // NÃO voltamos para a tela de "Carregando" em recargas — isso desmontaria
    // formulários/modais abertos. A tela de carregamento só aparece na 1ª vez.
    Promise.all([db.fetchRecords(), db.fetchTasks(), db.fetchHistory().catch(()=>[]), db.fetchProfiles().catch(()=>[]), db.fetchClients().catch(()=>[]), db.fetchTemplates().catch(()=>[]), db.fetchDeliveries().catch(()=>[])])
      .then(([r, t, h, p, c, tm, dv]) => { if (!active) return; setRecords(r); setTasks(t); setHistory(h); setProfiles(p); setClients(c); setTemplates(tm); setDeliveries(dv); })
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
        const { data: prof } = await supabase.from("profiles").select("name,is_admin").eq("id", session.user.id).single();
        if (!mounted) return;
        const next = { id: session.user.id, name: prof?.name || session.user.email, isAdmin: !!prof?.is_admin, email: session.user.email };
        const isNewLogin = userIdRef.current !== next.id;
        userIdRef.current = next.id;
        // Mantém a MESMA referência do usuário se nada mudou — evita recarregar
        // a tela (e fechar modais) quando o app volta do foco / renova o token.
        setUser(prev => (prev && prev.id===next.id && prev.name===next.name && prev.isAdmin===next.isAdmin) ? prev : next);
        if (greet && isNewLogin) toast(`Bem-vinda, ${(next.name||"").split(" ")[0]}!`, "info");
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

  async function handleUpdateBulk(updatedList) {
    try {
      await db.upsertRecords(updatedList);
      await reloadRecords();
      toast(`Passos atualizados — ${updatedList.length} profissional(is)`);
    } catch(e) { toast("Erro ao salvar os passos: "+e.message, "error"); }
  }

  async function handleImport({ records:newRecs, competencia, empresa, tipo, mode, note }) {
    try {
      if (mode==="replace") {
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
    if (!entry?.importId) { toast("Esta importação é antiga e não pode ser desfeita automaticamente.", "error"); return; }
    try {
      const removed = await db.deleteRecordsByImport(entry.importId);
      await db.deleteHistory(entry.id);
      await Promise.all([reloadRecords(), reloadHistory()]);
      toast(`Importação desfeita — ${removed} registro(s) removido(s)`, "info");
    } catch(e) { toast("Erro ao desfazer importação: "+e.message, "error"); }
  }

  function handleCompetencia(val) { setState(s=>({...s, competenciaAtual:val})); }

  async function handleRecordDelete(id) {
    try { await db.deleteRecord(id); await reloadRecords(); toast("Registro excluído", "info"); }
    catch(e) { toast("Erro ao excluir registro: "+e.message, "error"); }
  }

  async function handleTaskAdd(t)    { try { await db.insertTask(t); await reloadTasks(); toast("Tarefa criada"); } catch(e){ toast("Erro ao criar tarefa: "+e.message,"error"); } }
  async function handleTaskUpdate(u) { try { await db.updateTask(u); await reloadTasks(); } catch(e){ toast("Erro ao atualizar tarefa: "+e.message,"error"); } }
  async function handleTaskDelete(id){ try { await db.deleteTask(id); await reloadTasks(); toast("Tarefa excluída","info"); } catch(e){ toast("Erro ao excluir tarefa: "+e.message,"error"); } }

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
    // Impede rebaixar o último administrador.
    if (data.isAdmin === false) {
      const target = profiles.find(p => p.id === data.id);
      if (target?.isAdmin && profiles.filter(p => p.isAdmin).length <= 1) { toast("É preciso manter ao menos um administrador.", "error"); return; }
    }
    try {
      await db.updateProfile(data);
      await reloadProfiles();
      // Se o admin alterou o próprio papel/nome, reflete na sessão atual.
      if (data.id === user.id) setUser(u => ({ ...u, name: data.name, isAdmin: data.isAdmin }));
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
    try {
      if (c.id) { await db.updateClient(c); toast("Cliente atualizado"); }
      else { await db.insertClient(c); toast("Cliente cadastrado"); }
      await reloadClients();
    } catch(e) { toast("Erro ao salvar cliente: "+e.message, "error"); }
  }
  async function handleClientDelete(id) {
    try { await db.deleteClient(id); await reloadClients(); toast("Cliente excluído", "info"); }
    catch(e) { toast("Erro ao excluir cliente: "+e.message, "error"); }
  }
  async function handleClientsImport(list) {
    try { const n = await db.bulkInsertClients(list); await reloadClients(); toast(`${n} cliente(s) importados (cadastro incompleto)`); }
    catch(e) { toast("Erro ao importar clientes: "+e.message, "error"); }
  }
  // Agrupa N cadastros num só. baseId define qual cadastro é o destino (grupo);
  // sem baseId, usa o primeiro da lista. Reúne os CNPJs no destino e remove os demais.
  async function handleClientsMerge(ids, nome, baseId) {
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

  const responsaveis = [...new Set([...profiles.map(p=>p.name), ...records.map(r=>r.responsavel)].filter(Boolean))].sort();

  if (recovery) return (
    <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:`linear-gradient(135deg,${T.brandDark},${T.brand})`,padding:16}}>
      <RecoveryModal onClose={()=>{ setRecovery(false); }}/>
    </div>
  );

  if (!authReady) return (
    <div style={{minHeight:"100vh",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:14,background:T.canvas,color:T.muted,fontFamily:"system-ui,sans-serif"}}>
      <FcamaraLogo size={54}/>
      <div style={{fontSize:13}}>Carregando…</div>
    </div>
  );

  if (!user) return <Login/>;

  if (!dataReady) return (
    <div style={{minHeight:"100vh",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:14,background:T.canvas,color:T.muted,fontFamily:"system-ui,sans-serif"}}>
      <FcamaraLogo size={54}/>
      <div style={{fontSize:13}}>Carregando dados…</div>
    </div>
  );

  return (
    <div style={{fontFamily:"system-ui,-apple-system,sans-serif",color:T.ink,minHeight:"100vh",background:T.canvas,display:"flex",flexDirection:"column"}}>
      {showImport  && <ImportModal onImport={handleImport} onClose={()=>setImp(false)}/>}
      {showExport  && <ExportModal records={records} onClose={()=>setExp(false)} onDone={(n)=>toast(`CSV exportado — ${n} registros`)}/>}
      {showHistory && <HistoryModal history={history} onClose={()=>setHist(false)} onUndo={(entry)=>handleUndoImport(entry)}/>}
      {confirmLogout && <ConfirmDialog title="Sair da plataforma" message="Deseja realmente encerrar a sessão?" confirmLabel="Sair" onConfirm={()=>{ supabase.auth.signOut(); setUser(null); }} onClose={()=>setCL(false)}/>}

      <Topbar user={user} isAdmin={isAdmin} isMobile={isMobile} onMenu={()=>setDrawer(true)} onLogout={()=>setCL(true)}/>

      {isAdmin&&<div style={{background:T.warnBg,borderBottom:`1px solid ${T.warnLine}`,padding:"7px 20px",fontSize:12,color:T.warn,display:"flex",alignItems:"center",gap:8}}>
        <Badge label="Admin" color="blue" small/> Acesso completo a todos os analistas, empresas e competências.
      </div>}

      {isMobile && <MobileDrawer open={drawer} onClose={()=>setDrawer(false)} page={page} setPage={setPage} user={user} isAdmin={isAdmin}/>}

      <div style={{display:"flex",flex:1,minHeight:0}}>
        {!isMobile && <Sidebar page={page} setPage={setPage} user={user} isAdmin={isAdmin}/>}
        <main style={{flex:1,overflowX:"auto",minWidth:0}}>
          {(page==="time"||page==="dash")&&(
            <div style={{maxWidth:1140,margin:"0 auto",padding:isMobile?"18px 14px":"24px 22px"}}>
              {page==="time"&&<MyView records={records} analista={user.name} isAdmin={isAdmin} onUpdateBulk={handleUpdateBulk} onDeleteRecord={handleRecordDelete} competenciaAtual={state.competenciaAtual} onCompetenciaChange={handleCompetencia}/>}
              {page==="dash"&&<Dashboard records={records} analista={user.name} isAdmin={isAdmin}/>}
            </div>
          )}
          {page==="dados"&&isAdmin&&(
            <div style={{maxWidth:1140,margin:"0 auto",padding:isMobile?"18px 14px":"24px 22px"}}>
              <DataIOView recordsCount={records.length} clientsCount={clients.length}
                onImport={()=>setImp(true)} onExport={()=>setExp(true)} onHistory={()=>setHist(true)}
                onExportClients={()=>{ exportClientsCSV(clients); toast(`Clientes exportados — ${clients.length} cadastro(s)`); }}/>
            </div>
          )}
          {page==="clients"&&(
            <div style={{maxWidth:1140,margin:"0 auto",padding:isMobile?"18px 14px":"24px 22px"}}>
              <ClientsView clients={clients} isAdmin={isAdmin} onSave={handleClientSave} onDelete={handleClientDelete} onBulkImport={handleClientsImport} onMerge={handleClientsMerge}/>
            </div>
          )}
          {page==="tasks"&&(
            <div style={{padding:isMobile?"18px 14px":"24px 22px"}}>
              <Kanban tasks={tasks} responsaveis={responsaveis} isAdmin={isAdmin} competenciaAtual={state.competenciaAtual}
                templates={templates} deliveries={deliveries}
                onAdd={handleTaskAdd} onUpdate={handleTaskUpdate} onDelete={handleTaskDelete}
                onTemplateSave={handleTemplateSave} onTemplateDelete={handleTemplateDelete} onGenerate={handleGenerateDelivery}/>
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

