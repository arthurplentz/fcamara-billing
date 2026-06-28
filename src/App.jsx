import { useState, useEffect, useRef, useCallback } from "react";
import * as XLSX from "xlsx";

const APP_VERSION = "v3.0";
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
  { num: 1, title: "Extração de dados",    steps: ["p1_extrair"] },
  { num: 2, title: "Racional",             steps: ["p2_racional"] },
  { num: 3, title: "Validação comercial",  steps: ["p3_envio_com","p3_retorno_com","p3_data_retorno"] },
  { num: 4, title: "Aprovação do cliente", steps: ["p4_envio_cli","p4_aprovacao","p4_data_aprov"] },
  { num: 5, title: "Faturamento",          steps: ["p5_nf","p5_data_nf","p5_no_corte"] },
];

// Usuários com login/senha (senha = nome em minúsculas sem acento, simplificado)
const USERS = [
  { name: "Daniela",     password: "daniela",     isAdmin: true  },
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

// ─── HELPERS ─────────────────────────────────────────────────────────────────

const fmtShort = (n) => n == null ? "—" : "R$ " + Math.round(n).toLocaleString("pt-BR");
const nowISO   = () => new Date().toISOString();
const fmtDT    = (iso) => { if (!iso) return "—"; const d = new Date(iso); return d.toLocaleDateString("pt-BR") + " " + d.toLocaleTimeString("pt-BR", { hour:"2-digit", minute:"2-digit" }); };
const genId    = () => "r" + Date.now() + Math.random().toString(36).slice(2,7);
const makeProgress = () => Object.fromEntries(STEPS.map(s => [s.id, s.type==="date" ? "" : false]));

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

function initState() {
  const now = nowISO();
  const records = SAMPLE_RECORDS.map(r => ({ ...r, id: genId(), progress: makeProgress(), nfNumero: "", obs: "", updatedAt: now }));
  return {
    records,
    competenciaAtual: "05/2026",
    importHistory: [{ id: genId(), date: now, competencia:"05/2026", empresa:"BR02", tipo:"Time & Expenses", mode:"add", count: records.length, user: ADMIN_NAME, note:"Carga inicial de exemplo" }],
  };
}

function loadState() {
  try { const r = localStorage.getItem(LS_KEY); if (r) return JSON.parse(r); } catch {}
  return initState();
}

function saveState(s) { try { localStorage.setItem(LS_KEY, JSON.stringify(s)); } catch {} }

// ─── COLORS / STYLES ─────────────────────────────────────────────────────────

const C = {
  green:  { bg:"#dcfce7", text:"#166534", border:"#86efac" },
  teal:   { bg:"#ccfbf1", text:"#134e4a", border:"#5eead4" },
  blue:   { bg:"#dbeafe", text:"#1e40af", border:"#93c5fd" },
  yellow: { bg:"#fef9c3", text:"#854d0e", border:"#fde047" },
  orange: { bg:"#ffedd5", text:"#9a3412", border:"#fdba74" },
  gray:   { bg:"#f3f4f6", text:"#374151", border:"#d1d5db" },
  red:    { bg:"#fee2e2", text:"#991b1b", border:"#fca5a5" },
  purple: { bg:"#f3e8ff", text:"#6b21a8", border:"#d8b4fe" },
};

const inp = { padding:"7px 10px", borderRadius:8, border:"1px solid #d1d5db", fontSize:13, fontFamily:"inherit", background:"#fff", color:"#111827", width:"100%", boxSizing:"border-box" };

// ─── ATOMS ───────────────────────────────────────────────────────────────────

function Badge({ label, color="gray", small }) {
  const c = C[color]||C.gray;
  return <span style={{ fontSize:small?10:11, padding:small?"2px 7px":"3px 10px", borderRadius:20, background:c.bg, color:c.text, border:`1px solid ${c.border}`, fontWeight:500, whiteSpace:"nowrap" }}>{label}</span>;
}

function Btn({ children, onClick, primary, danger, small, disabled, style:s={} }) {
  const base = { padding:small?"5px 12px":"8px 18px", borderRadius:8, fontSize:small?12:13, fontWeight:600, cursor:disabled?"not-allowed":"pointer", display:"inline-flex", alignItems:"center", gap:6, border:"none", opacity:disabled?.5:1, ...s };
  const v = primary?{ background:"#1d4ed8", color:"#fff" }:danger?{ background:"#dc2626", color:"#fff" }:{ background:"#f3f4f6", color:"#374151", border:"1px solid #d1d5db" };
  return <button onClick={disabled?undefined:onClick} style={{...base,...v}}>{children}</button>;
}

function Modal({ title, onClose, children, wide, extraWide }) {
  const w = extraWide ? 960 : wide ? 780 : 520;
  return (
    <div style={{ position:"fixed", inset:0, background:"rgba(0,0,0,.5)", display:"flex", alignItems:"center", justifyContent:"center", zIndex:300 }} onClick={onClose}>
      <div style={{ background:"#fff", borderRadius:14, padding:"24px 28px", width:w, maxWidth:"96vw", maxHeight:"92vh", overflowY:"auto", boxShadow:"0 12px 40px rgba(0,0,0,.2)" }} onClick={e=>e.stopPropagation()}>
        <div style={{ display:"flex", alignItems:"center", marginBottom:20 }}>
          <h2 style={{ fontSize:17, fontWeight:700, flex:1 }}>{title}</h2>
          <button onClick={onClose} style={{ background:"none", border:"none", fontSize:22, cursor:"pointer", color:"#9ca3af", lineHeight:1 }}>×</button>
        </div>
        {children}
      </div>
    </div>
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

function ImportModal({ onImport, onClose }) {
  const [competencia,setComp]=useState("");
  const [empresa,setEmpresa]=useState("BR02");
  const [tipo,setTipo]=useState("Time & Expenses");
  const [mode,setMode]=useState("add");
  const [note,setNote]=useState("");
  const [preview,setPreview]=useState(null);
  const [fileName,setFileName]=useState("");
  const [msgs,setMsgs]=useState([]);
  const [loading,setLoading]=useState(false);
  const [dragOver,setDragOver]=useState(false);
  const fileRef=useRef();

  const reset=()=>{setPreview(null);setFileName("");setMsgs([]);};

  function readFile(file) {
    if(!competencia.match(/^\d{2}\/\d{4}$/)){setMsgs([{type:"error",text:"Informe a competência no formato MM/AAAA antes de carregar o arquivo."}]);return;}
    setLoading(true);setFileName(file.name);setPreview(null);setMsgs([]);
    const reader=new FileReader();
    reader.onload=e=>{
      try {
        const wb=XLSX.read(new Uint8Array(e.target.result),{type:"array",cellDates:false});
        const sheetName=wb.SheetNames.find(n=>n.toLowerCase().includes("time")&&n.toLowerCase().includes("expense"))||wb.SheetNames[0];
        const rows=XLSX.utils.sheet_to_json(wb.Sheets[sheetName],{header:1,defval:""});
        const {records,errors}=parseSheetRows(rows,empresa,tipo,competencia);
        const m=[];
        errors.forEach(e=>m.push({type:"warn",text:e}));
        if(records.length===0){m.push({type:"error",text:"Nenhum registro válido. Verifique a aba '📥 Time & Expenses'."});setMsgs(m);}
        else{m.push({type:"ok",text:`${records.length} registros encontrados na aba "${sheetName}".`});setMsgs(m);setPreview(records);}
      } catch(err){setMsgs([{type:"error",text:"Erro ao ler o arquivo: "+err.message}]);}
      setLoading(false);
    };
    reader.readAsArrayBuffer(file);
  }

  const onDrop=useCallback(e=>{e.preventDefault();setDragOver(false);const f=e.dataTransfer.files[0];if(f)readFile(f);},[competencia,empresa,tipo]);
  const mc={ok:{bg:"#f0fdf4",text:"#166534",border:"#86efac"},warn:{bg:"#fffbeb",text:"#92400e",border:"#fde68a"},error:{bg:"#fef2f2",text:"#991b1b",border:"#fca5a5"}};

  return (
    <Modal title="Importar — Time & Expenses" onClose={onClose} wide>
      <div style={{background:"#fffbeb",border:"1px solid #fde68a",borderRadius:8,padding:"10px 14px",marginBottom:16,fontSize:12,color:"#92400e"}}>
        🔒 Apenas a Daniela pode importar. Lê a aba <strong>📥 Time & Expenses</strong> diretamente do arquivo .xlsm/.xlsx.
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:12,marginBottom:14}}>
        <div><label style={{fontSize:12,color:"#6b7280",display:"block",marginBottom:4}}>Competência * <span style={{color:"#dc2626"}}>(preencha primeiro)</span></label><input style={inp} placeholder="05/2026" value={competencia} onChange={e=>{setComp(e.target.value);reset();}}/></div>
        <div><label style={{fontSize:12,color:"#6b7280",display:"block",marginBottom:4}}>Empresa</label><select style={inp} value={empresa} onChange={e=>{setEmpresa(e.target.value);reset();}}>{EMPRESAS.map(e=><option key={e.cod} value={e.cod}>{e.cod} — {e.nome}</option>)}</select></div>
        <div><label style={{fontSize:12,color:"#6b7280",display:"block",marginBottom:4}}>Tipo de projeto</label><select style={inp} value={tipo} onChange={e=>{setTipo(e.target.value);reset();}}>{TIPOS_PROJETO.map(t=><option key={t}>{t}</option>)}</select></div>
      </div>
      <div style={{marginBottom:14}}>
        <label style={{fontSize:12,color:"#6b7280",display:"block",marginBottom:6}}>Modo de importação</label>
        <div style={{display:"flex",gap:10}}>
          {[{v:"add",l:"➕ Incluir novos",d:"Adiciona sem apagar registros existentes."},{v:"replace",l:"🔄 Substituir",d:"Remove TODOS desta competência/empresa/tipo e reimporta."}].map(opt=>(
            <label key={opt.v} style={{flex:1,display:"flex",gap:8,padding:"10px 12px",borderRadius:8,border:`2px solid ${mode===opt.v?"#1d4ed8":"#e5e7eb"}`,cursor:"pointer",background:mode===opt.v?"#eff6ff":"#fff"}}>
              <input type="radio" name="mode" value={opt.v} checked={mode===opt.v} onChange={()=>setMode(opt.v)} style={{marginTop:2}}/>
              <div><div style={{fontSize:13,fontWeight:700,color:mode===opt.v?"#1d4ed8":"#374151"}}>{opt.l}</div><div style={{fontSize:11,color:"#6b7280",marginTop:2}}>{opt.d}</div></div>
            </label>
          ))}
        </div>
        {mode==="replace"&&<div style={{marginTop:8,fontSize:12,color:"#dc2626",fontWeight:600}}>⚠ O progresso já registrado para esta competência será perdido.</div>}
      </div>
      <div style={{marginBottom:14}}><label style={{fontSize:12,color:"#6b7280",display:"block",marginBottom:4}}>Nota da importação (opcional)</label><input style={inp} placeholder="Ex: Ajuste de valores de maio" value={note} onChange={e=>setNote(e.target.value)}/></div>
      <input type="file" ref={fileRef} style={{display:"none"}} accept=".xlsx,.xlsm,.xls" onChange={e=>{if(e.target.files[0])readFile(e.target.files[0]);e.target.value="";}}/>
      <div onDragOver={e=>{e.preventDefault();setDragOver(true);}} onDragLeave={()=>setDragOver(false)} onDrop={onDrop} onClick={()=>fileRef.current.click()}
        style={{border:`2px dashed ${dragOver?"#1d4ed8":fileName?"#86efac":"#d1d5db"}`,borderRadius:10,padding:"28px 20px",textAlign:"center",cursor:"pointer",background:dragOver?"#eff6ff":fileName?"#f0fdf4":"#fafafa",marginBottom:14}}>
        {loading?<div style={{color:"#6b7280",fontSize:13}}>⏳ Lendo arquivo...</div>:fileName?<><div style={{fontSize:24,marginBottom:6}}>✅</div><div style={{fontSize:13,fontWeight:700,color:"#166534"}}>{fileName}</div><div style={{fontSize:11,color:"#9ca3af",marginTop:4}}>Clique para trocar</div></>:<><div style={{fontSize:28,marginBottom:8}}>📂</div><div style={{fontSize:14,fontWeight:600,color:"#374151"}}>Clique ou arraste o arquivo aqui</div><div style={{fontSize:12,color:"#9ca3af",marginTop:4}}>Aceita .xlsm e .xlsx</div></>}
      </div>
      {msgs.map((m,i)=><div key={i} style={{marginBottom:6,fontSize:12,padding:"8px 12px",borderRadius:8,background:mc[m.type].bg,color:mc[m.type].text,border:`1px solid ${mc[m.type].border}`}}>{m.text}</div>)}
      {preview&&<div style={{marginBottom:14,padding:"12px 14px",borderRadius:8,background:"#f0fdf4",border:"1px solid #86efac"}}>
        <div style={{fontSize:13,fontWeight:700,color:"#166534",marginBottom:8}}>✓ {preview.length} registros prontos</div>
        <div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:11}}><thead><tr style={{background:"#dcfce7"}}>{["Responsável","Cliente","PEP","Profissional","Val. Total"].map(h=><th key={h} style={{padding:"4px 8px",textAlign:"left",color:"#166534",fontWeight:700}}>{h}</th>)}</tr></thead><tbody>{preview.slice(0,5).map(r=><tr key={r.id}><td style={{padding:"4px 8px"}}>{r.responsavel}</td><td style={{padding:"4px 8px"}}>{r.cliente}</td><td style={{padding:"4px 8px",fontFamily:"monospace"}}>{r.pep}</td><td style={{padding:"4px 8px"}}>{r.profissional}</td><td style={{padding:"4px 8px"}}>{fmtShort(r.valorTotal)}</td></tr>)}{preview.length>5&&<tr><td colSpan={5} style={{padding:"4px 8px",color:"#9ca3af",fontStyle:"italic"}}>... e mais {preview.length-5}</td></tr>}</tbody></table></div>
      </div>}
      <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
        <Btn onClick={onClose}>Cancelar</Btn>
        <Btn primary onClick={()=>{if(!preview)return;onImport({records:preview,competencia,empresa,tipo,mode,note:note||(mode==="replace"?"Substituição":"Adição")});onClose();}} disabled={!preview}>{mode==="replace"?"⚠ Confirmar substituição":"✓ Confirmar importação"}</Btn>
      </div>
    </Modal>
  );
}

// ─── EXPORT (admin) ──────────────────────────────────────────────────────────

function ExportModal({ records, onClose }) {
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
    const a=document.createElement("a");a.href=URL.createObjectURL(new Blob([csv],{type:"text/csv;charset=utf-8;"}));a.download=`FCamara_Billing_${[empresa,analista,comp].filter(v=>v!=="todas"&&v!=="todos").join("_")||"Tudo"}.csv`;a.click();onClose();
  }
  return(
    <Modal title="Exportar CSV" onClose={onClose}>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:16}}>
        <div><label style={{fontSize:12,color:"#6b7280",display:"block",marginBottom:4}}>Empresa</label><select style={inp} value={empresa} onChange={e=>setE(e.target.value)}><option value="todas">Todas</option>{EMPRESAS.map(e=><option key={e.cod} value={e.cod}>{e.cod} — {e.nome}</option>)}</select></div>
        <div><label style={{fontSize:12,color:"#6b7280",display:"block",marginBottom:4}}>Analista</label><select style={inp} value={analista} onChange={e=>setA(e.target.value)}><option value="todos">Todos</option>{analistas.map(a=><option key={a}>{a}</option>)}</select></div>
        <div><label style={{fontSize:12,color:"#6b7280",display:"block",marginBottom:4}}>Competência</label><select style={inp} value={comp} onChange={e=>setC(e.target.value)}><option value="todas">Todas</option>{comps.map(c=><option key={c}>{c}</option>)}</select></div>
        <div style={{display:"flex",alignItems:"flex-end",paddingBottom:2}}><label style={{display:"flex",alignItems:"center",gap:8,fontSize:13,cursor:"pointer"}}><input type="checkbox" checked={soNaoFat} onChange={e=>setSN(e.target.checked)} style={{width:16,height:16}}/>Somente não faturado</label></div>
      </div>
      <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}><Btn onClick={onClose}>Cancelar</Btn><Btn primary onClick={doExport}>⬇ Exportar CSV</Btn></div>
    </Modal>
  );
}

// ─── HISTORY (admin) ─────────────────────────────────────────────────────────

function HistoryModal({ history, onClose }) {
  return(
    <Modal title="Histórico de importações" onClose={onClose} wide>
      {history.length===0?<p style={{fontSize:13,color:"#9ca3af"}}>Nenhuma importação.</p>:
      <div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
        <thead><tr style={{background:"#f8fafc"}}>{["Data/Hora","Usuário","Competência","Empresa","Tipo","Modo","Registros","Nota"].map(h=><th key={h} style={{padding:"8px 10px",textAlign:"left",borderBottom:"1px solid #e5e7eb",fontWeight:600,color:"#6b7280",whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
        <tbody>{[...history].reverse().map(h=><tr key={h.id} style={{borderBottom:"1px solid #f3f4f6"}}>
          <td style={{padding:"8px 10px",whiteSpace:"nowrap"}}>{fmtDT(h.date)}</td>
          <td style={{padding:"8px 10px"}}><Badge label={h.user} color="purple" small/></td>
          <td style={{padding:"8px 10px"}}>{h.competencia}</td>
          <td style={{padding:"8px 10px"}}>{h.empresa}</td>
          <td style={{padding:"8px 10px"}}>{h.tipo}</td>
          <td style={{padding:"8px 10px"}}><Badge label={h.mode==="replace"?"Substituição":"Adição"} color={h.mode==="replace"?"red":"green"} small/></td>
          <td style={{padding:"8px 10px",fontWeight:700}}>{h.count}</td>
          <td style={{padding:"8px 10px",color:"#6b7280"}}>{h.note}</td>
        </tr>)}</tbody>
      </table></div>}
      <div style={{display:"flex",justifyContent:"flex-end",marginTop:16}}><Btn onClick={onClose}>Fechar</Btn></div>
    </Modal>
  );
}

// ─── BULK TIMELINE MODAL ─────────────────────────────────────────────────────
// Permite atualizar passos de múltiplos profissionais de um cliente de uma vez

function BulkTimelineModal({ cliente, pep, records, onSave, onClose }) {
  // Estado compartilhado dos passos (aplicado a todos os selecionados)
  const [selected, setSelected] = useState(new Set(records.map(r=>r.id)));
  const [sharedProg, setSharedProg] = useState(() => {
    // Inicializa com o estado do primeiro registro selecionado
    return { ...records[0]?.progress } || makeProgress();
  });
  const [nfNumero, setNfNumero] = useState(records[0]?.nfNumero || "");
  const [obs, setObs] = useState("");
  const [error, setError] = useState("");

  const toggleAll = () => setSelected(s => s.size === records.length ? new Set() : new Set(records.map(r=>r.id)));
  const toggle = (id) => setSelected(s => { const n=new Set(s); n.has(id)?n.delete(id):n.add(id); return n; });
  const setVal = (id, val) => setSharedProg(p=>({...p,[id]:val}));

  function handleSave() {
    if (sharedProg.p5_nf && !nfNumero.trim()) { setError("Informe o número da NF para marcar como emitida."); return; }
    if (selected.size === 0) { setError("Selecione ao menos um profissional."); return; }
    const now = nowISO();
    const updated = records.map(r => selected.has(r.id)
      ? { ...r, progress: { ...sharedProg }, nfNumero: sharedProg.p5_nf ? nfNumero.trim() : r.nfNumero, obs: obs || r.obs, updatedAt: now }
      : r
    );
    onSave(updated);
    onClose();
  }

  const groups = STEP_GROUPS;

  return (
    <Modal title={`Atualizar passos — ${cliente}`} onClose={onClose} extraWide>
      <div style={{fontSize:12,color:"#6b7280",marginBottom:14}}>{pep} · {records.length} profissionais</div>

      {/* Seleção de profissionais */}
      <div style={{marginBottom:20}}>
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:8}}>
          <span style={{fontSize:13,fontWeight:700}}>Selecione os profissionais que receberão esta atualização:</span>
          <button onClick={toggleAll} style={{fontSize:11,color:"#1d4ed8",background:"none",border:"none",cursor:"pointer",textDecoration:"underline"}}>
            {selected.size===records.length?"Desmarcar todos":"Selecionar todos"}
          </button>
        </div>
        <div style={{display:"flex",flexWrap:"wrap",gap:8}}>
          {records.map(r=>(
            <label key={r.id} style={{display:"flex",alignItems:"center",gap:6,padding:"6px 12px",borderRadius:8,border:`1.5px solid ${selected.has(r.id)?"#1d4ed8":"#e5e7eb"}`,background:selected.has(r.id)?"#eff6ff":"#f9fafb",cursor:"pointer",fontSize:13}}>
              <input type="checkbox" checked={selected.has(r.id)} onChange={()=>toggle(r.id)} style={{width:14,height:14}}/>
              <span style={{fontWeight:selected.has(r.id)?600:400,color:selected.has(r.id)?"#1d4ed8":"#374151"}}>{r.profissional}</span>
              <Badge label={calcStatus(r.progress)} color={calcStatusColor(r.progress)} small/>
            </label>
          ))}
        </div>
        {selected.size>0&&<div style={{fontSize:11,color:"#6b7280",marginTop:6}}>{selected.size} profissional(is) selecionado(s) — os passos abaixo serão aplicados a todos eles.</div>}
      </div>

      {/* Passos */}
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:12,marginBottom:16}}>
        {groups.map(g=>(
          <div key={g.num} style={{background:"#f9fafb",borderRadius:10,padding:"12px 14px"}}>
            <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:10}}>
              <div style={{width:22,height:22,borderRadius:"50%",background:"#1d4ed8",color:"#fff",display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:700,flexShrink:0}}>{g.num}</div>
              <span style={{fontWeight:700,fontSize:13}}>{STEP_GROUPS[g.num-1].title}</span>
            </div>
            {STEPS.filter(s=>g.steps.includes(s.id)).map(s=>(
              <div key={s.id} style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:8,gap:8}}>
                <span style={{fontSize:12,color:"#374151",flex:1}}>{s.name}</span>
                {s.type==="check"
                  ? <label style={{display:"flex",alignItems:"center",gap:5,cursor:"pointer",fontSize:12,whiteSpace:"nowrap"}}>
                      <input type="checkbox" checked={!!sharedProg[s.id]} onChange={e=>setVal(s.id,e.target.checked)} style={{width:15,height:15}}/>
                      <span style={{color:sharedProg[s.id]?"#166534":"#9ca3af"}}>{sharedProg[s.id]?"✓ Feito":"Pendente"}</span>
                    </label>
                  : <input type="date" value={sharedProg[s.id]||""} onChange={e=>setVal(s.id,e.target.value)} style={{...inp,width:150}}/>
                }
              </div>
            ))}
          </div>
        ))}
      </div>

      {/* NF Número */}
      <div style={{marginBottom:14,padding:"12px 14px",borderRadius:8,background:sharedProg.p5_nf?"#f0fdf4":"#f9fafb",border:`1px solid ${sharedProg.p5_nf?"#86efac":"#e5e7eb"}`}}>
        <label style={{fontSize:13,fontWeight:700,display:"block",marginBottom:6}}>
          Número da NF {sharedProg.p5_nf&&<span style={{color:"#dc2626"}}>*obrigatório</span>}
        </label>
        <input type="text" value={nfNumero} onChange={e=>{setNfNumero(e.target.value);setError("");}} placeholder="Ex: 123456" style={{...inp,maxWidth:260}}/>
        <div style={{fontSize:11,color:"#6b7280",marginTop:4}}>Será aplicado a todos os profissionais selecionados nesta nota.</div>
      </div>

      {/* Obs */}
      <div style={{marginBottom:16}}>
        <label style={{fontSize:13,fontWeight:700,display:"block",marginBottom:6}}>Observações (opcional)</label>
        <textarea value={obs} onChange={e=>setObs(e.target.value)} placeholder="Observações para todos os selecionados..." style={{...inp,minHeight:60,resize:"vertical"}}/>
      </div>

      {error&&<div style={{marginBottom:12,fontSize:13,padding:"8px 12px",borderRadius:8,background:"#fef2f2",color:"#991b1b",border:"1px solid #fca5a5"}}>{error}</div>}

      <div style={{display:"flex",gap:8,justifyContent:"flex-end"}}>
        <Btn onClick={onClose}>Cancelar</Btn>
        <Btn primary onClick={handleSave} disabled={selected.size===0}>Salvar — {selected.size} profissional(is)</Btn>
      </div>
    </Modal>
  );
}

// ─── MY VIEW (team) ──────────────────────────────────────────────────────────

function MyView({ records, analista, isAdmin, onUpdateBulk, competenciaAtual, onCompetenciaChange }) {
  const [empresa, setEmpresa]       = useState("");
  const [tipo, setTipo]             = useState("");
  const [filterComp, setFilterComp] = useState(competenciaAtual);
  const [filterAnalista, setFA]     = useState("todos");
  const [searchCliente, setSC]      = useState("");
  const [searchProf, setSP]         = useState("");
  const [expandedCliente, setExp]   = useState(null);
  const [bulkTarget, setBulk]       = useState(null); // {cliente, pep, records}

  const myRecords = isAdmin ? records : records.filter(r=>r.responsavel===analista);
  const competencias = [...new Set(records.map(r=>r.competencia))].sort();
  const analistas = [...new Set(records.map(r=>r.responsavel))].sort();
  const empresasUsed = [...new Set(myRecords.map(r=>r.empresa))];
  const tiposUsed = empresa ? [...new Set(myRecords.filter(r=>r.empresa===empresa).map(r=>r.tipo))] : [];

  let filtered = myRecords;
  if (empresa) filtered = filtered.filter(r=>r.empresa===empresa);
  if (tipo)    filtered = filtered.filter(r=>r.tipo===tipo);
  if (filterComp!=="todas") filtered = filtered.filter(r=>r.competencia===filterComp);
  if (isAdmin && filterAnalista!=="todos") filtered = filtered.filter(r=>r.responsavel===filterAnalista);
  if (searchCliente) filtered = filtered.filter(r=>r.cliente.toLowerCase().includes(searchCliente.toLowerCase()));
  if (searchProf)    filtered = filtered.filter(r=>r.profissional.toLowerCase().includes(searchProf.toLowerCase()));

  // Agrupar por cliente
  const grouped = {};
  filtered.forEach(r=>{
    const key = r.cliente+"|"+r.pep;
    if(!grouped[key]) grouped[key]={ cliente:r.cliente, pep:r.pep, records:[] };
    grouped[key].records.push(r);
  });

  return (
    <div>
      {bulkTarget&&<BulkTimelineModal {...bulkTarget} onClose={()=>setBulk(null)} onSave={updated=>{onUpdateBulk(updated);setBulk(null);}}/>}

      {/* Filtros */}
      <div style={{display:"flex",gap:8,marginBottom:14,flexWrap:"wrap",alignItems:"center"}}>
        <select style={{...inp,width:"auto",flex:1}} value={filterComp} onChange={e=>{setFilterComp(e.target.value);onCompetenciaChange(e.target.value);}}>
          {competencias.map(c=><option key={c} value={c}>{c}</option>)}
          <option value="todas">Todas as competências</option>
        </select>
        <select style={{...inp,width:"auto",flex:1}} value={empresa} onChange={e=>{setEmpresa(e.target.value);setTipo("");}}>
          <option value="">Todas as empresas</option>
          {(isAdmin?EMPRESAS:EMPRESAS.filter(e=>empresasUsed.includes(e.cod))).map(e=><option key={e.cod} value={e.cod}>{e.cod} — {e.nome}</option>)}
        </select>
        {empresa&&<select style={{...inp,width:"auto"}} value={tipo} onChange={e=>setTipo(e.target.value)}>
          <option value="">Todos os tipos</option>
          {tiposUsed.map(t=><option key={t}>{t}</option>)}
        </select>}
        {isAdmin&&<select style={{...inp,width:"auto"}} value={filterAnalista} onChange={e=>setFA(e.target.value)}>
          <option value="todos">Todos os analistas</option>
          {analistas.map(a=><option key={a}>{a}</option>)}
        </select>}
        <input style={{...inp,width:160}} placeholder="Filtrar cliente..." value={searchCliente} onChange={e=>setSC(e.target.value)}/>
        <input style={{...inp,width:160}} placeholder="Filtrar profissional..." value={searchProf} onChange={e=>setSP(e.target.value)}/>
      </div>

      {Object.values(grouped).length===0&&<div style={{textAlign:"center",padding:"3rem",color:"#9ca3af"}}>
        <div style={{fontSize:32,marginBottom:10}}>📭</div>
        <div style={{fontSize:14}}>Nenhum registro encontrado para os filtros selecionados.</div>
      </div>}

      {/* Cards de clientes */}
      {Object.values(grouped).map(g=>{
        const total   = g.records.reduce((a,r)=>a+(r.valorTotal||0),0);
        const faturados = g.records.filter(r=>r.progress?.p5_nf).length;
        const pct     = Math.round((faturados/g.records.length)*100)||0;
        const isOpen  = expandedCliente===(g.cliente+g.pep);
        const overallStatus = g.records.every(r=>r.progress?.p5_no_corte)?"Faturado no corte":g.records.every(r=>r.progress?.p5_nf)?"NF emitida":g.records.some(r=>r.progress?.p5_nf)?"Parcialmente faturado":"Em andamento";
        const overallColor  = g.records.every(r=>r.progress?.p5_no_corte)?"green":g.records.every(r=>r.progress?.p5_nf)?"teal":g.records.some(r=>r.progress?.p5_nf)?"blue":"yellow";

        return (
          <div key={g.cliente+g.pep} style={{marginBottom:10,background:"#fff",border:"1px solid #e5e7eb",borderRadius:12,overflow:"hidden"}}>
            {/* Cabeçalho do cliente — clicável */}
            <div onClick={()=>setExp(isOpen?null:(g.cliente+g.pep))} style={{display:"flex",alignItems:"center",gap:12,padding:"14px 18px",cursor:"pointer",userSelect:"none"}}>
              <div style={{flex:1}}>
                <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:4}}>
                  <span style={{fontSize:14,fontWeight:700}}>🏦 {g.cliente}</span>
                  <Badge label={overallStatus} color={overallColor} small/>
                </div>
                <div style={{fontSize:11,color:"#9ca3af"}}>{g.pep} · {g.records.length} profissionais · {fmtShort(total)}</div>
              </div>
              <div style={{textAlign:"center",marginRight:8}}>
                <div style={{fontSize:18,fontWeight:700,color:pct===100?"#16a34a":pct>50?"#2563eb":"#d97706"}}>{pct}%</div>
                <div style={{fontSize:10,color:"#9ca3af"}}>faturado</div>
              </div>
              <div style={{width:80}}>
                <div style={{height:6,background:"#f3f4f6",borderRadius:3}}>
                  <div style={{height:6,borderRadius:3,width:`${pct}%`,background:pct===100?"#16a34a":pct>50?"#2563eb":"#d97706"}}/>
                </div>
              </div>
              <Btn small onClick={e=>{e.stopPropagation();setBulk({cliente:g.cliente,pep:g.pep,records:g.records});}} style={{marginLeft:4}}>Atualizar passos</Btn>
              <span style={{fontSize:18,color:"#9ca3af",marginLeft:4}}>{isOpen?"▲":"▼"}</span>
            </div>

            {/* Detalhe dos profissionais */}
            {isOpen&&<div style={{borderTop:"1px solid #f3f4f6",padding:"0 18px 14px"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,marginTop:10}}>
                <thead><tr style={{background:"#f8fafc"}}>
                  {[isAdmin&&"Analista","Profissional","Período","Val. Venda","Hrs","Val. Total","NF","Status","Atualizado"].filter(Boolean).map(h=>
                    <th key={h} style={{padding:"7px 10px",textAlign:"left",borderBottom:"1px solid #e5e7eb",fontWeight:600,color:"#6b7280",whiteSpace:"nowrap"}}>{h}</th>
                  )}
                </tr></thead>
                <tbody>
                  {g.records.map(r=>(
                    <tr key={r.id} style={{borderBottom:"1px solid #f3f4f6"}}>
                      {isAdmin&&<td style={{padding:"7px 10px"}}><Badge label={r.responsavel} color="purple" small/></td>}
                      <td style={{padding:"7px 10px",fontWeight:500}}>{r.profissional}</td>
                      <td style={{padding:"7px 10px",color:"#6b7280",whiteSpace:"nowrap"}}>{r.inicio} → {r.fim}</td>
                      <td style={{padding:"7px 10px"}}>{fmtShort(r.valorVenda)}</td>
                      <td style={{padding:"7px 10px"}}>{r.hrsAprovadas}h</td>
                      <td style={{padding:"7px 10px",fontWeight:500}}>{fmtShort(r.valorTotal)}</td>
                      <td style={{padding:"7px 10px",fontFamily:"monospace",fontSize:11}}>{r.nfNumero||"—"}</td>
                      <td style={{padding:"7px 10px"}}><Badge label={calcStatus(r.progress)} color={calcStatusColor(r.progress)} small/></td>
                      <td style={{padding:"7px 10px",color:"#9ca3af",fontSize:11,whiteSpace:"nowrap"}}>{r.updatedAt?fmtDT(r.updatedAt):"—"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>}
          </div>
        );
      })}
    </div>
  );
}

// ─── DASHBOARD ───────────────────────────────────────────────────────────────

function Dashboard({ records, analista, isAdmin }) {
  const [filterEmpresa, setFE] = useState("todas");
  const [filterComp,    setFC] = useState("todas");
  const [filterAnalista,setFA] = useState(isAdmin?"todos":analista);
  const [filterEtapa,   setFEt]= useState("todas");

  const comps     = [...new Set(records.map(r=>r.competencia))].sort();
  const analistas = [...new Set(records.map(r=>r.responsavel))].sort();

  // Para não-admin, restringe sempre aos próprios clientes
  let base = isAdmin ? records : records.filter(r=>r.responsavel===analista);
  let f = base;
  if (filterEmpresa!=="todas") f=f.filter(r=>r.empresa===filterEmpresa);
  if (filterComp!=="todas")    f=f.filter(r=>r.competencia===filterComp);
  if (isAdmin&&filterAnalista!=="todos") f=f.filter(r=>r.responsavel===filterAnalista);
  if (filterEtapa!=="todas")   f=f.filter(r=>calcStatus(r.progress)===filterEtapa);

  const totalValor = f.reduce((a,r)=>a+(r.valorTotal||0),0);
  const faturados  = f.filter(r=>r.progress?.p5_nf);
  const naoFat     = f.filter(r=>!r.progress?.p5_nf);
  const valorFat   = faturados.reduce((a,r)=>a+(r.valorTotal||0),0);
  const valorRep   = naoFat.reduce((a,r)=>a+(r.valorTotal||0),0);

  // Valores por etapa
  const byEtapa = {};
  STATUS_ORDER.forEach(s=>{ byEtapa[s]={ count:0, valor:0 }; });
  f.forEach(r=>{ const s=calcStatus(r.progress); if(byEtapa[s]){byEtapa[s].count++;byEtapa[s].valor+=(r.valorTotal||0);} });

  // Por analista
  const byAnalista = {};
  f.forEach(r=>{
    if(!byAnalista[r.responsavel]) byAnalista[r.responsavel]={ total:0, fat:0, rep:0, cnt:0, fatCnt:0 };
    byAnalista[r.responsavel].total+=(r.valorTotal||0); byAnalista[r.responsavel].cnt++;
    if(r.progress?.p5_nf){byAnalista[r.responsavel].fat+=(r.valorTotal||0);byAnalista[r.responsavel].fatCnt++;}
    else byAnalista[r.responsavel].rep+=(r.valorTotal||0);
  });

  // Por empresa
  const byEmpresa = {};
  f.forEach(r=>{ if(!byEmpresa[r.empresa])byEmpresa[r.empresa]={total:0,fat:0}; byEmpresa[r.empresa].total+=(r.valorTotal||0); if(r.progress?.p5_nf)byEmpresa[r.empresa].fat+=(r.valorTotal||0); });

  // Não faturados agrupados por cliente
  const naoFatByCliente = {};
  naoFat.forEach(r=>{
    const key=r.cliente+"|"+r.pep;
    if(!naoFatByCliente[key]) naoFatByCliente[key]={ cliente:r.cliente, pep:r.pep, responsavel:r.responsavel, count:0, valor:0, status:calcStatus(r.progress), color:calcStatusColor(r.progress) };
    naoFatByCliente[key].count++; naoFatByCliente[key].valor+=(r.valorTotal||0);
  });

  const etapaColors = { "Faturado no corte":"green","NF emitida":"teal","Cliente aprovou":"blue","Aguard. aprovação cliente":"yellow","Retorno comercial recebido":"yellow","Aguard. retorno comercial":"orange","Racional montado":"orange","Dados extraídos":"orange","Não iniciado":"gray" };
  const MetCard=({label,value,color,sub})=><div style={{background:"#fff",border:"1px solid #e5e7eb",borderRadius:10,padding:"14px 16px"}}><div style={{fontSize:11,color:"#9ca3af",marginBottom:4}}>{label}</div><div style={{fontSize:19,fontWeight:700,color:color||"#111827"}}>{value}</div>{sub&&<div style={{fontSize:11,color:"#9ca3af",marginTop:2}}>{sub}</div>}</div>;

  return (
    <div>
      {/* Filtros */}
      <div style={{display:"flex",gap:8,marginBottom:20,flexWrap:"wrap"}}>
        <select style={{...inp,width:"auto"}} value={filterComp} onChange={e=>setFC(e.target.value)}><option value="todas">Todas as competências</option>{comps.map(c=><option key={c}>{c}</option>)}</select>
        <select style={{...inp,width:"auto"}} value={filterEmpresa} onChange={e=>setFE(e.target.value)}><option value="todas">Todas as empresas</option>{EMPRESAS.map(e=><option key={e.cod} value={e.cod}>{e.cod} — {e.nome}</option>)}</select>
        {isAdmin&&<select style={{...inp,width:"auto"}} value={filterAnalista} onChange={e=>setFA(e.target.value)}><option value="todos">Todos os analistas</option>{analistas.map(a=><option key={a}>{a}</option>)}</select>}
        <select style={{...inp,width:"auto"}} value={filterEtapa} onChange={e=>setFEt(e.target.value)}><option value="todas">Todas as etapas</option>{STATUS_ORDER.map(s=><option key={s}>{s}</option>)}</select>
      </div>

      {/* Métricas */}
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(140px,1fr))",gap:10,marginBottom:24}}>
        <MetCard label="Total registros" value={f.length}/>
        <MetCard label="Valor total" value={fmtShort(totalValor)} color="#1d4ed8"/>
        <MetCard label="Faturado" value={fmtShort(valorFat)} color="#166534" sub={`${faturados.length} registros`}/>
        <MetCard label="Represado" value={fmtShort(valorRep)} color="#854d0e" sub={`${naoFat.length} registros`}/>
      </div>

      {/* Valores por etapa */}
      <div style={{background:"#fff",border:"1px solid #e5e7eb",borderRadius:10,padding:16,marginBottom:16}}>
        <div style={{fontWeight:700,fontSize:14,marginBottom:12}}>Valor por etapa</div>
        <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(180px,1fr))",gap:8}}>
          {STATUS_ORDER.filter(s=>byEtapa[s]?.count>0).map(s=>{
            const d=byEtapa[s]; const c=C[etapaColors[s]]||C.gray;
            return <div key={s} style={{padding:"10px 12px",borderRadius:8,background:c.bg,border:`1px solid ${c.border}`}}>
              <div style={{fontSize:11,color:c.text,fontWeight:600,marginBottom:4}}>{s}</div>
              <div style={{fontSize:16,fontWeight:700,color:c.text}}>{fmtShort(d.valor)}</div>
              <div style={{fontSize:11,color:c.text,opacity:.7}}>{d.count} registro(s)</div>
            </div>;
          })}
        </div>
      </div>

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16,marginBottom:16}}>
        {/* Por analista — só para admin */}
        {isAdmin&&<div style={{background:"#fff",border:"1px solid #e5e7eb",borderRadius:10,padding:16}}>
          <div style={{fontWeight:700,fontSize:14,marginBottom:12}}>Por analista</div>
          {Object.entries(byAnalista).map(([a,d])=>{
            const pct=d.total>0?Math.round((d.fat/d.total)*100):0;
            const bar=pct===100?"#16a34a":pct>50?"#2563eb":"#d97706";
            return <div key={a} style={{marginBottom:14}}>
              <div style={{display:"flex",justifyContent:"space-between",fontSize:12,marginBottom:3}}><span style={{fontWeight:600}}>{a}</span><span style={{color:"#6b7280"}}>{pct}% · {d.cnt} reg.</span></div>
              <div style={{height:6,background:"#f3f4f6",borderRadius:3}}><div style={{height:6,borderRadius:3,width:`${pct}%`,background:bar}}/></div>
              <div style={{display:"flex",justifyContent:"space-between",fontSize:11,color:"#9ca3af",marginTop:3}}><span>Fat: {fmtShort(d.fat)}</span><span>Rep: {fmtShort(d.rep)}</span></div>
            </div>;
          })}
        </div>}

        {/* Por empresa */}
        <div style={{background:"#fff",border:"1px solid #e5e7eb",borderRadius:10,padding:16}}>
          <div style={{fontWeight:700,fontSize:14,marginBottom:12}}>Por empresa</div>
          {Object.entries(byEmpresa).map(([cod,d])=>{
            const emp=EMPRESAS.find(e=>e.cod===cod);
            const pct=d.total>0?Math.round((d.fat/d.total)*100):0;
            return <div key={cod} style={{marginBottom:12}}>
              <div style={{display:"flex",justifyContent:"space-between",fontSize:12,marginBottom:3}}><span style={{fontWeight:600}}>{cod} — {emp?.nome}</span><span style={{color:"#6b7280"}}>{fmtShort(d.total)}</span></div>
              <div style={{height:6,background:"#f3f4f6",borderRadius:3}}><div style={{height:6,borderRadius:3,width:`${pct}%`,background:"#1d4ed8"}}/></div>
            </div>;
          })}
        </div>
      </div>

      {/* Não faturados por cliente */}
      <div style={{background:"#fff",border:"1px solid #e5e7eb",borderRadius:10,padding:16}}>
        <div style={{fontWeight:700,fontSize:14,marginBottom:12}}>Não faturados — resumo por cliente ({Object.keys(naoFatByCliente).length})</div>
        {Object.keys(naoFatByCliente).length===0
          ?<div style={{textAlign:"center",padding:"1rem",color:"#9ca3af",fontSize:13}}>🎉 Tudo faturado!</div>
          :<div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
            <thead><tr style={{background:"#f8fafc"}}>{[isAdmin&&"Analista","Cliente","PEP","Profissionais","Val. Total","Etapa atual"].filter(Boolean).map(h=><th key={h} style={{padding:"8px 10px",textAlign:"left",borderBottom:"1px solid #e5e7eb",fontWeight:600,color:"#6b7280",whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
            <tbody>{Object.values(naoFatByCliente).map((d,i)=><tr key={i} style={{borderBottom:"1px solid #f3f4f6"}}>
              {isAdmin&&<td style={{padding:"7px 10px"}}>{d.responsavel}</td>}
              <td style={{padding:"7px 10px",fontWeight:500}}>{d.cliente}</td>
              <td style={{padding:"7px 10px",fontFamily:"monospace",fontSize:11}}>{d.pep}</td>
              <td style={{padding:"7px 10px",textAlign:"center"}}>{d.count}</td>
              <td style={{padding:"7px 10px",fontWeight:600}}>{fmtShort(d.valor)}</td>
              <td style={{padding:"7px 10px"}}><Badge label={d.status} color={d.color} small/></td>
            </tr>)}</tbody>
          </table></div>
        }
      </div>
    </div>
  );
}

// ─── APP ROOT ─────────────────────────────────────────────────────────────────

export default function App() {
  const [state, setState]       = useState(()=>loadState());
  const [user, setUser]         = useState(null); // { name, isAdmin }
  const [loginName, setLN]      = useState("");
  const [loginPass, setLP]      = useState("");
  const [loginErr, setLE]       = useState("");
  const [activeTab, setTab]     = useState("time");
  const [showImport, setImp]    = useState(false);
  const [showExport, setExp]    = useState(false);
  const [showHistory, setHist]  = useState(false);

  useEffect(()=>saveState(state),[state]);

  const isAdmin = user?.isAdmin || false;

  function handleLogin(e) {
    e.preventDefault();
    const found = USERS.find(u=>u.name.toLowerCase()===loginName.trim().toLowerCase()&&u.password===loginPass.trim().toLowerCase());
    if (!found) { setLE("Nome ou senha incorretos."); return; }
    setUser({ name: found.name, isAdmin: found.isAdmin });
    setLE("");
  }

  function handleUpdate(updated) {
    setState(s=>({...s, records:s.records.map(r=>r.id===updated.id?updated:r)}));
  }

  function handleUpdateBulk(updatedList) {
    const map = Object.fromEntries(updatedList.map(r=>[r.id,r]));
    setState(s=>({...s, records:s.records.map(r=>map[r.id]||r)}));
  }

  function handleImport({ records:newRecs, competencia, empresa, tipo, mode, note }) {
    setState(s=>{
      let base = mode==="replace" ? s.records.filter(r=>!(r.competencia===competencia&&r.empresa===empresa&&r.tipo===tipo)) : s.records;
      const entry={ id:genId(), date:nowISO(), competencia, empresa, tipo, mode, count:newRecs.length, user:ADMIN_NAME, note };
      return {...s, records:[...base,...newRecs], competenciaAtual:competencia, importHistory:[...s.importHistory,entry]};
    });
  }

  function handleCompetencia(val) {
    setState(s=>({...s, competenciaAtual:val}));
  }

  // ── LOGIN ──
  if (!user) {
    return (
      <div style={{minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center",background:"linear-gradient(135deg,#1e3a8a,#1d4ed8)",fontFamily:"system-ui,sans-serif"}}>
        <div style={{background:"#fff",borderRadius:16,padding:"36px 40px",width:380,maxWidth:"92vw",boxShadow:"0 16px 48px rgba(0,0,0,.2)"}}>
          <div style={{textAlign:"center",marginBottom:28}}>
            <div style={{fontSize:36,marginBottom:10}}>📊</div>
            <h1 style={{fontSize:18,fontWeight:800,color:"#1d4ed8",lineHeight:1.3,margin:0}}>Controle de Faturamento<br/>Grupo Fcamara</h1>
            <p style={{fontSize:12,color:"#9ca3af",marginTop:6}}>{APP_VERSION}</p>
          </div>
          <form onSubmit={handleLogin}>
            <div style={{marginBottom:12}}>
              <label style={{fontSize:12,color:"#6b7280",display:"block",marginBottom:4}}>Nome</label>
              <input style={inp} placeholder="Seu nome" value={loginName} onChange={e=>{setLN(e.target.value);setLE("");}} autoFocus/>
            </div>
            <div style={{marginBottom:16}}>
              <label style={{fontSize:12,color:"#6b7280",display:"block",marginBottom:4}}>Senha</label>
              <input style={inp} type="password" placeholder="Sua senha" value={loginPass} onChange={e=>{setLP(e.target.value);setLE("");}}/>
            </div>
            {loginErr&&<div style={{marginBottom:12,fontSize:12,padding:"8px 12px",borderRadius:8,background:"#fef2f2",color:"#991b1b",border:"1px solid #fca5a5"}}>{loginErr}</div>}
            <button type="submit" style={{width:"100%",padding:"10px",borderRadius:8,border:"none",background:"#1d4ed8",color:"#fff",fontSize:14,fontWeight:700,cursor:"pointer"}}>Entrar</button>
          </form>
          <p style={{fontSize:11,color:"#d1d5db",textAlign:"center",marginTop:16}}>A senha padrão é o seu nome em letras minúsculas.</p>
        </div>
      </div>
    );
  }

  // ── APP ──
  return (
    <div style={{fontFamily:"system-ui,-apple-system,sans-serif",color:"#111827",minHeight:"100vh",background:"#f8fafc"}}>
      {showImport  && <ImportModal onImport={handleImport} onClose={()=>setImp(false)}/>}
      {showExport  && <ExportModal records={state.records} onClose={()=>setExp(false)}/>}
      {showHistory && <HistoryModal history={state.importHistory} onClose={()=>setHist(false)}/>}

      {/* Topbar */}
      <div style={{background:"#1d4ed8",color:"#fff",padding:"0 20px",display:"flex",alignItems:"center",gap:8,height:52}}>
        <span style={{fontSize:14,fontWeight:800,flex:1}}>📊 Faturamento Grupo Fcamara</span>
        <span style={{fontSize:11,opacity:.8}}>
          {user.name}{isAdmin?" · Admin":""}
        </span>
        {isAdmin&&<>
          <button onClick={()=>setImp(true)}  style={{background:"rgba(255,255,255,.15)",border:"none",color:"#fff",borderRadius:6,padding:"4px 10px",fontSize:11,cursor:"pointer"}}>⬆ Importar</button>
          <button onClick={()=>setExp(true)}  style={{background:"rgba(255,255,255,.15)",border:"none",color:"#fff",borderRadius:6,padding:"4px 10px",fontSize:11,cursor:"pointer"}}>⬇ Exportar</button>
          <button onClick={()=>setHist(true)} style={{background:"rgba(255,255,255,.15)",border:"none",color:"#fff",borderRadius:6,padding:"4px 10px",fontSize:11,cursor:"pointer"}}>🕐 Histórico</button>
        </>}
        <button onClick={()=>setUser(null)} style={{background:"rgba(255,255,255,.12)",border:"none",color:"#fff",borderRadius:6,padding:"4px 10px",fontSize:11,cursor:"pointer"}}>Sair</button>
      </div>

      {isAdmin&&<div style={{background:"#fffbeb",borderBottom:"1px solid #fde68a",padding:"6px 20px",fontSize:12,color:"#92400e"}}>
        Admin — acesso completo a todos os analistas, empresas e competências.
      </div>}

      {/* Tabs */}
      <div style={{background:"#fff",borderBottom:"1px solid #e5e7eb",padding:"0 20px",display:"flex",gap:4}}>
        {[{id:"time",label:"Minha visão"},{id:"dash",label:"Dashboard"}].map(t=>(
          <button key={t.id} onClick={()=>setTab(t.id)} style={{padding:"14px 18px",border:"none",background:"transparent",cursor:"pointer",fontSize:13,fontWeight:activeTab===t.id?700:400,color:activeTab===t.id?"#1d4ed8":"#6b7280",borderBottom:activeTab===t.id?"2px solid #1d4ed8":"2px solid transparent"}}>
            {t.label}
          </button>
        ))}
      </div>

      <div style={{maxWidth:1100,margin:"0 auto",padding:"24px 20px"}}>
        {activeTab==="time"&&<MyView records={state.records} analista={user.name} isAdmin={isAdmin} onUpdateBulk={handleUpdateBulk} competenciaAtual={state.competenciaAtual} onCompetenciaChange={handleCompetencia}/>}
        {activeTab==="dash"&&<Dashboard records={state.records} analista={user.name} isAdmin={isAdmin}/>}
      </div>
    </div>
  );
}
