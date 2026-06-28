// Camada de dados — conversa com o Supabase e traduz entre o formato do banco
// (snake_case) e o formato usado no app (camelCase).
import { supabase } from "./supabase";

const nowISO = () => new Date().toISOString();

// ─── RECORDS ─────────────────────────────────────────────────────────────────
function dbToRec(row) {
  return {
    id: row.id,
    responsavel: row.responsavel, empresa: row.empresa, tipo: row.tipo,
    codCliente: row.cod_cliente || "", cliente: row.cliente, pep: row.pep,
    inicio: row.inicio || "", fim: row.fim || "", profissional: row.profissional || "",
    valorVenda: Number(row.valor_venda) || 0, hrsAprovadas: Number(row.hrs_aprovadas) || 0,
    valorTotal: Number(row.valor_total) || 0, valorLiquido: Number(row.valor_liquido) || 0,
    competencia: row.competencia, progress: row.progress || {},
    nfNumero: row.nf_numero || "", obs: row.obs || "", updatedAt: row.updated_at,
  };
}
function recToDb(r, withId) {
  const o = {
    responsavel: r.responsavel, empresa: r.empresa, tipo: r.tipo,
    cod_cliente: r.codCliente || null, cliente: r.cliente, pep: r.pep,
    inicio: r.inicio || null, fim: r.fim || null, profissional: r.profissional || null,
    valor_venda: r.valorVenda || 0, hrs_aprovadas: r.hrsAprovadas || 0,
    valor_total: r.valorTotal || 0, valor_liquido: r.valorLiquido || 0,
    competencia: r.competencia, progress: r.progress || {},
    nf_numero: r.nfNumero || "", obs: r.obs || "", updated_at: r.updatedAt || nowISO(),
  };
  if (withId && r.id) o.id = r.id;
  return o;
}

export async function fetchRecords() {
  const { data, error } = await supabase.from("records").select("*").order("created_at", { ascending: true });
  if (error) throw error;
  return data.map(dbToRec);
}
export async function insertRecords(list) {
  const { error } = await supabase.from("records").insert(list.map(r => recToDb(r, false)));
  if (error) throw error;
}
export async function upsertRecords(list) {
  const { error } = await supabase.from("records").upsert(list.map(r => recToDb(r, true)));
  if (error) throw error;
}
export async function deleteRecordsBy({ competencia, empresa, tipo }) {
  const { error } = await supabase.from("records").delete()
    .eq("competencia", competencia).eq("empresa", empresa).eq("tipo", tipo);
  if (error) throw error;
}
export async function deleteRecord(id) {
  const { error } = await supabase.from("records").delete().eq("id", id);
  if (error) throw error;
}

// ─── TASKS ───────────────────────────────────────────────────────────────────
function dbToTask(row) {
  return {
    id: row.id, title: row.title, desc: row.descricao || "",
    dueDate: row.due_date || "", assignee: row.assignee || "",
    status: row.status, createdAt: row.created_at, updatedAt: row.updated_at,
  };
}
function taskToDb(t) {
  return {
    title: t.title, descricao: t.desc || "", due_date: t.dueDate || null,
    assignee: t.assignee || null, status: t.status || "inbox", updated_at: nowISO(),
  };
}
export async function fetchTasks() {
  const { data, error } = await supabase.from("tasks").select("*").order("created_at", { ascending: true });
  if (error) throw error;
  return data.map(dbToTask);
}
export async function insertTask(t)  { const { error } = await supabase.from("tasks").insert(taskToDb(t));        if (error) throw error; }
export async function updateTask(t)  { const { error } = await supabase.from("tasks").update(taskToDb(t)).eq("id", t.id); if (error) throw error; }
export async function deleteTask(id) { const { error } = await supabase.from("tasks").delete().eq("id", id);      if (error) throw error; }

// ─── IMPORT HISTORY ──────────────────────────────────────────────────────────
function dbToHist(row) {
  return {
    id: row.id, date: row.date, competencia: row.competencia, empresa: row.empresa,
    tipo: row.tipo, mode: row.mode, count: row.count, user: row.user_name, note: row.note,
  };
}
export async function fetchHistory() {
  const { data, error } = await supabase.from("import_history").select("*").order("date", { ascending: true });
  if (error) throw error;
  return data.map(dbToHist);
}
export async function insertHistory(h) {
  const { error } = await supabase.from("import_history").insert({
    competencia: h.competencia, empresa: h.empresa, tipo: h.tipo,
    mode: h.mode, count: h.count, user_name: h.user, note: h.note,
  });
  if (error) throw error;
}

// ─── CLIENTS (perfil de faturamento) ─────────────────────────────────────────
const CLIENT_FIELDS = ["nome","cod_sap","grupo_empresa","tipos_contrato","tipos_peps","proposta_url","propostas","periodo_faturamento","calendario","tem_portal","portal_tipo","portal_link","portal_usuario","portal_senha","portal_passo_url","prazo_vencimento","forma_pagamento","contato_financeiro","contato_financeiro_email","account_manager","account_manager_email"];
const camel = s => s.replace(/_([a-z])/g, (_,c)=>c.toUpperCase());
function dbToClient(row) {
  const o = { id: row.id };
  CLIENT_FIELDS.forEach(f => { o[camel(f)] = f==="tem_portal" ? !!row[f] : (row[f] ?? ""); });
  return o;
}
function clientToDb(c) {
  const o = {};
  CLIENT_FIELDS.forEach(f => { const v = c[camel(f)]; o[f] = f==="tem_portal" ? !!v : (v || null); });
  o.updated_at = nowISO();
  return o;
}
export async function fetchClients() {
  const { data, error } = await supabase.from("clients").select("*").order("nome", { ascending: true });
  if (error) throw error;
  return data.map(dbToClient);
}
export async function insertClient(c) {
  const { data, error } = await supabase.from("clients").insert(clientToDb(c)).select().single();
  if (error) throw error;
  return dbToClient(data);
}
export async function updateClient(c) {
  const { error } = await supabase.from("clients").update(clientToDb(c)).eq("id", c.id);
  if (error) throw error;
}
export async function deleteClient(id) {
  const { error } = await supabase.from("clients").delete().eq("id", id);
  if (error) throw error;
}

// ─── PROFILES (gestão de acessos) ────────────────────────────────────────────
export async function fetchProfiles() {
  const { data, error } = await supabase.from("profiles").select("id,name,is_admin").order("name", { ascending: true });
  if (error) throw error;
  return data.map(p => ({ id: p.id, name: p.name, isAdmin: !!p.is_admin }));
}
export async function updateProfile({ id, name, isAdmin }) {
  const patch = {};
  if (name != null) patch.name = name;
  if (isAdmin != null) patch.is_admin = isAdmin;
  const { error } = await supabase.from("profiles").update(patch).eq("id", id);
  if (error) throw error;
}
export async function deleteProfile(id) {
  const { error } = await supabase.from("profiles").delete().eq("id", id);
  if (error) throw error;
}
