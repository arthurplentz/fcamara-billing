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
    importId: row.import_id || null,
    municipalNoteId: row.municipal_note_id || null,
    conciliadoEm: row.conciliado_em || null, conciliadoPor: row.conciliado_por || "",
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
    import_id: r.importId || null,
  };
  if (withId && r.id) o.id = r.id;
  return o;
}

export async function fetchRecords() {
  const rows = await fetchAllPaged("records", "*", "created_at");
  return rows.map(dbToRec);
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
export async function deleteRecordsByImport(importId) {
  const { error, count } = await supabase.from("records").delete({ count: "exact" }).eq("import_id", importId);
  if (error) throw error;
  return count || 0;
}
// Conciliação: amarra registros de receita a uma nota da prefeitura (decisão do
// analista). Atualiza só os campos de conciliação + número da NF — não toca nos
// dados de reconhecimento (o trigger records_guard permite estes campos).
export async function conciliateRecords(recordIds, { noteId, numero, userName }) {
  const { error } = await supabase.from("records")
    .update({ municipal_note_id: noteId, nf_numero: numero || "", conciliado_em: nowISO(), conciliado_por: userName || null, updated_at: nowISO() })
    .in("id", recordIds);
  if (error) throw error;
}
export async function unconciliateRecords(recordIds) {
  const { error } = await supabase.from("records")
    .update({ municipal_note_id: null, nf_numero: "", conciliado_em: null, conciliado_por: null, updated_at: nowISO() })
    .in("id", recordIds);
  if (error) throw error;
}
// Conciliação que também completa o funil de cada registro (progress já pronto
// no app, por registro). Atualiza um a um — o lote por conciliação é pequeno.
export async function conciliateRecordsWithProgress(items, { noteId, numero, userName }) {
  for (const it of items) {
    const { error } = await supabase.from("records").update({
      municipal_note_id: noteId, nf_numero: numero || "", progress: it.progress,
      conciliado_em: nowISO(), conciliado_por: userName || null, updated_at: nowISO(),
    }).eq("id", it.id);
    if (error) throw error;
  }
}

// ─── NOTAS DA PREFEITURA (NFS-e) ─────────────────────────────────────────────
function dbToNote(row) {
  return {
    id: row.id, municipio: row.municipio || "", empresa: row.empresa || "", numero: row.numero || "",
    emitidaEm: row.emitida_em || "", fatoGerador: row.fato_gerador || "",
    prestadorCnpj: row.prestador_cnpj || "", prestadorNome: row.prestador_nome || "",
    tomadorCnpj: row.tomador_cnpj || "", tomadorNome: row.tomador_nome || "",
    valorServicos: Number(row.valor_servicos) || 0, valorTotal: Number(row.valor_total) || 0,
    iss: Number(row.iss) || 0, situacao: row.situacao || "", cancelada: !!row.cancelada,
    pedidos: row.pedidos || "", competencias: row.competencias || "", profissionais: row.profissionais || "",
    discriminacao: row.discriminacao || "", importId: row.import_id || null,
  };
}
function noteToDb(n) {
  return {
    municipio: n.municipio || null, empresa: n.empresa || null, numero: n.numero || null,
    emitida_em: n.emitidaEm || null, fato_gerador: n.fatoGerador || null,
    prestador_cnpj: n.prestadorCnpj || null, prestador_nome: n.prestadorNome || null,
    tomador_cnpj: n.tomadorCnpj || null, tomador_nome: n.tomadorNome || null,
    valor_servicos: n.valorServicos || 0, valor_total: n.valorTotal || 0,
    iss: n.iss || 0, situacao: n.situacao || null, cancelada: !!n.cancelada,
    pedidos: n.pedidos || null, competencias: n.competencias || null, profissionais: n.profissionais || null,
    discriminacao: n.discriminacao || null, import_id: n.importId || null,
  };
}
export async function fetchMunicipalNotes() {
  const rows = await fetchAllPaged("municipal_notes", "*", "emitida_em");
  return rows.map(dbToNote);
}
export async function insertMunicipalNotes(list) {
  const size = 500;
  for (let i = 0; i < list.length; i += size) {
    const chunk = list.slice(i, i + size).map(noteToDb);
    const { error } = await supabase.from("municipal_notes").insert(chunk);
    if (error) throw error;
  }
  return list.length;
}
export async function deleteMunicipalNotesByImport(importId) {
  const { error, count } = await supabase.from("municipal_notes").delete({ count: "exact" }).eq("import_id", importId);
  if (error) throw error;
  return count || 0;
}
export async function deleteMunicipalNote(id) {
  const { error } = await supabase.from("municipal_notes").delete().eq("id", id);
  if (error) throw error;
}

// ─── TASKS ───────────────────────────────────────────────────────────────────
function dbToTask(row) {
  return {
    id: row.id, title: row.title, desc: row.descricao || "",
    dueDate: row.due_date || "", assignee: row.assignee || "",
    status: row.status, createdAt: row.created_at, updatedAt: row.updated_at,
    deliveryId: row.delivery_id || null, recorrente: !!row.recorrente, competencia: row.competencia || "",
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
    importId: row.import_id || null,
  };
}
export async function fetchHistory() {
  const { data, error } = await supabase.from("import_history").select("*").order("date", { ascending: true });
  if (error) throw error;
  return data.map(dbToHist);
}
export async function deleteHistory(id) {
  const { error } = await supabase.from("import_history").delete().eq("id", id);
  if (error) throw error;
}
// ─── ENTREGAS (modelos recorrentes + geração mensal) ─────────────────────────
export async function fetchTemplates() {
  const { data, error } = await supabase.from("delivery_templates").select("*").order("created_at", { ascending: true });
  if (error) throw error;
  return data.map(r => ({ id: r.id, title: r.title, items: Array.isArray(r.items) ? r.items : [] }));
}
export async function insertTemplate(t) {
  const { error } = await supabase.from("delivery_templates").insert({ title: t.title, items: t.items || [] });
  if (error) throw error;
}
export async function updateTemplate(t) {
  const { error } = await supabase.from("delivery_templates").update({ title: t.title, items: t.items || [] }).eq("id", t.id);
  if (error) throw error;
}
export async function deleteTemplate(id) {
  const { error } = await supabase.from("delivery_templates").delete().eq("id", id);
  if (error) throw error;
}
export async function fetchDeliveries() {
  const { data, error } = await supabase.from("deliveries").select("*").order("created_at", { ascending: false });
  if (error) throw error;
  return data.map(r => ({ id: r.id, templateId: r.template_id, title: r.title, competencia: r.competencia }));
}
// Gera uma entrega do mês: cria a delivery e dispara as tarefas para os analistas.
// Cada item pode ter: dia (do mês → vira data) e assignees (lista; vazio = todos).
export async function generateDelivery(template, competencia, allAnalysts) {
  const { data: dv, error } = await supabase.from("deliveries")
    .insert({ template_id: template.id, title: `${template.title} — ${competencia}`, competencia })
    .select().single();
  if (error) throw error;
  const [mm, yyyy] = String(competencia || "").split("/");
  const rows = [];
  (template.items || []).forEach(item => {
    const list = (Array.isArray(item.assignees) && item.assignees.length) ? item.assignees
               : (item.assignee ? [item.assignee] : (allAnalysts.length ? allAnalysts : [null]));
    const due = (item.dia && mm && yyyy) ? `${yyyy}-${String(mm).padStart(2,"0")}-${String(item.dia).padStart(2,"0")}` : null;
    list.forEach(a => rows.push({
      title: item.title, descricao: item.desc || "", assignee: a, status: "todo",
      due_date: due, delivery_id: dv.id, recorrente: true, competencia,
    }));
  });
  if (rows.length) { const { error: e2 } = await supabase.from("tasks").insert(rows); if (e2) throw e2; }
  return { delivery: dv, count: rows.length };
}

export async function insertHistory(h) {
  const { error } = await supabase.from("import_history").insert({
    competencia: h.competencia, empresa: h.empresa, tipo: h.tipo,
    mode: h.mode, count: h.count, user_name: h.user, note: h.note, import_id: h.importId || null,
  });
  if (error) throw error;
}

// ─── CLIENTS (perfil de faturamento) ─────────────────────────────────────────
const CLIENT_FIELDS = ["nome","cod_sap","cnpj","cnpjs","grupo_empresa","owner","incompleto","tipos_contrato","tipos_peps","proposta_url","propostas","periodo_faturamento","calendario","tem_portal","portal_tipo","portal_link","portal_usuario","portal_senha","portal_passo_url","prazo_vencimento","forma_pagamento","contato_financeiro","contato_financeiro_email","account_manager","account_manager_email"];
const BOOL_CLIENT_FIELDS = ["tem_portal","incompleto"];

// Busca paginada — traz TODAS as linhas (o Supabase pode limitar a 1000 por página).
async function fetchAllPaged(table, columns, orderCol) {
  const all = []; const size = 1000; let from = 0;
  for (;;) {
    const { data, error } = await supabase.from(table).select(columns).order(orderCol, { ascending: true }).range(from, from + size - 1);
    if (error) throw error;
    all.push(...data);
    if (data.length < size) break;
    from += size;
  }
  return all;
}
const camel = s => s.replace(/_([a-z])/g, (_,c)=>c.toUpperCase());
function dbToClient(row) {
  const o = { id: row.id };
  CLIENT_FIELDS.forEach(f => { o[camel(f)] = BOOL_CLIENT_FIELDS.includes(f) ? !!row[f] : (row[f] ?? ""); });
  return o;
}
function clientToDb(c) {
  const o = {};
  CLIENT_FIELDS.forEach(f => { const v = c[camel(f)]; o[f] = BOOL_CLIENT_FIELDS.includes(f) ? !!v : (v || null); });
  o.updated_at = nowISO();
  return o;
}
export async function fetchClients() {
  const rows = await fetchAllPaged("clients", "*", "nome");
  return rows.map(dbToClient);
}
export async function insertClient(c) {
  const { data, error } = await supabase.from("clients").insert(clientToDb(c)).select().single();
  if (error) throw error;
  return dbToClient(data);
}
// Carga em massa de clientes (importação). Insere em lotes.
export async function bulkInsertClients(list) {
  const size = 500;
  for (let i = 0; i < list.length; i += size) {
    const chunk = list.slice(i, i + size).map(clientToDb);
    const { error } = await supabase.from("clients").insert(chunk);
    if (error) throw error;
  }
  return list.length;
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
  const { data, error } = await supabase.from("profiles").select("id,name,is_admin,responsavel").order("name", { ascending: true });
  if (error) throw error;
  return data.map(p => ({ id: p.id, name: p.name, isAdmin: !!p.is_admin, responsavel: p.responsavel || "" }));
}
export async function updateProfile({ id, name, isAdmin, responsavel }) {
  const patch = {};
  if (name != null) patch.name = name;
  if (isAdmin != null) patch.is_admin = isAdmin;
  if (responsavel !== undefined) patch.responsavel = responsavel || null;
  const { error } = await supabase.from("profiles").update(patch).eq("id", id);
  if (error) throw error;
}
export async function deleteProfile(id) {
  const { error } = await supabase.from("profiles").delete().eq("id", id);
  if (error) throw error;
}
