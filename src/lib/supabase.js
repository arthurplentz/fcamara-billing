import { createClient } from "@supabase/supabase-js";

// Credenciais PÚBLICAS do projeto Supabase.
// A "publishable key" é feita para ir no frontend — os dados ficam protegidos
// pelas regras de segurança (RLS) do banco, não pela chave. Pode ficar no código.
// Em produção/CI é possível sobrescrever via variáveis de ambiente (Vite).
const SUPABASE_URL =
  import.meta.env.VITE_SUPABASE_URL || "https://lkrmzdrqlvppadompjmf.supabase.co";
const SUPABASE_KEY =
  import.meta.env.VITE_SUPABASE_ANON_KEY ||
  "sb_publishable_0BfGhFQWcPdaWFGMqqfpww_S2KRhrLb";

export const supabase = createClient(SUPABASE_URL, SUPABASE_KEY, {
  auth: { persistSession: true, autoRefreshToken: true, detectSessionInUrl: true },
});

export const SITE_URL =
  typeof window !== "undefined" ? window.location.origin + window.location.pathname : "/";
