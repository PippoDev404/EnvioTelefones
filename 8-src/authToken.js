import { supabase } from "./supabaseClient.js";

export async function getAccessTokenOrThrow() {
  const { data, error } = await supabase.auth.getSession();

  if (error) {
    throw new Error(error.message || "Não foi possível obter a sessão.");
  }

  const token = data?.session?.access_token;
  if (!token) {
    throw new Error("Usuário sem sessão ativa.");
  }

  return token;
}