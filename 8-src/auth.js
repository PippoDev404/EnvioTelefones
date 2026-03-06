import { supabase } from "./supabaseClient.js";

export async function requireAuth({ redirectTo = "../7-login/login.html" } = {}) {
  const { data } = await supabase.auth.getSession();

  if (!data.session) {
    window.location.href = redirectTo;
    return null;
  }

  return data.session.user; // tem user.id
}

export async function logout({ redirectTo = "../7-login/login.html" } = {}) {
  await supabase.auth.signOut();
  window.location.href = redirectTo;
}