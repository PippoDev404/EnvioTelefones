import { supabase } from "../8-src/supabaseClient.js";

const form = document.querySelector("#formularioLogin");
const inputEmail = document.querySelector("#emailAdmin");
const inputSenha = document.querySelector("#senhaAdmin");
const caixaAlerta = document.querySelector("#caixaAlerta");
const lembrarEmail = document.querySelector("#lembrarEmail");
const botaoEntrar = document.querySelector("#botaoEntrar");
const textoBotao = botaoEntrar?.querySelector(".textoBotao");
const botaoAlternarSenha = document.querySelector("#botaoAlternarSenha");
const linkEsqueciSenha = document.querySelector("#linkEsqueciSenha");
const anoAtual = document.querySelector("#anoAtual");
const pilulaAmbiente = document.querySelector("#pílulaAmbiente");

const LS_EMAIL_KEY = "ibespe:lembrarEmail";

function setMsg(texto = "", tipo = "info") {
  if (!caixaAlerta) return;
  caixaAlerta.textContent = texto;
  caixaAlerta.dataset.tipo = tipo;
  caixaAlerta.style.display = texto ? "block" : "none";
}

function setLoading(loading) {
  if (!botaoEntrar) return;
  botaoEntrar.disabled = !!loading;
  if (textoBotao) textoBotao.textContent = loading ? "Entrando..." : "Entrar";
}

function preencherAno() {
  if (anoAtual) anoAtual.textContent = String(new Date().getFullYear());
}

function preencherAmbiente() {
  if (!pilulaAmbiente) return;

  const host = window.location.hostname;
  const isLocal =
    host === "localhost" ||
    host === "127.0.0.1" ||
    host.endsWith(".local");

  pilulaAmbiente.textContent = isLocal ? "ENV: local" : "ENV: produção";
}

function restaurarEmail() {
  const salvo = localStorage.getItem(LS_EMAIL_KEY);
  if (salvo && inputEmail) {
    inputEmail.value = salvo;
    if (lembrarEmail) lembrarEmail.checked = true;
  }
}

function salvarPreferenciaEmail(email) {
  if (!lembrarEmail?.checked) {
    localStorage.removeItem(LS_EMAIL_KEY);
    return;
  }
  localStorage.setItem(LS_EMAIL_KEY, email);
}

async function redirecionarSeJaLogado() {
  const { data, error } = await supabase.auth.getUser();
  if (error) return;
  if (data?.user) {
    window.location.href = "/4-html/home.html";
  }
}

botaoAlternarSenha?.addEventListener("click", () => {
  if (!inputSenha) return;

  const mostrando = inputSenha.type === "text";

  // alterna tipo do input
  inputSenha.type = mostrando ? "password" : "text";

  // muda o ícone
  botaoAlternarSenha.innerHTML = mostrando
    ? '<i class="fa-regular fa-eye"></i>'
    : '<i class="fa-regular fa-eye-slash"></i>';

  // acessibilidade
  botaoAlternarSenha.setAttribute(
    "aria-label",
    mostrando ? "Mostrar senha" : "Ocultar senha"
  );
});

linkEsqueciSenha?.addEventListener("click", async (e) => {
  e.preventDefault();

  const email = String(inputEmail?.value || "").trim().toLowerCase();
  if (!email) {
    setMsg("Digite seu e-mail para recuperar a senha.", "erro");
    inputEmail?.focus();
    return;
  }

  setLoading(true);
  setMsg("Enviando link de recuperação...", "info");

  try {
    const { error } = await supabase.auth.resetPasswordForEmail(email, {
      redirectTo: `${window.location.origin}/7-login/login.html`
    });

    if (error) {
      setMsg(error.message, "erro");
      return;
    }

    setMsg("Se o e-mail existir, o link de recuperação foi enviado.", "sucesso");
  } catch (err) {
    console.error(err);
    setMsg("Não foi possível solicitar a recuperação de senha.", "erro");
  } finally {
    setLoading(false);
  }
});

form?.addEventListener("submit", async (e) => {
  e.preventDefault();

  const email = String(inputEmail?.value || "").trim().toLowerCase();
  const password = String(inputSenha?.value || "");

  if (!email) {
    setMsg("Informe seu e-mail.", "erro");
    inputEmail?.focus();
    return;
  }

  if (!password) {
    setMsg("Informe sua senha.", "erro");
    inputSenha?.focus();
    return;
  }

  setLoading(true);
  setMsg("Verificando acesso...", "info");

  try {
    const { error } = await supabase.auth.signInWithPassword({ email, password });

    if (error) {
      setMsg(error.message, "erro");
      return;
    }

    salvarPreferenciaEmail(email);
    setMsg("Login realizado com sucesso. Redirecionando...", "sucesso");

    window.location.href = "/4-html/home.html";
  } catch (err) {
    console.error(err);
    setMsg("Erro inesperado ao entrar.", "erro");
  } finally {
    setLoading(false);
  }
});

function boot() {
  preencherAno();
  preencherAmbiente();
  restaurarEmail();
  setMsg("");
  redirecionarSeJaLogado();
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", boot, { once: true });
} else {
  boot();
}
