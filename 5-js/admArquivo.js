// /5-js/admArquivo.js
import { requireAuth, logout } from "../8-src/auth.js";
import { supabase } from "../8-src/supabaseClient.js";

const user = await requireAuth({ redirectTo: "../7-login/login.html" });
if (!user) throw new Error("Sem sessão");

// exemplo de botão logout:
document.querySelector("#btnLogout")?.addEventListener("click", () => logout());

/* =========================
   CONFIG
========================= */

// ✅ AJUSTE AQUI SE NECESSÁRIO
const SUPABASE_ARQUIVOS_TABLE = "arquivos";
const SUPABASE_STORAGE_BUCKET = "arquivos";

// ✅ WEBHOOK DO LOTE (Lançar partes - geral, 1 POST)
const N8N_WEBHOOK_URL_LOTE = "https://n8n.srv962474.hstgr.cloud/webhook/envioArquivos";

// ✅ WEBHOOK DO INDIVIDUAL (botões "Lançar Pxx", 1 POST por item)
const N8N_WEBHOOK_URL_INDIVIDUAL = "https://n8n.srv962474.hstgr.cloud/webhook/envioArquivos";

// ✅ WEBHOOK EXCLUIR (apaga as 2 últimas mensagens enviadas)
const N8N_WEBHOOK_URL_EXCLUIR = "https://n8n.srv962474.hstgr.cloud/webhook/excluir";

const FETCH_TIMEOUT_MS = 30000;

// Coluna que define parte/sobra
const nomeColunaParte = "Nº PESQ.";
const maxLinhasPreview = 80;

/* =========================
   DOM
========================= */
const alertaBox = document.getElementById("alertaBox");
const alertaTexto = document.getElementById("alertaTexto");

const nomeArquivoEl = document.getElementById("nomeArquivo");
const dataArquivoEl = document.getElementById("dataArquivo");
const tamanhoArquivoEl = document.getElementById("tamanhoArquivo");
const statusArquivoEl = document.getElementById("statusArquivo");

const totalPartesEl = document.getElementById("totalPartes");
const totalSobrasEl = document.getElementById("totalSobras");
const totalUsuariosEl = document.getElementById("totalUsuarios");
const totalLinhasEl = document.getElementById("totalLinhas");

const corpoTabelaPartes = document.getElementById("corpoTabelaPartes");
const estadoVazioPartes = document.getElementById("estadoVazioPartes");

const corpoTabelaSobras = document.getElementById("corpoTabelaSobras");
const estadoVazioSobras = document.getElementById("estadoVazioSobras");

const botaoRecarregar = document.getElementById("botaoRecarregar");
const botaoLancarPartes = document.getElementById("botaoLancarPartes");
const textoDicaLancamento = document.getElementById("textoDicaLancamento");

// Modal visualizar
const overlayVisualizar = document.getElementById("overlayVisualizar");
const tituloVisualizar = document.getElementById("tituloVisualizar");
const botaoFecharVisualizar = document.getElementById("botaoFecharVisualizar");
const infoVisualizar = document.getElementById("infoVisualizar");
const tabelaModal = document.getElementById("tabelaModal");

/* =========================
   ESTADO
========================= */
let arquivoKeyAtual = null;
let idArquivoAtual = null;
let registroArquivoAtual = null;

let usuariosDisponiveis = []; // [{id, parte, telegramId, telegramUsername, status?}]

// principais e sobras separados
let partesGeradas = []; // [{chaveParte:"P01", ...}]
let sobrasGeradas = []; // [{chaveParte:"S01.1", labelVisivel:"P01.1", ...}]

let selecaoUsuarioPorParte = new Map(); // chaveParte(P01) -> usuarioId
let selecaoUsuarioPorSobra = new Map(); // chaveSobra(S01.1) -> usuarioId

let chaveParteAbertaNoModal = null;

/* =========================
   UTIL
========================= */
function mostrarAlerta(mensagem) {
  if (!alertaBox || !alertaTexto) return;
  alertaTexto.textContent = mensagem;
  alertaBox.hidden = false;
}

function esconderAlerta() {
  if (!alertaBox || !alertaTexto) return;
  alertaTexto.textContent = "";
  alertaBox.hidden = true;
}

function obterParametroUrl(nome) {
  const url = new URL(window.location.href);
  return url.searchParams.get(nome);
}

function normalizarTexto(txt) {
  return String(txt || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toLowerCase();
}

function formatarBytes(bytes) {
  if (!Number.isFinite(bytes)) return "—";
  const unidades = ["B", "KB", "MB", "GB"];
  let valor = bytes;
  let idx = 0;
  while (valor >= 1024 && idx < unidades.length - 1) {
    valor /= 1024;
    idx += 1;
  }
  return `${valor.toFixed(idx === 0 ? 0 : 2)} ${unidades[idx]}`;
}

function obterDataHoraBr() {
  const d = new Date();
  const dia = String(d.getDate()).padStart(2, "0");
  const mes = String(d.getMonth() + 1).padStart(2, "0");
  const ano = d.getFullYear();
  const hora = String(d.getHours()).padStart(2, "0");
  const min = String(d.getMinutes()).padStart(2, "0");
  return `${dia}/${mes}/${ano} ${hora}:${min}`;
}

function formatarDataBr(valor) {
  if (!valor) return "—";

  try {
    const d = new Date(valor);
    if (Number.isNaN(d.getTime())) return String(valor);
    return d.toLocaleString("pt-BR");
  } catch {
    return String(valor);
  }
}

function estimarTamanhoCsv(csv) {
  return new Blob([csv], { type: "text/csv;charset=utf-8" }).size;
}

function obterFlagAutoLancar() {
  const v = obterParametroUrl("autoLancar");
  return v === "1" || String(v).toLowerCase() === "true";
}

function encontrarChaveColunaPrimeiraLinha(objPrimeiraLinha) {
  const chaves = Object.keys(objPrimeiraLinha || {});
  const alvo = normalizarTexto(nomeColunaParte);

  for (const c of chaves) {
    if (normalizarTexto(c) === alvo) return c;
  }

  for (const c of chaves) {
    const n = normalizarTexto(c);
    if (n.includes("pesq") && (n.includes("n") || n.includes("nr") || n.includes("no") || n.includes("numero"))) return c;
    if (n.includes("pesq")) return c;
  }

  return null;
}

function obterAdminAtualLocal() {
  try {
    return JSON.parse(localStorage.getItem("adminAtual")) || { id: "admin-001" };
  } catch {
    return { id: "admin-001" };
  }
}

async function getAccessTokenOrThrow() {
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

/* =========================
   ✅ arquivoKey (estável)
========================= */
function obterArquivoKeyAtual() {
  const k =
    String(
      registroArquivoAtual?.arquivoKey ||
      registroArquivoAtual?.arquivo_key ||
      registroArquivoAtual?.key ||
      ""
    ).trim();

  if (k) return k;
  if (arquivoKeyAtual) return String(arquivoKeyAtual);
  if (Number.isFinite(idArquivoAtual)) return `ARQ-${String(idArquivoAtual)}`;
  return "";
}

/* =========================
   FETCH COM TIMEOUT
========================= */
async function fetchComTimeout(url, options = {}, timeoutMs = 30000) {
  const controller = new AbortController();
  const id = setTimeout(() => controller.abort(), timeoutMs);

  try {
    const resp = await fetch(url, { ...options, signal: controller.signal });
    return resp;
  } catch (err) {
    if (err?.name === "AbortError") {
      throw new Error(`Timeout após ${timeoutMs}ms ao chamar ${url}`);
    }
    throw err;
  } finally {
    clearTimeout(id);
  }
}

/* =========================
   ✅ EXCLUIR (ANTES DE ENVIAR)
========================= */
async function excluirUltimasMensagensAntesDeEnviar(contexto = "") {
  const token = await getAccessTokenOrThrow();

  const payload = {
    contexto: String(contexto || "").slice(0, 120),
    idArquivoOrigem: idArquivoAtual ?? null,
    nomeArquivoOrigem: registroArquivoAtual?.nome || "",
    arquivoKey: obterArquivoKeyAtual(),
    executadoEm: new Date().toISOString(),
  };

  const resp = await fetchComTimeout(
    N8N_WEBHOOK_URL_EXCLUIR,
    {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${token}`,
        "x-payload-tipo": "EXCLUIR",
      },
      body: JSON.stringify(payload),
    },
    FETCH_TIMEOUT_MS
  );

  if (!resp.ok) {
    const txt = await resp.text().catch(() => "");
    throw new Error(`Falha ao excluir últimas mensagens (HTTP ${resp.status}): ${txt || "sem resposta"}`);
  }

  return await resp.json().catch(() => ({}));
}

/* =========================
   TELEGRAM (tolerante)
========================= */
function obterTelegramDoUsuario(u) {
  const rawId =
    u?.telegramUserId ??
    u?.telegramId ??
    u?.telegramID ??
    u?.telegram_id ??
    u?.telegramid ??
    u?.tgId ??
    u?.tgID ??
    "";

  const rawUser =
    u?.telegramUsername ??
    u?.telegram_username ??
    u?.telegramUser ??
    u?.tgUsername ??
    u?.tg_user ??
    "";

  const telegramId = String(rawId ?? "").trim();
  const telegramUsername = String(rawUser ?? "").trim().replace(/^@/, "");

  return { telegramId, telegramUsername };
}

function normalizarUsuarioBruto(u) {
  const id = String(u?.id || "").trim();
  const parte = String(u?.parte || "").trim().toUpperCase();

  const { telegramId, telegramUsername } = obterTelegramDoUsuario(u);

  const tgIdFinal = telegramId || String(u?.telegram_user_id || "").trim();
  const tgUserFinal = telegramUsername || String(u?.telegram_username || "").trim().replace(/^@/, "");

  return {
    ...u,
    id,
    parte,
    telegramId: tgIdFinal,
    telegramUsername: tgUserFinal ? `@${tgUserFinal.replace(/^@/, "")}` : "",
    status: String(u?.status || "ativo").toLowerCase(),
  };
}

function montarLabelUsuario(u) {
  const id = String(u?.id || "").trim();
  const idCurto = id ? id.slice(-6) : "—";
  const { telegramUsername } = obterTelegramDoUsuario(u);
  if (telegramUsername) return `@${telegramUsername.replace(/^@/, "")} • ${idCurto}`;
  if (u?.telegramUsername) return `${String(u.telegramUsername)} • ${idCurto}`;
  return `Usuário ${idCurto}`;
}

/* =========================
   HELPERS
========================= */
function getUsuarioPorId(usuarioId) {
  const id = String(usuarioId || "");
  return usuariosDisponiveis.find((u) => String(u.id) === id) || null;
}

function getItemPorChave({ tipo, chave }) {
  const c = String(chave || "");
  if (tipo === "sobra") return sobrasGeradas.find((x) => String(x.chaveParte) === c) || null;
  return partesGeradas.find((x) => String(x.chaveParte) === c) || null;
}

function obterArquivoKeyDaUrl() {
  const url = new URL(window.location.href);

  const key = url.searchParams.get("key");
  if (key && String(key).trim()) return String(key).trim();

  const id = url.searchParams.get("id");
  if (id && String(id).trim()) return `ARQ-${String(id).trim()}`;

  return "";
}

function obterIdNumericoDaUrl() {
  const url = new URL(window.location.href);
  const id = url.searchParams.get("id");
  const n = Number(id);
  return Number.isFinite(n) ? n : null;
}

/* =========================
   PARTE vs SOBRA — NORMALIZAÇÃO
========================= */
function pad2(n) {
  return String(n).padStart(2, "0");
}

function classificarParte(valor) {
  if (valor === null || valor === undefined) return null;

  if (typeof valor === "number" && Number.isFinite(valor)) {
    const inteiro = Math.trunc(valor);
    const frac = Math.abs(valor - inteiro);

    if (inteiro <= 0) return null;

    if (frac > 0) {
      const fracStr = String(valor).includes(".") ? String(valor).split(".")[1] : "1";
      const key = `S${pad2(inteiro)}.${fracStr}`;
      return { tipo: "sobra", chave: key, label: `S${pad2(inteiro)}.${fracStr}`, nInt: inteiro, nFrac: parseInt(fracStr, 10) || 0, raw: String(valor) };
    }

    return { tipo: "parte", chave: `P${pad2(inteiro)}`, label: `P${pad2(inteiro)}`, nInt: inteiro, nFrac: 0, raw: String(valor) };
  }

  const raw = String(valor).trim();
  if (!raw) return null;

  const v = raw.replace(",", ".").toUpperCase();

  const mP = v.match(/^P\s*0*(\d+)(?:\.(\d+))?$/i);
  if (mP) {
    const inteiro = parseInt(mP[1], 10);
    if (!Number.isFinite(inteiro) || inteiro <= 0) return null;

    const fracStr = mP[2] ? String(mP[2]) : "";
    if (fracStr) {
      const key = `S${pad2(inteiro)}.${fracStr}`;
      return { tipo: "sobra", chave: key, label: `S${pad2(inteiro)}.${fracStr}`, nInt: inteiro, nFrac: parseInt(fracStr, 10) || 0, raw };
    }

    return { tipo: "parte", chave: `P${pad2(inteiro)}`, label: `P${pad2(inteiro)}`, nInt: inteiro, nFrac: 0, raw };
  }

  const mNum = v.match(/^(\d+)(?:\.(\d+))?$/);
  if (mNum) {
    const inteiro = parseInt(mNum[1], 10);
    if (!Number.isFinite(inteiro) || inteiro <= 0) return null;

    const fracStr = mNum[2] ? String(mNum[2]) : "";
    if (fracStr) {
      const key = `S${pad2(inteiro)}.${fracStr}`;
      return { tipo: "sobra", chave: key, label: `S${pad2(inteiro)}.${fracStr}`, nInt: inteiro, nFrac: parseInt(fracStr, 10) || 0, raw };
    }

    return { tipo: "parte", chave: `P${pad2(inteiro)}`, label: `P${pad2(inteiro)}`, nInt: inteiro, nFrac: 0, raw };
  }

  const extr = v.match(/(\d+)/);
  if (extr) {
    const inteiro = parseInt(extr[1], 10);
    if (!Number.isFinite(inteiro) || inteiro <= 0) return null;
    return { tipo: "parte", chave: `P${pad2(inteiro)}`, label: `P${pad2(inteiro)}`, nInt: inteiro, nFrac: 0, raw };
  }

  return null;
}

function ordenarPartesPrincipais(arr) {
  return [...arr].sort((a, b) => (a.nInt || 0) - (b.nInt || 0));
}

function ordenarSobras(arr) {
  return [...arr].sort((a, b) => {
    const ai = a.nInt || 0;
    const bi = b.nInt || 0;
    if (ai !== bi) return ai - bi;
    return (a.nFrac || 0) - (b.nFrac || 0);
  });
}

/* =========================
   MODAL VISUALIZAR
========================= */
function abrirModalVisualizar(chave, origem = "parte") {
  const lista = origem === "sobra" ? sobrasGeradas : partesGeradas;
  const item = lista.find((p) => p.chaveParte === chave);
  if (!item) return;

  chaveParteAbertaNoModal = chave;

  if (tituloVisualizar) tituloVisualizar.textContent = `Visualizar ${item.labelVisivel || item.labelParte || item.chaveParte}`;

  if (infoVisualizar) {
    const total = item.totalLinhas;
    const tamanho = item.tamanhoTexto;
    infoVisualizar.innerHTML =
      `${origem === "sobra" ? "Sobra" : "Parte"} <strong>${item.labelVisivel || item.labelParte || item.chaveParte}</strong> — ` +
      `<strong>${total}</strong> linha(s) — <strong>${tamanho}</strong>`;
  }

  renderizarTabelaPreviewModal(item.csv);
  if (overlayVisualizar) overlayVisualizar.hidden = false;

  const geoAtivo = listaGeoAtiva(registroArquivoAtual?.colunasVisiveis);
  infoVisualizar.innerHTML += `<br/>Colunas geo enviadas: <strong>${geoAtivo}</strong>`;
}

function fecharModalVisualizar() {
  chaveParteAbertaNoModal = null;
  if (overlayVisualizar) overlayVisualizar.hidden = true;
  if (tabelaModal) tabelaModal.innerHTML = "";
}

function renderizarTabelaPreviewModal(csv) {
  if (!tabelaModal) return;

  tabelaModal.innerHTML = "";

  const linhas = String(csv || "")
    .split("\n")
    .filter((l) => l.trim().length > 0);

  if (!linhas.length) {
    tabelaModal.innerHTML = "<tr><td>Nenhum dado para visualizar.</td></tr>";
    return;
  }

  const cabecalho = (linhas[0] || "").split(",").map((c) => c.replace(/^"|"$/g, ""));
  const dados = linhas.slice(1, 1 + maxLinhasPreview);

  const thead = document.createElement("thead");
  const trHead = document.createElement("tr");
  for (const col of cabecalho) {
    const th = document.createElement("th");
    th.textContent = col;
    trHead.appendChild(th);
  }
  thead.appendChild(trHead);

  const tbody = document.createElement("tbody");
  for (const linha of dados) {
    const tr = document.createElement("tr");
    const cols = linha.split(",");
    for (const c of cols) {
      const td = document.createElement("td");
      td.textContent = c.replace(/^"|"$/g, "");
      tr.appendChild(td);
    }
    tbody.appendChild(tr);
  }

  tabelaModal.appendChild(thead);
  tabelaModal.appendChild(tbody);
}

if (botaoFecharVisualizar) botaoFecharVisualizar.addEventListener("click", fecharModalVisualizar);
if (overlayVisualizar)
  overlayVisualizar.addEventListener("click", (e) => {
    if (e.target === overlayVisualizar) fecharModalVisualizar();
  });
window.addEventListener("keydown", (e) => {
  if (overlayVisualizar && !overlayVisualizar.hidden && e.key === "Escape") fecharModalVisualizar();
});

/* =========================
   BOTÃO "LANÇAR" (INDIVIDUAL)
========================= */
function criarBotaoLancar({ texto, tipo, chave }) {
  const btn = document.createElement("button");
  btn.type = "button";
  btn.className = "botaoLancarParte";
  btn.textContent = texto;
  btn.dataset.tipo = tipo;
  btn.dataset.chave = chave;

  btn.addEventListener("click", async (e) => {
    e.stopPropagation();

    try {
      btn.disabled = true;
      const oldText = btn.textContent;
      btn.textContent = "Excluindo últimas mensagens…";

      await excluirUltimasMensagensAntesDeEnviar(`INDIVIDUAL ${texto}`);

      btn.textContent = "Enviando…";

      const resp = await enviarParteIndividualParaN8n({ tipo, chave });

      const lancamentoId = resp?.lancamentoId || resp?.id || "";
      alert(
        `✅ Enviado!\n` +
        `Item: ${texto}\n` +
        (lancamentoId ? `LancamentoId: ${lancamentoId}\n` : "") +
        `Resposta do n8n recebida.`
      );

      btn.textContent = oldText;
      btn.disabled = false;
    } catch (err) {
      console.error(err);
      alert(`❌ Falha ao enviar ${texto}\n\n${err?.message || err}`);
      btn.disabled = false;
      btn.textContent = texto;
    }
  });

  return btn;
}

/* =========================
   SELECT DE USUÁRIOS (limpo)
========================= */
function criarSelectUsuarios({ valorSelecionado, onChange }) {
  const select = document.createElement("select");
  select.className = "selectUsuario";

  const optVazio = document.createElement("option");
  optVazio.value = "";
  optVazio.textContent = "Selecionar usuário…";
  select.appendChild(optVazio);

  for (const u of usuariosDisponiveis) {
    const opt = document.createElement("option");
    opt.value = String(u.id);
    opt.textContent = montarLabelUsuario(u);
    select.appendChild(opt);
  }

  if (valorSelecionado) select.value = String(valorSelecionado);

  select.addEventListener("change", () => onChange(select.value ? String(select.value) : ""));
  return select;
}

/* =========================
   RENDER TABELAS
========================= */
function renderizarTabelaPartes() {
  if (!corpoTabelaPartes) return;

  corpoTabelaPartes.innerHTML = "";

  if (!partesGeradas.length) {
    if (estadoVazioPartes) estadoVazioPartes.hidden = false;
    return;
  }
  if (estadoVazioPartes) estadoVazioPartes.hidden = true;

  for (const parte of partesGeradas) {
    const tr = document.createElement("tr");

    const tdParte = document.createElement("td");
    tdParte.textContent = parte.labelParte || parte.chaveParte;

    const tdLinhas = document.createElement("td");
    tdLinhas.textContent = String(parte.totalLinhas);

    const tdTamanho = document.createElement("td");
    tdTamanho.textContent = parte.tamanhoTexto;

    const tdVisualizar = document.createElement("td");
    const btnVisualizar = document.createElement("button");
    btnVisualizar.type = "button";
    btnVisualizar.className = "botaoSecundario";
    btnVisualizar.textContent = "👁 Visualizar";
    btnVisualizar.addEventListener("click", () => abrirModalVisualizar(parte.chaveParte, "parte"));
    tdVisualizar.appendChild(btnVisualizar);

    const tdUsuario = document.createElement("td");
    const sel = criarSelectUsuarios({
      valorSelecionado: selecaoUsuarioPorParte.get(parte.chaveParte),
      onChange: (val) => {
        if (!val) selecaoUsuarioPorParte.delete(parte.chaveParte);
        else selecaoUsuarioPorParte.set(parte.chaveParte, val);

        renderizarTabelaPartes();
        atualizarEstadoBotaoLancar();
      },
    });
    tdUsuario.appendChild(sel);

    const tdLancar = document.createElement("td");
    tdLancar.appendChild(
      criarBotaoLancar({
        texto: `Lançar ${parte.labelParte || parte.chaveParte}`,
        tipo: "parte",
        chave: parte.chaveParte,
      })
    );

    tr.appendChild(tdParte);
    tr.appendChild(tdLinhas);
    tr.appendChild(tdTamanho);
    tr.appendChild(tdVisualizar);
    tr.appendChild(tdUsuario);
    tr.appendChild(tdLancar);

    corpoTabelaPartes.appendChild(tr);
  }
}

function renderizarTabelaSobras() {
  if (!corpoTabelaSobras) return;

  corpoTabelaSobras.innerHTML = "";

  if (!sobrasGeradas.length) {
    if (estadoVazioSobras) estadoVazioSobras.hidden = false;
    return;
  }
  if (estadoVazioSobras) estadoVazioSobras.hidden = true;

  for (const sobra of sobrasGeradas) {
    const tr = document.createElement("tr");

    const tdSobra = document.createElement("td");
    tdSobra.textContent = sobra.labelVisivel || sobra.labelParte || sobra.chaveParte;

    const tdLinhas = document.createElement("td");
    tdLinhas.textContent = String(sobra.totalLinhas);

    const tdTamanho = document.createElement("td");
    tdTamanho.textContent = sobra.tamanhoTexto;

    const tdVisualizar = document.createElement("td");
    const btnVisualizar = document.createElement("button");
    btnVisualizar.type = "button";
    btnVisualizar.className = "botaoSecundario";
    btnVisualizar.textContent = "👁 Visualizar";
    btnVisualizar.addEventListener("click", () => abrirModalVisualizar(sobra.chaveParte, "sobra"));
    tdVisualizar.appendChild(btnVisualizar);

    const tdUsuario = document.createElement("td");
    const sel = criarSelectUsuarios({
      valorSelecionado: selecaoUsuarioPorSobra.get(sobra.chaveParte),
      onChange: (val) => {
        if (!val) selecaoUsuarioPorSobra.delete(sobra.chaveParte);
        else selecaoUsuarioPorSobra.set(sobra.chaveParte, val);

        renderizarTabelaSobras();
        atualizarEstadoBotaoLancar();
      },
    });
    tdUsuario.appendChild(sel);

    const tdLancar = document.createElement("td");
    tdLancar.appendChild(
      criarBotaoLancar({
        texto: `Lançar ${sobra.labelVisivel || sobra.labelParte || sobra.chaveParte}`,
        tipo: "sobra",
        chave: sobra.chaveParte,
      })
    );

    tr.appendChild(tdSobra);
    tr.appendChild(tdLinhas);
    tr.appendChild(tdTamanho);
    tr.appendChild(tdVisualizar);
    tr.appendChild(tdUsuario);
    tr.appendChild(tdLancar);

    corpoTabelaSobras.appendChild(tr);
  }
}

function atualizarEstadoBotaoLancar() {
  if (!botaoLancarPartes) return;

  const totalPrincipais = partesGeradas.length;
  const selPrincipais = selecaoUsuarioPorParte.size;

  const pode = usuariosDisponiveis.length > 0 && totalPrincipais > 0 && selPrincipais === totalPrincipais;

  botaoLancarPartes.disabled = !pode;

  if (textoDicaLancamento) {
    if (!usuariosDisponiveis.length) {
      textoDicaLancamento.textContent = "Nenhum usuário cadastrado.";
    } else if (!totalPrincipais && !sobrasGeradas.length) {
      textoDicaLancamento.textContent = "Nenhuma parte detectada.";
    } else if (!pode) {
      textoDicaLancamento.textContent = `Selecione usuários: ${selPrincipais}/${totalPrincipais}`;
    } else {
      const selecionadasSobras = selecaoUsuarioPorSobra.size;
      const totalSobras = sobrasGeradas.length;
      textoDicaLancamento.textContent =
        totalSobras > 0
          ? `Pronto para lançar. (Sobras opcionais: ${selecionadasSobras}/${totalSobras} selecionada(s))`
          : "Pronto para lançar.";
    }
  }
}

/* =========================
   COLUNAS GEO — FILTRO
========================= */
function normalizarChave(ch) {
  return String(ch || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toLowerCase();
}

function detectarChavesGeo(objPrimeiraLinha) {
  const chaves = Object.keys(objPrimeiraLinha || {});
  const norm = chaves.map((k) => ({ key: k, n: normalizarChave(k) }));

  const achar = (pred) => norm.find((x) => pred(x.n))?.key || null;

  const estadoKey =
    achar((n) => n === "estado") ||
    achar((n) => n === "uf") ||
    achar((n) => n.includes("estado")) ||
    achar((n) => n === "u.f.");

  const cidadeKey =
    achar((n) => n === "cidade") ||
    achar((n) => n === "municipio") ||
    achar((n) => n.includes("cidade")) ||
    achar((n) => n.includes("municip"));

  const regiaoKey =
    achar((n) => n === "regiao") ||
    achar((n) => n.includes("regiao")) ||
    achar((n) => n.includes("regional"));

  return { estadoKey, cidadeKey, regiaoKey };
}

function filtrarLinhasGeo(linhas, colunasVisiveis, geoKeys) {
  const vis = {
    estado: colunasVisiveis?.estado !== false,
    cidade: colunasVisiveis?.cidade !== false,
    regiao: colunasVisiveis?.regiao !== false,
  };

  const { estadoKey, cidadeKey, regiaoKey } = geoKeys || {};
  if (!estadoKey && !cidadeKey && !regiaoKey) return linhas;

  return (linhas || []).map((linha) => {
    const novo = { ...linha };

    if (estadoKey && !vis.estado) delete novo[estadoKey];
    if (cidadeKey && !vis.cidade) delete novo[cidadeKey];
    if (regiaoKey && !vis.regiao) delete novo[regiaoKey];

    return novo;
  });
}

function listaGeoAtiva(colunasVisiveis) {
  const vis = {
    estado: colunasVisiveis?.estado !== false,
    cidade: colunasVisiveis?.cidade !== false,
    regiao: colunasVisiveis?.regiao !== false,
  };
  const ativos = [];
  if (vis.estado) ativos.push("Estado");
  if (vis.cidade) ativos.push("Cidade");
  if (vis.regiao) ativos.push("Região");
  return ativos.length ? ativos.join(", ") : "Nenhuma (geo)";
}

/* =========================
   EXCEL -> PARTES / SOBRAS
========================= */
function gerarCsvDaParte(linhas) {
  const ws = window.XLSX.utils.json_to_sheet(linhas);
  const csv = window.XLSX.utils.sheet_to_csv(ws);
  return csv || "";
}

function separarPorPartesESobras(linhasJson, colunasVisiveis) {
  if (!Array.isArray(linhasJson) || linhasJson.length === 0) return { principais: [], sobras: [] };
  const geoKeys = detectarChavesGeo(linhasJson[0]);
  const chaveColuna = encontrarChaveColunaPrimeiraLinha(linhasJson[0]);
  if (!chaveColuna) throw new Error(`Não encontrei a coluna "${nomeColunaParte}". Verifique o nome no Excel.`);

  const mapaPrincipais = new Map();
  const mapaSobras = new Map();

  const metaPrincipais = new Map();
  const metaSobras = new Map();

  for (const linha of linhasJson) {
    const cls = classificarParte(linha[chaveColuna]);
    if (!cls) continue;

    if (cls.tipo === "parte") {
      if (!mapaPrincipais.has(cls.chave)) mapaPrincipais.set(cls.chave, []);
      mapaPrincipais.get(cls.chave).push(linha);
      metaPrincipais.set(cls.chave, cls);
    } else {
      if (!mapaSobras.has(cls.chave)) mapaSobras.set(cls.chave, []);
      mapaSobras.get(cls.chave).push(linha);
      metaSobras.set(cls.chave, cls);
    }
  }

  const principais = [];
  for (const [key, linhas] of mapaPrincipais.entries()) {
    const linhasFiltradas = filtrarLinhasGeo(linhas, colunasVisiveis, geoKeys);
    const csv = gerarCsvDaParte(linhasFiltradas);
    const tamanhoBytes = estimarTamanhoCsv(csv);
    const cls = metaPrincipais.get(key);

    principais.push({
      chaveParte: key,
      labelParte: cls?.label || key,
      nInt: cls?.nInt || 0,
      nFrac: 0,
      linhas,
      totalLinhas: linhas.length,
      csv,
      tamanhoBytes,
      tamanhoTexto: formatarBytes(tamanhoBytes),
    });
  }

  const sobras = [];
  for (const [key, linhas] of mapaSobras.entries()) {
    const linhasFiltradas = filtrarLinhasGeo(linhas, colunasVisiveis, geoKeys);
    const csv = gerarCsvDaParte(linhasFiltradas);
    const tamanhoBytes = estimarTamanhoCsv(csv);
    const cls = metaSobras.get(key);

    const labelVisivel = cls?.label ? cls.label.replace(/^S/, "P") : key.replace(/^S/, "P");

    sobras.push({
      chaveParte: key,
      labelParte: cls?.label || key,
      labelVisivel,
      nInt: cls?.nInt || 0,
      nFrac: cls?.nFrac || 0,
      linhas,
      totalLinhas: linhas.length,
      csv,
      tamanhoBytes,
      tamanhoTexto: formatarBytes(tamanhoBytes),
    });
  }

  return {
    principais: ordenarPartesPrincipais(principais),
    sobras: ordenarSobras(sobras),
  };
}

/* =========================
   ✅ CARREGAR USUÁRIOS (SUPABASE primeiro)
========================= */
async function carregarUsuariosDoSupabase() {
  const ownerId = user?.id;
  if (!ownerId) return [];

  {
    const { data, error } = await supabase
      .from("usuarios")
      .select("id,parte,telegram_user_id,telegram_username,status")
      .eq("owner_id", ownerId);

    if (!error) return data || [];

    const msg = String(error?.message || "").toLowerCase();
    if (!msg.includes("status")) throw error;
  }

  const { data, error } = await supabase
    .from("usuarios")
    .select("id,parte,telegram_user_id,telegram_username")
    .eq("owner_id", ownerId);

  if (error) throw error;
  return data || [];
}

async function carregarUsuariosDoIndexedDbFallback() {
  if (!window.ibeDb || typeof window.ibeDb.listarUsuariosDb !== "function") return [];

  const admin = obterAdminAtualLocal();
  const lista = await window.ibeDb.listarUsuariosDb({ adminId: admin.id });
  return lista || [];
}

async function carregarUsuarios() {
  try {
    const rows = await carregarUsuariosDoSupabase();

    usuariosDisponiveis = (rows || [])
      .map((r) => ({
        id: r.id,
        parte: r.parte,
        telegram_user_id: r.telegram_user_id,
        telegram_username: r.telegram_username,
        status: r.status ?? "ativo",
      }))
      .map(normalizarUsuarioBruto)
      .filter((u) => String(u.status || "").toLowerCase() === "ativo");

    return;
  } catch (e) {
    console.warn("Falha ao carregar usuários do Supabase, usando fallback IndexedDB:", e?.message || e);
  }

  const lista = await carregarUsuariosDoIndexedDbFallback();
  usuariosDisponiveis = (lista || [])
    .map(normalizarUsuarioBruto)
    .filter((u) => String(u.status || "").toLowerCase() === "ativo");
}

/* =========================
   ✅ SUPABASE - ARQUIVOS
========================= */
function isBlobLike(valor) {
  return (
    valor instanceof Blob ||
    (valor &&
      typeof valor === "object" &&
      typeof valor.arrayBuffer === "function")
  );
}

function base64ParaBlob(base64, mimeType = "application/octet-stream") {
  const clean = String(base64 || "").trim();
  const match = clean.match(/^data:(.*?);base64,(.*)$/);
  const realMime = match?.[1] || mimeType;
  const realBase64 = match?.[2] || clean;

  const bin = atob(realBase64);
  const len = bin.length;
  const bytes = new Uint8Array(len);

  for (let i = 0; i < len; i++) {
    bytes[i] = bin.charCodeAt(i);
  }

  return new Blob([bytes], { type: realMime });
}

function primeiroValor(obj, campos = []) {
  for (const c of campos) {
    if (obj?.[c] !== undefined && obj?.[c] !== null && obj?.[c] !== "") {
      return obj[c];
    }
  }
  return null;
}

function extrairArquivoKeyDoRegistro(reg) {
  return String(
    primeiroValor(reg, ["arquivoKey", "arquivo_key", "key"]) || ""
  ).trim();
}

function extrairNomeArquivoDoRegistro(reg) {
  return String(
    primeiroValor(reg, ["nome", "name", "nome_arquivo", "filename", "file_name"]) || "arquivo"
  ).trim();
}

function extrairStatusArquivoDoRegistro(reg) {
  return String(
    primeiroValor(reg, ["status", "situacao"]) || "—"
  ).trim();
}

function extrairTamanhoBytesDoRegistro(reg) {
  const bruto = primeiroValor(reg, ["tamanhoBytes", "tamanho_bytes", "size", "file_size", "bytes"]);
  const n = Number(bruto);
  return Number.isFinite(n) ? n : null;
}

function extrairDataArquivoDoRegistro(reg) {
  return (
    primeiroValor(reg, ["data", "created_at", "criado_em", "uploaded_at", "updated_at"]) || ""
  );
}

function extrairColunasVisiveisDoRegistro(reg) {
  return (
    primeiroValor(reg, ["colunasVisiveis", "colunas_visiveis"]) || {}
  );
}

function extrairOpcoesPesquisaDoRegistro(reg) {
  return (
    primeiroValor(reg, ["opcoesPesquisa", "opcoes_pesquisa"]) || {}
  );
}

function extrairMimeTypeDoRegistro(reg) {
  return (
    primeiroValor(reg, ["mimeType", "mime_type", "content_type", "tipoMime"]) ||
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  );
}

function extrairStoragePathDoRegistro(reg) {
  return String(
    primeiroValor(reg, [
      "storage_path",
      "storagePath",
      "caminho_storage",
      "caminhoStorage",
      "file_path",
      "filePath",
      "path",
      "caminho_arquivo",
      "caminhoArquivo",
    ]) || ""
  ).trim();
}

function normalizarRegistroArquivoSupabase(reg) {
  const tamanhoBytes = extrairTamanhoBytesDoRegistro(reg);

  return {
    ...reg,
    id: reg?.id ?? null,
    arquivoKey: extrairArquivoKeyDoRegistro(reg),
    nome: extrairNomeArquivoDoRegistro(reg),
    data: formatarDataBr(extrairDataArquivoDoRegistro(reg)),
    tamanho: tamanhoBytes !== null ? formatarBytes(tamanhoBytes) : String(primeiroValor(reg, ["tamanho", "file_size_text"]) || "—"),
    tamanhoBytes: tamanhoBytes ?? 0,
    status: extrairStatusArquivoDoRegistro(reg),
    opcoesPesquisa: extrairOpcoesPesquisaDoRegistro(reg),
    colunasVisiveis: extrairColunasVisiveisDoRegistro(reg),
    storagePath: extrairStoragePathDoRegistro(reg),
    mimeType: extrairMimeTypeDoRegistro(reg),
  };
}

async function buscarRegistroPorCampo({ campo, valor, ownerId }) {
  try {
    let query = supabase
      .from(SUPABASE_ARQUIVOS_TABLE)
      .select("*")
      .limit(1);

    if (ownerId) query = query.eq("owner_id", ownerId);
    query = query.eq(campo, valor);

    const { data, error } = await query.maybeSingle();

    if (error) {
      const msg = String(error?.message || "").toLowerCase();

      // ignora se a coluna não existir
      if (
        msg.includes("column") ||
        msg.includes("does not exist") ||
        msg.includes("could not find the") ||
        msg.includes("schema cache")
      ) {
        return null;
      }

      throw error;
    }

    return data || null;
  } catch (e) {
    const msg = String(e?.message || "").toLowerCase();
    if (
      msg.includes("column") ||
      msg.includes("does not exist") ||
      msg.includes("could not find the") ||
      msg.includes("schema cache")
    ) {
      return null;
    }
    throw e;
  }
}

async function buscarRegistroArquivoNoSupabase({ arquivoKey, idNumerico, ownerId }) {
  const candidatosKey = ["arquivo_key", "arquivoKey", "key"];
  const candidatosId = ["id"];

  if (arquivoKey) {
    for (const campo of candidatosKey) {
      const reg = await buscarRegistroPorCampo({ campo, valor: arquivoKey, ownerId });
      if (reg) return reg;
    }
  }

  if (Number.isFinite(idNumerico)) {
    for (const campo of candidatosId) {
      const reg = await buscarRegistroPorCampo({ campo, valor: idNumerico, ownerId });
      if (reg) return reg;
    }
  }

  return null;
}

async function baixarBlobDoStorage(path) {
  const { data, error } = await supabase.storage
    .from(SUPABASE_STORAGE_BUCKET)
    .download(path);

  if (error) throw error;
  if (!data) throw new Error("Storage retornou vazio ao baixar o arquivo.");

  return data;
}

async function baixarBlobPorUrl(url, mimeType = "application/octet-stream") {
  const resp = await fetchComTimeout(url, { method: "GET" }, FETCH_TIMEOUT_MS);
  if (!resp.ok) {
    throw new Error(`Falha ao baixar arquivo pela URL (HTTP ${resp.status}).`);
  }

  const blob = await resp.blob();
  if (blob && blob.size > 0) return blob;

  const ab = await resp.arrayBuffer();
  return new Blob([ab], { type: mimeType });
}

async function carregarArquivoDoSupabase({ arquivoKey, idNumerico }) {
  const ownerId = user?.id || null;

  const registroRaw = await buscarRegistroArquivoNoSupabase({
    arquivoKey,
    idNumerico,
    ownerId,
  });

  if (!registroRaw) {
    throw new Error(`Arquivo não encontrado na tabela "${SUPABASE_ARQUIVOS_TABLE}" para a chave/id informado.`);
  }

  const registro = normalizarRegistroArquivoSupabase(registroRaw);

  let blob = null;

  // 1) Storage path
  const storagePath = extrairStoragePathDoRegistro(registroRaw);
  if (storagePath) {
    blob = await baixarBlobDoStorage(storagePath);
  }

  // 2) URL direta
  if (!blob) {
    const fileUrl = primeiroValor(registroRaw, ["url", "file_url", "arquivo_url", "public_url"]);
    if (fileUrl) {
      blob = await baixarBlobPorUrl(String(fileUrl), extrairMimeTypeDoRegistro(registroRaw));
    }
  }

  // 3) base64 salvo na tabela
  if (!blob) {
    const b64 = primeiroValor(registroRaw, ["base64", "arquivoBase64", "conteudoBase64", "blobBase64"]);
    if (typeof b64 === "string" && b64.trim()) {
      blob = base64ParaBlob(b64, extrairMimeTypeDoRegistro(registroRaw));
    }
  }

  // 4) blob-like salvo no próprio registro
  if (!blob) {
    const possivelBlob = primeiroValor(registroRaw, ["blob", "arquivoBlob", "file"]);
    if (isBlobLike(possivelBlob)) {
      blob = possivelBlob;
    }
  }

  if (!blob) {
    console.error("Registro do arquivo encontrado, mas sem conteúdo utilizável:", registroRaw);
    throw new Error(
      `O registro foi encontrado no Supabase, mas não localizei o conteúdo do Excel. Verifique se a tabela "${SUPABASE_ARQUIVOS_TABLE}" possui o path do arquivo no bucket "${SUPABASE_STORAGE_BUCKET}".`
    );
  }

  return { registro, blob };
}

async function atualizarRegistroArquivoNoSupabaseParcial(idRegistro, patch) {
  if (!Number.isFinite(Number(idRegistro))) return false;

  const tentativas = [
    patch,
    {
      ...(patch.opcoesPesquisa ? { opcoes_pesquisa: patch.opcoesPesquisa } : {}),
      ...(patch.colunasVisiveis ? { colunas_visiveis: patch.colunasVisiveis } : {}),
    },
  ];

  for (const tentativa of tentativas) {
    const temAlgo = tentativa && Object.keys(tentativa).length > 0;
    if (!temAlgo) continue;

    const { error } = await supabase
      .from(SUPABASE_ARQUIVOS_TABLE)
      .update(tentativa)
      .eq("id", Number(idRegistro));

    if (!error) return true;

    const msg = String(error?.message || "").toLowerCase();
    if (
      msg.includes("column") ||
      msg.includes("does not exist") ||
      msg.includes("schema cache")
    ) {
      continue;
    }

    console.warn("Falha ao atualizar registro do arquivo no Supabase:", error);
    return false;
  }

  return false;
}

/* =========================
   ARQUIVO (SUPABASE)
========================= */
async function carregarArquivoPorKey(arquivoKey) {
  return await carregarArquivoDoSupabase({
    arquivoKey: String(arquivoKey || "").trim(),
    idNumerico: idArquivoAtual,
  });
}

async function lerExcelParaJson(blobExcel) {
  if (!window.XLSX) throw new Error("SheetJS (XLSX) não carregou. Verifique a tag script do CDN.");

  const arrayBuffer = await blobExcel.arrayBuffer();
  const workbook = window.XLSX.read(arrayBuffer, { type: "array" });

  const nomePrimeiraAba = workbook.SheetNames[0];
  const ws = workbook.Sheets[nomePrimeiraAba];
  if (!ws) throw new Error("Não consegui ler a primeira aba do Excel.");

  return window.XLSX.utils.sheet_to_json(ws, { defval: "" });
}

/* =========================
   AUTO MAP (somente Partes)
========================= */
function aplicarAutoMapPartesPrincipais() {
  selecaoUsuarioPorParte.clear();

  const mapaParteParaUsuario = new Map();
  for (const u of usuariosDisponiveis) {
    const parte = String(u.parte || "").trim().toUpperCase();
    if (parte) mapaParteParaUsuario.set(parte, String(u.id));
  }

  for (const p of partesGeradas) {
    const usuarioId = mapaParteParaUsuario.get(p.chaveParte);
    if (usuarioId) selecaoUsuarioPorParte.set(p.chaveParte, usuarioId);
  }

  selecaoUsuarioPorSobra.clear();
}

/* =========================
   CARDS
========================= */
function preencherCardsInfo() {
  if (!registroArquivoAtual) return;

  if (nomeArquivoEl) nomeArquivoEl.textContent = registroArquivoAtual.nome || "—";
  if (dataArquivoEl) dataArquivoEl.textContent = registroArquivoAtual.data || "—";
  if (tamanhoArquivoEl) tamanhoArquivoEl.textContent = registroArquivoAtual.tamanho || "—";
  if (statusArquivoEl) statusArquivoEl.textContent = registroArquivoAtual.status || "—";

  if (totalPartesEl) totalPartesEl.textContent = String(partesGeradas.length || 0);
  if (totalSobrasEl) totalSobrasEl.textContent = String(sobrasGeradas.length || 0);
  if (totalUsuariosEl) totalUsuariosEl.textContent = String(usuariosDisponiveis.length || 0);

  const linhasSomadas =
    partesGeradas.reduce((acc, p) => acc + (p.totalLinhas || 0), 0) +
    sobrasGeradas.reduce((acc, p) => acc + (p.totalLinhas || 0), 0);

  if (totalLinhasEl) totalLinhasEl.textContent = String(linhasSomadas || 0);
}

/* =========================
   LANÇAR (VALIDAÇÃO)
========================= */
function validarDistribuicao() {
  if (!partesGeradas.length && !sobrasGeradas.length) return "Nenhuma parte foi detectada.";
  if (!usuariosDisponiveis.length) return "Nenhum usuário cadastrado.";

  for (const p of partesGeradas) {
    if (!selecaoUsuarioPorParte.has(p.chaveParte)) {
      return `Selecione um usuário para a parte ${p.labelParte || p.chaveParte}.`;
    }
  }

  return null;
}

/* =========================
   ENVIO (INDIVIDUAL / LOTE)
========================= */
async function enviarJsonParaN8n(url, payload, headersExtra = {}) {
  const token = await getAccessTokenOrThrow();

  const resp = await fetch(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Authorization": `Bearer ${token}`,
      ...headersExtra,
    },
    body: JSON.stringify(payload),
  });

  if (!resp.ok) {
    const txt = await resp.text().catch(() => "");
    throw new Error(`n8n HTTP ${resp.status}: ${txt || "sem resposta"}`);
  }

  return await resp.json().catch(() => ({}));
}

async function enviarLoteParaN8n(payload, headersExtra = {}) {
  const token = await getAccessTokenOrThrow();

  const resp = await fetchComTimeout(
    N8N_WEBHOOK_URL_LOTE,
    {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${token}`,
        "x-payload-tipo": "LOTE",
        ...headersExtra,
      },
      body: JSON.stringify(payload),
    },
    FETCH_TIMEOUT_MS
  );

  if (!resp.ok) {
    const txt = await resp.text().catch(() => "");
    throw new Error(`n8n HTTP ${resp.status}: ${txt || "sem resposta"}`);
  }

  return await resp.json().catch(() => ({}));
}

/* =========================
   ✅ LINK DO QUESTIONÁRIO
========================= */
const QUESTIONARIO_BASE_URL = "https://www.ibespebr.com.br/questionario/pesquisa/";

function normalizarNumeroPesquisa(v) {
  return String(v || "").trim().replace(/[^\d]/g, "");
}

function normalizarDataPesquisa(v) {
  return String(v || "").trim().replaceAll("/", "-");
}

function dataPesquisaValida(v) {
  return /^\d{2}-\d{2}-\d{4}$/.test(String(v || ""));
}

function obterOpcoesPesquisaDaUrl() {
  const numeroPesquisa = normalizarNumeroPesquisa(obterParametroUrl("numeroPesquisa"));
  const dataPesquisa = normalizarDataPesquisa(obterParametroUrl("dataPesquisa"));

  return {
    numeroPesquisa,
    dataPesquisa,
    valida: !!numeroPesquisa && dataPesquisaValida(dataPesquisa),
  };
}

function obterOpcoesPesquisaDoArquivo() {
  const op = registroArquivoAtual?.opcoesPesquisa || {};
  const numeroPesquisa = normalizarNumeroPesquisa(op.numeroPesquisa || "");
  const dataPesquisa = normalizarDataPesquisa(op.dataPesquisa || "");
  return {
    numeroPesquisa,
    dataPesquisa,
    valida: !!numeroPesquisa && dataPesquisaValida(dataPesquisa),
  };
}

function obterOpcoesPesquisaPreferenciais() {
  const opUrl = obterOpcoesPesquisaDaUrl();
  if (opUrl.valida) return { ...opUrl, fonte: "url" };

  const opArq = obterOpcoesPesquisaDoArquivo();
  if (opArq.valida) return { ...opArq, fonte: "arquivo" };

  return { numeroPesquisa: "", dataPesquisa: "", valida: false, fonte: "nenhuma" };
}

async function persistirOpcoesPesquisaSeVieramDaUrl() {
  try {
    const opUrl = obterOpcoesPesquisaDaUrl();
    if (!opUrl.valida) return;

    const opAtual = obterOpcoesPesquisaDoArquivo();
    if (opAtual.valida && opAtual.numeroPesquisa === opUrl.numeroPesquisa && opAtual.dataPesquisa === opUrl.dataPesquisa) return;

    const atualizado = {
      ...registroArquivoAtual,
      opcoesPesquisa: {
        numeroPesquisa: opUrl.numeroPesquisa,
        dataPesquisa: opUrl.dataPesquisa,
      },
    };

    const ok = await atualizarRegistroArquivoNoSupabaseParcial(registroArquivoAtual?.id, {
      opcoesPesquisa: atualizado.opcoesPesquisa,
    });

    if (!ok) {
      console.warn("Não foi possível persistir opcoesPesquisa no Supabase. Mantendo apenas em memória.");
    }

    registroArquivoAtual = atualizado;
  } catch (e) {
    console.warn("Falha ao persistir opcoesPesquisa no arquivo:", e);
  }
}

function montarLinkQuestionarioParaParte(itemLabelPxx) {
  const op = obterOpcoesPesquisaPreferenciais();
  if (!op.valida) return "";

  const parte = String(itemLabelPxx || "").trim().toUpperCase();
  if (!/^P\d{2}$/.test(parte)) return "";

  return `${QUESTIONARIO_BASE_URL}${encodeURIComponent(op.numeroPesquisa)}/${encodeURIComponent(parte)}/${encodeURIComponent(op.dataPesquisa)}`;
}

/* =========================
   ✅ ENVIO INDIVIDUAL
========================= */
async function enviarParteIndividualParaN8n({ tipo, chave }) {
  const tipoNormalizado = tipo === "sobra" ? "sobra" : "parte";

  const item = getItemPorChave({ tipo: tipoNormalizado, chave });
  if (!item) throw new Error("Item não encontrado para lançar.");

  const usuarioId =
    tipoNormalizado === "sobra"
      ? (selecaoUsuarioPorSobra.get(chave) || "")
      : (selecaoUsuarioPorParte.get(chave) || "");

  if (!usuarioId) throw new Error("Selecione um usuário antes de lançar.");

  const u = getUsuarioPorId(usuarioId);
  if (!u) throw new Error("Usuário selecionado não encontrado.");

  const { telegramId, telegramUsername } = obterTelegramDoUsuario(u);

  const tgId = String(telegramId || "").trim();
  const tgUser = String(telegramUsername || "").trim().replace(/^@/, "");

  if (!tgId) {
    throw new Error("Este usuário está sem telegramId. Confira no Administrar Usuários se foi preenchido e salvo.");
  }

  const labelPxx = item.labelParte || item.chaveParte;
  const linkQuestionario = tipoNormalizado === "parte" ? montarLinkQuestionarioParaParte(labelPxx) : "";

  const arquivoKey = obterArquivoKeyAtual();
  const idArquivoOrigem = idArquivoAtual ?? registroArquivoAtual?.id ?? null;
  const nomeArquivoOrigem = registroArquivoAtual?.nome || "arquivo";

  const body = {
    arquivoKey: arquivoKey || "",
    idArquivoOrigem: Number.isFinite(Number(idArquivoOrigem)) ? Number(idArquivoOrigem) : null,
    nomeArquivoOrigem,

    categoria: tipoNormalizado === "sobra" ? "SOBRA" : "PARTE",
    chaveParte: item.chaveParte,
    labelVisivel: item.labelVisivel || item.labelParte || item.chaveParte,
    totalLinhas: Number(item.totalLinhas || 0),
    tamanhoBytes: Number(item.tamanhoBytes || 0),
    csv: item.csv || "",
    linkQuestionario: linkQuestionario || "",
  };

  const headersUsuario = {
    "x-user-id": String(u.id || ""),
    "x-telegram-id": tgId,
    "x-telegram-username": tgUser ? `@${tgUser}` : "",
    "x-tipo": body.categoria,
    "x-chave-parte": String(body.chaveParte || ""),
    "x-link-questionario": linkQuestionario || "",
    "x-arquivo-key": arquivoKey || "",
    "x-id-arquivo-origem": Number.isFinite(Number(idArquivoOrigem)) ? String(idArquivoOrigem) : "",
    "x-nome-arquivo-origem": String(nomeArquivoOrigem || ""),
  };

  return await enviarJsonParaN8n(N8N_WEBHOOK_URL_INDIVIDUAL, body, headersUsuario);
}

/* =========================
   ✅ ENVIO EM LOTE
========================= */
async function lancarPartes() {
  const erro = validarDistribuicao();
  if (erro) {
    alert(erro);
    return;
  }

  if (!window.ibeDb || typeof window.ibeDb.salvarDistribuicaoDb !== "function") {
    alert("db.js não tem salvarDistribuicaoDb(itens).");
    return;
  }

  const opPref = obterOpcoesPesquisaPreferenciais();
  if (partesGeradas.length > 0 && !opPref.valida) {
    alert("Faltam opções do link. Volte ao Dashboard e preencha ⚙️ Opções (Número da pesquisa e Data DD-MM-AAAA) antes de lançar.");
    return;
  }

  botaoLancarPartes.disabled = true;
  if (textoDicaLancamento) textoDicaLancamento.textContent = "Excluindo últimas mensagens…";

  try {
    await excluirUltimasMensagensAntesDeEnviar("LOTE (Lançar partes geral)");
  } catch (e) {
    console.error(e);
    botaoLancarPartes.disabled = false;
    if (textoDicaLancamento) textoDicaLancamento.textContent = "Falha ao excluir mensagens.";
    alert(`❌ Falha ao excluir as últimas mensagens antes do lote.\n\n${e?.message || e}`);
    return;
  }

  if (textoDicaLancamento) textoDicaLancamento.textContent = "Lançando…";

  const idOrigemFinal = idArquivoAtual ?? registroArquivoAtual?.id ?? null;

  const itensPrincipaisSalvar = partesGeradas.map((p) => {
    const usuarioId = String(selecaoUsuarioPorParte.get(p.chaveParte));
    const blobParte = new Blob([p.csv || ""], { type: "text/csv;charset=utf-8" });

    const linkQuestionario = montarLinkQuestionarioParaParte(p.labelParte || p.chaveParte);

    return {
      idArquivoOrigem: Number.isFinite(Number(idOrigemFinal)) ? Number(idOrigemFinal) : null,
      nomeArquivoOrigem: registroArquivoAtual?.nome || "arquivo",
      chaveParte: p.chaveParte,
      usuarioId,
      totalLinhas: p.totalLinhas,
      tamanhoBytes: p.tamanhoBytes,
      criadoEm: Date.now(),
      criadoEmTexto: obterDataHoraBr(),
      blobParte,
      tipoParte: "CSV",
      status: "Disponível",
      categoria: "PARTE",
      csv: p.csv || "",
      linkQuestionario: linkQuestionario || "",
    };
  });

  const itensSobrasSalvar = sobrasGeradas
    .filter((s) => selecaoUsuarioPorSobra.has(s.chaveParte))
    .map((s) => {
      const usuarioId = String(selecaoUsuarioPorSobra.get(s.chaveParte));
      const blobParte = new Blob([s.csv || ""], { type: "text/csv;charset=utf-8" });

      return {
        idArquivoOrigem: Number.isFinite(Number(idOrigemFinal)) ? Number(idOrigemFinal) : null,
        nomeArquivoOrigem: registroArquivoAtual?.nome || "arquivo",
        chaveParte: s.chaveParte,
        usuarioId,
        totalLinhas: s.totalLinhas,
        tamanhoBytes: s.tamanhoBytes,
        criadoEm: Date.now(),
        criadoEmTexto: obterDataHoraBr(),
        blobParte,
        tipoParte: "CSV",
        status: "Disponível",
        categoria: "SOBRA",
        labelVisivel: s.labelVisivel || s.labelParte || s.chaveParte,
        csv: s.csv || "",
        linkQuestionario: "",
      };
    });

  try {
    await window.ibeDb.salvarDistribuicaoDb([...itensPrincipaisSalvar, ...itensSobrasSalvar]);

    if (textoDicaLancamento) textoDicaLancamento.textContent = "Enviando lote para o servidor…";

    const admin = obterAdminAtualLocal();
    const itensLote = [];

    for (const i of itensPrincipaisSalvar) {
      const u = getUsuarioPorId(i.usuarioId);
      const { telegramId, telegramUsername } = obterTelegramDoUsuario(u || {});
      if (!telegramId) throw new Error(`Usuário ${i.usuarioId} está sem telegramId (PARTE ${i.chaveParte}).`);

      itensLote.push({
        usuario: {
          id: String(i.usuarioId),
          telegramId: String(telegramId),
          telegramUsername: telegramUsername ? `@${String(telegramUsername).replace(/^@/, "")}` : "",
        },
        item: {
          categoria: "PARTE",
          chaveParte: i.chaveParte,
          totalLinhas: Number(i.totalLinhas || 0),
          tamanhoBytes: Number(i.tamanhoBytes || 0),
          csv: i.csv || "",
          linkQuestionario: i.linkQuestionario || "",
        },
      });
    }

    for (const i of itensSobrasSalvar) {
      const u = getUsuarioPorId(i.usuarioId);
      const { telegramId, telegramUsername } = obterTelegramDoUsuario(u || {});
      if (!telegramId) throw new Error(`Usuário ${i.usuarioId} está sem telegramId (SOBRA ${i.chaveParte}).`);

      itensLote.push({
        usuario: {
          id: String(i.usuarioId),
          telegramId: String(telegramId),
          telegramUsername: telegramUsername ? `@${String(telegramUsername).replace(/^@/, "")}` : "",
        },
        item: {
          categoria: "SOBRA",
          chaveParte: i.chaveParte,
          labelVisivel: i.labelVisivel || i.chaveParte.replace(/^S/, "P"),
          totalLinhas: Number(i.totalLinhas || 0),
          tamanhoBytes: Number(i.tamanhoBytes || 0),
          csv: i.csv || "",
          linkQuestionario: "",
        },
      });
    }

    const opPrefFinal = obterOpcoesPesquisaPreferenciais();

    const payloadLote = {
      arquivo: {
        arquivoKey: obterArquivoKeyAtual(),
        idArquivoOrigem: Number.isFinite(Number(idOrigemFinal)) ? Number(idOrigemFinal) : null,
        nomeArquivoOrigem: registroArquivoAtual?.nome || "arquivo",
        dataArquivo: registroArquivoAtual?.data || "",
        tamanhoArquivo: registroArquivoAtual?.tamanho || "",
        statusArquivo: registroArquivoAtual?.status || "",
      },
      admin: {
        id: String(admin?.id || ""),
      },
      opcoesPesquisa: {
        numeroPesquisa: opPrefFinal.numeroPesquisa || "",
        dataPesquisa: opPrefFinal.dataPesquisa || "",
        baseUrl: QUESTIONARIO_BASE_URL,
        fonte: opPrefFinal.fonte || "",
      },
      opcoes: {
        criarLancamento: true,
        salvarDistribuicao: true,
      },
      itens: itensLote,
    };

    const resp = await enviarLoteParaN8n(payloadLote);
    const lancamentoId = resp?.lancamentoId || resp?.id || null;

    alert(
      `✅ Lote enviado ao n8n!\n` +
      `Partes: ${itensPrincipaisSalvar.length}\n` +
      `Sobras enviadas: ${itensSobrasSalvar.length} (opcional)\n` +
      (lancamentoId ? `LancamentoId: ${lancamentoId}\n` : "")
    );

    botaoLancarPartes.disabled = false;
    if (textoDicaLancamento) textoDicaLancamento.textContent = "Lote enviado com sucesso.";
    atualizarEstadoBotaoLancar();
  } catch (e) {
    console.error(e);
    alert(`❌ Falha ao lançar/enviar lote.\n\n${e?.message || e}`);
    botaoLancarPartes.disabled = false;
    if (textoDicaLancamento) textoDicaLancamento.textContent = "Falha ao enviar.";
    atualizarEstadoBotaoLancar();
  }
}

/* =========================
   INICIALIZAÇÃO
========================= */
async function inicializar() {
  esconderAlerta();

  try {
    arquivoKeyAtual = obterArquivoKeyDaUrl();
    idArquivoAtual = obterIdNumericoDaUrl();

    if (!arquivoKeyAtual && !Number.isFinite(idArquivoAtual)) {
      throw new Error("URL sem key/id. Abra como: admArquivo.html?key=ARQ-... ou ?id=123");
    }

    if (textoDicaLancamento) textoDicaLancamento.textContent = "Carregando…";
    if (botaoLancarPartes) botaoLancarPartes.disabled = true;

    await carregarUsuarios();

    const { registro, blob } = await carregarArquivoPorKey(arquivoKeyAtual);
    registroArquivoAtual = {
      ...registro,
      colunasVisiveis: {
        estado: true,
        cidade: true,
        regiao: true,
        ...(registro?.colunasVisiveis || {}),
      },
    };

    await persistirOpcoesPesquisaSeVieramDaUrl();

    const linhasJson = await lerExcelParaJson(blob);
    const { principais, sobras } = separarPorPartesESobras(linhasJson, registroArquivoAtual.colunasVisiveis);

    partesGeradas = principais;
    sobrasGeradas = sobras;

    aplicarAutoMapPartesPrincipais();

    preencherCardsInfo();
    renderizarTabelaPartes();
    renderizarTabelaSobras();

    if (!usuariosDisponiveis.length) {
      mostrarAlerta("Nenhum usuário cadastrado. Cadastre usuários antes de distribuir.");
    } else if (!partesGeradas.length && !sobrasGeradas.length) {
      mostrarAlerta(`Nenhuma parte detectada. Verifique a coluna "${nomeColunaParte}" (P01, P02...).`);
    } else if (sobrasGeradas.length) {
      mostrarAlerta(`Detectei ${sobrasGeradas.length} sobra(s). Elas ficam no final e precisam de seleção manual.`);
    }

    atualizarEstadoBotaoLancar();

    if (obterFlagAutoLancar()) {
      const podeLancarAgora = botaoLancarPartes && botaoLancarPartes.disabled === false;

      if (podeLancarAgora) {
        await lancarPartes();
      } else {
        alert(
          "Auto-lançar: faltam seleções obrigatórias nas PARTES principais (P01, P02...).\nSelecione e clique em 'Lançar partes (geral)'.\n\n(Sobras são opcionais.)"
        );
      }
    }
  } catch (e) {
    console.error(e);
    const msg = String(e?.message || e || "");

    if (
      msg.includes("não encontrado na tabela") ||
      msg.includes("não localizei o conteúdo do Excel") ||
      msg.includes("storage")
    ) {
      mostrarAlerta(`Erro: ${msg}`);
    } else {
      mostrarAlerta(`Erro: ${msg}`);
    }

    if (textoDicaLancamento) textoDicaLancamento.textContent = "Erro ao carregar.";
  }
}

/* =========================
   EVENTOS
========================= */
if (botaoRecarregar) botaoRecarregar.addEventListener("click", async () => await inicializar());
if (botaoLancarPartes) botaoLancarPartes.addEventListener("click", async () => await lancarPartes());

// ✅ Boot robusto
function boot() {
  inicializar().catch((e) => {
    console.error("Falha no boot do admArquivo:", e);
    mostrarAlerta(`Erro ao iniciar: ${e?.message || e}`);
    if (textoDicaLancamento) textoDicaLancamento.textContent = "Erro ao carregar.";
  });
}

if (document.readyState === "loading") {
  window.addEventListener("DOMContentLoaded", boot);
} else {
  boot();
}