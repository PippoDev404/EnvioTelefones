// /5-js/dashboard.js
import "../5-js/db.js";
import { requireAuth, logout } from "../8-src/auth.js";
import { supabase } from "../8-src/supabaseClient.js";

// ============================
// AUTH
// ============================
await requireAuth({ redirectTo: "../7-login/login.html" });
document.querySelector("#btnLogout")?.addEventListener("click", () => logout());

// ownerId
const { data: authData, error: authErr } = await supabase.auth.getUser();
if (authErr) console.warn("supabase.auth.getUser erro:", authErr.message);

const ownerId = authData?.user?.id;
console.log("ownerId do dashboard:", ownerId);
if (!ownerId) console.warn("Sem ownerId (sem sessão?)");

// ============================
// CONFIG
// ============================
const maxDias = 7;

const LS_KEYS = {
  LAST_DASH_VIEW: "ibespe:lastDashboardView"
};

function getLastDashKey() {
  return `${LS_KEYS.LAST_DASH_VIEW}:${String(ownerId || "anon")}`;
}

function registrarUltimaVisualizacaoDashboard(arquivo) {
  try {
    const payload = {
      fileName: arquivo?.nome || "arquivo",
      fileKey: arquivo?.arquivoKey || arquivo?.arquivo_key || null,
      at: Date.now()
    };
    localStorage.setItem(getLastDashKey(), JSON.stringify(payload));
  } catch {}
}

// ============================
// DOM (fixos)
// ============================
const inputBusca = document.getElementById("inputBusca");
const selectStatus = document.getElementById("selectStatus");
const estadoVazio = document.getElementById("estadoVazio");

const checkboxExcluirTodos = document.getElementById("checkboxExcluirTodos");
const botaoExcluirSelecionados = document.getElementById("botaoExcluirSelecionados");
const textoSelecionados = document.getElementById("textoSelecionados");

// Modal confirmação
const overlayConfirmacao = document.getElementById("overlayConfirmacao");
const mensagemConfirmacao = document.getElementById("mensagemConfirmacao");
const botaoCancelarConfirmacao = document.getElementById("botaoCancelarConfirmacao");
const botaoOkConfirmacao = document.getElementById("botaoOkConfirmacao");

// Tabela (inicializa no iniciarDashboard)
let corpoTabelaArquivos = null;

let acaoConfirmacao = null;
let eventosTabelaVinculados = false;

async function iniciarDashboard() {
  corpoTabelaArquivos = document.getElementById("corpoTabelaArquivos");
  if (!corpoTabelaArquivos) {
    console.warn("Não achei #corpoTabelaArquivos no HTML.");
    return;
  }

  await atualizar();
}

if (document.readyState === "loading") {
  window.addEventListener("DOMContentLoaded", iniciarDashboard);
} else {
  iniciarDashboard();
}

// ============================
// HELPERS OWNER
// ============================
function obterOwnerIdDoArquivo(arquivo) {
  return String(
    arquivo?.ownerId ??
    arquivo?.owner_id ??
    arquivo?.authUserId ??
    arquivo?.userId ??
    ""
  ).trim();
}

function arquivoPertenceAoOwnerAtual(arquivo) {
  return !!arquivo && !!ownerId && obterOwnerIdDoArquivo(arquivo) === String(ownerId);
}

async function listarArquivosDoOwnerDb() {
  const lista = await window.ibeDb.listarArquivosDb();
  return (lista || [])
    .filter(arquivoPertenceAoOwnerAtual)
    .sort((a, b) => (b.dataCriacao || 0) - (a.dataCriacao || 0));
}

async function pegarArquivoAtualPorKey(arquivoKey) {
  const arquivo = await window.ibeDb.pegarArquivoPorKeyDb(String(arquivoKey || ""));
  if (!arquivoPertenceAoOwnerAtual(arquivo)) return null;
  return arquivo;
}

async function excluirArquivosAntigosDoOwnerDb(maxDiasParam) {
  const agora = Date.now();
  const maxMs = maxDiasParam * 24 * 60 * 60 * 1000;

  const lista = await listarArquivosDoOwnerDb();
  const keysParaExcluir = lista
    .filter((a) => a.dataCriacao && (agora - a.dataCriacao) > maxMs)
    .map((a) => a.arquivoKey)
    .filter(Boolean);

  if (keysParaExcluir.length) {
    await window.ibeDb.excluirArquivosPorKeyDb(keysParaExcluir);
  }

  return keysParaExcluir.length;
}

// ============================
// MODAL CONFIRMAÇÃO
// ============================
function abrirModalConfirmacao(texto, aoConfirmar) {
  if (mensagemConfirmacao) mensagemConfirmacao.textContent = texto;
  acaoConfirmacao = typeof aoConfirmar === "function" ? aoConfirmar : null;
  if (overlayConfirmacao) overlayConfirmacao.hidden = false;
}

function fecharModalConfirmacao() {
  if (overlayConfirmacao) overlayConfirmacao.hidden = true;
  acaoConfirmacao = null;
}

botaoCancelarConfirmacao?.addEventListener("click", fecharModalConfirmacao);
botaoOkConfirmacao?.addEventListener("click", () => {
  if (acaoConfirmacao) acaoConfirmacao();
  fecharModalConfirmacao();
});
overlayConfirmacao?.addEventListener("click", (e) => {
  if (e.target === overlayConfirmacao) fecharModalConfirmacao();
});
window.addEventListener("keydown", (e) => {
  if (overlayConfirmacao && !overlayConfirmacao.hidden && e.key === "Escape") fecharModalConfirmacao();
});

// ============================
// MODAL OPÇÕES
// ============================
const overlayOpcoes = document.getElementById("overlayOpcoes");
const botaoCancelarOpcoes = document.getElementById("botaoCancelarOpcoes");
const botaoSalvarOpcoes = document.getElementById("botaoSalvarOpcoes");

const inputNumeroPesquisa = document.getElementById("inputNumeroPesquisa");
const inputDataPesquisa = document.getElementById("inputDataPesquisa");
const linkPreviewPesquisa = document.getElementById("linkPreviewPesquisa");

const togglesColunas = document.querySelectorAll(".colunaToggle");

let arquivoKeyEmEdicao = null;
let estadoOpcoesTemp = { numeroPesquisa: "", dataPesquisa: "" };
let estadoColunasTemp = { estado: true, cidade: true, regiao: true };

function obterHojeBr() {
  const d = new Date();
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const yyyy = d.getFullYear();
  return `${dd}-${mm}-${yyyy}`;
}

function normalizarNumeroPesquisa(v) {
  return String(v || "").trim().replace(/[^\d]/g, "");
}

function normalizarDataPesquisa(v) {
  return String(v || "").trim().replaceAll("/", "-");
}

function dataPesquisaValida(ddmmyyyy) {
  return /^\d{2}-\d{2}-\d{4}$/.test(String(ddmmyyyy || ""));
}

function montarLinkBase({ numeroPesquisa, dataPesquisa, parte = "P01" }) {
  const base = "https://www.ibespebr.com.br/questionario/pesquisa/";
  const num = normalizarNumeroPesquisa(numeroPesquisa);
  const data = normalizarDataPesquisa(dataPesquisa);
  if (!num || !dataPesquisaValida(data)) return "";
  return `${base}${encodeURIComponent(num)}/${encodeURIComponent(parte)}/${encodeURIComponent(data)}`;
}

function atualizarPreviewOpcoes() {
  if (!linkPreviewPesquisa) return;

  const href = montarLinkBase({
    numeroPesquisa: estadoOpcoesTemp.numeroPesquisa,
    dataPesquisa: estadoOpcoesTemp.dataPesquisa,
    parte: "P01"
  });

  if (!href) {
    linkPreviewPesquisa.textContent = "Preencha número e data para ver o link.";
    linkPreviewPesquisa.removeAttribute("href");
    return;
  }

  linkPreviewPesquisa.textContent = href;
  linkPreviewPesquisa.setAttribute("href", href);
  linkPreviewPesquisa.setAttribute("target", "_blank");
  linkPreviewPesquisa.setAttribute("rel", "noopener");
}

function garantirColunasVisiveis(arquivo) {
  const cv = arquivo?.colunasVisiveis || {};
  return {
    estado: cv.estado !== false,
    cidade: cv.cidade !== false,
    regiao: cv.regiao !== false
  };
}

function atualizarVisualToggles() {
  togglesColunas.forEach((el) => {
    const chave = el.dataset.coluna;
    if (!chave) return;
    if (estadoColunasTemp[chave]) el.classList.add("ativa");
    else el.classList.remove("ativa");
  });
}

togglesColunas.forEach((el) => {
  el.addEventListener("click", () => {
    const chave = el.dataset.coluna;
    if (!chave) return;
    estadoColunasTemp[chave] = !estadoColunasTemp[chave];
    atualizarVisualToggles();
  });
});

function abrirModalOpcoes(arquivo) {
  arquivoKeyEmEdicao = String(arquivo?.arquivoKey || arquivo?.arquivo_key || "");

  const op = arquivo?.opcoesPesquisa || {};
  estadoOpcoesTemp = {
    numeroPesquisa: String(op.numeroPesquisa || ""),
    dataPesquisa: String(op.dataPesquisa || "") || obterHojeBr()
  };

  const col = garantirColunasVisiveis(arquivo);
  estadoColunasTemp = { estado: !!col.estado, cidade: !!col.cidade, regiao: !!col.regiao };

  if (inputNumeroPesquisa) inputNumeroPesquisa.value = estadoOpcoesTemp.numeroPesquisa;
  if (inputDataPesquisa) inputDataPesquisa.value = estadoOpcoesTemp.dataPesquisa;

  atualizarPreviewOpcoes();
  atualizarVisualToggles();

  if (overlayOpcoes) overlayOpcoes.hidden = false;
}

function fecharModalOpcoes() {
  if (overlayOpcoes) overlayOpcoes.hidden = true;
  arquivoKeyEmEdicao = null;
}

inputNumeroPesquisa?.addEventListener("input", () => {
  estadoOpcoesTemp.numeroPesquisa = normalizarNumeroPesquisa(inputNumeroPesquisa.value);
  inputNumeroPesquisa.value = estadoOpcoesTemp.numeroPesquisa;
  atualizarPreviewOpcoes();
});

inputDataPesquisa?.addEventListener("input", () => {
  estadoOpcoesTemp.dataPesquisa = normalizarDataPesquisa(inputDataPesquisa.value);
  atualizarPreviewOpcoes();
});

botaoCancelarOpcoes?.addEventListener("click", fecharModalOpcoes);
overlayOpcoes?.addEventListener("click", (e) => {
  if (e.target === overlayOpcoes) fecharModalOpcoes();
});
window.addEventListener("keydown", (e) => {
  if (overlayOpcoes && !overlayOpcoes.hidden && e.key === "Escape") fecharModalOpcoes();
});

botaoSalvarOpcoes?.addEventListener("click", async () => {
  if (!arquivoKeyEmEdicao) return;

  const numero = normalizarNumeroPesquisa(estadoOpcoesTemp.numeroPesquisa);
  const data = normalizarDataPesquisa(estadoOpcoesTemp.dataPesquisa);

  if (!numero) return alert("Informe o número da pesquisa (ex: 744).");
  if (!dataPesquisaValida(data)) return alert("Informe a data no formato DD-MM-AAAA (ex: 27-02-2026).");

  const arquivo = await pegarArquivoAtualPorKey(arquivoKeyEmEdicao);
  if (!arquivo) return;

  const atualizado = {
    ...arquivo,
    opcoesPesquisa: { numeroPesquisa: numero, dataPesquisa: data },
    colunasVisiveis: { ...estadoColunasTemp }
  };

  await window.ibeDb.salvarArquivoDb(atualizado);
  fecharModalOpcoes();
  await atualizar();
});

// ============================
// UI HELPERS
// ============================
function classeBadgeTipo(tipo) {
  const t = String(tipo || "").toLowerCase();
  if (t === "xlsx") return "badgeTipo badgeXlsx";
  if (t === "xls") return "badgeTipo badgeXls";
  if (t === "csv") return "badgeTipo badgeCsv";
  return "badgeTipo badgeOutro";
}

function classeBadgeStatus(status) {
  const s = String(status || "").toLowerCase();
  if (s === "processado") return "badgeStatus statusProcessado";
  if (s === "pendente") return "badgeStatus statusPendente";
  if (s === "erro") return "badgeStatus statusErro";
  return "badgeStatus";
}

function classeBolinhaStatus(status) {
  const s = String(status || "").toLowerCase();
  if (s === "processado") return "bolinhaTabela verde";
  if (s === "pendente") return "bolinhaTabela laranja";
  if (s === "erro") return "bolinhaTabela vermelha";
  return "bolinhaTabela laranja";
}

function normalizarStatus(status) {
  const s = String(status || "").toLowerCase();
  if (s === "processado") return "Processado";
  if (s === "pendente") return "Pendente";
  if (s === "erro") return "Erro";
  return "Pendente";
}

// ============================
// SELEÇÃO
// ============================
function pegarCheckboxesLinha() {
  return Array.from(corpoTabelaArquivos?.querySelectorAll(".checkboxLinha") || []);
}

function pegarCheckboxesVisiveis() {
  return Array.from(corpoTabelaArquivos?.querySelectorAll("tr") || [])
    .filter((tr) => tr.style.display !== "none")
    .map((tr) => tr.querySelector(".checkboxLinha"))
    .filter(Boolean);
}

function pegarKeysSelecionados() {
  return pegarCheckboxesLinha()
    .filter((cb) => cb.checked)
    .map((cb) => String(cb.dataset.arquivoKey || ""));
}

function atualizarUiSelecao() {
  const checkboxes = pegarCheckboxesLinha();
  const marcados = checkboxes.filter((cb) => cb.checked).length;

  if (textoSelecionados) textoSelecionados.textContent = `${marcados} selecionado(s)`;
  if (botaoExcluirSelecionados) botaoExcluirSelecionados.disabled = marcados === 0;

  const visiveis = pegarCheckboxesVisiveis();
  const marcadosVisiveis = visiveis.filter((cb) => cb.checked).length;

  if (!checkboxExcluirTodos) return;

  if (visiveis.length === 0) {
    checkboxExcluirTodos.checked = false;
    checkboxExcluirTodos.indeterminate = false;
    return;
  }

  checkboxExcluirTodos.checked = marcadosVisiveis === visiveis.length;
  checkboxExcluirTodos.indeterminate = marcadosVisiveis > 0 && marcadosVisiveis < visiveis.length;
}

// ============================
// FILTROS
// ============================
function aplicarFiltros() {
  const termo = (inputBusca?.value || "").trim().toLowerCase();
  const status = (selectStatus?.value || "").trim().toLowerCase();

  const linhas = Array.from(corpoTabelaArquivos?.querySelectorAll("tr") || []);
  let visiveis = 0;

  for (const linha of linhas) {
    const nome = linha.querySelector("[data-col='nome']")?.textContent?.toLowerCase() || "";
    const statusLinha = linha.querySelector("[data-col='status']")?.textContent?.toLowerCase() || "";

    let mostrar = true;
    if (termo && !nome.includes(termo)) mostrar = false;
    if (status && statusLinha && statusLinha !== status) mostrar = false;

    linha.style.display = mostrar ? "" : "none";
    if (mostrar) visiveis += 1;
  }

  if (estadoVazio) estadoVazio.hidden = visiveis !== 0;
  atualizarUiSelecao();
}

inputBusca?.addEventListener("input", aplicarFiltros);
selectStatus?.addEventListener("change", aplicarFiltros);

// ============================
// EXCLUIR (INDEXEDDB + SUPABASE)
// ============================
async function excluirNoSupabasePorKeys(keysParaExcluir) {
  const keys = (keysParaExcluir || []).map((k) => String(k || "").trim()).filter(Boolean);
  if (!keys.length) return true;
  if (!ownerId) {
    console.warn("Sem ownerId: não dá pra excluir no Supabase.");
    return false;
  }

  const { error } = await supabase
    .from("arquivos")
    .delete()
    .eq("owner_id", ownerId)
    .in("arquivo_key", keys);

  if (error) {
    console.warn("Erro ao excluir no Supabase:", error.message);
    return false;
  }
  return true;
}

async function excluirNoIndexedDbPorKeys(keysParaExcluir) {
  const keys = (keysParaExcluir || []).map((k) => String(k || "").trim()).filter(Boolean);
  if (!keys.length) return true;
  await window.ibeDb.excluirArquivosPorKeyDb(keys);
  return true;
}

async function excluirPorKeys(keysParaExcluir) {
  const keys = (keysParaExcluir || []).map((k) => String(k || "").trim()).filter(Boolean);
  if (!keys.length) return;

  const okSupabase = await excluirNoSupabasePorKeys(keys);
  await excluirNoIndexedDbPorKeys(keys);

  if (!okSupabase) {
    alert("Arquivos removidos localmente, mas houve erro ao excluir no Supabase. Veja o console.");
  }
}

async function excluirArquivoUnico(arquivoKey) {
  const arquivo = await pegarArquivoAtualPorKey(arquivoKey);
  if (!arquivo) return;

  abrirModalConfirmacao(`Excluir "${arquivo.nome}"? Esta ação não pode ser desfeita.`, async () => {
    await excluirPorKeys([String(arquivoKey)]);
    await atualizar();
  });
}

function excluirSelecionados() {
  const keys = pegarKeysSelecionados().filter(Boolean);
  if (!keys.length) return;

  abrirModalConfirmacao(
    `Excluir ${keys.length} arquivo(s) selecionado(s)? Esta ação não pode ser desfeita.`,
    async () => {
      await excluirPorKeys(keys);
      await atualizar();
    }
  );
}

checkboxExcluirTodos?.addEventListener("change", () => {
  const visiveis = pegarCheckboxesVisiveis();
  for (const cb of visiveis) cb.checked = checkboxExcluirTodos.checked;
  atualizarUiSelecao();
});

botaoExcluirSelecionados?.addEventListener("click", excluirSelecionados);

// ============================
// RENDER
// ============================
function vincularEventosTabelaUmaVez() {
  if (eventosTabelaVinculados) return;
  if (!corpoTabelaArquivos) return;

  corpoTabelaArquivos.addEventListener("click", aoClicarTabela);
  corpoTabelaArquivos.addEventListener("change", aoAlterarTabela);

  eventosTabelaVinculados = true;
}

function renderizarTabela(arquivos) {
  if (!corpoTabelaArquivos) return;

  corpoTabelaArquivos.innerHTML = "";

  if (!arquivos.length) {
    if (estadoVazio) estadoVazio.hidden = false;
    atualizarUiSelecao();
    return;
  }

  if (estadoVazio) estadoVazio.hidden = true;

  for (const arquivoBruto of arquivos) {
    const arquivo = { ...arquivoBruto, status: normalizarStatus(arquivoBruto.status) };
    const arquivoKey = String(arquivo.arquivoKey || arquivo.arquivo_key || "");

    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td class="colunaExcluir">
        <div class="divExcluir">
          <p>Excluir: <input type="checkbox" class="checkboxLinha" data-arquivo-key="${arquivoKey}" /><p/>
        </div>
      </td>

      <td data-col="nome">${arquivo.nome || "-"}</td>

      <td>
        <span class="${classeBadgeTipo(arquivo.tipo)}">${String(arquivo.tipo || "-").toUpperCase()}</span>
      </td>

      <td>${arquivo.tamanho || "-"}</td>

      <td>${arquivo.data || "-"}</td>

      <td style="display:none">
        <span class="${classeBadgeStatus(arquivo.status)}" data-col="status">${arquivo.status}</span>
      </td>

      <td class="tdAcoes">
        <div class="acoesLinha">
          <button class="botaoAcao" data-acao="opcoes" data-arquivo-key="${arquivoKey}">⚙️ Opções</button>
          <button class="botaoAcao" data-acao="lancar" data-arquivo-key="${arquivoKey}">🚀 Lançar</button>
          <button class="botaoAcao" data-acao="ver" data-arquivo-key="${arquivoKey}">👁 Ver</button>
          <button class="botaoAcao" data-acao="excluir" data-arquivo-key="${arquivoKey}">🗑 Excluir</button>
        </div>
      </td>
    `;

    corpoTabelaArquivos.appendChild(tr);
  }

  atualizarUiSelecao();
}

async function marcarComoProcessadoSePendente(arquivoKey) {
  const arquivo = await pegarArquivoAtualPorKey(arquivoKey);
  if (!arquivo) return null;

  const statusAtual = normalizarStatus(arquivo.status);
  if (statusAtual.toLowerCase() !== "pendente") return arquivo;

  const atualizado = { ...arquivo, status: "Processado" };
  await window.ibeDb.salvarArquivoDb(atualizado);
  return atualizado;
}

async function aoClicarTabela(e) {
  const botao = e.target.closest("button");
  if (!botao) return;

  const acao = String(botao.dataset.acao || "").toLowerCase();
  const arquivoKey = String(botao.dataset.arquivoKey || "");
  if (!arquivoKey) return;

  if (acao === "opcoes") {
    const arquivo = await pegarArquivoAtualPorKey(arquivoKey);
    if (!arquivo) return;
    abrirModalOpcoes(arquivo);
    return;
  }

  if (acao === "excluir") {
    await excluirArquivoUnico(arquivoKey);
    return;
  }

  if (acao === "ver") {
    try {
      const arq = await pegarArquivoAtualPorKey(arquivoKey);
      if (arq) registrarUltimaVisualizacaoDashboard(arq);
    } catch {}

    try {
      await marcarComoProcessadoSePendente(arquivoKey);
    } catch {}

    window.location.href = `/4-html/admArquivo.html?key=${encodeURIComponent(arquivoKey)}`;
    return;
  }

  if (acao === "lancar") {
    try {
      const arq = await pegarArquivoAtualPorKey(arquivoKey);
      if (arq) registrarUltimaVisualizacaoDashboard(arq);
    } catch {}

    try {
      await marcarComoProcessadoSePendente(arquivoKey);
    } catch {}

    const arquivo = await pegarArquivoAtualPorKey(arquivoKey);
    const op = arquivo?.opcoesPesquisa || {};
    const numeroPesquisa = normalizarNumeroPesquisa(op.numeroPesquisa || "");
    const dataPesquisa = normalizarDataPesquisa(op.dataPesquisa || "");

    if (!numeroPesquisa || !dataPesquisaValida(dataPesquisa)) {
      alert("Antes de lançar, vá em ⚙️ Opções e preencha Número da pesquisa e Data (DD-MM-AAAA).");
      return;
    }

    const qs = new URLSearchParams({
      key: arquivoKey,
      autoLancar: "1",
      numeroPesquisa,
      dataPesquisa
    });

    window.location.href = `/4-html/admArquivo.html?${qs.toString()}`;
    return;
  }
}

function aoAlterarTabela(e) {
  const checkbox = e.target.closest(".checkboxLinha");
  if (!checkbox) return;
  atualizarUiSelecao();
}

// ============================
// SYNC SUPABASE -> INDEXEDDB
// ============================
function bytesParaTamanho(bytes) {
  const n = Number(bytes || 0);
  if (!n) return "0 B";
  const k = 1024;
  const sizes = ["B", "KB", "MB", "GB", "TB"];
  const i = Math.floor(Math.log(n) / Math.log(k));
  const val = (n / Math.pow(k, i)).toFixed(i === 0 ? 0 : 1);
  return `${val} ${sizes[i]}`;
}

function formatarDataISO(iso) {
  if (!iso) return "";
  const d = new Date(iso);
  if (isNaN(d.getTime())) return "";
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const yyyy = d.getFullYear();
  const hh = String(d.getHours()).padStart(2, "0");
  const mi = String(d.getMinutes()).padStart(2, "0");
  return `${dd}/${mm}/${yyyy} ${hh}:${mi}`;
}

function inferirTipo({ mime_type, storage_path, nome }) {
  const m = String(mime_type || "").toLowerCase();
  if (m.includes("spreadsheet") || m.includes("excel")) return "xlsx";
  if (m.includes("csv")) return "csv";
  const base = String(storage_path || nome || "").toLowerCase();
  const ext = base.split(".").pop();
  if (ext && ext.length <= 5) return ext;
  return "arquivo";
}

function parseJsonPossivel(v, fallback) {
  if (v == null) return fallback;
  if (typeof v === "object") return v;
  if (typeof v !== "string") return fallback;
  const s = v.trim();
  if (!s) return fallback;
  try {
    return JSON.parse(s);
  } catch {
    return fallback;
  }
}

async function esperarIbeDb({ tentativas = 80, delayMs = 50 } = {}) {
  for (let i = 0; i < tentativas; i++) {
    if (window.ibeDb?.listarArquivosDb && window.ibeDb?.salvarArquivoDb && window.ibeDb?.pegarArquivoPorKeyDb) return true;
    await new Promise((r) => setTimeout(r, delayMs));
  }
  return false;
}

async function syncArquivosSupabaseParaIndexedDb() {
  if (!ownerId) return;

  const ok = await esperarIbeDb();
  if (!ok) {
    console.warn("window.ibeDb não ficou pronto a tempo.");
    return;
  }

  const { data: arquivos, error } = await supabase
    .from("arquivos")
    .select("*")
    .eq("owner_id", ownerId)
    .order("data_criacao", { ascending: false });

  if (error) {
    console.warn("Erro ao buscar arquivos no Supabase:", error.message);
    return;
  }

  console.log("SUPABASE arquivos do owner:", arquivos);

  const fetchedKeys = new Set();

  if (Array.isArray(arquivos)) {
    for (const row of arquivos) {
      const arquivoKey = String(row?.arquivo_key || "").trim();
      if (!arquivoKey) continue;

      fetchedKeys.add(arquivoKey);

      const existente = await window.ibeDb.pegarArquivoPorKeyDb(arquivoKey);

      const dataCriacaoIso = row?.data_criacao || row?.atualizado_em || null;
      const dataCriacaoNum = dataCriacaoIso ? new Date(dataCriacaoIso).getTime() : Date.now();

      const mapped = {
        arquivoKey,
        arquivo_key: arquivoKey,

        ownerId: row?.owner_id || ownerId,
        owner_id: row?.owner_id || ownerId,
        idx: row?.idx ?? null,

        nome: row?.nome || "Arquivo",
        descricao: row?.descricao || "",

        mimeType: row?.mime_type || null,
        tipo: inferirTipo({
          mime_type: row?.mime_type,
          storage_path: row?.storage_path,
          nome: row?.nome
        }),

        tamanhoBytes: Number(row?.tamanho_bytes || 0),
        tamanho: bytesParaTamanho(row?.tamanho_bytes),

        dataCriacao: isNaN(dataCriacaoNum) ? Date.now() : dataCriacaoNum,
        data: formatarDataISO(dataCriacaoIso),

        status: existente?.status || "Pendente",

        colunasVisiveis: parseJsonPossivel(row?.colunas_visiveis, { estado: true, cidade: true, regiao: true }),

        storageBucket: row?.storage_bucket || null,
        storagePath: row?.storage_path || null,
        arquivoUrl: row?.arquivo_url || null,

        atualizadoEm: row?.atualizado_em || null,
        hash: row?.hash || null
      };

      await window.ibeDb.salvarArquivoDb({
        ...(existente || {}),
        ...mapped,
        opcoesPesquisa: existente?.opcoesPesquisa || mapped?.opcoesPesquisa || undefined
      });
    }
  }

  // remove do IndexedDB local apenas os arquivos do owner atual
  // que não existem mais no Supabase
  const locaisDoOwner = await listarArquivosDoOwnerDb();
  const staleKeys = locaisDoOwner
    .filter((a) => !fetchedKeys.has(String(a.arquivoKey || a.arquivo_key || "")))
    .map((a) => String(a.arquivoKey || a.arquivo_key || ""))
    .filter(Boolean);

  if (staleKeys.length) {
    await window.ibeDb.excluirArquivosPorKeyDb(staleKeys);
  }
}

// ============================
// ATUALIZAR
// ============================
async function atualizar() {
  const ok = await esperarIbeDb();
  if (!ok) {
    if (estadoVazio) estadoVazio.hidden = false;
    if (corpoTabelaArquivos) corpoTabelaArquivos.innerHTML = "";
    return;
  }

  await syncArquivosSupabaseParaIndexedDb();
  await excluirArquivosAntigosDoOwnerDb(maxDias);

  const arquivos = await listarArquivosDoOwnerDb();

  vincularEventosTabelaUmaVez();
  renderizarTabela(arquivos);
  aplicarFiltros();
}