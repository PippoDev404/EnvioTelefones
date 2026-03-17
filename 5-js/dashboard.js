// /5-js/dashboard.js
import "../5-js/db.js";
import { requireAuth, logout } from "../8-src/auth.js";
import { supabase } from "../8-src/supabaseClient.js";

// ============================
// AUTH
// ============================
await requireAuth({ redirectTo: "../7-login/login.html" });
document.querySelector("#btnLogout")?.addEventListener("click", () => logout());

const { data: authData, error: authErr } = await supabase.auth.getUser();
if (authErr) console.warn("supabase.auth.getUser erro:", authErr.message);

const ownerId = authData?.user?.id;
console.log("ownerId do dashboard:", ownerId);
if (!ownerId) console.warn("Sem ownerId (sem sessão?)");

// ============================
// CONFIG
// ============================
const DIAS_RETENCAO_ARQUIVOS = 7;
const SUPABASE_ARQUIVOS_TABLE = "arquivos";
const SUPABASE_STORAGE_BUCKET = "arquivos";

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

// Tabela
let corpoTabelaArquivos = null;

let acaoConfirmacao = null;
let eventosTabelaVinculados = false;

// cache em memória da listagem atual vinda do Supabase
let arquivosAtuais = [];

// trava simples
let dashboardBusy = false;

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
// HELPERS GERAIS
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

function normalizarStatus(status) {
  const s = String(status || "").toLowerCase();
  if (s === "processado") return "Processado";
  if (s === "pendente") return "Pendente";
  if (s === "erro") return "Erro";
  return "Pendente";
}

function obterArquivoKey(arquivo) {
  return String(arquivo?.arquivoKey || arquivo?.arquivo_key || "").trim();
}

function obterDataBaseArquivo(arquivo) {
  return (
    arquivo?.data_criacao ||
    arquivo?.atualizado_em ||
    arquivo?.created_at ||
    arquivo?.updated_at ||
    null
  );
}

function arquivoExpirado(arquivo, dias = DIAS_RETENCAO_ARQUIVOS) {
  const iso = obterDataBaseArquivo(arquivo);
  if (!iso) return false;

  const t = new Date(iso).getTime();
  if (Number.isNaN(t)) return false;

  const limiteMs = dias * 24 * 60 * 60 * 1000;
  return (Date.now() - t) > limiteMs;
}

function mapearArquivoSupabaseParaUI(row) {
  const arquivoKey = String(row?.arquivo_key || "").trim();
  const dataBase = obterDataBaseArquivo(row);
  const dataCriacaoNum = dataBase ? new Date(dataBase).getTime() : Date.now();

  return {
    ...row,

    arquivoKey,
    arquivo_key: arquivoKey,

    ownerId: row?.owner_id || ownerId,
    owner_id: row?.owner_id || ownerId,

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

    dataCriacao: Number.isNaN(dataCriacaoNum) ? Date.now() : dataCriacaoNum,
    data: formatarDataISO(dataBase),

    status: normalizarStatus(row?.status),

    opcoesPesquisa: parseJsonPossivel(row?.opcoes_pesquisa, {}),
    colunasVisiveis: parseJsonPossivel(row?.colunas_visiveis, { estado: true, cidade: true, regiao: true }),

    storageBucket: row?.storage_bucket || SUPABASE_STORAGE_BUCKET,
    storagePath: row?.storage_path || null,
    arquivoUrl: row?.arquivo_url || null,

    atualizadoEm: row?.atualizado_em || null,
    hash: row?.hash || null
  };
}

async function esperarIbeDb({ tentativas = 80, delayMs = 50 } = {}) {
  for (let i = 0; i < tentativas; i++) {
    if (window.ibeDb?.listarArquivosDb && window.ibeDb?.salvarArquivoDb && window.ibeDb?.pegarArquivoPorKeyDb) {
      return true;
    }
    await new Promise((r) => setTimeout(r, delayMs));
  }
  return false;
}

function setDashboardBusy(flag) {
  dashboardBusy = !!flag;
}

// ============================
// INDEXEDDB AUXILIAR
// ============================
async function listarArquivosDoOwnerDb() {
  if (!window.ibeDb?.listarArquivosDb) return [];
  const lista = await window.ibeDb.listarArquivosDb();
  return (lista || [])
    .filter((arquivo) => String(arquivo?.ownerId || arquivo?.owner_id || "") === String(ownerId || ""))
    .sort((a, b) => (b.dataCriacao || 0) - (a.dataCriacao || 0));
}

async function pegarArquivoAtualPorKey(arquivoKey) {
  const key = String(arquivoKey || "").trim();
  if (!key) return null;

  const emMemoria = arquivosAtuais.find((a) => obterArquivoKey(a) === key);
  if (emMemoria) return emMemoria;

  if (window.ibeDb?.pegarArquivoPorKeyDb) {
    const arquivoLocal = await window.ibeDb.pegarArquivoPorKeyDb(key);
    if (arquivoLocal) return arquivoLocal;
  }

  return null;
}

async function salvarArquivoLocalCache(arquivo) {
  if (!window.ibeDb?.salvarArquivoDb) return;
  await window.ibeDb.salvarArquivoDb(arquivo);
}

async function excluirArquivosLocaisPorKeys(keys) {
  const chaves = (keys || []).map((k) => String(k || "").trim()).filter(Boolean);
  if (!chaves.length) return;
  if (!window.ibeDb?.excluirArquivosPorKeyDb) return;
  await window.ibeDb.excluirArquivosPorKeyDb(chaves);
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
  arquivoKeyEmEdicao = obterArquivoKey(arquivo);

  const op = arquivo?.opcoesPesquisa || {};
  estadoOpcoesTemp = {
    numeroPesquisa: String(op.numeroPesquisa || ""),
    dataPesquisa: String(op.dataPesquisa || "") || obterHojeBr()
  };

  const col = garantirColunasVisiveis(arquivo);
  estadoColunasTemp = {
    estado: !!col.estado,
    cidade: !!col.cidade,
    regiao: !!col.regiao
  };

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

  const patch = {
    opcoes_pesquisa: {
      numeroPesquisa: numero,
      dataPesquisa: data
    },
    colunas_visiveis: {
      ...estadoColunasTemp
    },
    atualizado_em: new Date().toISOString()
  };

  const { error } = await supabase
    .from(SUPABASE_ARQUIVOS_TABLE)
    .update(patch)
    .eq("owner_id", ownerId)
    .eq("arquivo_key", arquivoKeyEmEdicao);

  if (error) {
    console.error(error);
    alert(`Falha ao salvar opções no Supabase: ${error.message}`);
    return;
  }

  const atualizadoLocal = {
    ...arquivo,
    opcoesPesquisa: { numeroPesquisa: numero, dataPesquisa: data },
    colunasVisiveis: { ...estadoColunasTemp }
  };

  await salvarArquivoLocalCache(atualizadoLocal);
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
// SUPABASE - LISTAR / EXCLUIR
// ============================
async function listarArquivosDoOwnerNoSupabase() {
  if (!ownerId) return [];

  const { data, error } = await supabase
    .from(SUPABASE_ARQUIVOS_TABLE)
    .select("*")
    .eq("owner_id", ownerId)
    .order("data_criacao", { ascending: false });

  if (error) {
    console.error("Erro ao buscar arquivos no Supabase:", error.message);
    return [];
  }

  return (data || []).map(mapearArquivoSupabaseParaUI);
}

async function excluirArquivosDoBucket(paths) {
  const lista = (paths || []).map((p) => String(p || "").trim()).filter(Boolean);
  if (!lista.length) return true;

  const { error } = await supabase.storage
    .from(SUPABASE_STORAGE_BUCKET)
    .remove(lista);

  if (error) {
    console.warn("Erro ao excluir no bucket:", error.message);
    return false;
  }

  return true;
}

async function excluirRegistrosDaTabelaPorKeys(keys) {
  const lista = (keys || []).map((k) => String(k || "").trim()).filter(Boolean);
  if (!lista.length) return true;

  const { error } = await supabase
    .from(SUPABASE_ARQUIVOS_TABLE)
    .delete()
    .eq("owner_id", ownerId)
    .in("arquivo_key", lista);

  if (error) {
    console.warn("Erro ao excluir registros da tabela:", error.message);
    return false;
  }

  return true;
}

async function excluirPorKeys(keysParaExcluir) {
  const keys = (keysParaExcluir || []).map((k) => String(k || "").trim()).filter(Boolean);
  if (!keys.length) return;

  const arquivosParaExcluir = arquivosAtuais.filter((a) => keys.includes(obterArquivoKey(a)));
  const storagePaths = arquivosParaExcluir.map((a) => a.storagePath).filter(Boolean);

  const okBucket = await excluirArquivosDoBucket(storagePaths);
  const okTabela = await excluirRegistrosDaTabelaPorKeys(keys);

  await excluirArquivosLocaisPorKeys(keys);

  if (!okBucket || !okTabela) {
    alert("Houve falha parcial ao excluir no Supabase/Storage. Veja o console.");
  }
}

async function excluirArquivosExpirados() {
  const lista = await listarArquivosDoOwnerNoSupabase();
  const expirados = lista.filter((a) => arquivoExpirado(a, DIAS_RETENCAO_ARQUIVOS));

  if (!expirados.length) return 0;

  const keys = expirados.map((a) => obterArquivoKey(a)).filter(Boolean);
  const paths = expirados.map((a) => a.storagePath).filter(Boolean);

  console.log(`Excluindo ${expirados.length} arquivo(s) expirado(s) com mais de ${DIAS_RETENCAO_ARQUIVOS} dia(s).`);

  const okBucket = await excluirArquivosDoBucket(paths);
  const okTabela = await excluirRegistrosDaTabelaPorKeys(keys);
  await excluirArquivosLocaisPorKeys(keys);

  if (!okBucket || !okTabela) {
    console.warn("Falha parcial ao excluir arquivos expirados.");
  }

  return expirados.length;
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
    const arquivoKey = obterArquivoKey(arquivo);

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

  const { error } = await supabase
    .from(SUPABASE_ARQUIVOS_TABLE)
    .update({
      status: "Processado",
      atualizado_em: new Date().toISOString()
    })
    .eq("owner_id", ownerId)
    .eq("arquivo_key", arquivoKey);

  if (error) {
    console.warn("Falha ao atualizar status para Processado:", error.message);
    return arquivo;
  }

  const atualizado = { ...arquivo, status: "Processado" };
  await salvarArquivoLocalCache(atualizado);
  return atualizado;
}

async function aoClicarTabela(e) {
  const botao = e.target.closest("button");
  if (!botao) return;
  if (dashboardBusy) return;

  const acao = String(botao.dataset.acao || "").toLowerCase();
  const arquivoKey = String(botao.dataset.arquivoKey || "");
  if (!arquivoKey) return;

  setDashboardBusy(true);

  try {
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
  } finally {
    setDashboardBusy(false);
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
async function syncArquivosSupabaseParaIndexedDb(listaSupabase) {
  const ok = await esperarIbeDb();
  if (!ok) {
    console.warn("window.ibeDb não ficou pronto a tempo.");
    return;
  }

  const fetchedKeys = new Set();

  for (const arquivo of listaSupabase || []) {
    const key = obterArquivoKey(arquivo);
    if (!key) continue;

    fetchedKeys.add(key);

    const existente = await window.ibeDb.pegarArquivoPorKeyDb(key);

    await window.ibeDb.salvarArquivoDb({
      ...(existente || {}),
      ...arquivo,
      opcoesPesquisa: arquivo.opcoesPesquisa || existente?.opcoesPesquisa || {},
      colunasVisiveis: arquivo.colunasVisiveis || existente?.colunasVisiveis || { estado: true, cidade: true, regiao: true }
    });
  }

  const locaisDoOwner = await listarArquivosDoOwnerDb();
  const staleKeys = locaisDoOwner
    .filter((a) => !fetchedKeys.has(obterArquivoKey(a)))
    .map((a) => obterArquivoKey(a))
    .filter(Boolean);

  if (staleKeys.length) {
    await excluirArquivosLocaisPorKeys(staleKeys);
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

  await excluirArquivosExpirados();

  const listaSupabase = await listarArquivosDoOwnerNoSupabase();
  arquivosAtuais = [...listaSupabase];

  await syncArquivosSupabaseParaIndexedDb(listaSupabase);

  vincularEventosTabelaUmaVez();
  renderizarTabela(listaSupabase);
  aplicarFiltros();
}