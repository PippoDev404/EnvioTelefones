// /5-js/admArquivo.js
import { requireAuth, logout } from "../8-src/auth.js";
import { supabase } from "../8-src/supabaseClient.js";

const user = await requireAuth({ redirectTo: "../7-login/login.html" });
if (!user) throw new Error("Sem sessão");

document.querySelector("#btnLogout")?.addEventListener("click", () => logout());

/* =========================
   CONFIG
========================= */
const SUPABASE_ARQUIVOS_TABLE = "arquivos";
const SUPABASE_STORAGE_BUCKET = "arquivos";
const SUPABASE_ENTREGAS_CONSOLIDADO_TABLE = "entregas_consolidado";

const N8N_WEBHOOK_URL_LOTE = "https://n8n.srv962474.hstgr.cloud/webhook/envioArquivos";
const N8N_WEBHOOK_URL_INDIVIDUAL = "https://n8n.srv962474.hstgr.cloud/webhook/envioArquivos";
const N8N_WEBHOOK_URL_EXCLUIR = "https://n8n.srv962474.hstgr.cloud/webhook/excluir";

const FETCH_TIMEOUT_MS = 100000;
const nomeColunaParte = "Nº PESQ.";
const maxLinhasPreview = 80;

const SESSION_AUTO_LANCAR_PREFIX = "ibespe:autoLancarExecutado:";

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
const botaoDownloadGeralCsv = document.getElementById("botaoDownloadGeralCsv");
const botaoDownloadGeralPdf = document.getElementById("botaoDownloadGeralPdf");
const botaoExcluirMensagensTelegram = document.getElementById("botaoExcluirMensagensTelegram");
const botaoLancarPartes = document.getElementById("botaoLancarPartes");
const textoDicaLancamento = document.getElementById("textoDicaLancamento");

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

let usuariosDisponiveis = [];
let partesGeradas = [];
let sobrasGeradas = [];

let selecaoUsuarioPorParte = new Map();
let selecaoUsuarioPorSobra = new Map();

let chaveParteAbertaNoModal = null;

let globalBusy = false;
let globalBusyReason = "";
let actionButtonsCache = [];

/* =========================
   MODAIS GERAIS (feedback/confirm)
========================= */
let modalUi = null;
let modalResolveFn = null;

function garantirModalUi() {
  if (modalUi) return modalUi;

  const style = document.createElement("style");
  style.id = "admArquivoModalUiStyle";
  style.textContent = `
    .ibe-modal-overlay{
      position:fixed;
      inset:0;
      background:rgba(0,0,0,.45);
      display:flex;
      align-items:center;
      justify-content:center;
      padding:16px;
      z-index:99999;
    }
    .ibe-modal-overlay[hidden]{
      display:none !important;
    }
    .ibe-modal-box{
      width:min(460px, 95vw);
      background:#fff;
      border:1px solid #ddd;
      border-radius:16px;
      box-shadow:0 18px 50px rgba(0,0,0,.20);
      overflow:hidden;
      font-family:Arial, Helvetica, sans-serif;
    }
    .ibe-modal-header{
      padding:14px 16px;
      border-bottom:1px solid #eee;
      background:#f7f7f7;
    }
    .ibe-modal-title{
      font-size:16px;
      font-weight:800;
      color:#111;
    }
    .ibe-modal-body{
      padding:16px;
      color:#222;
      font-size:14px;
      line-height:1.5;
      white-space:pre-wrap;
      word-break:break-word;
    }
    .ibe-modal-footer{
      padding:14px 16px;
      border-top:1px solid #eee;
      display:flex;
      justify-content:flex-end;
      gap:10px;
      flex-wrap:wrap;
      background:#fafafa;
    }
    .ibe-modal-btn{
      border:1px solid #d0d0d0;
      background:#fff;
      color:#111;
      border-radius:10px;
      padding:10px 14px;
      cursor:pointer;
      font-weight:700;
      font-size:13px;
    }
    .ibe-modal-btn:disabled{
      opacity:.6;
      cursor:not-allowed;
    }
    .ibe-modal-btn-primary{
      background:#111;
      color:#fff;
      border-color:#111;
    }
    .ibe-modal-btn-danger{
      background:#c62828;
      color:#fff;
      border-color:#c62828;
    }
  `;
  document.head.appendChild(style);

  const overlay = document.createElement("div");
  overlay.className = "ibe-modal-overlay";
  overlay.hidden = true;
  overlay.innerHTML = `
    <div class="ibe-modal-box" role="dialog" aria-modal="true" aria-labelledby="ibeModalTitle">
      <div class="ibe-modal-header">
        <div class="ibe-modal-title" id="ibeModalTitle">Mensagem</div>
      </div>
      <div class="ibe-modal-body" id="ibeModalBody"></div>
      <div class="ibe-modal-footer" id="ibeModalFooter"></div>
    </div>
  `;
  document.body.appendChild(overlay);

  overlay.addEventListener("click", (e) => {
    if (e.target === overlay && modalResolveFn) {
      const fn = modalResolveFn;
      modalResolveFn = null;
      overlay.hidden = true;
      fn(false);
    }
  });

  window.addEventListener("keydown", (e) => {
    if (!overlay.hidden && e.key === "Escape" && modalResolveFn) {
      const fn = modalResolveFn;
      modalResolveFn = null;
      overlay.hidden = true;
      fn(false);
    }
  });

  modalUi = {
    overlay,
    title: overlay.querySelector("#ibeModalTitle"),
    body: overlay.querySelector("#ibeModalBody"),
    footer: overlay.querySelector("#ibeModalFooter"),
  };

  return modalUi;
}

function fecharModalGeral(resultado = false) {
  const ui = garantirModalUi();
  ui.overlay.hidden = true;

  if (modalResolveFn) {
    const fn = modalResolveFn;
    modalResolveFn = null;
    fn(resultado);
  }
}

function abrirModalBase({ titulo = "Mensagem", mensagem = "", botoes = [] }) {
  const ui = garantirModalUi();

  ui.title.textContent = titulo;
  ui.body.textContent = mensagem;
  ui.footer.innerHTML = "";

  for (const botao of botoes) {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = `ibe-modal-btn ${botao.classe || ""}`.trim();
    btn.textContent = botao.texto || "OK";
    btn.disabled = !!botao.disabled;
    btn.addEventListener("click", () => {
      if (typeof botao.onClick === "function") botao.onClick();
    });
    ui.footer.appendChild(btn);
  }

  ui.overlay.hidden = false;
}

function modalInfo(titulo, mensagem) {
  return new Promise((resolve) => {
    modalResolveFn = resolve;
    abrirModalBase({
      titulo,
      mensagem,
      botoes: [
        {
          texto: "OK",
          classe: "ibe-modal-btn-primary",
          onClick: () => fecharModalGeral(true),
        },
      ],
    });
  });
}

function modalConfirm(titulo, mensagem, { textoConfirmar = "Confirmar", perigo = false } = {}) {
  return new Promise((resolve) => {
    modalResolveFn = resolve;
    abrirModalBase({
      titulo,
      mensagem,
      botoes: [
        {
          texto: "Cancelar",
          onClick: () => fecharModalGeral(false),
        },
        {
          texto: textoConfirmar,
          classe: perigo ? "ibe-modal-btn-danger" : "ibe-modal-btn-primary",
          onClick: () => fecharModalGeral(true),
        },
      ],
    });
  });
}

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

function obterOwnerIdAtual() {
  return String(user?.id || "").trim();
}

async function getAccessTokenOrThrow() {
  const { data, error } = await supabase.auth.getSession();

  if (error) throw new Error(error.message || "Não foi possível obter a sessão.");

  const token = data?.session?.access_token;
  if (!token) throw new Error("Usuário sem sessão ativa.");

  return token;
}

function slugNomeArquivo(txt) {
  return String(txt || "arquivo")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[<>:"/\\|?*\u0000-\u001F]/g, "")
    .replace(/\s+/g, "-")
    .replace(/-+/g, "-")
    .replace(/\.+$/g, "")
    .trim() || "arquivo";
}

function removerExtensao(nome) {
  return String(nome || "arquivo").replace(/\.[^.]+$/, "");
}

function escaparHtml(valor) {
  return String(valor ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function nomeArquivoDaParte(item, extensao = "csv") {
  const base = slugNomeArquivo(removerExtensao(registroArquivoAtual?.nome || "arquivo"));
  const label = slugNomeArquivo(item?.labelVisivel || item?.labelParte || item?.chaveParte || "parte");
  return `${base}-${label}.${extensao}`;
}

function baixarBlobComoArquivo(blob, nomeArquivo) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = nomeArquivo;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

function atualizarCacheBotoesAcao() {
  actionButtonsCache = Array.from(
    document.querySelectorAll(
      "button, select, input[type='button'], input[type='submit']"
    )
  ).filter(Boolean);
}

function setGlobalBusy(flag, reason = "") {
  globalBusy = !!flag;
  globalBusyReason = reason || "";

  atualizarCacheBotoesAcao();

  for (const el of actionButtonsCache) {
    if (!el) continue;

    if (el === botaoFecharVisualizar) continue;

    if (globalBusy) {
      if (!el.dataset.prevDisabledAdm) {
        el.dataset.prevDisabledAdm = el.disabled ? "1" : "0";
      }
      el.disabled = true;
    } else {
      const prev = el.dataset.prevDisabledAdm;
      if (prev === "0") el.disabled = false;
      if (prev === "1") el.disabled = true;
      delete el.dataset.prevDisabledAdm;
    }
  }

  if (!globalBusy) {
    atualizarEstadoBotaoLancar();
  } else if (textoDicaLancamento && reason) {
    textoDicaLancamento.textContent = reason;
  }
}

async function executarAcaoGlobal(nomeAcao, fn) {
  if (globalBusy) return;
  setGlobalBusy(true, nomeAcao);

  try {
    return await fn();
  } finally {
    setGlobalBusy(false);
  }
}

function montarChaveAutoLancar() {
  const op = obterOpcoesPesquisaPreferenciais();
  const key = obterArquivoKeyAtual() || String(idArquivoAtual || "");
  return `${SESSION_AUTO_LANCAR_PREFIX}${key}:${op.numeroPesquisa || ""}:${op.dataPesquisa || ""}`;
}

function autoLancarJaExecutado() {
  try {
    const chave = montarChaveAutoLancar();
    return sessionStorage.getItem(chave) === "1";
  } catch {
    return false;
  }
}

function marcarAutoLancarExecutado() {
  try {
    const chave = montarChaveAutoLancar();
    sessionStorage.setItem(chave, "1");
  } catch { }
}

function limparAutoLancarDaUrl() {
  try {
    const url = new URL(window.location.href);
    url.searchParams.delete("autoLancar");
    window.history.replaceState({}, "", url.toString());
  } catch { }
}

/* =========================
   CSV -> MATRIZ
========================= */
function csvParaMatriz(csv) {
  if (!window.XLSX) return [];

  try {
    const wb = window.XLSX.read(csv || "", {
      type: "string",
      raw: true,
      codepage: 65001,
      cellText: false,
      cellDates: false
    });

    const ws = wb.Sheets[wb.SheetNames[0]];
    return window.XLSX.utils.sheet_to_json(ws, {
      header: 1,
      defval: "",
      raw: true,
      blankrows: false
    });
  } catch (e) {
    console.warn("Falha ao converter CSV para matriz:", e);
    return [];
  }
}

function csvTextoParaJson(csv) {
  if (!window.XLSX) return [];

  try {
    const wb = window.XLSX.read(csv || "", {
      type: "string",
      raw: true,
      codepage: 65001,
      cellText: false,
      cellDates: false
    });

    const ws = wb.Sheets[wb.SheetNames[0]];
    return window.XLSX.utils.sheet_to_json(ws, {
      defval: "",
      raw: true
    }) || [];
  } catch (e) {
    console.warn("Falha ao converter CSV para JSON:", e);
    return [];
  }
}

function excelSerialParaDate(valor) {
  const numero = Number(valor);
  if (!Number.isFinite(numero)) return null;

  const dias = Math.floor(numero);
  const fracao = numero - dias;

  const utcDias = (dias - 25569) * 86400 * 1000;
  const utcFracao = Math.round(fracao * 86400 * 1000);

  const data = new Date(utcDias + utcFracao);
  if (Number.isNaN(data.getTime())) return null;

  return data;
}

function formatarDataHoraManualBr(data, incluirHora = false) {
  if (!(data instanceof Date) || Number.isNaN(data.getTime())) return "";

  const dia = String(data.getDate()).padStart(2, "0");
  const mes = String(data.getMonth() + 1).padStart(2, "0");
  const ano = data.getFullYear();

  if (!incluirHora) {
    return `${dia}/${mes}/${ano}`;
  }

  const hora = String(data.getHours()).padStart(2, "0");
  const minuto = String(data.getMinutes()).padStart(2, "0");
  const segundo = String(data.getSeconds()).padStart(2, "0");

  return `${dia}/${mes}/${ano} ${hora}:${minuto}:${segundo}`;
}

function normalizarNomeColunaControle(chave) {
  return String(chave || "")
    .replace(/^\uFEFF/, "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toUpperCase()
    .replace(/[.\-\/\\]/g, "_")
    .replace(/\s+/g, "_")
    .replace(/_+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function colunaPareceData(coluna) {
  const n = normalizarNomeColunaControle(coluna);

  if (
    n === "DIA_PESQ" ||
    n === "DIA_PESQUISA" ||
    n === "DATA" ||
    n === "DATA_PESQ" ||
    n === "DATA_PESQUISA"
  ) {
    return "data";
  }

  if (
    n === "DT_ALTERACAO" ||
    n === "DATA_HORA" ||
    n === "DATAHORA" ||
    n === "ULTIMA_ALTERACAO"
  ) {
    return "datahora";
  }

  return null;
}

function formatarValorParaDownload(coluna, valor) {
  if (valor === null || valor === undefined) return "";

  if (valor instanceof Date && !Number.isNaN(valor.getTime())) {
    const tipo = colunaPareceData(coluna);
    return formatarDataHoraManualBr(valor, tipo === "datahora");
  }

  const textoOriginal = String(valor).trim();
  if (!textoOriginal) return "";

  const tipoData = colunaPareceData(coluna);

  // Se já vier em formato de data BR, preserva exatamente
  if (tipoData && /^\d{2}\/\d{2}\/\d{4}$/.test(textoOriginal)) {
    return textoOriginal;
  }

  // Se já vier em formato data/hora BR, preserva
  if (tipoData && /^\d{2}\/\d{2}\/\d{4}\s+\d{2}:\d{2}(:\d{2})?$/.test(textoOriginal)) {
    return textoOriginal;
  }

  // Só tenta converter serial do Excel quando for realmente número puro
  if (tipoData && /^\d+([.,]\d+)?$/.test(textoOriginal)) {
    const num = Number(textoOriginal.replace(",", "."));
    if (Number.isFinite(num)) {
      const data = excelSerialParaDate(num);
      if (data) {
        return formatarDataHoraManualBr(data, tipoData === "datahora");
      }
    }
  }

  return textoOriginal;
}

function colunaPareceTelefone(coluna) {
  const n = normalizarNomeColunaControle(coluna);

  return (
    n === "TF1" ||
    n === "TF2" ||
    n === "TF3" ||
    n === "TF4" ||
    n === "TELEFONE" ||
    n === "CELULAR" ||
    n === "WHATSAPP"
  );
}

function formatarValorParaPdf(coluna, valor) {
  const valorFormatadoBase = formatarValorParaDownload(coluna, valor);

  if (valorFormatadoBase === null || valorFormatadoBase === undefined) return "";

  const texto = String(valorFormatadoBase).trim();
  if (!texto) return "";

  if (colunaPareceTelefone(coluna)) {
    const soDigitos = texto.replace(/\D/g, "");

    if (soDigitos.length >= 10 && !soDigitos.startsWith("0")) {
      return `0${soDigitos}`;
    }

    return soDigitos || texto;
  }

  return texto;
}

function escCsvComDelimitador(v, delimitador = ";") {
  const s = String(v ?? "");
  const regex = new RegExp(`[\"\\n\\r${delimitador}]`);
  if (regex.test(s)) return `"${s.replace(/"/g, '""')}"`;
  return s;
}

/* =========================
   DOWNLOADS SEM COLUNAS OPERACIONAIS
   + CSV em colunas corretas
   + Datas formatadas
   + PDF com linhas sem sobreposição
========================= */
function formatarDateParaBr(date, incluirHora = false) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";

  const dd = String(date.getDate()).padStart(2, "0");
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const yyyy = String(date.getFullYear());

  if (!incluirHora) {
    return `${dd}/${mm}/${yyyy}`;
  }

  const hh = String(date.getHours()).padStart(2, "0");
  const mi = String(date.getMinutes()).padStart(2, "0");
  const ss = String(date.getSeconds()).padStart(2, "0");

  return `${dd}/${mm}/${yyyy} ${hh}:${mi}:${ss}`;
}

function nomeDeColunaPareceData(nomeColuna) {
  const n = normalizarTexto(nomeColuna)
    .replace(/\./g, " ")
    .replace(/\//g, " ");

  return (
    n.includes("data") ||
    n.includes("dia pesq") ||
    n.includes("dt alteracao") ||
    n.includes("alteracao")
  );
}

function valorPareceSerialExcel(valor) {
  if (valor === null || valor === undefined || valor === "") return false;
  if (typeof valor === "number") return valor > 20000 && valor < 80000;

  const texto = String(valor).trim().replace(",", ".");
  if (!/^\d+(\.\d+)?$/.test(texto)) return false;

  const n = Number(texto);
  return Number.isFinite(n) && n > 20000 && n < 80000;
}

function removerColunasOperacionaisDaMatriz(matriz = []) {
  if (!Array.isArray(matriz) || !matriz.length) return [];

  const cabecalhoOriginal = Array.isArray(matriz[0]) ? matriz[0] : [];
  const indicesMantidos = [];
  const novoCabecalho = [];

  cabecalhoOriginal.forEach((coluna, idx) => {
    if (!familiaColunaOperacional(coluna)) {
      indicesMantidos.push(idx);
      novoCabecalho.push(coluna);
    }
  });

  const linhas = matriz.slice(1).map((linha) =>
    indicesMantidos.map((idx, pos) => {
      const valor = Array.isArray(linha) ? linha[idx] ?? "" : "";
      const header = novoCabecalho[pos] ?? "";
      return formatarValorParaPdf(header, valor);
    })
  );

  return [novoCabecalho, ...linhas];
}

function matrizParaCsv(matriz = [], delimitador = ";") {
  if (!Array.isArray(matriz) || !matriz.length) return "";

  return matriz
    .map((linha) =>
      (linha || [])
        .map((celula) => escCsvComDelimitador(celula ?? "", delimitador))
        .join(delimitador)
    )
    .join("\r\n");
}

function obterMatrizDownloadComOperacionais(csvOriginal = "") {
  const matrizOriginal = csvParaMatriz(csvOriginal);
  if (!matrizOriginal.length) return [];

  const cabecalho = Array.isArray(matrizOriginal[0]) ? matrizOriginal[0] : [];
  const linhas = matrizOriginal.slice(1).map((linha) =>
    cabecalho.map((coluna, idx) => {
      const valor = Array.isArray(linha) ? linha[idx] ?? "" : "";
      return formatarValorParaDownload(coluna, valor);
    })
  );

  return [cabecalho, ...linhas];
}

function gerarCsvDownloadComOperacionais(csvOriginal = "") {
  const matriz = obterMatrizDownloadComOperacionais(csvOriginal);
  if (!matriz.length) return "";
  return matrizParaCsv(matriz, ";");
}

function gerarCsvDownloadSemOperacionais(csvOriginal = "") {
  const matrizOriginal = csvParaMatriz(csvOriginal);
  if (!matrizOriginal.length) return "";

  const matrizLimpa = removerColunasOperacionaisDaMatriz(matrizOriginal);
  return matrizParaCsv(matrizLimpa, ";");
}

function obterMatrizDownloadSemOperacionais(csvOriginal = "") {
  const matrizOriginal = csvParaMatriz(csvOriginal);
  if (!matrizOriginal.length) return [];
  return removerColunasOperacionaisDaMatriz(matrizOriginal);
}

function baixarCsvItem(item) {
  const nome = nomeArquivoDaParte(item, "csv");
  const csvComOperacionais = gerarCsvDownloadComOperacionais(item?.csv || "");
  const conteudoFinal = "\uFEFF" + csvComOperacionais;
  const blob = new Blob([conteudoFinal], { type: "text/csv;charset=utf-8;" });
  baixarBlobComoArquivo(blob, nome);
}

function abrirJanelaImpressaoPdf(item) {
  const matriz = obterMatrizDownloadSemOperacionais(item?.csv || "");
  if (!matriz.length) {
    throw new Error("Não há dados para gerar o PDF.");
  }

  const nome = nomeArquivoDaParte(item, "pdf");
  const headers = matriz[0] || [];
  const linhas = matriz.slice(1);

  const thead = headers.map((c) => `<th>${escaparHtml(c)}</th>`).join("");
  const tbody = linhas
    .map((linha) => {
      const tds = headers.map((_, idx) => `<td>${escaparHtml(linha[idx] ?? "")}</td>`).join("");
      return `<tr>${tds}</tr>`;
    })
    .join("");

  const html = `
    <!DOCTYPE html>
    <html lang="pt-BR">
    <head>
      <meta charset="UTF-8" />
      <title>${escaparHtml(nome)}</title>
      <style>
        @page {
          size: portrait;
          margin: 10mm;
        }
        body{
          font-family: Arial, Helvetica, sans-serif;
          margin: 0;
          color:#111;
        }
        h1{
          font-size:16px;
          margin:0 0 8px 0;
        }
        .meta{
          font-size:11px;
          margin-bottom:10px;
          line-height:1.4;
        }
        table{
          width:100%;
          border-collapse:collapse;
          table-layout:fixed;
          font-size:10px;
        }
        th,td{
          border:1px solid #999;
          padding:5px 6px;
          text-align:left;
          vertical-align:top;
          word-break:break-word;
          white-space:normal;
          line-height:1.35;
        }
        th{
          background:#eee;
        }
      </style>
    </head>
    <body>
      <h1>${escaparHtml(nome)}</h1>
      <div class="meta">
        Parte: <strong>${escaparHtml(item?.labelVisivel || item?.labelParte || item?.chaveParte || "—")}</strong><br>
        Gerado em: <strong>${escaparHtml(obterDataHoraBr())}</strong>
      </div>
      <table>
        <thead><tr>${thead}</tr></thead>
        <tbody>${tbody}</tbody>
      </table>
    </body>
    </html>
  `;

  const win = window.open("", "_blank");
  if (!win) {
    throw new Error("O navegador bloqueou a janela de impressão/PDF.");
  }

  win.document.open();
  win.document.write(html);
  win.document.close();

  setTimeout(() => {
    win.focus();
    win.print();
  }, 300);
}

async function baixarPdfItem(item) {
  const nome = nomeArquivoDaParte(item, "pdf");
  const matriz = obterMatrizDownloadSemOperacionais(item?.csv || "");

  if (!matriz.length) {
    throw new Error("Não há dados para gerar o PDF.");
  }

  if (window.jspdf?.jsPDF) {
    try {
      const { jsPDF } = window.jspdf;
      const doc = new jsPDF({
        orientation: "portrait",
        unit: "pt",
        format: "a4"
      });

      const headers = matriz[0] || [];
      const rows = matriz.slice(1);

      const pageWidth = doc.internal.pageSize.getWidth();
      const pageHeight = doc.internal.pageSize.getHeight();

      const margemX = 24;
      const margemY = 24;
      const larguraUtil = pageWidth - (margemX * 2);

      let y = margemY;

      doc.setFont("helvetica", "bold");
      doc.setFontSize(12);
      doc.text(nome, margemX, y);

      y += 16;
      doc.setFont("helvetica", "normal");
      doc.setFontSize(9);
      doc.text(`Parte: ${item?.labelVisivel || item?.labelParte || item?.chaveParte || "—"}`, margemX, y);

      y += 12;
      doc.text(`Gerado em: ${obterDataHoraBr()}`, margemX, y);
      y += 18;

      const colCount = Math.max(headers.length, 1);
      const colWidth = larguraUtil / colCount;
      const paddingX = 3;
      const paddingY = 4;
      const minLineHeight = 10;

      function quebrarTexto(valor, largura) {
        const texto = String(valor ?? "");
        return doc.splitTextToSize(texto, Math.max(largura - paddingX * 2, 8));
      }

      function alturaDaLinha(celulasQuebradas) {
        let maxLinhas = 1;
        for (const linhas of celulasQuebradas) {
          maxLinhas = Math.max(maxLinhas, Array.isArray(linhas) ? linhas.length : 1);
        }
        return (maxLinhas * minLineHeight) + (paddingY * 2);
      }

      function desenharLinha(celulasQuebradas, altura, isHeader = false) {
        let x = margemX;

        if (isHeader) {
          doc.setFont("helvetica", "bold");
        } else {
          doc.setFont("helvetica", "normal");
        }

        for (let i = 0; i < celulasQuebradas.length; i++) {
          const linhasTexto = celulasQuebradas[i];
          doc.rect(x, y, colWidth, altura);

          const linhasArr = Array.isArray(linhasTexto) ? linhasTexto : [String(linhasTexto ?? "")];
          for (let j = 0; j < linhasArr.length; j++) {
            const textoLinha = String(linhasArr[j] ?? "");
            const textoY = y + paddingY + 8 + (j * minLineHeight);
            doc.text(textoLinha, x + paddingX, textoY);
          }

          x += colWidth;
        }

        y += altura;
      }

      function desenharCabecalho() {
        const headerQuebrado = headers.map((h) => quebrarTexto(h, colWidth));
        const altura = alturaDaLinha(headerQuebrado);
        desenharLinha(headerQuebrado, altura, true);
      }

      desenharCabecalho();

      for (const row of rows) {
        const celulasQuebradas = headers.map((_, idx) => quebrarTexto(row[idx] ?? "", colWidth));
        const altura = alturaDaLinha(celulasQuebradas);

        if (y + altura > pageHeight - 24) {
          doc.addPage();
          y = margemY;
          desenharCabecalho();
        }

        desenharLinha(celulasQuebradas, altura, false);
      }

      doc.save(nome);
      return;
    } catch (e) {
      console.warn("Falha ao gerar PDF com jsPDF. Usando impressão.", e);
    }
  }

  abrirJanelaImpressaoPdf(item);
}

function baixarCsvDireto(nomeArquivo, conteudo) {
  const blob = new Blob(["\uFEFF" + conteudo], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);

  const link = document.createElement("a");
  link.href = url;
  link.download = nomeArquivo;
  link.style.display = "none";

  document.body.appendChild(link);
  link.click();

  setTimeout(() => {
    URL.revokeObjectURL(url);
    link.remove();
  }, 1000);
}

async function baixarTodasAsPartesGeraisCsv() {
  if (!partesGeradas.length) {
    throw new Error("Nenhuma parte encontrada para download.");
  }

  for (let i = 0; i < partesGeradas.length; i++) {
    const parte = partesGeradas[i];
    const nome = nomeArquivoDaParte(parte, "csv");
    const csvComOperacionais = gerarCsvDownloadComOperacionais(parte.csv || "");
    baixarCsvDireto(nome, csvComOperacionais);
    await new Promise((resolve) => setTimeout(resolve, 800));
  }
}

async function baixarTodasAsPartesGeraisPdf() {
  if (!partesGeradas.length) {
    throw new Error("Nenhuma parte encontrada para download.");
  }

  if (!window.jspdf?.jsPDF) {
    throw new Error("Para download geral em PDF, adicione o jsPDF no HTML.");
  }

  for (let i = 0; i < partesGeradas.length; i++) {
    const parte = partesGeradas[i];
    await baixarPdfItem(parte);
    await new Promise((resolve) => setTimeout(resolve, 800));
  }
}

/* =========================
   arquivoKey (estável)
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
    return await fetch(url, { ...options, signal: controller.signal });
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
   EXCLUIR TELEGRAM (ESPECÍFICO)
========================= */
async function excluirMensagensTelegramEspecificas({
  contexto = "",
  ownerId = "",
  usuarioId = "",
  telegramId = "",
  chaveParte = "",
  categoria = "",
  arquivoKey = "",
  idArquivoOrigem = null,
  nomeArquivoOrigem = "",
  modo = "ESPECIFICO",
} = {}) {
  const token = await getAccessTokenOrThrow();

  const payload = {
    modo,
    contexto: String(contexto || "").slice(0, 200),

    ownerId: String(ownerId || obterOwnerIdAtual() || ""),
    usuarioId: String(usuarioId || "").trim(),
    telegramId: String(telegramId || "").trim(),

    chaveParte: String(chaveParte || "").trim(),
    categoria: String(categoria || "").trim(),

    arquivoKey: String(arquivoKey || obterArquivoKeyAtual() || "").trim(),
    idArquivoOrigem: Number.isFinite(Number(idArquivoOrigem)) ? Number(idArquivoOrigem) : null,
    nomeArquivoOrigem: String(nomeArquivoOrigem || registroArquivoAtual?.nome || "").trim(),

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
        "x-owner-id": payload.ownerId,
        "x-usuario-id": payload.usuarioId || "",
        "x-telegram-id": payload.telegramId || "",
        "x-chave-parte": payload.chaveParte || "",
        "x-categoria": payload.categoria || "",
        "x-arquivo-key": payload.arquivoKey || "",
        "x-id-arquivo-origem": payload.idArquivoOrigem ? String(payload.idArquivoOrigem) : "",
        "x-nome-arquivo-origem": payload.nomeArquivoOrigem || "",
        "x-modo-exclusao": payload.modo || "ESPECIFICO",
      },
      body: JSON.stringify(payload),
    },
    FETCH_TIMEOUT_MS
  );

  if (!resp.ok) {
    const txt = await resp.text().catch(() => "");
    throw new Error(`Falha ao excluir mensagens do Telegram (HTTP ${resp.status}): ${txt || "sem resposta"}`);
  }

  return await resp.json().catch(() => ({}));
}

async function excluirMensagensDoArquivoInteiro() {
  const ownerId = obterOwnerIdAtual();
  const arquivoKey = obterArquivoKeyAtual();
  const idArquivoOrigem = idArquivoAtual ?? registroArquivoAtual?.id ?? null;
  const nomeArquivoOrigem = registroArquivoAtual?.nome || "";

  if (!ownerId) {
    throw new Error("OwnerId não encontrado.");
  }

  if (!arquivoKey && !Number.isFinite(Number(idArquivoOrigem))) {
    throw new Error("Não encontrei arquivoKey nem idArquivoOrigem para excluir.");
  }

  return await excluirMensagensTelegramEspecificas({
    modo: "ARQUIVO_INTEIRO",
    contexto: "EXCLUIR MENSAGENS DO ARQUIVO INTEIRO",
    ownerId,
    arquivoKey,
    idArquivoOrigem,
    nomeArquivoOrigem,
  });
}

/* =========================
   TELEGRAM
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
    u?.telegram_user_id ??
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
   PARTE vs SOBRA
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

  if (tituloVisualizar) {
    tituloVisualizar.textContent = `Visualizar ${item.labelVisivel || item.labelParte || item.chaveParte}`;
  }

  if (infoVisualizar) {
    infoVisualizar.innerHTML =
      `${origem === "sobra" ? "Sobra" : "Parte"} <strong>${item.labelVisivel || item.labelParte || item.chaveParte}</strong> — ` +
      `<strong>${item.totalLinhas}</strong> linha(s) — <strong>${item.tamanhoTexto}</strong>`;
  }

  renderizarTabelaPreviewModal(item.csv);
  if (overlayVisualizar) overlayVisualizar.hidden = false;

  const geoAtivo = listaGeoAtiva(registroArquivoAtual?.colunasVisiveis);
  if (infoVisualizar) {
    infoVisualizar.innerHTML += `<br/>Colunas geo enviadas: <strong>${geoAtivo}</strong>`;
  }
}

function fecharModalVisualizar() {
  chaveParteAbertaNoModal = null;
  if (overlayVisualizar) overlayVisualizar.hidden = true;
  if (tabelaModal) tabelaModal.innerHTML = "";
}

function renderizarTabelaPreviewModal(csv) {
  if (!tabelaModal) return;

  tabelaModal.innerHTML = "";

  const matriz = csvParaMatriz(csv);

  if (!matriz.length) {
    tabelaModal.innerHTML = "<tr><td>Nenhum dado para visualizar.</td></tr>";
    return;
  }

  const cabecalho = matriz[0] || [];
  const dados = matriz.slice(1, 1 + maxLinhasPreview);

  const thead = document.createElement("thead");
  const trHead = document.createElement("tr");

  for (const col of cabecalho) {
    const th = document.createElement("th");
    th.textContent = String(col ?? "");
    trHead.appendChild(th);
  }

  thead.appendChild(trHead);

  const tbody = document.createElement("tbody");

  for (const linha of dados) {
    const tr = document.createElement("tr");

    for (let i = 0; i < cabecalho.length; i++) {
      const td = document.createElement("td");
      td.textContent = String(linha[i] ?? "");
      tr.appendChild(td);
    }

    tbody.appendChild(tr);
  }

  tabelaModal.appendChild(thead);
  tabelaModal.appendChild(tbody);
}

botaoFecharVisualizar?.addEventListener("click", fecharModalVisualizar);

overlayVisualizar?.addEventListener("click", (e) => {
  if (e.target === overlayVisualizar) fecharModalVisualizar();
});

window.addEventListener("keydown", (e) => {
  if (overlayVisualizar && !overlayVisualizar.hidden && e.key === "Escape") {
    fecharModalVisualizar();
  }
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
    if (globalBusy) return;

    await executarAcaoGlobal(`Processando ${texto}…`, async () => {
      try {
        const item = getItemPorChave({ tipo, chave });
        if (!item) throw new Error("Item não encontrado para lançar.");

        const usuarioId =
          tipo === "sobra"
            ? (selecaoUsuarioPorSobra.get(chave) || "")
            : (selecaoUsuarioPorParte.get(chave) || "");

        if (!usuarioId) throw new Error("Selecione um usuário antes de lançar.");

        const usuario = getUsuarioPorId(usuarioId);
        if (!usuario) throw new Error("Usuário selecionado não encontrado.");

        const { telegramId } = obterTelegramDoUsuario(usuario);
        if (!telegramId) throw new Error("Usuário sem telegramId.");

        if (textoDicaLancamento) textoDicaLancamento.textContent = `Excluindo mensagens de ${texto}…`;

        await excluirMensagensTelegramEspecificas({
          modo: "ESPECIFICO",
          contexto: `INDIVIDUAL ${texto}`,
          ownerId: obterOwnerIdAtual(),
          usuarioId: String(usuario.id || ""),
          telegramId: String(telegramId || ""),
          chaveParte: String(item.chaveParte || ""),
          categoria: tipo === "sobra" ? "SOBRA" : "PARTE",
          arquivoKey: obterArquivoKeyAtual(),
          idArquivoOrigem: idArquivoAtual ?? registroArquivoAtual?.id ?? null,
          nomeArquivoOrigem: registroArquivoAtual?.nome || "",
        });

        if (textoDicaLancamento) textoDicaLancamento.textContent = `Enviando ${texto}…`;

        const resp = await enviarParteIndividualParaN8n({ tipo, chave });
        const lancamentoId = resp?.lancamentoId || resp?.id || "";

        await modalInfo(
          "Envio concluído",
          `✅ Enviado com sucesso!\n\nItem: ${texto}\n${lancamentoId ? `LancamentoId: ${lancamentoId}\n` : ""}Resposta do n8n recebida.`
        );
      } catch (err) {
        console.error(err);
        await modalInfo("Falha no envio", `❌ Falha ao enviar ${texto}\n\n${err?.message || err}`);
      }
    });
  });

  return btn;
}

/* =========================
   BOTÕES DE AÇÕES
========================= */
function criarBotaoAcaoPequeno({ texto, titulo, onClick }) {
  const btn = document.createElement("button");
  btn.type = "button";
  btn.className = "botaoSecundario botaoAcaoPequeno";
  btn.textContent = texto;
  btn.title = titulo || "";
  btn.addEventListener("click", onClick);
  return btn;
}

function criarContainerAcoesItem(item, origem = "parte") {
  const div = document.createElement("div");
  div.className = "containerAcoes";

  const btnVisualizar = criarBotaoAcaoPequeno({
    texto: "👁 Visualizar",
    titulo: "Visualizar",
    onClick: () => {
      if (globalBusy) return;
      abrirModalVisualizar(item.chaveParte, origem);
    },
  });

  const btnPdf = criarBotaoAcaoPequeno({
    texto: "📄 PDF",
    titulo: "Baixar PDF",
    onClick: async () => {
      if (globalBusy) return;
      await executarAcaoGlobal("Gerando PDF…", async () => {
        try {
          await baixarPdfItem(item);
        } catch (e) {
          console.error(e);
          await modalInfo("Falha ao gerar PDF", `❌ Falha ao gerar PDF.\n\n${e?.message || e}`);
        }
      });
    },
  });

  const btnCsv = criarBotaoAcaoPequeno({
    texto: "⬇ CSV",
    titulo: "Baixar CSV",
    onClick: async () => {
      if (globalBusy) return;
      await executarAcaoGlobal("Preparando CSV…", async () => {
        try {
          baixarCsvItem(item);
        } catch (e) {
          console.error(e);
          await modalInfo("Falha no download CSV", `❌ Falha ao baixar CSV.\n\n${e?.message || e}`);
        }
      });
    },
  });

  div.appendChild(btnVisualizar);
  div.appendChild(btnPdf);
  div.appendChild(btnCsv);

  return div;
}

/* =========================
   SELECT DE USUÁRIOS
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

  select.addEventListener("change", () => {
    if (globalBusy) return;
    onChange(select.value ? String(select.value) : "");
  });

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

    const tdAcoes = document.createElement("td");
    tdAcoes.appendChild(criarContainerAcoesItem(parte, "parte"));

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
    tr.appendChild(tdAcoes);
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

    const tdAcoes = document.createElement("td");
    tdAcoes.appendChild(criarContainerAcoesItem(sobra, "sobra"));

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
    tr.appendChild(tdAcoes);
    tr.appendChild(tdUsuario);
    tr.appendChild(tdLancar);

    corpoTabelaSobras.appendChild(tr);
  }
}

function atualizarEstadoBotaoLancar() {
  if (!botaoLancarPartes) return;

  if (globalBusy) {
    botaoLancarPartes.disabled = true;
    return;
  }

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
   COLUNAS GEO
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
   NORMALIZAÇÃO DAS COLUNAS OPERACIONAIS
========================= */
function familiaColunaOperacional(chave) {
  const n = normalizarNomeColunaControle(chave).replace(/_\d+$/, "");

  if (n === "STATUS") return "STATUS";
  if (n === "OBSERVACAO") return "OBSERVACAO";

  const compacto = n.replace(/_/g, "");
  if (compacto === "DTALTERACAO") return "DT.ALTERACAO";
  if (compacto === "ULTIMAALTERACAO") return "DT.ALTERACAO";

  return null;
}

function valorPreenchido(v) {
  return String(v ?? "").trim() !== "";
}

function statusCanonico(v) {
  const s = String(v ?? "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toUpperCase();

  if (!s) return "";
  if (s === "SEM_RESPOSTA") return "RETORNO";
  if (s === "LIGAR_MAIS_TARDE") return "RETORNO";
  if (s.startsWith("RETORNO")) return "RETORNO";

  return s;
}

function escCsv(v) {
  const s = String(v ?? "");
  if (/[",\n\r]/.test(s)) return `"${s.replace(/"/g, '""')}"`;
  return s;
}

function obterCabecalhoBaseSemOperacionais(linhas) {
  const ordem = [];
  const vistos = new Set();

  for (const linha of linhas || []) {
    for (const chave of Object.keys(linha || {})) {
      const fam = familiaColunaOperacional(chave);
      if (fam) continue;

      const nome = String(chave);
      if (!vistos.has(nome)) {
        vistos.add(nome);
        ordem.push(nome);
      }
    }
  }

  return ordem;
}

function consolidarLinhaOperacional(linha) {
  const out = {};
  const pool = {
    STATUS: [],
    OBSERVACAO: [],
    "DT.ALTERACAO": [],
  };

  for (const [k, v] of Object.entries(linha || {})) {
    const fam = familiaColunaOperacional(k);

    if (fam) {
      pool[fam].push(v ?? "");
      continue;
    }

    out[k] = v ?? "";
  }

  let statusValor = "";
  for (const v of pool.STATUS) {
    const s = statusCanonico(v);
    if (valorPreenchido(s)) statusValor = s;
  }

  let obsValor = "";
  for (const v of pool.OBSERVACAO) {
    if (valorPreenchido(v)) obsValor = String(v ?? "");
  }

  let dtValor = "";
  for (const v of pool["DT.ALTERACAO"]) {
    if (valorPreenchido(v)) dtValor = String(v ?? "");
  }

  out["STATUS"] = statusValor || "PENDENTE";
  out["OBSERVACAO"] = obsValor || "";
  out["DT.ALTERACAO"] = dtValor || "";

  return out;
}

function consolidarColunasOperacionais(linhas) {
  return (linhas || []).map((linha) => consolidarLinhaOperacional(linha));
}

function normalizarLinhaParaCsvFinal(linha, cabecalhoFinal) {
  const out = {};

  for (const col of cabecalhoFinal) {
    const valor = linha?.[col] ?? "";
    out[col] = formatarValorParaDownload(col, valor);
  }

  return out;
}

function gerarCsvDaParte(linhas) {
  const linhasNorm = consolidarColunasOperacionais(linhas || []);
  if (!linhasNorm.length) return "";

  const cabecalhoBase = obterCabecalhoBaseSemOperacionais(linhasNorm);
  const cabecalhoFinal = [
    ...cabecalhoBase,
    "STATUS",
    "OBSERVACAO",
    "DT.ALTERACAO",
  ];

  const linhasCsv = [];
  linhasCsv.push(cabecalhoFinal.map(escCsv).join(","));

  for (const linha of linhasNorm) {
    const linhaFormatada = normalizarLinhaParaCsvFinal(linha, cabecalhoFinal);
    const row = cabecalhoFinal.map((col) => escCsv(linhaFormatada?.[col] ?? ""));
    linhasCsv.push(row.join(","));
  }

  return linhasCsv.join("\n");
}

/* =========================
   EXCEL -> PARTES / SOBRAS
========================= */
function separarPorPartesESobras(linhasJson, colunasVisiveis) {
  if (!Array.isArray(linhasJson) || linhasJson.length === 0) return { principais: [], sobras: [] };

  const linhasConsolidadasOperacionais = consolidarColunasOperacionais(linhasJson);
  const geoKeys = detectarChavesGeo(linhasConsolidadasOperacionais[0]);
  const chaveColuna = encontrarChaveColunaPrimeiraLinha(linhasConsolidadasOperacionais[0]);

  if (!chaveColuna) {
    throw new Error(`Não encontrei a coluna "${nomeColunaParte}". Verifique o nome no Excel/consolidado.`);
  }

  const mapaPrincipais = new Map();
  const mapaSobras = new Map();

  const metaPrincipais = new Map();
  const metaSobras = new Map();

  for (const linha of linhasConsolidadasOperacionais) {
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
   CARREGAR USUÁRIOS
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
  return await window.ibeDb.listarUsuariosDb({ adminId: admin.id }) || [];
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
   SUPABASE - HELPERS GENÉRICOS
========================= */
function isBlobLike(valor) {
  return (
    valor instanceof Blob ||
    (valor && typeof valor === "object" && typeof valor.arrayBuffer === "function")
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

  for (let i = 0; i < len; i++) bytes[i] = bin.charCodeAt(i);

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
  return String(primeiroValor(reg, ["arquivoKey", "arquivo_key", "key"]) || "").trim();
}

function extrairNomeArquivoDoRegistro(reg) {
  return String(primeiroValor(reg, ["nome", "name", "nome_arquivo", "filename", "file_name"]) || "arquivo").trim();
}

function extrairStatusArquivoDoRegistro(reg) {
  return String(primeiroValor(reg, ["status", "situacao"]) || "—").trim();
}

function extrairTamanhoBytesDoRegistro(reg) {
  const bruto = primeiroValor(reg, ["tamanhoBytes", "tamanho_bytes", "size", "file_size", "bytes"]);
  const n = Number(bruto);
  return Number.isFinite(n) ? n : null;
}

function extrairDataArquivoDoRegistro(reg) {
  return primeiroValor(reg, ["data", "created_at", "criado_em", "uploaded_at", "updated_at"]) || "";
}

function extrairColunasVisiveisDoRegistro(reg) {
  return primeiroValor(reg, ["colunasVisiveis", "colunas_visiveis"]) || {};
}

function extrairOpcoesPesquisaDoRegistro(reg) {
  return primeiroValor(reg, ["opcoesPesquisa", "opcoes_pesquisa"]) || {};
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

/* =========================
   SUPABASE - BUSCA GENÉRICA
========================= */
async function buscarRegistroPorCampoNaTabela({ tabela, campo, valor, ownerId }) {
  try {
    let query = supabase.from(tabela).select("*").limit(1);

    if (ownerId) query = query.eq("owner_id", ownerId);
    query = query.eq(campo, valor);

    const { data, error } = await query.maybeSingle();

    if (error) {
      const msg = String(error?.message || "").toLowerCase();
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
      const reg = await buscarRegistroPorCampoNaTabela({
        tabela: SUPABASE_ARQUIVOS_TABLE,
        campo,
        valor: arquivoKey,
        ownerId
      });
      if (reg) return reg;
    }
  }

  if (Number.isFinite(idNumerico)) {
    for (const campo of candidatosId) {
      const reg = await buscarRegistroPorCampoNaTabela({
        tabela: SUPABASE_ARQUIVOS_TABLE,
        campo,
        valor: idNumerico,
        ownerId
      });
      if (reg) return reg;
    }
  }

  return null;
}

async function buscarRegistroEntregasConsolidado({ arquivoKey, idNumerico, ownerId }) {
  const candidatosKey = [
    "arquivo_key",
    "arquivoKey",
    "key",
    "arquivo_origem_key",
    "arquivoOrigemKey"
  ];

  const candidatosId = [
    "id",
    "id_arquivo_origem",
    "idArquivoOrigem",
    "arquivo_id",
    "arquivoId"
  ];

  if (arquivoKey) {
    for (const campo of candidatosKey) {
      const reg = await buscarRegistroPorCampoNaTabela({
        tabela: SUPABASE_ENTREGAS_CONSOLIDADO_TABLE,
        campo,
        valor: arquivoKey,
        ownerId
      });
      if (reg) return reg;
    }
  }

  if (Number.isFinite(idNumerico)) {
    for (const campo of candidatosId) {
      const reg = await buscarRegistroPorCampoNaTabela({
        tabela: SUPABASE_ENTREGAS_CONSOLIDADO_TABLE,
        campo,
        valor: idNumerico,
        ownerId
      });
      if (reg) return reg;
    }
  }

  return null;
}

async function baixarBlobDoStorage(path) {
  const { data, error } = await supabase.storage.from(SUPABASE_STORAGE_BUCKET).download(path);
  if (error) throw error;
  if (!data) throw new Error("Storage retornou vazio ao baixar o arquivo.");
  return data;
}

async function baixarBlobPorUrl(url, mimeType = "application/octet-stream") {
  const resp = await fetchComTimeout(url, { method: "GET" }, FETCH_TIMEOUT_MS);
  if (!resp.ok) throw new Error(`Falha ao baixar arquivo pela URL (HTTP ${resp.status}).`);

  const blob = await resp.blob();
  if (blob && blob.size > 0) return blob;

  const ab = await resp.arrayBuffer();
  return new Blob([ab], { type: mimeType });
}

function extrairLinhasJsonDoRegistroConsolidado(registroRaw) {
  const candidatos = [
    "linhasJson",
    "linhas_json",
    "json",
    "json_data",
    "data_json",
    "conteudo_json",
    "conteudoJson",
    "payload_json",
    "payloadJson",
    "registros",
    "rows",
    "itens",
    "dados"
  ];

  for (const campo of candidatos) {
    const valor = registroRaw?.[campo];

    if (Array.isArray(valor)) {
      return valor;
    }

    if (typeof valor === "string" && valor.trim()) {
      try {
        const parsed = JSON.parse(valor);
        if (Array.isArray(parsed)) return parsed;
        if (Array.isArray(parsed?.rows)) return parsed.rows;
        if (Array.isArray(parsed?.dados)) return parsed.dados;
        if (Array.isArray(parsed?.itens)) return parsed.itens;
      } catch {
        // segue tentando outros formatos
      }
    }

    if (valor && typeof valor === "object") {
      if (Array.isArray(valor.rows)) return valor.rows;
      if (Array.isArray(valor.dados)) return valor.dados;
      if (Array.isArray(valor.itens)) return valor.itens;
    }
  }

  return null;
}

function extrairCsvDoRegistroConsolidado(registroRaw) {
  const candidatos = [
    "csv",
    "conteudo_csv",
    "conteudoCsv",
    "csv_texto",
    "csvTexto",
    "texto_csv",
    "textoCsv"
  ];

  for (const campo of candidatos) {
    const valor = registroRaw?.[campo];
    if (typeof valor === "string" && valor.trim()) {
      return valor;
    }
  }

  return "";
}

async function extrairBlobDoRegistro(registroRaw) {
  let blob = null;

  const storagePath = extrairStoragePathDoRegistro(registroRaw);
  if (storagePath) {
    try {
      blob = await baixarBlobDoStorage(storagePath);
      if (blob) return blob;
    } catch (e) {
      console.warn("Falha ao baixar blob por storagePath:", e);
    }
  }

  const fileUrl = primeiroValor(registroRaw, ["url", "file_url", "arquivo_url", "public_url"]);
  if (fileUrl) {
    try {
      blob = await baixarBlobPorUrl(String(fileUrl), extrairMimeTypeDoRegistro(registroRaw));
      if (blob) return blob;
    } catch (e) {
      console.warn("Falha ao baixar blob por URL:", e);
    }
  }

  const b64 = primeiroValor(registroRaw, ["base64", "arquivoBase64", "conteudoBase64", "blobBase64"]);
  if (typeof b64 === "string" && b64.trim()) {
    try {
      blob = base64ParaBlob(b64, extrairMimeTypeDoRegistro(registroRaw));
      if (blob) return blob;
    } catch (e) {
      console.warn("Falha ao converter base64 em blob:", e);
    }
  }

  const possivelBlob = primeiroValor(registroRaw, ["blob", "arquivoBlob", "file"]);
  if (isBlobLike(possivelBlob)) return possivelBlob;

  return null;
}

/* =========================
   FONTE DO ARQUIVO PARA SEPARAÇÃO
========================= */
async function carregarLinhasDoEntregasConsolidado({ arquivoKey, idNumerico, ownerId }) {
  const registroRaw = await buscarRegistroEntregasConsolidado({
    arquivoKey,
    idNumerico,
    ownerId
  });

  if (!registroRaw) return null;

  const linhasJsonDireto = extrairLinhasJsonDoRegistroConsolidado(registroRaw);
  if (Array.isArray(linhasJsonDireto) && linhasJsonDireto.length) {
    return {
      registro: normalizarRegistroArquivoSupabase(registroRaw),
      linhasJson: linhasJsonDireto,
      fonte: "entregas_consolidado"
    };
  }

  const csv = extrairCsvDoRegistroConsolidado(registroRaw);
  if (csv) {
    const linhasJsonCsv = csvTextoParaJson(csv);
    if (Array.isArray(linhasJsonCsv) && linhasJsonCsv.length) {
      return {
        registro: normalizarRegistroArquivoSupabase(registroRaw),
        linhasJson: linhasJsonCsv,
        fonte: "entregas_consolidado"
      };
    }
  }

  const blob = await extrairBlobDoRegistro(registroRaw);
  if (blob) {
    const linhasJsonBlob = await lerExcelParaJson(blob);
    if (Array.isArray(linhasJsonBlob) && linhasJsonBlob.length) {
      return {
        registro: normalizarRegistroArquivoSupabase(registroRaw),
        linhasJson: linhasJsonBlob,
        fonte: "entregas_consolidado"
      };
    }
  }

  return null;
}

async function carregarArquivoBaseDoSupabase({ arquivoKey, idNumerico }) {
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
  const blob = await extrairBlobDoRegistro(registroRaw);

  if (!blob) {
    console.error("Registro encontrado, mas sem conteúdo utilizável:", registroRaw);
    throw new Error(
      `O registro foi encontrado no Supabase, mas não localizei o conteúdo do Excel. Verifique se a tabela "${SUPABASE_ARQUIVOS_TABLE}" possui o path do arquivo no bucket "${SUPABASE_STORAGE_BUCKET}".`
    );
  }

  const linhasJson = await lerExcelParaJson(blob);

  return {
    registro,
    linhasJson,
    fonte: "arquivo_base"
  };
}

async function carregarMetadadosArquivoBase({ arquivoKey, idNumerico }) {
  const ownerId = user?.id || null;

  const registroRaw = await buscarRegistroArquivoNoSupabase({
    arquivoKey,
    idNumerico,
    ownerId,
  });

  if (!registroRaw) return null;

  return normalizarRegistroArquivoSupabase(registroRaw);
}

async function carregarFonteArquivoParaSeparacao({ arquivoKey, idNumerico }) {
  const ownerId = user?.id || null;

  try {
    const consolidado = await carregarLinhasDoEntregasConsolidado({
      arquivoKey,
      idNumerico,
      ownerId
    });

    if (consolidado?.linhasJson?.length) {
      console.info("Separação carregada de:", consolidado.fonte);
      return consolidado;
    }
  } catch (e) {
    console.warn("Falha ao carregar de entregas_consolidado. Indo para arquivo base:", e);
  }

  const base = await carregarArquivoBaseDoSupabase({
    arquivoKey,
    idNumerico
  });

  console.info("Separação carregada de:", base.fonte);
  return base;
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
    if (!tentativa || !Object.keys(tentativa).length) continue;

    const { error } = await supabase
      .from(SUPABASE_ARQUIVOS_TABLE)
      .update(tentativa)
      .eq("id", Number(idRegistro));

    if (!error) return true;

    const msg = String(error?.message || "").toLowerCase();
    if (msg.includes("column") || msg.includes("does not exist") || msg.includes("schema cache")) {
      continue;
    }

    console.warn("Falha ao atualizar registro do arquivo no Supabase:", error);
    return false;
  }

  return false;
}

/* =========================
   ARQUIVO
========================= */
async function carregarArquivoPorKey(arquivoKey) {
  return await carregarFonteArquivoParaSeparacao({
    arquivoKey: String(arquivoKey || "").trim(),
    idNumerico: idArquivoAtual
  });
}

async function lerExcelParaJson(blobExcel) {
  if (!window.XLSX) {
    throw new Error("SheetJS (XLSX) não carregou. Verifique a tag script do CDN.");
  }

  const arrayBuffer = await blobExcel.arrayBuffer();
  const workbook = window.XLSX.read(arrayBuffer, {
    type: "array",
    cellDates: false,
    cellText: false
  });

  const nomePrimeiraAba = workbook.SheetNames[0];
  const ws = workbook.Sheets[nomePrimeiraAba];
  if (!ws) throw new Error("Não consegui ler a primeira aba do Excel.");

  return window.XLSX.utils.sheet_to_json(ws, {
    defval: "",
    raw: false,
    dateNF: "dd/mm/yyyy"
  });
}

/* =========================
   AUTO MAP
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
   ENVIO
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
   LINK DO QUESTIONÁRIO
========================= */
const QUESTIONARIO_BASE_URL = "https://www.ibespebr.com.br/questionario/pesquisa/";

function normalizarNumeroPesquisa(v) {
  return String(v || "").trim().replace(/[^\d]/g, "");
}

function normalizarDataPesquisa(v) {
  return String(v || "").trim().replace(/\//g, "-");
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
    if (
      opAtual.valida &&
      opAtual.numeroPesquisa === opUrl.numeroPesquisa &&
      opAtual.dataPesquisa === opUrl.dataPesquisa
    ) return;

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

    if (!ok) console.warn("Não foi possível persistir opcoesPesquisa no Supabase.");

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
   ENVIO INDIVIDUAL
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
    ownerId: obterOwnerIdAtual(),
    usuarioId: String(u.id || ""),

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
    "x-owner-id": obterOwnerIdAtual(),
    "x-usuario-id": String(u.id || ""),
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
   ENVIO EM LOTE
========================= */
async function lancarPartes() {
  const erro = validarDistribuicao();
  if (erro) {
    await modalInfo("Distribuição incompleta", erro);
    return;
  }

  if (!window.ibeDb || typeof window.ibeDb.salvarDistribuicaoDb !== "function") {
    await modalInfo("Erro", "db.js não tem salvarDistribuicaoDb(itens).");
    return;
  }

  const opPref = obterOpcoesPesquisaPreferenciais();
  if (partesGeradas.length > 0 && !opPref.valida) {
    await modalInfo(
      "Opções da pesquisa ausentes",
      "Faltam opções do link. Volte ao Dashboard e preencha ⚙️ Opções (Número da pesquisa e Data DD-MM-AAAA) antes de lançar."
    );
    return;
  }

  if (textoDicaLancamento) textoDicaLancamento.textContent = "Excluindo mensagens do arquivo…";

  await excluirMensagensTelegramEspecificas({
    modo: "ARQUIVO_INTEIRO",
    contexto: "LOTE (Lançar partes geral)",
    ownerId: obterOwnerIdAtual(),
    arquivoKey: obterArquivoKeyAtual(),
    idArquivoOrigem: idArquivoAtual ?? registroArquivoAtual?.id ?? null,
    nomeArquivoOrigem: registroArquivoAtual?.nome || "",
  });

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
    ownerId: obterOwnerIdAtual(),
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

  const resp = await enviarLoteParaN8n(payloadLote, {
    "x-owner-id": obterOwnerIdAtual(),
    "x-arquivo-key": obterArquivoKeyAtual(),
    "x-id-arquivo-origem": Number.isFinite(Number(idOrigemFinal)) ? String(idOrigemFinal) : "",
    "x-nome-arquivo-origem": String(registroArquivoAtual?.nome || ""),
  });
  const lancamentoId = resp?.lancamentoId || resp?.id || null;

  if (textoDicaLancamento) textoDicaLancamento.textContent = "Lote enviado com sucesso.";
  atualizarEstadoBotaoLancar();

  await modalInfo(
    "Lote enviado",
    `✅ Lote enviado ao n8n!\n\nPartes: ${itensPrincipaisSalvar.length}\nSobras enviadas: ${itensSobrasSalvar.length} (opcional)\n${lancamentoId ? `LancamentoId: ${lancamentoId}\n` : ""}`
  );
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
    if (botaoExcluirMensagensTelegram) botaoExcluirMensagensTelegram.disabled = false;

    await carregarUsuarios();

    const metaBase = await carregarMetadadosArquivoBase({
      arquivoKey: arquivoKeyAtual,
      idNumerico: idArquivoAtual,
    });

    const { registro, linhasJson, fonte } = await carregarArquivoPorKey(arquivoKeyAtual);

    registroArquivoAtual = {
      ...registro,
      ...(metaBase || {}),
      opcoesPesquisa: metaBase?.opcoesPesquisa || registro?.opcoesPesquisa || {},
      colunasVisiveis: {
        estado: true,
        cidade: true,
        regiao: true,
        ...(metaBase?.colunasVisiveis || registro?.colunasVisiveis || {}),
      },
      fonteSeparacao: fonte || "",
    };

    await persistirOpcoesPesquisaSeVieramDaUrl();

    const { principais, sobras } = separarPorPartesESobras(
      linhasJson,
      registroArquivoAtual.colunasVisiveis
    );

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

    if (registroArquivoAtual?.fonteSeparacao) {
      console.info(`Fonte usada para separação: ${registroArquivoAtual.fonteSeparacao}`);
    }

    atualizarEstadoBotaoLancar();
    limparAutoLancarDaUrl();

    if (obterFlagAutoLancar()) {
      const podeLancarAgora = botaoLancarPartes && botaoLancarPartes.disabled === false;
      const jaExecutado = autoLancarJaExecutado();

      if (!jaExecutado && podeLancarAgora) {
        marcarAutoLancarExecutado();

        await executarAcaoGlobal("Executando auto-lançamento…", async () => {
          try {
            await lancarPartes();
          } catch (e) {
            console.error(e);
            await modalInfo("Falha no auto-lançamento", `❌ Falha ao lançar automaticamente.\n\n${e?.message || e}`);
          }
        });
      } else if (!jaExecutado && !podeLancarAgora) {
        await modalInfo(
          "Auto-lançar não executado",
          "Faltam seleções obrigatórias nas PARTES principais (P01, P02...).\nSelecione e clique em 'Lançar partes (geral)'.\n\n(Sobras são opcionais.)"
        );
      }
    }
  } catch (e) {
    console.error(e);
    mostrarAlerta(`Erro: ${String(e?.message || e || "")}`);
    if (textoDicaLancamento) textoDicaLancamento.textContent = "Erro ao carregar.";
  }
}

/* =========================
   EVENTOS
========================= */
botaoRecarregar?.addEventListener("click", async () => {
  if (globalBusy) return;
  await executarAcaoGlobal("Recarregando…", async () => {
    await inicializar();
  });
});

botaoDownloadGeralCsv?.addEventListener("click", async () => {
  if (globalBusy) return;

  await executarAcaoGlobal("Baixando CSV geral…", async () => {
    try {
      await baixarTodasAsPartesGeraisCsv();
      await modalInfo("Download concluído", "✅ Download geral CSV iniciado com sucesso.");
    } catch (e) {
      console.error(e);
      await modalInfo("Falha no download CSV", `❌ Falha no download geral CSV.\n\n${e?.message || e}`);
    }
  });
});

botaoDownloadGeralPdf?.addEventListener("click", async () => {
  if (globalBusy) return;

  await executarAcaoGlobal("Baixando PDF geral…", async () => {
    try {
      await baixarTodasAsPartesGeraisPdf();
      await modalInfo("Download concluído", "✅ Download geral PDF iniciado com sucesso.");
    } catch (e) {
      console.error(e);
      await modalInfo("Falha no download PDF", `❌ Falha no download geral PDF.\n\n${e?.message || e}`);
    }
  });
});

botaoExcluirMensagensTelegram?.addEventListener("click", async () => {
  if (globalBusy) return;

  const confirmar = await modalConfirm(
    "Excluir mensagens do Telegram",
    "Deseja excluir as mensagens do Telegram relacionadas a este arquivo?\n\nIsso deve apagar somente as mensagens deste arquivo.",
    { textoConfirmar: "Excluir", perigo: true }
  );

  if (!confirmar) return;

  await executarAcaoGlobal("Excluindo mensagens do Telegram…", async () => {
    try {
      const resp = await excluirMensagensDoArquivoInteiro();

      await modalInfo(
        "Exclusão concluída",
        `✅ Exclusão solicitada com sucesso.\n\nArquivo: ${registroArquivoAtual?.nome || "—"}\n${resp?.apagadas !== undefined ? `Mensagens apagadas: ${resp.apagadas}\n` : ""}`
      );
    } catch (e) {
      console.error(e);
      await modalInfo("Falha na exclusão", `❌ Falha ao excluir mensagens do Telegram.\n\n${e?.message || e}`);
    }
  });
});

botaoLancarPartes?.addEventListener("click", async () => {
  if (globalBusy) return;

  await executarAcaoGlobal("Lançando partes…", async () => {
    try {
      await lancarPartes();
    } catch (e) {
      console.error(e);
      await modalInfo("Falha no lote", `❌ Falha ao lançar/enviar lote.\n\n${e?.message || e}`);
      if (textoDicaLancamento) textoDicaLancamento.textContent = "Falha ao enviar.";
      atualizarEstadoBotaoLancar();
    }
  });
});

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