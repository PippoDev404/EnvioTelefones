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

function baixarCsvItem(item) {
  const nome = nomeArquivoDaParte(item, "csv");
  const blob = new Blob([item?.csv || ""], { type: "text/csv;charset=utf-8" });
  baixarBlobComoArquivo(blob, nome);
}

/* =========================
   CSV -> MATRIZ
========================= */
function csvParaMatriz(csv) {
  if (!window.XLSX) return [];

  try {
    const wb = window.XLSX.read(csv || "", {
      type: "string",
      raw: false,
      codepage: 65001
    });

    const ws = wb.Sheets[wb.SheetNames[0]];
    return window.XLSX.utils.sheet_to_json(ws, {
      header: 1,
      defval: "",
      raw: false,
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
      raw: false,
      codepage: 65001
    });
    const ws = wb.Sheets[wb.SheetNames[0]];
    return window.XLSX.utils.sheet_to_json(ws, { defval: "" }) || [];
  } catch (e) {
    console.warn("Falha ao converter CSV para JSON:", e);
    return [];
  }
}

function abrirJanelaImpressaoPdf(item) {
  const matriz = csvParaMatriz(item?.csv || "");
  if (!matriz.length) {
    alert("Não há dados para gerar o PDF.");
    return;
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
        body{
          font-family: Arial, Helvetica, sans-serif;
          margin: 16px;
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
          table-layout:auto;
          font-size:9px;
        }
        th,td{
          border:1px solid #999;
          padding:4px 5px;
          text-align:left;
          vertical-align:top;
          word-break:break-word;
        }
        thead{
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
    alert("O navegador bloqueou a janela de impressão/PDF.");
    return;
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
  const matriz = csvParaMatriz(item?.csv || "");

  if (!matriz.length) {
    alert("Não há dados para gerar o PDF.");
    return;
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
      y += 18;

      const colCount = Math.max(headers.length, 1);
      const colWidth = larguraUtil / colCount;
      const lineHeight = 12;

      const desenharCabecalho = () => {
        doc.setFont("helvetica", "bold");
        headers.forEach((header, idx) => {
          const x = margemX + (idx * colWidth);
          doc.rect(x, y - 9, colWidth, lineHeight);
          doc.text(String(header ?? "").slice(0, 20), x + 2, y);
        });
        y += lineHeight;
        doc.setFont("helvetica", "normal");
      };

      desenharCabecalho();

      for (const row of rows) {
        if (y > pageHeight - 30) {
          doc.addPage();
          y = margemY;
          desenharCabecalho();
        }

        headers.forEach((_, idx) => {
          const x = margemX + (idx * colWidth);
          const valor = String(row[idx] ?? "").slice(0, 20);
          doc.rect(x, y - 9, colWidth, lineHeight);
          doc.text(valor, x + 2, y);
        });

        y += lineHeight;
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
  const blob = new Blob([conteudo], { type: "text/csv;charset=utf-8;" });
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
    alert("Nenhuma parte encontrada para download.");
    return;
  }

  for (let i = 0; i < partesGeradas.length; i++) {
    const parte = partesGeradas[i];
    const nome = nomeArquivoDaParte(parte, "csv");
    baixarCsvDireto(nome, parte.csv || "");
    await new Promise((resolve) => setTimeout(resolve, 800));
  }
}

async function baixarTodasAsPartesGeraisPdf() {
  if (!partesGeradas.length) {
    alert("Nenhuma parte encontrada para download.");
    return;
  }

  if (!window.jspdf?.jsPDF) {
    alert("Para download geral em PDF, adicione o jsPDF no HTML. Sem ele, o navegador pode bloquear várias janelas.");
    return;
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

    try {
      btn.disabled = true;
      const oldText = btn.textContent;

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

      btn.textContent = "Excluindo mensagens…";

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
    onClick: () => abrirModalVisualizar(item.chaveParte, origem),
  });

  const btnPdf = criarBotaoAcaoPequeno({
    texto: "📄 PDF",
    titulo: "Baixar PDF",
    onClick: async () => {
      try {
        await baixarPdfItem(item);
      } catch (e) {
        console.error(e);
        alert(`Falha ao gerar PDF.\n\n${e?.message || e}`);
      }
    },
  });

  const btnCsv = criarBotaoAcaoPequeno({
    texto: "⬇ CSV",
    titulo: "Baixar CSV",
    onClick: () => baixarCsvItem(item),
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
    const row = cabecalhoFinal.map((col) => escCsv(linha?.[col] ?? ""));
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
  if (!window.XLSX) throw new Error("SheetJS (XLSX) não carregou. Verifique a tag script do CDN.");

  const arrayBuffer = await blobExcel.arrayBuffer();
  const workbook = window.XLSX.read(arrayBuffer, { type: "array" });

  const nomePrimeiraAba = workbook.SheetNames[0];
  const ws = workbook.Sheets[nomePrimeiraAba];
  if (!ws) throw new Error("Não consegui ler a primeira aba do Excel.");

  return window.XLSX.utils.sheet_to_json(ws, { defval: "" });
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
  if (textoDicaLancamento) textoDicaLancamento.textContent = "Excluindo mensagens do arquivo…";

  try {
    await excluirMensagensTelegramEspecificas({
      modo: "ARQUIVO_INTEIRO",
      contexto: "LOTE (Lançar partes geral)",
      ownerId: obterOwnerIdAtual(),
      arquivoKey: obterArquivoKeyAtual(),
      idArquivoOrigem: idArquivoAtual ?? registroArquivoAtual?.id ?? null,
      nomeArquivoOrigem: registroArquivoAtual?.nome || "",
    });
  } catch (e) {
    console.error(e);
    botaoLancarPartes.disabled = false;
    if (textoDicaLancamento) textoDicaLancamento.textContent = "Falha ao excluir mensagens.";
    alert(`❌ Falha ao excluir as mensagens do arquivo antes do lote.\n\n${e?.message || e}`);
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
    mostrarAlerta(`Erro: ${String(e?.message || e || "")}`);
    if (textoDicaLancamento) textoDicaLancamento.textContent = "Erro ao carregar.";
  }
}

/* =========================
   EVENTOS
========================= */
botaoRecarregar?.addEventListener("click", async () => await inicializar());

botaoDownloadGeralCsv?.addEventListener("click", async () => {
  try {
    await baixarTodasAsPartesGeraisCsv();
  } catch (e) {
    console.error(e);
    alert(`Falha no download geral CSV.\n\n${e?.message || e}`);
  }
});

botaoDownloadGeralPdf?.addEventListener("click", async () => {
  try {
    await baixarTodasAsPartesGeraisPdf();
  } catch (e) {
    console.error(e);
    alert(`Falha no download geral PDF.\n\n${e?.message || e}`);
  }
});

botaoExcluirMensagensTelegram?.addEventListener("click", async () => {
  try {
    const confirmar = window.confirm(
      "Deseja excluir as mensagens do Telegram relacionadas a este arquivo?\n\nIsso deve apagar somente as mensagens deste arquivo."
    );

    if (!confirmar) return;

    botaoExcluirMensagensTelegram.disabled = true;
    const textoOriginal = botaoExcluirMensagensTelegram.textContent;
    botaoExcluirMensagensTelegram.textContent = "Excluindo...";

    const resp = await excluirMensagensDoArquivoInteiro();

    alert(
      `✅ Exclusão solicitada com sucesso.\n` +
      `Arquivo: ${registroArquivoAtual?.nome || "—"}\n` +
      (resp?.apagadas !== undefined ? `Mensagens apagadas: ${resp.apagadas}\n` : "")
    );

    botaoExcluirMensagensTelegram.textContent = textoOriginal;
    botaoExcluirMensagensTelegram.disabled = false;
  } catch (e) {
    console.error(e);
    alert(`❌ Falha ao excluir mensagens do Telegram.\n\n${e?.message || e}`);
    if (botaoExcluirMensagensTelegram) {
      botaoExcluirMensagensTelegram.textContent = "🗑️ Excluir mensagens do Telegram";
      botaoExcluirMensagensTelegram.disabled = false;
    }
  }
});

botaoLancarPartes?.addEventListener("click", async () => await lancarPartes());

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