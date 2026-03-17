import { requireAuth, logout } from "../8-src/auth.js";
import { supabase } from "../8-src/supabaseClient.js";

const TABELA = "entregas_consolidado";

const user = await requireAuth({ redirectTo: "../7-login/login.html" });
if (!user) throw new Error("Sem sessão");

document.querySelector("#btnLogout")?.addEventListener("click", () => logout());

const ownerId = String(user?.id || "").trim();

// ======================
// DOM
// ======================
const inputBusca = document.getElementById("inputBusca");
const selectOrdenacao = document.getElementById("selectOrdenacao");
const txtTotal = document.getElementById("txtTotal");

const corpoTabelaArquivos = document.getElementById("corpoTabelaArquivos");
const estadoVazio = document.getElementById("estadoVazio");

// modal ver
const modalVer = document.getElementById("modalVer");
const btnFecharVer = document.getElementById("btnFecharVer");
const modalVerTitulo = document.getElementById("modalVerTitulo");
const modalVerSub = document.getElementById("modalVerSub");
const previewHead = document.getElementById("previewHead");
const previewBody = document.getElementById("previewBody");

// modal partes
const modalPartes = document.getElementById("modalPartes");
const btnFecharPartes = document.getElementById("btnFecharPartes");
const modalPartesTitulo = document.getElementById("modalPartesTitulo");
const selectParte = document.getElementById("selectParte");
const btnVerParte = document.getElementById("btnVerParte");
const btnDownloadParte = document.getElementById("btnDownloadParte");
const nomeDownloadParte = document.getElementById("nomeDownloadParte");
const parteHead = document.getElementById("parteHead");
const parteBody = document.getElementById("parteBody");

// modal excluir
const modalExcluir = document.getElementById("modalExcluir");
const btnFecharExcluir = document.getElementById("btnFecharExcluir");
const btnCancelarExcluir = document.getElementById("btnCancelarExcluir");
const btnConfirmarExcluir = document.getElementById("btnConfirmarExcluir");
const modalExcluirSub = document.getElementById("modalExcluirSub");
const txtExcluirNome = document.getElementById("txtExcluirNome");

// ======================
// STATE
// ======================
let arquivosAtuais = [];          // [{ arquivoKey, nome, atualizadoEm, csv, raw, ownerId }]
let arquivoSelecionadoKey = null; // arquivoKey
let partesDoArquivo = [];         // [{key, nome, csv}]
let isLoading = false;
let realtimeChannel = null;

// estado exclusão
let excluirPendente = null; // { arquivoKey, nome, ownerId }

init();

async function init() {
  bind();
  await carregarTudo(true);
  iniciarRealtime();
}

function bind() {
  inputBusca?.addEventListener("input", aplicarFiltros);

  selectOrdenacao?.addEventListener("change", () => {
    renderTabelaArquivos(arquivosAtuais);
    aplicarFiltros();
  });

  btnFecharVer?.addEventListener("click", () => fecharDialog(modalVer));

  btnFecharPartes?.addEventListener("click", () => {
    limparPreviewParte();
    fecharDialog(modalPartes);
  });

  btnVerParte?.addEventListener("click", () => {
    const key = selectParte?.value;
    if (!arquivoSelecionadoKey || !key) return;
    verParteLocal(arquivoSelecionadoKey, key);
  });

  btnDownloadParte?.addEventListener("click", () => {
    const key = selectParte?.value;
    if (!arquivoSelecionadoKey || !key) return;

    const manual = (nomeDownloadParte?.value || "").trim();
    const forcedName = manual
      ? (manual.toLowerCase().endsWith(".xlsx") ? manual : `${manual}.xlsx`)
      : null;

    baixarParteLocal(arquivoSelecionadoKey, key, forcedName);
  });

  btnFecharExcluir?.addEventListener("click", () => fecharModalExcluir());
  btnCancelarExcluir?.addEventListener("click", () => fecharModalExcluir());
  btnConfirmarExcluir?.addEventListener("click", () => confirmarExclusao());

  window.addEventListener("keydown", (e) => {
    if (e.key !== "Escape") return;

    if (modalExcluir?.open) fecharModalExcluir();

    if (modalPartes?.open) {
      limparPreviewParte();
      fecharDialog(modalPartes);
    }

    if (modalVer?.open) fecharDialog(modalVer);
  });

  modalExcluir?.addEventListener("click", (e) => {
    const r = modalExcluir.getBoundingClientRect();
    const clicouFora =
      e.clientX < r.left ||
      e.clientX > r.right ||
      e.clientY < r.top ||
      e.clientY > r.bottom;

    if (clicouFora) fecharModalExcluir();
  });

  window.addEventListener("beforeunload", () => {
    encerrarRealtime();
  });
}

// ======================
// FETCH SUPABASE (GET)
// ======================
async function carregarTudo(mostrarPrimeiraVez = false) {
  if (isLoading) return;
  isLoading = true;

  try {
    const { data, error } = await supabase
      .from(TABELA)
      .select("arquivo_key, csv, atualizado_em, nome_arquivo_origem, owner_id")
      .eq("owner_id", ownerId)
      .order("atualizado_em", { ascending: false });

    if (error) throw error;

    const linhas = Array.isArray(data) ? data : [];

    const novaLista = linhas
      .map((r) => ({
        arquivoKey: String(r.arquivo_key ?? "").trim(),
        nome: String(r.nome_arquivo_origem ?? "—").trim(),
        atualizadoEm: r.atualizado_em ?? null,
        csv: garantirTexto(r.csv ?? ""),
        ownerId: String(r.owner_id ?? "").trim(),
        raw: r,
      }))
      .filter((x) => x.arquivoKey);

    const mudou = hashLista(arquivosAtuais) !== hashLista(novaLista);

    if (mostrarPrimeiraVez || mudou) {
      arquivosAtuais = novaLista;
      renderTabelaArquivos(arquivosAtuais);
      aplicarFiltros();

      if (txtTotal) txtTotal.textContent = String(arquivosAtuais.length);
    }
  } catch (e) {
    console.error("Erro ao carregar arquivos:", e);

    if (mostrarPrimeiraVez) {
      arquivosAtuais = [];
      renderTabelaArquivos(arquivosAtuais);
      aplicarFiltros();

      if (txtTotal) txtTotal.textContent = "0";
    }
  } finally {
    isLoading = false;
  }
}

// ======================
// REALTIME
// ======================
function iniciarRealtime() {
  encerrarRealtime();

  realtimeChannel = supabase
    .channel(`watch-${TABELA}-${ownerId}`)
    .on(
      "postgres_changes",
      {
        event: "*",
        schema: "public",
        table: TABELA,
      },
      async (payload) => {
        const row = payload?.new || payload?.old || {};
        const rowOwnerId = String(row.owner_id ?? "").trim();

        // Evita recarregar por alterações de outros usuários
        if (rowOwnerId && rowOwnerId !== ownerId) return;

        await carregarTudo(false);
      }
    )
    .subscribe((status) => {
      console.log("[Realtime]", status);
    });
}

function encerrarRealtime() {
  if (realtimeChannel) {
    supabase.removeChannel(realtimeChannel);
    realtimeChannel = null;
  }
}

// ======================
// RENDER TABELA + AÇÕES
// ======================
function renderTabelaArquivos(lista) {
  if (!corpoTabelaArquivos) return;

  const ordenacao = String(selectOrdenacao?.value || "desc");

  const ordenada = [...(lista || [])].sort((a, b) => {
    const da = new Date(a.atualizadoEm || 0).getTime();
    const db = new Date(b.atualizadoEm || 0).getTime();

    if (ordenacao === "asc") return da - db;
    if (ordenacao === "nome") {
      return String(a.nome || "").localeCompare(String(b.nome || ""), "pt-BR");
    }
    return db - da;
  });

  corpoTabelaArquivos.innerHTML = "";

  if (!ordenada.length) {
    if (estadoVazio) estadoVazio.hidden = false;
    return;
  }

  if (estadoVazio) estadoVazio.hidden = true;

  for (const a of ordenada) {
    const baseNome = (a.nome || `arquivo-${a.arquivoKey}`).trim();
    const nomeExibido = baseNome.toLowerCase().endsWith(".xlsx")
      ? baseNome
      : baseNome.toLowerCase().endsWith(".csv")
        ? baseNome.replace(/\.csv$/i, ".xlsx")
        : `${baseNome}.xlsx`;

    const updated = formatarData(a.atualizadoEm) || "—";
    const size = formatarBytes(estimarBytes(a.csv || "")) || "—";
    const partsCount = contarPartesNoCsv(a.csv || "");

    const tr = document.createElement("tr");
    tr.dataset.arquivokey = String(a.arquivoKey);

    tr.innerHTML = `
      <td><span class="badgePartes" title="Partes detectadas">${escapeHtml(String(partsCount))}</span></td>
      <td class="nomeArquivo">${escapeHtml(nomeExibido)}</td>
      <td>${escapeHtml(updated)}</td>
      <td>${escapeHtml(size)}</td>
      <td>
        <div class="acoesLinha">
          <button class="botaoAcao" type="button" data-acao="ver">👁 Ver</button>
          <button class="botaoAcao" type="button" data-acao="download">⬇️ Excel</button>
          <button class="botaoAcao" type="button" data-acao="partes">🧩 Partes</button>
          <button class="botaoAcao" type="button" data-acao="excluir">🗑 Excluir</button>
        </div>
      </td>
    `;

    corpoTabelaArquivos.appendChild(tr);
  }

  corpoTabelaArquivos.onclick = (e) => {
    const btn = e.target.closest("button");
    if (!btn) return;

    const acao = String(btn.dataset.acao || "");
    const tr = btn.closest("tr");
    const arquivoKey = tr?.dataset?.arquivokey;
    if (!arquivoKey) return;

    if (acao === "ver") verArquivoLocal(arquivoKey);
    if (acao === "download") baixarArquivoLocal(arquivoKey);
    if (acao === "partes") abrirPartesLocal(arquivoKey);
    if (acao === "excluir") abrirModalExcluir(arquivoKey);
  };
}

function aplicarFiltros() {
  const termo = String(inputBusca?.value || "").trim().toLowerCase();

  const linhas = Array.from(corpoTabelaArquivos?.querySelectorAll("tr") || []);
  let visiveis = 0;

  for (const tr of linhas) {
    const nome = tr.querySelector(".nomeArquivo")?.textContent?.toLowerCase() || "";
    const mostrar = !termo || nome.includes(termo);
    tr.style.display = mostrar ? "" : "none";
    if (mostrar) visiveis += 1;
  }

  if (estadoVazio) estadoVazio.hidden = visiveis !== 0;
}

// ======================
// EXCLUIR (SUPABASE)
// ======================
function abrirModalExcluir(arquivoKey) {
  const arq = arquivosAtuais.find((x) => String(x.arquivoKey) === String(arquivoKey));
  if (!arq) return;

  excluirPendente = {
    arquivoKey: String(arq.arquivoKey),
    nome: String(arq.nome || ""),
    ownerId: String(arq.ownerId || ownerId),
  };

  if (txtExcluirNome) {
    const nomeExib = (arq.nome || `arquivo-${arq.arquivoKey}`).trim();
    txtExcluirNome.textContent = `Arquivo: ${nomeExib} (key: ${arq.arquivoKey})`;
  }

  if (modalExcluirSub) {
    modalExcluirSub.textContent = "Tem certeza que deseja excluir este arquivo do sistema?";
  }

  abrirDialog(modalExcluir);
}

function fecharModalExcluir() {
  excluirPendente = null;
  fecharDialog(modalExcluir);
}

async function confirmarExclusao() {
  if (!excluirPendente) return;

  const payload = { ...excluirPendente };

  removerDaListaLocal(payload.arquivoKey);
  fecharModalExcluir();

  try {
    const { error } = await supabase
      .from(TABELA)
      .delete()
      .eq("arquivo_key", payload.arquivoKey)
      .eq("owner_id", ownerId);

    if (error) {
      await carregarTudo(true);
      throw error;
    }

    await carregarTudo(true);
  } catch (e) {
    console.error("Falha ao excluir:", e);
    await carregarTudo(true);
  }
}

function removerDaListaLocal(arquivoKey) {
  arquivosAtuais = arquivosAtuais.filter((x) => String(x.arquivoKey) !== String(arquivoKey));
  renderTabelaArquivos(arquivosAtuais);
  aplicarFiltros();

  if (txtTotal) {
    txtTotal.textContent = String(arquivosAtuais.length);
  }
}

function verArquivoLocal(arquivoKey) {
  limparPreviewMaster();

  const arq = arquivosAtuais.find((x) => String(x.arquivoKey) === String(arquivoKey));
  if (!arq) return;

  if (modalVerTitulo) modalVerTitulo.textContent = `Ver: ${arq.nome || arq.arquivoKey}`;
  if (modalVerSub) modalVerSub.textContent = "Prévia (20 linhas). Cada vírgula (,) é uma coluna.";

  const { headers, rows } = csvPreview(arq.csv || "", 20);
  renderTabela(previewHead, previewBody, headers, rows);

  abrirDialog(modalVer);
}

function baixarArquivoLocal(arquivoKey) {
  const arq = arquivosAtuais.find((x) => String(x.arquivoKey) === String(arquivoKey));
  if (!arq || !window.XLSX) return;

  const base = (arq.nome || `arquivo-${arq.arquivoKey}`).trim()
    .replace(/\.csv$/i, "")
    .replace(/\.xlsx$/i, "");

  salvarXlsxComoArquivoCsv(arq.csv || "", `${base}.xlsx`);
}

function abrirPartesLocal(arquivoKey) {
  arquivoSelecionadoKey = arquivoKey;
  partesDoArquivo = [];

  if (selectParte) selectParte.innerHTML = "";
  if (nomeDownloadParte) nomeDownloadParte.value = "";

  limparPreviewParte();

  const arq = arquivosAtuais.find((x) => String(x.arquivoKey) === String(arquivoKey));
  if (!arq) return;

  if (modalPartesTitulo) {
    modalPartesTitulo.textContent = `Partes: ${arq.nome || arq.arquivoKey}`;
  }

  const partes = gerarPartesDoCsv(arq.csv || "", "Nº PESQ.");

  if (!partes.length) {
    if (selectParte) {
      selectParte.innerHTML = `<option value="">(sem partes detectadas)</option>`;
    }
    abrirDialog(modalPartes);
    return;
  }

  partesDoArquivo = partes;

  if (selectParte) {
    selectParte.innerHTML = partes
      .map((p) => `<option value="${escapeHtml(String(p.key))}">${escapeHtml(String(p.nome))}</option>`)
      .join("");
  }

  abrirDialog(modalPartes);
}

function verParteLocal(_arquivoKey, parteKey) {
  limparPreviewParte();

  const parte = partesDoArquivo.find((p) => String(p.key) === String(parteKey));
  if (!parte) return;

  const { headers, rows } = csvPreview(parte.csv || "", 20);
  renderTabela(parteHead, parteBody, headers, rows);
}

function baixarParteLocal(arquivoKey, parteKey, forcedName = null) {
  const arq = arquivosAtuais.find((x) => String(x.arquivoKey) === String(arquivoKey));
  const parte = partesDoArquivo.find((p) => String(p.key) === String(parteKey));

  if (!arq || !parte || !window.XLSX) return;

  const base = (arq.nome || `arquivo-${arq.arquivoKey}`)
    .replace(/\.csv$/i, "")
    .replace(/\.xlsx$/i, "");

  let filename = forcedName || `${base}-${parte.key}.xlsx`;
  filename = filename.toLowerCase().endsWith(".xlsx") ? filename : `${filename}.xlsx`;

  salvarXlsxComoArquivoCsv(parte.csv || "", filename);
}

function garantirTexto(txt) {
  return String(txt || "").replace(/\r\n/g, "\n").replace(/\r/g, "\n");
}

function estimarBytes(txt) {
  return new Blob([String(txt || "")], { type: "text/plain;charset=utf-8" }).size;
}

function csvPreview(csvText, limitLinhas = 20) {
  const linhas = String(csvText || "").split("\n").filter((l) => l.trim().length > 0);
  if (!linhas.length) return { headers: [], rows: [] };

  const headers = parseCsvLine(linhas[0]);
  const dados = linhas.slice(1, 1 + limitLinhas).map(parseCsvLine);

  const colCount = headers.length;
  const rows = dados.map((r) => {
    const rr = [...r];
    while (rr.length < colCount) rr.push("");
    if (rr.length > colCount) rr.length = colCount;
    return rr;
  });

  return { headers, rows };
}

function parseCsvLine(line) {
  const s = String(line ?? "");
  const out = [];
  let cur = "";
  let inQuotes = false;

  for (let i = 0; i < s.length; i++) {
    const ch = s[i];

    if (ch === '"') {
      const next = s[i + 1];
      if (inQuotes && next === '"') {
        cur += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (ch === "," && !inQuotes) {
      out.push(cur);
      cur = "";
      continue;
    }

    cur += ch;
  }

  out.push(cur);
  return out.map(limparCampoCsv);
}

function limparCampoCsv(v) {
  return String(v ?? "")
    .replace(/\t+/g, " ")
    .replace(/\s{2,}/g, " ")
    .replace(/^"|"$/g, "")
    .trim();
}

function gerarPartesDoCsv(csvText, nomeColunaParte = "Nº PESQ.") {
  const linhas = String(csvText || "").split("\n").filter((l) => l.trim().length > 0);
  if (linhas.length < 2) return [];

  const headerLine = linhas[0];
  const headers = parseCsvLine(headerLine);

  const idxParte = headers.findIndex(
    (h) => normalizarTexto(h) === normalizarTexto(nomeColunaParte)
  );
  if (idxParte < 0) return [];

  const mapa = new Map();

  for (let i = 1; i < linhas.length; i++) {
    const cols = parseCsvLine(linhas[i]);
    const parteRaw = String(cols[idxParte] ?? "").trim().toUpperCase();
    if (!parteRaw) continue;

    const parteKey = normalizarParteKey(parteRaw);
    if (!parteKey) continue;

    if (!mapa.has(parteKey)) mapa.set(parteKey, []);
    mapa.get(parteKey).push(linhas[i]);
  }

  return [...mapa.entries()]
    .sort((a, b) => a[0].localeCompare(b[0]))
    .map(([key, lines]) => ({
      key,
      nome: key,
      csv: [headerLine, ...lines].join("\n"),
      totalLinhas: lines.length,
    }));
}

function normalizarParteKey(v) {
  const s = String(v || "").trim().toUpperCase().replace(/\s+/g, "");
  if (!s) return null;

  const mP = s.match(/^P0*(\d{1,3})$/);
  if (mP) return `P${String(parseInt(mP[1], 10)).padStart(2, "0")}`;

  const mN = s.match(/^0*(\d{1,3})$/);
  if (mN) return `P${String(parseInt(mN[1], 10)).padStart(2, "0")}`;

  return null;
}

function contarPartesNoCsv(csvText) {
  return gerarPartesDoCsv(csvText || "", "Nº PESQ.").length;
}

function normalizarTexto(txt) {
  return String(txt || "")
    .replace(/\t+/g, " ")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toLowerCase();
}

function salvarXlsxComoArquivoCsv(csvText, filenameXlsx) {
  if (!window.XLSX) return;

  const linhas = String(csvText || "")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .split("\n")
    .filter((l) => l.trim().length > 0);

  const aoa = linhas.map(parseCsvLine);
  const maxCols = aoa.reduce((m, r) => Math.max(m, (r || []).length), 0);

  const aoaNorm = aoa.map((r) => {
    const rr = [...(r || [])];
    while (rr.length < maxCols) rr.push("");
    return rr;
  });

  const ws = XLSX.utils.aoa_to_sheet(aoaNorm);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Dados");

  let nome = filenameXlsx || "arquivo.xlsx";
  nome = nome.toLowerCase().endsWith(".xlsx") ? nome : `${nome}.xlsx`;

  XLSX.writeFile(wb, nome);
}

function renderTabela(elHead, elBody, headers, rows) {
  if (!elHead || !elBody) return;

  elHead.innerHTML = "";
  elBody.innerHTML = "";

  const trh = document.createElement("tr");
  (headers || []).forEach((h) => {
    const th = document.createElement("th");
    th.textContent = String(h);
    trh.appendChild(th);
  });
  elHead.appendChild(trh);

  (rows || []).forEach((r) => {
    const tr = document.createElement("tr");
    (r || []).forEach((cell) => {
      const td = document.createElement("td");
      td.textContent = cell == null ? "" : String(cell);
      tr.appendChild(td);
    });
    elBody.appendChild(tr);
  });
}

function limparPreviewMaster() {
  if (previewHead) previewHead.innerHTML = "";
  if (previewBody) previewBody.innerHTML = "";
}

function limparPreviewParte() {
  if (parteHead) parteHead.innerHTML = "";
  if (parteBody) parteBody.innerHTML = "";
}

function abrirDialog(dlg) {
  if (!dlg) return;
  try {
    if (typeof dlg.showModal === "function") dlg.showModal();
    else dlg.setAttribute("open", "open");
  } catch {
    dlg.setAttribute("open", "open");
  }
}

function fecharDialog(dlg) {
  if (!dlg) return;
  try {
    if (typeof dlg.close === "function") dlg.close();
    else dlg.removeAttribute("open");
  } catch {
    dlg.removeAttribute("open");
  }
}

function escapeHtml(str) {
  return String(str)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function formatarBytes(bytes) {
  if (bytes == null || isNaN(Number(bytes))) return null;

  const b = Number(bytes);
  const units = ["B", "KB", "MB", "GB"];
  let u = 0;
  let v = b;

  while (v >= 1024 && u < units.length - 1) {
    v /= 1024;
    u++;
  }

  return `${v.toFixed(u === 0 ? 0 : 2)} ${units[u]}`;
}

function formatarData(iso) {
  if (!iso) return null;

  try {
    const d = new Date(iso);
    if (isNaN(d.getTime())) return String(iso);
    return d.toLocaleString("pt-BR");
  } catch {
    return String(iso);
  }
}

function hashLista(lista) {
  if (!Array.isArray(lista)) return "";
  return lista
    .map(
      (x) =>
        `${x.arquivoKey ?? ""}|${x.atualizadoEm ?? ""}|${x.nome ?? ""}|${(x.csv ?? "").length}|${x.ownerId ?? ""}`
    )
    .sort()
    .join("||");
}