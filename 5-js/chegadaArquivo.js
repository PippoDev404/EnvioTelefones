import { requireAuth, logout } from "../8-src/auth.js";
import { supabase } from "../8-src/supabaseClient.js";

const TABELA = "entregas_consolidado";
const TABELA_ROWS = "entregas_linhas";
const cacheCsvOperacional = new Map();

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

const modalVer = document.getElementById("modalVer");
const btnFecharVer = document.getElementById("btnFecharVer");
const btnRecarregarVer = document.getElementById("btnRecarregarVer");
const modalVerTitulo = document.getElementById("modalVerTitulo");
const modalVerSub = document.getElementById("modalVerSub");
const relatorioVer = document.getElementById("relatorioVer");

const modalPartes = document.getElementById("modalPartes");
const btnFecharPartes = document.getElementById("btnFecharPartes");
const modalPartesTitulo = document.getElementById("modalPartesTitulo");
const selectParte = document.getElementById("selectParte");
const btnVerParte = document.getElementById("btnVerParte");
const btnDownloadParte = document.getElementById("btnDownloadParte");
const nomeDownloadParte = document.getElementById("nomeDownloadParte");
const parteHead = document.getElementById("parteHead");
const parteBody = document.getElementById("parteBody");

const modalExcluir = document.getElementById("modalExcluir");
const btnFecharExcluir = document.getElementById("btnFecharExcluir");
const btnCancelarExcluir = document.getElementById("btnCancelarExcluir");
const btnConfirmarExcluir = document.getElementById("btnConfirmarExcluir");
const modalExcluirSub = document.getElementById("modalExcluirSub");
const txtExcluirNome = document.getElementById("txtExcluirNome");

// ======================
// STATE
// ======================
let arquivosAtuais = [];
let arquivoSelecionadoKey = null;
let arquivoSelecionadoVerKey = null;
let partesDoArquivo = [];
let isLoading = false;
let realtimeChannel = null;
let excluirPendente = null;

let relatorioAtualBase = null;
let pesquisadoresSelecionados = new Set();
let dropdownPesquisadoresAberto = false;
let recarregandoModalVer = false;
let salvandoFiltroPromise = null;

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

  btnFecharVer?.addEventListener("click", () => {
    relatorioAtualBase = null;
    pesquisadoresSelecionados = new Set();
    arquivoSelecionadoVerKey = null;
    dropdownPesquisadoresAberto = false;
    fecharDialog(modalVer);
  });

  btnRecarregarVer?.addEventListener("click", async () => {
    await recarregarModalVer();
  });

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

    if (modalVer?.open) {
      fecharDropdownPesquisadores();
      relatorioAtualBase = null;
      pesquisadoresSelecionados = new Set();
      arquivoSelecionadoVerKey = null;
      dropdownPesquisadoresAberto = false;
      fecharDialog(modalVer);
    }
  });

  document.addEventListener("click", (e) => {
    const dentroDoDropdown = e.target.closest?.(".selectPesquisadoresWrap");
    if (!dentroDoDropdown) {
      fecharDropdownPesquisadores();
    }
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

async function obterCsvOperacional(arquivoKey, forceReload = false) {
  const key = String(arquivoKey || "").trim();
  if (!key) return "";

  if (!forceReload && cacheCsvOperacional.has(key)) {
    return cacheCsvOperacional.get(key);
  }

  const { data, error } = await supabase
    .from(TABELA_ROWS)
    .select("linha_idx, line, idp, estado, cidade, regiao_cidade, tf1, tf2, status, observacao, dt_alteracao, pesquisador, arquivo_key")
    .eq("arquivo_key", key)
    .order("linha_idx", { ascending: true });

  if (error) throw error;

  const rows = Array.isArray(data) ? data : [];

  const headers = [
    "LINE",
    "IDP",
    "ESTADO",
    "CIDADE",
    "REGIAO_CIDADE",
    "TF1",
    "TF2",
    "STATUS",
    "OBSERVACAO",
    "DT_ALTERACAO",
    "Nº PESQ."
  ];

  function esc(v) {
    const s = String(v ?? "");
    if (/[",\n\r]/.test(s)) {
      return `"${s.replace(/"/g, '""')}"`;
    }
    return s;
  }

  const linhas = rows.map((r, idx) => {
    return [
      esc(r.line ?? idx + 1),
      esc(r.idp ?? ""),
      esc(r.estado ?? ""),
      esc(r.cidade ?? ""),
      esc(r.regiao_cidade ?? ""),
      esc(r.tf1 ?? ""),
      esc(r.tf2 ?? ""),
      esc(r.status ?? ""),
      esc(r.observacao ?? ""),
      esc(r.dt_alteracao ?? ""),
      esc(r.pesquisador ?? "SEM PARTE")
    ].join(",");
  });

  const csv = [headers.join(","), ...linhas].join("\n");
  cacheCsvOperacional.set(key, csv);
  return csv;
}

// ======================
// FETCH SUPABASE
// ======================
async function carregarTudo(mostrarPrimeiraVez = false) {
  if (isLoading) return;
  isLoading = true;

  try {
    const { data, error } = await supabase
      .from(TABELA)
      .select("arquivo_key, csv, atualizado_em, nome_arquivo_origem, owner_id, pesquisadores_desmarcados")
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
        pesquisadoresDesmarcados: Array.isArray(r.pesquisadores_desmarcados)
          ? r.pesquisadores_desmarcados.map((x) => String(x || "").trim()).filter(Boolean)
          : [],
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

async function carregarArquivoPorKey(arquivoKey) {
  const { data, error } = await supabase
    .from(TABELA)
    .select("arquivo_key, csv, atualizado_em, nome_arquivo_origem, owner_id, pesquisadores_desmarcados")
    .eq("owner_id", ownerId)
    .eq("arquivo_key", arquivoKey)
    .maybeSingle();

  if (error) throw error;
  if (!data) return null;

  return {
    arquivoKey: String(data.arquivo_key ?? "").trim(),
    nome: String(data.nome_arquivo_origem ?? "—").trim(),
    atualizadoEm: data.atualizado_em ?? null,
    csv: garantirTexto(data.csv ?? ""),
    ownerId: String(data.owner_id ?? "").trim(),
    pesquisadoresDesmarcados: Array.isArray(data.pesquisadores_desmarcados)
      ? data.pesquisadores_desmarcados.map((x) => String(x || "").trim()).filter(Boolean)
      : [],
    raw: data,
  };
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
// EXCLUIR
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

  cacheCsvOperacional.delete(payload.arquivoKey);
  removerDaListaLocal(payload.arquivoKey);
  fecharModalExcluir();

  try {
    const { error: errorRows } = await supabase
      .from(TABELA_ROWS)
      .delete()
      .eq("arquivo_key", payload.arquivoKey);

    if (errorRows) {
      await carregarTudo(true);
      throw errorRows;
    }

    const { error: errorConsolidado } = await supabase
      .from(TABELA)
      .delete()
      .eq("arquivo_key", payload.arquivoKey)
      .eq("owner_id", ownerId);

    if (errorConsolidado) {
      await carregarTudo(true);
      throw errorConsolidado;
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

// ======================
// VER ARQUIVO = RELATÓRIO
// ======================
async function verArquivoLocal(arquivoKey) {
  limparRelatorioMaster();

  const arq = arquivosAtuais.find((x) => String(x.arquivoKey) === String(arquivoKey));
  if (!arq) return;

  arquivoSelecionadoVerKey = String(arquivoKey);

  if (modalVerTitulo) {
    modalVerTitulo.textContent = `Relatório: ${arq.nome || arq.arquivoKey}`;
  }

  try {
    const csvOperacional = await obterCsvOperacional(arquivoKey, true);
    relatorioAtualBase = gerarRelatorioCsv(csvOperacional || "");
  } catch (e) {
    console.error("Erro ao carregar CSV operacional:", e);
    relatorioAtualBase = gerarRelatorioCsv(arq.csv || "");
  }

  const pesquisadoresBase = relatorioAtualBase?.tabelaPesquisadores || [];
  const todosPesquisadores = pesquisadoresBase.map((item) => item.parte);
  const desmarcados = new Set(
    (arq.pesquisadoresDesmarcados || [])
      .map((x) => String(x || "").trim())
      .filter(Boolean)
  );

  pesquisadoresSelecionados = new Set(
    todosPesquisadores.filter((parte) => !desmarcados.has(parte))
  );

  dropdownPesquisadoresAberto = false;
  atualizarRelatorioFiltrado();
  abrirDialog(modalVer);
}

async function recarregarModalVer() {
  if (!arquivoSelecionadoVerKey || recarregandoModalVer) return;

  recarregandoModalVer = true;

  if (btnRecarregarVer) {
    btnRecarregarVer.disabled = true;
    btnRecarregarVer.innerHTML = `<i class="fa-solid fa-rotate-right fa-spin"></i> Recarregando`;
  }

  try {
    if (salvandoFiltroPromise) {
      await salvandoFiltroPromise;
    }

    const arqAtualizado = await carregarArquivoPorKey(arquivoSelecionadoVerKey);
    if (!arqAtualizado) return;

    const idx = arquivosAtuais.findIndex(
      (x) => String(x.arquivoKey) === String(arquivoSelecionadoVerKey)
    );

    if (idx >= 0) {
      arquivosAtuais[idx] = arqAtualizado;
    } else {
      arquivosAtuais.unshift(arqAtualizado);
    }

    renderTabelaArquivos(arquivosAtuais);
    aplicarFiltros();

    if (txtTotal) txtTotal.textContent = String(arquivosAtuais.length);

    try {
      const csvOperacional = await obterCsvOperacional(arquivoSelecionadoVerKey, true);
      relatorioAtualBase = gerarRelatorioCsv(csvOperacional || "");
    } catch (e) {
      console.error("Erro ao recarregar CSV operacional:", e);
      relatorioAtualBase = gerarRelatorioCsv(arqAtualizado.csv || "");
    }

    const pesquisadoresBase = relatorioAtualBase?.tabelaPesquisadores || [];
    const todosPesquisadores = pesquisadoresBase.map((item) => item.parte);
    const desmarcados = new Set(
      (arqAtualizado.pesquisadoresDesmarcados || [])
        .map((x) => String(x || "").trim())
        .filter(Boolean)
    );

    pesquisadoresSelecionados = new Set(
      todosPesquisadores.filter((parte) => !desmarcados.has(parte))
    );

    atualizarRelatorioFiltrado();
  } catch (e) {
    console.error("Erro ao recarregar modal:", e);
  } finally {
    recarregandoModalVer = false;

    if (btnRecarregarVer) {
      btnRecarregarVer.disabled = false;
      btnRecarregarVer.innerHTML = `<i class="fa-solid fa-rotate-right"></i> Recarregar`;
    }
  }
}

function atualizarRelatorioFiltrado() {
  if (!relatorioAtualBase) return;

  const relatorioFiltrado = gerarRelatorioFiltrado(
    relatorioAtualBase,
    pesquisadoresSelecionados
  );

  if (modalVerSub) {
    modalVerSub.textContent =
      `Total de registros: ${formatarNumero(relatorioFiltrado.totais.totalRegistros)} • ` +
      `Coluna de status: ${relatorioAtualBase.nomeColunaStatus || "não identificada"} • ` +
      `Coluna de parte: ${relatorioAtualBase.nomeColunaParte || "não identificada"}`;
  }

  renderRelatorioArquivo(relatorioFiltrado, relatorioAtualBase, pesquisadoresSelecionados);
}

function renderRelatorioArquivo(relatorio, relatorioBase, selecionadosSet) {
  if (!relatorioVer) return;

  if (!relatorio || !relatorio.totais) {
    relatorioVer.innerHTML = `<div class="relatorioVazio">Não foi possível gerar o relatório deste arquivo.</div>`;
    return;
  }

  const totais = relatorio.totais;
  const statusGeralRows = relatorio.statusGeralRows || [];
  const tabelaPorPesquisador = relatorio.tabelaPesquisadores || [];
  const statusColumns = relatorio.statusColumns || [];
  const pesquisadoresBase = relatorioBase?.tabelaPesquisadores || [];
  const totalSelecionados = pesquisadoresBase.filter((item) => selecionadosSet.has(item.parte)).length;

  const cardsHtml = `
    <div class="relatorioCards">
      <div class="cardResumo">
        <p class="cardResumoTitulo">Total de números</p>
        <p class="cardResumoValor">${formatarNumero(totais.totalRegistros)}</p>
      </div>
      <div class="cardResumo">
        <p class="cardResumoTitulo">Tratados</p>
        <p class="cardResumoValor">${formatarNumero(totais.tratados)}</p>
      </div>
      <div class="cardResumo">
        <p class="cardResumoTitulo">Pendentes</p>
        <p class="cardResumoValor">${formatarNumero(totais.pendentes)}</p>
      </div>
      <div class="cardResumo">
        <p class="cardResumoTitulo">Quantidade de Pesquisadores</p>
        <p class="cardResumoValor">${formatarNumero(totalSelecionados)}</p>
      </div>
    </div>
  `;

  const filtroHtml = `
    <section class="filtroPesquisadores">
      <div class="filtroPesquisadoresTopo">
        <div>
          <h4 class="filtroPesquisadoresTitulo">Filtro de pesquisadores</h4>
          <p class="filtroPesquisadoresSub">
            Remova pesquisadores da soma e das tabelas usando o seletor abaixo.
          </p>
        </div>
      </div>

      <div class="selectPesquisadoresWrap">
        <button type="button" class="botaoSelectPesquisadores" id="btnAbrirSelectPesquisadores">
          <span>${montarTextoResumoSelecionados(pesquisadoresBase, selecionadosSet)}</span>
          <i class="fa-solid fa-chevron-down"></i>
        </button>

        <div class="dropdownPesquisadores" id="dropdownPesquisadores" ${dropdownPesquisadoresAberto ? "" : "hidden"}>
          <div class="dropdownPesquisadoresTopo">
            <button class="botaoMini" type="button" data-filtro-p="todos">Selecionar todos</button>
            <button class="botaoMini" type="button" data-filtro-p="nenhum">Limpar todos</button>
          </div>

          <div class="dropdownPesquisadoresLista">
            ${pesquisadoresBase.length
      ? pesquisadoresBase.map((item) => {
        const checked = selecionadosSet.has(item.parte);
        return `
                      <label class="itemPesquisadorSelect ${checked ? "" : "desmarcado"}">
                        <input
                          type="checkbox"
                          data-pesquisador="${escapeHtml(item.parte)}"
                          ${checked ? "checked" : ""}
                        />
                        <span>${escapeHtml(item.parte)}</span>
                      </label>
                    `;
      }).join("")
      : `<div class="relatorioVazio" style="width:100%;">Nenhum pesquisador identificado.</div>`
    }
          </div>
        </div>
      </div>
    </section>
  `;

  const statusGeralHtml = `
    <section class="blocoRelatorio">
      <div class="blocoRelatorioTopo">
        <div>
          <h4 class="blocoRelatorioTitulo">Resumo geral por status</h4>
          <p class="blocoRelatorioSub">Contagem consolidada dos pesquisadores atualmente selecionados.</p>
        </div>
      </div>

      <div class="tabelaRelatorioWrap">
        <table class="tabelaRelatorio" aria-label="Resumo geral por status">
          <thead>
            <tr>
              <th>Status</th>
              <th class="colTotal">Quantidade</th>
            </tr>
          </thead>
          <tbody>
            ${statusGeralRows.length
      ? statusGeralRows.map((item) => `
                  <tr>
                    <td style="text-align:left;">${escapeHtml(item.status)}</td>
                    <td class="colTotal">${formatarNumero(item.quantidade)}</td>
                  </tr>
                `).join("")
      : `
                  <tr>
                    <td colspan="2">Nenhum status encontrado.</td>
                  </tr>
                `
    }
          </tbody>
        </table>
      </div>
    </section>
  `;

  const pesquisadoresHtml = `
    <section class="blocoRelatorio">
      <div class="blocoRelatorioTopo">
        <div>
          <h4 class="blocoRelatorioTitulo">Resumo por pesquisador</h4>
          <p class="blocoRelatorioSub">Cada linha representa um pesquisador (P) e cada coluna representa um status.</p>
        </div>
      </div>

      <div class="tabelaRelatorioWrap">
        <table class="tabelaRelatorio" aria-label="Resumo por pesquisador e status">
          <thead>
            <tr>
              <th class="colPesquisador">Pesquisador</th>
              <th class="colTotal">Total</th>
              <th class="colTotal">Tratados</th>
              ${statusColumns.map((status) => `<th class="colStatus">${escapeHtml(status)}</th>`).join("")}
            </tr>
          </thead>
          <tbody>
            ${tabelaPorPesquisador.length
      ? tabelaPorPesquisador.map((linha) => `
                  <tr>
                    <td class="colPesquisador">${escapeHtml(linha.parte)}</td>
                    <td class="colTotal">${formatarNumero(linha.total)}</td>
                    <td class="colTotal">${formatarNumero(linha.total - (linha.statuses?.["Pendente"] || 0))}</td>
                    ${statusColumns.map((status) => `<td class="colStatus">${formatarNumero(linha.statuses?.[status] || 0)}</td>`).join("")}</tr>`).join("") : `
                  <tr>
                    <td colspan="${2 + statusColumns.length}">Nenhum pesquisador selecionado.</td>
                  </tr>
                `
    }
          </tbody>
        </table>
      </div>
    </section>
  `;

  relatorioVer.innerHTML = `${cardsHtml}${filtroHtml}${statusGeralHtml}${pesquisadoresHtml}`;
  bindFiltroPesquisadores();
}

function bindFiltroPesquisadores() {
  const btnAbrir = relatorioVer?.querySelector("#btnAbrirSelectPesquisadores");
  const dropdown = relatorioVer?.querySelector("#dropdownPesquisadores");

  btnAbrir?.addEventListener("click", (e) => {
    e.stopPropagation();
    if (!dropdown) return;

    dropdown.hidden = !dropdown.hidden;
    dropdownPesquisadoresAberto = !dropdown.hidden;
  });

  relatorioVer?.querySelectorAll('input[data-pesquisador]')?.forEach((input) => {
    input.addEventListener("click", (e) => e.stopPropagation());

    input.addEventListener("change", async (e) => {
      e.stopPropagation();

      const checkbox = e.currentTarget;
      const parte = String(checkbox?.dataset?.pesquisador || "").trim();
      if (!parte) return;

      if (checkbox.checked) pesquisadoresSelecionados.add(parte);
      else pesquisadoresSelecionados.delete(parte);

      dropdownPesquisadoresAberto = true;

      await salvarFiltroPesquisadoresNoBanco(arquivoSelecionadoVerKey);
      atualizarRelatorioFiltrado();
    });

    atualizarClasseVisualCheckbox(input);
  });

  relatorioVer?.querySelector('[data-filtro-p="todos"]')?.addEventListener("click", async (e) => {
    e.stopPropagation();

    pesquisadoresSelecionados = new Set(
      (relatorioAtualBase?.tabelaPesquisadores || []).map((item) => item.parte)
    );

    dropdownPesquisadoresAberto = true;

    await salvarFiltroPesquisadoresNoBanco(arquivoSelecionadoVerKey);
    atualizarRelatorioFiltrado();
  });

  relatorioVer?.querySelector('[data-filtro-p="nenhum"]')?.addEventListener("click", async (e) => {
    e.stopPropagation();

    pesquisadoresSelecionados = new Set();

    dropdownPesquisadoresAberto = true;

    await salvarFiltroPesquisadoresNoBanco(arquivoSelecionadoVerKey);
    atualizarRelatorioFiltrado();
  });
}

function atualizarClasseVisualCheckbox(input) {
  const label = input?.closest?.(".itemPesquisadorSelect");
  if (!label) return;

  if (input.checked) label.classList.remove("desmarcado");
  else label.classList.add("desmarcado");
}

function fecharDropdownPesquisadores() {
  const dropdown = document.getElementById("dropdownPesquisadores");
  if (dropdown) dropdown.hidden = true;
  dropdownPesquisadoresAberto = false;
}

function montarTextoResumoSelecionados(base, selecionadosSet) {
  const total = Array.isArray(base) ? base.length : 0;
  const marcados = Array.from(selecionadosSet || []).length;

  if (!total) return "Nenhum pesquisador";
  if (marcados === total) return `Todos os pesquisadores (${total})`;
  if (marcados === 0) return "Nenhum pesquisador selecionado";
  return `${marcados} de ${total} pesquisador(es) selecionado(s)`;
}

async function salvarFiltroPesquisadoresNoBanco(arquivoKey) {
  if (!arquivoKey || !relatorioAtualBase) return false;

  const arq = arquivosAtuais.find(
    (x) => String(x.arquivoKey || "").trim() === String(arquivoKey || "").trim()
  );

  if (!arq) {
    console.warn("Arquivo não encontrado em arquivosAtuais", { arquivoKey });
    return false;
  }

  const keyTrimmed = String(arquivoKey || "").trim();
  const ownerTrimmed = String(arq.ownerId || arq.owner_id || "").trim();

  const todos = (relatorioAtualBase.tabelaPesquisadores || [])
    .map((item) => String(item.parte || "").trim())
    .filter(Boolean);

  const desmarcados = todos.filter((parte) => !pesquisadoresSelecionados.has(parte));

  arq.pesquisadoresDesmarcados = [...desmarcados];

  if (salvandoFiltroPromise) {
    try {
      await salvandoFiltroPromise;
    } catch (_) { }
  }

  salvandoFiltroPromise = (async () => {
    const payload = {
      pesquisadores_desmarcados: desmarcados,
      ultima_alteracao_em: new Date().toISOString()
    };

    try {
      console.log("Salvando filtro de pesquisadores", {
        tabela: TABELA,
        arquivoKey: keyTrimmed,
        ownerId: ownerTrimmed,
        payload
      });

      // 1) Verifica se a linha existe com arquivo_key + owner_id
      let rowCheck = null;
      let checkError = null;

      if (ownerTrimmed) {
        const resCheckOwner = await supabase
          .from(TABELA)
          .select("arquivo_key, owner_id, pesquisadores_desmarcados")
          .eq("arquivo_key", keyTrimmed)
          .eq("owner_id", ownerTrimmed)
          .maybeSingle();

        rowCheck = resCheckOwner.data;
        checkError = resCheckOwner.error;
      }

      // 2) Se não achou com owner_id, tenta só por arquivo_key
      if (!rowCheck) {
        const resCheckKey = await supabase
          .from(TABELA)
          .select("arquivo_key, owner_id, pesquisadores_desmarcados")
          .eq("arquivo_key", keyTrimmed)
          .maybeSingle();

        if (!checkError) checkError = resCheckKey.error;
        rowCheck = resCheckKey.data;
      }

      if (checkError) {
        console.warn("Erro na verificação antes do update", checkError);
      }

      if (!rowCheck) {
        console.warn("Nenhuma linha encontrada para salvar filtro", {
          tabela: TABELA,
          arquivoKey: keyTrimmed,
          ownerId: ownerTrimmed
        });
        return false;
      }

      console.log("Linha encontrada para update", rowCheck);

      // 3) Tenta update com arquivo_key + owner_id
      let updateRes = null;

      if (ownerTrimmed) {
        updateRes = await supabase
          .from(TABELA)
          .update(payload)
          .eq("arquivo_key", keyTrimmed)
          .eq("owner_id", ownerTrimmed)
          .select("arquivo_key, owner_id, pesquisadores_desmarcados, ultima_alteracao_em");
      }

      if (updateRes?.error) {
        console.error("Erro no update com owner_id", updateRes.error);
      }

      // 4) Se não atualizou nada, tenta só com arquivo_key
      if (!updateRes || updateRes.error || !Array.isArray(updateRes.data) || updateRes.data.length === 0) {
        console.warn("Update com owner_id não alterou linha; tentando apenas arquivo_key", {
          arquivoKey: keyTrimmed,
          ownerId: ownerTrimmed
        });

        updateRes = await supabase
          .from(TABELA)
          .update(payload)
          .eq("arquivo_key", keyTrimmed)
          .select("arquivo_key, owner_id, pesquisadores_desmarcados, ultima_alteracao_em");
      }

      if (updateRes.error) {
        console.error("Erro no update por arquivo_key", updateRes.error);
        return false;
      }

      if (!Array.isArray(updateRes.data) || updateRes.data.length === 0) {
        console.warn("Update executado mas nenhuma linha retornou", {
          arquivoKey: keyTrimmed,
          ownerId: ownerTrimmed
        });

        // checagem final para confirmar se salvou mesmo assim
        const confirmRes = await supabase
          .from(TABELA)
          .select("arquivo_key, owner_id, pesquisadores_desmarcados, ultima_alteracao_em")
          .eq("arquivo_key", keyTrimmed)
          .maybeSingle();

        if (confirmRes.error) {
          console.error("Erro na confirmação final", confirmRes.error);
          return false;
        }

        const salvo = Array.isArray(confirmRes.data?.pesquisadores_desmarcados)
          && JSON.stringify(confirmRes.data.pesquisadores_desmarcados) === JSON.stringify(desmarcados);

        if (salvo) {
          console.log("Filtro salvo com sucesso (confirmado por leitura)", confirmRes.data);
          return true;
        }

        console.warn("Confirmação final não encontrou o valor salvo", confirmRes.data);
        return false;
      }

      console.log("Filtro salvo com sucesso", updateRes.data);
      return true;
    } catch (e) {
      console.error("Erro ao salvar filtro de pesquisadores:", e);
      return false;
    } finally {
      salvandoFiltroPromise = null;
    }
  })();

  return await salvandoFiltroPromise;
}

function gerarRelatorioFiltrado(relatorioBase, selecionadosSet) {
  const linhasSelecionadas = (relatorioBase?.tabelaPesquisadores || [])
    .filter((item) => selecionadosSet.has(item.parte));

  const statusColumns = [...(relatorioBase?.statusColumns || [])];
  const statusGeralMap = new Map();
  const totais = {
    totalRegistros: 0,
    tratados: 0,
    pendentes: 0,
  };

  for (const linha of linhasSelecionadas) {
    totais.totalRegistros += Number(linha.total || 0);

    for (const status of statusColumns) {
      const qtd = Number(linha.statuses?.[status] || 0);
      if (!qtd) continue;

      statusGeralMap.set(status, Number(statusGeralMap.get(status) || 0) + qtd);

      if (normalizarTexto(status) === "pendente") totais.pendentes += qtd;
      else totais.tratados += qtd;
    }
  }

  return {
    nomeColunaStatus: relatorioBase?.nomeColunaStatus || null,
    nomeColunaParte: relatorioBase?.nomeColunaParte || null,
    quantidadePesquisadores: linhasSelecionadas.length,
    totais,
    statusColumns,
    statusGeralRows: mapContagemParaRows(statusGeralMap),
    tabelaPesquisadores: linhasSelecionadas,
  };
}

function gerarRelatorioCsv(csvText) {
  const linhas = String(csvText || "").split("\n").filter((l) => l.trim().length > 0);

  if (linhas.length < 2) {
    return {
      nomeColunaStatus: null,
      nomeColunaParte: null,
      quantidadePesquisadores: 0,
      totais: {
        totalRegistros: 0,
        tratados: 0,
        pendentes: 0,
      },
      statusGeralRows: [],
      statusColumns: [],
      tabelaPesquisadores: [],
    };
  }

  const headers = parseCsvLine(linhas[0]);
  const dados = linhas.slice(1).map(parseCsvLine);

  const idxStatus = encontrarIndiceColuna(headers, [
    "status",
    "status final",
    "status ligação",
    "status ligacao",
    "situação",
    "situacao",
    "resultado",
    "resultado final",
  ]);

  const idxParte = encontrarIndiceColuna(headers, [
    "nº pesq.",
    "n° pesq.",
    "no pesq.",
    "num pesq",
    "numero pesq",
    "n pesq",
    "pesq",
  ]);

  const nomeColunaStatus = idxStatus >= 0 ? String(headers[idxStatus] || "") : null;
  const nomeColunaParte = idxParte >= 0 ? String(headers[idxParte] || "") : null;

  const totais = {
    totalRegistros: 0,
    tratados: 0,
    pendentes: 0,
  };

  const statusGeralMap = new Map();
  const partesMap = new Map();
  const statusSet = new Set();

  for (const row of dados) {
    totais.totalRegistros += 1;

    const statusRaw = idxStatus >= 0 ? limparCampoCsv(row[idxStatus] ?? "") : "";
    const parteRaw = idxParte >= 0 ? limparCampoCsv(row[idxParte] ?? "") : "";

    const parteKey = normalizarParteKey(parteRaw) || "SEM PARTE";
    const statusInfo = classificarStatus(statusRaw);

    if (!partesMap.has(parteKey)) {
      partesMap.set(parteKey, {
        parte: parteKey,
        total: 0,
        statuses: {},
      });
    }

    const parte = partesMap.get(parteKey);
    parte.total += 1;

    if (statusInfo.tipo === "pendente") {
      totais.pendentes += 1;
    } else if (statusInfo.tipo !== "sem_status") {
      totais.tratados += 1;
    }

    if (statusInfo.tipo !== "sem_status") {
      incrementarMapaContagem(statusGeralMap, statusInfo.label, 1);
      parte.statuses[statusInfo.label] = Number(itemOrZero(parte.statuses[statusInfo.label])) + 1;
      statusSet.add(statusInfo.label);
    }
  }

  const statusColumnsBrutas = [...statusSet].sort((a, b) => {
    const qa = Number(statusGeralMap.get(a) || 0);
    const qb = Number(statusGeralMap.get(b) || 0);
    if (qb !== qa) return qb - qa;
    return a.localeCompare(b, "pt-BR");
  });

  const statusColumns = statusColumnsBrutas.map(formatarNomeStatusExibicao);

  const tabelaPesquisadores = [...partesMap.values()]
    .sort((a, b) => ordenarPartes(a.parte, b.parte))
    .map((item) => {
      const statuses = {};
      for (const statusBruto of statusColumnsBrutas) {
        const nomeExibicao = formatarNomeStatusExibicao(statusBruto);
        statuses[nomeExibicao] = Number(item.statuses[statusBruto] || 0);
      }
      return {
        parte: item.parte,
        total: item.total,
        statuses,
      };
    });

  return {
    nomeColunaStatus,
    nomeColunaParte,
    quantidadePesquisadores: tabelaPesquisadores.length,
    totais,
    statusGeralRows: mapContagemParaRows(statusGeralMap).map((item) => ({
      ...item,
      status: formatarNomeStatusExibicao(item.status),
    })),
    statusColumns,
    tabelaPesquisadores,
  };
}

function itemOrZero(v) {
  return Number(v || 0);
}

function classificarStatus(statusRaw) {
  const original = String(statusRaw || "").trim();
  const norm = normalizarTexto(original);

  if (!norm) {
    return {
      tipo: "sem_status",
      label: "Sem status",
    };
  }

  if (
    norm.includes("pendente") ||
    norm === "pend" ||
    norm.includes("aguardando")
  ) {
    return {
      tipo: "pendente",
      label: original,
    };
  }

  return {
    tipo: "tratado",
    label: original,
  };
}

function formatarNomeStatusExibicao(status) {
  const bruto = String(status || "").trim();
  if (!bruto) return "Sem status";

  const normalizado = normalizarTexto(bruto).replace(/\s+/g, "_");

  const mapaFixos = {
    pendente: "Pendente",
    pesquisa_feita: "Pesquisa feita",
    retorno: "Retorno",
    nao_atendeu: "Nõ atendeu/caixa postal",
    nao_pode_fazer_pesquisa: "Não pode fazer pesquisa",
    numero_nao_existe: "N° não existe/ n° bloqueado",
    outra_cidade: "Outra cidade",
    recusa: "Recusa",
    remover_da_lista: "Remover da lista",
    sem_status: "Sem status",
  };

  if (mapaFixos[normalizado]) {
    return mapaFixos[normalizado];
  }

  return bruto
    .replace(/_/g, " ")
    .toLowerCase()
    .replace(/\b\w/g, (letra) => letra.toUpperCase());
}

function encontrarIndiceColuna(headers, candidatos = []) {
  const normHeaders = (headers || []).map((h) => normalizarTexto(h));

  for (const candidato of candidatos) {
    const normCand = normalizarTexto(candidato);
    const idxExato = normHeaders.findIndex((h) => h === normCand);
    if (idxExato >= 0) return idxExato;
  }

  for (const candidato of candidatos) {
    const normCand = normalizarTexto(candidato);
    const idxParcial = normHeaders.findIndex((h) => h.includes(normCand));
    if (idxParcial >= 0) return idxParcial;
  }

  return -1;
}

function incrementarMapaContagem(mapa, chave, valor = 1) {
  const key = String(chave || "").trim() || "Sem status";
  mapa.set(key, Number(mapa.get(key) || 0) + Number(valor || 0));
}

function mapContagemParaRows(mapa) {
  return [...(mapa?.entries() || [])]
    .map(([status, quantidade]) => ({
      status: String(status || ""),
      quantidade: Number(quantidade || 0),
    }))
    .sort((a, b) => {
      if (b.quantidade !== a.quantidade) return b.quantidade - a.quantidade;
      return a.status.localeCompare(b.status, "pt-BR");
    });
}

function ordenarPartes(a, b) {
  const pa = extrairNumeroParte(a);
  const pb = extrairNumeroParte(b);

  if (pa !== null && pb !== null) return pa - pb;
  if (pa !== null) return -1;
  if (pb !== null) return 1;

  return String(a || "").localeCompare(String(b || ""), "pt-BR");
}

function extrairNumeroParte(nome) {
  const m = String(nome || "").match(/^P(\d{1,3})$/i);
  if (!m) return null;
  return Number(m[1]);
}

// ======================
// DOWNLOAD / PARTES
// ======================
async function baixarArquivoLocal(arquivoKey) {
  const arq = arquivosAtuais.find((x) => String(x.arquivoKey) === String(arquivoKey));
  if (!arq || !window.XLSX) return;

  const base = (arq.nome || `arquivo-${arq.arquivoKey}`).trim()
    .replace(/\.csv$/i, "")
    .replace(/\.xlsx$/i, "");

  try {
    const csvOperacional = await obterCsvOperacional(arquivoKey, true);
    salvarXlsxComoArquivoCsv(csvOperacional || "", `${base}.xlsx`);
  } catch (e) {
    console.error("Erro ao baixar arquivo operacional:", e);
  }
}

async function abrirPartesLocal(arquivoKey) {
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

  try {
    const csvOperacional = await obterCsvOperacional(arquivoKey, true);
    const partes = gerarPartesDoCsv(csvOperacional || "", "Nº PESQ.");

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
  } catch (e) {
    console.error("Erro ao carregar partes do arquivo:", e);
    if (selectParte) {
      selectParte.innerHTML = `<option value="">(erro ao carregar partes)</option>`;
    }
    abrirDialog(modalPartes);
  }
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

// ======================
// HELPERS CSV
// ======================
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
    .sort((a, b) => a[0].localeCompare(b[0], "pt-BR"))
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

// ======================
// UI HELPERS
// ======================
function limparRelatorioMaster() {
  if (relatorioVer) relatorioVer.innerHTML = "";
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

function formatarNumero(valor) {
  return Number(valor || 0).toLocaleString("pt-BR");
}

function hashLista(lista) {
  if (!Array.isArray(lista)) return "";
  return lista
    .map(
      (x) =>
        `${x.arquivoKey ?? ""}|${x.atualizadoEm ?? ""}|${x.nome ?? ""}|${(x.csv ?? "").length}|${x.ownerId ?? ""}|${JSON.stringify(x.pesquisadoresDesmarcados ?? [])}`
    )
    .sort()
    .join("||");
}