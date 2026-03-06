import { requireAuth, logout } from "../8-src/auth.js";
import { supabase } from "../8-src/supabaseClient.js";

await requireAuth({ redirectTo: "../7-login/login.html" });

document.querySelector("#btnLogout")?.addEventListener("click", () => logout());

// ===== Seletores =====
const botaoSelecionar = document.querySelector("[data-botao-selecionar]");
const inputArquivo = document.querySelector("[data-entrada]");
const listaArquivos = document.querySelector("[data-lista]");
const areaSoltar = document.querySelector("[data-area]");
const formularioEnvio = document.getElementById("formularioEnvio");

// ===== Modal sucesso =====
const modalSucessoOverlay = document.getElementById("modalSucessoOverlay");
const modalSucessoSegundos = document.getElementById("modalSucessoSegundos");
const modalSucessoTexto = document.getElementById("modalSucessoTexto");
const btnFecharModalSucesso = document.getElementById("btnFecharModalSucesso");
const btnVerArquivo = document.getElementById("btnVerArquivo");
const btnIrDashboard = document.getElementById("btnIrDashboard");

let timerModalSucesso = null;
let segundosRestantes = 5;
let modalFoiFechado = false;
let ultimoArquivoKeyEnviado = null;

// ===== Utilitários =====
function formatarTamanho(bytes) {
  if (!bytes) return "0 Bytes";
  const k = 1024;
  const sizes = ["Bytes", "KB", "MB", "GB"];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return Math.round((bytes / Math.pow(k, i)) * 100) / 100 + " " + sizes[i];
}

function obterDataBr() {
  const d = new Date();
  const dia = String(d.getDate()).padStart(2, "0");
  const mes = String(d.getMonth() + 1).padStart(2, "0");
  const ano = d.getFullYear();
  return `${dia}/${mes}/${ano}`;
}

function gerarArquivoKey() {
  return `ARQ-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function pegarExt(nome) {
  const p = String(nome || "").split(".");
  const ext = p.length > 1 ? p.pop() : "";
  return String(ext || "").toLowerCase();
}

// ===== Modal =====
function abrirModalSucesso(nomeArquivo, arquivoKey) {
  if (!modalSucessoOverlay) return;

  ultimoArquivoKeyEnviado = arquivoKey;

  modalFoiFechado = false;
  segundosRestantes = 5;

  if (modalSucessoTexto) modalSucessoTexto.textContent = `"${nomeArquivo}" enviado com sucesso!`;
  if (modalSucessoSegundos) modalSucessoSegundos.textContent = segundosRestantes;

  modalSucessoOverlay.hidden = false;

  clearInterval(timerModalSucesso);
  timerModalSucesso = setInterval(() => {
    if (modalFoiFechado) {
      clearInterval(timerModalSucesso);
      return;
    }

    segundosRestantes--;
    if (modalSucessoSegundos) modalSucessoSegundos.textContent = segundosRestantes;

    if (segundosRestantes <= 0) {
      clearInterval(timerModalSucesso);
      window.location.href = `/4-html/dashboard.html`;
    }
  }, 1000);
}

function fecharModalSucesso() {
  modalFoiFechado = true;
  clearInterval(timerModalSucesso);
  if (modalSucessoOverlay) modalSucessoOverlay.hidden = true;
}

btnFecharModalSucesso?.addEventListener("click", fecharModalSucesso);

btnIrDashboard?.addEventListener("click", () => {
  window.location.href = "/4-html/dashboard.html";
});

btnVerArquivo?.addEventListener("click", () => {
  if (ultimoArquivoKeyEnviado) {
    window.location.href = `/4-html/admArquivo.html?key=${encodeURIComponent(ultimoArquivoKeyEnviado)}`;
  } else {
    window.location.href = "/4-html/dashboard.html";
  }
});

modalSucessoOverlay?.addEventListener("click", (e) => {
  if (e.target === modalSucessoOverlay) fecharModalSucesso();
});

window.addEventListener("keydown", (e) => {
  if (modalSucessoOverlay && !modalSucessoOverlay.hidden && e.key === "Escape") fecharModalSucesso();
});

// ===== Render =====
function renderizarArquivo(arquivo) {
  if (!arquivo) return;

  if (!arquivo.name.match(/\.(xlsx|xls|csv)$/i)) {
    inputArquivo.value = "";
    listaArquivos.innerHTML = "";
    listaArquivos.hidden = true;
    return;
  }

  listaArquivos.innerHTML = `
    <article class="itemArquivo">
      <div class="itemArquivoEsq">
        <div class="itemArquivoBadge">${arquivo.name.split(".").pop().toUpperCase()}</div>
        <div class="itemArquivoMeta">
          <div class="itemArquivoNome">${arquivo.name}</div>
          <div class="itemArquivoTamanho">${formatarTamanho(arquivo.size)}</div>
        </div>
      </div>
      <button type="button" class="botao" id="removerArquivo">Remover</button>
    </article>
  `;

  listaArquivos.hidden = false;

  document.getElementById("removerArquivo")?.addEventListener("click", () => {
    inputArquivo.value = "";
    listaArquivos.innerHTML = "";
    listaArquivos.hidden = true;
    areaSoltar?.classList.remove("dragover");
  });
}

// ===== Seleção manual =====
botaoSelecionar?.addEventListener("click", (e) => {
  e.preventDefault();
  inputArquivo?.click();
});

inputArquivo?.addEventListener("change", () => {
  renderizarArquivo(inputArquivo.files?.[0]);
});

// ===== Drag and drop =====
if (areaSoltar && inputArquivo) {
  areaSoltar.addEventListener("dragover", (e) => {
    e.preventDefault();
    areaSoltar.classList.add("dragover");
  });

  areaSoltar.addEventListener("dragleave", () => {
    areaSoltar.classList.remove("dragover");
  });

  areaSoltar.addEventListener("drop", (e) => {
    e.preventDefault();
    areaSoltar.classList.remove("dragover");

    const arquivo = e.dataTransfer.files?.[0];
    if (!arquivo) return;

    inputArquivo.files = e.dataTransfer.files;
    renderizarArquivo(arquivo);
  });
}

// ===== Envio (IndexedDB + Supabase metadata) =====
formularioEnvio?.addEventListener("submit", async (e) => {
  e.preventDefault();

  const arquivo = inputArquivo.files?.[0];
  if (!arquivo) return;

  if (!arquivo.name.match(/\.(xlsx|xls|csv)$/i)) return;

  if (!window.ibeDb?.salvarArquivoDb) {
    alert("db.js não carregou (window.ibeDb.salvarArquivoDb).");
    return;
  }

  // 1) gera a key UMA VEZ (mesma para IndexedDB e Supabase)
  const arquivoKey = gerarArquivoKey();
  const ext = pegarExt(arquivo.name);

  // 2) salva no IndexedDB (local)
  const registro = {
    arquivoKey,
    idNumerico: Date.now(),
    nome: arquivo.name,
    tipo: ext, // ✅ minúsculo: xlsx/xls/csv
    tamanho: formatarTamanho(arquivo.size),
    tamanhoBytes: arquivo.size,
    data: obterDataBr(),
    dataCriacao: Date.now(),
    status: "Pendente",
    blob: arquivo
  };

  try {
    await window.ibeDb.salvarArquivoDb(registro);
  } catch (err) {
    console.error(err);
    alert("Falha ao salvar no banco do navegador (IndexedDB).");
    return;
  }

  // 3) cria/atualiza no Supabase (metadata) — NÃO quebra o envio local se der erro
  try {
    const { data: authData, error: authErr } = await supabase.auth.getUser();
    if (authErr) throw authErr;

    const ownerId = authData?.user?.id;
    if (!ownerId) throw new Error("Sem ownerId no Supabase");

    const { error } = await supabase
      .from("arquivos")
      .upsert(
        {
          arquivo_key: arquivoKey,
          owner_id: ownerId,
          nome: arquivo.name,
          tamanho_bytes: arquivo.size,
          mime_type: arquivo.type || null,
          atualizado_em: new Date().toISOString(),
          data_criacao: new Date().toISOString()
        },
        { onConflict: "arquivo_key" }
      );

    if (error) throw error;
  } catch (err) {
    console.warn("Aviso: não consegui salvar metadata no Supabase (envio local ok):", err?.message || err);
  }

  // 4) limpa UI + modal
  inputArquivo.value = "";
  listaArquivos.innerHTML = "";
  listaArquivos.hidden = true;

  abrirModalSucesso(arquivo.name, arquivoKey);
});