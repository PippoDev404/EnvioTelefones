import { requireAuth, logout } from "../8-src/auth.js";
import { supabase } from "../8-src/supabaseClient.js";

await requireAuth({ redirectTo: "../7-login/login.html" });

document.querySelector("#btnLogout")?.addEventListener("click", () => logout());

/* =========================
   CONFIG
========================= */
const SUPABASE_ARQUIVOS_TABLE = "arquivos";
const SUPABASE_STORAGE_BUCKET = "arquivos";

/* =========================
   SELETORES
========================= */
const botaoSelecionar = document.querySelector("[data-botao-selecionar]");
const inputArquivo = document.querySelector("[data-entrada]");
const listaArquivos = document.querySelector("[data-lista]");
const areaSoltar = document.querySelector("[data-area]");
const formularioEnvio = document.getElementById("formularioEnvio");

/* =========================
   MODAL SUCESSO
========================= */
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

/* =========================
   UTILITÁRIOS
========================= */
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

function gerarNomeSeguro(nome) {
  return String(nome || "arquivo.xlsx")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^\w.\-]+/g, "_");
}

function tipoMimePadraoPorExt(ext) {
  const e = String(ext || "").toLowerCase();

  if (e === "xlsx") {
    return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  }
  if (e === "xls") {
    return "application/vnd.ms-excel";
  }
  if (e === "csv") {
    return "text/csv";
  }

  return "application/octet-stream";
}

function validarArquivoPlanilha(arquivo) {
  if (!arquivo) return false;
  return /\.(xlsx|xls|csv)$/i.test(arquivo.name || "");
}

/* =========================
   MODAL
========================= */
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

/* =========================
   RENDER
========================= */
function renderizarArquivo(arquivo) {
  if (!arquivo) return;

  if (!validarArquivoPlanilha(arquivo)) {
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

/* =========================
   SELEÇÃO MANUAL
========================= */
botaoSelecionar?.addEventListener("click", (e) => {
  e.preventDefault();
  inputArquivo?.click();
});

inputArquivo?.addEventListener("change", () => {
  renderizarArquivo(inputArquivo.files?.[0]);
});

/* =========================
   DRAG AND DROP
========================= */
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

/* =========================
   SUPABASE - UPLOAD
========================= */
async function obterUsuarioLogado() {
  const { data, error } = await supabase.auth.getUser();
  if (error) throw error;

  const user = data?.user || null;
  if (!user?.id) {
    throw new Error("Usuário sem sessão ativa no Supabase.");
  }

  return user;
}

async function uploadArquivoParaBucket({ arquivo, ownerId, arquivoKey }) {
  const ext = pegarExt(arquivo.name);
  const nomeSeguro = gerarNomeSeguro(arquivo.name);
  const mimeType = arquivo.type || tipoMimePadraoPorExt(ext);

  const storagePath = `${ownerId}/${arquivoKey}/${nomeSeguro}`;

  const { data, error } = await supabase.storage
    .from(SUPABASE_STORAGE_BUCKET)
    .upload(storagePath, arquivo, {
      contentType: mimeType,
      upsert: false,
    });

  if (error) {
    throw new Error(`Falha ao enviar arquivo para o bucket: ${error.message}`);
  }

  return {
    storagePath: data?.path || storagePath,
    mimeType,
  };
}

async function salvarMetadataArquivoNoSupabase({
  arquivo,
  ownerId,
  arquivoKey,
  storagePath,
  mimeType,
}) {
  const agoraIso = new Date().toISOString();

  const payload = {
    arquivo_key: arquivoKey,
    owner_id: ownerId,
    nome: arquivo.name,
    tamanho_bytes: arquivo.size,
    mime_type: mimeType || arquivo.type || null,
    storage_path: storagePath,
    atualizado_em: agoraIso,
    data_criacao: agoraIso,
  };

  const { data, error } = await supabase
    .from(SUPABASE_ARQUIVOS_TABLE)
    .upsert(payload, { onConflict: "arquivo_key" })
    .select()
    .single();

  if (error) {
    throw new Error(`Falha ao salvar metadados no Supabase: ${error.message}`);
  }

  return data;
}

async function removerArquivoDoBucketSeExistir(storagePath) {
  if (!storagePath) return;

  try {
    await supabase.storage
      .from(SUPABASE_STORAGE_BUCKET)
      .remove([storagePath]);
  } catch (err) {
    console.warn("Falha ao remover arquivo do bucket após erro:", err?.message || err);
  }
}

/* =========================
   INDEXEDDB - CACHE LOCAL
========================= */
async function salvarArquivoNoIndexedDb({ arquivo, arquivoKey }) {
  if (!window.ibeDb?.salvarArquivoDb) {
    throw new Error("db.js não carregou (window.ibeDb.salvarArquivoDb).");
  }

  const ext = pegarExt(arquivo.name);

  const registro = {
    arquivoKey,
    idNumerico: Date.now(),
    nome: arquivo.name,
    tipo: ext,
    tamanho: formatarTamanho(arquivo.size),
    tamanhoBytes: arquivo.size,
    data: obterDataBr(),
    dataCriacao: Date.now(),
    status: "Pendente",
    blob: arquivo,
  };

  await window.ibeDb.salvarArquivoDb(registro);

  return registro;
}

/* =========================
   ENVIO
========================= */
formularioEnvio?.addEventListener("submit", async (e) => {
  e.preventDefault();

  const arquivo = inputArquivo.files?.[0];
  if (!arquivo) return;

  if (!validarArquivoPlanilha(arquivo)) {
    alert("Selecione um arquivo Excel ou CSV válido.");
    return;
  }

  const arquivoKey = gerarArquivoKey();

  let ownerId = null;
  let storagePath = null;
  let mimeType = arquivo.type || tipoMimePadraoPorExt(pegarExt(arquivo.name));

  try {
    const usuario = await obterUsuarioLogado();
    ownerId = usuario.id;

    // 1) upload do arquivo real para o bucket
    const uploadInfo = await uploadArquivoParaBucket({
      arquivo,
      ownerId,
      arquivoKey,
    });

    storagePath = uploadInfo.storagePath;
    mimeType = uploadInfo.mimeType;

    // 2) salva metadados no Supabase
    await salvarMetadataArquivoNoSupabase({
      arquivo,
      ownerId,
      arquivoKey,
      storagePath,
      mimeType,
    });

    // 3) opcional: mantém cache local no IndexedDB
    await salvarArquivoNoIndexedDb({
      arquivo,
      arquivoKey,
    });

    // 4) limpa UI + modal
    inputArquivo.value = "";
    listaArquivos.innerHTML = "";
    listaArquivos.hidden = true;

    abrirModalSucesso(arquivo.name, arquivoKey);
  } catch (err) {
    console.error(err);

    // rollback do bucket se upload ocorreu mas metadados/cache falharam
    if (storagePath) {
      await removerArquivoDoBucketSeExistir(storagePath);
    }

    alert(err?.message || "Falha ao enviar arquivo.");
  }
});