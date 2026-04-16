import { requireAuth, logout } from "../8-src/auth.js";

const user = await requireAuth({ redirectTo: "../7-login/login.html" });
if (!user) throw new Error("Sem sessão");

document.querySelector("#btnLogout")?.addEventListener("click", () => logout());

const form = document.getElementById("formAuditoria");
const statusBox = document.getElementById("statusBox");
const resultadoContainer = document.getElementById("resultadoContainer");
const toggleTheme = document.getElementById("toggleTheme");

const N8N_WEBHOOK_URL = "https://n8n.srv962474.hstgr.cloud/webhook-test/chat";
const LIMITE_AUDIOS = 5;

function setStatus(tipo, mensagem) {
  if (!statusBox) return;

  statusBox.className = "StatusBox";
  if (tipo === "loading") statusBox.classList.add("status-loading");
  if (tipo === "success") statusBox.classList.add("status-sucesso");
  if (tipo === "error") statusBox.classList.add("status-erro");
  if (tipo === "neutral") statusBox.classList.add("status-neutro");

  statusBox.textContent = mensagem;
}

function limparResultados() {
  if (!resultadoContainer) return;
  resultadoContainer.innerHTML = "";
}

function criarBlocoResultado(item, index) {
  const wrapper = document.createElement("article");
  wrapper.className = "BlocoEntrevista";

  const titulo = document.createElement("h3");
  titulo.className = "BlocoEntrevistaTitulo";
  titulo.textContent = `Entrevista ${index + 1} — ${item.nomeEntrevista || item.audioName || "Sem nome"}`;

  const grade = document.createElement("div");
  grade.className = "BlocoEntrevistaGrade";

  const colunaResposta = document.createElement("div");
  colunaResposta.className = "BlocoInterno";

  const tituloResposta = document.createElement("h4");
  tituloResposta.textContent = "Resposta da IA";

  const conteudoResposta = document.createElement("div");
  conteudoResposta.className = "CaixaResultado";
  conteudoResposta.textContent = item.resposta || "Nenhuma resposta retornada.";

  colunaResposta.appendChild(tituloResposta);
  colunaResposta.appendChild(conteudoResposta);

  const colunaTranscricao = document.createElement("div");
  colunaTranscricao.className = "BlocoInterno";

  const tituloTranscricao = document.createElement("h4");
  tituloTranscricao.textContent = "Transcrição";

  const conteudoTranscricao = document.createElement("div");
  conteudoTranscricao.className = "CaixaResultado";
  conteudoTranscricao.textContent = item.transcricao || "Nenhuma transcrição retornada.";

  colunaTranscricao.appendChild(tituloTranscricao);
  colunaTranscricao.appendChild(conteudoTranscricao);

  grade.appendChild(colunaResposta);
  grade.appendChild(colunaTranscricao);

  wrapper.appendChild(titulo);
  wrapper.appendChild(grade);

  return wrapper;
}

function renderizarResultados(resultados) {
  limparResultados();

  if (!resultadoContainer) return;

  if (!Array.isArray(resultados) || resultados.length === 0) {
    resultadoContainer.innerHTML = `
      <div class="BlocoEntrevistaPlaceholder">
        Nenhum resultado foi retornado.
      </div>
    `;
    return;
  }

  resultados.forEach((item, index) => {
    resultadoContainer.appendChild(criarBlocoResultado(item, index));
  });
}

function normalizarResultados(data) {
  if (Array.isArray(data?.resultados)) {
    return data.resultados;
  }

  if (Array.isArray(data?.data?.resultados)) {
    return data.data.resultados;
  }

  if (Array.isArray(data)) {
    return data;
  }

  if (data?.resposta || data?.transcricao) {
    return [
      {
        nomeEntrevista: data.nomeEntrevista || data.audioName || "Entrevista",
        resposta: data.resposta,
        transcricao: data.transcricao
      }
    ];
  }

  return [];
}

toggleTheme?.addEventListener("click", () => {
  document.documentElement.classList.toggle("active");
  document.body.classList.toggle("active");
});

async function extrairTextoQuestionario(file) {
  const extensao = file.name.split(".").pop()?.toLowerCase();

  if (extensao !== "docx") {
    throw new Error("O questionário deve estar em formato .docx.");
  }

  if (!window.mammoth) {
    throw new Error("Mammoth não foi carregado no navegador.");
  }

  const arrayBuffer = await file.arrayBuffer();
  const result = await window.mammoth.extractRawText({ arrayBuffer });

  if (result.messages?.length) {
    console.log("Avisos do Mammoth:", result.messages);
  }

  return result.value;
}

form?.addEventListener("submit", async (event) => {
  event.preventDefault();

  const audios = Array.from(document.getElementById("audio")?.files || []);
  const questionario = document.getElementById("questionario")?.files?.[0];

  if (!audios.length || !questionario) {
    setStatus("error", "Selecione os áudios e o questionário antes de continuar.");
    return;
  }

  if (audios.length > LIMITE_AUDIOS) {
    setStatus("error", `Você pode enviar no máximo ${LIMITE_AUDIOS} áudios por vez.`);
    return;
  }

  limparResultados();

  try {
    setStatus("loading", "Lendo questionário...");
    const questionarioTexto = await extrairTextoQuestionario(questionario);

    if (!questionarioTexto || !questionarioTexto.trim()) {
      throw new Error("Não foi possível extrair texto do questionário.");
    }

    setStatus("loading", `Enviando ${audios.length} áudio(s) para análise...`);

    const formData = new FormData();

    audios.forEach((audioFile) => {
      formData.append("audio", audioFile);
    });

    formData.append("questionarioTexto", questionarioTexto);
    formData.append("questionarioNome", questionario.name);

    const response = await fetch(N8N_WEBHOOK_URL, {
      method: "POST",
      body: formData
    });

    const contentType = response.headers.get("content-type") || "";
    let data;

    if (contentType.includes("application/json")) {
      data = await response.json();
    } else {
      const rawText = await response.text();
      throw new Error(`O webhook não retornou JSON válido. Retorno: ${rawText.slice(0, 300)}`);
    }

    console.log("Resposta completa do n8n:", data);

    if (!response.ok) {
      throw new Error(data.message || data.error || "Erro ao processar a análise.");
    }

    const resultados = normalizarResultados(data);

    if (!resultados.length) {
      throw new Error("O webhook não retornou entrevistas no formato esperado.");
    }

    renderizarResultados(resultados);
    setStatus("success", "Análise concluída com sucesso.");
  } catch (error) {
    console.error(error);

    limparResultados();

    if (resultadoContainer) {
      resultadoContainer.innerHTML = `
        <div class="BlocoEntrevistaPlaceholder">
          Ocorreu um erro ao consultar o workflow.
        </div>
      `;
    }

    setStatus("error", error.message || "Erro inesperado.");
  }
});