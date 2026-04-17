import { requireAuth, logout } from "../8-src/auth.js";

const user = await requireAuth({ redirectTo: "../7-login/login.html" });
if (!user) throw new Error("Sem sessão");

document.querySelector("#btnLogout")?.addEventListener("click", () => logout());

const form = document.getElementById("formAuditoria");
const statusBox = document.getElementById("statusBox");
const resultadoContainer = document.getElementById("resultadoContainer");
const toggleTheme = document.getElementById("toggleTheme");

const N8N_WEBHOOK_URL = "https://n8n.srv962474.hstgr.cloud/webhook/chat";
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
  if (Array.isArray(data?.resultados)) return data.resultados;
  if (Array.isArray(data?.data?.resultados)) return data.data.resultados;
  if (Array.isArray(data)) return data;

  if (data?.resposta || data?.transcricao) {
    return [{
      nomeEntrevista: data.nomeEntrevista || data.audioName || "Entrevista",
      resposta: data.resposta,
      transcricao: data.transcricao
    }];
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

  return result.value;
}

/* ============================= */
/* NORMALIZAÇÃO DE ÁUDIO (NOVO) */
/* ============================= */

async function converterParaWav16kMono(file) {
  const audioContext = new (window.AudioContext || window.webkitAudioContext)();
  const arrayBuffer = await file.arrayBuffer();
  const audioBuffer = await audioContext.decodeAudioData(arrayBuffer);

  const offlineCtx = new OfflineAudioContext(1, audioBuffer.duration * 16000, 16000);
  const source = offlineCtx.createBufferSource();

  const monoBuffer = offlineCtx.createBuffer(1, audioBuffer.length, audioBuffer.sampleRate);

  for (let i = 0; i < audioBuffer.numberOfChannels; i++) {
    const channelData = audioBuffer.getChannelData(i);
    const monoData = monoBuffer.getChannelData(0);
    for (let j = 0; j < channelData.length; j++) {
      monoData[j] += channelData[j] / audioBuffer.numberOfChannels;
    }
  }

  source.buffer = monoBuffer;
  source.connect(offlineCtx.destination);
  source.start();

  const renderedBuffer = await offlineCtx.startRendering();

  return bufferToWave(renderedBuffer, 16000);
}

function bufferToWave(abuffer, sampleRate) {
  const numOfChan = abuffer.numberOfChannels;
  const length = abuffer.length * numOfChan * 2 + 44;
  const buffer = new ArrayBuffer(length);
  const view = new DataView(buffer);

  let offset = 0;

  function writeString(str) {
    for (let i = 0; i < str.length; i++) {
      view.setUint8(offset++, str.charCodeAt(i));
    }
  }

  writeString("RIFF");
  view.setUint32(offset, 36 + abuffer.length * 2, true); offset += 4;
  writeString("WAVE");
  writeString("fmt ");
  view.setUint32(offset, 16, true); offset += 4;
  view.setUint16(offset, 1, true); offset += 2;
  view.setUint16(offset, 1, true); offset += 2;
  view.setUint32(offset, sampleRate, true); offset += 4;
  view.setUint32(offset, sampleRate * 2, true); offset += 4;
  view.setUint16(offset, 2, true); offset += 2;
  view.setUint16(offset, 16, true); offset += 2;
  writeString("data");
  view.setUint32(offset, abuffer.length * 2, true); offset += 4;

  const channelData = abuffer.getChannelData(0);
  let index = 0;

  for (let i = 0; i < channelData.length; i++, offset += 2) {
    let sample = Math.max(-1, Math.min(1, channelData[i]));
    view.setInt16(offset, sample * 0x7fff, true);
  }

  return new Blob([view], { type: "audio/wav" });
}

/* ============================= */

form?.addEventListener("submit", async (event) => {
  event.preventDefault();

  const audios = Array.from(document.getElementById("audio")?.files || []);
  const questionario = document.getElementById("questionario")?.files?.[0];

  if (!audios.length || !questionario) {
    setStatus("error", "Selecione os áudios e o questionário.");
    return;
  }

  if (audios.length > LIMITE_AUDIOS) {
    setStatus("error", `Máximo de ${LIMITE_AUDIOS} áudios.`);
    return;
  }

  try {
    setStatus("loading", "Lendo questionário...");
    const questionarioTexto = await extrairTextoQuestionario(questionario);

    setStatus("loading", "Normalizando áudios...");

    const audiosConvertidos = [];

    for (const audio of audios) {
      try {
        const wavBlob = await converterParaWav16kMono(audio);
        const novoArquivo = new File([wavBlob], audio.name + ".wav", { type: "audio/wav" });
        audiosConvertidos.push(novoArquivo);
      } catch (e) {
        console.warn("Falha ao converter:", audio.name);
      }
    }

    const formData = new FormData();

    audiosConvertidos.forEach((audioFile) => {
      formData.append("audio", audioFile);
    });

    formData.append("questionarioTexto", questionarioTexto);
    formData.append("questionarioNome", questionario.name);

    setStatus("loading", "Enviando para análise...");

    const response = await fetch(N8N_WEBHOOK_URL, {
      method: "POST",
      body: formData
    });

    const data = await response.json();

    const resultados = normalizarResultados(data);

    renderizarResultados(resultados);
    setStatus("success", "Análise concluída.");
  } catch (error) {
    console.error(error);
    setStatus("error", error.message);
  }
});