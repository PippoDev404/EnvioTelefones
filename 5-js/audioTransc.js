import { requireAuth, logout } from "../8-src/auth.js";

const user = await requireAuth({ redirectTo: "../7-login/login.html" });
if (!user) throw new Error("Sem sessão");

document.querySelector("#btnLogout")?.addEventListener("click", () => logout());

const form = document.getElementById("formAuditoria");
const statusBox = document.getElementById("statusBox");
const resultadoContainer = document.getElementById("resultadoContainer");
const toggleTheme = document.getElementById("toggleTheme");
const btnDownloadTranscricoes = document.getElementById("btnDownloadTranscricoes");

const LIMITE_AUDIOS = 10;

// URL da API Python (HuggingFace Spaces em produção, localhost em desenvolvimento)
const API_URL = "https://fmbirl-transcricao-audio.hf.space";

let transcricoesAtuais = [];

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
  titulo.textContent = `Áudio ${index + 1} — ${item.nomeEntrevista || "Sem nome"}`;

  const blocoTranscricao = document.createElement("div");
  blocoTranscricao.className = "BlocoInterno";

  const tituloTranscricao = document.createElement("h4");
  tituloTranscricao.textContent = "Transcrição";

  const conteudoTranscricao = document.createElement("div");
  conteudoTranscricao.className = "CaixaResultado";
  conteudoTranscricao.textContent = item.transcricao || "Nenhuma transcrição retornada.";

  blocoTranscricao.appendChild(tituloTranscricao);
  blocoTranscricao.appendChild(conteudoTranscricao);

  wrapper.appendChild(titulo);
  wrapper.appendChild(blocoTranscricao);

  return wrapper;
}

function renderizarResultados(resultados) {
  limparResultados();

  if (!resultadoContainer) return;

  if (!Array.isArray(resultados) || resultados.length === 0) {
    resultadoContainer.innerHTML = `
      <div class="BlocoEntrevistaPlaceholder">
        Nenhuma transcrição foi retornada.
      </div>
    `;
    if (btnDownloadTranscricoes) btnDownloadTranscricoes.disabled = true;
    return;
  }

  resultados.forEach((item, index) => {
    resultadoContainer.appendChild(criarBlocoResultado(item, index));
  });

  if (btnDownloadTranscricoes) btnDownloadTranscricoes.disabled = false;
}

function baixarTodasAsTranscricoes() {
  if (!Array.isArray(transcricoesAtuais) || transcricoesAtuais.length === 0) {
    setStatus("error", "Não há transcrições para baixar.");
    return;
  }

  const conteudo = transcricoesAtuais
    .map((item, index) => {
      const nome = item.nomeEntrevista || `Áudio ${index + 1}`;
      const texto = item.transcricao || "";
      return `Áudio ${index + 1} — ${nome}\nTranscrição\n${texto}`;
    })
    .join("\n\n----------------------------------------\n\n");

  const blob = new Blob([conteudo], { type: "text/plain;charset=utf-8" });
  const url = URL.createObjectURL(blob);

  const a = document.createElement("a");
  a.href = url;
  a.download = "transcricoes.txt";
  document.body.appendChild(a);
  a.click();
  a.remove();

  URL.revokeObjectURL(url);
}

btnDownloadTranscricoes?.addEventListener("click", baixarTodasAsTranscricoes);

toggleTheme?.addEventListener("click", () => {
  document.documentElement.classList.toggle("active");
  document.body.classList.toggle("active");
});

function limparEspacosBasico(texto) {
  return (texto || "")
    .replace(/\s+/g, " ")
    .replace(/\s+([,.!?;:])/g, "$1")
    .trim();
}

/**
 * Envia arquivo de áudio para a API Python e retorna a transcrição
 */
async function transcreverArquivo(file) {
  setStatus("loading", `Enviando: ${file.name}`);

  const formData = new FormData();
  formData.append("arquivo", file);

  const response = await fetch(`${API_URL}/transcrever`, {
    method: "POST",
    body: formData
  });

  if (!response.ok) {
    const error = await response.json().catch(() => ({}));
    throw new Error(error.detail || `Erro na API: ${response.status}`);
  }

  const resultado = await response.json();

  return {
    nomeEntrevista: resultado.nomeEntrevista || file.name,
    transcricao: resultado.transcricao || ""
  };
}

form?.addEventListener("submit", async (event) => {
  event.preventDefault();

  const audios = Array.from(document.getElementById("audio")?.files || []);

  if (!audios.length) {
    setStatus("error", "Selecione ao menos um áudio.");
    return;
  }

  if (audios.length > LIMITE_AUDIOS) {
    setStatus("error", `Máximo de ${LIMITE_AUDIOS} áudios.`);
    return;
  }

  try {
    limparResultados();
    transcricoesAtuais = [];
    if (btnDownloadTranscricoes) btnDownloadTranscricoes.disabled = true;

    const resultados = [];

    for (const audio of audios) {
      const item = await transcreverArquivo(audio);
      resultados.push(item);
      renderizarResultados(resultados);
    }

    transcricoesAtuais = resultados;
    renderizarResultados(resultados);

    setStatus("success", "Transcrição concluída.");
  } catch (error) {
    console.error(error);
    setStatus("error", error?.message || "Erro ao transcrever os áudios.");
  }
});