import { requireAuth, logout } from "../8-src/auth.js";
import { supabase } from "../8-src/supabaseClient.js";

const user = await requireAuth({ redirectTo: "../7-login/login.html" });
if (!user) throw new Error("Sem sessão");

document.querySelector("#btnLogout")?.addEventListener("click", () => logout());

const inputArquivo = document.getElementById("inputArquivo");
const areaDrop = document.getElementById("areaDrop");
const btnSelecionarArquivo = document.getElementById("btnSelecionarArquivo");
const btnTrocarArquivo = document.getElementById("btnTrocarArquivo");
const btnRemoverArquivo = document.getElementById("btnRemoverArquivo");

const uploadStateEmpty = document.getElementById("uploadStateEmpty");
const uploadStateFilled = document.getElementById("uploadStateFilled");
const arquivoRecebidoNome = document.getElementById("arquivoRecebidoNome");
const arquivoRecebidoMeta = document.getElementById("arquivoRecebidoMeta");

const btnProcessarArquivo = document.getElementById("btnProcessarArquivo");

const emptyFiles = document.getElementById("emptyFiles");
const listaArquivos = document.getElementById("listaArquivos");

const modalPreview = document.getElementById("modalPreview");
const modalBackdrop = document.getElementById("modalBackdrop");
const btnFecharModal = document.getElementById("btnFecharModal");
const arquivoPreviewHead = document.getElementById("arquivoPreviewHead");
const arquivoPreviewBody = document.getElementById("arquivoPreviewBody");

const modalConfirmacaoRemocao = document.getElementById("modalConfirmacaoRemocao");
const modalConfirmacaoBackdrop = document.getElementById("modalConfirmacaoBackdrop");
const btnCancelarRemocao = document.getElementById("btnCancelarRemocao");
const btnConfirmarRemocao = document.getElementById("btnConfirmarRemocao");

// Opcional no HTML:
// <select id="selectTipoCota">
//   <option value="auto">Automático</option>
//   <option value="numero">Número</option>
//   <option value="estado">Estado</option>
// </select>
const selectTipoCota = document.getElementById("selectTipoCota");

// Opcional no HTML para exibir detecção:
// <div id="tipoCotaInfo"></div>
const tipoCotaInfo = document.getElementById("tipoCotaInfo");

let arquivoSelecionado = null;
let arquivoGeradoUnico = null;

inicializar();

function inicializar() {
  inicializarUpload();
  inicializarProcessamentoUnico();
  inicializarModal();
  inicializarModalRemocao();
  atualizarEstadoArquivo();
}

function inicializarUpload() {
  btnSelecionarArquivo?.addEventListener("click", () => {
    inputArquivo?.click();
  });

  btnTrocarArquivo?.addEventListener("click", () => {
    inputArquivo?.click();
  });

  btnRemoverArquivo?.addEventListener("click", () => {
    abrirModalConfirmacaoRemocao();
  });

  inputArquivo?.addEventListener("change", (event) => {
    const file = event.target.files?.[0] || null;
    definirArquivoSelecionado(file);
  });

  if (areaDrop) {
    areaDrop.addEventListener("click", (event) => {
      const clicouEmBotao = event.target.closest("button");
      if (!clicouEmBotao) inputArquivo?.click();
    });

    areaDrop.addEventListener("keydown", (event) => {
      if (event.key === "Enter" || event.key === " ") {
        event.preventDefault();
        inputArquivo?.click();
      }
    });

    ["dragenter", "dragover"].forEach((evento) => {
      areaDrop.addEventListener(evento, (event) => {
        event.preventDefault();
        areaDrop.classList.add("dragover");
      });
    });

    ["dragleave", "drop"].forEach((evento) => {
      areaDrop.addEventListener(evento, (event) => {
        event.preventDefault();

        if (evento === "dragleave") {
          const destino = event.relatedTarget;
          if (destino && areaDrop.contains(destino)) return;
        }

        areaDrop.classList.remove("dragover");
      });
    });

    areaDrop.addEventListener("drop", (event) => {
      const file = event.dataTransfer?.files?.[0] || null;
      if (!file) return;

      definirArquivoSelecionado(file);

      if (inputArquivo) {
        try {
          const dt = new DataTransfer();
          dt.items.add(file);
          inputArquivo.files = dt.files;
        } catch (error) {
          console.warn("Não foi possível sincronizar o arquivo no input.", error);
        }
      }
    });
  }
}

function inicializarProcessamentoUnico() {
  if (!btnProcessarArquivo) return;

  btnProcessarArquivo.addEventListener("click", async () => {
    if (!arquivoSelecionado) {
      alert("Selecione ou arraste um arquivo de cota antes de continuar.");
      return;
    }

    try {
      setBotaoCarregando(btnProcessarArquivo, true, "Lendo e gerando...");

      const cotas = await lerArquivoCota(arquivoSelecionado);

      if (!cotas.length) {
        alert("A cota foi lida, mas nenhuma linha válida foi encontrada.");
        return;
      }

      const tipoDetectado = cotas[0]?.tipoCota || "desconhecido";
      atualizarInfoTipoCota(tipoDetectado);

      const registrosEncontrados = await buscarRegistrosParaCota(cotas);

      if (!registrosEncontrados.length) {
        alert("Nenhum registro foi encontrado no banco para essa cota.");
        return;
      }

      const agora = new Date();

      arquivoGeradoUnico = {
        id: agora.getTime(),
        nomeArquivo: montarNomeArquivoUnico(arquivoSelecionado?.name, agora),
        origem: arquivoSelecionado?.name || "cota.csv",
        criadoEm: formatarDataHora(agora),
        headers: [
          "IDP",
          "ESTADO",
          "CIDADE",
          "Cod_estado",
          "Cod_municipio",
          "Cod_meso",
          "Cod_micro",
          "REGIÃO CIDADE",
          "SETOR DENTRO DA CIDADE",
          "SEXO",
          "IDADE",
          "GRAU DE INSTRUÇÃO",
          "OCUPAÇÃO",
          "RENDA FAMILIAR",
          "DATA DA PESQUISA",
          "CATEGORIA",
          "NOME",
          "ENDEREÇO COMPLETO",
          "ENDEREÇO",
          "NÚMERO",
          "BAIRRO",
          "CEP",
          "EMAIL",
          "TF1",
          "TF2",
          "TF3",
          "TF4",
          "Nº PESQ.",
          "DIA PESQ.",
          "QT TF1",
          "QT TF2",
          "QT TF3",
          "QT TF4"
        ],
        rows: registrosEncontrados.map(mapearRegistroParaLinhaFinal)
      };

      renderizarArquivoUnico();

      alert(`Arquivo único gerado com sucesso. ${arquivoGeradoUnico.rows.length} linha(s) encontradas.`);
    } catch (error) {
      console.error("Erro ao processar a cota:", error);
      alert(`Não foi possível processar a planilha: ${error.message || error}`);
    } finally {
      setBotaoCarregando(btnProcessarArquivo, false, "Ler e gerar arquivo");
    }
  });
}

function inicializarModal() {
  btnFecharModal?.addEventListener("click", fecharModalPreview);
  modalBackdrop?.addEventListener("click", fecharModalPreview);

  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape" && modalPreview && !modalPreview.classList.contains("hidden")) {
      fecharModalPreview();
    }
  });
}

function inicializarModalRemocao() {
  btnCancelarRemocao?.addEventListener("click", fecharModalConfirmacaoRemocao);
  modalConfirmacaoBackdrop?.addEventListener("click", fecharModalConfirmacaoRemocao);

  btnConfirmarRemocao?.addEventListener("click", () => {
    resetarArquivoSelecionado();
    fecharModalConfirmacaoRemocao();
  });

  document.addEventListener("keydown", (event) => {
    if (
      event.key === "Escape" &&
      modalConfirmacaoRemocao &&
      !modalConfirmacaoRemocao.classList.contains("hidden")
    ) {
      fecharModalConfirmacaoRemocao();
    }
  });
}

function abrirModalConfirmacaoRemocao() {
  const temAlgoParaRemover = Boolean(arquivoSelecionado || arquivoGeradoUnico);
  if (!temAlgoParaRemover) return;

  modalConfirmacaoRemocao?.classList.remove("hidden");
  modalConfirmacaoRemocao?.setAttribute("aria-hidden", "false");
}

function fecharModalConfirmacaoRemocao() {
  modalConfirmacaoRemocao?.classList.add("hidden");
  modalConfirmacaoRemocao?.setAttribute("aria-hidden", "true");
}

function definirArquivoSelecionado(file) {
  if (!file) {
    resetarArquivoSelecionado();
    return;
  }

  if (!ehArquivoPermitido(file)) {
    alert("Envie um arquivo válido (.xlsx, .xls ou .csv).");
    resetarArquivoSelecionado();
    return;
  }

  arquivoSelecionado = file;
  arquivoGeradoUnico = null;

  if (arquivoRecebidoNome) arquivoRecebidoNome.textContent = file.name || "arquivo.xlsx";
  if (arquivoRecebidoMeta) {
    arquivoRecebidoMeta.textContent = `${formatarTamanhoArquivo(file.size)} • ${obterBadgeArquivo(file)}`;
  }

  limparArquivoUnico();
  atualizarEstadoArquivo();

  if (btnProcessarArquivo) {
    btnProcessarArquivo.disabled = false;
  }
}

function resetarArquivoSelecionado() {
  arquivoSelecionado = null;
  arquivoGeradoUnico = null;

  if (inputArquivo) inputArquivo.value = "";
  if (btnProcessarArquivo) btnProcessarArquivo.disabled = true;

  limparArquivoSelecionadoVisual();
  limparArquivoUnico();
  fecharModalPreview();
  atualizarEstadoArquivo();
  atualizarInfoTipoCota("");
}

function atualizarEstadoArquivo() {
  const temArquivo = Boolean(arquivoSelecionado);

  areaDrop?.classList.toggle("has-file", temArquivo);

  uploadStateEmpty?.classList.toggle("hidden", temArquivo);
  uploadStateFilled?.classList.toggle("hidden", !temArquivo);

  if (!temArquivo) {
    if (arquivoRecebidoNome) arquivoRecebidoNome.textContent = "arquivo.xlsx";
    if (arquivoRecebidoMeta) arquivoRecebidoMeta.textContent = "0 KB • XLSX";
  }
}

function limparArquivoSelecionadoVisual() {
  if (arquivoRecebidoNome) arquivoRecebidoNome.textContent = "arquivo.xlsx";
  if (arquivoRecebidoMeta) arquivoRecebidoMeta.textContent = "0 KB • XLSX";
}

function limparArquivoUnico() {
  arquivoGeradoUnico = null;

  if (!listaArquivos || !emptyFiles) return;
  listaArquivos.innerHTML = "";
  listaArquivos.classList.add("hidden");
  emptyFiles.classList.remove("hidden");
}

function atualizarInfoTipoCota(tipo) {
  if (!tipoCotaInfo) return;

  if (!tipo) {
    tipoCotaInfo.textContent = "";
    return;
  }

  const mapa = {
    auto: "Automático",
    numero: "Número",
    estado: "Estado",
    desconhecido: "Desconhecido"
  };

  tipoCotaInfo.textContent = `Tipo de cota: ${mapa[tipo] || tipo}`;
}

async function lerArquivoCota(file) {
  const nome = file.name.toLowerCase();

  let matriz = [];

  if (nome.endsWith(".csv")) {
    const texto = await lerTextoArquivoCsv(file);
    matriz = texto
      .split(/\r?\n/)
      .map((linha) => separarLinhaCsv(linha))
      .filter((linha) => linha.some((celula) => String(celula).trim() !== ""));
  } else {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array" });
    const primeiraAba = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[primeiraAba];

    matriz = XLSX.utils.sheet_to_json(worksheet, {
      header: 1,
      defval: ""
    });
  }

  if (!matriz.length) return [];

  const linhasNaoVazias = matriz.filter((linha) =>
    Array.isArray(linha) && linha.some((celula) => String(celula).trim() !== "")
  );

  if (!linhasNaoVazias.length) return [];

  const header = linhasNaoVazias[0].map((item) => normalizarCabecalho(item));
  const linhas = linhasNaoVazias.slice(1);

  const idx = {
    dia: encontrarIndiceCabecalho(header, ["dia"]),
    pesquisador: encontrarIndiceCabecalho(header, ["pesq", "pesq."]),
    regiao: encontrarIndiceCabecalho(header, ["regiao", "regiao cidade", "regiao_cidade"]),
    setor1: encontrarIndiceCabecalho(header, ["setores1", "setor1", "setor 1", "numero1", "numero 1", "estado1", "estado 1"]),
    qt1: encontrarIndiceCabecalho(header, ["qt1", "qt 1", "quantidade1", "quantidade 1"]),
    setor2: encontrarIndiceCabecalho(header, ["setores2", "setor2", "setor 2", "numero2", "numero 2", "estado2", "estado 2"]),
    qt2: encontrarIndiceCabecalho(header, ["qt2", "qt 2", "quantidade2", "quantidade 2"]),
    setor3: encontrarIndiceCabecalho(header, ["setores3", "setor3", "setor 3", "numero3", "numero 3", "estado3", "estado 3"]),
    qt3: encontrarIndiceCabecalho(header, ["qt3", "qt 3", "quantidade3", "quantidade 3"])
  };

  if (idx.pesquisador < 0) {
    throw new Error("A coluna PESQ. não foi encontrada no arquivo de cota.");
  }

  const setoresColetados = [];

  linhas.forEach((linha) => {
    [idx.setor1, idx.setor2, idx.setor3].forEach((indice) => {
      const valor = String(pegarValorLinha(linha, indice) || "").trim();
      if (valor) setoresColetados.push(valor);
    });
  });

  const tipoManual = (selectTipoCota?.value || "auto").toLowerCase();
  const tipoDetectado = tipoManual !== "auto" ? tipoManual : detectarTipoDaCota(setoresColetados);

  const cotas = [];

  linhas.forEach((linha) => {
    const pesquisador = String(pegarValorLinha(linha, idx.pesquisador) || "").trim();
    const dia = String(pegarValorLinha(linha, idx.dia) || "").trim();
    const regiao = String(pegarValorLinha(linha, idx.regiao) || "").trim();

    if (!pesquisador) return;

    const itens = [];

    [
      { valor: idx.setor1, qt: idx.qt1 },
      { valor: idx.setor2, qt: idx.qt2 },
      { valor: idx.setor3, qt: idx.qt3 }
    ].forEach(({ valor, qt }) => {
      const valorCota = String(pegarValorLinha(linha, valor) || "").trim();
      const quantidade = converterNumero(pegarValorLinha(linha, qt));

      if (valorCota && quantidade > 0) {
        itens.push({
          valor: limparValorCota(valorCota, tipoDetectado),
          quantidade,
          regiao
        });
      }
    });

    if (itens.length) {
      cotas.push({
        dia,
        pesquisador,
        tipoCota: tipoDetectado,
        itens
      });
    }
  });

  return cotas;
}

function detectarTipoDaCota(valores) {
  const validos = valores
    .map((v) => String(v ?? "").trim())
    .filter(Boolean);

  if (!validos.length) return "desconhecido";

  const totalNumericos = validos.filter((v) => /^\d+$/.test(v)).length;
  const proporcaoNumericos = totalNumericos / validos.length;

  if (proporcaoNumericos >= 0.7) return "numero";

  const ufs = new Set([
    "AC", "AL", "AP", "AM", "BA", "CE", "DF", "ES", "GO", "MA",
    "MT", "MS", "MG", "PA", "PB", "PR", "PE", "PI", "RJ", "RN",
    "RS", "RO", "RR", "SC", "SP", "SE", "TO"
  ]);

  const totalUFs = validos.filter((v) => ufs.has(String(v).trim().toUpperCase())).length;
  const proporcaoUFs = totalUFs / validos.length;

  if (proporcaoUFs >= 0.7) return "estado";

  return "desconhecido";
}

function limparValorCota(valor, tipoCota) {
  const texto = String(valor ?? "").trim();
  if (tipoCota === "estado") return texto.toUpperCase();
  return texto;
}

function obterColunaBusca(tipoCota) {
  if (tipoCota === "estado") return "estado";
  return "setor_dentro_cidade";
}

async function buscarRegistrosParaCota(cotas) {
  const resultadoFinal = [];
  const idsUsados = new Set();

  for (const cota of cotas) {
    for (const item of cota.itens) {
      const registros = await buscarNoBancoPorCota({
        valor: item.valor,
        quantidade: item.quantidade,
        pesquisador: cota.pesquisador,
        dia: cota.dia,
        tipoCota: cota.tipoCota,
        regiao: item.regiao,
        idsIgnorados: idsUsados
      });

      for (const reg of registros) {
        const chave = String(reg.id ?? reg.idp ?? "");
        if (chave && idsUsados.has(chave)) continue;
        if (chave) idsUsados.add(chave);
        resultadoFinal.push(reg);
      }
    }
  }

  return resultadoFinal;
}

async function buscarNoBancoPorCota({
  valor,
  quantidade,
  pesquisador,
  dia,
  tipoCota,
  regiao,
  idsIgnorados = new Set()
}) {
  const quantidadeDesejada = Number(quantidade) || 0;
  if (quantidadeDesejada <= 0) return [];

  const colunaBusca = obterColunaBusca(tipoCota);
  const limiteBusca = Math.max(quantidadeDesejada * 3, quantidadeDesejada + 50);

  let query = supabase
    .from("base_mestra")
    .select(`
      id,
      idp,
      estado,
      cidade,
      cod_estado,
      cod_municipio,
      cod_meso,
      cod_micro,
      regiao_cidade,
      setor_dentro_cidade,
      sexo,
      idade,
      grau_instrucao,
      ocupacao,
      renda_familiar,
      data_pesquisa,
      categoria,
      nome,
      endereco_completo,
      endereco,
      numero,
      bairro,
      cep,
      email,
      tf1,
      tf2,
      tf3,
      tf4,
      numero_pesq,
      dia_pesq,
      qt_tf1,
      qt_tf2,
      qt_tf3,
      qt_tf4
    `)
    .eq(colunaBusca, valor)
    .limit(limiteBusca);

  if (regiao) {
    query = query.eq("regiao_cidade", regiao);
  }

  const { data, error } = await query;

  if (error) {
    console.error("Erro ao buscar no banco:", error);
    return [];
  }

  const registrosFiltrados = [];
  const idsLocais = new Set();

  for (const reg of data || []) {
    const idAtual = reg?.id != null ? String(reg.id) : "";

    if (idAtual && (idsIgnorados.has(idAtual) || idsLocais.has(idAtual))) {
      continue;
    }

    registrosFiltrados.push({
      ...reg,
      data_pesquisa: reg.data_pesquisa ?? dia,
      numero_pesq: pesquisador,
      dia_pesq: dia
    });

    if (idAtual) idsLocais.add(idAtual);

    if (registrosFiltrados.length >= quantidadeDesejada) break;
  }

  return registrosFiltrados;
}

function mapearRegistroParaLinhaFinal(registro) {
  return [
    registro.idp ?? "",
    registro.estado ?? "",
    registro.cidade ?? "",
    registro.cod_estado ?? "",
    registro.cod_municipio ?? "",
    registro.cod_meso ?? "",
    registro.cod_micro ?? "",
    registro.regiao_cidade ?? "",
    registro.setor_dentro_cidade ?? "",
    registro.sexo ?? "",
    registro.idade ?? "",
    registro.grau_instrucao ?? "",
    registro.ocupacao ?? "",
    registro.renda_familiar ?? "",
    registro.data_pesquisa ?? "",
    registro.categoria ?? "",
    registro.nome ?? "",
    registro.endereco_completo ?? "",
    registro.endereco ?? "",
    registro.numero ?? "",
    registro.bairro ?? "",
    registro.cep ?? "",
    registro.email ?? "",
    registro.tf1 ?? "",
    registro.tf2 ?? "",
    registro.tf3 ?? "",
    registro.tf4 ?? "",
    registro.numero_pesq ?? "",
    registro.dia_pesq ?? "",
    registro.qt_tf1 ?? "",
    registro.qt_tf2 ?? "",
    registro.qt_tf3 ?? "",
    registro.qt_tf4 ?? ""
  ];
}

function renderizarArquivoUnico() {
  if (!listaArquivos || !emptyFiles || !arquivoGeradoUnico) return;

  emptyFiles.classList.add("hidden");
  listaArquivos.classList.remove("hidden");

  listaArquivos.innerHTML = `
    <div class="file-row">
      <div class="file-info">
        <div class="file-name">${escapeHtml(arquivoGeradoUnico.nomeArquivo)}</div>
        <div class="file-meta">
          Origem: ${escapeHtml(arquivoGeradoUnico.origem)} •
          Linhas: ${arquivoGeradoUnico.rows.length} •
          Colunas: ${arquivoGeradoUnico.headers.length} •
          Criado em: ${escapeHtml(arquivoGeradoUnico.criadoEm)}
        </div>
      </div>

      <div class="file-actions">
        <button class="btn btn-secondary" id="btnVerArquivoUnico" type="button">VER</button>
        <button class="btn btn-success" id="btnDownloadArquivoUnico" type="button">DOWNLOAD EXCEL</button>
        <button class="btn btn-warning" id="btnEnviarDashboard" type="button">Enviar para a Dashboard</button>
      </div>
    </div>
  `;

  document.getElementById("btnVerArquivoUnico")?.addEventListener("click", () => {
    abrirModalPreview(arquivoGeradoUnico);
  });

  document.getElementById("btnDownloadArquivoUnico")?.addEventListener("click", () => {
    baixarArquivoUnico(arquivoGeradoUnico);
  });

  document.getElementById("btnEnviarDashboard")?.addEventListener("click", () => {
    alert("A ação de envio para a Dashboard ainda não foi implementada.");
  });
}

function abrirModalPreview(arquivo) {
  if (!modalPreview || !arquivoPreviewHead || !arquivoPreviewBody) return;

  arquivoPreviewHead.innerHTML = `
    <tr>
      ${arquivo.headers.map((header) => `<th>${escapeHtml(header)}</th>`).join("")}
    </tr>
  `;

  arquivoPreviewBody.innerHTML = "";

  const linhasPreview = arquivo.rows.slice(0, 20);

  linhasPreview.forEach((linha) => {
    const tr = document.createElement("tr");
    tr.innerHTML = linha
      .map((valor) => `<td>${escapeHtml(String(valor ?? ""))}</td>`)
      .join("");
    arquivoPreviewBody.appendChild(tr);
  });

  modalPreview.classList.remove("hidden");
  modalPreview.setAttribute("aria-hidden", "false");
}

function fecharModalPreview() {
  modalPreview?.classList.add("hidden");
  modalPreview?.setAttribute("aria-hidden", "true");
}

function baixarArquivoUnico(arquivo) {
  const dados = [arquivo.headers, ...arquivo.rows];
  const worksheet = XLSX.utils.aoa_to_sheet(dados);
  const workbook = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(workbook, worksheet, "Telefones");
  XLSX.writeFile(workbook, arquivo.nomeArquivo);
}

function montarNomeArquivoUnico(nomeOriginal, data) {
  const base = String(nomeOriginal || "arquivo")
    .replace(/\.[^.]+$/, "")
    .replace(/[^\w\-]+/g, "_");

  const stamp = [
    data.getFullYear(),
    String(data.getMonth() + 1).padStart(2, "0"),
    String(data.getDate()).padStart(2, "0"),
    "_",
    String(data.getHours()).padStart(2, "0"),
    String(data.getMinutes()).padStart(2, "0"),
    String(data.getSeconds()).padStart(2, "0")
  ].join("");

  return `${base}_gerado_${stamp}.xlsx`;
}

function setBotaoCarregando(botao, carregando, textoPadrao = "Processar") {
  if (!botao) return;

  if (carregando) {
    botao.dataset.labelOriginal = botao.textContent;
    botao.textContent = textoPadrao;
    botao.disabled = true;
    return;
  }

  botao.textContent = botao.dataset.labelOriginal || textoPadrao;
  botao.disabled = false;

  if (!arquivoSelecionado && botao === btnProcessarArquivo) {
    botao.disabled = true;
  }
}

function ehArquivoPermitido(file) {
  const nome = String(file?.name || "").toLowerCase();
  return [".xlsx", ".xls", ".csv"].some((ext) => nome.endsWith(ext));
}

function obterBadgeArquivo(file) {
  return obterExtensaoArquivo(file?.name || "").toUpperCase() || "ARQ";
}

function obterExtensaoArquivo(nome) {
  const partes = String(nome || "").split(".");
  return partes.length > 1 ? partes.pop() : "arquivo";
}

function formatarTamanhoArquivo(bytes) {
  const valor = Number(bytes || 0);

  if (valor < 1024) return `${valor} B`;
  if (valor < 1024 * 1024) return `${(valor / 1024).toFixed(1)} KB`;
  if (valor < 1024 * 1024 * 1024) return `${(valor / (1024 * 1024)).toFixed(1)} MB`;

  return `${(valor / (1024 * 1024 * 1024)).toFixed(1)} GB`;
}

function formatarDataHora(data) {
  const dd = String(data.getDate()).padStart(2, "0");
  const mm = String(data.getMonth() + 1).padStart(2, "0");
  const yyyy = data.getFullYear();
  const hh = String(data.getHours()).padStart(2, "0");
  const mi = String(data.getMinutes()).padStart(2, "0");

  return `${dd}/${mm}/${yyyy} ${hh}:${mi}`;
}

function normalizarCabecalho(valor) {
  return String(valor || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toLowerCase();
}

function encontrarIndiceCabecalho(headers, alternativas) {
  return headers.findIndex((header) =>
    alternativas.some((alt) => header === normalizarCabecalho(alt))
  );
}

function pegarValorLinha(linha, indice) {
  if (indice < 0) return "";
  return linha?.[indice] ?? "";
}

function converterNumero(valor) {
  const texto = String(valor ?? "").trim().replace(",", ".");
  const numero = Number(texto);
  return Number.isFinite(numero) ? numero : 0;
}

function separarLinhaCsv(linha) {
  const resultado = [];
  let atual = "";
  let dentroDeAspas = false;

  for (let i = 0; i < linha.length; i++) {
    const char = linha[i];
    const prox = linha[i + 1];

    if (char === '"' && dentroDeAspas && prox === '"') {
      atual += '"';
      i++;
      continue;
    }

    if (char === '"') {
      dentroDeAspas = !dentroDeAspas;
      continue;
    }

    if (char === ";" && !dentroDeAspas) {
      resultado.push(atual);
      atual = "";
      continue;
    }

    if (char === "," && !dentroDeAspas) {
      resultado.push(atual);
      atual = "";
      continue;
    }

    atual += char;
  }

  resultado.push(atual);
  return resultado;
}

function lerTextoArquivoCsv(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();

    reader.onload = () => {
      const buffer = reader.result;
      const uint8 = new Uint8Array(buffer);

      let texto = new TextDecoder("utf-8", { fatal: false }).decode(uint8);

      const temCaractereSubstituicao = texto.includes("�");
      if (temCaractereSubstituicao) {
        texto = new TextDecoder("windows-1252", { fatal: false }).decode(uint8);
      }

      resolve(texto.replace(/^﻿/, ""));
    };

    reader.onerror = () => {
      reject(reader.error || new Error("Não foi possível ler o arquivo CSV."));
    };

    reader.readAsArrayBuffer(file);
  });
}

function escapeHtml(valor) {
  return String(valor ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}