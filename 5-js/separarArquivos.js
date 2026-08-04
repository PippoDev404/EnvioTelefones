// 5-js/separarArquivos.js
// Aceita os 2 tipos de planilha:
//  1) número seco: 31971013437
//  2) formatado:  (011) 9819-26957  -> vira 11981926957
// E também notação científica (5,539E+12)
// + mostra números DUPLICADOS e em quais linhas
import * as XLSX from "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm";
import JSZip from "https://cdn.jsdelivr.net/npm/jszip@3.10.1/+esm";

const inputPasta = document.getElementById("inputPasta");
const inputDocumento = document.getElementById("inputDocumento");
const resumoPasta = document.getElementById("resumoPasta");
const resumoDocumento = document.getElementById("resumoDocumento");
const btnProcessar = document.getElementById("btnProcessar");
const btnBaixarZip = document.getElementById("btnBaixarZip");
const btnSalvarPasta = document.getElementById("btnSalvarPasta");
const statusTexto = document.getElementById("statusTexto");
const listaResultado = document.getElementById("listaResultado");

let arquivosEncontrados = [];
let ultimoResultado = { achados: [], nao: [], duplicados: [] };

/* ---------- números ---------- */

function sequenciasDe(texto) {
    return String(texto).match(/\d+/g) || [];
}

function limpar(numero) {
    return numero.replace(/^0+(?=\d)/, "");
}

// "apelidos" do número pra comparação flexível (sem 55, últimos 10/11)
function variantes(numero) {
    const n = limpar(numero);
    const vars = new Set([n]);
    if (n.startsWith("55") && n.length >= 12) vars.add(n.slice(2));
    if (n.length > 11) vars.add(n.slice(-11));
    if (n.length > 10) vars.add(n.slice(-10));
    return [...vars].filter((v) => v.length >= 8);
}

/* ---------- leitura de célula (resolve notação científica) ---------- */

function textoDaCelula(celula) {
    if (celula === null || celula === undefined || celula === "") return "";
    if (typeof celula === "number") return celula.toFixed(0);

    let t = String(celula).trim();
    if (/^[+-]?\d+(?:[.,]\d+)?[eE][+-]?\d+$/.test(t)) {
        t = parseFloat(t.replace(",", ".")).toFixed(0);
    }
    return t;
}

/* ---------- reconhece telefone FORMATADO: (011) 9819-26957 ---------- */

const PADRAO_TEL = /^\(?\d{2,3}\)?[\s.\-]*\d{4,5}[\s.\-]*\d{4,5}$/;

// extrai números de uma célula OU de um nome de arquivo
function extrairNumerosDe(texto) {
    const t = String(texto).trim();
    const nums = [];
    if (!t) return nums;

    // tipo 2: formatado -> junta todos os dígitos (011 vira 11)
    if (PADRAO_TEL.test(t)) {
        nums.push(limpar(t.replace(/\D/g, "")));
    }

    // tipo 1: sequências grandes de dígitos (11+ secos, IDs, etc.)
    for (const seq of sequenciasDe(t)) {
        if (seq.length >= 8) nums.push(limpar(seq));
    }

    return nums;
}

/* ---------- reconhece as colunas de telefone pelo cabeçalho ---------- */

function ehColunaTelefone(cabecalho) {
    const t = String(cabecalho ?? "").trim().toUpperCase();
    return (
        t.includes("[SYS20]") ||
        t.includes("NO. TF") ||
        t.includes("[V48]") ||
        t.includes("CONFIRMAR TELEFONE") ||
        t === "TEL." ||
        t === "TEL"
    );
}

/* ---------- lê TODAS as abas e puxa as colunas de telefone ---------- */

async function lerPlanilha(file) {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array", raw: true });
    const mapaBase = new Map();    // numero -> ID da primeira linha
    const ocorrencias = new Map(); // numero -> [IDs das linhas onde aparece]
    let totalRegistros = 0;

    for (const nomeAba of workbook.SheetNames) {
        const linhas = XLSX.utils.sheet_to_json(workbook.Sheets[nomeAba], {
            header: 1,
            raw: true,
            defval: "",
        });
        if (!linhas.length) continue;

        // acha os índices das colunas de telefone na linha de cabeçalho
        const colunasTelefone = [];
        linhas[0].forEach((c, i) => {
            if (ehColunaTelefone(c)) colunasTelefone.push(i);
        });

        for (const linha of linhas) {
            const sujeito = limpar(sequenciasDe(String(linha[0] ?? "")).join(""));

            // se achou as colunas, usa só elas; senão, lê a linha inteira (fallback)
            const celulasAlvo = colunasTelefone.length
                ? colunasTelefone.map((i) => linha[i])
                : linha;

            let achouNaLinha = false;
            const numsDaLinha = new Set();

            for (const celula of celulasAlvo) {
                const texto = textoDaCelula(celula);
                if (!texto) continue;

                const nums = extrairNumerosDe(texto);
                if (nums.length) achouNaLinha = true;

                for (const n of nums) {
                    numsDaLinha.add(n);
                    if (!mapaBase.has(n)) mapaBase.set(n, sujeito);
                }
            }

            // conta 1 ocorrência por linha (mesmo que o número esteja em 2 colunas da mesma linha)
            for (const n of numsDaLinha) {
                if (!ocorrencias.has(n)) ocorrencias.set(n, []);
                ocorrencias.get(n).push(sujeito || "(sem ID)");
            }

            if (achouNaLinha) totalRegistros++;
        }
    }

    // números que aparecem em mais de 1 linha
    const duplicados = [...ocorrencias.entries()]
        .filter(([, ids]) => ids.length > 1)
        .map(([numero, ids]) => ({ numero, ids }));

    // variante -> { numero original, sujeito }
    const mapaVariante = new Map();
    for (const [numero, sujeito] of mapaBase) {
        for (const v of variantes(numero)) {
            if (!mapaVariante.has(v)) mapaVariante.set(v, { numero, sujeito });
        }
    }

    return { mapaVariante, totalUnicos: mapaBase.size, totalRegistros, duplicados };
}

/* ---------- botão de relatório (criado via JS) ---------- */

let btnRelatorio = document.getElementById("btnRelatorio");
if (!btnRelatorio) {
    btnRelatorio = document.createElement("button");
    btnRelatorio.id = "btnRelatorio";
    btnRelatorio.className = "botaoTerciario";
    btnRelatorio.innerHTML = '<i class="fa-solid fa-clipboard-check"></i> Relatório de conferência';
    btnRelatorio.disabled = true;
    document.querySelector(".acoesSeparar").appendChild(btnRelatorio);
    btnRelatorio.addEventListener("click", baixarRelatorio);
}

function baixarRelatorio() {
    const { achados, nao, duplicados } = ultimoResultado;
    const linhas = [
        ["TIPO", "ARQUIVO_OU_NUMERO", "NUMERO_NO_ARQUIVO", "NUMERO_NA_PLANILHA", "ID_LINHA_PLANILHA"].join(";"),
    ];

    for (const a of achados) {
        linhas.push(["OK", a.nome, a.numeroArquivo, a.numeroPlanilha, a.sujeito].join(";"));
    }
    for (const n of nao) {
        linhas.push(["FORA_DA_PLANILHA", n.nome, n.numeros.join("|"), "", ""].join(";"));
    }
    for (const d of duplicados) {
        linhas.push(["DUPLICADO_NA_PLANILHA", d.numero, "", "", d.ids.join("|")].join(";"));
    }

    const blob = new Blob(["\ufeff" + linhas.join("\r\n")], { type: "text/csv;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "relatorio_conferencia.csv";
    link.click();
    URL.revokeObjectURL(url);

    statusTexto.textContent = "Relatório de conferência baixado. Abra no Excel e audite linha por linha.";
}

/* ---------- diagnóstico ---------- */

function mostrarDiagnostico(amostraPlanilha, nao, duplicados) {
    let box = document.getElementById("diagnostico");
    if (!box) {
        box = document.createElement("div");
        box.id = "diagnostico";
        box.style.cssText =
            "margin-top:14px;padding:14px 16px;border:1px solid #c9c9c9;border-radius:12px;background:#efefef;font-size:.85rem;line-height:1.7;word-break:break-all;";
        listaResultado.insertAdjacentElement("afterend", box);
    }

    const fora = nao.slice(0, 8).map((n) => `${n.nome} → [${n.numeros.join(", ") || "sem números"}]`);

    box.innerHTML =
        `<strong>🔎 Conferência:</strong><br>` +
        `Exemplos de números lidos das colunas de telefone: ${amostraPlanilha.join(", ") || "—"}` +
        (duplicados.length
            ? `<br><br><strong>⚠️ Números duplicados (${duplicados.length}):</strong><br>` +
            duplicados.slice(0, 10)
                .map((d) => `${d.numero} → aparece em ${d.ids.length} linhas: ${d.ids.join(", ")}`)
                .join("<br>") +
            (duplicados.length > 10 ? "<br>… (veja o relatório completo)" : "")
            : "<br><br>Nenhum número duplicado. ✅") +
        (nao.length
            ? `<br><br><strong>Arquivos que NÃO estão na planilha (${nao.length}):</strong><br>` + fora.join("<br>") +
            (nao.length > 8 ? "<br>… (veja o relatório completo)" : "")
            : "<br><br>Todos os arquivos da pasta bateram com a planilha. ✅");
}

/* ---------- eventos ---------- */

inputPasta.addEventListener("change", () => {
    const total = inputPasta.files?.length || 0;
    resumoPasta.textContent = total ? `${total} arquivo(s) na pasta` : "Nenhuma pasta selecionada";
});

inputDocumento.addEventListener("change", () => {
    resumoDocumento.textContent = inputDocumento.files[0]
        ? inputDocumento.files[0].name
        : "Nenhuma planilha selecionada";
});

btnProcessar.addEventListener("click", async () => {
    listaResultado.innerHTML = "";
    arquivosEncontrados = [];
    btnBaixarZip.disabled = true;
    btnSalvarPasta.disabled = true;
    btnRelatorio.disabled = true;
    statusTexto.textContent = "Lendo planilha e processando...";

    try {
        const arquivos = Array.from(inputPasta.files || []);
        const documento = inputDocumento.files?.[0];

        if (!arquivos.length) throw new Error("Selecione a pasta com os arquivos.");
        if (!documento) throw new Error("Selecione o Excel ou CSV com os números.");

        const { mapaVariante, totalUnicos, totalRegistros, duplicados } = await lerPlanilha(documento);
        if (!mapaVariante.size) throw new Error("Nenhum número válido nas colunas de telefone.");

        const achados = [];
        const nao = [];

        for (const f of arquivos) {
            const numsArquivo = extrairNumerosDe(f.name);
            let evidencia = null;

            for (const n of numsArquivo) {
                for (const v of variantes(n)) {
                    const hit = mapaVariante.get(v);
                    if (hit) {
                        evidencia = { numeroArquivo: n, numeroPlanilha: hit.numero, sujeito: hit.sujeito };
                        break;
                    }
                }
                if (evidencia) break;
            }

            if (evidencia) {
                achados.push({ file: f, nome: f.webkitRelativePath || f.name, ...evidencia });
            } else {
                nao.push({ nome: f.webkitRelativePath || f.name, numeros: numsArquivo });
            }
        }

        ultimoResultado = { achados, nao, duplicados };
        arquivosEncontrados = achados.map((a) => a.file);

        statusTexto.textContent =
            `Concluído. Telefones na planilha: ${totalRegistros} ` +
            `(${totalUnicos} únicos, ${duplicados.length} duplicado(s)). ` +
            `Arquivos encontrados: ${achados.length} de ${arquivos.length} • Fora: ${nao.length}.`;

        for (const a of achados) {
            const li = document.createElement("li");
            li.textContent = a.nome;
            listaResultado.appendChild(li);
        }

        mostrarDiagnostico([...new Set([...mapaVariante.values()].map((h) => h.numero))].slice(0, 10), nao, duplicados);

        if (achados.length) {
            btnBaixarZip.disabled = false;
            btnSalvarPasta.disabled = false;
        }
        btnRelatorio.disabled = false;
    } catch (erro) {
        console.error(erro);
        statusTexto.textContent = "Erro: " + erro.message;
    }
});

btnBaixarZip.addEventListener("click", async () => {
    try {
        statusTexto.textContent = "Gerando ZIP...";
        btnBaixarZip.disabled = true;

        const zip = new JSZip();
        const usados = new Set();

        for (const file of arquivosEncontrados) {
            let nome = file.name;
            let i = 1;
            while (usados.has(nome)) {
                const ponto = file.name.lastIndexOf(".");
                nome = ponto > 0
                    ? `${file.name.slice(0, ponto)} (${i})${file.name.slice(ponto)}`
                    : `${file.name} (${i})`;
                i++;
            }
            usados.add(nome);
            zip.file(nome, file);
        }

        const blob = await zip.generateAsync({ type: "blob" });
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = "arquivos_separados_ibespe.zip";
        link.click();
        URL.revokeObjectURL(url);

        statusTexto.textContent = "ZIP gerado e baixado com sucesso!";
    } catch (erro) {
        console.error(erro);
        statusTexto.textContent = "Erro ao gerar ZIP: " + erro.message;
    } finally {
        btnBaixarZip.disabled = arquivosEncontrados.length === 0;
    }
});

btnSalvarPasta.addEventListener("click", async () => {
    if (!("showDirectoryPicker" in window)) {
        alert("Seu navegador não suporta salvar direto em pasta. Use Chrome/Edge ou baixe o ZIP.");
        return;
    }

    try {
        statusTexto.textContent = "Aguardando seleção da pasta de destino...";
        btnSalvarPasta.disabled = true;

        const dirHandle = await window.showDirectoryPicker({ mode: "readwrite" });
        const usados = new Set();

        for (const file of arquivosEncontrados) {
            let nome = file.name;
            let i = 1;
            while (usados.has(nome)) {
                const ponto = file.name.lastIndexOf(".");
                nome = ponto > 0
                    ? `${file.name.slice(0, ponto)} (${i})${file.name.slice(ponto)}`
                    : `${file.name} (${i})`;
                i++;
            }
            usados.add(nome);

            const fileHandle = await dirHandle.getFileHandle(nome, { create: true });
            const writable = await fileHandle.createWritable();
            await writable.write(file);
            await writable.close();
        }

        statusTexto.textContent = `${arquivosEncontrados.length} arquivo(s) salvo(s) na pasta selecionada.`;
    } catch (erro) {
        if (erro.name !== "AbortError") {
            console.error(erro);
            statusTexto.textContent = "Erro ao salvar: " + erro.message;
        } else {
            statusTexto.textContent = "Operação cancelada.";
        }
    } finally {
        btnSalvarPasta.disabled = arquivosEncontrados.length === 0;
    }
});