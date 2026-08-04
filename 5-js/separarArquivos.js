// 5-js/separarArquivos.js
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

/* ---------- números ---------- */

function sequenciasDe(texto) {
    return String(texto).match(/\d+/g) || [];
}

function limpar(numero) {
    return numero.replace(/^0+(?=\d)/, "");
}

// "apelidos" do mesmo número pra comparação ficar flexível:
// sem zeros à esquerda, sem o 55 do começo (DDI), e últimos 10/11 dígitos
function variantes(numero) {
    const n = limpar(numero);
    const vars = new Set([n]);
    if (n.startsWith("55") && n.length >= 12) vars.add(n.slice(2));
    if (n.length > 11) vars.add(n.slice(-11));
    if (n.length > 10) vars.add(n.slice(-10));
    return [...vars].filter((v) => v.length >= 8);
}

/* ---------- leitura da planilha (todas as abas/colunas) ---------- */

async function lerNumerosDaPlanilha(file) {
    const buffer = await file.arrayBuffer();
    const base = new Set();

    const workbook = XLSX.read(buffer, { type: "array", raw: true });

    for (const nomeAba of workbook.SheetNames) {
        const linhas = XLSX.utils.sheet_to_json(workbook.Sheets[nomeAba], {
            header: 1,
            raw: true,
            defval: "",
        });

        for (const linha of linhas) {
            for (const celula of linha) {
                if (celula === null || celula === undefined || celula === "") continue;

                let texto = String(celula);

                // número puro -> sem notação científica
                if (typeof celula === "number") {
                    texto = celula.toFixed(0);
                } else if (/^[+-]?\d+(?:[.,]\d+)?[eE][+-]?\d+$/.test(texto.trim())) {
                    texto = parseFloat(texto.replace(",", ".")).toFixed(0);
                }

                for (const seq of sequenciasDe(texto)) {
                    if (seq.length >= 8) base.add(limpar(seq));
                }
            }
        }
    }

    return [...base];
}

/* ---------- diagnóstico visual ---------- */

function mostrarDiagnostico(amostraPlanilha, amostraArquivos) {
    let box = document.getElementById("diagnostico");
    if (!box) {
        box = document.createElement("div");
        box.id = "diagnostico";
        box.style.cssText =
            "margin-top:14px;padding:14px 16px;border:1px solid #c9c9c9;border-radius:12px;background:#efefef;font-size:.85rem;line-height:1.7;word-break:break-all;";
        listaResultado.insertAdjacentElement("afterend", box);
    }

    box.innerHTML =
        `<strong>🔎 Diagnóstico (pra ver por que não bateu):</strong><br>` +
        `Exemplos de números lidos da planilha: ${amostraPlanilha.join(", ") || "—"}<br>` +
        `Exemplos de números nos nomes dos arquivos: ${amostraArquivos.join(", ") || "️ nenhum número com 8+ dígitos nos nomes dos arquivos"}`;
}

/* ---------- eventos ---------- */

inputPasta.addEventListener("change", () => {
    const total = inputPasta.files?.length || 0;
    resumoPasta.textContent = total
        ? `${total} arquivo(s) na pasta`
        : "Nenhuma pasta selecionada";
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
    statusTexto.textContent = "Lendo planilha e processando...";

    try {
        const arquivos = Array.from(inputPasta.files || []);
        const documento = inputDocumento.files?.[0];

        if (!arquivos.length) throw new Error("Selecione a pasta com os arquivos.");
        if (!documento) throw new Error("Selecione o Excel ou CSV com os números.");

        const numerosBase = await lerNumerosDaPlanilha(documento);
        if (!numerosBase.length) {
            throw new Error("Nenhum número válido (8+ dígitos) encontrado na planilha.");
        }

        // todas as variantes dos números da planilha
        const setPlanilha = new Set();
        for (const n of numerosBase) {
            for (const v of variantes(n)) setPlanilha.add(v);
        }

        // amostra dos números que aparecem nos NOMES dos arquivos
        const amostraArquivos = [];
        for (const f of arquivos) {
            for (const s of sequenciasDe(f.name)) {
                if (s.length >= 8 && amostraArquivos.length < 10) amostraArquivos.push(limpar(s));
            }
            if (amostraArquivos.length >= 10) break;
        }

        // match: algum número do nome do arquivo bate com alguma variante da planilha
        arquivosEncontrados = arquivos.filter((f) =>
            sequenciasDe(f.name).some(
                (s) => s.length >= 8 && variantes(s).some((v) => setPlanilha.has(v))
            )
        );

        mostrarDiagnostico(numerosBase.slice(0, 10), amostraArquivos);

        statusTexto.textContent =
            `Concluído. Números únicos na planilha: ${numerosBase.length}. ` +
            `Arquivos encontrados: ${arquivosEncontrados.length} de ${arquivos.length}.`;

        for (const f of arquivosEncontrados) {
            const li = document.createElement("li");
            li.textContent = f.webkitRelativePath || f.name;
            listaResultado.appendChild(li);
        }

        if (arquivosEncontrados.length) {
            btnBaixarZip.disabled = false;
            btnSalvarPasta.disabled = false;
        } else {
            statusTexto.textContent += " (Nenhum arquivo correspondeu — veja o diagnóstico abaixo.)";
        }
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
                nome =
                    ponto > 0
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
                nome =
                    ponto > 0
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