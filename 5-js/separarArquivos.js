// 5-js/separarArquivos.js
// 3 botões: Processar | Baixar encontrados (.zip) | Baixar NÃO encontrados (.zip)
import * as XLSX from "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm";
import JSZip from "https://cdn.jsdelivr.net/npm/jszip@3.10.1/+esm";

console.log("✅ separarArquivos.js carregou");

const inputPasta = document.getElementById("inputPasta");
const inputDocumento = document.getElementById("inputDocumento");
const resumoPasta = document.getElementById("resumoPasta");
const resumoDocumento = document.getElementById("resumoDocumento");
const btnProcessar = document.getElementById("btnProcessar");
const btnBaixarZip = document.getElementById("btnBaixarZip");
const btnNaoEncontrados = document.getElementById("btnSalvarPasta"); // reaproveitado p/ "não encontrados"
const statusTexto = document.getElementById("statusTexto");
const listaResultado = document.getElementById("listaResultado");

/* renomeia os botões pra ficar com só 3 ações */
if (btnBaixarZip) btnBaixarZip.innerHTML = '<i class="fa-solid fa-file-zipper"></i> Baixar encontrados (.zip)';
if (btnNaoEncontrados) btnNaoEncontrados.innerHTML = '<i class="fa-solid fa-file-zipper"></i> Baixar NÃO encontrados (.zip)';

let arquivosEncontrados = [];
let arquivosNaoEncontrados = [];
let ultimoResultado = { achados: [], nao: [], duplicados: [] };

/* ---------- números ---------- */

function sequenciasDe(texto) {
    return String(texto).match(/\d+/g) || [];
}

function limpar(numero) {
    return numero.replace(/^0+(?=\d)/, "");
}

function variantes(numero) {
    const n = limpar(numero);
    const vars = new Set([n]);
    if (n.startsWith("55") && n.length >= 12) vars.add(n.slice(2));
    if (n.length > 11) vars.add(n.slice(-11));
    if (n.length > 10) vars.add(n.slice(-10));
    if (n.length > 9) vars.add(n.slice(-9));
    if (n.length > 8) vars.add(n.slice(-8));
    return [...vars].filter((v) => v.length >= 8);
}

function textoDaCelula(celula) {
    if (celula === null || celula === undefined || celula === "") return "";
    if (typeof celula === "number") return celula.toFixed(0);

    let t = String(celula).replace(/\u00a0/g, " ").trim();
    if (/^[+-]?\d+(?:[.,]\d+)?[eE][+-]?\d+$/.test(t)) {
        t = parseFloat(t.replace(",", ".")).toFixed(0);
    }
    return t;
}

// com DDD: (011) 9819-26957 | 55 11 98192-6957 | +55 (11) 9819-26957
const PADRAO_TEL = /^(?:\+?55[\s.\-]*)?\(?\d{2,3}\)?[\s.\-]*\d{4,5}[\s.\-]*\d{4,5}$/;
// sem DDD: 9819-26957 | 981926957
const PADRAO_TEL_SEM_DDD = /^\d{4,5}[\s.\-]*\d{4,5}$/;

function extrairNumerosDe(texto) {
    const t = String(texto).replace(/\u00a0/g, " ").trim();
    const nums = [];
    if (!t) return nums;

    if (PADRAO_TEL.test(t) || PADRAO_TEL_SEM_DDD.test(t)) {
        nums.push(limpar(t.replace(/\D/g, "")));
    }

    for (const seq of sequenciasDe(t)) {
        if (seq.length >= 8) nums.push(limpar(seq));
    }

    return nums;
}

function ehColunaTelefone(cabecalho) {
    const t = String(cabecalho ?? "").trim().toUpperCase();
    return (
        t.startsWith("TEL") ||
        t.includes("[SYS20]") ||
        t.includes("NO. TF") ||
        t.includes("[V48]") ||
        t.includes("CONFIRMAR TELEFONE")
    );
}

/* ---------- lê a planilha (todas as abas) ---------- */

async function lerPlanilha(file) {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array", raw: true });
    const mapaBase = new Map();
    const ocorrencias = new Map();
    let totalRegistros = 0;

    for (const nomeAba of workbook.SheetNames) {
        const linhas = XLSX.utils.sheet_to_json(workbook.Sheets[nomeAba], {
            header: 1,
            raw: true,
            defval: "",
        });
        if (!linhas.length) continue;

        const colunasTelefone = [];
        linhas[0].forEach((c, i) => {
            if (ehColunaTelefone(c)) colunasTelefone.push(i);
        });

        for (const linha of linhas) {
            const sujeito = limpar(sequenciasDe(String(linha[0] ?? "")).join(""));

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

            for (const n of numsDaLinha) {
                if (!ocorrencias.has(n)) ocorrencias.set(n, []);
                ocorrencias.get(n).push(sujeito || "(sem ID)");
            }

            if (achouNaLinha) totalRegistros++;
        }
    }

    const duplicados = [...ocorrencias.entries()]
        .filter(([, ids]) => ids.length > 1)
        .map(([numero, ids]) => ({ numero, ids }));

    const mapaVariante = new Map();
    for (const [numero, sujeito] of mapaBase) {
        for (const v of variantes(numero)) {
            if (!mapaVariante.has(v)) mapaVariante.set(v, { numero, sujeito });
        }
    }

    return { mapaVariante, totalUnicos: mapaBase.size, totalRegistros, duplicados };
}

/* ---------- ZIP genérico (encontrados ou não) ---------- */

async function baixarZipDe(arquivos, nomeZip, botao) {
    if (!arquivos.length) return;

    try {
        if (statusTexto) statusTexto.textContent = `Gerando ${nomeZip}...`;
        botao.disabled = true;

        const zip = new JSZip();
        const usados = new Set();

        for (const file of arquivos) {
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
        link.download = nomeZip;
        link.click();
        URL.revokeObjectURL(url);

        if (statusTexto) statusTexto.textContent = `${nomeZip} baixado com ${arquivos.length} arquivo(s).`;
    } catch (erro) {
        console.error(erro);
        if (statusTexto) statusTexto.textContent = "Erro ao gerar ZIP: " + erro.message;
    } finally {
        atualizarBotoesZip();
    }
}

function atualizarBotoesZip() {
    if (btnBaixarZip) btnBaixarZip.disabled = arquivosEncontrados.length === 0;
    if (btnNaoEncontrados) btnNaoEncontrados.disabled = arquivosNaoEncontrados.length === 0;
}

atualizarBotoesZip(); // começa desabilitado

/* ---------- diagnóstico ---------- */

function mostrarDiagnostico(amostraPlanilha, nao, duplicados) {
    let box = document.getElementById("diagnostico");
    if (!box) {
        box = document.createElement("div");
        box.id = "diagnostico";
        box.style.cssText =
            "margin-top:14px;padding:14px 16px;border:1px solid #c9c9c9;border-radius:12px;background:#efefef;font-size:.85rem;line-height:1.7;word-break:break-all;";
        (listaResultado || document.body).insertAdjacentElement("afterend", box);
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
            (nao.length > 8 ? "<br>… (baixe o ZIP dos não encontrados pra ver todos)" : "")
            : "<br><br>Todos os arquivos da pasta bateram com a planilha. ✅");
}

/* ---------- eventos ---------- */

inputPasta?.addEventListener("change", () => {
    const total = inputPasta.files?.length || 0;
    if (resumoPasta) {
        resumoPasta.textContent = total ? `${total} arquivo(s) na pasta` : "Nenhuma pasta selecionada";
    }
    console.log("📁 Pasta selecionada:", total, "arquivo(s)");
});

inputDocumento?.addEventListener("change", () => {
    if (resumoDocumento) {
        resumoDocumento.textContent = inputDocumento.files[0]
            ? inputDocumento.files[0].name
            : "Nenhuma planilha selecionada";
    }
});

btnProcessar?.addEventListener("click", async () => {
    if (listaResultado) listaResultado.innerHTML = "";
    arquivosEncontrados = [];
    arquivosNaoEncontrados = [];
    atualizarBotoesZip();
    if (statusTexto) statusTexto.textContent = "Lendo planilha e processando...";

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
                nao.push({ file: f, nome: f.webkitRelativePath || f.name, numeros: numsArquivo });
            }
        }

        ultimoResultado = { achados, nao, duplicados };
        arquivosEncontrados = achados.map((a) => a.file);
        arquivosNaoEncontrados = nao.map((n) => n.file);

        if (statusTexto) {
            statusTexto.textContent =
                `Concluído. Telefones na planilha: ${totalRegistros} ` +
                `(${totalUnicos} únicos, ${duplicados.length} duplicado(s)). ` +
                `Encontrados: ${achados.length} • Não encontrados: ${nao.length}.`;
        }

        for (const a of achados) {
            const li = document.createElement("li");
            li.textContent = a.nome;
            listaResultado?.appendChild(li);
        }

        mostrarDiagnostico([...new Set([...mapaVariante.values()].map((h) => h.numero))].slice(0, 10), nao, duplicados);
        atualizarBotoesZip();
    } catch (erro) {
        console.error(erro);
        if (statusTexto) statusTexto.textContent = "Erro: " + erro.message;
    }
});

btnBaixarZip?.addEventListener("click", () => {
    baixarZipDe(arquivosEncontrados, "arquivos_encontrados.zip", btnBaixarZip);
});

btnNaoEncontrados?.addEventListener("click", () => {
    baixarZipDe(arquivosNaoEncontrados, "arquivos_nao_encontrados.zip", btnNaoEncontrados);
});