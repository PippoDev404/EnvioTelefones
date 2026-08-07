// 5-js/separarArquivos.js
// Match por 3 chaves (em ordem):
//  1) telefone
//  2) data/hora com JANELA DE ±15 MIN (+ confirma agente quando disponível)
//  3) agente (só se o arquivo não tiver data no nome)
// Botões: Processar | encontrados (.zip) | NÃO encontrados (.xlsx) | fora da planilha (.zip)
import * as XLSX from "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm";
import JSZip from "https://cdn.jsdelivr.net/npm/jszip@3.10.1/+esm";

console.log("✅ separarArquivos.js carregou");

/* ⏱️ janela de tolerância da data/hora (em minutos) — mexa aqui se precisar */
const TOLERANCIA_MIN = 15;

const inputPasta = document.getElementById("inputPasta");
const inputDocumento = document.getElementById("inputDocumento");
const resumoPasta = document.getElementById("resumoPasta");
const resumoDocumento = document.getElementById("resumoDocumento");
const btnProcessar = document.getElementById("btnProcessar");
const btnBaixarZip = document.getElementById("btnBaixarZip");
const btnNaoEncontrados = document.getElementById("btnSalvarPasta");
const statusTexto = document.getElementById("statusTexto");
const listaResultado = document.getElementById("listaResultado");

if (btnBaixarZip) btnBaixarZip.innerHTML = '<i class="fa-solid fa-file-zipper"></i> Baixar encontrados (.zip)';
if (btnNaoEncontrados) btnNaoEncontrados.innerHTML = '<i class="fa-solid fa-file-excel"></i> Baixar NÃO encontrados (.xlsx)';

let arquivosEncontrados = [];
let arquivosForaDaPlanilha = [];
let estruturaPlanilha = null;
let numerosEncontrados = new Set();
let linhasEncontradas = new Set();
let linhasRestantes = 0;
let ultimoResultado = { achados: [], nao: [], duplicados: [] };

/* ---------- utilidades ---------- */

function sequenciasDe(texto) {
    return String(texto).match(/\d+/g) || [];
}

function limpar(numero) {
    return numero.replace(/^0+(?=\d)/, "");
}

function normalizarTexto(t) {
    return String(t ?? "")
        .toLowerCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .replace(/[^a-z0-9]+/g, " ")
        .replace(/\s+/g, " ")
        .trim();
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
    if (celula instanceof Date) return "";

    let t = String(celula).replace(/\u00a0/g, " ").trim();
    if (/^[+-]?\d+(?:[.,]\d+)?[eE][+-]?\d+$/.test(t)) {
        t = parseFloat(t.replace(",", ".")).toFixed(0);
    }
    return t;
}

/* ---------- data/hora em "minutos UTC" (pra comparar com tolerância) ---------- */

function tsMinutosDataHora(celula) {
    if (celula instanceof Date && !isNaN(celula)) {
        return Math.floor(Date.UTC(
            celula.getFullYear(), celula.getMonth(), celula.getDate(),
            celula.getHours(), celula.getMinutes()
        ) / 60000);
    }

    const t = String(celula ?? "").trim();
    const m = t.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})[^\d]*(\d{1,2}):(\d{2})/);
    if (m) {
        const [, dd, mm, yy, h, mi] = m;
        const y = yy.length === 2 ? 2000 + Number(yy) : Number(yy);
        return Math.floor(Date.UTC(y, Number(mm) - 1, Number(dd), Number(h), Number(mi)) / 60000);
    }
    return null;
}

// do nome do arquivo: 20260806_103015 -> minutos UTC
function tsMinutosDoArquivo(nome) {
    const m = String(nome).match(/(\d{4})(\d{2})(\d{2})[_\-](\d{2})(\d{2})(\d{2})/);
    if (!m) return null;
    const [, y, mo, d, h, mi, s] = m;
    return Math.floor(Date.UTC(Number(y), Number(mo) - 1, Number(d), Number(h), Number(mi), Number(s)) / 60000);
}

/* ---------- agente ---------- */

function infoAgente(texto) {
    const t = normalizarTexto(texto);
    if (!t) return null;
    const nums = sequenciasDe(t).filter((n) => n.length >= 3);
    const nomeSo = t.replace(/\d+/g, " ").replace(/\s+/g, " ").trim();
    if (!nomeSo && !nums.length) return null;
    return { t, numero: nums.length ? nums[nums.length - 1] : "", nomeSo };
}

function infoAgenteDoArquivo(nome) {
    const m = String(nome).match(/agente[\s_]+(.+?)[\s_]+fila/i);
    if (!m) return null;
    return infoAgente(m[1]);
}

function agenteBate(a, b) {
    if (!a || !b) return true; // se um lado não tem agente, não bloqueia
    return (
        (a.numero && b.numero && a.numero === b.numero) ||
        (a.nomeSo && b.nomeSo && (a.nomeSo.includes(b.nomeSo) || b.nomeSo.includes(a.nomeSo)))
    );
}

/* ---------- telefone ---------- */

const PADRAO_TEL = /^(?:\+?55[\s.\-]*)?\(?\d{2,3}\)?[\s.\-]*\d{4,5}[\s.\-]*\d{4,5}$/;
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

function ehColunaData(cabecalho) {
    return String(cabecalho ?? "").toUpperCase().includes("DATA");
}

function ehColunaAgente(cabecalho) {
    return String(cabecalho ?? "").toUpperCase().includes("AGENTE");
}

/* ---------- lê a planilha ---------- */

async function lerPlanilha(file) {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array", raw: true, cellDates: true });
    const mapaBase = new Map();
    const ocorrencias = new Map();
    const registrosData = []; // { ts, agente, sujeito }
    const agentesPlanilha = [];
    const folhas = [];
    let totalRegistros = 0;

    for (const nomeAba of workbook.SheetNames) {
        const linhas = XLSX.utils.sheet_to_json(workbook.Sheets[nomeAba], {
            header: 1,
            raw: true,
            defval: "",
        });
        if (!linhas.length) continue;

        const colunasTelefone = [];
        let colunaData = -1;
        let colunaAgente = -1;

        linhas[0].forEach((c, i) => {
            if (ehColunaTelefone(c)) colunasTelefone.push(i);
            if (colunaData < 0 && ehColunaData(c)) colunaData = i;
            if (colunaAgente < 0 && ehColunaAgente(c)) colunaAgente = i;
        });

        const rows = [];

        linhas.forEach((linha, idxLinha) => {
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

            const ts = colunaData >= 0 ? tsMinutosDataHora(linha[colunaData]) : null;
            const agente = colunaAgente >= 0 ? infoAgente(linha[colunaAgente]) : null;

            if (ts !== null) registrosData.push({ ts, agente, sujeito });
            if (agente) agentesPlanilha.push({ ...agente, sujeito });

            for (const n of numsDaLinha) {
                if (!ocorrencias.has(n)) ocorrencias.set(n, []);
                ocorrencias.get(n).push(sujeito || "(sem ID)");
            }

            if (idxLinha > 0 && (achouNaLinha || ts !== null || agente)) totalRegistros++;

            rows.push({ aoa: linha, nums: numsDaLinha, sujeito });
        });

        folhas.push({ nome: nomeAba, rows });
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

    return {
        mapaVariante,
        registrosData,
        agentesPlanilha,
        folhas,
        totalUnicos: mapaBase.size,
        totalRegistros,
        duplicados,
    };
}

/* ---------- busca data/hora com tolerância ---------- */

function buscarPorDataAgente(registrosData, tsArq, agArq) {
    if (tsArq === null) return null;

    let melhor = null;
    let melhorDiff = Infinity;

    for (const r of registrosData) {
        const diff = Math.abs(r.ts - tsArq);
        if (diff > TOLERANCIA_MIN) continue;
        if (!agenteBate(r.agente, agArq)) continue;

        if (diff < melhorDiff) {
            melhorDiff = diff;
            melhor = r;
        }
    }

    return melhor;
}

/* ---------- ZIP genérico (STORE = rápido) ---------- */

async function baixarZipDe(arquivos, nomeZip, botao) {
    if (!arquivos.length) return;

    try {
        if (statusTexto) statusTexto.textContent = `Gerando ${nomeZip} (modo rápido)...`;
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

        const blob = await zip.generateAsync({ type: "blob", compression: "STORE" });
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
        atualizarBotoes();
    }
}

/* ---------- EXCEL dos NÃO ENCONTRADOS ---------- */

function linhaTemMatch(row) {
    if (linhasEncontradas.has(row.sujeito)) return true;
    for (const n of row.nums) {
        if (numerosEncontrados.has(n)) return true;
    }
    return false;
}

function baixarPlanilhaRestante() {
    if (!estruturaPlanilha) return;

    try {
        if (statusTexto) statusTexto.textContent = "Gerando planilha restante...";

        const wb = XLSX.utils.book_new();

        for (const folha of estruturaPlanilha) {
            const aoa = [];

            folha.rows.forEach((row, idx) => {
                if (idx === 0) {
                    aoa.push(row.aoa);
                    return;
                }
                if (!linhaTemMatch(row)) aoa.push(row.aoa);
            });

            const ws = XLSX.utils.aoa_to_sheet(aoa);
            XLSX.utils.book_append_sheet(wb, ws, folha.nome.slice(0, 31));
        }

        XLSX.writeFile(wb, "planilha_nao_encontrados.xlsx");

        if (statusTexto) {
            statusTexto.textContent = `Planilha restante baixada: ${linhasRestantes} linha(s) sem áudio encontrado.`;
        }
    } catch (erro) {
        console.error(erro);
        if (statusTexto) statusTexto.textContent = "Erro ao gerar planilha: " + erro.message;
    }
}

/* ---------- 4º botão: fora da planilha ---------- */

let btnFora = document.getElementById("btnForaPlanilha");
if (!btnFora) {
    const containerAcoes =
        document.querySelector(".acoesSeparar") ||
        btnProcessar?.parentElement ||
        document.body;

    btnFora = document.createElement("button");
    btnFora.id = "btnForaPlanilha";
    btnFora.className = "botaoTerciario";
    btnFora.innerHTML = '<i class="fa-solid fa-file-zipper"></i> Baixar fora da planilha (.zip)';
    btnFora.disabled = true;
    containerAcoes.appendChild(btnFora);
    btnFora.addEventListener("click", () =>
        baixarZipDe(arquivosForaDaPlanilha, "arquivos_fora_da_planilha.zip", btnFora)
    );
}

/* ---------- estado dos botões ---------- */

function atualizarBotoes() {
    if (btnBaixarZip) btnBaixarZip.disabled = arquivosEncontrados.length === 0;
    if (btnNaoEncontrados) btnNaoEncontrados.disabled = linhasRestantes === 0;
    if (btnFora) btnFora.disabled = arquivosForaDaPlanilha.length === 0;
}

atualizarBotoes();

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
        `Exemplos de números lidos das colunas de telefone: ${amostraPlanilha.join(", ") || "— (nessa planilha o match foi por data/hora ou agente)"}` +
        (duplicados.length
            ? `<br><br><strong>⚠️ Números duplicados (${duplicados.length}):</strong><br>` +
            duplicados.slice(0, 10)
                .map((d) => `${d.numero} → aparece em ${d.ids.length} linhas: ${d.ids.join(", ")}`)
                .join("<br>")
            : "") +
        (nao.length
            ? `<br><br><strong>Arquivos da pasta que NÃO estão na planilha (${nao.length}):</strong><br>` + fora.join("<br>") +
            (nao.length > 8 ? "<br>… (baixe o ZIP “fora da planilha” pra levar todos)" : "")
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
    arquivosForaDaPlanilha = [];
    numerosEncontrados = new Set();
    linhasEncontradas = new Set();
    linhasRestantes = 0;
    atualizarBotoes();
    if (statusTexto) statusTexto.textContent = "Lendo planilha e processando...";

    try {
        const arquivos = Array.from(inputPasta.files || []);
        const documento = inputDocumento.files?.[0];

        if (!arquivos.length) throw new Error("Selecione a pasta com os arquivos.");
        if (!documento) throw new Error("Selecione o Excel ou CSV com os números.");

        const { mapaVariante, registrosData, agentesPlanilha, folhas, totalUnicos, totalRegistros, duplicados } =
            await lerPlanilha(documento);

        if (!mapaVariante.size && !registrosData.length && !agentesPlanilha.length) {
            throw new Error("Planilha sem telefone, sem data e sem agente — nada pra comparar.");
        }

        estruturaPlanilha = folhas;

        const achados = [];
        const nao = [];
        let cTel = 0, cData = 0, cAgente = 0;

        for (const f of arquivos) {
            const numsArquivo = extrairNumerosDe(f.name);
            const tsArq = tsMinutosDoArquivo(f.name);
            const agArq = infoAgenteDoArquivo(f.name);
            let evidencia = null;
            let regra = "";

            // 1) telefone
            for (const n of numsArquivo) {
                for (const v of variantes(n)) {
                    const hit = mapaVariante.get(v);
                    if (hit) {
                        evidencia = { numeroArquivo: n, numeroPlanilha: hit.numero, sujeito: hit.sujeito };
                        regra = "tel";
                        break;
                    }
                }
                if (evidencia) break;
            }

            // 2) data/hora com tolerância (+ agente confirmando)
            if (!evidencia) {
                const hit = buscarPorDataAgente(registrosData, tsArq, agArq);
                if (hit) {
                    evidencia = { numeroArquivo: String(tsArq), numeroPlanilha: String(hit.ts), sujeito: hit.sujeito };
                    regra = "data";
                }
            }

            // 3) agente (só se o arquivo NÃO tem data no nome)
            if (!evidencia && tsArq === null && agArq) {
                const hit = agentesPlanilha.find((a) => agenteBate(a, agArq) && (a.nomeSo || a.numero));
                if (hit) {
                    evidencia = { numeroArquivo: agArq.t, numeroPlanilha: hit.t, sujeito: hit.sujeito };
                    regra = "agente";
                }
            }

            if (evidencia) {
                if (regra === "tel") { cTel++; numerosEncontrados.add(evidencia.numeroPlanilha); }
                if (regra === "data") cData++;
                if (regra === "agente") cAgente++;
                linhasEncontradas.add(evidencia.sujeito);
                achados.push({ file: f, nome: f.webkitRelativePath || f.name, ...evidencia });
            } else {
                nao.push({ file: f, nome: f.webkitRelativePath || f.name, numeros: numsArquivo });
            }
        }

        linhasRestantes = 0;
        for (const folha of estruturaPlanilha) {
            folha.rows.forEach((row, idx) => {
                if (idx > 0 && !linhaTemMatch(row)) linhasRestantes++;
            });
        }

        ultimoResultado = { achados, nao, duplicados };
        arquivosEncontrados = achados.map((a) => a.file);
        arquivosForaDaPlanilha = nao.map((n) => n.file);

        if (statusTexto) {
            statusTexto.textContent =
                `Concluído. Registros na planilha: ${totalRegistros}. ` +
                `Encontrados: ${achados.length} (telefone: ${cTel} • data/hora ±${TOLERANCIA_MIN}min: ${cData} • agente: ${cAgente}) • ` +
                `Fora da planilha: ${nao.length} • Linhas restantes: ${linhasRestantes}.`;
        }

        for (const a of achados) {
            const li = document.createElement("li");
            li.textContent = a.nome;
            listaResultado?.appendChild(li);
        }

        mostrarDiagnostico([...new Set([...mapaVariante.values()].map((h) => h.numero))].slice(0, 10), nao, duplicados);
        atualizarBotoes();
    } catch (erro) {
        console.error(erro);
        if (statusTexto) statusTexto.textContent = "Erro: " + erro.message;
    }
});

btnBaixarZip?.addEventListener("click", () =>
    baixarZipDe(arquivosEncontrados, "arquivos_encontrados.zip", btnBaixarZip)
);

btnNaoEncontrados?.addEventListener("click", baixarPlanilhaRestante);