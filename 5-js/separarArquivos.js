// 5-js/separarArquivos.js
import * as XLSX from "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/+esm";
import JSZip from "https://cdn.jsdelivr.net/npm/jszip@3.10.1/+esm";

console.log("✅ separarArquivos.js carregou");

const TOLERANCIA_MIN = 30;

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
if (btnNaoEncontrados) btnNaoEncontrados.innerHTML = '<i class="fa-solid fa-file-csv"></i> Baixar NÃO encontrados (.csv)';

let arquivosEncontrados = [];
let arquivosEncontradosPorData = [];
let arquivosForaDaPlanilha = [];
let estruturaPlanilha = null;
let numerosEncontrados = new Set();
let linhasEncontradas = new Set();
let linhasRestantes = 0;
let ultimoResultado = { achados: [], nao: [], duplicados: [] };

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

function normalizarTelefone(texto) {
    const todosDigitos = String(texto).replace(/\D/g, "");
    if (!todosDigitos || todosDigitos.length < 8) return [];
    
    const resultados = new Set();
    resultados.add(limpar(todosDigitos));
    
    if (todosDigitos.length >= 13) {
        const semPrefixoOperadora = todosDigitos.replace(/^0(\d{2})/, "$1");
        resultados.add(limpar(semPrefixoOperadora));
    }
    
    const semZerosInicio = todosDigitos.replace(/^0+/, "");
    if (semZerosInicio.length >= 10) {
        resultados.add(semZerosInicio);
    }
    
    const matchDDDDup = todosDigitos.match(/^(\d{2})\1+(\d{8,9})$/);
    if (matchDDDDup) {
        const ddd = matchDDDDup[1];
        const numero = matchDDDDup[2];
        resultados.add(ddd + numero);
    }
    
    if (todosDigitos.startsWith("5555")) {
        resultados.add(limpar(todosDigitos.slice(2)));
    }
    
    if (todosDigitos.startsWith("55") && todosDigitos.length >= 12) {
        resultados.add(todosDigitos.slice(2));
    }
    
    if (!todosDigitos.startsWith("55") && todosDigitos.length >= 10) {
        resultados.add("55" + todosDigitos);
    }
    
    if (todosDigitos.length === 11) {
        const semNove = todosDigitos.slice(0, 2) + todosDigitos.slice(3);
        resultados.add(semNove);
    }
    if (todosDigitos.length === 10) {
        const comNove = todosDigitos.slice(0, 2) + "9" + todosDigitos.slice(2);
        resultados.add(comNove);
    }
    
    if (todosDigitos.length === 12 && todosDigitos.startsWith("55")) {
        const sem55 = todosDigitos.slice(2);
        resultados.add(sem55);
        const comNove = sem55.slice(0, 2) + "9" + sem55.slice(2);
        resultados.add("55" + comNove);
        const semNove = sem55.slice(0, 2) + sem55.slice(3);
        resultados.add("55" + semNove);
    }
    
    if (todosDigitos.length === 13 && todosDigitos.startsWith("55")) {
        const sem55 = todosDigitos.slice(2);
        resultados.add(sem55);
        const comNove = sem55.slice(0, 2) + "9" + sem55.slice(2);
        resultados.add("55" + comNove);
    }
    
    const matchDDDBaguncado = todosDigitos.match(/^0*(\d)(\d)\1*\2+(\d{8,9})$/);
    if (matchDDDBaguncado) {
        const ddd = matchDDDBaguncado[1] + matchDDDBaguncado[2];
        const numero = matchDDDBaguncado[3];
        resultados.add(ddd + numero);
        resultados.add("55" + ddd + numero);
    }
    
    if (todosDigitos.length > 13) {
        const ultimos11 = todosDigitos.slice(-11);
        const ultimos10 = todosDigitos.slice(-10);
        resultados.add(ultimos11);
        resultados.add(ultimos10);
        resultados.add("55" + ultimos11);
        resultados.add("55" + ultimos10);
    }
    
    return [...resultados].filter(n => n.length >= 8 && n.length <= 15);
}

function variantes(numero) {
    const n = limpar(numero);
    const vars = new Set([n]);
    
    if (n.startsWith("55") && n.length >= 12) vars.add(n.slice(2));
    if (n.length > 11) vars.add(n.slice(-11));
    if (n.length > 10) vars.add(n.slice(-10));
    if (n.length > 9) vars.add(n.slice(-9));
    if (n.length > 8) vars.add(n.slice(-8));
    
    if (n.length >= 12) {
        const semUmExtra = n.replace(/^(\d)\1+/, "$1");
        if (semUmExtra !== n) vars.add(semUmExtra);
        if (!n.startsWith("55")) {
            vars.add("55" + n);
        }
    }
    
    if (n.length === 11 && !n.startsWith("55")) {
        const semNove = n.slice(0, 2) + n.slice(3);
        vars.add(semNove);
    }
    if (n.length === 10 && !n.startsWith("55")) {
        const comNove = n.slice(0, 2) + "9" + n.slice(2);
        vars.add(comNove);
    }
    
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

function tsMinutosDataHora(celula) {
    if (celula instanceof Date && !isNaN(celula)) {
        return Math.floor(Date.UTC(
            celula.getFullYear(), celula.getMonth(), celula.getDate(),
            celula.getHours(), celula.getMinutes()
        ) / 60000);
    }

    const t = String(celula ?? "").trim();
    if (!t) return null;
    
    const m = t.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})[^\d]*(\d{1,2}):(\d{2})/);
    if (m) {
        const [, dd, mm, yy, h, mi] = m;
        const y = yy.length === 2 ? 2000 + Number(yy) : Number(yy);
        return Math.floor(Date.UTC(y, Number(mm) - 1, Number(dd), Number(h), Number(mi)) / 60000);
    }
    
    const d = new Date(t);
    if (!isNaN(d.getTime())) {
        return Math.floor(Date.UTC(
            d.getFullYear(), d.getMonth(), d.getDate(),
            d.getHours(), d.getMinutes()
        ) / 60000);
    }
    
    return null;
}

function tsMinutosDoArquivo(nome) {
    const m = String(nome).match(/(\d{4})(\d{2})(\d{2})[_\-](\d{2})(\d{2})(\d{2})/);
    if (!m) return null;
    const [, y, mo, d, h, mi, s] = m;
    return Math.floor(Date.UTC(Number(y), Number(mo) - 1, Number(d), Number(h), Number(mi), Number(s)) / 60000);
}

function infoAgente(texto) {
    const t = normalizarTexto(texto);
    if (!t) return null;
    const nums = sequenciasDe(t).filter((n) => n.length >= 3);
    const nomeSo = t.replace(/\d+/g, " ").replace(/\s+/g, " ").trim();
    if (!nomeSo && !nums.length) return null;
    return { t, numero: nums.length ? nums[nums.length - 1] : "", nomeSo };
}

function infoAgenteDoArquivo(nome) {
    const m = String(nome).match(/agente[\s_]+(.+?)[\s_]+(?:tel|fila)/i);
    if (!m) return null;
    
    const textoAgente = m[1];
    const partes = textoAgente.split(/[\s_]+/);
    
    let nomeLimpo = partes;
    if (partes.length > 1 && /^\d+$/.test(partes[partes.length - 1])) {
        nomeLimpo = partes.slice(0, -1);
    }
    
    return infoAgente(nomeLimpo.join(" "));
}

function agenteBate(a, b) {
    if (!a || !b) return true;
    return (
        (a.numero && b.numero && a.numero === b.numero) ||
        (a.nomeSo && b.nomeSo && (a.nomeSo.includes(b.nomeSo) || b.nomeSo.includes(a.nomeSo)))
    );
}

const PADRAO_TEL = /^(?:\+?55[\s.\-]*)?\(?\d{2,3}\)?[\s.\-]*\d{4,5}[\s.\-]*\d{4,5}$/;
const PADRAO_TEL_SEM_DDD = /^\d{4,5}[\s.\-]*\d{4,5}$/;

// NOVO: Extração inteligente de telefone do nome do arquivo
function extrairTelefoneDoNome(nome) {
    // Formato: ...Tel_(11) 9 9951-9737_Fila... ou ...Tel_(14) 9 9603-2075_Fila...
    const match = String(nome).match(/tel[_\s]+(\(?[\d\s\-\(\)]+\d{4,5}[\s\-]*\d{4,5})/i);
    if (match) {
        return match[1];
    }
    return null;
}

// NOVO: Extrai números de telefone do nome SEM "Tel_" (ignora timestamp e IDs curtos)
function extrairNumerosDoNomeSemTel(nome) {
    const nums = [];
    const sequencias = sequenciasDe(nome);
    
    for (const seq of sequencias) {
        // Ignora timestamp do início (8 dígitos tipo 20260930)
        if (seq.length === 8 && /^\d{4}(0[1-9]|1[0-2])(0[1-9]|[12]\d|3[01])$/.test(seq)) continue;
        // Ignora IDs curtos de agente (3-4 dígitos)
        if (seq.length >= 3 && seq.length <= 4) continue;
        // Pega apenas sequências que pareçam telefone (8-11 dígitos)
        if (seq.length >= 8 && seq.length <= 11) {
            nums.push(limpar(seq));
        }
    }
    
    return nums;
}

function extrairNumerosDe(texto) {
    const t = String(texto).replace(/\u00a0/g, " ").trim();
    const nums = [];
    if (!t) return nums;

    const normalizados = normalizarTelefone(t);
    nums.push(...normalizados);

    for (const seq of sequenciasDe(t)) {
        if (seq.length >= 8) {
            const limpo = limpar(seq);
            if (!nums.includes(limpo)) {
                nums.push(limpo);
            }
        }
    }

    return nums;
}

function ehColunaTelefone(cabecalho) {
    const t = String(cabecalho ?? "").trim().toUpperCase();
    return (
        t.startsWith("TEL") ||
        t.includes("TELEFONE") ||
        t.includes("CELULAR") ||
        t.includes("FONE") ||
        t.includes("CONTATO") ||
        t.includes("NÚMERO") ||
        t.includes("NUMERO") ||
        t.includes("PHONE") ||
        t.includes("MOBILE") ||
        t.includes("WHATSAPP") ||
        t.includes("ZAP") ||
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
    const t = String(cabecalho ?? "").toUpperCase();
    return t.includes("AGENTE");
}

function ehColunaNumeroAgente(cabecalho) {
    const t = String(cabecalho ?? "").toUpperCase();
    return t.includes("NÚMERO DO AGENTE") || t.includes("NUMERO DO AGENTE");
}

function ehColunaProtocolo(cabecalho) {
    const t = String(cabecalho ?? "").trim().toLowerCase();
    return t === "protocolo" || t.includes("protocolo");
}

async function lerPlanilha(file) {
    const buffer = await file.arrayBuffer();
    const workbook = XLSX.read(buffer, { type: "array", raw: true, cellDates: true });
    const mapaBase = new Map();
    const mapaProtocolo = new Map();
    const ocorrencias = new Map();
    const registrosData = [];
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
        let colunaNumeroAgente = -1;
        let colunaProtocolo = -1;

        console.log(` Cabeçalhos da aba "${nomeAba}":`, linhas[0]);

        linhas[0].forEach((c, i) => {
            if (ehColunaTelefone(c)) colunasTelefone.push(i);
            if (colunaData < 0 && ehColunaData(c)) colunaData = i;
            if (colunaAgente < 0 && ehColunaAgente(c) && !ehColunaNumeroAgente(c)) colunaAgente = i;
            if (colunaNumeroAgente < 0 && ehColunaNumeroAgente(c)) colunaNumeroAgente = i;
            if (colunaProtocolo < 0 && ehColunaProtocolo(c)) colunaProtocolo = i;
        });

        console.log(" Colunas encontradas:", {
            telefone: colunasTelefone,
            data: colunaData,
            agente: colunaAgente,
            numeroAgente: colunaNumeroAgente,
            protocolo: colunaProtocolo
        });

        const rows = [];

        linhas.forEach((linha, idxLinha) => {
            const sujeito = limpar(sequenciasDe(String(linha[0] ?? "")).join(""));

            if (colunaProtocolo >= 0) {
                const proto = textoDaCelula(linha[colunaProtocolo]).replace(/\D/g, "");
                if (proto && proto.length >= 4) {
                    const protoLimpo = limpar(proto);
                    if (!mapaProtocolo.has(protoLimpo)) {
                        mapaProtocolo.set(protoLimpo, sujeito);
                    }
                }
            }

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
            
            let agente = null;
            if (colunaAgente >= 0) {
                agente = infoAgente(linha[colunaAgente]);
                if (agente && colunaNumeroAgente >= 0) {
                    const numAgente = textoDaCelula(linha[colunaNumeroAgente]).replace(/\D/g, "");
                    if (numAgente) {
                        agente.numero = numAgente;
                    }
                }
            }

            if (idxLinha === 1 && colunaData >= 0) {
                console.log("🕐 DEBUG DATA:", {
                    celulaRaw: linha[colunaData],
                    celulaString: String(linha[colunaData]),
                    tsMinutos: ts,
                    dataObj: new Date(linha[colunaData])
                });
            }

            if (idxLinha === 1 && colunaAgente >= 0) {
                console.log("👤 DEBUG AGENTE:", {
                    nomeAgente: linha[colunaAgente],
                    numeroAgente: colunaNumeroAgente >= 0 ? linha[colunaNumeroAgente] : "N/A",
                    agenteObj: agente
                });
            }

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

    console.log("📋 Total registros data/hora:", registrosData.length);
    console.log(" Total agentes:", agentesPlanilha.length);
    console.log("📞 Total telefones:", mapaBase.size);

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
        mapaProtocolo,
        registrosData,
        agentesPlanilha,
        folhas,
        totalUnicos: mapaBase.size,
        totalRegistros,
        duplicados,
    };
}

async function baixarZipDe(arquivos, nomeZip, botao) {
    if (!arquivos.length) {
        alert("⚠️ Nenhum arquivo para baixar!");
        return;
    }

    try {
        if (statusTexto) statusTexto.textContent = `Gerando ${nomeZip} (modo rápido)...`;
        if (botao) botao.disabled = true;

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
        if (botao) botao.disabled = false;
        atualizarBotoes();
    }
}

async function baixarZipEmLotes(arquivos, nomeZip, botao, tamanhoLote = 50) {
    if (!arquivos.length) {
        alert("⚠️ Nenhum arquivo para baixar!");
        return;
    }

    try {
        if (statusTexto) statusTexto.textContent = `Gerando ${nomeZip} em lotes de ${tamanhoLote}...`;
        if (botao) botao.disabled = true;

        const zip = new JSZip();
        const nomesBaseUsados = new Set();

        for (let i = 0; i < arquivos.length; i++) {
            const loteIndex = Math.floor(i / tamanhoLote) + 1;
            const nomePasta = `lote_${String(loteIndex).padStart(2, '0')}`;
            const pastaLote = zip.folder(nomePasta);
            
            const file = arquivos[i];
            let nomeFinal = file.name;
            
            let contador = 1;
            while (nomesBaseUsados.has(nomeFinal)) {
                const ponto = file.name.lastIndexOf(".");
                nomeFinal = ponto > 0
                    ? `${file.name.slice(0, ponto)} (${contador})${file.name.slice(ponto)}`
                    : `${file.name} (${contador})`;
                contador++;
            }
            nomesBaseUsados.add(nomeFinal);
            
            pastaLote.file(nomeFinal, file);
        }

        const blob = await zip.generateAsync({ type: "blob", compression: "STORE" });
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = nomeZip;
        link.click();
        URL.revokeObjectURL(url);

        const totalLotes = Math.ceil(arquivos.length / tamanhoLote);
        if (statusTexto) {
            statusTexto.textContent = `✅ ${nomeZip} baixado: ${arquivos.length} arquivo(s) dividido(s) em ${totalLotes} pasta(s).`;
        }
    } catch (erro) {
        console.error(erro);
        if (statusTexto) statusTexto.textContent = "Erro ao gerar ZIP em lotes: " + erro.message;
    } finally {
        if (botao) botao.disabled = false;
        atualizarBotoes();
    }
}

function baixarEncontradosCSV() {
    if (!ultimoResultado.achados || !ultimoResultado.achados.length) {
        alert("⚠️ Nenhum arquivo encontrado para exportar!");
        return;
    }

    try {
        if (statusTexto) statusTexto.textContent = "Gerando CSV dos encontrados...";

        const cabecalho = [
            "Nome do Arquivo",
            "Regra de Match",
            "Identificador no Arquivo",
            "Identificador na Planilha",
            "Sujeito (ID da linha)"
        ];

        const linhas = [cabecalho.map(c => `"${c}"`).join(";")];

        for (const achado of ultimoResultado.achados) {
            const linha = [
                achado.nome || "",
                achado.regra || "",
                achado.numeroArquivo || "",
                achado.numeroPlanilha || "",
                achado.sujeito || ""
            ];

            const linhaCSV = linha.map(celula => {
                let str = String(celula ?? "");
                str = str.replace(/"/g, '""').replace(/[\n\r]+/g, ' ').trim();
                return `"${str}"`;
            }).join(";");

            linhas.push(linhaCSV);
        }

        const csvContent = linhas.join("\r\n");
        const blob = new Blob(["\ufeff" + csvContent], { type: "text/csv;charset=utf-8;" });
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = "arquivos_encontrados.csv";
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);

        if (statusTexto) {
            statusTexto.textContent = `✅ CSV dos encontrados baixado com ${ultimoResultado.achados.length} linha(s)!`;
        }
    } catch (erro) {
        console.error("❌ ERRO ao gerar CSV dos encontrados:", erro);
        alert("Erro ao gerar arquivo: " + erro.message);
    }
}

function linhaTemMatch(row) {
    if (linhasEncontradas.has(row.sujeito)) return true;
    for (const n of row.nums) {
        if (numerosEncontrados.has(n)) return true;
    }
    return false;
}

function baixarPlanilhaRestante() {
    if (!estruturaPlanilha) {
        console.error("❌ estruturaPlanilha é nula!");
        return;
    }

    try {
        if (statusTexto) statusTexto.textContent = "Gerando arquivo CSV...";

        let csvContent = "";
        let totalExportado = 0;

        for (const folha of estruturaPlanilha) {
            folha.rows.forEach((row, idx) => {
                if (!row || !row.aoa) return;

                const temDados = row.aoa.some(celula => {
                    if (celula === null || celula === undefined) return false;
                    return String(celula).trim() !== "";
                });

                if (!temDados) return;
                if (idx > 0 && linhaTemMatch(row)) return;

                const linhaCSV = row.aoa.map(celula => {
                    if (celula === null || celula === undefined) return '""';
                    
                    let str = String(celula);
                    str = str.replace(/"/g, '""').replace(/[\n\r]+/g, ' ').trim();
                    
                    return `"${str}"`;
                }).join(";");

                csvContent += linhaCSV + "\r\n";
                if (idx > 0) totalExportado++;
            });
        }

        if (totalExportado === 0) {
            alert("⚠️ Nenhuma linha restante para exportar!");
            if (statusTexto) statusTexto.textContent = "Nenhuma linha restante para exportar.";
            return;
        }

        const blob = new Blob(["\ufeff" + csvContent], { type: "text/csv;charset=utf-8;" });
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = "planilha_nao_encontrados.csv";
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);

        if (statusTexto) {
            statusTexto.textContent = `✅ CSV baixado com ${totalExportado} linha(s)!`;
        }
    } catch (erro) {
        console.error("❌ ERRO FATAL:", erro);
        alert("Erro ao gerar arquivo: " + erro.message);
    }
}

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

let btnEncontradosPorData = document.getElementById("btnEncontradosPorData");
if (!btnEncontradosPorData) {
    const containerAcoes =
        document.querySelector(".acoesSeparar") ||
        btnProcessar?.parentElement ||
        document.body;

    btnEncontradosPorData = document.createElement("button");
    btnEncontradosPorData.id = "btnEncontradosPorData";
    btnEncontradosPorData.className = "botaoTerciario";
    btnEncontradosPorData.innerHTML = '<i class="fa-solid fa-clock"></i> Baixar encontrados por data/hora (.zip)';
    btnEncontradosPorData.disabled = true;
    containerAcoes.appendChild(btnEncontradosPorData);
    btnEncontradosPorData.addEventListener("click", () =>
        baixarZipDe(arquivosEncontradosPorData, "encontrados_por_data_hora.zip", btnEncontradosPorData)
    );
}

let btnEncontradosCSV = document.getElementById("btnEncontradosCSV");
if (!btnEncontradosCSV) {
    const containerAcoes =
        document.querySelector(".acoesSeparar") ||
        btnProcessar?.parentElement ||
        document.body;

    btnEncontradosCSV = document.createElement("button");
    btnEncontradosCSV.id = "btnEncontradosCSV";
    btnEncontradosCSV.className = "botaoTerciario";
    btnEncontradosCSV.innerHTML = '<i class="fa-solid fa-file-csv"></i> Baixar encontrados (.csv)';
    btnEncontradosCSV.disabled = true;
    containerAcoes.appendChild(btnEncontradosCSV);
    btnEncontradosCSV.addEventListener("click", baixarEncontradosCSV);
}

let chkModoMatch = document.getElementById("chkModoMatch");
if (!chkModoMatch) {
    const containerAcoes = document.querySelector(".acoesSeparar") || btnProcessar?.parentElement || document.body;
    
    const wrapper = document.createElement("div");
    wrapper.style.cssText = "margin-bottom: 12px; display: flex; align-items: center; gap: 8px; font-family: sans-serif; font-size: 0.9rem; color: #333;";
    
    chkModoMatch = document.createElement("input");
    chkModoMatch.type = "checkbox";
    chkModoMatch.id = "chkModoMatch";
    chkModoMatch.style.cursor = "pointer";
    chkModoMatch.style.width = "18px";
    chkModoMatch.style.height = "18px";
    
    const label = document.createElement("label");
    label.htmlFor = "chkModoMatch";
    label.textContent = "Forçar busca por Data/Hora e Agente (ignora Protocolo e Telefone)";
    label.style.cursor = "pointer";
    label.style.fontWeight = "500";
    
    wrapper.appendChild(chkModoMatch);
    wrapper.appendChild(label);
    
    if (btnProcessar) {
        containerAcoes.insertBefore(wrapper, btnProcessar);
    } else {
        containerAcoes.appendChild(wrapper);
    }
}

function atualizarBotoes() {
    if (btnBaixarZip) btnBaixarZip.disabled = arquivosEncontrados.length === 0;
    if (btnEncontradosCSV) btnEncontradosCSV.disabled = !ultimoResultado.achados || ultimoResultado.achados.length === 0;
    if (btnNaoEncontrados) btnNaoEncontrados.disabled = linhasRestantes === 0;
    if (btnFora) btnFora.disabled = arquivosForaDaPlanilha.length === 0;
    if (btnEncontradosPorData) btnEncontradosPorData.disabled = arquivosEncontradosPorData.length === 0;
}

atualizarBotoes();

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
        `<strong> Conferência:</strong><br>` +
        `Exemplos de números lidos das colunas de telefone: ${amostraPlanilha.join(", ") || "— (nessa planilha o match foi por data/hora ou agente)"}` +
        (duplicados.length
            ? `<br><br><strong>️ Números duplicados (${duplicados.length}):</strong><br>` +
            duplicados.slice(0, 10)
                .map((d) => `${d.numero} → aparece em ${d.ids.length} linhas: ${d.ids.join(", ")}`)
                .join("<br>")
            : "") +
        (nao.length
            ? `<br><br><strong>Arquivos da pasta que NÃO estão na planilha (${nao.length}):</strong><br>` + fora.join("<br>") +
            (nao.length > 8 ? "<br>… (baixe o ZIP 'fora da planilha' pra levar todos)" : "")
            : "<br><br>Todos os arquivos da pasta bateram com a planilha. ✅");
}

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
    arquivosEncontradosPorData = [];
    arquivosForaDaPlanilha = [];
    numerosEncontrados = new Set();
    linhasEncontradas = new Set();
    linhasRestantes = 0;
    
    let linhasUtilizadas = new Set();
    
    atualizarBotoes();
    if (statusTexto) statusTexto.textContent = "Lendo planilha e processando...";

    try {
        const arquivos = Array.from(inputPasta.files || []);
        const documento = inputDocumento.files?.[0];

        if (!arquivos.length) throw new Error("Selecione a pasta com os arquivos.");
        if (!documento) throw new Error("Selecione o Excel ou CSV com os números.");

        const forcarDataEAgente = chkModoMatch?.checked || false;

        const { mapaVariante, mapaProtocolo, registrosData, agentesPlanilha, folhas, totalUnicos, totalRegistros, duplicados } =
            await lerPlanilha(documento);

        if (!mapaVariante.size && !mapaProtocolo.size && !registrosData.length && !agentesPlanilha.length) {
            throw new Error("Planilha sem telefone, sem protocolo, sem data e sem agente — nada pra comparar.");
        }

        estruturaPlanilha = folhas;

        const achados = [];
        const nao = [];
        let cProto = 0, cTel = 0, cData = 0, cAgente = 0;

        for (let idx = 0; idx < arquivos.length; idx++) {
            const f = arquivos[idx];
            
            // NOVO: Lógica inteligente de extração
            const telDoNome = extrairTelefoneDoNome(f.name);
            let numsArquivo = [];
            let temTelefoneNoNome = false;
            
            if (telDoNome) {
                // Tem "Tel_" no nome - extrai só o telefone
                numsArquivo = extrairNumerosDe(telDoNome);
                temTelefoneNoNome = true;
            } else {
                // NÃO tem "Tel_" - extrai números que pareçam telefone (ignora timestamp e IDs)
                numsArquivo = extrairNumerosDoNomeSemTel(f.name);
                temTelefoneNoNome = false;
            }
            
            const protoArquivo = String(f.name).match(/(\d+)/)?.[1] || null;
            const tsArq = tsMinutosDoArquivo(f.name);
            const agArq = infoAgenteDoArquivo(f.name);
            
            if (idx < 5) {
                console.log(`\n🔍 ARQUIVO ${idx + 1}:`, f.name);
                console.log("  - Tem Tel_ no nome?", temTelefoneNoNome);
                console.log("  - Telefone extraído:", telDoNome);
                console.log("  - Números para match:", numsArquivo);
                console.log("  - Timestamp arquivo:", tsArq, tsArq ? new Date(tsArq * 60000).toISOString() : "");
                console.log("  - Agente arquivo:", agArq);
            }
            
            let evidencia = null;
            let regra = "";

            if (!forcarDataEAgente) {
                if (!evidencia && mapaProtocolo.size > 0 && protoArquivo) {
                    const hit = mapaProtocolo.get(protoArquivo);
                    if (hit && !linhasUtilizadas.has(hit)) {
                        evidencia = { numeroArquivo: protoArquivo, numeroPlanilha: protoArquivo, sujeito: hit };
                        regra = "protocolo";
                    }
                }

                if (!evidencia && temTelefoneNoNome) {
                    for (const n of numsArquivo) {
                        for (const v of variantes(n)) {
                            const hit = mapaVariante.get(v);
                            if (hit && !linhasUtilizadas.has(hit.sujeito)) {
                                evidencia = { numeroArquivo: n, numeroPlanilha: hit.numero, sujeito: hit.sujeito };
                                regra = "tel";
                                break;
                            }
                        }
                        if (evidencia) break;
                    }
                }
            }

            if (!evidencia && tsArq !== null) {
                let melhorHit = null;
                let melhorDiff = Infinity;
                
                for (const r of registrosData) {
                    if (linhasUtilizadas.has(r.sujeito)) continue;
                    
                    const diff = Math.abs(r.ts - tsArq);
                    if (diff > TOLERANCIA_MIN) continue;
                    
                    // Se tem agente no arquivo, confirma com agente. Se não, aceita qualquer um.
                    if (agArq && !agenteBate(r.agente, agArq)) continue;

                    if (diff < melhorDiff) {
                        melhorDiff = diff;
                        melhorHit = r;
                    }
                }
                
                if (melhorHit) {
                    evidencia = { numeroArquivo: String(tsArq), numeroPlanilha: String(melhorHit.ts), sujeito: melhorHit.sujeito };
                    regra = "data";
                }
            }

            if (!evidencia && agArq) {
                for (const a of agentesPlanilha) {
                    if (linhasUtilizadas.has(a.sujeito)) continue;
                    
                    if (agenteBate(a, agArq) && (a.nomeSo || a.numero)) {
                        evidencia = { numeroArquivo: agArq.t, numeroPlanilha: a.t, sujeito: a.sujeito };
                        regra = "agente";
                        break;
                    }
                }
            }

            if (evidencia) {
                linhasUtilizadas.add(evidencia.sujeito);
                
                if (regra === "protocolo") cProto++;
                if (regra === "tel") { cTel++; numerosEncontrados.add(evidencia.numeroPlanilha); }
                if (regra === "data") {
                    cData++;
                    arquivosEncontradosPorData.push(f);
                }
                if (regra === "agente") cAgente++;
                linhasEncontradas.add(evidencia.sujeito);
                achados.push({ file: f, nome: f.webkitRelativePath || f.name, regra, ...evidencia });
                
                if (idx < 3) {
                    console.log("  ✅ ENCONTRADO por:", regra);
                }
            } else {
                nao.push({ file: f, nome: f.webkitRelativePath || f.name, numeros: numsArquivo });
                if (idx < 3) {
                    console.log("  ❌ NÃO ENCONTRADO");
                }
            }
        }

        linhasRestantes = 0;
        for (const folha of estruturaPlanilha) {
            folha.rows.forEach((row, idx) => {
                if (idx > 0) {
                    if (!row || !row.aoa) return;
                    
                    const temDados = row.aoa.some(celula => {
                        if (celula === null || celula === undefined) return false;
                        return String(celula).trim() !== "";
                    });
                    
                    if (temDados && !linhaTemMatch(row)) linhasRestantes++;
                }
            });
        }

        ultimoResultado = { achados, nao, duplicados };
        arquivosEncontrados = achados.map((a) => a.file);
        arquivosForaDaPlanilha = nao.map((n) => n.file);

        if (statusTexto) {
            const modoTexto = forcarDataEAgente ? "🕒 MODO: Data/Hora e Agente" : "📞 MODO: Protocolo e Telefone";
            statusTexto.textContent =
                `Concluído (${modoTexto}). Registros na planilha: ${totalRegistros}. ` +
                `Encontrados: ${achados.length} (protocolo: ${cProto} • telefone: ${cTel} • data/hora ±${TOLERANCIA_MIN}min: ${cData} • agente: ${cAgente}) • ` +
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
    baixarZipEmLotes(arquivosEncontrados, "arquivos_encontrados.zip", btnBaixarZip, 50)
);

btnNaoEncontrados?.addEventListener("click", baixarPlanilhaRestante);