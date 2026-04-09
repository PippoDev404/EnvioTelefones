import { requireAuth, logout } from "../8-src/auth.js";

const user = await requireAuth({ redirectTo: "../7-login/login.html" });
if (!user) throw new Error("Sem sessão");

document.querySelector("#btnLogout")?.addEventListener("click", () => logout());

const baseFileInput = document.querySelector("#baseFile");
const cotaFileInput = document.querySelector("#cotaFile");
const hasCotaSelect = document.querySelector("#hasCota");
const maxNumbersPerPesqInput = document.querySelector("#maxNumbersPerPesq");
const processBtn = document.querySelector("#processBtn");
const downloadBtn = document.querySelector("#downloadBtn");

const manualDaysSection = document.querySelector("#manualDaysSection");
const manualDaysList = document.querySelector("#manualDaysList");
const addDayBtn = document.querySelector("#addDayBtn");

const logBox = document.querySelector("#logBox");
const resultSection = document.querySelector("#resultSection");
const previewTable = document.querySelector("#previewTable");
const previewHead = previewTable.querySelector("thead");
const previewBody = previewTable.querySelector("tbody");

let manualDayCounter = 0;
let generatedBlob = null;
let generatedFileName = "";

const CATEGORY_PRIORITY = ["PESQUISA", "VIVO", "CLARO", "OI", "TIM", "BRASIL"];
const SOBRA_LABEL = "SOBRA";

hasCotaSelect?.addEventListener("change", syncModeUI);
addDayBtn?.addEventListener("click", () => createManualDayCard());
processBtn?.addEventListener("click", handleProcess);
downloadBtn?.addEventListener("click", handleDownload);

syncModeUI();

/* =========================
   UI
========================= */

function syncModeUI() {
    const hasCota = hasCotaSelect?.value === "sim";

    document.querySelectorAll(".cota-only").forEach((el) => {
        el.classList.toggle("hidden", !hasCota);
    });

    document.querySelectorAll(".no-cota-only").forEach((el) => {
        el.classList.toggle("hidden", hasCota);
    });

    if (manualDaysSection) {
        manualDaysSection.classList.toggle("hidden", hasCota);
    }

    if (!hasCota && manualDaysList && !manualDaysList.children.length) {
        createManualDayCard();
    }
}

function createManualDayCard(data = {}) {
    if (!manualDaysList) return;

    manualDayCounter += 1;
    const card = document.createElement("div");
    card.className = "day-card";
    card.dataset.dayId = String(manualDayCounter);

    card.innerHTML = `
        <div class="day-card-header">
            <div class="day-card-title">Dia de pesquisa ${manualDayCounter}</div>
            <button type="button" class="danger-btn remove-day-btn">Remover</button>
        </div>

        <div class="day-grid">
            <div class="field">
                <label>Quantidade de pesquisadores</label>
                <input
                    type="number"
                    class="manual-p-count"
                    min="1"
                    step="1"
                    placeholder="Ex.: 36"
                    value="${data.pCount ?? ""}"
                />
            </div>

            <div class="field">
                <label>Data do dia de pesquisa</label>
                <input
                    type="text"
                    class="manual-date"
                    placeholder="Ex.: 12/04/2026"
                    value="${data.date ?? ""}"
                />
            </div>
        </div>

        <div class="day-help">
            Em qualquer modo, os telefones são sempre distribuídos por categoria em rodízio:
            PESQUISA, VIVO, CLARO, OI, TIM, BRASIL.
        </div>
    `;

    card.querySelector(".remove-day-btn")?.addEventListener("click", () => {
        card.remove();
        refreshManualDayTitles();
        if (!manualDaysList.children.length) {
            createManualDayCard();
        }
    });

    manualDaysList.appendChild(card);
    refreshManualDayTitles();
}

function refreshManualDayTitles() {
    if (!manualDaysList) return;

    [...manualDaysList.children].forEach((card, index) => {
        const title = card.querySelector(".day-card-title");
        if (title) title.textContent = `Dia de pesquisa ${index + 1}`;
    });
}

/* =========================
   UTILITÁRIOS
========================= */

function showLog() {
    if (logBox) logBox.classList.remove("hidden");
}

function addLog(message, type = "ok") {
    if (!logBox) return;

    showLog();

    let ul = logBox.querySelector("ul");
    if (!ul) {
        ul = document.createElement("ul");
        logBox.innerHTML = "";
        logBox.appendChild(ul);
    }

    const li = document.createElement("li");
    li.className = type;
    li.textContent = message;
    ul.appendChild(li);

    logBox.scrollTop = logBox.scrollHeight;
}

function clearLog() {
    if (!logBox) return;
    logBox.innerHTML = "";
    logBox.classList.add("hidden");
}

function normalizeText(value) {
    return String(value ?? "")
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "")
        .replace(/\s+/g, " ")
        .trim()
        .toUpperCase();
}

function isRowEmpty(row = []) {
    return row.every((cell) => String(cell ?? "").trim() === "");
}

function excelDateToJSDate(serial) {
    if (typeof serial !== "number") return null;
    const utcDays = Math.floor(serial - 25569);
    const utcValue = utcDays * 86400;
    return new Date(utcValue * 1000);
}

function formatDateToDDMMYYYY(date) {
    if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";
    const dd = String(date.getDate()).padStart(2, "0");
    const mm = String(date.getMonth() + 1).padStart(2, "0");
    const yyyy = date.getFullYear();
    return `${dd}/${mm}/${yyyy}`;
}

function normalizeDateString(value) {
    const raw = String(value ?? "").trim();
    if (!raw) return "";

    const direct = raw.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
    if (direct) return raw;

    const iso = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/);
    if (iso) return `${iso[3]}/${iso[2]}/${iso[1]}`;

    return raw;
}

function safeFileName(name) {
    return String(name || "arquivo_final")
        .replace(/\.[^.]+$/, "")
        .replace(/[\\/:*?"<>|]+/g, "_")
        .replace(/\s+/g, "_");
}

function triggerDownload(blob, fileName) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
}

function getArrayBuffer(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

function decodeCsvArrayBuffer(arrayBuffer) {
    try {
        return new TextDecoder("utf-8", { fatal: true }).decode(arrayBuffer);
    } catch {
        return new TextDecoder("latin1").decode(arrayBuffer);
    }
}

function parseCSV(text) {
    const parsed = Papa.parse(text, {
        header: true,
        skipEmptyLines: true,
        delimiter: "",
        dynamicTyping: false,
        transformHeader: (header) => String(header ?? "").trim(),
    });

    return parsed.data || [];
}

function findMainSheetName(workbook) {
    const names = workbook.SheetNames || [];
    if (!names.length) return null;

    const telefones = names.find(
        (name) => normalizeText(name) === "TELEFONES"
    );
    if (telefones) return telefones;

    const firstNotContagem = names.find(
        (name) => normalizeText(name) !== "CONTAGEM"
    );

    return firstNotContagem || names[0];
}

function getSheetRows(workbook, sheetName) {
    const sheet = workbook.Sheets[sheetName];
    return XLSX.utils.sheet_to_json(sheet, {
        header: 1,
        defval: "",
        raw: true,
        blankrows: false,
    });
}

function findHeaderIndex(headers, candidates) {
    const normalizedHeaders = headers.map(normalizeText);

    for (const candidate of candidates) {
        const idx = normalizedHeaders.findIndex((h) => h === normalizeText(candidate));
        if (idx !== -1) return idx;
    }

    return -1;
}

function getDisplayDate(rawValue) {
    if (typeof rawValue === "number") {
        return formatDateToDDMMYYYY(excelDateToJSDate(rawValue));
    }
    return normalizeDateString(rawValue);
}

function clearColumnsInRows(rows, dataStartIndex, targetIndexes = []) {
    for (let i = dataStartIndex; i < rows.length; i++) {
        if (!rows[i]) rows[i] = [];
        if (isRowEmpty(rows[i])) continue;

        for (const colIndex of targetIndexes) {
            if (colIndex >= 0) rows[i][colIndex] = "";
        }
    }
}

function ensureRequiredColumn(index, name) {
    return index !== -1 ? true : (() => {
        throw new Error(`Não encontrei a coluna "${name}".`);
    })();
}

function getMaxNumbersPerPesq() {
    const value = Number(maxNumbersPerPesqInput?.value);
    if (!Number.isInteger(value) || value <= 0) {
        throw new Error("Informe a quantidade máxima de telefones por P.");
    }
    return value;
}

function getManualDayConfigs() {
    if (!manualDaysList) return [];

    const cards = [...manualDaysList.querySelectorAll(".day-card")];

    return cards
        .map((card) => {
            const pCount = Number(card.querySelector(".manual-p-count")?.value);
            const date = normalizeDateString(card.querySelector(".manual-date")?.value);

            return { pCount, date };
        })
        .filter((item) => item.pCount > 0 && item.date);
}

function createPesquisadoresFromCount(count) {
    const qtd = Number(count);
    if (!Number.isInteger(qtd) || qtd <= 0) return [];

    return Array.from({ length: qtd }, (_, index) => {
        const n = String(index + 1).padStart(2, "0");
        return `P${n}`;
    });
}

function normalizeCategory(value) {
    return normalizeText(value);
}

function getCategoryPriorityIndex(category) {
    const idx = CATEGORY_PRIORITY.indexOf(normalizeCategory(category));
    return idx === -1 ? 999 : idx;
}

function sortRowsByCategory(rows) {
    return [...rows].sort((a, b) => {
        const diff = getCategoryPriorityIndex(a.categoria) - getCategoryPriorityIndex(b.categoria);
        if (diff !== 0) return diff;
        return a.originalIndex - b.originalIndex;
    });
}

function getPesqSortOrder(pesqValue) {
    const normalized = String(pesqValue ?? "").trim().toUpperCase();

    if (normalized === SOBRA_LABEL) return Number.MAX_SAFE_INTEGER - 1;
    if (!normalized) return Number.MAX_SAFE_INTEGER;

    const match = normalized.match(/^P(\d+)$/);
    if (match) return Number(match[1]);

    return Number.MAX_SAFE_INTEGER - 2;
}

function sortFinalRows(rows) {
    return [...rows].sort((a, b) => {
        const pesqDiff = getPesqSortOrder(a[8]) - getPesqSortOrder(b[8]);
        if (pesqDiff !== 0) return pesqDiff;

        const aDate = String(a[9] ?? "");
        const bDate = String(b[9] ?? "");
        if (aDate !== bDate) return aDate.localeCompare(bDate, "pt-BR");

        const aCategoryDiff = getCategoryPriorityIndex(a[5]) - getCategoryPriorityIndex(b[5]);
        if (aCategoryDiff !== 0) return aCategoryDiff;

        const aIdp = String(a[0] ?? "");
        const bIdp = String(b[0] ?? "");
        return aIdp.localeCompare(bIdp, "pt-BR", { numeric: true, sensitivity: "base" });
    });
}

function buildFinalRows(rows, dataStartIndex, columnIndexes, diaPesqColIndex) {
    const finalRows = [];

    for (let i = dataStartIndex; i < rows.length; i++) {
        const row = rows[i] || [];
        if (isRowEmpty(row)) continue;

        const pesqValue = String(row[columnIndexes.pesq] ?? "").trim();
        if (!pesqValue) continue;

        const diaPesqValue =
            diaPesqColIndex >= 0 ? getDisplayDate(row[diaPesqColIndex]) : "";

        finalRows.push([
            row[columnIndexes.idp] ?? "",
            row[columnIndexes.cidade] ?? "",
            row[columnIndexes.estado] ?? "",
            row[columnIndexes.regiao] ?? "",
            row[columnIndexes.setor] ?? "",
            row[columnIndexes.categoria] ?? "",
            row[columnIndexes.tf1] ?? "",
            row[columnIndexes.tf2] ?? "",
            pesqValue,
            diaPesqValue ?? "",
        ]);
    }

    return sortFinalRows(finalRows);
}

function autoFitSheetColumns(sheetData, worksheet) {
    const widths = [];

    sheetData.forEach((row) => {
        row.forEach((cell, index) => {
            const len = String(cell ?? "").length;
            widths[index] = Math.max(widths[index] || 10, Math.min(len + 2, 35));
        });
    });

    worksheet["!cols"] = widths.map((wch) => ({ wch }));
}

function handleDownload() {
    if (!generatedBlob || !generatedFileName) {
        addLog("Nenhum arquivo final foi gerado ainda.", "warn");
        return;
    }
    triggerDownload(generatedBlob, generatedFileName);
}

function assignRowsToPesquisador(rowsToAssign, pesquisador, date, rows, pesqCol, dataPesquisaCol) {
    rowsToAssign.forEach((item) => {
        const matrixIndex = item.rowNumberExcel - 1;
        if (!rows[matrixIndex]) rows[matrixIndex] = [];
        rows[matrixIndex][pesqCol] = pesquisador;
        rows[matrixIndex][dataPesquisaCol] = date;
    });
}

function computeMaxPerPByAvailable(totalRows, totalPesquisadores) {
    if (!totalPesquisadores) return 0;
    return Math.floor(totalRows / totalPesquisadores);
}

function isNumericLike(value) {
    const raw = String(value ?? "").trim();
    return /^-?\d+(?:[.,]\d+)?$/.test(raw);
}

function normalizeSetorValue(value) {
    const raw = String(value ?? "").trim();
    if (!raw) return "";
    return raw.replace(",", ".");
}

function normalizeSetorKey(value) {
    const raw = String(value ?? "").trim();
    if (!raw) return "";
    return isNumericLike(raw) ? normalizeSetorValue(raw) : normalizeText(raw);
}

function detectSetorMode(setorValue) {
    return isNumericLike(setorValue) ? "setor" : "estado";
}

function buildCotaDemandListForDate(cotaRowsForDate) {
    const demands = [];
    const headers = Object.keys(cotaRowsForDate[0] || {});
    const setorCols = headers.filter((h) => /^SETORES\d+$/i.test(String(h).trim()));

    cotaRowsForDate.forEach((row) => {
        const pesquisador = String(row["PESQ."] ?? row["PESQ"] ?? "").trim();
        if (!pesquisador) return;

        setorCols.forEach((_, index) => {
            const setorCol = `SETORES${index + 1}`;
            const qtCol = `QT${index + 1}`;

            const setorRaw = String(row[setorCol] ?? "").trim();
            const questionarios = Number(String(row[qtCol] ?? "").replace(",", "."));

            if (!setorRaw || !Number.isFinite(questionarios) || questionarios <= 0) return;

            demands.push({
                pesquisador,
                setorRaw,
                setorMode: detectSetorMode(setorRaw),
                setorKey: normalizeSetorKey(setorRaw),
                questionarios
            });
        });
    });

    return demands;
}

function splitCotaRowsByDate(cotaRows) {
    const groups = new Map();

    cotaRows.forEach((row) => {
        const date = normalizeDateString(row["DIA"]);
        if (!date) return;

        if (!groups.has(date)) groups.set(date, []);
        groups.get(date).push(row);
    });

    return groups;
}

function buildPoolsByMode(usefulRows) {
    const estadoPools = new Map();
    const setorPools = new Map();

    usefulRows.forEach((row) => {
        const estadoKey = normalizeText(row.estado);
        const setorKey = normalizeSetorValue(row.setor);

        if (!estadoPools.has(estadoKey)) estadoPools.set(estadoKey, []);
        estadoPools.get(estadoKey).push(row);

        if (!setorPools.has(setorKey)) setorPools.set(setorKey, []);
        setorPools.get(setorKey).push(row);
    });

    for (const [key, value] of estadoPools.entries()) {
        estadoPools.set(key, sortRowsByCategory(value));
    }

    for (const [key, value] of setorPools.entries()) {
        setorPools.set(key, sortRowsByCategory(value));
    }

    return { estadoPools, setorPools };
}

/* =========================
   DISTRIBUIÇÃO POR CATEGORIA
========================= */

function buildOrderedCategoryGroups(rows) {
    const grouped = new Map();
    const others = new Map();

    rows.forEach((item) => {
        const normalized = normalizeCategory(item.categoria);

        if (CATEGORY_PRIORITY.includes(normalized)) {
            if (!grouped.has(normalized)) grouped.set(normalized, []);
            grouped.get(normalized).push(item);
        } else {
            if (!others.has(normalized)) others.set(normalized, []);
            others.get(normalized).push(item);
        }
    });

    const orderedGroups = [];

    CATEGORY_PRIORITY.forEach((category) => {
        const arr = grouped.get(category) || [];
        if (arr.length) {
            orderedGroups.push({
                category,
                rows: [...arr].sort((a, b) => a.originalIndex - b.originalIndex)
            });
        }
    });

    [...others.entries()]
        .sort((a, b) => a[0].localeCompare(b[0]))
        .forEach(([category, arr]) => {
            orderedGroups.push({
                category,
                rows: [...arr].sort((a, b) => a.originalIndex - b.originalIndex)
            });
        });

    return orderedGroups;
}

function distributeRowsByCategoryRoundRobin(rowsPool, targets) {
    const categoryGroups = buildOrderedCategoryGroups(rowsPool);
    let totalAssigned = 0;

    for (let groupIndex = 0; groupIndex < categoryGroups.length; groupIndex++) {
        const group = categoryGroups[groupIndex];
        let cursor = 0;

        while (cursor < group.rows.length) {
            const targetsNeedingRows = targets.filter((target) => target.remaining > 0);

            if (!targetsNeedingRows.length) {
                const remainingPool = [];

                for (let i = groupIndex; i < categoryGroups.length; i++) {
                    const currentGroup = categoryGroups[i];
                    if (i === groupIndex) {
                        remainingPool.push(...currentGroup.rows.slice(cursor));
                    } else {
                        remainingPool.push(...currentGroup.rows);
                    }
                }

                return { totalAssigned, remainingPool };
            }

            let assignedInThisRound = false;

            for (const target of targetsNeedingRows) {
                if (cursor >= group.rows.length) break;
                if (target.remaining <= 0) continue;

                const rowItem = group.rows[cursor++];
                target.rows.push(rowItem);
                target.remaining -= 1;
                totalAssigned += 1;
                assignedInThisRound = true;
            }

            if (!assignedInThisRound) break;
        }
    }

    const assignedSet = new Set(targets.flatMap((t) => t.rows));
    const remainingPool = rowsPool.filter((row) => !assignedSet.has(row));

    return { totalAssigned, remainingPool };
}

function assignTargetsToSheet(targets, rows, pesqCol, dataPesquisaCol) {
    let total = 0;

    targets.forEach((target) => {
        target.rows.forEach((item) => {
            const matrixIndex = item.rowNumberExcel - 1;
            if (!rows[matrixIndex]) rows[matrixIndex] = [];
            rows[matrixIndex][pesqCol] = target.pesquisador;
            rows[matrixIndex][dataPesquisaCol] = target.date;
            total += 1;
        });
    });

    return total;
}

function markUnusedRowsAsSobra(usefulRows, rows, pesqCol, dataPesquisaCol) {
    let sobraCount = 0;

    usefulRows.forEach((item) => {
        const matrixIndex = item.rowNumberExcel - 1;
        if (!rows[matrixIndex]) rows[matrixIndex] = [];

        const currentPesq = String(rows[matrixIndex][pesqCol] ?? "").trim();

        if (!currentPesq) {
            rows[matrixIndex][pesqCol] = SOBRA_LABEL;
            rows[matrixIndex][dataPesquisaCol] = "";
            sobraCount += 1;
        }
    });

    return sobraCount;
}

/* =========================
   RATEIO POR DATA NO MODO COTA
========================= */

function groupDemandsByBucketAcrossDates(cotaRowsByDate) {
    const bucketMap = new Map();

    for (const [date, cotaRowsForDate] of cotaRowsByDate.entries()) {
        const demands = buildCotaDemandListForDate(cotaRowsForDate);

        demands.forEach((demand) => {
            const bucketKey = `${demand.setorMode}::${demand.setorKey}`;

            if (!bucketMap.has(bucketKey)) {
                bucketMap.set(bucketKey, {
                    mode: demand.setorMode,
                    key: demand.setorKey,
                    dates: new Map()
                });
            }

            const bucket = bucketMap.get(bucketKey);

            if (!bucket.dates.has(date)) {
                bucket.dates.set(date, []);
            }

            bucket.dates.get(date).push(demand);
        });
    }

    return bucketMap;
}

function splitIntegerEqually(total, parts) {
    if (parts <= 0) return [];
    const base = Math.floor(total / parts);
    const remainder = total % parts;

    return Array.from({ length: parts }, (_, index) => base + (index < remainder ? 1 : 0));
}

function buildSharedPoolsForCota(usefulRows, cotaRowsByDate) {
    const bucketGroups = groupDemandsByBucketAcrossDates(cotaRowsByDate);
    const { estadoPools, setorPools } = buildPoolsByMode(usefulRows);

    const sharedPools = new Map();

    for (const [bucketKey, bucketInfo] of bucketGroups.entries()) {
        const sourcePool =
            bucketInfo.mode === "estado"
                ? (estadoPools.get(bucketInfo.key) || [])
                : (setorPools.get(bucketInfo.key) || []);

        const dates = [...bucketInfo.dates.keys()].sort((a, b) => {
            const [da, ma, ya] = a.split("/").map(Number);
            const [db, mb, yb] = b.split("/").map(Number);
            return new Date(ya, ma - 1, da) - new Date(yb, mb - 1, db);
        });

        const equalParts = splitIntegerEqually(sourcePool.length, dates.length);

        let cursor = 0;
        const perDatePools = new Map();

        dates.forEach((date, index) => {
            const qty = equalParts[index];
            const slice = sourcePool.slice(cursor, cursor + qty);
            cursor += qty;
            perDatePools.set(date, slice);
        });

        sharedPools.set(bucketKey, {
            mode: bucketInfo.mode,
            key: bucketInfo.key,
            perDatePools
        });
    }

    return sharedPools;
}

/* =========================
   PROCESSAMENTO PRINCIPAL
========================= */

async function handleProcess() {
    try {
        clearLog();
        if (resultSection) resultSection.classList.add("hidden");
        if (previewHead) previewHead.innerHTML = "";
        if (previewBody) previewBody.innerHTML = "";
        if (downloadBtn) downloadBtn.classList.add("hidden");
        generatedBlob = null;
        generatedFileName = "";

        const baseFile = baseFileInput?.files?.[0];
        const hasCota = hasCotaSelect?.value === "sim";

        if (!baseFile) {
            addLog("Selecione o arquivo base (.xlsx).", "error");
            return;
        }

        if (hasCota && !cotaFileInput?.files?.[0]) {
            addLog("Selecione o arquivo COTA (.csv).", "error");
            return;
        }

        if (!hasCota) {
            const manualConfigs = getManualDayConfigs();
            if (!manualConfigs.length) {
                addLog("Informe ao menos um dia de pesquisa com quantidade de pesquisadores e data.", "error");
                return;
            }
        }

        processBtn.disabled = true;
        processBtn.textContent = "Processando...";

        addLog("Lendo arquivo Excel...", "ok");
        const baseBuffer = await getArrayBuffer(baseFile);
        const workbook = XLSX.read(baseBuffer, { type: "array", cellDates: false });

        const mainSheetName = findMainSheetName(workbook);
        if (!mainSheetName) {
            addLog("Não foi possível localizar a aba principal.", "error");
            return;
        }

        const rows = getSheetRows(workbook, mainSheetName);
        if (!rows.length || rows.length < 3) {
            addLog("A aba principal está vazia ou incompleta.", "error");
            return;
        }

        addLog(`Aba principal encontrada: ${mainSheetName}`, "ok");

        const headerRowIndex = 1;
        const headerRow = rows[headerRowIndex] || [];
        const dataStartIndex = headerRowIndex + 1;
        const dataRows = rows.slice(dataStartIndex);

        const idpCol = findHeaderIndex(headerRow, ["IDP"]);
        const cidadeCol = findHeaderIndex(headerRow, ["CIDADE"]);
        const estadoCol = findHeaderIndex(headerRow, ["ESTADO"]);
        const regiaoCol = findHeaderIndex(headerRow, ["REGIÃO CIDADE", "REGIAO CIDADE", "REGIÃO", "REGIAO"]);
        const setorCol = findHeaderIndex(headerRow, [
            "SETOR DENTRO DA CIDADE",
            "SETOR",
            "SETOR CIDADE",
            "SETOR_CIDADE"
        ]);
        const categoriaCol = findHeaderIndex(headerRow, ["CATEGORIA"]);
        const tf1Col = findHeaderIndex(headerRow, ["TF1"]);
        const tf2Col = findHeaderIndex(headerRow, ["TF2"]);
        const pesqCol = findHeaderIndex(headerRow, [
            "Nº PESQ.",
            "Nº PESQ",
            "NO PESQ.",
            "NO PESQ",
            "NUMERO PESQ",
            "N° PESQ",
            "N°PESQ"
        ]);
        const dataPesquisaCol = findHeaderIndex(headerRow, [
            "DIA PESQ",
            "DIA DA PESQ",
            "DIA DA PESQUISA",
            "DIA PESQUISA",
            "DATA DA PESQUISA",
            "DATA PESQUISA"
        ]);

        ensureRequiredColumn(idpCol, "IDP");
        ensureRequiredColumn(cidadeCol, "CIDADE");
        ensureRequiredColumn(estadoCol, "ESTADO");
        ensureRequiredColumn(regiaoCol, "REGIÃO");
        ensureRequiredColumn(setorCol, "SETOR DENTRO DA CIDADE");
        ensureRequiredColumn(categoriaCol, "CATEGORIA");
        ensureRequiredColumn(tf1Col, "TF1");
        ensureRequiredColumn(tf2Col, "TF2");
        ensureRequiredColumn(pesqCol, "N° PESQ");
        ensureRequiredColumn(dataPesquisaCol, "DIA PESQ");

        const usefulRows = dataRows
            .map((row, dataIndex) => ({
                row,
                dataIndex,
                rowNumberExcel: dataStartIndex + dataIndex + 1,
                originalIndex: dataIndex,
                cidade: String(row[cidadeCol] ?? "").trim(),
                estado: String(row[estadoCol] ?? "").trim(),
                regiao: String(row[regiaoCol] ?? "").trim(),
                setor: String(row[setorCol] ?? "").trim(),
                categoria: String(row[categoriaCol] ?? "").trim(),
            }))
            .filter(({ row }) => !isRowEmpty(row));

        if (!usefulRows.length) {
            addLog("Não há linhas úteis na aba principal.", "error");
            return;
        }

        addLog(`Total de linhas úteis encontradas: ${usefulRows.length}`, "ok");
        addLog("Limpando colunas antigas de N° PESQ e DIA PESQ...", "ok");
        clearColumnsInRows(rows, dataStartIndex, [pesqCol, dataPesquisaCol]);

        if (hasCota) {
            await processWithCota({
                rows,
                usefulRows,
                dataPesquisaCol,
                pesqCol
            });
        } else {
            processWithoutCota({
                rows,
                usefulRows,
                dataPesquisaCol,
                pesqCol
            });
        }

        const sobraCount = markUnusedRowsAsSobra(usefulRows, rows, pesqCol, dataPesquisaCol);
        addLog(`Linhas marcadas como SOBRA: ${sobraCount}`, sobraCount > 0 ? "ok" : "warn");

        const finalHeaders = ["IDP", "CIDADE", "ESTADO", "REGIÃO", "SETOR", "CATEGORIA", "TF1", "TF2", "N° PESQ", "DIA PESQ"];
        const finalDataRows = buildFinalRows(
            rows,
            dataStartIndex,
            {
                idp: idpCol,
                cidade: cidadeCol,
                estado: estadoCol,
                regiao: regiaoCol,
                setor: setorCol,
                categoria: categoriaCol,
                tf1: tf1Col,
                tf2: tf2Col,
                pesq: pesqCol,
            },
            dataPesquisaCol
        );

        if (!finalDataRows.length) {
            addLog("Nenhuma linha final foi gerada.", "warn");
            return;
        }

        renderPreview(finalHeaders, finalDataRows);
        addLog(`Linhas exportadas no arquivo final: ${finalDataRows.length}`, "ok");

        const finalWorkbook = XLSX.utils.book_new();
        const finalSheetData = [finalHeaders, ...finalDataRows];
        const finalSheet = XLSX.utils.aoa_to_sheet(finalSheetData);

        autoFitSheetColumns(finalSheetData, finalSheet);

        XLSX.utils.book_append_sheet(finalWorkbook, finalSheet, "RESULTADO");

        const wbout = XLSX.write(finalWorkbook, {
            bookType: "xlsx",
            type: "array",
        });

        generatedFileName = `${safeFileName(baseFile.name)}_PREENCHIDO.xlsx`;
        generatedBlob = new Blob([wbout], {
            type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        });

        if (downloadBtn) downloadBtn.classList.remove("hidden");
        addLog(`Arquivo final pronto para download: ${generatedFileName}`, "ok");
    } catch (error) {
        console.error(error);
        addLog(`Erro ao processar: ${error.message || error}`, "error");
    } finally {
        processBtn.disabled = false;
        processBtn.textContent = "Processar arquivo";
    }
}

function processWithoutCota({
    rows,
    usefulRows,
    dataPesquisaCol,
    pesqCol
}) {
    addLog("Modo selecionado: SEM COTA", "ok");

    const maxPerPesq = getMaxNumbersPerPesq();
    const dayConfigs = getManualDayConfigs();

    if (!dayConfigs.length) {
        throw new Error("Informe ao menos um dia de pesquisa.");
    }

    dayConfigs.forEach((item, index) => {
        if (!item.pCount || item.pCount <= 0) {
            throw new Error(`Dia ${index + 1}: quantidade de pesquisadores inválida.`);
        }
        if (!item.date) {
            throw new Error(`Dia ${index + 1}: data inválida.`);
        }
    });

    const totalPesquisadores = dayConfigs.reduce((sum, item) => sum + item.pCount, 0);
    const totalNeeded = totalPesquisadores * maxPerPesq;

    addLog(`Total de pesquisadores informados: ${totalPesquisadores}`, "ok");
    addLog(`Telefones necessários: ${totalNeeded}`, "ok");

    if (usefulRows.length < totalNeeded) {
        const maxPerP = computeMaxPerPByAvailable(usefulRows.length, totalPesquisadores);
        throw new Error(`O máximo de telefones por P são ${maxPerP} de acordo com a quantidade de linhas.`);
    }

    const targets = [];

    dayConfigs.forEach((dayConfig) => {
        const pesquisadores = createPesquisadoresFromCount(dayConfig.pCount);

        pesquisadores.forEach((pesquisador) => {
            targets.push({
                pesquisador,
                date: dayConfig.date,
                remaining: maxPerPesq,
                rows: []
            });
        });
    });

    const { totalAssigned } = distributeRowsByCategoryRoundRobin(usefulRows, targets);

    const totalRemaining = targets.reduce((sum, target) => sum + target.remaining, 0);
    if (totalRemaining > 0) {
        const maxPerP = computeMaxPerPByAvailable(usefulRows.length, totalPesquisadores);
        throw new Error(`O máximo de telefones por P são ${maxPerP} de acordo com a quantidade de linhas.`);
    }

    const totalAtribuido = assignTargetsToSheet(targets, rows, pesqCol, dataPesquisaCol);

    dayConfigs.forEach((dayConfig, dayIndex) => {
        const totalNoDia = targets
            .filter((t) => t.date === dayConfig.date)
            .reduce((sum, t) => sum + t.rows.length, 0);

        addLog(
            `Dia ${dayIndex + 1} (${dayConfig.date}): ${dayConfig.pCount} pesquisadores preenchidos com ${Math.floor(totalNoDia / dayConfig.pCount)} telefones em média.`,
            "ok"
        );
    });

    addLog(`Total preenchido em N° PESQ: ${totalAtribuido}`, "ok");
    addLog(`Telefones usados na distribuição por categoria: ${totalAssigned}`, "ok");
    addLog(`Distribuição por categoria em rodízio: ${CATEGORY_PRIORITY.join(" → ")}`, "ok");
}

async function processWithCota({
    rows,
    usefulRows,
    dataPesquisaCol,
    pesqCol
}) {
    addLog("Modo selecionado: COM COTA", "ok");

    const cotaFile = cotaFileInput?.files?.[0];
    const cotaBuffer = await getArrayBuffer(cotaFile);
    const cotaText = decodeCsvArrayBuffer(cotaBuffer);
    const cotaRows = parseCSV(cotaText);

    if (!cotaRows.length) {
        throw new Error("O arquivo COTA está vazio ou inválido.");
    }

    const cotaRowsByDate = splitCotaRowsByDate(cotaRows);

    if (!cotaRowsByDate.size) {
        throw new Error('Não encontrei a coluna "DIA" preenchida corretamente no arquivo COTA.');
    }

    const sharedPools = buildSharedPoolsForCota(usefulRows, cotaRowsByDate);
    let totalAtribuido = 0;

    for (const [date, cotaRowsForDate] of cotaRowsByDate.entries()) {
        addLog(`Processando data ${date}...`, "ok");

        const demands = buildCotaDemandListForDate(cotaRowsForDate);

        if (!demands.length) {
            addLog(`Data ${date}: nenhuma demanda válida encontrada.`, "warn");
            continue;
        }

        const totalsByBucket = new Map();

        demands.forEach((demand) => {
            const bucketKey = `${demand.setorMode}::${demand.setorKey}`;

            if (!totalsByBucket.has(bucketKey)) {
                totalsByBucket.set(bucketKey, {
                    mode: demand.setorMode,
                    key: demand.setorKey,
                    totalQuestionarios: 0
                });
            }

            totalsByBucket.get(bucketKey).totalQuestionarios += demand.questionarios;
        });

        const multiplierByBucket = new Map();

        for (const [bucketKey, bucketInfo] of totalsByBucket.entries()) {
            const shared = sharedPools.get(bucketKey);
            const poolForDate = shared?.perDatePools?.get(date) || [];
            const available = poolForDate.length;

            const multiplier = bucketInfo.totalQuestionarios > 0
                ? Math.min(60, Math.floor(available / bucketInfo.totalQuestionarios))
                : 0;

            multiplierByBucket.set(bucketKey, multiplier);

            addLog(
                `Data ${date} | ${bucketInfo.mode === "estado" ? "Estado" : "Setor"} ${bucketInfo.key}: ${available} telefones disponíveis para esta data, ${bucketInfo.totalQuestionarios} questionários, multiplicador final ${multiplier}.`,
                multiplier > 0 ? "ok" : "warn"
            );
        }

        let totalData = 0;

        for (const [bucketKey] of totalsByBucket.entries()) {
            const bucketDemands = demands.filter((demand) => `${demand.setorMode}::${demand.setorKey}` === bucketKey);
            const multiplier = multiplierByBucket.get(bucketKey) || 0;
            if (multiplier <= 0) continue;

            const shared = sharedPools.get(bucketKey);
            const poolForDate = shared?.perDatePools?.get(date) || [];
            if (!poolForDate.length) continue;

            const targets = bucketDemands.map((demand) => ({
                pesquisador: demand.pesquisador,
                date,
                remaining: Math.floor(demand.questionarios * multiplier),
                rows: []
            }));

            const { totalAssigned, remainingPool } = distributeRowsByCategoryRoundRobin(poolForDate, targets);

            shared.perDatePools.set(date, remainingPool);

            totalData += totalAssigned;
            totalAtribuido += assignTargetsToSheet(targets, rows, pesqCol, dataPesquisaCol);

            addLog(
                `Data ${date} | Grupo ${bucketKey}: ${totalAssigned} telefones distribuídos em rodízio por categoria.`,
                totalAssigned > 0 ? "ok" : "warn"
            );
        }

        addLog(`Data ${date}: total preenchido ${totalData}.`, "ok");
    }

    addLog(`Total preenchido em N° PESQ: ${totalAtribuido}`, "ok");
    addLog(`Distribuição por categoria em rodízio por grupo: ${CATEGORY_PRIORITY.join(" → ")}`, "ok");
}

function renderPreview(headers, dataRows, limit = 60) {
    if (!previewHead || !previewBody || !resultSection) return;

    previewHead.innerHTML = "";
    previewBody.innerHTML = "";

    const trHead = document.createElement("tr");
    headers.forEach((header) => {
        const th = document.createElement("th");
        th.textContent = header ?? "";
        trHead.appendChild(th);
    });
    previewHead.appendChild(trHead);

    dataRows.slice(0, limit).forEach((row) => {
        const tr = document.createElement("tr");
        row.forEach((cell) => {
            const td = document.createElement("td");
            td.textContent = cell ?? "";
            tr.appendChild(td);
        });
        previewBody.appendChild(tr);
    });

    resultSection.classList.remove("hidden");
}