import { requireAuth, logout } from "../8-src/auth.js";

const user = await requireAuth({ redirectTo: "../7-login/login.html" });
if (!user) throw new Error("Sem sessão");

document.querySelector("#btnLogout")?.addEventListener("click", () => logout());

const baseFileInput = document.querySelector("#baseFile");
const cotaFileInput = document.querySelector("#cotaFile");
const hasCotaSelect = document.querySelector("#hasCota");
const maxNumbersPerPesqInput = document.querySelector("#maxNumbersPerPesq");
const maxPesquisaPerPesqInput = document.querySelector("#maxPesquisaPerPesq");
const processBtn = document.querySelector("#processBtn");
const downloadBtn = document.querySelector("#downloadBtn");

const manualDaysSection = document.querySelector("#manualDaysSection");
const manualDaysList = document.querySelector("#manualDaysList");
const addDayBtn = document.querySelector("#addDayBtn");

const logBox = document.querySelector("#logBox");

const previewSection = document.querySelector("#previewSection");
const previewMeta = document.querySelector("#previewMeta");
const previewTableWrapper = document.querySelector("#previewTableWrapper");

let manualDayCounter = 0;
let generatedBlob = null;
let generatedFileName = "";

const CATEGORY_PRIORITY = ["PESQUISA", "VIVO", "CLARO", "OI", "TIM", "BRASIL"];
const SOBRA_LABEL = "SOBRA";
const EXPORT_CHUNK_SIZE = 20000;
const UI_YIELD_EVERY = 5000;

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

function clearPreview() {
    if (previewTableWrapper) previewTableWrapper.innerHTML = "";
    if (previewMeta) {
        previewMeta.innerHTML = "";
        previewMeta.classList.add("hidden");
    }
    if (previewSection) previewSection.classList.add("hidden");
}

function renderPreviewTable(rows) {
    if (!previewSection || !previewTableWrapper) return;

    previewSection.classList.remove("hidden");

    let startIndex = 0;

    for (let i = 0; i < rows.length; i++) {
        const row = rows[i] || [];
        const normalizedRow = row.map(normalizeText);

        if (
            normalizedRow.includes("IDP") &&
            normalizedRow.includes("CIDADE") &&
            normalizedRow.includes("ESTADO")
        ) {
            startIndex = i;
            break;
        }
    }

    const previewRows = rows.slice(startIndex, startIndex + 30);

    if (!previewRows.length) {
        previewTableWrapper.innerHTML = `<div class="preview-empty">Nenhuma linha disponível para pré-visualização.</div>`;
        if (previewMeta) previewMeta.classList.add("hidden");
        return;
    }

    const headerRow = previewRows[0] || [];
    const bodyRows = previewRows.slice(1);

    let maxCols = 0;
    for (const row of previewRows) {
        if (Array.isArray(row) && row.length > maxCols) maxCols = row.length;
    }

    if (previewMeta) {
        previewMeta.classList.remove("hidden");
        previewMeta.innerHTML = `
            <strong>Cabeçalho detectado na linha:</strong> ${startIndex + 1}
            &nbsp;•&nbsp;
            <strong>Linhas exibidas:</strong> ${previewRows.length}
            &nbsp;•&nbsp;
            <strong>Colunas visíveis:</strong> ${maxCols}
        `;
    }

    let html = `
        <div class="preview-table-wrap">
            <table class="preview-table">
                <thead>
                    <tr>
                        <th class="preview-col-index">#</th>
    `;

    for (let colIndex = 0; colIndex < maxCols; colIndex++) {
        html += `<th>${escapeHtml(String(headerRow[colIndex] ?? ""))}</th>`;
    }

    html += `
                    </tr>
                </thead>
                <tbody>
    `;

    bodyRows.forEach((row, rowIndex) => {
        html += `<tr>`;
        html += `<td class="preview-col-index">${startIndex + 2 + rowIndex}</td>`;

        for (let colIndex = 0; colIndex < maxCols; colIndex++) {
            html += `<td>${escapeHtml(String(row?.[colIndex] ?? ""))}</td>`;
        }

        html += `</tr>`;
    });

    html += `
                </tbody>
            </table>
        </div>
    `;

    previewTableWrapper.innerHTML = html;
}

function escapeHtml(value) {
    return String(value ?? "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

/* =========================
   UTILITÁRIOS
========================= */

function sleepFrame() {
    return new Promise((resolve) => {
        setTimeout(resolve, 0);
    });
}

async function maybeYield(counter) {
    if (counter > 0 && counter % UI_YIELD_EVERY === 0) {
        await sleepFrame();
    }
}

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

async function removeFullyEmptyRows(rows = []) {
    const cleanedRows = [];
    let removedCount = 0;

    for (let i = 0; i < rows.length; i++) {
        const row = Array.isArray(rows[i]) ? rows[i] : [];

        if (isRowEmpty(row)) {
            removedCount += 1;
        } else {
            cleanedRows.push(row);
        }

        await maybeYield(i);
    }

    return { cleanedRows, removedCount };
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

    if (typeof value === "number" && Number.isFinite(value)) {
        const converted = formatDateToDDMMYYYY(excelDateToJSDate(value));
        if (converted) return converted;
    }

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

function parseCSVToRows(text) {
    const parsed = Papa.parse(text, {
        header: false,
        skipEmptyLines: false,
        delimiter: "",
        dynamicTyping: false,
    });

    return parsed.data || [];
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

function getMaxPesquisaPerPesq() {
    const value = Number(maxPesquisaPerPesqInput?.value);

    if (!Number.isFinite(value) || value < 0) {
        return 0;
    }

    return Math.floor(value);
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

function getPesqSortOrder(pesqValue) {
    const normalized = String(pesqValue ?? "").trim().toUpperCase();

    if (normalized === SOBRA_LABEL) return Number.MAX_SAFE_INTEGER - 1;
    if (!normalized) return Number.MAX_SAFE_INTEGER;

    const match = normalized.match(/^P(\d+)$/);
    if (match) return Number(match[1]);

    return Number.MAX_SAFE_INTEGER - 2;
}

function handleDownload() {
    if (!generatedBlob || !generatedFileName) {
        addLog("Nenhum arquivo final foi gerado ainda.", "warn");
        return;
    }
    triggerDownload(generatedBlob, generatedFileName);
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

function findHeaderRowIndex(rows, maxScan = 20) {
    for (let i = 0; i < Math.min(rows.length, maxScan); i++) {
        const row = rows[i] || [];
        const normalized = row.map(normalizeText);

        const hasIDP = normalized.includes("IDP");
        const hasCidade = normalized.includes("CIDADE");
        const hasEstado = normalized.includes("ESTADO");

        if (hasIDP && hasCidade && hasEstado) {
            return i;
        }
    }
    return -1;
}

/* =========================
   LINHAS LEVES
========================= */

async function buildUsefulRowRefs(rows, dataStartIndex, indexes) {
    const usefulRefs = [];

    for (let i = dataStartIndex; i < rows.length; i++) {
        const row = rows[i] || [];
        if (isRowEmpty(row)) continue;

        usefulRefs.push({
            rowIndex: i,
            rowNumberExcel: i + 1,
            originalIndex: i - dataStartIndex,
            estado: String(row[indexes.estado] ?? "").trim(),
            cidade: String(row[indexes.cidade] ?? "").trim(),
            regiao: String(row[indexes.regiao] ?? "").trim(),
            setor: String(row[indexes.setor] ?? "").trim(),
            categoria: String(row[indexes.categoria] ?? "").trim(),
        });

        await maybeYield(i - dataStartIndex);
    }

    return usefulRefs;
}

async function buildPoolsByMode(usefulRefs) {
    const estadoPools = new Map();
    const setorPools = new Map();

    for (let i = 0; i < usefulRefs.length; i++) {
        const ref = usefulRefs[i];
        const estadoKey = normalizeText(ref.estado);
        const setorKey = normalizeSetorValue(ref.setor);

        if (!estadoPools.has(estadoKey)) estadoPools.set(estadoKey, []);
        if (!setorPools.has(setorKey)) setorPools.set(setorKey, []);

        estadoPools.get(estadoKey).push(ref);
        setorPools.get(setorKey).push(ref);

        await maybeYield(i);
    }

    for (const [key, value] of estadoPools.entries()) {
        value.sort((a, b) => {
            const diff = getCategoryPriorityIndex(a.categoria) - getCategoryPriorityIndex(b.categoria);
            if (diff !== 0) return diff;
            return a.originalIndex - b.originalIndex;
        });
        estadoPools.set(key, value);
        await sleepFrame();
    }

    for (const [key, value] of setorPools.entries()) {
        value.sort((a, b) => {
            const diff = getCategoryPriorityIndex(a.categoria) - getCategoryPriorityIndex(b.categoria);
            if (diff !== 0) return diff;
            return a.originalIndex - b.originalIndex;
        });
        setorPools.set(key, value);
        await sleepFrame();
    }

    return { estadoPools, setorPools };
}

/* =========================
   DISTRIBUIÇÃO POR CATEGORIA
========================= */

async function buildOrderedCategoryGroups(rowRefs) {
    const grouped = new Map();
    const others = new Map();

    for (let i = 0; i < rowRefs.length; i++) {
        const item = rowRefs[i];
        const normalized = normalizeCategory(item.categoria);

        if (CATEGORY_PRIORITY.includes(normalized)) {
            if (!grouped.has(normalized)) grouped.set(normalized, []);
            grouped.get(normalized).push(item);
        } else {
            if (!others.has(normalized)) others.set(normalized, []);
            others.get(normalized).push(item);
        }

        await maybeYield(i);
    }

    const orderedGroups = [];

    for (let i = 0; i < CATEGORY_PRIORITY.length; i++) {
        const category = CATEGORY_PRIORITY[i];
        const arr = grouped.get(category) || [];
        if (arr.length) {
            arr.sort((a, b) => a.originalIndex - b.originalIndex);
            orderedGroups.push({
                category,
                rows: arr
            });
        }
        await sleepFrame();
    }

    const otherEntries = [...others.entries()].sort((a, b) => a[0].localeCompare(b[0]));

    for (let i = 0; i < otherEntries.length; i++) {
        const [category, arr] = otherEntries[i];
        arr.sort((a, b) => a.originalIndex - b.originalIndex);

        orderedGroups.push({
            category,
            rows: arr
        });

        if (i % 20 === 0) {
            await sleepFrame();
        }
    }

    return orderedGroups;
}

async function distributeRowsByCategoryRoundRobin(rowPool, targets, options = {}) {
    const maxPesquisaPerPesq =
        Number.isInteger(options.maxPesquisaPerPesq) && options.maxPesquisaPerPesq > 0
            ? options.maxPesquisaPerPesq
            : 0;

    const categoryGroups = await buildOrderedCategoryGroups(rowPool);
    let totalAssigned = 0;

    for (let t = 0; t < targets.length; t++) {
        if (!Array.isArray(targets[t].rows)) targets[t].rows = [];
        if (!Number.isFinite(targets[t].pesquisaAssigned)) targets[t].pesquisaAssigned = 0;
    }

    for (let groupIndex = 0; groupIndex < categoryGroups.length; groupIndex++) {
        const group = categoryGroups[groupIndex];
        let cursor = 0;
        const isPesquisaCategory = normalizeCategory(group.category) === "PESQUISA";

        while (cursor < group.rows.length) {
            const targetsNeedingRows = [];

            for (let t = 0; t < targets.length; t++) {
                if (targets[t].remaining > 0) {
                    targetsNeedingRows.push(targets[t]);
                }
            }

            if (!targetsNeedingRows.length) {
                const remainingPool = [];

                for (let i = groupIndex; i < categoryGroups.length; i++) {
                    const currentGroup = categoryGroups[i];

                    if (i === groupIndex) {
                        for (let j = cursor; j < currentGroup.rows.length; j++) {
                            remainingPool.push(currentGroup.rows[j]);
                        }
                    } else {
                        for (let j = 0; j < currentGroup.rows.length; j++) {
                            remainingPool.push(currentGroup.rows[j]);
                        }
                    }

                    await sleepFrame();
                }

                return { totalAssigned, remainingPool };
            }

            let assignedInThisRound = false;

            for (let i = 0; i < targetsNeedingRows.length; i++) {
                const target = targetsNeedingRows[i];

                if (cursor >= group.rows.length) break;
                if (target.remaining <= 0) continue;

                if (
                    isPesquisaCategory &&
                    maxPesquisaPerPesq > 0 &&
                    target.pesquisaAssigned >= maxPesquisaPerPesq
                ) {
                    continue;
                }

                const rowRef = group.rows[cursor];
                cursor += 1;

                target.rows.push(rowRef);
                target.remaining -= 1;
                totalAssigned += 1;
                assignedInThisRound = true;

                if (isPesquisaCategory) {
                    target.pesquisaAssigned += 1;
                }
            }

            if (!assignedInThisRound) {
                break;
            }

            if (cursor % UI_YIELD_EVERY === 0) {
                await sleepFrame();
            }
        }

        await sleepFrame();
    }

    const assignedSet = new Set();

    for (let i = 0; i < targets.length; i++) {
        const targetRows = targets[i].rows;
        for (let j = 0; j < targetRows.length; j++) {
            assignedSet.add(targetRows[j]);
        }
        await sleepFrame();
    }

    const remainingPool = [];
    for (let i = 0; i < rowPool.length; i++) {
        const rowRef = rowPool[i];
        if (!assignedSet.has(rowRef)) {
            remainingPool.push(rowRef);
        }
        await maybeYield(i);
    }

    return { totalAssigned, remainingPool };
}

function assignTargetsToSheet(targets, rows, pesqCol, dataPesquisaCol) {
    let total = 0;

    targets.forEach((target) => {
        target.rows.forEach((ref) => {
            if (!rows[ref.rowIndex]) rows[ref.rowIndex] = [];
            rows[ref.rowIndex][pesqCol] = target.pesquisador;
            rows[ref.rowIndex][dataPesquisaCol] = target.date;
            total += 1;
        });
    });

    return total;
}

function markUnusedRowsAsSobra(usefulRefs, rows, pesqCol, dataPesquisaCol) {
    let sobraCount = 0;

    usefulRefs.forEach((ref) => {
        if (!rows[ref.rowIndex]) rows[ref.rowIndex] = [];

        const currentPesq = String(rows[ref.rowIndex][pesqCol] ?? "").trim();

        if (!currentPesq) {
            rows[ref.rowIndex][pesqCol] = SOBRA_LABEL;
            rows[ref.rowIndex][dataPesquisaCol] = "";
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

async function buildSharedPoolsForCota(usefulRefs, cotaRowsByDate) {
    const bucketGroups = groupDemandsByBucketAcrossDates(cotaRowsByDate);
    const { estadoPools, setorPools } = await buildPoolsByMode(usefulRefs);

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

        for (let index = 0; index < dates.length; index++) {
            const date = dates[index];
            const qty = equalParts[index];
            const slice = sourcePool.slice(cursor, cursor + qty);
            cursor += qty;
            perDatePools.set(date, slice);
        }

        sharedPools.set(bucketKey, {
            mode: bucketInfo.mode,
            key: bucketInfo.key,
            perDatePools
        });

        await sleepFrame();
    }

    return sharedPools;
}

/* =========================
   SAÍDA FINAL
========================= */

async function buildFinalRowRefs(rows, dataStartIndex, columnIndexes, diaPesqColIndex) {
    const refs = [];

    for (let i = dataStartIndex; i < rows.length; i++) {
        const row = rows[i] || [];
        if (isRowEmpty(row)) continue;

        const pesqValue = String(row[columnIndexes.pesq] ?? "").trim();
        if (!pesqValue) continue;

        const diaPesqValue =
            diaPesqColIndex >= 0 ? getDisplayDate(row[diaPesqColIndex]) : "";

        refs.push({
            rowIndex: i,
            pesqOrder: getPesqSortOrder(pesqValue),
            dateValue: String(diaPesqValue ?? ""),
            categoryOrder: getCategoryPriorityIndex(row[columnIndexes.categoria] ?? ""),
            idpValue: String(row[columnIndexes.idp] ?? "")
        });

        await maybeYield(i - dataStartIndex);
    }

    refs.sort((a, b) => {
        const pesqDiff = a.pesqOrder - b.pesqOrder;
        if (pesqDiff !== 0) return pesqDiff;

        const dateDiff = a.dateValue.localeCompare(b.dateValue, "pt-BR");
        if (dateDiff !== 0) return dateDiff;

        const catDiff = a.categoryOrder - b.categoryOrder;
        if (catDiff !== 0) return catDiff;

        return a.idpValue.localeCompare(b.idpValue, "pt-BR", {
            numeric: true,
            sensitivity: "base"
        });
    });

    await sleepFrame();
    return refs;
}

function buildFinalOutputRow(rawRow, columnIndexes, diaPesqColIndex) {
    const diaPesqValue =
        diaPesqColIndex >= 0 ? getDisplayDate(rawRow[diaPesqColIndex]) : "";

    return [
        rawRow[columnIndexes.idp] ?? "",
        rawRow[columnIndexes.cidade] ?? "",
        rawRow[columnIndexes.estado] ?? "",
        rawRow[columnIndexes.regiao] ?? "",
        rawRow[columnIndexes.setor] ?? "",
        rawRow[columnIndexes.categoria] ?? "",
        rawRow[columnIndexes.tf1] ?? "",
        rawRow[columnIndexes.tf2] ?? "",
        rawRow[columnIndexes.pesq] ?? "",
        diaPesqValue ?? "",
    ];
}

function escapeCSVValue(value) {
    const str = String(value ?? "");
    if (/[;"\n\r,]/.test(str)) {
        return `"${str.replace(/"/g, '""')}"`;
    }
    return str;
}

async function buildFinalCSVInChunks({
    rows,
    sortedRefs,
    finalHeaders,
    columnIndexes,
    diaPesqColIndex
}) {
    const parts = [];

    parts.push(
        "\uFEFF" + finalHeaders.map(escapeCSVValue).join(";") + "\r\n"
    );

    for (let start = 0; start < sortedRefs.length; start += EXPORT_CHUNK_SIZE) {
        const end = Math.min(start + EXPORT_CHUNK_SIZE, sortedRefs.length);
        let chunkText = "";

        for (let i = start; i < end; i++) {
            const ref = sortedRefs[i];
            const outputRow = buildFinalOutputRow(
                rows[ref.rowIndex] || [],
                columnIndexes,
                diaPesqColIndex
            );

            chunkText += outputRow.map(escapeCSVValue).join(";") + "\r\n";
        }

        parts.push(chunkText);
        await sleepFrame();
    }

    return parts.join("");
}

/* =========================
   PROCESSAMENTO PRINCIPAL
========================= */

async function handleProcess() {
    try {
        clearLog();
        clearPreview();

        if (downloadBtn) downloadBtn.classList.add("hidden");
        generatedBlob = null;
        generatedFileName = "";

        const baseFile = baseFileInput?.files?.[0];
        const hasCota = hasCotaSelect?.value === "sim";

        if (!baseFile) {
            addLog("Selecione o arquivo base em CSV.", "error");
            return;
        }

        if (!String(baseFile.name || "").toLowerCase().endsWith(".csv")) {
            addLog("O arquivo base agora aceita somente CSV.", "error");
            addLog('Abra sua planilha e salve a aba principal como "CSV UTF-8".', "warn");
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

        let rows = [];
        let sourceLabel = "";

        addLog("Lendo base CSV...", "ok");
        const baseBuffer = await getArrayBuffer(baseFile);
        const baseText = decodeCsvArrayBuffer(baseBuffer);
        rows = parseCSVToRows(baseText);
        sourceLabel = "CSV";

        addLog(`Fonte carregada: CSV | ${sourceLabel}`, "ok");
        addLog(`Total bruto de linhas lidas: ${rows.length}`, "ok");

        const emptyRowCleanup = await removeFullyEmptyRows(rows);
        rows = emptyRowCleanup.cleanedRows;

        addLog(
            `Linhas totalmente vazias removidas: ${emptyRowCleanup.removedCount}`,
            emptyRowCleanup.removedCount > 0 ? "ok" : "warn"
        );
        addLog(`Total de linhas restantes após limpeza: ${rows.length}`, "ok");

        renderPreviewTable(rows);

        if (!rows.length) {
            addLog("Não consegui materializar as linhas da base CSV.", "error");
            return;
        }

        const headerRowIndex = findHeaderRowIndex(rows);

        if (headerRowIndex === -1) {
            addLog("Não consegui localizar a linha de cabeçalho na base CSV.", "error");
            console.log("Prévia sem cabeçalho detectado:", rows.slice(0, 30));
            return;
        }

        addLog(`Linha de cabeçalho detectada: ${headerRowIndex + 1}`, "ok");

        const headerRow = rows[headerRowIndex] || [];
        const dataStartIndex = headerRowIndex + 1;

        const idpCol = findHeaderIndex(headerRow, ["IDP"]);
        const cidadeCol = findHeaderIndex(headerRow, ["CIDADE"]);
        const estadoCol = findHeaderIndex(headerRow, ["ESTADO"]);
        const regiaoCol = findHeaderIndex(headerRow, ["REGIÃO CIDADE", "REGIAO CIDADE", "REGIÃO", "REGIAO"]);
        const setorCol = findHeaderIndex(headerRow, [
            "SETOR DENTRO DA CIDADE",
            "SETOR",
            "SETOR CIDADE",
            "SETOR_CIDADE",
            "SETO"
        ]);
        const categoriaCol = findHeaderIndex(headerRow, ["CATEGORIA", "CATE"]);
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

        const usefulRefs = await buildUsefulRowRefs(rows, dataStartIndex, {
            cidade: cidadeCol,
            estado: estadoCol,
            regiao: regiaoCol,
            setor: setorCol,
            categoria: categoriaCol
        });

        if (!usefulRefs.length) {
            addLog("Não há linhas úteis na base principal.", "error");
            return;
        }

        addLog(`Total de linhas úteis encontradas: ${usefulRefs.length}`, "ok");
        addLog("Limpando colunas antigas de N° PESQ e DIA PESQ...", "ok");
        clearColumnsInRows(rows, dataStartIndex, [pesqCol, dataPesquisaCol]);

        if (hasCota) {
            await processWithCota({
                rows,
                usefulRefs,
                dataPesquisaCol,
                pesqCol
            });
        } else {
            await processWithoutCota({
                rows,
                usefulRefs,
                dataPesquisaCol,
                pesqCol
            });
        }

        const sobraCount = markUnusedRowsAsSobra(usefulRefs, rows, pesqCol, dataPesquisaCol);
        addLog(`Linhas marcadas como SOBRA: ${sobraCount}`, sobraCount > 0 ? "ok" : "warn");

        const finalHeaders = [
            "IDP",
            "CIDADE",
            "ESTADO",
            "REGIÃO",
            "SETOR",
            "CATEGORIA",
            "TF1",
            "TF2",
            "N° PESQ",
            "DIA PESQ"
        ];

        addLog("Montando referências finais para exportação...", "ok");

        const columnIndexes = {
            idp: idpCol,
            cidade: cidadeCol,
            estado: estadoCol,
            regiao: regiaoCol,
            setor: setorCol,
            categoria: categoriaCol,
            tf1: tf1Col,
            tf2: tf2Col,
            pesq: pesqCol,
        };

        const finalRowRefs = await buildFinalRowRefs(
            rows,
            dataStartIndex,
            columnIndexes,
            dataPesquisaCol
        );

        if (!finalRowRefs.length) {
            addLog("Nenhuma linha final foi gerada.", "warn");
            return;
        }

        addLog(`Linhas exportadas no arquivo final: ${finalRowRefs.length}`, "ok");
        addLog("Gerando arquivo CSV final em blocos para economizar memória...", "ok");

        const csvText = await buildFinalCSVInChunks({
            rows,
            sortedRefs: finalRowRefs,
            finalHeaders,
            columnIndexes,
            diaPesqColIndex: dataPesquisaCol
        });

        generatedFileName = `${safeFileName(baseFile.name)}_PREENCHIDO.csv`;
        generatedBlob = new Blob([csvText], {
            type: "text/csv;charset=utf-8;",
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

async function processWithoutCota({
    rows,
    usefulRefs,
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

    if (usefulRefs.length < totalNeeded) {
        const maxPerP = computeMaxPerPByAvailable(usefulRefs.length, totalPesquisadores);
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

    const { totalAssigned } = await distributeRowsByCategoryRoundRobin(usefulRefs, targets);

    const totalRemaining = targets.reduce((sum, target) => sum + target.remaining, 0);
    if (totalRemaining > 0) {
        const maxPerP = computeMaxPerPByAvailable(usefulRefs.length, totalPesquisadores);
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
    usefulRefs,
    dataPesquisaCol,
    pesqCol
}) {
    addLog("Modo selecionado: COM COTA", "ok");

    const cotaFile = cotaFileInput?.files?.[0];
    if (!String(cotaFile?.name || "").toLowerCase().endsWith(".csv")) {
        throw new Error("O arquivo COTA deve estar em CSV.");
    }

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

    const sharedPools = await buildSharedPoolsForCota(usefulRefs, cotaRowsByDate);
    let totalAtribuido = 0;

    for (const [date, cotaRowsForDate] of cotaRowsByDate.entries()) {
        addLog(`Processando data ${date}...`, "ok");
        await sleepFrame();

        const demands = buildCotaDemandListForDate(cotaRowsForDate);

        if (!demands.length) {
            addLog(`Data ${date}: nenhuma demanda válida encontrada.`, "warn");
            continue;
        }

        const totalsByBucket = new Map();
        const demandsByBucket = new Map();

        for (let i = 0; i < demands.length; i++) {
            const demand = demands[i];
            const bucketKey = `${demand.setorMode}::${demand.setorKey}`;

            if (!totalsByBucket.has(bucketKey)) {
                totalsByBucket.set(bucketKey, {
                    mode: demand.setorMode,
                    key: demand.setorKey,
                    totalQuestionarios: 0
                });
            }

            if (!demandsByBucket.has(bucketKey)) {
                demandsByBucket.set(bucketKey, []);
            }

            totalsByBucket.get(bucketKey).totalQuestionarios += demand.questionarios;
            demandsByBucket.get(bucketKey).push(demand);
        }

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

            await sleepFrame();
        }

        let totalData = 0;

        for (const [bucketKey] of totalsByBucket.entries()) {
            const bucketDemands = demandsByBucket.get(bucketKey) || [];
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

            const { totalAssigned, remainingPool } = await distributeRowsByCategoryRoundRobin(poolForDate, targets);

            shared.perDatePools.set(date, remainingPool);

            totalData += totalAssigned;
            totalAtribuido += assignTargetsToSheet(targets, rows, pesqCol, dataPesquisaCol);

            addLog(
                `Data ${date} | Grupo ${bucketKey}: ${totalAssigned} telefones distribuídos em rodízio por categoria.`,
                totalAssigned > 0 ? "ok" : "warn"
            );

            await sleepFrame();
        }

        addLog(`Data ${date}: total preenchido ${totalData}.`, "ok");
        await sleepFrame();
    }

    addLog(`Total preenchido em N° PESQ: ${totalAtribuido}`, "ok");
    addLog(`Distribuição por categoria em rodízio por grupo: ${CATEGORY_PRIORITY.join(" → ")}`, "ok");
}