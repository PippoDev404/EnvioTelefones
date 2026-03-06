// /5-js/db.js (PT + camelCase) — IndexedDB central do sistema
// Stores: arquivos, usuarios, distribuicoes
//
// ✅ PK de ARQUIVOS: arquivoKey (string) ex: "ARQ-1710000000000"
// ✅ Distribuições referenciam arquivo por arquivoKeyOrigem (string)

const nomeBanco = "ibespeBanco";
const versaoBanco = 4;

const nomeStoreArquivos = "arquivos";
const nomeStoreUsuarios = "usuarios";
const nomeStoreDistribuicoes = "distribuicoes";

function gerarArquivoKey() {
  return `ARQ-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function parseJsonPossivel(v, fallback) {
  if (v == null) return fallback;
  if (typeof v === "object") return v;
  if (typeof v !== "string") return fallback;
  const s = v.trim();
  if (!s) return fallback;
  try {
    return JSON.parse(s);
  } catch {
    return fallback;
  }
}

function abrirBanco() {
  return new Promise((resolve, reject) => {
    const req = indexedDB.open(nomeBanco, versaoBanco);

    req.onupgradeneeded = () => {
      const db = req.result;

      // ======================
      // ARQUIVOS (PK: arquivoKey string)
      // ======================
      if (db.objectStoreNames.contains(nomeStoreArquivos)) {
        // ⚠️ isso apaga tudo ao aumentar versão. Ok se você quer "resetar"
        db.deleteObjectStore(nomeStoreArquivos);
      }

      const stArq = db.createObjectStore(nomeStoreArquivos, { keyPath: "arquivoKey" });
      stArq.createIndex("porDataCriacao", "dataCriacao");
      stArq.createIndex("porArquivoKey", "arquivoKey", { unique: true });

      // ======================
      // USUÁRIOS
      // ======================
      if (!db.objectStoreNames.contains(nomeStoreUsuarios)) {
        const stUsr = db.createObjectStore(nomeStoreUsuarios, { keyPath: "id" });

        stUsr.createIndex("porEmail", "email", { unique: true });
        stUsr.createIndex("porAdminId", "adminId");
        stUsr.createIndex("porParte", "parte");

        stUsr.createIndex("porTelegramUserId", "telegramUserId", { unique: false });
        stUsr.createIndex("porTelegramUsername", "telegramUsername", { unique: false });
      } else {
        const stUsr = req.transaction.objectStore(nomeStoreUsuarios);

        if (!stUsr.indexNames.contains("porEmail")) stUsr.createIndex("porEmail", "email", { unique: true });
        if (!stUsr.indexNames.contains("porAdminId")) stUsr.createIndex("porAdminId", "adminId");
        if (!stUsr.indexNames.contains("porParte")) stUsr.createIndex("porParte", "parte");

        if (!stUsr.indexNames.contains("porTelegramUserId")) stUsr.createIndex("porTelegramUserId", "telegramUserId", { unique: false });
        if (!stUsr.indexNames.contains("porTelegramUsername")) stUsr.createIndex("porTelegramUsername", "telegramUsername", { unique: false });
      }

      // ======================
      // DISTRIBUIÇÕES
      // ======================
      if (!db.objectStoreNames.contains(nomeStoreDistribuicoes)) {
        const stDist = db.createObjectStore(nomeStoreDistribuicoes, { keyPath: "id" });

        stDist.createIndex("porArquivoKeyOrigem", "arquivoKeyOrigem");
        stDist.createIndex("porUsuarioId", "usuarioId");
        stDist.createIndex("porCriadoEm", "criadoEm");
        stDist.createIndex("porChaveParte", "chaveParte");
      } else {
        const stDist = req.transaction.objectStore(nomeStoreDistribuicoes);

        if (!stDist.indexNames.contains("porArquivoKeyOrigem")) stDist.createIndex("porArquivoKeyOrigem", "arquivoKeyOrigem");
        if (!stDist.indexNames.contains("porUsuarioId")) stDist.createIndex("porUsuarioId", "usuarioId");
        if (!stDist.indexNames.contains("porCriadoEm")) stDist.createIndex("porCriadoEm", "criadoEm");
        if (!stDist.indexNames.contains("porChaveParte")) stDist.createIndex("porChaveParte", "chaveParte");
      }
    };

    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
  });
}

// =========================
// HELPERS TX
// =========================
async function txStore(nomeStore, modo = "readonly") {
  const db = await abrirBanco();
  const tx = db.transaction(nomeStore, modo);
  return { db, tx, store: tx.objectStore(nomeStore) };
}

function normalizarArquivoKey(arquivoKeyOuObj) {
  if (typeof arquivoKeyOuObj === "string") return arquivoKeyOuObj.trim();
  const o = arquivoKeyOuObj || {};
  const k = o.arquivoKey || o.arquivo_key || o.key || o.id;
  return k != null ? String(k).trim() : "";
}

// =========================
// ARQUIVOS
// =========================
async function salvarArquivoDb(registro) {
  const { tx, store } = await txStore(nomeStoreArquivos, "readwrite");

  const colunasVisiveisPadrao = { estado: true, cidade: true, regiao: true };

  let arquivoKey = normalizarArquivoKey(registro);
  if (!arquivoKey) arquivoKey = gerarArquivoKey();

  // ✅ aceita string JSON do supabase/n8n
  const colunasVisiveisNormalizadas = parseJsonPossivel(registro?.colunasVisiveis, registro?.colunasVisiveis || {});
  const opcoesPesquisaNormalizadas = parseJsonPossivel(registro?.opcoesPesquisa, registro?.opcoesPesquisa || null);

  const registroSeguro = {
    ...registro,
    arquivoKey,
    arquivo_key: registro?.arquivo_key || arquivoKey,

    opcoesPesquisa: opcoesPesquisaNormalizadas || registro?.opcoesPesquisa || undefined,

    colunasVisiveis: {
      ...colunasVisiveisPadrao,
      ...(colunasVisiveisNormalizadas || {}),
    },
  };

  return new Promise((resolve, reject) => {
    store.put(registroSeguro);
    tx.oncomplete = () => resolve(true);
    tx.onerror = () => reject(tx.error);
  });
}

async function listarArquivosDb() {
  const { store } = await txStore(nomeStoreArquivos, "readonly");
  return new Promise((resolve, reject) => {
    const req = store.getAll();
    req.onsuccess = () => {
      const lista = (req.result || []).sort((a, b) => (b.dataCriacao || 0) - (a.dataCriacao || 0));
      resolve(lista);
    };
    req.onerror = () => reject(req.error);
  });
}

async function pegarArquivoPorKeyDb(arquivoKey) {
  const { store } = await txStore(nomeStoreArquivos, "readonly");
  const k = normalizarArquivoKey(String(arquivoKey || ""));
  return new Promise((resolve, reject) => {
    const req = store.get(k);
    req.onsuccess = () => resolve(req.result || null);
    req.onerror = () => reject(req.error);
  });
}

// compat
async function pegarArquivoDb(idOuKey) {
  return await pegarArquivoPorKeyDb(String(idOuKey || ""));
}

async function excluirArquivosPorKeyDb(chaves) {
  const { tx, store } = await txStore(nomeStoreArquivos, "readwrite");
  const keys = (chaves || []).map((k) => String(k).trim()).filter(Boolean);
  return new Promise((resolve, reject) => {
    for (const k of keys) store.delete(k);
    tx.oncomplete = () => resolve(true);
    tx.onerror = () => reject(tx.error);
  });
}

// compat
async function excluirArquivosDb(idsOuKeys) {
  return await excluirArquivosPorKeyDb(idsOuKeys);
}

async function excluirArquivosAntigosDb(maxDias) {
  const agora = Date.now();
  const maxMs = maxDias * 24 * 60 * 60 * 1000;

  const lista = await listarArquivosDb();
  const keysParaExcluir = lista
    .filter((a) => a.dataCriacao && (agora - a.dataCriacao) > maxMs)
    .map((a) => a.arquivoKey)
    .filter(Boolean);

  if (keysParaExcluir.length) await excluirArquivosPorKeyDb(keysParaExcluir);
  return keysParaExcluir.length;
}

// =========================
// USUÁRIOS
// =========================
async function salvarUsuarioDb(usuario) {
  const { tx, store } = await txStore(nomeStoreUsuarios, "readwrite");
  return new Promise((resolve, reject) => {
    store.put(usuario);
    tx.oncomplete = () => resolve(true);
    tx.onerror = () => reject(tx.error);
  });
}

async function listarUsuariosDb({ adminId } = {}) {
  const { store } = await txStore(nomeStoreUsuarios, "readonly");
  return new Promise((resolve, reject) => {
    if (!adminId) {
      const req = store.getAll();
      req.onsuccess = () => resolve(req.result || []);
      req.onerror = () => reject(req.error);
      return;
    }

    const idx = store.index("porAdminId");
    const req = idx.getAll(adminId);
    req.onsuccess = () => resolve(req.result || []);
    req.onerror = () => reject(req.error);
  });
}

async function pegarUsuarioDb(id) {
  const { store } = await txStore(nomeStoreUsuarios, "readonly");
  return new Promise((resolve, reject) => {
    const req = store.get(String(id));
    req.onsuccess = () => resolve(req.result || null);
    req.onerror = () => reject(req.error);
  });
}

async function pegarUsuarioPorEmailDb(email) {
  const { store } = await txStore(nomeStoreUsuarios, "readonly");
  return new Promise((resolve, reject) => {
    const idx = store.index("porEmail");
    const req = idx.get(String(email || "").toLowerCase());
    req.onsuccess = () => resolve(req.result || null);
    req.onerror = () => reject(req.error);
  });
}

async function pegarUsuarioPorTelegramIdDb(telegramUserId) {
  const { store } = await txStore(nomeStoreUsuarios, "readonly");
  return new Promise((resolve, reject) => {
    const idx = store.index("porTelegramUserId");
    const req = idx.get(String(telegramUserId || ""));
    req.onsuccess = () => resolve(req.result || null);
    req.onerror = () => reject(req.error);
  });
}

async function excluirUsuariosDb(ids) {
  const { tx, store } = await txStore(nomeStoreUsuarios, "readwrite");
  return new Promise((resolve, reject) => {
    for (const id of ids) store.delete(String(id));
    tx.oncomplete = () => resolve(true);
    tx.onerror = () => reject(tx.error);
  });
}

// =========================
// DISTRIBUIÇÃO
// =========================
function gerarIdDistribuicao() {
  return `dist-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

async function salvarDistribuicaoDb(itens) {
  if (!Array.isArray(itens) || !itens.length) return true;

  const { tx, store } = await txStore(nomeStoreDistribuicoes, "readwrite");

  return new Promise((resolve, reject) => {
    for (const item of itens) {
      const registro = { ...item, id: item.id || gerarIdDistribuicao() };

      if (registro.arquivoKeyOrigem == null) {
        const fallback = registro.arquivoKey || registro.arquivo_key || registro.idArquivoOrigem;
        if (fallback != null) registro.arquivoKeyOrigem = String(fallback);
      } else {
        registro.arquivoKeyOrigem = String(registro.arquivoKeyOrigem);
      }

      if (registro.idArquivoOrigem != null && typeof registro.idArquivoOrigem !== "string") {
        registro.idArquivoOrigem = String(registro.idArquivoOrigem);
      }

      store.put(registro);
    }

    tx.oncomplete = () => resolve(true);
    tx.onerror = () => reject(tx.error);
  });
}

async function listarDistribuicoesPorUsuarioDb(usuarioId) {
  const { store } = await txStore(nomeStoreDistribuicoes, "readonly");
  return new Promise((resolve, reject) => {
    const idx = store.index("porUsuarioId");
    const req = idx.getAll(String(usuarioId));
    req.onsuccess = () => resolve(req.result || []);
    req.onerror = () => reject(req.error);
  });
}

async function listarDistribuicoesPorArquivoKeyDb(arquivoKeyOrigem) {
  const { store } = await txStore(nomeStoreDistribuicoes, "readonly");
  return new Promise((resolve, reject) => {
    const idx = store.index("porArquivoKeyOrigem");
    const req = idx.getAll(String(arquivoKeyOrigem));
    req.onsuccess = () => resolve(req.result || []);
    req.onerror = () => reject(req.error);
  });
}

// compat
async function listarDistribuicoesPorArquivoDb(idOuKey) {
  return await listarDistribuicoesPorArquivoKeyDb(String(idOuKey || ""));
}

// Exporta no window
window.ibeDb = {
  // arquivos
  salvarArquivoDb,
  listarArquivosDb,
  pegarArquivoDb,
  pegarArquivoPorKeyDb,
  excluirArquivosDb,
  excluirArquivosPorKeyDb,
  excluirArquivosAntigosDb,

  // usuarios
  salvarUsuarioDb,
  listarUsuariosDb,
  pegarUsuarioDb,
  pegarUsuarioPorEmailDb,
  pegarUsuarioPorTelegramIdDb,
  excluirUsuariosDb,

  // distribuicao
  salvarDistribuicaoDb,
  listarDistribuicoesPorUsuarioDb,
  listarDistribuicoesPorArquivoDb,
  listarDistribuicoesPorArquivoKeyDb,
};