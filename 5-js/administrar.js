// /5-js/administrar.js
import "../5-js/db.js";
import { requireAuth, logout } from "../8-src/auth.js";
import { supabase } from "../8-src/supabaseClient.js";

const user = await requireAuth({ redirectTo: "../7-login/login.html" });
if (!user) throw new Error("Sem sessão"); // segurança

document.querySelector("#btnLogout")?.addEventListener("click", () => logout());

// ============================
// CONFIG / ADMIN LOCAL
// ============================
const chaveAdminAtual = "adminAtual";
const adminPadrao = { id: "admin-001", nome: "Administrador", email: "admin@ibespe.com.br" };

// ============================
// UTIL
// ============================
function carregarJson(chave, padrao) {
  try {
    return JSON.parse(localStorage.getItem(chave)) ?? padrao;
  } catch {
    return padrao;
  }
}
function salvarJson(chave, valor) {
  localStorage.setItem(chave, JSON.stringify(valor));
}
function formatarDataBr(data = new Date()) {
  const d = String(data.getDate()).padStart(2, "0");
  const m = String(data.getMonth() + 1).padStart(2, "0");
  const a = data.getFullYear();
  return `${d}/${m}/${a}`;
}

function normalizarParte(valor) {
  const v = String(valor || "").trim().toUpperCase();
  if (!v) return "";
  const m = v.match(/^P\s*0*(\d+)$/i) || v.match(/^P(\d+)$/i);
  if (m) {
    const n = parseInt(m[1], 10);
    if (!Number.isFinite(n) || n <= 0) return "";
    return `P${String(n).padStart(2, "0")}`;
  }
  const m2 = v.match(/^(\d+)(?:\.\d+)?$/);
  if (m2) {
    const n = parseInt(m2[1], 10);
    if (!Number.isFinite(n) || n <= 0) return "";
    return `P${String(n).padStart(2, "0")}`;
  }
  return "";
}

function normalizarTelegramId(valor) {
  const s = String(valor || "").trim();
  if (!s) return "";
  const onlyDigits = s.replace(/[^\d]/g, "");
  if (!onlyDigits) return "";
  if (onlyDigits.length < 5) return "";
  return onlyDigits;
}

function normalizarTelegramUser(valor) {
  let s = String(valor || "").trim();
  if (!s) return "";
  if (s.startsWith("@")) s = s.slice(1);
  s = s.replace(/\s+/g, "");
  if (!s) return "";
  return `@${s}`;
}

function garantirArroba(v) {
  const s = String(v || "").trim();
  if (!s) return "";
  return s.startsWith("@") ? s : `@${s}`;
}

const gerarIdUsuario = () => `usr-${Date.now()}-${Math.random().toString(16).slice(2)}`;

// ============================
// DOM (resumo)
// ============================
const adminNome = document.getElementById("adminNome");
const adminEmail = document.getElementById("adminEmail");
const qtdUsuarios = document.getElementById("qtdUsuarios");
const qtdAtivos = document.getElementById("qtdAtivos");

// DOM (form)
const formVincular = document.getElementById("formVincular");
const inputTelegramId = document.getElementById("inputTelegramId");
const inputTelegramUser = document.getElementById("inputTelegramUser");

// DOM (filtros)
const inputBusca = document.getElementById("inputBusca");
const selectStatus = document.getElementById("selectStatus");
const botaoDesvincularSelecionados = document.getElementById("botaoDesvincularSelecionados");
const textoSelecionados = document.getElementById("textoSelecionados");
const checkboxSelecionarTodos = document.getElementById("checkboxSelecionarTodos");

// DOM (tabela)
const corpoTabelaUsuarios = document.getElementById("corpoTabelaUsuarios");
const estadoVazio = document.getElementById("estadoVazio");

// Modal confirmação
const overlayConfirmacao = document.getElementById("overlayConfirmacao");
const mensagemConfirmacao = document.getElementById("mensagemConfirmacao");
const botaoCancelar = document.getElementById("botaoCancelar");
const botaoConfirmar = document.getElementById("botaoConfirmar");
let acaoConfirmacao = null;

// ============================
// UI HELPERS
// ============================
function classeStatus(status) {
  const s = String(status || "").toLowerCase();
  return s === "ativo" ? "badge badgeAtivo" : "badge badgeBloqueado";
}

function aviso(msg) {
  abrirModalConfirmacao(msg, () => {});
}

// ============================
// ADMIN (local)
// ============================
function garantirAdmin() {
  const adminAtual = carregarJson(chaveAdminAtual, null);
  if (!adminAtual) salvarJson(chaveAdminAtual, adminPadrao);
}
function obterAdminAtual() {
  return carregarJson(chaveAdminAtual, adminPadrao);
}

// ============================
// DB helpers (IndexedDB)
// ============================
async function listarUsuariosDoAdminIndexedDb() {
  const admin = obterAdminAtual();
  if (!window.ibeDb?.listarUsuariosDb) return [];
  const lista = await window.ibeDb.listarUsuariosDb({ adminId: admin.id });
  return (lista || []).sort((a, b) => (b.criadoEm || 0) - (a.criadoEm || 0));
}

async function salvarUsuarioIndexedDb(usuario) {
  if (!window.ibeDb?.salvarUsuarioDb) return;
  await window.ibeDb.salvarUsuarioDb(usuario);
}

async function excluirUsuariosIndexedDb(ids) {
  if (!window.ibeDb?.excluirUsuariosDb) return;
  await window.ibeDb.excluirUsuariosDb(ids);
}

// ============================
// SUPABASE helpers
// ============================
// ✅ você disse que precisa: id, owner_id, parte, telegram_user_id, telegram_username, criado_em
// status é opcional (só funciona se existir a coluna).

function mapSupabaseParaUsuario(row) {
  const tgId = String(row?.telegram_user_id || "").trim();
  const tgUser = garantirArroba(String(row?.telegram_username || "").trim());

  const status = String(row?.status || "ativo").toLowerCase() === "bloqueado" ? "bloqueado" : "ativo";

  return {
    id: String(row?.id || "").trim(),
    status,
    parte: String(row?.parte || "").trim(),
    telegramUserId: tgId,
    telegramUsername: tgUser,
    criadoEm: row?.criado_em ? new Date(row.criado_em).getTime() : Date.now(),
    dataTexto: row?.criado_em ? new Date(row.criado_em).toLocaleDateString("pt-BR") : "",
  };
}

async function listarUsuariosSupabase() {
  const ownerId = user?.id;
  if (!ownerId) return [];

  // tenta com status
  {
    const { data, error } = await supabase
      .from("usuarios")
      .select("id,status,parte,telegram_user_id,telegram_username,criado_em")
      .eq("owner_id", ownerId)
      .order("criado_em", { ascending: false });

    if (!error) return (data || []).map(mapSupabaseParaUsuario);

    const msg = String(error?.message || "");
    if (!msg.toLowerCase().includes("status")) {
      console.warn("Erro ao listar usuários no Supabase:", msg);
      return [];
    }
  }

  // fallback sem status
  const { data, error } = await supabase
    .from("usuarios")
    .select("id,parte,telegram_user_id,telegram_username,criado_em")
    .eq("owner_id", ownerId)
    .order("criado_em", { ascending: false });

  if (error) {
    console.warn("Erro ao listar usuários no Supabase:", error.message);
    return [];
  }

  return (data || []).map(mapSupabaseParaUsuario);
}

async function salvarUsuarioSupabase(usuario) {
  const ownerId = user?.id;
  if (!ownerId) throw new Error("Sem sessão/ownerId.");

  const payloadBase = {
    id: String(usuario?.id || "").trim(),
    owner_id: ownerId,
    parte: normalizarParte(usuario?.parte || ""),
    telegram_user_id: String(usuario?.telegramUserId || "").trim(),
    telegram_username: String(usuario?.telegramUsername || "").trim().replace(/^@/, ""),
    criado_em: usuario?.criadoEm ? new Date(Number(usuario.criadoEm)).toISOString() : new Date().toISOString(),
  };

  // tenta com status se existir
  const payloadComStatus = {
    ...payloadBase,
    status: String(usuario?.status || "ativo"),
  };

  {
    const { error } = await supabase.from("usuarios").upsert([payloadComStatus], { onConflict: "id" });
    if (!error) return true;

    const msg = String(error?.message || "");
    if (!msg.toLowerCase().includes("status")) throw new Error(msg);
  }

  const { error: e2 } = await supabase.from("usuarios").upsert([payloadBase], { onConflict: "id" });
  if (e2) throw new Error(e2.message);
  return true;
}

async function excluirUsuariosSupabase(ids) {
  const ownerId = user?.id;
  if (!ownerId) throw new Error("Sem sessão/ownerId.");

  const lista = (ids || []).map((x) => String(x || "").trim()).filter(Boolean);
  if (!lista.length) return true;

  const { error } = await supabase.from("usuarios").delete().eq("owner_id", ownerId).in("id", lista);
  if (error) throw new Error(error.message);
  return true;
}

async function atualizarStatusUsuarioSupabase(usuarioId, novoStatus) {
  const ownerId = user?.id;
  if (!ownerId) throw new Error("Sem sessão/ownerId.");

  const status = String(novoStatus || "").toLowerCase();
  if (!["ativo", "bloqueado"].includes(status)) throw new Error("Status inválido.");

  const { error } = await supabase.from("usuarios").update({ status }).eq("owner_id", ownerId).eq("id", String(usuarioId));

  if (!error) return true;

  const msg = String(error?.message || "");
  if (msg.toLowerCase().includes("status")) {
    throw new Error("A coluna 'status' não existe no Supabase. Crie a coluna status (texto) para habilitar Bloquear/Ativar.");
  }
  throw new Error(msg);
}

// ============================
// SYNC IndexedDB -> Supabase
// ============================
async function syncUsuariosIndexedDbParaSupabase() {
  try {
    const ownerId = user?.id;
    if (!ownerId) throw new Error("Sem sessão/ownerId.");

    const usuarios = await listarUsuariosDoAdminIndexedDb();

    const payload = (usuarios || [])
      .map((u) => {
        const id = String(u?.id || "").trim();
        if (!id) return null;

        const telegram_user_id = String(u?.telegramUserId || u?.telegramId || "").trim();
        const telegram_username = String(u?.telegramUsername || "").trim().replace(/^@/, "");

        return {
          id,
          owner_id: ownerId,
          parte: normalizarParte(u?.parte || ""),
          telegram_user_id,
          telegram_username,
          criado_em: u?.criadoEm ? new Date(Number(u.criadoEm)).toISOString() : new Date().toISOString(),
        };
      })
      .filter(Boolean);

    if (!payload.length) return { ok: true, enviados: 0 };

    const { error } = await supabase.from("usuarios").upsert(payload, { onConflict: "id" });
    if (error) throw error;

    console.log("SYNC usuários OK:", payload.length);
    return { ok: true, enviados: payload.length };
  } catch (e) {
    console.error("Erro ao sincronizar usuários (IndexedDB -> Supabase):", e?.message || e);
    return { ok: false, enviados: 0, error: e?.message || String(e) };
  }
}

// ============================
// Próximo P automático
// ============================
function gerarProximoP(usuarios) {
  const usados = new Set(
    (usuarios || [])
      .map((u) => normalizarParte(u.parte))
      .filter(Boolean)
      .map((p) => parseInt(p.replace(/\D/g, ""), 10))
      .filter((n) => Number.isFinite(n) && n > 0)
  );

  for (let i = 1; i <= 99; i++) {
    if (!usados.has(i)) return `P${String(i).padStart(2, "0")}`;
  }
  return `P${String((usuarios?.length || 0) + 1).padStart(2, "0")}`;
}

// ============================
// Modal confirmação
// ============================
function abrirModalConfirmacao(texto, aoConfirmar) {
  if (mensagemConfirmacao) mensagemConfirmacao.textContent = texto;
  acaoConfirmacao = typeof aoConfirmar === "function" ? aoConfirmar : null;
  if (overlayConfirmacao) overlayConfirmacao.hidden = false;
}
function fecharModalConfirmacao() {
  if (overlayConfirmacao) overlayConfirmacao.hidden = true;
  acaoConfirmacao = null;
}
botaoCancelar?.addEventListener("click", fecharModalConfirmacao);
botaoConfirmar?.addEventListener("click", () => {
  if (acaoConfirmacao) acaoConfirmacao();
  fecharModalConfirmacao();
});
overlayConfirmacao?.addEventListener("click", (e) => {
  if (e.target === overlayConfirmacao) fecharModalConfirmacao();
});

// ============================
// Seleção (checkbox)
// ============================
function pegarCheckboxesLinha() {
  return Array.from(corpoTabelaUsuarios.querySelectorAll(".checkboxLinha"));
}
function pegarCheckboxesVisiveis() {
  return Array.from(corpoTabelaUsuarios.querySelectorAll("tr"))
    .filter((tr) => tr.style.display !== "none")
    .map((tr) => tr.querySelector(".checkboxLinha"))
    .filter(Boolean);
}
function pegarIdsSelecionados() {
  return pegarCheckboxesLinha()
    .filter((cb) => cb.checked)
    .map((cb) => String(cb.dataset.usuarioId));
}
function atualizarUiSelecao() {
  const checkboxes = pegarCheckboxesLinha();
  const marcados = checkboxes.filter((c) => c.checked).length;

  if (textoSelecionados) textoSelecionados.textContent = `${marcados} selecionado(s)`;
  if (botaoDesvincularSelecionados) botaoDesvincularSelecionados.disabled = marcados === 0;

  const visiveis = pegarCheckboxesVisiveis();
  const marcadosVisiveis = visiveis.filter((c) => c.checked).length;

  if (!checkboxSelecionarTodos) return;

  if (visiveis.length === 0) {
    checkboxSelecionarTodos.checked = false;
    checkboxSelecionarTodos.indeterminate = false;
    return;
  }

  checkboxSelecionarTodos.checked = marcadosVisiveis === visiveis.length;
  checkboxSelecionarTodos.indeterminate = marcadosVisiveis > 0 && marcadosVisiveis < visiveis.length;
}

// ============================
// Filtros
// ============================
function aplicarFiltros() {
  const termo = (inputBusca?.value || "").trim().toLowerCase();
  const status = (selectStatus?.value || "").trim().toLowerCase();

  const linhas = Array.from(corpoTabelaUsuarios.querySelectorAll("tr"));
  let visiveis = 0;

  for (const linha of linhas) {
    const tgId = (linha.querySelector("[data-col='telegramId']")?.dataset?.telegramid || "").toLowerCase();
    const tgUser = linha.querySelector("[data-col='telegramUser'] input")?.value?.toLowerCase() || "";
    const statusLinha = linha.querySelector("[data-col='status']")?.dataset?.status || "";

    let mostrar = true;
    if (termo && !(tgId.includes(termo) || tgUser.includes(termo))) mostrar = false;
    if (status && statusLinha !== status) mostrar = false;

    linha.style.display = mostrar ? "" : "none";
    if (mostrar) visiveis += 1;
  }

  if (estadoVazio) estadoVazio.hidden = visiveis !== 0;
  atualizarUiSelecao();
}
inputBusca?.addEventListener("input", aplicarFiltros);
selectStatus?.addEventListener("change", aplicarFiltros);

// ============================
// Parte SELECT
// ============================
function montarOpcoesParte(usuarios, parteAtual) {
  const atual = normalizarParte(parteAtual) || "";
  const usados = new Set((usuarios || []).map((u) => normalizarParte(u.parte)).filter(Boolean));

  const frag = document.createDocumentFragment();

  const optVazio = document.createElement("option");
  optVazio.value = "";
  optVazio.textContent = "—";
  frag.appendChild(optVazio);

  for (let i = 1; i <= 99; i++) {
    const p = `P${String(i).padStart(2, "0")}`;
    const opt = document.createElement("option");
    opt.value = p;
    opt.textContent = p;

    if (usados.has(p) && p !== atual) opt.disabled = true;
    frag.appendChild(opt);
  }

  return frag;
}

// ============================
// Render tabela
// ============================
function renderizarTabela(usuarios) {
  corpoTabelaUsuarios.innerHTML = "";

  if (!usuarios.length) {
    if (estadoVazio) estadoVazio.hidden = false;
    atualizarUiSelecao();
    return;
  }

  if (estadoVazio) estadoVazio.hidden = true;

  for (const u of usuarios) {
    const tr = document.createElement("tr");

    const tgId = u.telegramUserId ? String(u.telegramUserId) : "";
    const tgUser = u.telegramUsername ? String(u.telegramUsername) : "";

    tr.innerHTML = `
      <td>
        <input type="checkbox" class="checkboxLinha" data-usuario-id="${u.id}" />
      </td>

      <td>
        <span class="${classeStatus(u.status)}" data-col="status" data-status="${u.status}">
          ${u.status}
        </span>
      </td>

      <td data-col="parte">
        <select class="selectParte" data-usuario-id="${u.id}"></select>
      </td>

      <td data-col="telegramId" data-telegramid="${tgId}">
        <span class="textoFixo">${tgId || "—"}</span>
      </td>

      <td data-col="telegramUser">
        <input class="inputTelegramUser inputPadrao" data-usuario-id="${u.id}" placeholder="@usuario" value="${tgUser}" />
      </td>

      <td>${u.dataTexto || ""}</td>

      <td>
        <div class="acoesLinha">
          <button class="botaoAcao" data-acao="toggle" data-usuario-id="${u.id}">
            ${String(u.status).toLowerCase() === "bloqueado" ? "✅ Ativar" : "⛔ Bloquear"}
          </button>
          <button class="botaoAcao" data-acao="desvincular" data-usuario-id="${u.id}">🗑 Desvincular</button>
        </div>
      </td>
    `;

    corpoTabelaUsuarios.appendChild(tr);

    const sel = tr.querySelector(".selectParte");
    sel.appendChild(montarOpcoesParte(usuarios, u.parte));
    sel.value = normalizarParte(u.parte) || "";
  }

  atualizarUiSelecao();
}

// ============================
// Eventos da tabela
// ============================
let eventosTabelaOk = false;
function vincularEventosTabelaUmaVez() {
  if (eventosTabelaOk) return;

  // cliques (ações)
  corpoTabelaUsuarios.addEventListener("click", async (e) => {
    const botao = e.target.closest("button");
    if (!botao) return;

    const acao = String(botao.dataset.acao || "");
    const usuarioId = String(botao.dataset.usuarioId || "");
    if (!usuarioId) return;

    const usuarios = await listarUsuariosSupabase();
    const usuario = usuarios.find((x) => String(x.id) === usuarioId);
    if (!usuario) return;

    if (acao === "toggle") {
      const vaiAtivar = String(usuario.status).toLowerCase() === "bloqueado";
      const novoStatus = vaiAtivar ? "ativo" : "bloqueado";

      abrirModalConfirmacao(`Deseja ${vaiAtivar ? "ativar" : "bloquear"} este usuário?`, async () => {
        try {
          await atualizarStatusUsuarioSupabase(usuarioId, novoStatus);
          try {
            await salvarUsuarioIndexedDb({ ...usuario, status: novoStatus });
          } catch {}
          await atualizar();
        } catch (err) {
          console.error(err);
          aviso(`Falha ao atualizar status: ${err?.message || err}`);
        }
      });
      return;
    }

    if (acao === "desvincular") {
      abrirModalConfirmacao("Desvincular (apagar) este usuário? Esta ação remove do banco e da tela.", async () => {
        try {
          await excluirUsuariosSupabase([usuarioId]);
          try {
            await excluirUsuariosIndexedDb([usuarioId]);
          } catch {}
          await atualizar();
        } catch (err) {
          console.error(err);
          aviso(`Falha ao apagar usuário: ${err?.message || err}`);
        }
      });
      return;
    }
  });

  // checkbox
  corpoTabelaUsuarios.addEventListener("change", (e) => {
    if (e.target.closest(".checkboxLinha")) atualizarUiSelecao();
  });

  // mudança de Parte via SELECT
  corpoTabelaUsuarios.addEventListener("change", async (e) => {
    const sel = e.target.closest(".selectParte");
    if (!sel) return;

    const usuarioId = String(sel.dataset.usuarioId || "");
    const novo = normalizarParte(sel.value);

    const usuarios = await listarUsuariosSupabase();
    const usuario = usuarios.find((u) => String(u.id) === usuarioId);
    if (!usuario) return;

    if (!novo) {
      usuario.parte = "";
      try {
        await salvarUsuarioSupabase(usuario);
      } catch (err) {
        console.error(err);
        aviso(`Falha ao salvar parte: ${err?.message || err}`);
        return;
      }
      try {
        await salvarUsuarioIndexedDb(usuario);
      } catch {}
      await atualizar();
      return;
    }

    const repetido = usuarios.some((u) => String(u.id) !== usuarioId && normalizarParte(u.parte) === novo);
    if (repetido) {
      aviso(`Já existe outro usuário com ${novo}. Escolha outro P.`);
      sel.value = normalizarParte(usuario.parte) || "";
      return;
    }

    usuario.parte = novo;
    try {
      await salvarUsuarioSupabase(usuario);
    } catch (err) {
      console.error(err);
      aviso(`Falha ao salvar parte: ${err?.message || err}`);
      return;
    }
    try {
      await salvarUsuarioIndexedDb(usuario);
    } catch {}
    await atualizar();
  });

  // edição inline: Telegram Username
  corpoTabelaUsuarios.addEventListener(
    "blur",
    async (e) => {
      const inpTgUser = e.target.closest(".inputTelegramUser");
      if (!inpTgUser) return;

      const usuarioId = String(inpTgUser.dataset.usuarioId || "");
      if (!usuarioId) return;

      const usuarios = await listarUsuariosSupabase();
      const usuario = usuarios.find((u) => String(u.id) === usuarioId);
      if (!usuario) return;

      const novoUser = normalizarTelegramUser(inpTgUser.value);
      usuario.telegramUsername = novoUser || "";

      try {
        await salvarUsuarioSupabase(usuario);
      } catch (err) {
        console.error(err);
        aviso(`Falha ao salvar @: ${err?.message || err}`);
        return;
      }

      try {
        await salvarUsuarioIndexedDb(usuario);
      } catch {}

      inpTgUser.value = usuario.telegramUsername ? String(usuario.telegramUsername) : "";
    },
    true
  );

  eventosTabelaOk = true;
}

// ============================
// Ações topo (criar usuário com P automático)
// ============================
formVincular?.addEventListener("submit", async (e) => {
  e.preventDefault();

  const tgId = normalizarTelegramId(inputTelegramId?.value || "");
  const tgUser = normalizarTelegramUser(inputTelegramUser?.value || "");

  if (!tgId) {
    aviso("Informe o Telegram ID (obrigatório para o bot enviar no privado).");
    return;
  }

  const usuariosExistentes = await listarUsuariosSupabase();
  const repetidoTg = usuariosExistentes.some((u) => String(u.telegramUserId || "") === tgId);
  if (repetidoTg) {
    aviso(`Já existe outro usuário com esse Telegram ID (${tgId}).`);
    return;
  }

  const parteAutomatica = gerarProximoP(usuariosExistentes);

  const usuario = {
    id: gerarIdUsuario(),
    status: "ativo",
    criadoEm: Date.now(),
    dataTexto: formatarDataBr(),
    parte: parteAutomatica,
    telegramUserId: tgId,
    telegramUsername: tgUser || "",
  };

  try {
    await salvarUsuarioSupabase(usuario);
  } catch (err) {
    console.error(err);
    aviso(`Falha ao criar usuário no Supabase: ${err?.message || err}`);
    return;
  }

  try {
    await salvarUsuarioIndexedDb(usuario);
  } catch {}

  if (inputTelegramId) inputTelegramId.value = "";
  if (inputTelegramUser) inputTelegramUser.value = "";

  await atualizar();
});

checkboxSelecionarTodos?.addEventListener("change", () => {
  const visiveis = pegarCheckboxesVisiveis();
  for (const cb of visiveis) cb.checked = checkboxSelecionarTodos.checked;
  atualizarUiSelecao();
});

botaoDesvincularSelecionados?.addEventListener("click", async () => {
  const ids = pegarIdsSelecionados();
  if (!ids.length) return;

  abrirModalConfirmacao(
    `Desvincular (apagar) ${ids.length} usuário(s) selecionado(s)? Esta ação remove do banco e da tela.`,
    async () => {
      try {
        await excluirUsuariosSupabase(ids);
        try {
          await excluirUsuariosIndexedDb(ids);
        } catch {}
        await atualizar();
      } catch (err) {
        console.error(err);
        aviso(`Falha ao apagar selecionados: ${err?.message || err}`);
      }
    }
  );
});

// ============================
// Resumo
// ============================
function atualizarResumo(usuarios) {
  if (qtdUsuarios) qtdUsuarios.textContent = String(usuarios.length);
  if (qtdAtivos) qtdAtivos.textContent = String(usuarios.filter((u) => String(u.status).toLowerCase() === "ativo").length);
}

// ============================
// Atualizar
// ============================
async function atualizar() {
  const admin = obterAdminAtual();
  if (adminNome) adminNome.textContent = admin.nome;
  if (adminEmail) adminEmail.textContent = admin.email;

  // 1) sync local -> supabase (traz os “sumidos”)
  const r = await syncUsuariosIndexedDbParaSupabase();
  if (r?.ok) console.log(`SYNC usuários: ${r.enviados} upsert(s) para Supabase`);
  else console.warn("SYNC usuários falhou:", r);

  // 2) lista supabase (fonte de verdade)
  const usuarios = await listarUsuariosSupabase();

  vincularEventosTabelaUmaVez();
  renderizarTabela(usuarios);
  atualizarResumo(usuarios);
  aplicarFiltros();
}

// ============================
// BOOT
// ============================
function boot() {
  garantirAdmin();
  atualizar().catch((e) => {
    console.error("Falha no boot administrar:", e);
    aviso(`Erro ao iniciar: ${e?.message || e}`);
  });
}

if (document.readyState === "loading") {
  window.addEventListener("DOMContentLoaded", boot);
} else {
  boot();
}