import { requireAuth, logout } from "../8-src/auth.js";

const user = await requireAuth({ redirectTo: "../7-login/login.html" });
if (!user) throw new Error("Sem sessão");

document.querySelector("#btnLogout")?.addEventListener("click", () => logout());

(() => {
  const ownerId = String(user?.id || "").trim();

  const LS_KEYS = {
    LAST_DASH_VIEW: "ibespe:lastDashboardView"
  };

  const $ = (id) => document.getElementById(id);

  function getLastDashKey() {
    return `${LS_KEYS.LAST_DASH_VIEW}:${ownerId}`;
  }

  function safeJsonParse(raw, fallback) {
    try {
      if (!raw) return fallback;
      return JSON.parse(raw);
    } catch {
      return fallback;
    }
  }

  function formatDateTimeBR(value) {
    if (!value) return "—";
    const d = new Date(value);
    if (Number.isNaN(d.getTime())) return "—";
    const pad = (n) => String(n).padStart(2, "0");
    return `${pad(d.getDate())}/${pad(d.getMonth() + 1)}/${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
  }

  function formatRelative(value) {
    if (!value) return "—";
    const d = new Date(value);
    if (Number.isNaN(d.getTime())) return "—";

    const diffMs = Date.now() - d.getTime();
    if (diffMs < 0) return "agora há pouco";

    const sec = Math.floor(diffMs / 1000);
    const min = Math.floor(sec / 60);
    const hr = Math.floor(min / 60);
    const day = Math.floor(hr / 24);

    if (sec < 45) return "agora há pouco";
    if (min < 60) return `há ${min} min`;
    if (hr < 24) return `há ${hr} h`;
    return `há ${day} dia(s)`;
  }

  function getLastDashboardView() {
    return safeJsonParse(localStorage.getItem(getLastDashKey()), null);
  }

  function renderLastDashboard(lastDash) {
    const elMain = $("ultimaDashboard");
    const elDet = $("ultimaDashboardDetalhe");
    if (!elMain || !elDet) return;

    if (!lastDash || !lastDash.at) {
      elMain.textContent = "—";
      elDet.textContent = "Abra um arquivo na dashboard para registrar.";
      return;
    }

    elMain.textContent = `${formatDateTimeBR(lastDash.at)} (${formatRelative(lastDash.at)})`;

    const nome = lastDash.fileName || "arquivo";
    elDet.textContent = `Último arquivo aberto: ${nome}`;
  }

  async function init() {
    const lastDash = getLastDashboardView();
    renderLastDashboard(lastDash);
  }

  function boot() {
    init().catch((e) => {
      console.error("Falha ao iniciar home:", e);
    });
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", boot, { once: true });
  } else {
    boot();
  }
})();