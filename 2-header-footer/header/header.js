import { logout } from "../../8-src/auth.js";

const LS_AVATAR_KEY = "profileAvatarBase64";
const DEFAULT_AVATAR = "/0-imgs/user.svg";

function aplicarAvatarNoHeader() {
  const img = document.getElementById("profileAvatar");
  if (!img) return;

  const saved = localStorage.getItem(LS_AVATAR_KEY);
  img.src = saved || DEFAULT_AVATAR;
}

function escutarMudancasDeAvatar() {
  if (window.__ibespeAvatarListenerReady) return;
  window.__ibespeAvatarListenerReady = true;

  window.addEventListener("ibespe:avatarUpdated", aplicarAvatarNoHeader);

  window.addEventListener("storage", (e) => {
    if (e.key === LS_AVATAR_KEY) aplicarAvatarNoHeader();
  });
}

function destacarMenuAtivo() {
  const paginaAtual = location.pathname
    .split("/")
    .pop()
    .split(".")[0]
    .toLowerCase();

  document.querySelectorAll(".navLink[dataRota]").forEach((link) => {
    link.classList.remove("active");

    const rota = (link.getAttribute("dataRota") || "").toLowerCase();
    if (rota === paginaAtual) {
      link.classList.add("active");
      link.setAttribute("aria-current", "page");
    } else {
      link.removeAttribute("aria-current");
    }
  });
}

function fecharMenuPerfil(profileMenu, profileToggle) {
  if (!profileMenu || !profileToggle) return;
  profileMenu.classList.remove("open");
  profileMenu.setAttribute("aria-hidden", "true");
  profileToggle.setAttribute("aria-expanded", "false");
}

function abrirMenuPerfil(profileMenu, profileToggle) {
  if (!profileMenu || !profileToggle) return;
  profileMenu.classList.add("open");
  profileMenu.setAttribute("aria-hidden", "false");
  profileToggle.setAttribute("aria-expanded", "true");
}

function alternarMenuPerfil(profileMenu, profileToggle) {
  if (!profileMenu || !profileToggle) return;

  const aberto = profileMenu.classList.contains("open");
  if (aberto) {
    fecharMenuPerfil(profileMenu, profileToggle);
  } else {
    abrirMenuPerfil(profileMenu, profileToggle);
  }
}

function inicializarMenuPerfil() {
  const profileToggle = document.getElementById("profileToggle");
  const profileMenu = document.getElementById("profileMenu");
  const profileLogout = document.getElementById("profileLogout");

  if (!profileToggle || !profileMenu) return;

  // evita bind duplicado no mesmo header
  if (profileToggle.dataset.bound === "1") return;
  profileToggle.dataset.bound = "1";

  profileToggle.addEventListener("click", (e) => {
    e.preventDefault();
    e.stopPropagation();
    alternarMenuPerfil(profileMenu, profileToggle);
  });

  profileMenu.addEventListener("click", (e) => {
    e.stopPropagation();
  });

  profileLogout?.addEventListener("click", async (e) => {
    e.preventDefault();
    e.stopPropagation();

    try {
      await logout();
    } catch (err) {
      console.error("Erro ao sair:", err);
      window.location.href = "/7-login/login.html";
    }
  });

  document.addEventListener("click", (e) => {
    if (!profileMenu.contains(e.target) && !profileToggle.contains(e.target)) {
      fecharMenuPerfil(profileMenu, profileToggle);
    }
  });

  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") {
      fecharMenuPerfil(profileMenu, profileToggle);
    }
  });
}

function aplicarTema() {
  const tema = localStorage.getItem("tema") || "claro";
  const body = document.body;
  const root = document.documentElement;

  if (tema === "claro") {
    body.classList.add("active");
    root.classList.add("active");
  } else {
    body.classList.remove("active");
    root.classList.remove("active");
  }

  atualizarEstadoBotaoTema();
}

function alternarTema() {
  const temaAtual = localStorage.getItem("tema") || "claro";
  const novoTema = temaAtual === "escuro" ? "claro" : "escuro";
  localStorage.setItem("tema", novoTema);
  aplicarTema();
}

function atualizarEstadoBotaoTema() {
  const botao = document.getElementById("themeToggle");
  if (!botao) return;
  const tema = localStorage.getItem("tema") || "claro";
  if (tema === "escuro") {
    botao.classList.add("active");
  } else {
    botao.classList.remove("active");
  }
}

function inicializarTema() {
  aplicarTema();
  atualizarEstadoBotaoTema();
  const themeToggle = document.getElementById("themeToggle");
  if (themeToggle) {
    themeToggle.addEventListener("click", (e) => {
      e.preventDefault();
      e.stopPropagation();
      alternarTema();
    });
  }
}

function inicializarHeader() {
  destacarMenuAtivo();
  aplicarAvatarNoHeader();
  escutarMudancasDeAvatar();
  inicializarMenuPerfil();
  inicializarTema();
}

window.inicializarHeader = inicializarHeader;
export { inicializarHeader };