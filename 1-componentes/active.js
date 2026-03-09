// active.js - Controla os temas das páginas aplicando classes .active onde necessário

function aplicarTemaAtivo() {
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
}

function inicializarTemaAtivo() {
  aplicarTemaAtivo();

  window.addEventListener("storage", (e) => {
    if (e.key === "tema") {
      aplicarTemaAtivo();
    }
  });
}

document.addEventListener("DOMContentLoaded", inicializarTemaAtivo);

export { aplicarTemaAtivo };