async function carregarComponente(id, caminho) {
  const resp = await fetch(caminho);
  if (!resp.ok) throw new Error(`Erro ao carregar ${caminho}: ${resp.status}`);
  const html = await resp.text();

  const el = document.getElementById(id);
  if (!el) throw new Error(`Elemento #${id} não existe nesta página`);
  el.innerHTML = html;
}

function carregarScriptUmaVez(src, { module = false } = {}) {
  return new Promise((resolve, reject) => {
    const seletor = module
      ? `script[type="module"][src="${src}"]`
      : `script[src="${src}"]`;

    if (document.querySelector(seletor)) {
      resolve();
      return;
    }

    const script = document.createElement("script");
    script.src = src;

    if (module) {
      script.type = "module";
    } else {
      script.defer = true;
    }

    script.onload = resolve;
    script.onerror = () => reject(new Error(`Falha ao carregar script: ${src}`));
    document.head.appendChild(script);
  });
}

function atualizarAvatarHeader() {
  const avatarHeader = document.getElementById("profileAvatar");
  if (!avatarHeader) return;

  const avatarSalvo = localStorage.getItem("profileAvatarBase64");
  avatarHeader.src = avatarSalvo || "/0-imgs/user.svg";
}

async function bootHeaderFooter() {
  try {
    await carregarComponente("header", "/2-header-footer/header/header.html");
    await carregarComponente("footer", "/2-header-footer/footer/footer.html");

    atualizarAvatarHeader();

    // ✅ header.js agora é módulo
    await carregarScriptUmaVez("/2-header-footer/header/header.js", { module: true });

    // footer.js pode continuar comum
    await carregarScriptUmaVez("/2-header-footer/footer/footer.js");

    // se o header.js expuser a função globalmente, chama
    if (typeof window.inicializarHeader === "function") {
      window.inicializarHeader();
    }
  } catch (e) {
    console.error("Erro ao carregar header/footer:", e);
  }
}

if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", bootHeaderFooter, { once: true });
} else {
  bootHeaderFooter();
}