/** Select buscable */
window.APP = window.APP || {};

APP.PreciseSelect = (() => {
  let overlay, titleEl, searchEl, listEl, onSelect;

  function ensure() {
    if (overlay) return;
    overlay = document.createElement("div");
    overlay.className = "overlay";
    overlay.innerHTML = `
      <div class="precise-modal">
        <div class="precise-modal-head">
          <h3 id="ps-title">Buscar</h3>
          <button type="button" id="ps-close" aria-label="Cerrar">✕</button>
        </div>
        <div class="precise-search">
          <input type="search" id="ps-search" placeholder="Buscar..." autocomplete="off" />
        </div>
        <div class="precise-list" id="ps-list"></div>
      </div>`;
    document.body.appendChild(overlay);
    titleEl = overlay.querySelector("#ps-title");
    searchEl = overlay.querySelector("#ps-search");
    listEl = overlay.querySelector("#ps-list");
    overlay.querySelector("#ps-close").onclick = close;
    overlay.addEventListener("click", (e) => {
      if (e.target === overlay) close();
    });
    searchEl.addEventListener("input", () => paint(searchEl.value));
  }

  let getOptions = () => [];

  function paint(q) {
    const opts = getOptions(q) || [];
    listEl.innerHTML = opts.length
      ? opts
          .map(
            (o) => `<button type="button" class="precise-opt" data-id="${String(o.id).replace(/"/g, "")}">
          <strong>${o.label}</strong>${o.meta ? `<small>${o.meta}</small>` : ""}
        </button>`
          )
          .join("")
      : `<p class="empty" style="padding:12px">Sin resultados</p>`;
    listEl.querySelectorAll(".precise-opt").forEach((btn) => {
      btn.onclick = () => {
        const opt = opts.find((o) => String(o.id) === btn.dataset.id);
        if (opt && onSelect) onSelect(opt);
        close();
      };
    });
  }

  function open(cfg) {
    ensure();
    titleEl.textContent = cfg.title || "Buscar";
    getOptions = cfg.getOptions || (() => []);
    onSelect = cfg.onSelect || null;
    searchEl.value = "";
    paint("");
    overlay.classList.add("open");
    searchEl.focus();
  }

  function close() {
    if (overlay) overlay.classList.remove("open");
  }

  return { open, close };
})();
