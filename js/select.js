/** Select buscable */
window.APP = window.APP || {};

APP.PreciseSelect = (() => {
  let overlay, titleEl, searchEl, listEl, onSelect;
  let manualEl, dniEl, nombreEl, useEl, hintEl;
  let allowManual = false;
  let findKnown = null;
  let bound = false;

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
          <svg class="precise-search-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" aria-hidden="true">
            <circle cx="11" cy="11" r="7"/><path d="M20 20l-3.5-3.5"/>
          </svg>
          <input type="search" id="ps-search" placeholder="Buscar DNI o nombre…" autocomplete="off" inputmode="search" />
        </div>
        <div class="precise-list" id="ps-list"></div>
        <div class="precise-manual" id="ps-manual" hidden>
          <p class="precise-manual-hint" id="ps-hint">No está en la lista. Ingresa DNI y nombre.</p>
          <label class="precise-manual-field">
            <small>DNI</small>
            <input id="ps-dni" inputmode="numeric" maxlength="8" autocomplete="off" />
          </label>
          <label class="precise-manual-field">
            <small>Nombres</small>
            <input id="ps-nombre" autocomplete="off" autocapitalize="characters" disabled />
          </label>
          <button type="button" class="btn-main" id="ps-use" disabled>Usar</button>
        </div>
      </div>`;
    document.body.appendChild(overlay);
    titleEl = overlay.querySelector("#ps-title");
    searchEl = overlay.querySelector("#ps-search");
    listEl = overlay.querySelector("#ps-list");
    manualEl = overlay.querySelector("#ps-manual");
    hintEl = overlay.querySelector("#ps-hint");
    dniEl = overlay.querySelector("#ps-dni");
    nombreEl = overlay.querySelector("#ps-nombre");
    useEl = overlay.querySelector("#ps-use");
    overlay.querySelector("#ps-close").onclick = close;
    overlay.addEventListener("click", (e) => {
      if (e.target === overlay) close();
    });
    let searchTick = 0;
    searchEl.addEventListener("input", () => {
      const q = searchEl.value;
      if (searchTick) clearTimeout(searchTick);
      // DNI: responde al toque. Nombre: 60ms para no trabar el teclado.
      const digits = String(q || "").replace(/\D/g, "");
      const delay = digits.length >= 1 && digits === String(q || "").replace(/\s/g, "") ? 0 : 60;
      searchTick = setTimeout(() => {
        searchTick = 0;
        paint(q);
      }, delay);
    });
    bindManual();
  }

  function bindManual() {
    if (bound) return;
    bound = true;
    dniEl.addEventListener("input", () => {
      dniEl.value = String(dniEl.value || "").replace(/\D/g, "").slice(0, 8);
      syncManualFields();
    });
    nombreEl.addEventListener("input", () => {
      const cur = nombreEl.selectionStart;
      nombreEl.value = APP.Data.cleanNombre(nombreEl.value);
      try {
        nombreEl.setSelectionRange(cur, cur);
      } catch (e) {}
      syncManualFields();
    });
    useEl.addEventListener("click", () => {
      const dni = APP.Data.cleanDni(dniEl.value);
      const nombre = APP.Data.cleanNombre(nombreEl.value);
      if (dni.length !== 8 || nombre.length < 3) return;
      const added = APP.Data.addScanner(dni, nombre);
      close();
      if (onSelect && added) onSelect({ id: added.dni, dni: added.dni, nombre: added.nombre, label: added.nombre });
    });
  }

  function syncManualFields() {
    const dni = APP.Data.cleanDni(dniEl.value);
    const known = dni.length === 8 && findKnown && findKnown(dni);
    const nameOn = allowManual && dni.length === 8 && !known;
    nombreEl.disabled = !nameOn;
    if (!nameOn && !nombreEl.value) nombreEl.value = "";
    if (known) {
      hintEl.textContent = "Ese DNI ya está en la lista.";
    } else if (dni.length === 8) {
      hintEl.textContent = "No está en la lista. Escribe el nombre en mayúsculas.";
    } else {
      hintEl.textContent = "No está en la lista. Ingresa DNI de 8 dígitos y nombre.";
    }
    const nombre = APP.Data.cleanNombre(nombreEl.value);
    useEl.disabled = !(nameOn && nombre.length >= 3);
  }

  function showManual(q, opts) {
    if (!allowManual) {
      manualEl.hidden = true;
      return;
    }
    const digits = APP.Data.cleanDni(q);
    const typed = String(q || "").trim();
    const known = digits.length === 8 && findKnown && findKnown(digits);
    const need = !known && (digits.length === 8 || (!opts.length && typed));
    manualEl.hidden = !need;
    if (!need) return;
    if (digits.length) dniEl.value = digits;
    syncManualFields();
    if (digits.length === 8 && !known && !nombreEl.value) {
      nombreEl.focus();
    }
  }

  let getOptions = () => [];
  let emptyText = "Sin resultados";

  function paint(q) {
    const opts = getOptions(q) || [];
    listEl.innerHTML = opts.length
      ? opts
          .map((o) => {
            const metaRaw = String(o.meta || "").trim();
            const meta =
              metaRaw && /^\d{6,8}$/.test(metaRaw.replace(/\D/g, "")) && metaRaw.replace(/\D/g, "").length === 8
                ? "DNI - " + metaRaw.replace(/\D/g, "")
                : metaRaw;
            return `<button type="button" class="precise-opt${o.marked ? " marked" : ""}" data-id="${String(o.id).replace(/"/g, "")}">
          <strong>${o.label}</strong>${o.note ? `<small class="precise-note">${o.note}</small>` : ""}${meta ? `<small>${meta}</small>` : ""}
        </button>`;
          })
          .join("")
      : `<p class="empty">Sin resultados</p>`;
    listEl.querySelectorAll(".precise-opt").forEach((btn) => {
      btn.onclick = () => {
        const opt = opts.find((o) => String(o.id) === btn.dataset.id);
        close();
        if (opt && onSelect) onSelect(opt);
      };
    });
    showManual(q, opts);
  }

  function open(cfg) {
    ensure();
    titleEl.textContent = cfg.title || "Buscar";
    getOptions = cfg.getOptions || (() => []);
    onSelect = cfg.onSelect || null;
    emptyText = cfg.empty || "Sin resultados";
    allowManual = !!cfg.allowManual;
    findKnown = cfg.findKnown || null;
    searchEl.placeholder = cfg.placeholder || "Buscar DNI o nombre…";
    searchEl.value = "";
    dniEl.value = "";
    nombreEl.value = "";
    nombreEl.disabled = true;
    useEl.disabled = true;
    manualEl.hidden = true;
    paint("");
    overlay.classList.add("open");
    searchEl.focus();
  }

  function close() {
    if (overlay) overlay.classList.remove("open");
  }

  return { open, close };
})();
