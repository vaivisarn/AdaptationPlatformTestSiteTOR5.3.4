/* Adaptation Platform Library — Card library view.
   - Loads xlsx via SheetJS file picker (or shows DEMO_ROWS).
   - Highlights missing cells.
   - Click a card to open the full-detail drawer. */

(() => {

  // ===== Column schema (matches the spreadsheet) =====================

  const COLS = [
    { key: "product_id",          label: "Product ID",        th: "รหัส" },
    { key: "product_name",        label: "Product name",      th: "ชื่อผลิตภัณฑ์" },
    { key: "delivery_format",     label: "Delivery format",   th: "รูปแบบการให้บริการ" },
    { key: "service_type",        label: "Service type",      th: "ประเภทเนื้อหา",            multi: true },
    { key: "owner",               label: "Owner",             th: "เจ้าของ" },
    { key: "developer",           label: "Developer",         th: "ผู้พัฒนา" },
    { key: "geographic_scope",    label: "Geographic scope",  th: "ขอบเขตทางภูมิศาสตร์" },
    { key: "description",         label: "Description",       th: "คำอธิบาย" },
    { key: "group_of_use_case",   label: "Use-case group",    th: "หมวดใหญ่ของกรณีการใช้งาน", multi: true },
    { key: "use_case",            label: "Use case",          th: "กรณีการใช้งาน",            multi: true },
    { key: "sectors",             label: "Sectors (NAP)",     th: "สาขา",                     multi: true },
    { key: "url",                 label: "URL",               th: "url",                       url: true },
    { key: "GFCS Pillar Related", label: "GFCS pillar",       th: "สาขา GFCS ที่เกี่ยวข้อง",   multi: true },
    { key: "remarks",             label: "Remarks",           th: "หมายเหตุ" },
  ];

  const FILTERS = [
    { key: "delivery_format",     label: "Delivery format", single: true },
    { key: "geographic_scope",    label: "Geographic scope", single: true },
    { key: "service_type",        label: "Service type",     multi: true },
    { key: "group_of_use_case",   label: "Use-case group",   multi: true },
    { key: "sectors",             label: "Sectors (NAP)",    multi: true },
    { key: "GFCS Pillar Related", label: "GFCS pillar",      multi: true },
  ];

  // ===== State =======================================================

  // source: "bundled" (data/platforms.xlsx) | "demo" | "uploaded"
  const state = {
    rows: [],
    source: "bundled",
    sourceLabel: "loading…",
    bundledRows: null,         // cached after first fetch
    bundledLabel: null,
    bundledError: null,
    uploadedRows: null,        // last user-picked file
    uploadedLabel: null,
    filters: {},
    search: "",
    onlyGaps: false,
    sort: "gaps_desc",
    openId: null,
  };

  // Path to the bundled spreadsheet — committed alongside index.html.
  const BUNDLED_XLSX_PATH = "data/platforms.xlsx";

  // ===== Helpers =====================================================

  const isEmpty = (v) => v === undefined || v === null || String(v).trim() === "";
  const splitMulti = (v) => isEmpty(v) ? [] : String(v).split(";").map(s => s.trim()).filter(Boolean);

  const rowFilledCount = (row) => COLS.filter(c => !isEmpty(row[c.key])).length;
  const rowGapCount    = (row) => COLS.length - rowFilledCount(row);
  const rowHasGaps     = (row) => rowGapCount(row) > 0;

  const uniqueValues = (rows, key, multi) => {
    const set = new Set();
    rows.forEach(r => {
      if (multi) splitMulti(r[key]).forEach(v => set.add(v));
      else if (!isEmpty(r[key])) set.add(String(r[key]).trim());
    });
    return [...set].sort((a,b) => a.localeCompare(b));
  };

  const matchesFilters = (row) => {
    if (state.search) {
      const q = state.search.toLowerCase();
      const hay = [row.product_name, row.description, row.product_id, row.owner, row.developer]
        .map(v => (v||"").toString().toLowerCase()).join(" ");
      if (!hay.includes(q)) return false;
    }
    if (state.onlyGaps && !rowHasGaps(row)) return false;
    for (const f of FILTERS) {
      const sel = state.filters[f.key];
      if (!sel) continue;
      if (f.single) {
        if (sel && String(row[f.key] || "").trim() !== sel) return false;
      } else {
        if (sel.size === 0) continue;
        const vals = splitMulti(row[f.key]);
        let ok = false;
        for (const s of sel) if (vals.includes(s)) { ok = true; break; }
        if (!ok) return false;
      }
    }
    return true;
  };

  const sortRows = (rows) => {
    const arr = [...rows];
    const cmpStr = (a,b) => String(a||"").toLowerCase().localeCompare(String(b||"").toLowerCase());
    switch (state.sort) {
      case "gaps_desc":
        arr.sort((a,b) => rowGapCount(b) - rowGapCount(a) || cmpStr(a.product_id, b.product_id));
        break;
      case "gaps_asc":
        arr.sort((a,b) => rowGapCount(a) - rowGapCount(b) || cmpStr(a.product_id, b.product_id));
        break;
      case "id":
        arr.sort((a,b) => cmpStr(a.product_id, b.product_id));
        break;
      case "name":
        arr.sort((a,b) => cmpStr(a.product_name, b.product_name));
        break;
      case "owner":
        arr.sort((a,b) => cmpStr(a.owner, b.owner));
        break;
    }
    return arr;
  };

  const filteredRows = () => sortRows(state.rows.filter(matchesFilters));

  // ===== DOM helpers =================================================

  const $  = (sel, root=document) => root.querySelector(sel);
  const $$ = (sel, root=document) => [...root.querySelectorAll(sel)];
  const el = (tag, attrs={}, children=[]) => {
    const n = document.createElement(tag);
    for (const k in attrs) {
      if (k === "class") n.className = attrs[k];
      else if (k === "html") n.innerHTML = attrs[k];
      else if (k.startsWith("on")) n.addEventListener(k.slice(2), attrs[k]);
      else if (attrs[k] !== false && attrs[k] !== null && attrs[k] !== undefined) n.setAttribute(k, attrs[k]);
    }
    (Array.isArray(children) ? children : [children]).forEach(c => {
      if (c == null || c === false) return;
      n.appendChild(typeof c === "string" ? document.createTextNode(c) : c);
    });
    return n;
  };
  const missing = () => el("span", { class: "missing" });

  const activeFilterCount = () => {
    let n = 0;
    for (const f of FILTERS) {
      const sel = state.filters[f.key];
      if (!sel) continue;
      if (f.single) n += sel ? 1 : 0;
      else n += sel.size;
    }
    if (state.search) n += 1;
    if (state.onlyGaps) n += 1;
    return n;
  };

  // ===== XLSX loading ================================================
  // Real workbook layout (from the TOR template):
  //   sheet "Introduction"    — purpose/instructions, skip
  //   sheet "Products"        — actual records  ← use this one
  //   sheet "Ref_Lookups"     — controlled vocabs, skip
  //   sheet "Products_backup" — legacy structure, skip
  //
  // Header cells are bilingual two-liners, e.g.
  //   "รหัส\nproduct_id"     →  key = "product_id"
  // We normalise each header by splitting on newline and taking the
  // last non-empty piece (the English machine-name).

  const normaliseHeader = (raw) => {
    if (raw === undefined || raw === null) return "";
    const lines = String(raw).split(/\r?\n/).map(s => s.trim()).filter(Boolean);
    if (!lines.length) return "";
    // Prefer the line that looks like a machine-name (ascii, snake_case-ish)
    const ascii = lines.find(l => /^[A-Za-z][A-Za-z0-9_ ]*$/.test(l));
    return (ascii || lines[lines.length - 1]).trim();
  };

  const pickSheet = (wb) => {
    // Prefer a sheet named "Products"; fall back to the 2nd sheet, then 1st.
    const named = wb.SheetNames.find(n => /^products$/i.test(n.trim()));
    if (named) return { name: named, sheet: wb.Sheets[named] };
    if (wb.SheetNames.length >= 2) {
      const n = wb.SheetNames[1];
      return { name: n, sheet: wb.Sheets[n] };
    }
    const n = wb.SheetNames[0];
    return { name: n, sheet: wb.Sheets[n] };
  };

  const parseProductsSheet = (sheet) => {
    const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "", blankrows: false });
    if (!matrix.length) return [];
    const headers = matrix[0].map(normaliseHeader);
    const rows = [];
    for (let i = 1; i < matrix.length; i++) {
      const row = matrix[i];
      // Skip rows that are entirely empty
      if (!row || row.every(c => c === "" || c === null || c === undefined)) continue;
      const obj = {};
      headers.forEach((h, idx) => {
        if (!h) return;
        const v = row[idx];
        obj[h] = v === undefined || v === null ? "" : v;
      });
      // Skip rows with no product_id AND no product_name (defensive)
      if (isEmpty(obj.product_id) && isEmpty(obj.product_name)) continue;
      // Normalise product_id to string (sheet often stores it as a number)
      if (!isEmpty(obj.product_id)) obj.product_id = String(obj.product_id).trim();
      rows.push(obj);
    }
    return rows;
  };

  // Parse an ArrayBuffer/Uint8Array into { rows, sheetName }.
  const parseXlsxBuffer = (buf) => {
    const data = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
    const wb = XLSX.read(data, { type: "array" });
    const { name, sheet } = pickSheet(wb);
    if (!sheet) throw new Error("No usable sheet found");
    const rows = parseProductsSheet(sheet);
    if (!rows.length) throw new Error(`Sheet "${name}" has no data rows`);
    return { rows, sheetName: name };
  };

  // File-picker upload (overrides the bundled file for this session).
  const loadUploadedFile = (file) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const { rows, sheetName } = parseXlsxBuffer(e.target.result);
        state.uploadedRows  = rows;
        state.uploadedLabel = `${file.name} · sheet "${sheetName}" · ${rows.length} rows`;
        setSource("uploaded");
      } catch (err) {
        console.error(err);
        alert("Could not read file: " + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // Auto-load the spreadsheet that ships with the page.
  const loadBundled = async () => {
    try {
      const resp = await fetch(BUNDLED_XLSX_PATH, { cache: "no-cache" });
      if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
      const buf = await resp.arrayBuffer();
      const { rows, sheetName } = parseXlsxBuffer(buf);
      state.bundledRows  = rows;
      state.bundledLabel = `${BUNDLED_XLSX_PATH} · sheet "${sheetName}" · ${rows.length} rows`;
      state.bundledError = null;
    } catch (err) {
      console.warn("Bundled spreadsheet failed to load:", err);
      state.bundledError = err.message || String(err);
    }
  };

  // Switch the active data source and re-render.
  const setSource = (which) => {
    state.source = which;
    state.filters = {};
    state.openId  = null;
    if (which === "bundled") {
      if (state.bundledRows) {
        state.rows = state.bundledRows;
        state.sourceLabel = state.bundledLabel;
      } else {
        state.rows = [];
        state.sourceLabel = state.bundledError
          ? `couldn't load ${BUNDLED_XLSX_PATH} (${state.bundledError})`
          : "loading…";
      }
    } else if (which === "demo") {
      state.rows = window.DEMO_ROWS;
      state.sourceLabel = `demo data · ${window.DEMO_ROWS.length} rows`;
    } else if (which === "uploaded") {
      state.rows = state.uploadedRows || [];
      state.sourceLabel = state.uploadedLabel || "no file uploaded yet";
    }
    renderAll();
  };

  // ===== Top meta ====================================================

  const renderTopMeta = () => {
    const total = state.rows.length;
    const shown = state.rows.filter(matchesFilters).length;
    const gaps  = state.rows.filter(rowHasGaps).length;
    $("#meta").innerHTML =
      `<strong>${shown}</strong>/${total} shown · <strong>${gaps}</strong> with gaps · src: ${state.sourceLabel}`;
  };

  // ===== Filter bar ==================================================

  const renderFilters = () => {
    const bar = $("#filterbar");
    bar.innerHTML = "";
    FILTERS.forEach(f => {
      const group = el("div", { class: "filter-group" });
      group.appendChild(el("span", { class: "filter-label" }, f.label));
      const chipsRow = el("div", { class: "chips" });

      const values = uniqueValues(state.rows, f.key, f.multi);

      if (f.single) {
        const sel = state.filters[f.key] || null;
        const allChip = el("span", { class: "chip" + (sel === null ? " on" : "") }, "All");
        allChip.onclick = () => { delete state.filters[f.key]; renderAll(); };
        chipsRow.appendChild(allChip);
        values.forEach(v => {
          const c = el("span", { class: "chip" + (sel === v ? " on tone-accent" : "") }, v);
          c.onclick = () => { state.filters[f.key] = v; renderAll(); };
          chipsRow.appendChild(c);
        });
      } else {
        const sel = state.filters[f.key] || new Set();
        values.forEach(v => {
          const on = sel.has(v);
          const c = el("span", { class: "chip" + (on ? " on" : "") }, v);
          c.onclick = () => {
            const cur = state.filters[f.key] || new Set();
            if (cur.has(v)) cur.delete(v); else cur.add(v);
            if (cur.size === 0) delete state.filters[f.key];
            else state.filters[f.key] = cur;
            renderAll();
          };
          chipsRow.appendChild(c);
        });
        if (!values.length) chipsRow.appendChild(el("span", { class: "chip", style: "opacity:.4" }, "— no values —"));
      }

      group.appendChild(chipsRow);
      bar.appendChild(group);
    });

    // "Clear all filters" enabled state
    const n = activeFilterCount();
    const clearBtn = $("#clear-filters");
    clearBtn.disabled = n === 0;
    clearBtn.textContent = n > 0 ? `Clear all filters (${n})` : "Clear all filters";
  };

  // ===== Cards =======================================================

  const renderCards = () => {
    const root = $("#view-cards");
    root.innerHTML = "";
    const rows = filteredRows();
    if (!rows.length) {
      root.appendChild(el("div", { class: "empty-state" }, "No records match the current filters."));
      return;
    }

    const grid = el("div", { class: "cards" });
    rows.forEach(r => {
      const filled = rowFilledCount(r);
      const pct = Math.round(filled / COLS.length * 100);
      const fillClass = pct === 100 ? "ok" : (pct < 70 ? "warn" : "");
      const gaps = rowGapCount(r);

      const card = el("div", { class: "card", role: "button", tabindex: "0" });
      card.addEventListener("click", () => openDrawer(r.product_id));
      card.addEventListener("keydown", (e) => {
        if (e.key === "Enter" || e.key === " ") { e.preventDefault(); openDrawer(r.product_id); }
      });

      // head
      card.appendChild(el("div", { class: "card-head" }, [
        el("span", { class: "card-id" }, r.product_id || "—"),
        isEmpty(r.delivery_format) ? missing() : el("span", { class: "card-format" }, r.delivery_format),
      ]));

      // title
      const title = el("h3", { class: "card-title" });
      if (isEmpty(r.product_name)) title.appendChild(missing());
      else title.textContent = r.product_name;
      card.appendChild(title);

      // description
      const desc = el("p", { class: "card-desc" });
      if (isEmpty(r.description)) desc.appendChild(missing());
      else desc.textContent = r.description;
      card.appendChild(desc);

      // tag chips: service_type + scope
      const tagsWrap = el("div", { class: "card-tags" });
      splitMulti(r.service_type).forEach(v => tagsWrap.appendChild(el("span", { class: "tag-soft accent" }, v)));
      if (!isEmpty(r.geographic_scope)) tagsWrap.appendChild(el("span", { class: "tag-soft" }, r.geographic_scope));
      if (!tagsWrap.children.length) tagsWrap.appendChild(missing());
      card.appendChild(tagsWrap);

      // meta grid
      const meta = el("dl", { class: "card-meta" });
      const addMeta = (label, key, joinMulti=false) => {
        meta.appendChild(el("dt", {}, label));
        const dd = el("dd");
        if (joinMulti) {
          const vals = splitMulti(r[key]);
          if (!vals.length) dd.appendChild(missing());
          else dd.textContent = vals.join(", ");
        } else {
          if (isEmpty(r[key])) dd.appendChild(missing());
          else dd.textContent = r[key];
        }
        meta.appendChild(dd);
      };
      addMeta("Owner", "owner");
      addMeta("Developer", "developer");
      addMeta("Sectors", "sectors", true);
      addMeta("GFCS pillar", "GFCS Pillar Related", true);
      card.appendChild(meta);

      // completeness
      const comp = el("div", { class: "completeness" });
      comp.appendChild(el("span", {}, `${filled}/${COLS.length} fields`));
      const bar = el("div", { class: "completeness-bar" });
      bar.appendChild(el("div", { class: "completeness-fill " + fillClass, style: `width:${pct}%` }));
      comp.appendChild(bar);
      comp.appendChild(el("span", {}, gaps > 0 ? `${gaps} missing` : `complete`));
      card.appendChild(comp);

      grid.appendChild(card);
    });
    root.appendChild(grid);
  };

  // ===== Drawer (full detail) ========================================

  const openDrawer = (id) => {
    state.openId = id;
    renderDrawer();
    $("#drawer-overlay").classList.remove("hidden");
    requestAnimationFrame(() => {
      $("#drawer-overlay").classList.add("visible");
      $("#drawer").setAttribute("aria-hidden", "false");
    });
  };
  const closeDrawer = () => {
    state.openId = null;
    $("#drawer-overlay").classList.remove("visible");
    $("#drawer").setAttribute("aria-hidden", "true");
    setTimeout(() => $("#drawer-overlay").classList.add("hidden"), 200);
  };

  const renderDrawer = () => {
    const r = state.rows.find(x => x.product_id === state.openId);
    const drawer = $("#drawer");
    drawer.innerHTML = "";
    if (!r) return;

    const filled = rowFilledCount(r);
    const pct = Math.round(filled / COLS.length * 100);
    const gaps = rowGapCount(r);

    // header
    const head = el("div", { class: "drawer-head" });
    head.appendChild(el("div", { class: "drawer-head-row" }, [
      el("span", { class: "card-id" }, r.product_id || "—"),
      el("button", { class: "btn ghost drawer-close", onclick: closeDrawer, type: "button", "aria-label": "Close" }, "✕"),
    ]));
    const h2 = el("h2", { class: "drawer-title" });
    if (isEmpty(r.product_name)) h2.appendChild(missing()); else h2.textContent = r.product_name;
    head.appendChild(h2);

    const subParts = el("div", { class: "drawer-sub" });
    if (!isEmpty(r.delivery_format)) subParts.appendChild(el("span", { class: "tag-soft" }, r.delivery_format));
    else subParts.appendChild(missing());
    if (!isEmpty(r.geographic_scope)) subParts.appendChild(el("span", { class: "tag-soft" }, r.geographic_scope));
    else subParts.appendChild(missing());
    head.appendChild(subParts);

    // completeness summary
    const comp = el("div", { class: "drawer-completeness" });
    comp.appendChild(el("span", { class: "filter-label" }, `Metadata completeness — ${filled}/${COLS.length} fields`));
    const bar = el("div", { class: "completeness-bar" });
    const fc = pct === 100 ? "ok" : (pct < 70 ? "warn" : "");
    bar.appendChild(el("div", { class: "completeness-fill " + fc, style: `width:${pct}%` }));
    comp.appendChild(bar);
    if (gaps > 0) comp.appendChild(el("span", { class: "gap-badge" }, `${gaps} missing`));
    head.appendChild(comp);
    drawer.appendChild(head);

    // all fields
    const dl = el("dl", { class: "drawer-fields" });
    COLS.forEach(c => {
      const dt = el("dt");
      dt.appendChild(document.createTextNode(c.label));
      dt.appendChild(el("small", {}, c.th));
      dl.appendChild(dt);
      const dd = el("dd", { class: isEmpty(r[c.key]) ? "is-missing" : "" });
      if (isEmpty(r[c.key])) {
        dd.appendChild(missing());
      } else if (c.multi) {
        const vals = splitMulti(r[c.key]);
        const wrap = el("div", { class: "tag-list" });
        vals.forEach(v => wrap.appendChild(el("span", { class: "tag-soft" + (c.key === "service_type" ? " accent" : "") }, v)));
        dd.appendChild(wrap);
      } else if (c.url) {
        dd.appendChild(el("a", { href: r[c.key], target: "_blank", rel: "noopener" }, r[c.key]));
      } else {
        dd.textContent = r[c.key];
      }
      dl.appendChild(dd);
    });
    drawer.appendChild(dl);
  };

  // ===== Render ======================================================

  const renderAll = () => {
    renderTopMeta();
    renderFilters();
    renderCards();
    renderSourceToggle();
    if (state.openId) renderDrawer();
  };

  // ===== Init ========================================================

  const renderSourceToggle = () => {
    $$("#source-toggle .seg-btn").forEach(btn => {
      const v = btn.dataset.source;
      const enabled = v !== "uploaded" || !!state.uploadedRows;
      btn.classList.toggle("active", state.source === v);
      btn.disabled = !enabled && state.source !== v;
      if (v === "uploaded") {
        btn.textContent = state.uploadedLabel
          ? "Uploaded file"
          : "Uploaded — none yet";
      }
    });
    // Show bundled error inline if present and active source is bundled
    const warn = $("#bundled-warn");
    if (state.source === "bundled" && state.bundledError) {
      warn.classList.remove("hidden");
      warn.textContent = `Couldn’t load ${BUNDLED_XLSX_PATH}. Serve over http:// (e.g. python3 -m http.server) or pick a file manually.`;
    } else {
      warn.classList.add("hidden");
    }
  };

  window.addEventListener("DOMContentLoaded", async () => {
    // Start with demo data while the bundled file streams in,
    // so the page never sits blank.
    state.source = "bundled";
    state.sourceLabel = "loading…";
    state.rows = [];
    renderAll();

    await loadBundled();
    if (state.bundledRows) {
      setSource("bundled");
    } else {
      // Fall back to demo if bundled fetch failed (e.g. opened via file://)
      setSource("demo");
    }

    // Source toggle
    $$("#source-toggle .seg-btn").forEach(btn => {
      btn.addEventListener("click", () => {
        const v = btn.dataset.source;
        if (v === "uploaded" && !state.uploadedRows) {
          $("#file").click();
          return;
        }
        setSource(v);
      });
    });

    // File picker (always available)
    $("#file").addEventListener("change", (e) => {
      const f = e.target.files[0];
      if (f) loadUploadedFile(f);
      e.target.value = "";
    });
    $("#upload-btn").addEventListener("click", () => $("#file").click());

    $("#search").addEventListener("input", (e) => {
      state.search = e.target.value;
      renderTopMeta();
      renderCards();
    });

    $("#gaps-toggle").addEventListener("change", (e) => {
      state.onlyGaps = e.target.checked;
      $("#gaps-toggle-wrap").classList.toggle("on", state.onlyGaps);
      renderTopMeta();
      renderFilters();
      renderCards();
    });

    $("#sort-select").addEventListener("change", (e) => {
      state.sort = e.target.value;
      renderCards();
    });

    $("#clear-filters").addEventListener("click", () => {
      state.filters = {};
      state.search = "";
      state.onlyGaps = false;
      $("#search").value = "";
      $("#gaps-toggle").checked = false;
      $("#gaps-toggle-wrap").classList.remove("on");
      renderAll();
    });

    // close drawer
    $("#drawer-overlay").addEventListener("click", (e) => {
      if (e.target.id === "drawer-overlay") closeDrawer();
    });
    document.addEventListener("keydown", (e) => {
      if (e.key === "Escape" && state.openId) closeDrawer();
    });

    renderAll();
  });

})();
