"use strict";
const $ = (id) => document.getElementById(id);
let worker = null;                      // started by boot(), after the service worker is in control
let file = null, ready = false, busy = false, urls = [];
let CATALOG = [];                       // graph types offered by the Python engine

// ---------------------------------------------------------------- worker messages
const onWorkerMessage = (e) => {
  const m = e.data;
  if (m.type === "status") setStatus(m.text, true);
  else if (m.type === "ready") { ready = true; CATALOG = m.catalog; setStatus("Ready. Choose a Gen5 .xpt file.", false); updateGo(); }
  else if (m.type === "log") addLog(m.text);
  else if (m.type === "fatal") { showError("Could not start the Python engine:\n" + m.text + "\n\nOpen this page through a web server (see README), not by double-clicking index.html."); setStatus("Failed to start.", false); }
  else if (m.type === "failed") { busy = false; showError(m.text); setStatus("Extraction failed.", false); updateGo(); }
  else if (m.type === "custom_done") onDrawn(m);
  else if (m.type === "labels_done") onLabels(m);
  else if (m.type === "rebuilt") onRebuilt(m);
  else if (m.type === "done") { busy = false; setStatus("Done.", false); updateGo(); showResults(m); }
};
setStatus("Loading Python engine...", true);

function setStatus(text, spin) { $("status-text").textContent = text; $("status").classList.toggle("ok", !spin); }
function addLog(t) {
  const l = $("log"); l.textContent += t + "\n"; l.scrollTop = l.scrollHeight;
  $("log-toggle").hidden = false;
  if (busy) setStatus(t, true);                 // the status line shows the latest step
}
$("log-toggle").onclick = () => { const l = $("log"); l.hidden = !l.hidden; $("log-toggle").textContent = l.hidden ? "Details" : "Hide details"; };
function showError(t) { const el = $("error"); el.hidden = false; el.textContent = t; }
function updateGo() { $("go").disabled = !(ready && file && !busy); }

// ---------------------------------------------------------------- file input
function setFile(f) {
  if (!f) return;
  file = f;
  $("drop").classList.add("has");
  $("drop-title").textContent = f.name;
  $("drop-sub").textContent = (f.size / 1024).toFixed(0) + " KB";
  $("drop-action").textContent = "Change file";
  $("error").hidden = true;
  updateGo();
}
$("file").addEventListener("change", (e) => setFile(e.target.files[0]));
const drop = $("drop");
["dragenter", "dragover"].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add("over"); }));
["dragleave", "drop"].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove("over"); }));
drop.addEventListener("drop", (e) => setFile(e.dataTransfer.files[0]));
drop.addEventListener("keydown", (e) => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); $("file").click(); } });

$("go").addEventListener("click", async () => {
  if (!file) return;
  busy = true; updateGo();
  $("results").hidden = true; $("error").hidden = true; $("log").textContent = ""; $("log").hidden = true; $("log-toggle").hidden = true; $("log-toggle").textContent = "Details";
  setStatus("Working...", true);
  const buffer = await file.arrayBuffer();
  worker.postMessage({ type: "run", name: file.name, buffer, plots: $("plots").checked }, [buffer]);
});

// ---------------------------------------------------------------- results
const mime = (p) => p.endsWith(".png") ? "image/png" : p.endsWith(".xlsx") ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" : p.endsWith(".csv") ? "text/csv" : "text/plain";

let STEM = "experiment", zipUrl = null, dirty = false, pendingDownload = false;

function applyFiles(m) {
  urls.forEach(URL.revokeObjectURL); urls = [];
  const files = m.files.map((f) => {
    const url = URL.createObjectURL(new Blob([f.data], { type: mime(f.path) }));
    urls.push(url);
    return { ...f, url, size: f.data.byteLength };
  });
  zipUrl = URL.createObjectURL(new Blob([m.zip], { type: "application/zip" })); urls.push(zipUrl);
  fillTable(m.overview);
  const readme = files.find((f) => f.path === "README.txt");
  $("readme").textContent = readme ? new TextDecoder().decode(readme.data) : "";
  fillFiles(files);
  setDirty(false);
}

function setDirty(v) { dirty = v; $("stale").hidden = $("stale2").hidden = !v; }

function showResults(m) {
  STEM = m.stem;
  applyFiles(m);
  $("results-title").textContent = m.stem;
  const rows = m.overview || [];
  $("results-meta").textContent = `${new Set(rows.map((r) => r.plate)).size} plates  \u00b7  ${new Set(rows.map((r) => r.read)).size} reads  \u00b7  ${rows.reduce((s, r) => s + r.points, 0).toLocaleString()} measurements`;
  document.body.classList.add("has-results");
  initGraphs(m.state, false);
  initNames(m.state);
  $("results").hidden = false;
  selectTab("graphs");
  $("results").scrollIntoView({ behavior: "smooth" });
}

// "Download everything": if names/notes changed since the files were built, rebuild them first
$("zip").onclick = () => {
  if (!dirty) { download(zipUrl, STEM + "_extracted.zip"); return; }
  pendingDownload = true;
  $("zip").disabled = true; $("zip-label").textContent = "Preparing your download...";
  worker.postMessage({ type: "rebuild", plots: $("plots").checked });
};
function onRebuilt(m) {
  $("zip").disabled = false; $("zip-label").textContent = "Download everything";
  if (!m.ok) { pendingDownload = false; showError("Could not rebuild the files: " + m.error); return; }
  applyFiles(m);
  if (pendingDownload) { pendingDownload = false; download(zipUrl, STEM + "_extracted.zip"); }
}

function download(url, name) { const a = document.createElement("a"); a.href = url; a.download = name; document.body.appendChild(a); a.click(); a.remove(); }

function fillTable(rows, t = $("tbl"), only = null) {
  t.innerHTML = "";
  if (!rows || !rows.length) return;
  const cols = only ? only.filter((c) => c in rows[0]) : Object.keys(rows[0]);
  const head = t.createTHead().insertRow();
  cols.forEach((c) => { const th = document.createElement("th"); th.textContent = c; head.append(th); });
  const body = t.createTBody();
  for (const r of rows) {
    const tr = body.insertRow();
    cols.forEach((c) => {
      const td = tr.insertCell(); const v = r[c];
      if (typeof v === "number") { td.className = "num"; td.textContent = Number.isInteger(v) ? v : Number(v.toPrecision(4)); }
      else td.textContent = v ?? "";
    });
  }
}

function fillFiles(files) {
  const groups = [["graphs/by_plate_read/", "Graphs: plate x read"], ["graphs/by_read/", "Graphs: read across plates"], ["graphs/by_plate/", "Graphs: plate overview"],
    ["by_read/", "Excel: one workbook per read"], ["by_plate/", "Excel: one workbook per plate"], ["csv/", "CSV: one file per plate + read"], ["", "Summary, analysis and data"]];
  const used = new Set(), root = $("filelist"); root.innerHTML = "";
  for (const [prefix, title] of groups) {
    const items = files.filter((f) => !used.has(f.path) && (prefix ? f.path.startsWith(prefix) : !f.path.includes("/"))).sort((a, b) => a.path.localeCompare(b.path));
    if (!items.length) continue;
    items.forEach((f) => used.add(f.path));
    const d = document.createElement("div"); d.className = "filegroup";
    d.innerHTML = `<h4>${title} (${items.length})</h4>`;
    for (const f of items) {
      const a = document.createElement("a"); a.className = "filelink"; a.href = f.url; a.download = f.path.split("/").pop();
      a.innerHTML = `<span></span><small>${(f.size / 1024).toFixed(f.size > 102400 ? 0 : 1)} KB</small>`;
      a.firstChild.textContent = f.path.slice(prefix.length);
      d.append(a);
    }
    root.append(d);
  }
}

// ---------------------------------------------------------------- tabs + lightbox
function selectTab(name) {
  document.querySelectorAll("#tabs button").forEach((b) => b.classList.toggle("active", b.dataset.tab === name));
  document.querySelectorAll(".pane").forEach((p) => (p.hidden = p.id !== "pane-" + (name === "graphs" ? "wells" : name)));
}
$("tabs").addEventListener("click", (e) => { if (e.target.dataset.tab) selectTab(e.target.dataset.tab); });
document.addEventListener("click", (e) => {
  if (e.target.matches(".fig img")) { $("lightbox-img").src = e.target.src; $("lightbox").hidden = false; }
  else if (e.target.closest("#lightbox")) $("lightbox").hidden = true;
});
document.addEventListener("keydown", (e) => { if (e.key === "Escape") $("lightbox").hidden = true; });

// ---------------------------------------------------------------- graphs: everything redraws as soon as you change anything
const ROWS = "ABCDEFGH";
// mode: "all" = every well with data | "single" = click one well | "multi" = click/drag/type several wells
const W = { avail: {}, reads: [], spec: {}, plates: [], rawOf: {}, plate: new Set(), read: null, kind: "dashboard", mode: "all", sel: new Set(), id: 0, timer: null, png: null, csv: null, xlsx: null, urls: [] };
const wellKey = (w) => [ROWS.indexOf(w[0]), parseInt(w.slice(1), 10)];
const sortWells = (a) => [...a].sort((x, y) => { const [r1, c1] = wellKey(x), [r2, c2] = wellKey(y); return r1 - r2 || c1 - c2; });

function initGraphs(st, keep) {
  const prev = keep ? { read: W.read, kind: W.kind, mode: W.mode, origs: [...W.plate].map((l) => W.rawOf[l]).filter(Boolean), sel: [...W.sel] } : null;
  const labelOf = (o) => st.plate_names[o] || o;
  W.avail = st.avail; W.reads = st.reads.map((r) => r.label); W.spec = Object.fromEntries(st.reads.map((r) => [r.label, r.spectrum]));
  W.plates = st.raw_plates.map(labelOf); W.rawOf = Object.fromEntries(st.raw_plates.map((o) => [labelOf(o), o]));
  const origs = prev && prev.origs.length ? prev.origs : st.raw_plates.slice(0, 1);
  W.plate = new Set(origs.map(labelOf));
  W.read = prev && W.reads.includes(prev.read) ? prev.read : W.reads[0];
  W.kind = prev ? prev.kind : "dashboard"; W.mode = prev ? prev.mode : "all"; W.sel = new Set(prev ? prev.sel : []);
  const rs = $("w-read"); rs.innerHTML = ""; W.reads.forEach((r) => rs.add(new Option(r.replace(/:/, ": "), r))); rs.value = W.read;
  fillKinds(); buildChips(); buildMap(); syncMode(); schedule(0);
}

function buildChips() {
  const box = $("w-plates"); box.innerHTML = "";
  const mk = (label, on, fn) => { const b = document.createElement("button"); b.type = "button"; b.className = "chip" + (on ? " on" : ""); b.textContent = label; b.onclick = fn; box.append(b); return b; };
  for (const p of W.plates) {
    mk(p, W.plate.has(p), () => {
      if (W.plate.has(p)) { if (W.plate.size > 1) W.plate.delete(p); } else W.plate.add(p);
      buildChips(); refreshMap(); schedule(0);
    });
  }
  if (W.plates.length > 1) mk("All plates", W.plate.size === W.plates.length, () => { W.plate = new Set(W.plates); buildChips(); refreshMap(); schedule(0); });
}

const enabledWells = () => {
  const s = new Set();
  for (const p of W.plate) for (const w of (W.avail[p] && W.avail[p][W.read]) || []) s.add(w);
  return s;
};

function fillKinds() {                  // only graph types that make sense for this read (spectrum vs endpoint)
  const want = W.spec[W.read] ? "spectrum" : "endpoint";
  const ok = CATALOG.filter((c) => c.applies === "both" || c.applies === want);
  const sel = $("w-kind"); sel.innerHTML = "";
  ok.forEach((c) => sel.add(new Option(c.label, c.id)));
  if (!ok.some((c) => c.id === W.kind)) W.kind = "dashboard";
  sel.value = W.kind;
}

function buildMap() {
  const m = $("platemap"); m.innerHTML = "";
  m.append(document.createElement("span"));
  for (let c = 1; c <= 12; c++) { const h = document.createElement("button"); h.className = "hd"; h.textContent = c; h.dataset.col = c; h.type = "button"; m.append(h); }
  for (const r of ROWS) {
    const h = document.createElement("button"); h.className = "hd"; h.textContent = r; h.dataset.row = r; h.type = "button"; m.append(h);
    for (let c = 1; c <= 12; c++) { const b = document.createElement("button"); b.className = "well"; b.dataset.well = r + c; b.textContent = r + c; b.type = "button"; m.append(b); }
  }
}

function refreshMap() {
  const en = enabledWells();
  if (W.mode === "all") W.sel = new Set(en);
  else for (const w of [...W.sel]) if (!en.has(w)) W.sel.delete(w);
  document.querySelectorAll("#platemap .well").forEach((b) => {
    const w = b.dataset.well, has = en.has(w);
    b.disabled = !has; b.classList.toggle("has", has); b.classList.toggle("sel", W.sel.has(w));
  });
  document.querySelectorAll("#platemap .hd").forEach((h) => h.classList.toggle("act", W.mode === "multi"));
  $("w-hint").textContent = W.mode === "multi"
    ? "Click wells to add or remove them, drag across wells to paint, click a row letter or column number to toggle a whole row/column. Greyed wells have no data for the chosen plates and read."
    : "Click one well. Switch to Multiple wells to pick several. Greyed wells have no data for the chosen plates and read.";
}

function syncMode() {
  for (const m of ["all", "single", "multi"]) $("mode-" + m).classList.toggle("active", W.mode === m);
  $("mapcol").hidden = W.mode === "all";
  $("wellbody").classList.toggle("all", W.mode === "all");
  $("w-all").hidden = W.mode !== "multi";
  refreshMap();
}

function setMode(mode) {
  const prev = new Set(W.sel);
  W.mode = mode;
  if (mode === "single") W.sel = new Set(sortWells(prev).slice(0, 1));
  else if (mode === "multi") W.sel = prev;
  syncMode(); schedule(0);
}

function setSel(arr) { W.sel = new Set(arr); refreshMap(); schedule(0); }

function toggleGroup(wells) {
  const en = enabledWells(), group = wells.filter((w) => en.has(w));
  if (!group.length) return;
  const allOn = group.every((w) => W.sel.has(w)), next = new Set(W.sel);
  group.forEach((w) => (allOn ? next.delete(w) : next.add(w)));
  setSel(next);
}

(function wire() {
  const map = $("platemap"); let painting = null;       // true = adding, false = removing
  const wellOf = (el) => el && el.closest && el.closest(".well");
  map.addEventListener("pointerdown", (e) => {
    const b = wellOf(e.target);
    if (b && !b.disabled) {
      e.preventDefault(); const w = b.dataset.well;
      if (W.mode === "single") { setSel([w]); return; }
      painting = !W.sel.has(w); paint(w);
    }
    const h = e.target.closest(".hd");
    if (h && W.mode === "multi") {
      if (h.dataset.col) toggleGroup(ROWS.split("").map((r) => r + h.dataset.col));
      else if (h.dataset.row) toggleGroup(Array.from({ length: 12 }, (_, i) => h.dataset.row + (i + 1)));
    }
  });
  function paint(w) { const n = new Set(W.sel); painting ? n.add(w) : n.delete(w); W.sel = n; refreshMap(); schedule(150); }
  map.addEventListener("pointerover", (e) => {
    if (painting === null || W.mode !== "multi" || !(e.buttons & 1)) return;
    const b = wellOf(e.target); if (b && !b.disabled && W.sel.has(b.dataset.well) !== painting) paint(b.dataset.well);
  });
  document.addEventListener("pointerup", () => { painting = null; });
  map.addEventListener("dragstart", (e) => e.preventDefault());

  $("mode-all").onclick = () => setMode("all");
  $("mode-single").onclick = () => setMode("single");
  $("mode-multi").onclick = () => setMode("multi");
  $("w-clear").onclick = () => { setSel([]); };
  $("w-all").onclick = () => setSel([...enabledWells()]);
  $("w-read").onchange = (e) => { W.read = e.target.value; fillKinds(); refreshMap(); schedule(0); };
  $("w-kind").onchange = (e) => { W.kind = e.target.value; schedule(0); };
  const apply = () => {
    const parsed = parseWells($("w-text").value);
    if (parsed.error) { setWStatus(parsed.error, false); return; }
    const en = enabledWells(), keep = parsed.wells.filter((w) => en.has(w));
    if (!keep.length) { setWStatus("None of those wells have data for the chosen plates and read.", false); return; }
    if (W.mode === "all" || (keep.length > 1 && W.mode === "single")) { W.mode = "multi"; syncMode(); }
    setSel(keep);
  };
  $("w-apply").onclick = apply;
  $("w-text").addEventListener("keydown", (e) => { if (e.key === "Enter") apply(); });
  $("w-dl-png").onclick = () => W.png && download(W.png, "graph.png");
  $("w-dl-csv").onclick = () => W.csv && download(W.csv, "selected_data.csv");
  $("w-dl-xlsx").onclick = () => W.xlsx && download(W.xlsx, "selected_data_with_notes.xlsx");
})();

// same syntax as the Python tool:  A1  |  A1:C3  |  A (row)  |  2 (column)
function parseWells(text) {
  const out = new Set();
  for (const tok of text.toUpperCase().split(/[,\s;]+/).filter(Boolean)) {
    let m;
    if ((m = tok.match(/^([A-H])(\d{1,2})[:\-]([A-H])(\d{1,2})$/))) {
      const [r1, r2] = [ROWS.indexOf(m[1]), ROWS.indexOf(m[3])].sort((a, b) => a - b), [c1, c2] = [+m[2], +m[4]].sort((a, b) => a - b);
      for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) out.add(ROWS[r] + c);
    } else if (/^[A-H]\d{1,2}$/.test(tok)) out.add(tok);
    else if (/^[A-H]$/.test(tok)) for (let c = 1; c <= 12; c++) out.add(tok + c);
    else if (/^\d{1,2}$/.test(tok)) for (const r of ROWS) out.add(r + +tok);
    else return { error: `Cannot understand "${tok}". Use A1, A1:C3, a row letter (A) or a column number (2).` };
  }
  return { wells: sortWells(out).filter((w) => { const c = wellKey(w)[1]; return c >= 1 && c <= 12; }) };
}

function setWStatus(t, busy) { const el = $("w-status"); el.textContent = t; el.classList.toggle("busy", !!busy); }
function showEmpty(msg) { $("w-fig").hidden = $("w-dl").hidden = $("w-stats-wrap").hidden = true; setWStatus(msg, false); }

function schedule(delay) {
  clearTimeout(W.timer);
  if (!W.sel.size) { showEmpty(W.mode === "all" ? "No data for this read in the chosen plates." : "Click a well on the plate map."); return; }
  $("w-fig").classList.add("stale");                       // dim the old graph so it is obvious it is being replaced
  setWStatus("Drawing...", true);
  W.timer = setTimeout(() => {
    const wells = sortWells(W.sel), id = ++W.id;
    worker.postMessage({ type: "custom", id, plates: [...W.plate], read: W.read, wells, kind: W.kind });
  }, delay);
}

function onDrawn(m) {
  if (m.id !== W.id) return;                               // a newer request is already on its way
  $("w-fig").classList.remove("stale");
  if (!m.ok) { showEmpty(m.error ? "Could not draw this: " + m.error : "No data for this selection."); return; }
  W.urls.forEach(URL.revokeObjectURL);
  W.png = URL.createObjectURL(new Blob([m.png], { type: "image/png" }));
  W.csv = URL.createObjectURL(new Blob([m.csv], { type: "text/csv" }));
  W.xlsx = URL.createObjectURL(new Blob([m.xlsx], { type: mime("x.xlsx") }));
  W.urls = [W.png, W.csv, W.xlsx];
  $("w-img").src = W.png;
  $("w-fig").hidden = $("w-dl").hidden = false;
  const kindName = (CATALOG.find((c) => c.id === W.kind) || {}).label || W.kind;
  const n = [...W.plate].reduce((t, p) => t + ((W.avail[p] && W.avail[p][W.read]) || []).filter((w) => W.sel.has(w)).length, 0);
  const where = W.plate.size > 1 ? ` on ${W.plate.size} plates` : "";
  setWStatus(`${kindName}  -  ${n} well${n > 1 ? "s" : ""}${where}${n <= 14 && W.plate.size === 1 ? ": " + sortWells(W.sel).join(", ") : ""}`, false);
  $("w-stats-wrap").hidden = !m.stats.length;
  fillTable(m.stats, $("w-stats"), ["plate", "well", "sample", "peak_wavelength_nm", "peak_value", "mean_value", "AUC", "FWHM_nm"]);
}

// ---------------------------------------------------------------- names & notes
const NM = { raw: [], plateNames: {}, wellNames: {}, wells: {}, cur: null };

function initNames(st) {
  NM.raw = st.raw_plates; NM.plateNames = { ...st.plate_names }; NM.wells = st.wells; NM.cur = NM.raw[0];
  NM.wellNames = {}; st.well_names.forEach(([p, w, n]) => (NM.wellNames[p + "|" + w] = n));
  $("nm-notes").value = st.notes || "";
  const box = $("nm-plates"); box.innerHTML = "";
  for (const p of NM.raw) {
    const sp = document.createElement("span"); sp.textContent = p;
    const inp = document.createElement("input"); inp.type = "text"; inp.value = NM.plateNames[p] || ""; inp.placeholder = p; inp.maxLength = 40;
    inp.oninput = () => { NM.plateNames[p] = inp.value.trim(); fillNmPlates(); };
    box.append(sp, inp);
  }
  fillNmPlates(); buildWellTable();
  $("nm-status").textContent = "";
}

function fillNmPlates() {
  const sel = $("nm-plate"); sel.innerHTML = "";
  NM.raw.forEach((p) => sel.add(new Option(NM.plateNames[p] ? `${p} - ${NM.plateNames[p]}` : p, p)));
  sel.value = NM.cur;
}

function buildWellTable() {
  const t = $("nm-wells"); t.innerHTML = "";
  const wells = NM.wells[NM.cur] || [];
  if (!wells.length) { t.innerHTML = "<tr><td>No wells with data on this plate.</td></tr>"; return; }
  for (const w of wells) {
    const tr = t.insertRow(), key = NM.cur + "|" + w;
    tr.insertCell().textContent = w;
    const inp = document.createElement("input"); inp.type = "text"; inp.placeholder = "e.g. Blank, Sample 3"; inp.maxLength = 60; inp.value = NM.wellNames[key] || "";
    inp.oninput = () => { const v = inp.value.trim(); if (v) NM.wellNames[key] = v; else delete NM.wellNames[key]; };
    tr.insertCell().append(inp);
  }
}

$("nm-plate").onchange = (e) => { NM.cur = e.target.value; buildWellTable(); };
$("nm-load").onclick = () => $("nm-file").click();
$("nm-file").addEventListener("change", async (e) => {
  const f = e.target.files[0]; if (!f) return;
  const text = await f.text(), box = $("nm-notes");
  box.value = box.value.trim() ? box.value.trim() + "\n\n" + text : text;
  $("nm-file").value = "";
});
$("nm-apply").onclick = () => {
  const wells = Object.entries(NM.wellNames).filter(([, n]) => n).map(([k, n]) => { const [p, w] = k.split("|"); return [p, w, n]; });
  const plateMap = Object.fromEntries(Object.entries(NM.plateNames).filter(([, n]) => n));
  $("nm-status").textContent = "Applying..."; $("nm-apply").disabled = true;
  worker.postMessage({ type: "labels", plateMap, wells, notes: $("nm-notes").value });
};
function onLabels(m) {
  $("nm-apply").disabled = false;
  if (!m.ok) { $("nm-status").textContent = m.error; return; }
  setDirty(true);
  initGraphs(m.state, true);
  $("nm-status").textContent = "Applied. The graphs are updated, and your downloads will include these names and notes.";
}

// ---------------------------------------------------------------- start-up, offline support, install button
function startWorker() {
  worker = new Worker("worker.js?v=10");
  worker.onmessage = onWorkerMessage;
  worker.postMessage({ type: "init" });
}

async function boot() {
  const canSW = "serviceWorker" in navigator && (location.protocol === "https:" || ["localhost", "127.0.0.1"].includes(location.hostname));
  if (canSW) {
    try {
      const wasControlled = !!navigator.serviceWorker.controller;
      await navigator.serviceWorker.register("sw.js");
      await navigator.serviceWorker.ready;
      // First visit: reload once so the page and the Python engine are loaded through the service worker, which saves them for offline use.
      if (!wasControlled && !sessionStorage.getItem("sw-reloaded")) {
        sessionStorage.setItem("sw-reloaded", "1");
        location.reload();
        return;
      }
    } catch (err) { /* private window or blocked storage: run without offline support */ }
  }
  startWorker();
}

let installEvent = null;
window.addEventListener("beforeinstallprompt", (e) => { e.preventDefault(); installEvent = e; $("install").hidden = false; });
window.addEventListener("appinstalled", () => { installEvent = null; $("install").hidden = true; });
$("install").onclick = async () => {
  if (!installEvent) return;
  installEvent.prompt();
  try { await installEvent.userChoice; } catch (err) { /* ignored */ }
  installEvent = null; $("install").hidden = true;
};

boot();
