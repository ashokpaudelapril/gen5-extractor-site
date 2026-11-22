// Runs Python (Pyodide) in a background thread so the page stays responsive.
// Everything happens inside the visitor's browser: the .xpt file is never uploaded anywhere.
const PYODIDE = "https://cdn.jsdelivr.net/pyodide/v0.28.3/full/";
importScripts(PYODIDE + "pyodide.js");

const BASE = self.location.href.replace(/[^/]*$/, "");
const WHEELS = [
  "wheels/et_xmlfile-2.0.0-py3-none-any.whl",
  "wheels/olefile-0.47-py2.py3-none-any.whl",
  "wheels/openpyxl-3.1.5-py2.py3-none-any.whl",
];
const post = (type, o = {}, transfer = []) => postMessage({ type, ...o }, transfer);
self.logmsg = (m) => post("log", { text: String(m) });

let py = null;
let ready = null;
let stem = "experiment";

async function init() {
  post("status", { text: "Loading Python (first visit downloads ~40 MB, then it is cached)..." });
  py = await loadPyodide({ indexURL: PYODIDE });
  post("status", { text: "Loading numpy, pandas, matplotlib..." });
  await py.loadPackage(["numpy", "pandas", "matplotlib", "micropip"]);
  post("status", { text: "Loading Excel / Gen5 readers..." });
  const urls = JSON.stringify(WHEELS.map((w) => BASE + w));
  await py.runPythonAsync(`
import micropip
await micropip.install(${urls}, deps=False)
`);
  const src = await (await fetch(BASE + "gen5_extract.py?v=10", { cache: "no-cache" })).text();
  py.FS.mkdirTree("/work/in");
  py.FS.writeFile("/work/gen5_extract.py", src);
  await py.runPythonAsync(`
import sys, os, io, json, shutil, zipfile
sys.path.insert(0, "/work")
import gen5_extract as g
from js import logmsg

def make_zip(stem):
    out = f"/work/{stem}_extracted"
    with zipfile.ZipFile(f"/work/{stem}_extracted.zip", "w", zipfile.ZIP_DEFLATED) as z:
        for root, _, files in os.walk(out):
            for f in files:
                full = os.path.join(root, f)
                z.write(full, os.path.join(f"{stem}_extracted", os.path.relpath(full, out)))

def overview(ov):
    cols = ["plate", "read", "unit", "wells", "points", "invalid", "min", "max", "mean"]
    return json.loads(ov[cols].to_json(orient="records"))

def state():
    df, plates, info = g.labeled()
    L = g.LAST
    avail = {}
    for (p, r), gg in df.groupby(["plate", "read"], sort=False):
        avail.setdefault(p, {})[r] = sorted(g.usable(gg).well.unique(), key=g.well_key)
    pm = L["plate_names"]
    wells = {p: sorted({w for ws in avail.get(pm.get(p, p), {}).values() for w in ws}, key=g.well_key) for p in L["raw_plates"]}
    return json.dumps({
        "avail": avail,
        "reads": [{"label": r["label"], "spectrum": r["start"] is not None} for r in L["reads"]],
        "raw_plates": L["raw_plates"], "plate_names": pm,
        "well_names": [[p, w, n] for (p, w), n in L["well_names"].items()],
        "wells": wells, "notes": L["notes"]})

def run_job(path, stem, plots):
    out = f"/work/{stem}_extracted"
    shutil.rmtree(out, ignore_errors=True)
    try:
        _, ov = g.process(path, out, plots, log=logmsg)
    except SystemExit as e:
        return json.dumps({"ok": False, "error": str(e.code)})
    except Exception as e:
        return json.dumps({"ok": False, "error": f"{type(e).__name__}: {e}"})
    make_zip(stem)
    return json.dumps({"ok": True, "out": out, "overview": overview(ov)})

def apply_labels(plate_map, triples, notes):
    try:
        g.set_labels_struct(json.loads(plate_map), [tuple(t) for t in json.loads(triples)], notes)
    except SystemExit as e:
        return json.dumps({"ok": False, "error": str(e.code)})
    return json.dumps({"ok": True})

def rebuild(stem, plots):
    out = f"/work/{stem}_extracted"
    shutil.rmtree(out, ignore_errors=True)
    try:
        ov = g.regenerate(out, plots, logmsg)
    except SystemExit as e:
        return json.dumps({"ok": False, "error": str(e.code)})
    except Exception as e:
        return json.dumps({"ok": False, "error": f"{type(e).__name__}: {e}"})
    make_zip(stem)
    return json.dumps({"ok": True, "out": out, "overview": overview(ov)})

def custom(plates, read, wells, kind):
    df, _, _ = g.labeled()
    res = g.custom_view(df, g.LAST["reads"], plates, read, wells, "/work/custom", kind)
    if res is None:
        return json.dumps({"ok": False})
    return json.dumps({"ok": True, "stats": json.loads(res["stats"].to_json(orient="records"))})
`);
  const cat = await py.runPythonAsync("import json; json.dumps(g.graph_catalog())");
  post("ready", { catalog: JSON.parse(cat) });
}

function walk(dir, base, acc) {
  for (const name of py.FS.readdir(dir)) {
    if (name === "." || name === "..") continue;
    const full = dir + "/" + name;
    if (py.FS.isDir(py.FS.stat(full).mode)) walk(full, base, acc);
    else acc.push({ path: full.slice(base.length + 1), data: py.FS.readFile(full) });
  }
  return acc;
}

function collect(res) {
  const files = walk(res.out, res.out, []);
  const zip = py.FS.readFile(`/work/${stem}_extracted.zip`);
  return { files, zip, transfer: files.map((f) => f.data.buffer).concat([zip.buffer]) };
}

self.onmessage = async (e) => {
  const m = e.data;
  try {
    if (m.type === "init") {
      ready = init().catch((err) => post("fatal", { text: String(err) }));
      return;
    }
    await ready;
    if (m.type === "run") {
      stem = m.name.replace(/\.[^.]*$/, "").replace(/[^A-Za-z0-9._-]+/g, "_") || "experiment";
      const path = `/work/in/${stem}.xpt`;
      py.FS.writeFile(path, new Uint8Array(m.buffer));
      const res = JSON.parse(await py.runPythonAsync(`run_job(${JSON.stringify(path)}, ${JSON.stringify(stem)}, ${m.plots ? "True" : "False"})`));
      if (!res.ok) return post("failed", { text: res.error });
      const c = collect(res);
      const st = JSON.parse(await py.runPythonAsync("state()"));
      post("done", { stem, overview: res.overview, state: st, files: c.files, zip: c.zip }, c.transfer);
    } else if (m.type === "labels") {
      py.globals.set("lab_plates", JSON.stringify(m.plateMap));      // plain values, never pasted into code
      py.globals.set("lab_wells", JSON.stringify(m.wells));
      py.globals.set("lab_notes", m.notes || "");
      const res = JSON.parse(await py.runPythonAsync("apply_labels(lab_plates, lab_wells, lab_notes)"));
      if (!res.ok) return post("labels_done", { ok: false, error: res.error });
      post("labels_done", { ok: true, state: JSON.parse(await py.runPythonAsync("state()")) });
    } else if (m.type === "rebuild") {
      const res = JSON.parse(await py.runPythonAsync(`rebuild(${JSON.stringify(stem)}, ${m.plots ? "True" : "False"})`));
      if (!res.ok) return post("rebuilt", { ok: false, error: res.error });
      const c = collect(res);
      post("rebuilt", { ok: true, overview: res.overview, files: c.files, zip: c.zip }, c.transfer);
    } else if (m.type === "custom") {
      const res = JSON.parse(await py.runPythonAsync(`custom(${JSON.stringify(m.plates)}, ${JSON.stringify(m.read)}, ${JSON.stringify(m.wells)}, ${JSON.stringify(m.kind || "dashboard")})`));
      if (!res.ok) return post("custom_done", { id: m.id, ok: false });
      const png = py.FS.readFile("/work/custom/selection.png"), csv = py.FS.readFile("/work/custom/selection.csv"), xlsx = py.FS.readFile("/work/custom/selection.xlsx");
      post("custom_done", { id: m.id, ok: true, png, csv, xlsx, stats: res.stats }, [png.buffer, csv.buffer, xlsx.buffer]);
    }
  } catch (err) {
    if (m.type === "custom") post("custom_done", { id: m.id, ok: false, error: String(err) });
    else if (m.type === "labels") post("labels_done", { ok: false, error: String(err) });
    else if (m.type === "rebuild") post("rebuilt", { ok: false, error: String(err) });
    else post("failed", { text: String(err) });
  }
};
