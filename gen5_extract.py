#!/usr/bin/env python3
"""
gen5_extract.py - extract data from BioTek Gen5 experiment files (.xpt)

Usage:
    python gen5_extract.py Experiment1.xpt            # output goes to Experiment1_extracted/
    python gen5_extract.py Experiment1.xpt -o MyOut   # choose the output folder
    python gen5_extract.py Experiment1.xpt --no-plots # extraction only, skip analysis/graphs
    python gen5_extract.py                            # asks for the file name

Output layout (one folder per input file):
    README.txt                  what is in the folder + experiment info
    Analysis.xlsx               per-well peaks/AUC/FWHM, per plate+read mean/SD/CV
    graphs/by_plate_read/       one detailed figure per plate x read (Plate 1 + Read 1, ...)
    graphs/by_read/ by_plate/   comparisons across plates / overview per plate
    Summary.xlsx                Overview, Info, Peaks (max per well), Endpoint reads
    by_plate/Plate1.xlsx ...    one workbook per plate, one sheet per read
    by_read/Read1_xxx.xlsx ...  one workbook per read, one sheet per plate
    csv/Plate1_Read1_xxx.csv    one CSV per plate+read (same tables as the sheets)
    all_data_long.csv           everything in one tidy table (plate, read, well, wavelength, value)

Spectrum reads  -> table with wavelength down the rows, one column per well (+ line chart)
Endpoint reads  -> 8x12 plate-layout grid
Only wells that were actually measured are included (unread wells are dropped).

Requires: olefile, numpy, pandas, openpyxl, matplotlib   (matplotlib only for graphs; --no-plots skips them)
Window mode also needs PySide6.  Run with no arguments (or --gui) to open the window.
"""
import argparse
import os
import re
import struct
import sys
import textwrap
import zlib

import numpy as np
import olefile
import pandas as pd
from openpyxl.chart import LineChart, Reference

try:                                    # graphs are optional: extraction still works without matplotlib
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
except ImportError:
    plt = None

ROWS = "ABCDEFGH"
WELLS = [f"{r}{c}" for r in ROWS for c in range(1, 13)]   # row-major, 96 wells
STATUS_NOT_READ, STATUS_OK, STATUS_INVALID = 1, 0, 100     # status byte in each record
RECORD = 24                                                # bytes per data record
DPI = 150                                                  # graph resolution


# --------------------------------------------------------------------------
# Parsing
# --------------------------------------------------------------------------
def safe(name):
    return re.sub(r"[^A-Za-z0-9._-]+", "_", name).strip("_")


def well_key(w):
    return (w[0], int(w[1:]))


def parse_wells(spec):
    """'A1, B2, A1:C3, A, 2'  ->  sorted list of wells (None if spec is empty).
    A1 = one well | A1:C3 = rectangle | A = whole row | 2 = whole column"""
    if spec is None or not str(spec).strip():
        return None
    out = set()
    for tok in re.split(r"[,\s;]+", str(spec).strip().upper()):
        if not tok:
            continue
        m = re.fullmatch(r"([A-H])(\d{1,2})[:\-]([A-H])(\d{1,2})", tok)
        if m:
            r1, c1, r2, c2 = m.group(1), int(m.group(2)), m.group(3), int(m.group(4))
            for r in ROWS[min(ROWS.index(r1), ROWS.index(r2)): max(ROWS.index(r1), ROWS.index(r2)) + 1]:
                for c in range(min(c1, c2), max(c1, c2) + 1):
                    out.add(f"{r}{c}")
        elif re.fullmatch(r"[A-H]\d{1,2}", tok):
            out.add(tok)
        elif re.fullmatch(r"[A-H]", tok):
            out.update(f"{tok}{c}" for c in range(1, 13))
        elif re.fullmatch(r"\d{1,2}", tok):
            out.update(f"{r}{int(tok)}" for r in ROWS)
        else:
            sys.exit(f"ERROR: cannot understand well '{tok}'. Use A1, A1:C3, a row letter (A) or a column number (2).")
    bad = [w for w in out if w not in WELLS]
    if bad:
        sys.exit(f"ERROR: not a well on a 96-well plate: {', '.join(sorted(bad))}")
    return sorted(out, key=well_key)


def parse_plate_names(spec):
    """'Control, Treated 1h, Treated 24h' (commas) or one name per line -> list in plate order.
    A blank entry keeps that plate's default name."""
    if spec is None:
        return None
    if isinstance(spec, (list, tuple)):
        items = [str(x).strip() for x in spec]
    else:
        text = str(spec)
        items = [x.strip() for x in (text.splitlines() if "\n" in text else text.split(","))]
    while items and not items[-1]:
        items.pop()
    return items or None


def wlab(well, sample=""):
    """display label of a well: 'A2' or 'A2 (Blank)'"""
    return f"{well} ({sample})" if isinstance(sample, str) and sample else well


def label_map(g):
    """well -> display label, using the 'sample' column when there is one"""
    if "sample" not in g.columns:
        return {w: w for w in g.well.unique()}
    first = g.drop_duplicates("well").set_index("well")["sample"]
    return {w: wlab(w, sm) for w, sm in first.items()}


def parse_well_names(spec, raw_plates, plate_labels):
    """'Plate 1 | A1 = Blank; Plate 1 | A2:A4 = Sample; Control | B1 = Std'  ->  [(raw_plate, well, name), ...]
    Names belong to ONE plate (the same well holds different samples on different plates), so every entry
    names its plate. With a single plate in the file the 'plate |' part may be left out."""
    if spec is None or not str(spec).strip():
        return []
    by_name = {p.lower(): p for p in raw_plates}
    by_name.update({l.lower(): p for p, l in plate_labels.items() if l})
    out = []
    for entry in re.split(r"[;\n]+", str(spec)):
        entry = entry.strip()
        if not entry:
            continue
        if "=" not in entry:
            sys.exit(f"ERROR: well name '{entry}' needs an '=', e.g.  A1 = Blank   or   A2:A4 = Sample   or   Control | B1 = Std")
        left, name = entry.split("=", 1)
        targets = list(raw_plates)
        if "|" not in left and len(raw_plates) > 1:
            sys.exit(f"ERROR: well name '{entry}' must say which plate it is for, because each plate can hold different samples.\n"
                     f"       Example:  {raw_plates[0]} | A1 = Blank   (plates: {', '.join(raw_plates)}, or the names you gave them)")
        if "|" in left:
            pp, left = left.split("|", 1)
            if pp.strip().lower() not in by_name:
                sys.exit(f"ERROR: unknown plate '{pp.strip()}' in well names. Plates: {', '.join(raw_plates)}")
            targets = [by_name[pp.strip().lower()]]
        for w in parse_wells(left) or []:
            for p in targets:
                out.append((p, w, name.strip()))
    return out


def read_streams(path):
    if not olefile.isOleFile(path):
        sys.exit(f"ERROR: '{path}' is not a Gen5 (OLE2) file.")
    return olefile.OleFileIO(path)


def parse_plate_stream(raw):
    """Decompress one SUBSETS/n/DATA stream -> (values, status) arrays in stored order."""
    k = raw.find(b"\x78\x9c")
    if k < 0:
        raise ValueError("no compressed block found")
    u = zlib.decompress(raw[k:])
    # every record contains the 8-byte timestamp tail 'd3 01 00'; records repeat every 24 bytes
    anchors = [m.start() - 3 for m in re.finditer(rb"\xd3\x01\x00", u)]
    s = set(anchors)
    pos = [p for p in anchors if p >= 17 and (p + RECORD in s or p - RECORD in s)]
    vals = np.array([struct.unpack_from("<d", u, p - 17)[0] for p in pos])
    status = np.array([u[p - 9] for p in pos])
    return u, vals, status


def read_definitions(u, contents):
    """Read names (in order) from the dataset header of a plate stream."""
    names = []
    head = u
    for m in re.finditer(rb"([\x05-\x40])(Read (\d+):)", head):
        L = m.group(1)[0]
        text = head[m.start() + 1: m.start() + 1 + L].decode("latin1")
        if len(text) == L and re.fullmatch(r"Read \d+:[ -~]+", text) and text not in names:
            names.append(text)
    names.sort(key=lambda t: int(re.match(r"Read (\d+):", t).group(1)))
    reads = []
    for text in names:
        idx = int(re.match(r"Read (\d+):", text).group(1))
        enc = text.encode("latin1")
        n, start, step, unit = 1, None, None, None
        # spectrum parameters: <name> 01 00 ff*8 01 00 00 00 <u32 count> <f64 start> <f64 step>
        sp = re.search(re.escape(enc) + rb"\x01\x00\xff{8}\x01\x00\x00\x00(.{4})(.{8})(.{8})", contents, re.S)
        if sp:
            cnt = struct.unpack("<I", sp.group(1))[0]
            st, sw = struct.unpack("<dd", sp.group(2) + sp.group(3))
            if 1 < cnt < 5000 and 100 < st < 3000 and 0 < sw < 100:
                n, start, step = cnt, st, sw
        # unit string (length-prefixed, e.g. 'OD', 'RFU') shortly after the name
        um = re.search(re.escape(enc) + rb".{0,260}?[\x02-\x04](OD|RFU|RLU|LUM|OD600)\b", contents, re.S)
        unit = um.group(1).decode() if um else None
        reads.append(dict(index=idx, label=text, n=n, start=start, step=step, unit=unit))
    return reads


def experiment_info(ole, contents, plate_streams):
    info = {}
    txt = [s.decode("latin1") for s in re.findall(rb"[\x20-\x7e]{4,}", contents)]
    for s in txt:
        if "Take" in s or "Default" in s:
            info["Protocol / template string"] = s
            break
    for s in plate_streams[:1]:
        found = [x.decode("latin1") for x in re.findall(rb"[\x20-\x7e]{4,}", s[-4000:])]
        for x in found:
            if x.startswith("Synergy") or x.startswith("Epoch") or x.startswith("Cytation"):
                info["Instrument"] = x
        nums = [x for x in found if re.fullmatch(r"\d{5,8}", x)]
        if nums:
            info["Instrument serial / ID"] = ", ".join(nums)
        dates = [x for x in found if re.search(r"\d+\.\d+\.\d+", x)]
        if dates:
            info["Version / date strings"] = ", ".join(dates)
    return info


def extract(path):
    ole = read_streams(path)
    contents = ole.openstream("Contents").read()
    plate_ids = sorted({p[1] for p in ole.listdir() if len(p) == 3 and p[0] == "SUBSETS"}, key=int)
    if not plate_ids:
        sys.exit("ERROR: no plate data found in this file.")

    records, streams, plate_names = [], [], {}
    reads = None
    for pid in plate_ids:
        hdr = ole.openstream(["SUBSETS", pid, "HEADER"]).read()
        m = re.search(rb"Plate \d+", hdr)
        plate_names[pid] = m.group().decode() if m else f"Plate {pid}"
        u, vals, status = parse_plate_stream(ole.openstream(["SUBSETS", pid, "DATA"]).read())
        streams.append(u)
        if reads is None:
            reads = read_definitions(u, contents)
            if not reads:
                sys.exit("ERROR: could not find read definitions.")
        expected = 96 * sum(r["n"] for r in reads)
        if len(vals) != expected:
            sys.exit(f"ERROR: {plate_names[pid]}: found {len(vals)} data records but the protocol implies {expected}. "
                     "This file has a layout the extractor does not understand yet.")
        k = 0
        for r in reads:
            n = r["n"]
            v = vals[k:k + 96 * n].reshape(96, n)
            s = status[k:k + 96 * n].reshape(96, n)
            k += 96 * n
            for w in range(96):
                for j in range(n):
                    if s[w, j] == STATUS_NOT_READ:
                        continue
                    val = np.nan if s[w, j] == STATUS_INVALID else v[w, j]
                    wl = None if r["start"] is None else r["start"] + r["step"] * j
                    records.append((plate_names[pid], r["label"], r["unit"] or "", WELLS[w], wl, val))
    df = pd.DataFrame(records, columns=["plate", "read", "unit", "well", "wavelength_nm", "value"])
    return df, reads, experiment_info(ole, contents, streams), list(plate_names.values())


# --------------------------------------------------------------------------
# Output
# --------------------------------------------------------------------------
def spectrum_table(g):
    t = g.pivot(index="wavelength_nm", columns="well", values="value")
    t = t[sorted(t.columns, key=well_key)]
    return t.rename(columns=label_map(g))


def grid_table(g):
    t = pd.DataFrame(index=list(ROWS), columns=range(1, 13), dtype=float)
    for w, v in zip(g.well, g.value):
        t.loc[w[0], int(w[1:])] = v
    t.index.name = "row"
    return t


def table_for(g):
    return spectrum_table(g) if g.wavelength_nm.notna().any() else grid_table(g)


def add_chart(ws, ncols, nrows, title, ytitle):
    ch = LineChart()
    ch.title, ch.height, ch.width = title, 9, 22
    ch.y_axis.title, ch.x_axis.title = ytitle, "Wavelength (nm)"
    ch.add_data(Reference(ws, min_col=2, max_col=1 + ncols, min_row=1, max_row=1 + nrows), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=2, max_row=1 + nrows))
    ch.x_axis.delete = ch.y_axis.delete = False
    ws.add_chart(ch, f"{chr(65 + min(ncols, 20) + 2)}2")


def about_rows(plates_shown):
    """Field/Value table with source, plate names, well names and notes (goes on an 'About' sheet)"""
    L = LAST
    pm = L.get("plate_names", {})
    rows = [("Source file", os.path.basename(L.get("src", "")))]
    for p in L.get("raw_plates", []):
        if pm.get(p, p) in plates_shown:
            rows.append((f"Plate {p.replace('Plate ', '')}", pm.get(p, p)))
    for (p, w), n in sorted(L.get("well_names", {}).items(), key=lambda kv: (kv[0][0], well_key(kv[0][1]))):
        if pm.get(p, p) in plates_shown:
            rows.append((f"Well {pm.get(p, p)} / {w}", n))
    if L.get("notes"):
        rows.append(("Notes", L["notes"]))
    return pd.DataFrame(rows, columns=["Field", "Value"])


def write_about(xw, about):
    about.to_excel(xw, sheet_name="About", index=False)
    ws = xw.sheets["About"]
    ws.column_dimensions["A"].width = 26
    ws.column_dimensions["B"].width = 90


def write_book(path, sheets, about=None):
    """sheets: list of (sheet_name, table, is_spectrum, title, unit)"""
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        for name, t, is_spec, title, unit in sheets:
            t.to_excel(xw, sheet_name=name[:31])
            ws = xw.sheets[name[:31]]
            ws.freeze_panes = "B2"
            ws.column_dimensions["A"].width = 14
            if is_spec and t.shape[1] and t.shape[0] > 1:
                add_chart(ws, t.shape[1], t.shape[0], title, unit)
        if about is not None:
            write_about(xw, about)


def write_all(df, reads, info, plates, src, out):
    os.makedirs(out, exist_ok=True)
    for d in ("by_plate", "by_read", "csv"):
        os.makedirs(os.path.join(out, d), exist_ok=True)
    unit_of = {r["label"]: r["unit"] or "" for r in reads}
    short = lambda lab: safe(lab.replace(":", "_"))

    df.to_csv(os.path.join(out, "all_data_long.csv"), index=False)

    # by_plate: one workbook per plate, a sheet per read
    for p in plates:
        sheets = []
        for r in reads:
            g = df[(df.plate == p) & (df.read == r["label"])]
            if g.empty:
                continue
            t = table_for(g)
            sheets.append((f"R{r['index']}_{short(r['label'].split(':', 1)[1])}"[:31], t,
                           g.wavelength_nm.notna().any(), f"{p} - {r['label']}", unit_of[r["label"]]))
            t.to_csv(os.path.join(out, "csv", f"{safe(p)}_{short(r['label'])}.csv"))
        if sheets:
            write_book(os.path.join(out, "by_plate", f"{safe(p)}.xlsx"), sheets, about_rows({p}))

    # by_read: one workbook per read, a sheet per plate
    for r in reads:
        sheets = []
        for p in plates:
            g = df[(df.plate == p) & (df.read == r["label"])]
            if g.empty:
                continue
            sheets.append((safe(p), table_for(g), g.wavelength_nm.notna().any(),
                           f"{p} - {r['label']}", unit_of[r["label"]]))
        if sheets:
            write_book(os.path.join(out, "by_read", f"{short(r['label'])}.xlsx"), sheets, about_rows(set(plates)))

    # Summary workbook
    overview = (df.groupby(["plate", "read"], sort=False)
                  .agg(unit=("unit", "first"), wells=("well", "nunique"),
                       wells_list=("well", lambda s: " ".join(sorted(s.unique(), key=well_key))),
                       points=("value", "size"), invalid=("value", lambda s: int(s.isna().sum())),
                       min=("value", "min"), max=("value", "max"), mean=("value", "mean"))
                  .reset_index())
    spec = df[df.wavelength_nm.notna() & df.value.notna()]
    peaks = (spec.loc[spec.groupby(["plate", "read", "well"], sort=False).value.idxmax()]
                 [["plate", "read", "unit", "well", "sample", "wavelength_nm", "value"]]
                 .rename(columns={"wavelength_nm": "peak_wavelength_nm", "value": "peak_value"}))
    end = df[df.wavelength_nm.isna()].pivot_table(index=["plate", "well", "sample"], columns="read", values="value").reset_index()
    inf = pd.DataFrame(list({"Source file": os.path.basename(src), **info}.items()), columns=["Field", "Value"])
    with pd.ExcelWriter(os.path.join(out, "Summary.xlsx"), engine="openpyxl") as xw:
        overview.to_excel(xw, sheet_name="Overview", index=False)
        peaks.to_excel(xw, sheet_name="Peaks", index=False)
        if not end.empty:
            end.to_excel(xw, sheet_name="Endpoint_reads", index=False)
        inf.to_excel(xw, sheet_name="Info", index=False)
        for ws in xw.book.worksheets:
            ws.freeze_panes = "A2"
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = min(
                    60, max(10, max(len(str(c.value)) if c.value is not None else 0 for c in col[:50]) + 2))

    lines = [f"Extracted from: {os.path.basename(src)}", ""]
    lines += [f"{k}: {v}" for k, v in info.items()]
    lines += ["", "Reads:"]
    for r in reads:
        extra = f"  {r['n']} points, {r['start']:g}-{r['start'] + r['step'] * (r['n'] - 1):g} nm step {r['step']:g}" if r["start"] else "  endpoint"
        lines.append(f"  {r['label']} [{r['unit'] or '?'}]{extra}")
    lines += ["", "Folders:",
              "  by_plate/   one workbook per plate (sheet per read)",
              "  by_read/    one workbook per read (sheet per plate)  <- best for comparing plates",
              "  csv/        one CSV per plate+read",
              "  Summary.xlsx  Overview, Peaks, Endpoint_reads, Info",
              "  all_data_long.csv  everything, one row per measurement",
              "", "Reminders: unread wells are omitted; blank cells = invalid reading (status 100, value -636).",
              "Excitation settings for emission-scan reads are not recovered from the file."]
    open(os.path.join(out, "README.txt"), "w").write("\n".join(lines) + "\n")
    return overview


# --------------------------------------------------------------------------
# Analysis + graphs
# --------------------------------------------------------------------------
def _legend(ax, n, **kw):
    """few series: legend inside the plot; many: outside on the right so it never hides data; huge: leave it out"""
    if n > 40:
        return
    if n > 8:
        ax.legend(loc="upper left", bbox_to_anchor=(1.01, 1.0), fontsize=8, frameon=False, ncol=1, **kw)
    else:
        ax.legend(fontsize=8.5, frameon=True, framealpha=0.85, loc="upper right", **kw)


def _save(fig, path):
    """save a figure; if the user wrote notes, print them small at the bottom of every graph"""
    notes = (LAST.get("notes") or "").strip()
    if notes:
        txt = textwrap.shorten(" ".join(notes.split()), width=330, placeholder=" ...")
        fig.text(0.01, -0.012, "\n".join(textwrap.wrap("Notes: " + txt, width=150)), fontsize=8, color="#555555", va="top", ha="left")
    fig.savefig(path, dpi=DPI, bbox_inches="tight")
    plt.close(fig)


def _setup_plots():
    plt.rcParams.update({"font.size": 10, "axes.grid": True, "grid.alpha": 0.3, "axes.spines.top": False,
                         "axes.spines.right": False, "figure.facecolor": "white"})


def colors(n):
    cmap = plt.get_cmap("tab10" if n <= 10 else "tab20" if n <= 20 else "turbo")
    return [cmap(i % cmap.N if n <= 20 else i / max(n - 1, 1)) for i in range(n)]


def ylabel(unit, read):
    return {"OD": "Absorbance (OD)", "RFU": "Fluorescence (RFU)"}.get(unit, f"Signal ({unit})" if unit else "Signal")


def is_spectrum(g):
    return g.wavelength_nm.notna().any()


def usable(g):
    """drop wells whose values are all invalid"""
    ok = g.groupby("well").value.apply(lambda s: s.notna().any())
    return g[g.well.isin(ok[ok].index)]


# --------------------------------------------------------------------------
# numbers
# --------------------------------------------------------------------------
def analyze(df):
    wells, summ, means = [], [], []
    for (p, r), g in df.groupby(["plate", "read"], sort=False):
        g = usable(g)
        if g.empty:
            continue
        unit = g.unit.iloc[0]
        smp = g.drop_duplicates("well").set_index("well")["sample"].to_dict() if "sample" in g.columns else {}
        if is_spectrum(g):
            t = g.pivot(index="wavelength_nm", columns="well", values="value")
            for w in sorted(t.columns, key=well_key):
                s = t[w].dropna()
                pk = s.idxmax()
                half = s.max() / 2
                above = s[s >= half].index
                wells.append(dict(plate=p, read=r, unit=unit, well=w, sample=smp.get(w, ""), peak_wavelength_nm=pk, peak_value=s.max(),
                                  min_value=s.min(), mean_value=s.mean(),
                                  AUC=np.trapezoid(s.values, s.index.values),
                                  FWHM_nm=(above.max() - above.min()) if s.max() > 0 else np.nan,
                                  value_at_start=s.iloc[0], value_at_end=s.iloc[-1]))
            m = t.mean(axis=1)
            sd = t.std(axis=1, ddof=1) if t.shape[1] > 1 else pd.Series(0, index=t.index)
            for wl in t.index:
                means.append(dict(plate=p, read=r, unit=unit, wavelength_nm=wl, mean=m[wl], sd=sd[wl], n_wells=t.shape[1]))
        else:
            for w, v in zip(g.well, g.value):
                wells.append(dict(plate=p, read=r, unit=unit, well=w, sample=smp.get(w, ""), peak_value=v, mean_value=v))
    wdf = pd.DataFrame(wells)
    rows = []
    for (p, r), g in wdf.groupby(["plate", "read"], sort=False):
        v = g.peak_value
        d = dict(plate=p, read=r, unit=g.unit.iloc[0], n_wells=len(g), mean_peak_value=v.mean(),
                 sd_peak_value=v.std(ddof=1) if len(v) > 1 else np.nan, min=v.min(), max=v.max())
        d["CV_percent"] = 100 * d["sd_peak_value"] / v.mean() if len(v) > 1 and v.mean() else np.nan
        if "peak_wavelength_nm" in g and g.peak_wavelength_nm.notna().any():
            d["mean_peak_wavelength_nm"] = g.peak_wavelength_nm.mean()
            d["mean_AUC"] = g.AUC.mean()
        rows.append(d)
    return wdf, pd.DataFrame(rows), pd.DataFrame(means)


# --------------------------------------------------------------------------
# figures
# --------------------------------------------------------------------------
def fig_spectrum(g, plate, read, unit, path, note=""):
    t = g.pivot(index="wavelength_nm", columns="well", values="value")
    t = t[sorted(t.columns, key=well_key)]
    cs = colors(t.shape[1])
    lm = label_map(g)
    fig = plt.figure(figsize=(14, 9.5))
    gs = fig.add_gridspec(2, 2, height_ratios=[1.6, 1], hspace=0.32, wspace=0.22, top=0.92)
    ax = fig.add_subplot(gs[0, :])
    peaks = []
    for c, w in zip(cs, t.columns):
        s = t[w].dropna()
        pk = s.idxmax()
        peaks.append((lm[w], pk, s.max()))
        ax.plot(s.index, s.values, color=c, lw=1.6, label=f"{lm[w]}:  peak {s.max():.3g} @ {pk:g} nm")
        ax.plot(pk, s.max(), "o", color=c, ms=6, mec="black", mew=0.6)
    top = max(p[2] for p in peaks)
    ax.set_xlabel("Wavelength (nm)")
    ax.set_ylabel(ylabel(unit, read))
    ax.set_ylim(min(0, t.min().min() * 1.1), top * 1.12 if top > 0 else 1)
    ax.set_title(("Selected wells" if note else "All wells") + " (dots = peak of each well)", loc="left", fontweight="bold")
    _legend(ax, len(t.columns), title="Well")

    ax2 = fig.add_subplot(gs[1, 0])
    m = t.mean(axis=1)
    sd = t.std(axis=1, ddof=1) if t.shape[1] > 1 else pd.Series(0, index=t.index)
    ax2.plot(m.index, m.values, color="black", lw=1.8, label="mean")
    ax2.fill_between(m.index, m - sd, m + sd, color="tab:blue", alpha=0.25, label="± 1 SD")
    mp = m.idxmax()
    ax2.axvline(mp, color="red", ls="--", lw=1)
    ax2.annotate(f"mean peak {m.max():.3g} @ {mp:g} nm", (mp, m.max()), textcoords="offset points",
                 xytext=(8, -4), color="red", fontsize=9)
    ax2.set_xlabel("Wavelength (nm)")
    ax2.set_ylabel(ylabel(unit, read))
    ax2.set_title(f"Mean ± SD of {t.shape[1]} wells" if t.shape[1] > 1 else "Single well (no SD)", loc="left")
    ax2.legend(frameon=False, fontsize=8)

    ax3 = fig.add_subplot(gs[1, 1])
    ws = [p[0] for p in peaks]
    bars = ax3.bar(ws, [p[2] for p in peaks], color=cs, edgecolor="black", lw=0.5)
    for b, p in zip(bars, peaks):
        ax3.text(b.get_x() + b.get_width() / 2, b.get_height(), f"{p[1]:g} nm", ha="center", va="bottom", fontsize=8)
    ax3.set_ylabel(f"Peak {ylabel(unit, read)}")
    ax3.set_title("Peak value per well (label = peak wavelength)", loc="left")
    ax3.set_ylim(0, max(p[2] for p in peaks) * 1.18 if top > 0 else 1)
    ax3.tick_params(axis="x", rotation=45)

    fig.suptitle(f"{plate}   |   {read}   |   {t.index.min():g}-{t.index.max():g} nm" + (f"   |   {note}" if note else ""), fontsize=13, y=0.97)
    _save(fig, path)


def fig_endpoint(g, plate, read, unit, path, note=""):
    g = g.copy()
    g = g.sort_values("well", key=lambda s: s.map(well_key))
    fig, (ax, ax2) = plt.subplots(1, 2, figsize=(14, 5.2), gridspec_kw=dict(width_ratios=[1.1, 1]))
    cs = colors(len(g))
    lm = label_map(g)
    bars = ax.bar([lm[w] for w in g.well], g.value, color=cs, edgecolor="black", lw=0.5)
    for b, v in zip(bars, g.value):
        ax.text(b.get_x() + b.get_width() / 2, b.get_height(), f"{v:g}", ha="center", va="bottom", fontsize=9)
    mu = g.value.mean()
    ax.axhline(mu, color="red", ls="--", lw=1, label=f"mean = {mu:.4g}")
    ax.set_ylabel(ylabel(unit, read))
    ax.set_title("Value per well", loc="left")
    ax.legend(frameon=False)
    ax.set_ylim(0, g.value.max() * 1.15 if g.value.max() > 0 else 1)
    ax.tick_params(axis="x", rotation=45)
    grid = np.full((8, 12), np.nan)
    for w, v in zip(g.well, g.value):
        grid[ROWS.index(w[0]), int(w[1:]) - 1] = v
    im = ax2.imshow(grid, cmap="viridis", aspect="equal")
    ax2.grid(False)
    ax2.set_xticks(range(12), range(1, 13))
    ax2.set_yticks(range(8), list(ROWS))
    ax2.xaxis.tick_top()
    for i in range(8):
        for j in range(12):
            if not np.isnan(grid[i, j]):
                ax2.text(j, i, f"{grid[i, j]:g}", ha="center", va="center", fontsize=8,
                         color="white" if grid[i, j] < np.nanmean(grid) else "black")
    fig.colorbar(im, ax=ax2, shrink=0.8, label=ylabel(unit, read))
    ax2.set_title("Plate layout (unread wells blank)", loc="left", pad=22)
    fig.suptitle(f"{plate}   |   {read}   |   endpoint" + (f"   |   {note}" if note else ""), fontsize=13)
    fig.tight_layout()
    _save(fig, path)


def fig_read_across_plates(df, read, unit, plates, path, note=""):
    d = df[df.read == read]
    if is_spectrum(d):
        fig, axes = plt.subplots(1, 2, figsize=(15, 5.8))
        cs = colors(len(plates))
        for c, p in zip(cs, plates):
            g = usable(d[d.plate == p])
            if g.empty:
                continue
            t = g.pivot(index="wavelength_nm", columns="well", values="value")
            m = t.mean(axis=1)
            sd = t.std(axis=1, ddof=1) if t.shape[1] > 1 else pd.Series(0, index=t.index)
            axes[0].plot(m.index, m, color=c, lw=1.8, label=f"{p} (n={t.shape[1]})")
            axes[0].fill_between(m.index, m - sd, m + sd, color=c, alpha=0.18)
            pk = m.idxmax()
            axes[0].plot(pk, m.max(), "o", color=c, mec="black", ms=6)
            for w in t.columns:
                axes[1].plot(t.index, t[w], color=c, lw=0.8, alpha=0.7)
        axes[0].set_title("Mean ± SD per plate", loc="left")
        axes[1].set_title("Every well, coloured by plate", loc="left")
        for a in axes:
            a.set_xlabel("Wavelength (nm)")
            a.set_ylabel(ylabel(unit, read))
        axes[0].legend(frameon=False)
    else:
        fig, ax = plt.subplots(figsize=(15, 5.5))
        pv = usable(d).pivot_table(index="well", columns="plate", values="value")
        pv = pv.reindex(sorted(pv.index, key=well_key))
        pv.plot.bar(ax=ax, edgecolor="black", lw=0.4, width=0.8, color=colors(len(pv.columns)))
        ax.set_ylabel(ylabel(unit, read))
        ax.set_xlabel("Well")
        ax.set_title("Value per well, grouped by plate", loc="left")
        ax.tick_params(axis="x", rotation=0)
        ax.legend(frameon=False, title="Plate")
    fig.suptitle(f"{read} - comparison across plates" + (f"   |   {note}" if note else ""), fontsize=13)
    fig.tight_layout()
    _save(fig, path)


def fig_plate_overview(df, plate, reads, path):
    d = df[df.plate == plate]
    fig, axes = plt.subplots(len(reads), 1, figsize=(13, 3.6 * len(reads)))
    axes = np.atleast_1d(axes)
    for ax, r in zip(axes, reads):
        g = usable(d[d.read == r["label"]])
        unit = r["unit"] or ""
        if g.empty:
            ax.set_visible(False)
            continue
        if is_spectrum(g):
            t = g.pivot(index="wavelength_nm", columns="well", values="value")
            for c, w in zip(colors(t.shape[1]), sorted(t.columns, key=well_key)):
                ax.plot(t.index, t[w], color=c, lw=1.2, label=label_map(g)[w])
            ax.set_xlabel("Wavelength (nm)")
            ax.legend(ncol=8, fontsize=7, frameon=False)
        else:
            g = g.sort_values("well", key=lambda s: s.map(well_key))
            ax.bar([label_map(g)[w] for w in g.well], g.value, color=colors(len(g)), edgecolor="black", lw=0.4)
        ax.set_ylabel(ylabel(unit, r["label"]))
        ax.set_title(r["label"], loc="left", fontweight="bold")
    fig.suptitle(f"{plate} - overview of all reads", fontsize=13, y=1.0)
    fig.tight_layout()
    _save(fig, path)


# --------------------------------------------------------------------------
# --------------------------------------------------------------------------
# Graph catalog: every graph type the user can choose
# --------------------------------------------------------------------------
GRAPH_CATALOG = [   # (id, label, applies to)
    ("dashboard", "Combined dashboard (all wells + mean ± SD + peaks)", "both"),
    ("overlay", "All wells overlaid, peaks marked", "spectrum"),
    ("mean_sd", "Mean ± SD", "spectrum"),
    ("normalized", "Normalized overlay (each well scaled to its own peak)", "spectrum"),
    ("peak_bar", "Peak value per well (bar chart)", "spectrum"),
    ("auc_bar", "Area under the curve per well (bar chart)", "spectrum"),
    ("heatmap", "Heatmap (wavelength x well)", "spectrum"),
    ("bar", "Value per well (bar chart)", "endpoint"),
    ("plate_map", "Plate map (8x12 heatmap)", "endpoint"),
    ("compare", "Comparison across plates (one per read)", "extra"),
    ("overview", "Plate overview (all reads on one page)", "extra"),
]
GRAPH_IDS = [c[0] for c in GRAPH_CATALOG]
DEFAULT_GRAPHS = ["dashboard", "compare", "overview"]


def graph_catalog():
    return [dict(id=i, label=l, applies=a) for i, l, a in GRAPH_CATALOG]


def parse_graphs(spec):
    """'overlay,mean_sd' | 'all' | '' -> list of ids (None = defaults)."""
    if spec is None or not str(spec).strip():
        return None
    if str(spec).strip().lower() == "all":
        return list(GRAPH_IDS)
    kinds = [k.strip().lower() for k in re.split(r"[,\s;]+", str(spec)) if k.strip()]
    bad = [k for k in kinds if k not in GRAPH_IDS]
    if bad:
        sys.exit(f"ERROR: unknown graph type(s): {', '.join(bad)}.\nChoose from: {', '.join(GRAPH_IDS)}  (or 'all'). Use --list-graphs to see what each does.")
    return kinds


def graphs_for(spectrum):
    """ids that make sense for one plate x read (spectrum or endpoint)"""
    want = "spectrum" if spectrum else "endpoint"
    return [i for i, _, a in GRAPH_CATALOG if a in ("both", want)]


def series_table(g):
    """wavelength x series table; series are wells (or 'Plate N well' when several plates are mixed)"""
    multi = g.plate.nunique() > 1
    cols = {}
    for (p, w), gg in sorted(g.groupby(["plate", "well"]), key=lambda kv: (kv[0][0], well_key(kv[0][1]))):
        lab = wlab(w, gg["sample"].iloc[0] if "sample" in gg.columns else "")
        cols[f"{p} {lab}" if multi else lab] = gg.set_index("wavelength_nm").value
    return pd.DataFrame(cols)


def endpoint_series(g):
    multi = g.plate.nunique() > 1
    smp = g["sample"] if "sample" in g.columns else [""] * len(g)
    rows = sorted(zip(g.plate, g.well, g.value, smp), key=lambda x: (x[0], well_key(x[1])))
    return [(f"{p} {wlab(w, sm)}" if multi else wlab(w, sm), v) for p, w, v, sm in rows]


def _single_spectrum(kind, g, plate, read, unit, path, note=""):
    t = series_table(g)
    n = t.shape[1]
    cs = colors(n)
    yl = ylabel(unit, read)
    label = dict((i, l) for i, l, _ in GRAPH_CATALOG)[kind]
    multi = g.plate.nunique() > 1
    if kind == "mean_sd" and multi:                       # several plates: mean per plate
        fig_read_across_plates(g, read, unit, sorted(g.plate.unique()), path, note=note)
        return
    height = max(4.5, 0.42 * n + 2.6) if kind == "heatmap" else 6.2
    fig, ax = plt.subplots(figsize=(13, height))
    if kind == "overlay":
        for c, w in zip(cs, t.columns):
            s = t[w].dropna()
            pk = s.idxmax()
            ax.plot(s.index, s.values, color=c, lw=1.6, label=f"{w}:  peak {s.max():.3g} @ {pk:g} nm")
            ax.plot(pk, s.max(), "o", color=c, ms=6, mec="black", mew=0.6)
        ax.set_ylabel(yl)
        _legend(ax, n)
    elif kind == "mean_sd":
        m = t.mean(axis=1)
        sd = t.std(axis=1, ddof=1) if n > 1 else pd.Series(0, index=t.index)
        ax.plot(m.index, m.values, color="black", lw=1.8, label=f"mean of {n} well{'s' if n > 1 else ''}")
        if n > 1:
            ax.fill_between(m.index, m - sd, m + sd, color="tab:blue", alpha=0.25, label="± 1 SD")
        mp = m.idxmax()
        ax.axvline(mp, color="red", ls="--", lw=1)
        ax.annotate(f"mean peak {m.max():.3g} @ {mp:g} nm", (mp, m.max()), textcoords="offset points",
                    xytext=(8, -4), color="red", fontsize=10)
        ax.set_ylabel(yl)
        ax.legend(frameon=False)
    elif kind == "normalized":
        for c, w in zip(cs, t.columns):
            s = t[w].dropna()
            mx = s.max()
            if mx > 0:
                ax.plot(s.index, s.values / mx, color=c, lw=1.5, label=f"{w}:  peak @ {s.idxmax():g} nm")
        ax.set_ylabel("Normalized signal (each well's peak = 1)")
        _legend(ax, n)
    elif kind in ("peak_bar", "auc_bar"):
        if kind == "peak_bar":
            vals = [t[w].max() for w in t.columns]
            marks = [f"{t[w].idxmax():g} nm" for w in t.columns]
            ax.set_ylabel(f"Peak {yl}")
        else:
            vals = [np.trapezoid(t[w].dropna().values, t[w].dropna().index.values) for w in t.columns]
            marks = [f"{v:.4g}" for v in vals]
            ax.set_ylabel(f"Area under curve ({unit or 'signal'} x nm)")
        bars = ax.bar(list(t.columns), vals, color=cs, edgecolor="black", lw=0.5)
        for b, mk in zip(bars, marks):
            ax.text(b.get_x() + b.get_width() / 2, b.get_height(), mk, ha="center", va="bottom", fontsize=8.5)
        ax.set_ylim(0, max(vals) * 1.15 if max(vals) > 0 else 1)
        ax.tick_params(axis="x", rotation=45 if n > 6 else 0)
    elif kind == "heatmap":
        im = ax.imshow(t.T.values.astype(float), aspect="auto", cmap="viridis",
                       extent=[t.index.min(), t.index.max(), n, 0])
        ax.set_yticks(np.arange(n) + 0.5, list(t.columns))
        ax.grid(False)
        fig.colorbar(im, ax=ax, label=yl, pad=0.01)
    if kind not in ("peak_bar", "auc_bar"):
        ax.set_xlabel("Wavelength (nm)")
    fig.suptitle(f"{plate}   |   {read}   |   {label}" + (f"   |   {note}" if note else ""), fontsize=13)
    fig.tight_layout()
    _save(fig, path)


def _single_endpoint(kind, g, plate, read, unit, path, note=""):
    yl = ylabel(unit, read)
    label = dict((i, l) for i, l, _ in GRAPH_CATALOG)[kind]
    if kind == "bar":
        ser = endpoint_series(g)
        fig, ax = plt.subplots(figsize=(13, 5.5))
        bars = ax.bar([s[0] for s in ser], [s[1] for s in ser], color=colors(len(ser)), edgecolor="black", lw=0.5)
        for b, (_, v) in zip(bars, ser):
            ax.text(b.get_x() + b.get_width() / 2, b.get_height(), f"{v:g}", ha="center", va="bottom", fontsize=9)
        mu = np.mean([s[1] for s in ser])
        ax.axhline(mu, color="red", ls="--", lw=1, label=f"mean = {mu:.4g}")
        ax.set_ylabel(yl)
        ax.legend(frameon=False)
        ax.set_ylim(0, max(s[1] for s in ser) * 1.15 if max(s[1] for s in ser) > 0 else 1)
        ax.tick_params(axis="x", rotation=45 if len(ser) > 6 else 0)
    else:                                                  # plate_map: one 8x12 map per plate, shared colour scale
        plates = sorted(g.plate.unique())
        fig, axes = plt.subplots(1, len(plates), figsize=(7.5 * len(plates), 5.2), squeeze=False)
        lo, hi = g.value.min(), g.value.max()
        for ax, p in zip(axes[0], plates):
            grid = np.full((8, 12), np.nan)
            for w, v in zip(g[g.plate == p].well, g[g.plate == p].value):
                grid[ROWS.index(w[0]), int(w[1:]) - 1] = v
            im = ax.imshow(grid, cmap="viridis", vmin=lo, vmax=hi, aspect="equal")
            ax.grid(False)
            ax.set_xticks(range(12), range(1, 13))
            ax.set_yticks(range(8), list(ROWS))
            ax.xaxis.tick_top()
            for i in range(8):
                for j in range(12):
                    if not np.isnan(grid[i, j]):
                        ax.text(j, i, f"{grid[i, j]:g}", ha="center", va="center", fontsize=8,
                                color="white" if grid[i, j] < (lo + hi) / 2 else "black")
            if len(plates) > 1:
                ax.set_title(p, loc="left", pad=22)
        fig.colorbar(im, ax=axes.ravel().tolist(), shrink=0.8, label=yl)
    fig.suptitle(f"{plate}   |   {read}   |   {label}" + (f"   |   {note}" if note else ""), fontsize=13)
    _save(fig, path)


def draw_graph(kind, g, plate, read, unit, path, note=""):
    """Draw one graph type for one plate x read. Returns False if the type does not apply to this read."""
    spec = is_spectrum(g)
    if kind == "dashboard":
        (fig_spectrum if spec else fig_endpoint)(g, plate, read, unit, path, note=note)
    elif spec and kind in graphs_for(True):
        _single_spectrum(kind, g, plate, read, unit, path, note)
    elif not spec and kind in graphs_for(False):
        _single_endpoint(kind, g, plate, read, unit, path, note)
    else:
        return False
    return True


def run_analysis(df, reads, plates, out, graphs=None):
    _setup_plots()
    gdir = os.path.join(out, "graphs")
    for d in ("by_plate_read", "by_read", "by_plate"):
        os.makedirs(os.path.join(gdir, d), exist_ok=True)
    unit_of = {r["label"]: r["unit"] or "" for r in reads}
    kinds = graphs or DEFAULT_GRAPHS
    n = 0
    for p in plates:
        for r in reads:
            g = usable(df[(df.plate == p) & (df.read == r["label"])])
            if g.empty:
                continue
            for k in kinds:
                if k in ("compare", "overview"):
                    continue
                suffix = "" if k == "dashboard" else f"__{k}"
                folder = os.path.join(gdir, "by_plate_read", safe(p))
                os.makedirs(folder, exist_ok=True)
                f = os.path.join(folder, f"{safe(p)}__{safe(r['label'].replace(':', '_'))}{suffix}.png")
                if draw_graph(k, g, p, r["label"], unit_of[r["label"]], f):
                    n += 1
        if "overview" in kinds:
            fig_plate_overview(df, p, reads, os.path.join(gdir, "by_plate", f"{safe(p)}_overview.png"))
    if "compare" in kinds:
        for r in reads:
            fig_read_across_plates(df, r["label"], unit_of[r["label"]], plates,
                                   os.path.join(gdir, "by_read", f"{safe(r['label'].replace(':', '_'))}__all_plates.png"))
    wdf, sdf, mdf = analyze(df)
    with pd.ExcelWriter(os.path.join(out, "Analysis.xlsx"), engine="openpyxl") as xw:
        sdf.to_excel(xw, sheet_name="Plate_Read_stats", index=False)
        wdf.to_excel(xw, sheet_name="Well_stats", index=False)
        mdf.to_excel(xw, sheet_name="Spectrum_means", index=False)
        for ws in xw.book.worksheets:
            ws.freeze_panes = "A2"
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = min(
                    40, max(10, max(len(str(c.value)) if c.value is not None else 0 for c in col[:60]) + 2))
    return n, sdf


LAST = {}          # last extraction (used by the website for interactive well selection)


def custom_view(df, reads, plates, read_label, wells, outdir, kind="dashboard"):
    """Graph + table for chosen wells of one read, for one plate or several plates side by side.
    Returns dict(png, csv, stats) or None when nothing matches."""
    d = usable(df[df.read.eq(read_label) & df.plate.isin(plates) & df.well.isin(wells)])
    if d.empty:
        return None
    os.makedirs(outdir, exist_ok=True)
    unit = d.unit.iloc[0]
    sel = sorted(d.well.unique(), key=well_key)
    n_series = d.groupby(["plate", "well"]).ngroups
    if d.plate.nunique() == 1:
        note = f"wells: {', '.join(sel) if len(sel) <= 10 else str(len(sel)) + ' wells'}"
    else:
        note = f"{n_series} wells on {d.plate.nunique()} plates"
    png = os.path.join(outdir, "selection.png")
    present = [p for p in plates if p in set(d.plate)]
    if kind not in graphs_for(is_spectrum(d)) or kind in ("compare", "overview"):
        kind = "dashboard"                                 # type does not apply to this read
    if kind == "dashboard" and len(present) > 1:
        fig_read_across_plates(d, read_label, unit, present, png, note=note)
    else:
        draw_graph(kind, d, present[0] if len(present) == 1 else "Plates " + ", ".join(p.replace("Plate ", "") for p in present),
                   read_label, unit, png, note=note)
    csv = os.path.join(outdir, "selection.csv")
    d.to_csv(csv, index=False)
    wdf, _, _ = analyze(d)
    xlsx = os.path.join(outdir, "selection.xlsx")
    with pd.ExcelWriter(xlsx, engine="openpyxl") as xw:
        d.to_excel(xw, sheet_name="Data", index=False)
        wdf.to_excel(xw, sheet_name="Stats", index=False)
        write_about(xw, about_rows(set(present)))
    return dict(png=png, csv=csv, xlsx=xlsx, stats=wdf, kind=kind)


def set_labels(plate_names=None, well_spec=None, notes=None):
    """text form (command line / window): plate names list or 'a, b, c'; well spec 'A1 = Blank; ...'; notes text"""
    raw = LAST["raw_plates"]
    pm = {p: n for p, n in zip(raw, parse_plate_names(plate_names) or []) if n}
    set_labels_struct(pm, parse_well_names(well_spec, raw, pm), notes)


def set_labels_struct(plate_map, triples, notes=None):
    """plate_map {original plate: name}; triples [(original plate, well, name)]; replaces all labels at once"""
    raw = LAST["raw_plates"]
    pm = {p: str(n).strip()[:40] for p, n in (plate_map or {}).items() if p in raw and str(n).strip()}
    finals = [pm.get(p, p).lower() for p in raw]
    if len(set(finals)) != len(finals):
        sys.exit("ERROR: every plate needs a different name.")
    have = set(zip(LAST["raw_df"].plate, LAST["raw_df"].well))      # only wells that actually have data can be named
    wn = {}
    for p, w, n in triples or []:
        if (p, w) in have and str(n).strip():
            wn[(p, w)] = str(n).strip()[:60]
    LAST.update(plate_names=pm, well_names=wn, notes=(notes or "").strip())


def labeled():
    """the extracted data with the current plate names / well names applied -> (df, plates, info)"""
    L = LAST
    df = L["raw_df"].copy()
    wn, pm = L["well_names"], L["plate_names"]
    df["sample"] = [wn.get((p, w), "") for p, w in zip(df.plate, df.well)]
    df["plate"] = df["plate"].replace(pm)
    df = df[["plate", "read", "unit", "well", "sample", "wavelength_nm", "value"]]
    plates = [pm.get(p, p) for p in L["raw_plates"]]
    info = dict(L["base_info"])
    if pm:
        info["Plate names"] = "; ".join(f"{k} = {v}" for k, v in pm.items())
    if wn:
        info["Well names"] = "; ".join(f"{pm.get(p, p)} {w} = {n}" for (p, w), n in
                                       sorted(wn.items(), key=lambda kv: (kv[0][0], well_key(kv[0][1]))))
    if L["notes"]:
        info["Notes"] = L["notes"]
    return df, plates, info


def regenerate(out, plots=True, log=print, sel=None, graphs=None):
    """(re)write every output file from the stored data using the current labels"""
    df, plates, info = labeled()
    reads = LAST["reads"]
    if sel:
        df = df[df.well.isin(sel)]
        if df.empty:
            sys.exit(f"ERROR: none of the selected wells ({', '.join(sel)}) have data in this file.")
        log(f"Well filter: {len(sel)} wells selected ({', '.join(sel) if len(sel) <= 12 else ', '.join(sel[:12]) + ', ...'})")
    log(f"Found {len(plates)} plates, {len(reads)} reads, {len(df):,} measurements. Writing files ...")
    ov = write_all(df, reads, info, plates, LAST["src"], out)
    if LAST["notes"]:
        with open(os.path.join(out, "Notes.txt"), "w") as fh:
            fh.write(LAST["notes"] + "\n")
    if sel:
        with open(os.path.join(out, "README.txt"), "a") as fh:
            fh.write(f"\nWell filter applied: {', '.join(sel)}\n")
    if plots and (graphs is None or graphs):
        if plt is None:
            sys.exit("Graphs need matplotlib:  pip install matplotlib   (or turn graphs off)")
        log("Analysing and drawing graphs ...")
        n, _ = run_analysis(df, reads, plates, out, graphs if isinstance(graphs, list) else parse_graphs(graphs))
        log(f"{n} plate x read graphs written (plus any comparison/overview graphs chosen).")
    return ov


def process(path, out=None, plots=True, log=print, wells=None, graphs=None, plate_names=None, notes=None, well_names=None):
    """Whole pipeline: extract -> apply names/notes -> workbooks/CSVs -> (optional) analysis + graphs."""
    if not os.path.isfile(path):
        sys.exit(f"ERROR: file not found: {path}")
    out = out or os.path.splitext(path)[0] + "_extracted"
    log(f"Reading {os.path.basename(path)} ...")
    df, reads, info, plates = extract(path)
    LAST.clear()
    LAST.update(raw_df=df, reads=reads, raw_plates=plates, base_info=info, src=path, plate_names={}, well_names={}, notes="")
    set_labels(plate_names, well_names, notes)
    if LAST["plate_names"]:
        log("Plate names: " + "; ".join(f"{k} = {v}" for k, v in LAST["plate_names"].items()))
    if LAST["well_names"]:
        log(f"{len(LAST['well_names'])} well name(s) set.")
    sel = parse_wells(wells) if isinstance(wells, str) else (sorted(wells, key=well_key) if wells else None)
    ov = regenerate(out, plots, log, sel, graphs)
    log(f"Done -> {out}")
    return out, ov


def main():
    ap = argparse.ArgumentParser(description="Extract BioTek Gen5 .xpt files into organized Excel/CSV files, with analysis and graphs.")
    ap.add_argument("file", nargs="?", help="path to the .xpt file (omit to open the window)")
    ap.add_argument("-o", "--out", help="output folder (default: <file name>_extracted)")
    ap.add_argument("--no-plots", action="store_true", help="skip analysis and graphs")
    ap.add_argument("--wells", help="only these wells, e.g. \"A1\" or \"A1,B2\" or \"A1:C3\" or \"A\" (row) or \"2\" (column)")
    ap.add_argument("--graphs", help="graph types, comma separated, or 'all' (default: dashboard,compare,overview). See --list-graphs")
    ap.add_argument("--list-graphs", action="store_true", help="show the available graph types and exit")
    ap.add_argument("--plate-names", help="names for the plates, in order, comma separated, e.g. \"Control,Treated 1h,Treated 24h\"")
    ap.add_argument("--well-names", help="names for wells, one plate at a time, e.g. \"Plate 1 | A1=Blank; Plate 1 | A2:A4=Sample X; Plate 2 | A1=Control\"")
    ap.add_argument("--notes", help="path to a .txt file with your notes (saved as Notes.txt and added to the README)")
    ap.add_argument("--gui", action="store_true", help="open the window")
    a, _ = ap.parse_known_args()          # ignore extra args the macOS launcher may add
    if a.list_graphs:
        for i, l, ap_to in GRAPH_CATALOG:
            print(f"  {i:<11} {l}   [{ap_to}]")
        return
    if a.gui or not a.file:
        return gui()
    notes = None
    if a.notes:
        if not os.path.isfile(a.notes):
            sys.exit(f"ERROR: notes file not found: {a.notes}")
        notes = open(a.notes, encoding="utf-8", errors="replace").read()
    out, ov = process(a.file, a.out, not a.no_plots, wells=a.wells, graphs=parse_graphs(a.graphs),
                      plate_names=a.plate_names, notes=notes, well_names=a.well_names)
    print()
    print(ov[["plate", "read", "unit", "wells", "points", "invalid"]].to_string(index=False))


# --------------------------------------------------------------------------
# Window (PySide6)
# --------------------------------------------------------------------------
def gui():
    try:
        from PySide6 import QtCore, QtGui, QtWidgets
    except ImportError:
        path = input("Window needs PySide6 (pip install PySide6). Enter the .xpt file name instead: ").strip().strip("'\"")
        out, ov = process(path)
        print(ov[["plate", "read", "unit", "wells", "points", "invalid"]].to_string(index=False))
        return

    class Worker(QtCore.QThread):
        message = QtCore.Signal(str)
        finished_ok = QtCore.Signal(str)
        failed = QtCore.Signal(str)

        def __init__(self, path, out, plots, wells=None, graphs=None, names=None, notes=None, well_names=None):
            super().__init__()
            self.path, self.out, self.plots, self.wells, self.graphs = path, out, plots, wells, graphs
            self.names, self.notes, self.well_names = names, notes, well_names

        def run(self):
            try:
                out, _ = process(self.path, self.out, self.plots, log=self.message.emit, wells=self.wells, graphs=self.graphs,
                                 plate_names=self.names, notes=self.notes, well_names=self.well_names)
                self.finished_ok.emit(out)
            except SystemExit as e:               # extract() reports problems with sys.exit(msg)
                self.failed.emit(str(e.code))
            except Exception as e:                # anything unexpected
                self.failed.emit(f"{type(e).__name__}: {e}")

    class Window(QtWidgets.QWidget):
        def __init__(self):
            super().__init__()
            self.setWindowTitle("Gen5 Data Extractor")
            self.resize(1180, 760)
            self.setAcceptDrops(True)
            self.out = None
            self.worker = None

            self.file_edit = QtWidgets.QLineEdit()
            self.file_edit.setPlaceholderText("Choose or drop a Gen5 .xpt file here")
            browse = QtWidgets.QPushButton("Browse...")
            browse.clicked.connect(self.browse)
            self.out_edit = QtWidgets.QLineEdit()
            self.out_edit.setPlaceholderText("Output folder (default: next to the file)")
            out_btn = QtWidgets.QPushButton("Change...")
            out_btn.clicked.connect(self.pick_out)
            self.wells_edit = QtWidgets.QLineEdit()
            self.wells_edit.setPlaceholderText("All wells.  Or e.g.  A2   |   A1,B2,C3   |   A1:D2   |   A (whole row)   |   3 (whole column)")
            self.names_edit = QtWidgets.QLineEdit()
            self.names_edit.setPlaceholderText("Optional plate names in order, e.g.  Control, Treated 1h, Treated 24h")
            self.wnames_edit = QtWidgets.QLineEdit()
            self.wnames_edit.setPlaceholderText("Optional well names per plate, e.g.  Plate 1 | A1 = Blank; Plate 1 | A2:A4 = Sample X; Plate 2 | A1 = Control")
            self.notes_edit = QtWidgets.QLineEdit()
            self.notes_edit.setPlaceholderText("Optional notes file (.txt)")
            notes_btn = QtWidgets.QPushButton("Choose...")
            notes_btn.clicked.connect(self.pick_notes)
            self.gbox = QtWidgets.QGroupBox("Graphs to create")
            gl = QtWidgets.QGridLayout(self.gbox)
            self.gchecks = {}
            for k, (gid, glabel, applies) in enumerate(GRAPH_CATALOG):
                cb = QtWidgets.QCheckBox(glabel)
                cb.setChecked(gid in DEFAULT_GRAPHS)
                self.gchecks[gid] = cb
                gl.addWidget(cb, k // 2, k % 2)
            self.plots = QtWidgets.QCheckBox("Create analysis and graphs")
            self.plots.toggled.connect(self.gbox.setEnabled)
            self.plots.setChecked(plt is not None)
            self.plots.setEnabled(plt is not None)
            self.go = QtWidgets.QPushButton("Extract")
            self.go.setDefault(True)
            self.go.setMinimumHeight(34)
            self.go.clicked.connect(self.start)
            self.open_btn = QtWidgets.QPushButton("Open output folder")
            self.open_btn.setEnabled(False)
            self.open_btn.clicked.connect(self.open_folder)
            self.bar = QtWidgets.QProgressBar()
            self.bar.setRange(0, 1)
            self.bar.setTextVisible(False)
            self.log = QtWidgets.QPlainTextEdit(readOnly=True)
            self.log.setMaximumHeight(120)

            self.tree = QtWidgets.QTreeWidget()
            self.tree.setHeaderLabel("Results  (double-click a file to open it)")
            self.tree.setMinimumWidth(330)
            self.tree.currentItemChanged.connect(self.preview)
            self.tree.itemDoubleClicked.connect(self.open_item)
            self.img = QtWidgets.QLabel("Graphs appear here after extraction.\nSelect a graph on the left to preview it.")
            self.img.setAlignment(QtCore.Qt.AlignCenter)
            self.img.setMinimumSize(500, 400)
            scroll = QtWidgets.QScrollArea()
            scroll.setWidgetResizable(True)
            scroll.setWidget(self.img)
            split = QtWidgets.QSplitter()
            split.addWidget(self.tree)
            split.addWidget(scroll)
            split.setStretchFactor(1, 1)

            top = QtWidgets.QGridLayout()
            top.addWidget(QtWidgets.QLabel("Gen5 file"), 0, 0)
            top.addWidget(self.file_edit, 0, 1)
            top.addWidget(browse, 0, 2)
            top.addWidget(QtWidgets.QLabel("Output"), 1, 0)
            top.addWidget(self.out_edit, 1, 1)
            top.addWidget(out_btn, 1, 2)
            top.addWidget(QtWidgets.QLabel("Wells"), 2, 0)
            top.addWidget(self.wells_edit, 2, 1, 1, 2)
            top.addWidget(QtWidgets.QLabel("Plate names"), 3, 0)
            top.addWidget(self.names_edit, 3, 1, 1, 2)
            top.addWidget(QtWidgets.QLabel("Well names"), 4, 0)
            top.addWidget(self.wnames_edit, 4, 1, 1, 2)
            top.addWidget(QtWidgets.QLabel("Notes"), 5, 0)
            top.addWidget(self.notes_edit, 5, 1)
            top.addWidget(notes_btn, 5, 2)
            row = QtWidgets.QHBoxLayout()
            row.addWidget(self.plots)
            row.addStretch()
            row.addWidget(self.open_btn)
            row.addWidget(self.go)
            lay = QtWidgets.QVBoxLayout(self)
            lay.addLayout(top)
            lay.addWidget(self.gbox)
            lay.addLayout(row)
            lay.addWidget(self.bar)
            lay.addWidget(self.log)
            lay.addWidget(split, 1)
            self._pix = None

        # -- input -----------------------------------------------------
        def dragEnterEvent(self, e):
            if e.mimeData().hasUrls():
                e.acceptProposedAction()

        def dropEvent(self, e):
            for u in e.mimeData().urls():
                if u.isLocalFile():
                    self.file_edit.setText(u.toLocalFile())
                    break

        def browse(self):
            f, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Choose Gen5 file", "", "Gen5 files (*.xpt);;All files (*)")
            if f:
                self.file_edit.setText(f)

        def pick_notes(self):
            f, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Choose notes file", "", "Text files (*.txt);;All files (*)")
            if f:
                self.notes_edit.setText(f)

        def pick_out(self):
            d = QtWidgets.QFileDialog.getExistingDirectory(self, "Choose output folder")
            if d:
                self.out_edit.setText(d)

        # -- run -------------------------------------------------------
        def say(self, text):
            self.log.appendPlainText(text)

        def start(self):
            path = self.file_edit.text().strip().strip("'\"")
            if not path:
                self.say("Choose a .xpt file first.")
                return
            self.log.clear()
            self.tree.clear()
            self.img.setText("Working ...")
            self.go.setEnabled(False)
            self.open_btn.setEnabled(False)
            self.bar.setRange(0, 0)           # busy indicator
            notes = None
            nf = self.notes_edit.text().strip().strip("'\"")
            if nf:
                try:
                    notes = open(nf, encoding="utf-8", errors="replace").read()
                except OSError as e:
                    self.say(f"Cannot read notes file: {e}")
                    return
            self.worker = Worker(path, self.out_edit.text().strip() or None, self.plots.isChecked(), self.wells_edit.text().strip() or None,
                                 [k for k, c in self.gchecks.items() if c.isChecked()],
                                 self.names_edit.text().strip() or None, notes,
                                 self.wnames_edit.text().strip() or None)
            self.worker.message.connect(self.say)
            self.worker.finished_ok.connect(self.done)
            self.worker.failed.connect(self.fail)
            self.worker.start()

        def fail(self, msg):
            self.bar.setRange(0, 1)
            self.go.setEnabled(True)
            self.img.setText("Extraction failed.")
            self.say(msg)

        def done(self, out):
            self.bar.setRange(0, 1)
            self.bar.setValue(1)
            self.go.setEnabled(True)
            self.open_btn.setEnabled(True)
            self.out = out
            self.fill_tree(out)
            first = self.first_graph
            if first:
                self.tree.setCurrentItem(first)
            else:
                self.img.setText("Finished. Open the output folder to see the files.")

        # -- results browser -------------------------------------------
        def fill_tree(self, out):
            self.first_graph = None
            labels = [("graphs/by_plate_read", "Graphs: plate x read"), ("graphs/by_read", "Graphs: read across plates"),
                      ("graphs/by_plate", "Graphs: plate overview"), ("by_read", "Excel: by read"),
                      ("by_plate", "Excel: by plate"), ("csv", "CSV files"), (".", "Summary and analysis")]
            for sub, title in labels:
                base = os.path.join(out, sub)
                if not os.path.isdir(base):
                    continue
                top = QtWidgets.QTreeWidgetItem([title])
                found = False
                for root, _, files in sorted(os.walk(base)):
                    if sub == "." and root != base:
                        continue
                    if sub != "." and sub in ("by_read", "by_plate") and root != base:
                        continue
                    for fn in sorted(files):
                        ext = os.path.splitext(fn)[1].lower()
                        if sub == "." and ext not in (".xlsx", ".txt", ".csv"):
                            continue
                        it = QtWidgets.QTreeWidgetItem([os.path.relpath(os.path.join(root, fn), base)])
                        it.setData(0, QtCore.Qt.UserRole, os.path.join(root, fn))
                        top.addChild(it)
                        found = True
                        if ext == ".png" and self.first_graph is None:
                            self.first_graph = it
                if found:
                    self.tree.addTopLevelItem(top)
                    top.setExpanded(title.startswith("Graphs: plate x read"))

        def preview(self, item, _prev=None):
            path = item.data(0, QtCore.Qt.UserRole) if item else None
            if not path or not path.lower().endswith(".png"):
                return
            self._pix = QtGui.QPixmap(path)
            self.show_pix()

        def resizeEvent(self, e):
            super().resizeEvent(e)
            self.show_pix()

        def show_pix(self):
            if self._pix and not self._pix.isNull():
                w = max(self.img.parentWidget().width() - 20, 300)
                self.img.setPixmap(self._pix.scaledToWidth(w, QtCore.Qt.SmoothTransformation))

        def open_item(self, item, _col=0):
            path = item.data(0, QtCore.Qt.UserRole)
            if path:
                QtGui.QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(path))

        def open_folder(self):
            if self.out:
                QtGui.QDesktopServices.openUrl(QtCore.QUrl.fromLocalFile(self.out))

    app = QtWidgets.QApplication.instance() or QtWidgets.QApplication(sys.argv)
    w = Window()
    w.show()
    app.exec()


if __name__ == "__main__":
    main()
