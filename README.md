# Gen5 Data Extractor (website)

A static website: no server code, no uploads. The visitor's browser runs the same Python
program (`gen5_extract.py`) through Pyodide (Python compiled to WebAssembly).

## Using the Graphs tab
After extracting, the **Graphs** tab opens with a graph already drawn. Change anything and the graph redraws immediately:
* **Read** - which measurement to look at.
* **Graph type** - dashboard, overlay, mean +/- SD, normalized, peak bars, area under curve, heatmap (spectrum reads);
  dashboard, bar chart, plate map (endpoint reads).
* **Plates** - one plate, several, or all (several plates are compared side by side).
* **Wells** - *All wells*, *Single well* (click one), or *Multiple wells* (click, drag, click a row letter / column number,
  or type `A1, B2, A1:C3, A, 2`).
The graph (PNG), the selected data (CSV) and a per-well stats table can be downloaded from the same tab.

## Names and notes (Names & notes tab)
Name your plates and individual wells and add notes (type them or load a .txt file), then click **Apply**.
The graphs update straight away. The names and notes are saved everywhere: graph titles, legends and axis labels,
a small notes line at the bottom of every graph, Excel/CSV column headers and an *About* sheet in every workbook,
`Notes.txt`, and the README. **Download everything** rebuilds the files with your latest names first.
Command line: `--plate-names "Control,Treated"`, `--well-names "Plate 1 | A1=Blank; Plate 1 | A2:A4=Sample X; Control | B1=Std"` (names belong to one plate, so each entry names its plate), `--notes notes.txt`.

## Run it on your computer
    cd website
    python3 -m http.server 8000
    # open http://localhost:8000

(Opening index.html by double-click does not work: browsers block web workers on file:// pages.)

## Put it online
Upload the whole `website/` folder to any static host (GitHub Pages, Netlify, Cloudflare Pages,
a university web server...). Nothing else is needed. Files stay in each visitor's browser.
If the data is confidential, host it somewhere only your group can reach.

## Notes
* First visit downloads ~40 MB (Python + numpy/pandas/matplotlib from the jsDelivr CDN); browsers cache it afterwards.
* `gen5_extract.py` here is a copy of `../gen5_extract.py`. After changing the original, copy it again:
      cp ../gen5_extract.py gen5_extract.py
* `wheels/` holds the three small pure-Python packages Pyodide doesn't ship (openpyxl, olefile, et_xmlfile).
* To work fully offline, self-host Pyodide: download a release from https://github.com/pyodide/pyodide/releases
  and change `PYODIDE` at the top of `worker.js` to your copy.

## Install as an app (works offline)
Open the website in Chrome or Edge and click **Install app** (top right). It then opens in its own
window like a normal program and keeps working without internet. In Safari: File, then Add to Dock.

## Install with pip (command line)
Needs Python 3.10 or newer and git.

    pip install "git+https://github.com/ashokpaudelapril/gen5-extractor-site"
    gen5-extract Experiment1.xpt

With the desktop window:

    pip install "gen5-extractor[gui] @ git+https://github.com/ashokpaudelapril/gen5-extractor-site"
    gen5-extract-gui

Without installing anything permanently:

    pipx run --spec git+https://github.com/ashokpaudelapril/gen5-extractor-site gen5-extract Experiment1.xpt
    uvx --from git+https://github.com/ashokpaudelapril/gen5-extractor-site gen5-extract Experiment1.xpt
