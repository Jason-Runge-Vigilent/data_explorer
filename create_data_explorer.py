"""
create_data_explorer.py
────────────────────────
Generates a self-contained data_explorer.html in the same directory.

The HTML has three independent tabs.  Each tab lets you:
  • Upload one or more CSV or Excel files (client-side, no server needed)
  • Pick an X-axis column
  • Tick Y-axis columns and click "Add Figure" to create Plotly charts
  • Rename tabs, remove files, remove / clear figures

Run from any Python 3 environment – no third-party packages required.
The generated HTML only needs a modern browser and internet access
(Plotly and SheetJS are loaded from CDN).
"""

import os

BASE     = os.path.dirname(os.path.abspath(__file__))
OUT_PATH = os.path.join(BASE, "data_explorer.html")

# ── Number of tabs ────────────────────────────────────────────────────────
TAB_LABELS = ["Tab 1", "Tab 2", "Tab 3"]   # edit or extend as needed


def build_html(tab_labels: list[str]) -> str:
    tabs_json = "[" + ",".join(
        f'{{"id":"tab{i+1}","label":"{lbl}"}}'
        for i, lbl in enumerate(tab_labels)
    ) + "]"

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Data Explorer</title>
<script src="https://cdn.plot.ly/plotly-2.27.0.min.js"></script>
<script src="https://cdn.sheetjs.com/xlsx-0.20.1/package/dist/xlsx.full.min.js"></script>
<style>
  *, *::before, *::after {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: 'Segoe UI', Arial, sans-serif; background: #f0f2f5; color: #222;
         height: 100vh; display: flex; flex-direction: column; overflow: hidden; min-height: 0; }}

  /* ── Header ── */
  header {{ background: #1a3a5c; color: #fff; padding: 12px 20px; flex-shrink: 0;
            display: flex; align-items: center; gap: 12px; }}
  header h1 {{ font-size: 1.1rem; font-weight: 600; }}

  /* ── Tab bar ── */
  .tab-bar {{ display: flex; background: #fff; border-bottom: 2px solid #dde1e7;
              padding: 0 16px; flex-shrink: 0; align-items: center; }}
  .tab-btn {{ padding: 10px 22px; cursor: pointer; font-size: 0.88rem; font-weight: 500;
              border: none; background: none; color: #555;
              border-bottom: 3px solid transparent; margin-bottom: -2px; transition: color .15s; }}
  .tab-btn:hover  {{ color: #1a3a5c; }}
  .tab-btn.active {{ color: #1a3a5c; border-bottom-color: #1a7fd4; }}
  .tab-bar-spacer {{ flex: 1; }}
  .tab-rename-btn {{ font-size: 0.75rem; color: #888; background: none; border: none;
                    cursor: pointer; padding: 4px 8px; border-radius: 4px; }}
  .tab-rename-btn:hover {{ background: #eef2f7; color: #333; }}

  /* ── Tab content ── */
  .tab-content {{ display: none; flex: 1; min-height: 0; overflow: hidden; padding: 12px; gap: 12px; }}
  .tab-content.active {{ display: flex; }}

  /* ── Left panel ── */
  .left-panel {{ width: 270px; min-width: 230px; background: #fff; border-radius: 8px;
                box-shadow: 0 1px 4px rgba(0,0,0,.1); display: flex; flex-direction: column;
                overflow: hidden; flex-shrink: 0; min-height: 0; }}
  .lp-header {{ background: #1a3a5c; color: #fff; padding: 10px 13px;
                font-size: 0.85rem; font-weight: 600; flex-shrink: 0; }}

  .de-section {{ padding: 8px 12px; border-bottom: 1px solid #eee; flex-shrink: 0; }}
  .de-section-title {{ font-size: 0.72rem; font-weight: 700; color: #1a3a5c;
                       text-transform: uppercase; letter-spacing: .04em; margin-bottom: 5px; }}
  .de-upload-btn {{ display: block; width: 100%; padding: 7px 10px;
                   background: #eef2f7; border: 1px dashed #aab; border-radius: 6px;
                   cursor: pointer; font-size: 0.78rem; text-align: center; color: #445;
                   transition: background .15s; }}
  .de-upload-btn:hover {{ background: #dce4f0; }}
  .de-file-list {{ flex: 0 0 auto; max-height: 160px; overflow-y: auto;
                   border-bottom: 1px solid #eee; }}
  .de-file-item {{ display: flex; align-items: center; padding: 5px 12px;
                   font-size: 0.775rem; cursor: pointer; gap: 6px; color: #333;
                   border-bottom: 1px solid #f2f2f2; }}
  .de-file-item:hover {{ background: #f0f5fb; }}
  .de-file-item.selected {{ background: #e3eefa; color: #1a3a5c; font-weight: 600; }}
  .de-file-name {{ white-space: nowrap; overflow: hidden; text-overflow: ellipsis; flex: 1; }}
  .de-file-del {{ background: none; border: none; cursor: pointer; color: #ccc;
                  font-size: 0.85rem; padding: 0 3px; flex-shrink: 0; }}
  .de-file-del:hover {{ color: #c0392b; }}

  .de-x-select {{ width: 100%; padding: 5px 8px; font-size: 0.78rem;
                  border: 1px solid #ccd; border-radius: 5px; background: #fff;
                  color: #333; margin-top: 3px; }}
  .de-col-label {{ padding: 7px 12px 3px; font-size: 0.72rem; font-weight: 700;
                   color: #1a3a5c; text-transform: uppercase; letter-spacing: .04em;
                   flex-shrink: 0; border-bottom: 1px solid #eee; }}
  .de-col-list {{ flex: 1; overflow-y: auto; }}
  .de-col-placeholder {{ padding: 14px 12px; font-size: 0.78rem; color: #aaa; text-align: center; }}
  .sig-item {{ display: flex; align-items: center; padding: 5px 12px;
               font-size: 0.775rem; cursor: pointer; gap: 7px; color: #333; }}
  .sig-item:hover {{ background: #f0f5fb; }}
  .sig-item input {{ cursor: pointer; accent-color: #1a7fd4; }}

  .lp-actions {{ padding: 10px 12px; border-top: 1px solid #eee;
                  flex-shrink: 0; display: flex; flex-direction: column; gap: 5px; }}
  .btn {{ padding: 8px 10px; border-radius: 6px; border: none; cursor: pointer;
          font-size: 0.8rem; font-weight: 500; transition: background .15s; text-align: center; }}
  .btn-primary   {{ background: #1a7fd4; color: #fff; }}
  .btn-primary:hover   {{ background: #155faa; }}
  .btn-secondary {{ background: #eef2f7; color: #333; }}
  .btn-secondary:hover {{ background: #dce4f0; }}
  .btn-row {{ display: flex; gap: 5px; }}
  .btn-row .btn {{ flex: 1; }}

  /* ── Right panel ── */
  .right-panel {{ flex: 1; overflow-y: auto; display: flex; flex-direction: column;
                  gap: 10px; min-width: 0; min-height: 0; }}
  .empty-state {{ display: flex; align-items: center; justify-content: center;
                  flex: 1; background: #fff; border-radius: 8px; color: #aaa;
                  font-size: 0.9rem; box-shadow: 0 1px 4px rgba(0,0,0,.08);
                  min-height: 200px; text-align: center; padding: 20px; }}
  .empty-state strong {{ color: #555; }}

  .fig-card {{ background: #fff; border-radius: 8px; box-shadow: 0 1px 4px rgba(0,0,0,.1); overflow: hidden; }}
  .fig-card-hdr {{ display: flex; align-items: center; justify-content: space-between;
                   padding: 8px 13px; background: #f4f7fb; border-bottom: 1px solid #e5eaf2; }}
  .fig-card-title {{ font-size: 0.82rem; font-weight: 600; color: #1a3a5c; flex: 1;
                     white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }}
  .fig-remove {{ background: none; border: none; cursor: pointer; color: #aaa;
                 font-size: 1rem; padding: 2px 6px; border-radius: 4px; flex-shrink: 0; }}
  .fig-remove:hover {{ background: #fee; color: #c0392b; }}

  ::-webkit-scrollbar {{ width: 5px; }}
  ::-webkit-scrollbar-thumb {{ background: #ccc; border-radius: 3px; }}
</style>
</head>
<body>

<header>
  <h1>&#128202; Data Explorer</h1>
  <span style="font-size:0.8rem;opacity:0.7;">Upload CSV or Excel files &mdash; select columns &mdash; create plots</span>
</header>

<div class="tab-bar" id="tabBar">
  <span class="tab-bar-spacer"></span>
  <button class="tab-rename-btn" onclick="renameActiveTab()">&#9998;&nbsp; Rename tab</button>
</div>
<div id="tabContents" style="flex:1;min-height:0;display:flex;flex-direction:column;overflow:hidden;"></div>

<script>
const PALETTE = [
  '#1f77b4','#d62728','#2ca02c','#ff7f0e','#9467bd',
  '#8c564b','#e377c2','#17becf','#bcbd22','#7f7f7f',
  '#aec7e8','#ff9896','#98df8a','#ffbb78','#c5b0d5',
  '#c49c94','#f7b6d2','#dbdb8d','#9edae5','#c7c7c7',
];

const TABS = {tabs_json};
let activeTabId = TABS[0].id;
let figCounter  = 0;

const tabState = {{}};
TABS.forEach(t => {{ tabState[t.id] = {{ files: {{}}, activeFile: null }}; }});

/* ── Build UI ── */
function buildUI() {{
  const bar    = document.getElementById('tabBar');
  const contents = document.getElementById('tabContents');
  const spacer = bar.querySelector('.tab-bar-spacer');

  TABS.forEach((tab, i) => {{
    const btn = document.createElement('button');
    btn.className   = 'tab-btn' + (i === 0 ? ' active' : '');
    btn.textContent = tab.label;
    btn.dataset.tid = tab.id;
    btn.addEventListener('click', () => activateTab(tab.id));
    bar.insertBefore(btn, spacer);

    const div = document.createElement('div');
    div.id        = 'tc-' + tab.id;
    div.className = 'tab-content' + (i === 0 ? ' active' : '');
    div.innerHTML = buildPanelHTML(tab.id);
    contents.appendChild(div);

    document.getElementById('fi-' + tab.id)
      .addEventListener('change', e => handleUpload(e, tab.id));
  }});
}}

function buildPanelHTML(tid) {{
  return `
    <div class="left-panel">
      <div class="lp-header">Signal Selector</div>
      <div class="de-section">
        <div class="de-section-title">Files</div>
        <label class="de-upload-btn">
          &#128194;&nbsp; Upload CSV / Excel
          <input type="file" id="fi-${{tid}}" accept=".csv,.xlsx,.xls" multiple style="display:none">
        </label>
      </div>
      <div class="de-file-list" id="fl-${{tid}}"></div>
      <div class="de-section">
        <div class="de-section-title">X-Axis column</div>
        <select class="de-x-select" id="xc-${{tid}}">
          <option value="">— first column —</option>
        </select>
      </div>
      <div class="de-col-label">Y-Axis columns</div>
      <div class="de-col-list" id="cl-${{tid}}">
        <div class="de-col-placeholder">Upload a file and select it to see columns.</div>
      </div>
      <div class="lp-actions">
        <button class="btn btn-primary"   onclick="addFigure('${{tid}}')">&#43;&nbsp; Add Figure</button>
        <div class="btn-row">
          <button class="btn btn-secondary" onclick="selectAll('${{tid}}', true)">Select All</button>
          <button class="btn btn-secondary" onclick="selectAll('${{tid}}', false)">Deselect All</button>
        </div>
        <button class="btn btn-secondary" onclick="clearAll('${{tid}}')">&#128465;&nbsp; Clear All Figures</button>
      </div>
    </div>
    <div class="right-panel" id="rp-${{tid}}">
      <div class="empty-state" id="es-${{tid}}">
        Upload a CSV or Excel file, select columns, then click<br><strong>+ Add Figure</strong>
      </div>
    </div>`;
}}

/* ── Tab switching ── */
function activateTab(tid) {{
  activeTabId = tid;
  document.querySelectorAll('.tab-btn').forEach(b =>
    b.classList.toggle('active', b.dataset.tid === tid));
  document.querySelectorAll('.tab-content').forEach(d =>
    d.classList.toggle('active', d.id === 'tc-' + tid));
}}

function renameActiveTab() {{
  const tab = TABS.find(t => t.id === activeTabId);
  if (!tab) return;
  const name = prompt('Rename tab:', tab.label);
  if (name && name.trim()) {{
    tab.label = name.trim();
    document.querySelector(`.tab-btn[data-tid="${{activeTabId}}"]`).textContent = tab.label;
  }}
}}

/* ── File upload ── */
function handleUpload(e, tid) {{
  [...e.target.files].forEach(file => {{
    const reader = new FileReader();
    if (/\\.csv$/i.test(file.name)) {{
      reader.onload = ev => parseCSV(tid, file.name, ev.target.result);
      reader.readAsText(file);
    }} else {{
      reader.onload = ev => parseExcel(tid, file.name, ev.target.result);
      reader.readAsArrayBuffer(file);
    }}
  }});
  e.target.value = '';
}}

function parseCSV(tid, name, text) {{
  const lines = text.split(/\\r?\\n/).filter(l => l.trim());
  if (!lines.length) return;
  const sep     = lines[0].includes('\\t') ? '\\t' : ',';
  const columns = lines[0].split(sep).map(c => c.replace(/^"|"$/g, '').trim());
  const rows    = [];
  for (let i = 1; i < lines.length; i++) {{
    const vals = lines[i].split(sep).map(v => v.replace(/^"|"$/g, '').trim());
    const row  = {{}};
    columns.forEach((c, j) => row[c] = vals[j] ?? '');
    rows.push(row);
  }}
  storeFile(tid, name, columns, rows);
}}

function parseExcel(tid, name, buffer) {{
  if (typeof XLSX === 'undefined') {{
    alert('SheetJS failed to load. Check your internet connection or convert the file to CSV.');
    return;
  }}
  const wb      = XLSX.read(buffer, {{ type: 'array' }});
  const ws      = wb.Sheets[wb.SheetNames[0]];
  const arr     = XLSX.utils.sheet_to_json(ws, {{ header: 1, defval: '' }});
  if (!arr.length) return;
  const columns = arr[0].map(String);
  const rows    = arr.slice(1).map(r => {{
    const row = {{}};
    columns.forEach((c, j) => row[c] = r[j] ?? '');
    return row;
  }});
  storeFile(tid, name, columns, rows);
}}

function storeFile(tid, name, columns, rows) {{
  tabState[tid].files[name] = {{ columns, rows }};
  refreshFileList(tid);
  activateFile(tid, name);
}}

/* ── File list ── */
function refreshFileList(tid) {{
  const list   = document.getElementById('fl-' + tid);
  const active = tabState[tid].activeFile;
  list.innerHTML = '';
  Object.keys(tabState[tid].files).forEach(name => {{
    const item = document.createElement('div');
    item.className = 'de-file-item' + (name === active ? ' selected' : '');
    item.innerHTML = `<span>&#128196;</span>
      <span class="de-file-name" title="${{name}}">${{name}}</span>
      <button class="de-file-del" title="Remove" onclick="removeFile(event,'${{tid}}','${{name}}')">&#10005;</button>`;
    item.addEventListener('click', () => activateFile(tid, name));
    list.appendChild(item);
  }});
}}

function removeFile(e, tid, name) {{
  e.stopPropagation();
  const st = tabState[tid];
  delete st.files[name];
  if (st.activeFile === name) {{
    const remaining = Object.keys(st.files);
    st.activeFile = remaining.length ? remaining[0] : null;
    if (st.activeFile) activateFile(tid, st.activeFile);
    else {{
      document.getElementById('xc-' + tid).innerHTML = '<option value="">— first column —</option>';
      document.getElementById('cl-' + tid).innerHTML =
        '<div class="de-col-placeholder">Upload a file and select it to see columns.</div>';
    }}
  }}
  refreshFileList(tid);
}}

/* ── Column list ── */
function activateFile(tid, name) {{
  tabState[tid].activeFile = name;
  refreshFileList(tid);
  const info = tabState[tid].files[name];
  if (!info) return;

  const xSel = document.getElementById('xc-' + tid);
  xSel.innerHTML = '<option value="">— first column —</option>';
  info.columns.forEach(c => {{
    const opt = document.createElement('option');
    opt.value = c; opt.textContent = c;
    xSel.appendChild(opt);
  }});

  const colList = document.getElementById('cl-' + tid);
  colList.innerHTML = '';
  info.columns.forEach(c => {{
    const lbl = document.createElement('label');
    lbl.className = 'sig-item';
    const chk = document.createElement('input');
    chk.type = 'checkbox'; chk.value = c;
    lbl.appendChild(chk);
    lbl.appendChild(document.createTextNode(c));
    colList.appendChild(lbl);
  }});
}}

function selectAll(tid, checked) {{
  const cl = document.getElementById('cl-' + tid);
  if (cl) cl.querySelectorAll('input[type=checkbox]').forEach(c => c.checked = checked);
}}

/* ── Figures ── */
function addFigure(tid) {{
  const st = tabState[tid];
  if (!st.activeFile) {{ alert('Upload and select a file first.'); return; }}
  const cl      = document.getElementById('cl-' + tid);
  const checked = cl ? [...cl.querySelectorAll('input[type=checkbox]:checked')] : [];
  if (!checked.length) {{ alert('Please tick at least one column before adding a figure.'); return; }}

  const xCol  = document.getElementById('xc-' + tid).value;
  const yCols = checked.map(c => c.value);
  const info  = st.files[st.activeFile];
  const figId = 'fig-' + (++figCounter);
  const title = yCols.join('  |  ');

  const es = document.getElementById('es-' + tid);
  if (es) es.remove();

  const rp   = document.getElementById('rp-' + tid);
  const card = document.createElement('div');
  card.className = 'fig-card';
  card.id        = figId;
  card.innerHTML = `
    <div class="fig-card-hdr">
      <span class="fig-card-title" title="${{st.activeFile}} — ${{title}}">${{st.activeFile}} &mdash; ${{title}}</span>
      <button class="fig-remove" onclick="removeFigure('${{figId}}','${{tid}}')" title="Remove">&#10005;</button>
    </div>
    <div id="plot-${{figId}}"></div>`;
  rp.appendChild(card);
  renderPlot(figId, info, xCol, yCols);
}}

function renderPlot(figId, info, xCol, yCols) {{
  const xKey    = xCol || info.columns[0];
  const xValues = info.rows.map(r => r[xKey]);

  const traces = yCols.map((col, idx) => ({{
    x:    xValues,
    y:    info.rows.map(r => {{ const v = parseFloat(r[col]); return isNaN(v) ? r[col] : v; }}),
    mode: 'lines+markers',
    name: col,
    line:   {{ color: PALETTE[idx % PALETTE.length], width: 2 }},
    marker: {{ size: 4 }},
    hovertemplate: '<b>' + col + '</b><br>%{{x}}<br>%{{y:.4g}}<extra></extra>',
  }}));

  Plotly.newPlot('plot-' + figId, traces, {{
    height:        420,
    margin:        {{ l: 60, r: 15, t: 20, b: 60 }},
    xaxis:         {{ title: xKey, showgrid: true, gridcolor: '#ebebeb' }},
    yaxis:         {{ showgrid: true, gridcolor: '#ebebeb', autorange: true }},
    legend:        {{ orientation: 'h', y: -0.18, font: {{ size: 11 }} }},
    plot_bgcolor:  '#fff',
    paper_bgcolor: '#fff',
    hovermode:     'x unified',
  }}, {{
    responsive: true,
    displayModeBar: true,
    modeBarButtonsToRemove: ['select2d', 'lasso2d'],
    toImageButtonOptions: {{ filename: figId, scale: 2 }},
  }});
}}

function removeFigure(figId, tid) {{
  const el = document.getElementById(figId);
  if (el) el.remove();
  const rp = document.getElementById('rp-' + tid);
  if (!rp.querySelector('.fig-card')) {{
    rp.innerHTML = `<div class="empty-state" id="es-${{tid}}">
      Upload a CSV or Excel file, select columns, then click<br><strong>+ Add Figure</strong></div>`;
  }}
}}

function clearAll(tid) {{
  document.getElementById('rp-' + tid).innerHTML =
    `<div class="empty-state" id="es-${{tid}}">
      Upload a CSV or Excel file, select columns, then click<br><strong>+ Add Figure</strong></div>`;
}}

buildUI();
</script>
</body>
</html>"""


if __name__ == "__main__":
    print("Building HTML...")
    html = build_html(TAB_LABELS)
    with open(OUT_PATH, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Done!  Saved to:\n  {OUT_PATH}")
