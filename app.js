/**
 * ExcelAutomate - In-Browser Python WebAssembly Data & Excel Engine
 * Handles Pyodide initialization, client-side dataset parsing, interactive charts,
 * in-browser Python execution, and XLSX report compilation.
 */

// Global State
const state = {
  pyodide: null,
  pyodideReady: false,
  rawDataset: null,
  parsedData: [],
  columns: [],
  fileName: "sample_sales_data.csv",
  chartInstance: null,
  deferredPrompt: null
};

// Embedded Sample Dataset for 1-Click Instant Demo
const SAMPLE_CSV = `Date,Region,Sales_Rep,Category,Product,Units_Sold,Unit_Price,Total_Revenue,Cost,Profit,Status
2026-01-05,North America,Shane Weickum,Software,Enterprise Suite,14,1250.00,17500.00,5250.00,12250.00,Closed
2026-01-08,Europe,Emma Watson,Hardware,High-Density Server,6,3400.00,20400.00,12240.00,8160.00,Closed
2026-01-12,North America,Michael Scott,Consulting,Workflow Optimization,25,300.00,7500.00,2250.00,5250.00,Closed
2026-01-15,Asia Pacific,David Chen,Software,Cloud Storage Plan,80,95.00,7600.00,1520.00,6080.00,Closed
2026-01-19,Europe,Sophie Taylor,Software,Enterprise Suite,10,1250.00,12500.00,3750.00,8750.00,Closed
2026-01-22,Latin America,Carlos Gomez,Hardware,Workstation Pro,12,1850.00,22200.00,13320.00,8880.00,Closed
2026-01-27,North America,Shane Weickum,Consulting,Custom Integration,40,350.00,14000.00,4200.00,9800.00,Closed
2026-02-02,Asia Pacific,David Chen,Hardware,High-Density Server,8,3400.00,27200.00,16320.00,10880.00,Closed
2026-02-06,North America,Jim Halpert,Software,Analytics Dashboard,35,450.00,15750.00,3150.00,12600.00,Closed
2026-02-10,Europe,Emma Watson,Software,Enterprise Suite,16,1250.00,20000.00,6000.00,14000.00,Closed
2026-02-14,Latin America,Carlos Gomez,Consulting,Security Audit,18,650.00,11700.00,3510.00,8190.00,Closed
2026-02-18,North America,Shane Weickum,Hardware,Workstation Pro,15,1850.00,27750.00,16650.00,11100.00,Closed
2026-02-23,Asia Pacific,Mei Lin,Software,Analytics Dashboard,45,450.00,20250.00,4050.00,16200.00,Closed
2026-02-27,Europe,Sophie Taylor,Consulting,Workflow Optimization,30,300.00,9000.00,2700.00,6300.00,Closed
2026-03-03,North America,Michael Scott,Hardware,High-Density Server,5,3400.00,17000.00,10200.00,6800.00,Closed
2026-03-07,Latin America,Carlos Gomez,Software,Cloud Storage Plan,120,95.00,11400.00,2280.00,9120.00,Closed
2026-03-12,North America,Shane Weickum,Software,Enterprise Suite,22,1250.00,27500.00,8250.00,19250.00,Closed
2026-03-18,Europe,Emma Watson,Hardware,Workstation Pro,11,1850.00,20350.00,12210.00,8140.00,Closed
2026-03-24,Asia Pacific,David Chen,Consulting,Custom Integration,35,350.00,12250.00,3675.00,8575.00,Closed
2026-03-29,North America,Jim Halpert,Software,Enterprise Suite,18,1250.00,22500.00,6750.00,15750.00,Closed`;

// Python Transformation Script to Execute in Browser WebAssembly
const PYTHON_RUNNER_SCRIPT = `
import io
import json
import base64
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, PieChart, DoughnutChart, Reference

def build_excel_report(json_data_str, config_json_str):
    # 1. Parse inputs
    data_list = json.loads(json_data_str)
    config = json.loads(config_json_str)
    
    df = pd.DataFrame(data_list)
    df.columns = [str(c).strip() for c in df.columns]
    
    report_title = config.get("report_title", "Executive Business Performance Report")
    author = config.get("author", "Shane Weickum")
    source_name = config.get("source_name", "Data_Upload")
    theme = config.get("theme", "navy")
    enable_charts = config.get("enable_charts", True)
    chart_type = config.get("chart_type", "bar")
    chart_cat = config.get("chart_cat")
    chart_metric = config.get("chart_metric")
    chart_agg = config.get("chart_agg", "sum")
    
    # 2. Theme palettes
    palettes = {
        "navy": {"primary": "1E293B", "accent": "2563EB", "subtle": "F8FAFC", "card": "F1F5F9"},
        "emerald": {"primary": "064E3B", "accent": "10B981", "subtle": "F0FDF4", "card": "ECFDF5"},
        "indigo": {"primary": "312E81", "accent": "6366F1", "subtle": "EEF2FF", "card": "E0E7FF"},
        "slate": {"primary": "18181B", "accent": "64748B", "subtle": "F8FAFC", "card": "F1F5F9"}
    }
    pal = palettes.get(theme, palettes["navy"])
    
    # Fonts & Fills
    font_family = "Segoe UI"
    font_title = Font(name=font_family, size=15, bold=True, color="FFFFFF")
    font_sub = Font(name=font_family, size=9, italic=True, color="E2E8F0")
    font_section = Font(name=font_family, size=11, bold=True, color=pal["primary"])
    font_card_lbl = Font(name=font_family, size=8, bold=True, color="64748B")
    font_card_val = Font(name=font_family, size=15, bold=True, color=pal["primary"])
    font_card_val_green = Font(name=font_family, size=15, bold=True, color="059669")
    font_tbl_hdr = Font(name=font_family, size=10, bold=True, color="FFFFFF")
    font_data = Font(name=font_family, size=10, color="1E293B")
    font_data_bold = Font(name=font_family, size=10, bold=True, color="1E293B")
    
    fill_hdr = PatternFill(start_color=pal["primary"], end_color=pal["primary"], fill_type="solid")
    fill_card = PatternFill(start_color=pal["card"], end_color=pal["card"], fill_type="solid")
    fill_zebra = PatternFill(start_color=pal["subtle"], end_color=pal["subtle"], fill_type="solid")
    fill_total = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
    
    thin_side = Side(border_style="thin", color="CBD5E1")
    thin_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
    double_bottom = Border(top=thin_side, bottom=Side(border_style="double", color=pal["primary"]))
    
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    
    # 3. Detect column types
    numeric_cols = [c for c in df.columns if pd.api.types.is_numeric_dtype(df[c])]
    currency_keys = ["revenue", "sales", "profit", "price", "cost", "amount", "fee", "total"]
    currency_cols = [c for c in numeric_cols if any(k in c.lower() for k in currency_keys)]
    qty_keys = ["units", "qty", "quantity", "count", "volume"]
    qty_cols = [c for c in numeric_cols if any(k in c.lower() for k in qty_keys)]
    date_cols = [c for c in df.columns if "date" in c.lower() or "time" in c.lower()]
    cat_cols = [c for c in df.columns if c not in numeric_cols and c not in date_cols]
    
    # Select primary metric & category
    pri_metric = chart_metric if (chart_metric and chart_metric in df.columns) else (currency_cols[0] if currency_cols else (numeric_cols[0] if numeric_cols else None))
    pri_cat = chart_cat if (chart_cat and chart_cat in df.columns) else (cat_cols[0] if cat_cols else None)
    
    wb = openpyxl.Workbook()
    
    # =========================================================================
    # SHEET 1: EXECUTIVE SUMMARY
    # =========================================================================
    ws_sum = wb.active
    ws_sum.title = config.get("summary_sheet_name", "Executive Summary")
    ws_sum.views.sheetView[0].showGridLines = True
    
    # Title Banner
    ws_sum.merge_cells("B2:I2")
    ws_sum.merge_cells("B3:I3")
    ws_sum["B2"].value = f"  {report_title}"
    ws_sum["B2"].font = font_title
    ws_sum["B2"].fill = fill_hdr
    ws_sum["B2"].alignment = align_left
    
    ws_sum["B3"].value = f"  Source: {source_name}   |   Prepared by: {author}"
    ws_sum["B3"].font = font_sub
    ws_sum["B3"].fill = fill_hdr
    ws_sum["B3"].alignment = align_left
    
    for r in range(2, 4):
        for c in range(2, 10):
            ws_sum.cell(r, c).fill = fill_hdr
            
    # KPI Metric Cards
    kpis = [{"lbl": "TOTAL RECORDS", "val": f"{len(df):,}", "grn": False}]
    if pri_metric:
        tot = df[pri_metric].sum()
        is_c = pri_metric in currency_cols
        kpis.append({"lbl": f"TOTAL {pri_metric.replace('_',' ').upper()}", "val": f"\${tot:,.2f}" if is_c else f"{tot:,.0f}", "grn": True})
        avg = df[pri_metric].mean()
        kpis.append({"lbl": f"AVG {pri_metric.replace('_',' ').upper()} / ROW", "val": f"\${avg:,.2f}" if is_c else f"{avg:,.1f}", "grn": False})
        
    sec_metric = next((c for c in currency_cols if c != pri_metric), None)
    if sec_metric:
        sec_tot = df[sec_metric].sum()
        kpis.append({"lbl": f"TOTAL {sec_metric.replace('_',' ').upper()}", "val": f"\${sec_tot:,.2f}", "grn": False})
        
    card_cols = [("B","C"), ("D","E"), ("F","G"), ("H","I")]
    for idx, k in enumerate(kpis[:4]):
        cs, ce = card_cols[idx]
        ws_sum.merge_cells(f"{cs}5:{ce}5")
        ws_sum.merge_cells(f"{cs}6:{ce}6")
        
        c_l = ws_sum[f"{cs}5"]
        c_l.value = k["lbl"]
        c_l.font = font_card_lbl
        c_l.alignment = align_center
        
        c_v = ws_sum[f"{cs}6"]
        c_v.value = k["val"]
        c_v.font = font_card_val_green if k["grn"] else font_card_val
        c_v.alignment = align_center
        
        cs_i = openpyxl.utils.column_index_from_string(cs)
        ce_i = openpyxl.utils.column_index_from_string(ce)
        for r in range(5, 7):
            for c in range(cs_i, ce_i + 1):
                cell = ws_sum.cell(r, c)
                cell.fill = fill_card
                cell.border = thin_border

    # Breakdown Table
    curr_row = 9
    t_start = 0
    t_end = 0
    if pri_cat and pri_metric:
        ws_sum.cell(curr_row, 2, f"Performance Breakdown by {pri_cat.replace('_',' ').title()}").font = font_section
        curr_row += 1
        
        grp_df = df.groupby(pri_cat, as_index=False).agg({pri_metric: chart_agg if chart_agg in ["sum","mean","count","max","min"] else "sum"})
        grp_df = grp_df.sort_values(by=pri_metric, ascending=False).reset_index(drop=True)
        
        headers = [pri_cat, pri_metric]
        for col_i, h in enumerate(headers, start=2):
            cell = ws_sum.cell(curr_row, col_i, h.replace('_',' ').title())
            cell.font = font_tbl_hdr
            cell.fill = fill_hdr
            cell.alignment = align_left if col_i == 2 else align_right
            cell.border = thin_border
        curr_row += 1
        t_start = curr_row
        
        for r_i, r_data in grp_df.iterrows():
            r_fill = fill_zebra if r_i % 2 == 1 else PatternFill(fill_type=None)
            for c_i, col_name in enumerate(headers, start=2):
                val = r_data[col_name]
                cell = ws_sum.cell(curr_row, c_i, val)
                cell.font = font_data
                cell.border = thin_border
                if r_fill.fill_type:
                    cell.fill = r_fill
                if col_name in currency_cols:
                    cell.number_format = '"$"#,##0.00'
                    cell.alignment = align_right
                elif col_name in numeric_cols:
                    cell.number_format = '#,##0'
                    cell.alignment = align_right
            curr_row += 1
        t_end = curr_row - 1
        
        # Total row
        ws_sum.cell(curr_row, 2, "Total").font = font_data_bold
        ws_sum.cell(curr_row, 2).fill = fill_total
        ws_sum.cell(curr_row, 2).border = double_bottom
        
        col_let = get_column_letter(3)
        cell_t = ws_sum.cell(curr_row, 3)
        cell_t.value = f"=SUM({col_let}{t_start}:{col_let}{t_end})"
        cell_t.font = font_data_bold
        cell_t.fill = fill_total
        cell_t.border = double_bottom
        if pri_metric in currency_cols:
            cell_t.number_format = '"$"#,##0.00'
            cell_t.alignment = align_right
            
        # Embedded Chart
        if enable_charts and t_end >= t_start:
            if chart_type == "line":
                ch = LineChart()
            elif chart_type == "pie":
                ch = PieChart()
            elif chart_type == "doughnut":
                ch = DoughnutChart()
            else:
                ch = BarChart()
                ch.type = "col"
                
            ch.title = f"{pri_metric.replace('_',' ').title()} by {pri_cat.replace('_',' ').title()}"
            ch.width = 15
            ch.height = 9
            
            d_ref = Reference(ws_sum, min_col=3, min_row=t_start - 1, max_row=t_end)
            c_ref = Reference(ws_sum, min_col=2, min_row=t_start, max_row=t_end)
            ch.add_data(d_ref, titles_from_data=True)
            ch.set_categories(c_ref)
            ws_sum.add_chart(ch, "F9")

    # Column widths
    for col in ws_sum.columns:
        col_letter = get_column_letter(col[0].column)
        max_l = 0
        for cell in col:
            if cell.row not in [2, 3] and cell.value:
                max_l = max(max_l, len(str(cell.value)))
        ws_sum.column_dimensions[col_letter].width = max(max_l + 4, 14)
    ws_sum.column_dimensions["A"].width = 3

    # =========================================================================
    # SHEET 2: DETAILED DATA
    # =========================================================================
    ws_dat = wb.create_sheet(title=config.get("data_sheet_name", "Detailed Data"))
    ws_dat.views.sheetView[0].showGridLines = True
    
    # Header
    for c_i, c_n in enumerate(df.columns, start=1):
        c = ws_dat.cell(1, c_i, c_n.replace('_',' ').title())
        c.font = font_tbl_hdr
        c.fill = fill_hdr
        c.border = thin_border
        
    for r_i, r_vals in enumerate(df.values, start=2):
        r_fill = fill_zebra if r_i % 2 == 1 else PatternFill(fill_type=None)
        for c_i, val in enumerate(r_vals, start=1):
            col_n = df.columns[c_i - 1]
            cell = ws_dat.cell(r_i, c_i, val)
            cell.font = font_data
            cell.border = thin_border
            if r_fill.fill_type:
                cell.fill = r_fill
            if col_n in currency_cols:
                cell.number_format = '"$"#,##0.00'
                cell.alignment = align_right
            elif col_n in qty_cols:
                cell.number_format = '#,##0'
                cell.alignment = align_right
            elif col_n in numeric_cols:
                cell.number_format = '#,##0.00'
                cell.alignment = align_right

    ws_dat.auto_filter.ref = ws_dat.dimensions
    ws_dat.freeze_panes = "A2"
    
    for col in ws_dat.columns:
        col_letter = get_column_letter(col[0].column)
        max_l = max([len(str(c.value or '')) for c in col])
        ws_dat.column_dimensions[col_letter].width = min(max(max_l + 4, 12), 45)

    # Save to in-memory bytes buffer
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return base64.b64encode(buf.getvalue()).decode('utf-8')
`;

// Initialize Pyodide WebAssembly Engine
async function initPyodideEngine() {
  const statusEl = document.getElementById("pyodide-status");
  logToConsole("Initializing Pyodide WebAssembly Python environment...");

  try {
    state.pyodide = await loadPyodide({
      stdout: (msg) => logToConsole(`[Python stdout] ${msg}`),
      stderr: (msg) => logToConsole(`[Python stderr] ${msg}`)
    });

    logToConsole("Loading essential Python wheels: pandas, micropip...");
    statusEl.innerHTML = '<span class="status-dot"></span><span class="status-text">Loading pandas...</span>';

    await state.pyodide.loadPackage(["pandas", "micropip"]);

    // openpyxl is pure-Python and isn't in Pyodide's built package repository,
    // so it has to come from PyPI via micropip instead of loadPackage.
    logToConsole("Installing openpyxl from PyPI via micropip...");
    statusEl.innerHTML = '<span class="status-dot"></span><span class="status-text">Installing openpyxl...</span>';
    const micropip = state.pyodide.pyimport("micropip");
    await micropip.install("openpyxl");

    logToConsole("Compiling Python transformation engine...");
    await state.pyodide.runPythonAsync(PYTHON_RUNNER_SCRIPT);

    state.pyodideReady = true;
    statusEl.className = "status-pill status-ready";
    statusEl.innerHTML = '<span class="status-dot"></span><span class="status-text">Python Engine Ready</span>';
    logToConsole("✓ Pyodide WebAssembly Python Engine is fully ready!");
    showToast("Python WebAssembly engine loaded successfully", "success");
    
    // Auto-load sample dataset on first visit
    loadSampleData();
  } catch (err) {
    console.error("Pyodide init error:", err);
    statusEl.className = "status-pill status-loading";
    statusEl.innerHTML = '<span class="status-dot"></span><span class="status-text">Python Load Failed</span>';
    logToConsole(`[ERROR] Failed to initialize Pyodide: ${err.message}`);
    showToast("Could not load Python WebAssembly engine", "error");
  }
}

// Log to Virtual Python Console — the persistent in-app debugger.
// Survives reloads via sessionStorage and flags unseen errors on the tab
// itself, so a failure is never only visible in the browser devtools.
const CONSOLE_LOG_KEY = "excel_automate_console_log";
const CONSOLE_LOG_MAX_LINES = 500;

function detectLogLevel(msg) {
  if (/^\[ERROR\]/.test(msg)) return "error";
  if (/^\[WARN\]/.test(msg)) return "warn";
  if (msg.startsWith("✓")) return "success";
  return "info";
}

function logToConsole(msg, level = null) {
  const consoleEl = document.getElementById("console-output");
  if (!consoleEl) return;
  const resolvedLevel = level || detectLogLevel(msg);
  const timestamp = new Date().toLocaleTimeString();
  const line = document.createElement("span");
  line.className = `console-line log-${resolvedLevel}`;
  line.textContent = `[${timestamp}] ${msg}\n`;
  consoleEl.appendChild(line);
  consoleEl.scrollTop = consoleEl.scrollHeight;

  while (consoleEl.childNodes.length > CONSOLE_LOG_MAX_LINES) {
    consoleEl.removeChild(consoleEl.firstChild);
  }

  try {
    sessionStorage.setItem(CONSOLE_LOG_KEY, consoleEl.innerHTML);
  } catch (e) { /* sessionStorage unavailable or full — logging still works in-page */ }

  if (resolvedLevel === "error") {
    flagConsoleError();
  }
}

function flagConsoleError() {
  const badge = document.getElementById("console-error-badge");
  const consolePanel = document.getElementById("console-tab");
  if (badge && !consolePanel?.classList.contains("active")) {
    badge.classList.remove("hidden");
  }
}

function restoreConsoleLog() {
  const consoleEl = document.getElementById("console-output");
  if (!consoleEl) return;
  const saved = sessionStorage.getItem(CONSOLE_LOG_KEY);
  consoleEl.innerHTML = saved || "";
  consoleEl.scrollTop = consoleEl.scrollHeight;
}

// Surface any otherwise-uncaught JS error/rejection into the same debugger
// panel, so future updates (from any tool) still get logged, not just
// errors this file explicitly calls logToConsole() for.
window.addEventListener("error", (event) => {
  logToConsole(`[ERROR] Uncaught exception: ${event.message} (${event.filename}:${event.lineno})`, "error");
});
window.addEventListener("unhandledrejection", (event) => {
  logToConsole(`[ERROR] Unhandled promise rejection: ${event.reason?.message || event.reason}`, "error");
});

// Toast Notifications
function showToast(msg, type = "info") {
  const container = document.getElementById("toast-container");
  const toast = document.createElement("div");
  toast.className = `toast ${type}`;
  toast.innerHTML = `<span>${msg}</span>`;
  container.appendChild(toast);
  setTimeout(() => {
    toast.style.opacity = "0";
    setTimeout(() => toast.remove(), 300);
  }, 4000);
}

// Tab Switching
function initTabNavigation() {
  const tabBtns = document.querySelectorAll(".tab-btn");
  const panels = document.querySelectorAll(".tab-panel");

  tabBtns.forEach(btn => {
    btn.addEventListener("click", () => {
      const targetId = btn.getAttribute("data-tab");
      tabBtns.forEach(b => b.classList.remove("active"));
      panels.forEach(p => p.classList.remove("active"));

      btn.classList.add("active");
      const targetPanel = document.getElementById(targetId);
      if (targetPanel) targetPanel.classList.add("active");

      // Re-render chart if switching to visuals tab
      if (targetId === "charts-tab") {
        updateChartPreview();
      }

      // Acknowledge any pending error badge when the console is opened
      if (targetId === "console-tab") {
        document.getElementById("console-error-badge")?.classList.add("hidden");
      }
    });
  });
}

// Parse Raw Text (CSV, TSV, JSON)
function parseRawDataset(rawText, fileName = "data.csv") {
  state.fileName = fileName;
  state.rawDataset = rawText;
  logToConsole(`Parsing data file: ${fileName}...`);

  try {
    let rows = [];
    if (fileName.endsWith(".json")) {
      const parsed = JSON.parse(rawText);
      rows = Array.isArray(parsed) ? parsed : [parsed];
    } else {
      const workbook = XLSX.read(rawText, { type: "string" });
      const firstSheet = workbook.SheetNames[0];
      rows = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheet]);
    }

    if (!rows || rows.length === 0) {
      showToast("Uploaded file contains no data rows", "error");
      return;
    }

    state.parsedData = rows;
    state.columns = Object.keys(rows[0]);

    document.getElementById("loaded-filename").textContent = fileName;
    document.getElementById("loaded-filename").classList.remove("hidden");
    document.getElementById("data-stats").classList.remove("hidden");

    renderPreviewTable(rows);
    updateDatasetStats(rows);
    populateChartSelects();
    calculateLiveKPIs(rows);
    updateChartPreview();

    logToConsole(`✓ Successfully loaded ${rows.length} records with ${state.columns.length} columns.`);
    showToast(`Loaded ${rows.length} records from ${fileName}`, "success");
  } catch (err) {
    console.error("Data parse error:", err);
    logToConsole(`[ERROR] Failed to parse data: ${err.message}`);
    showToast("Error reading file format", "error");
  }
}

// Render Preview Table
function renderPreviewTable(rows) {
  const container = document.getElementById("table-container");
  const countBadge = document.getElementById("preview-count");
  if (!rows || rows.length === 0) {
    container.innerHTML = '<div class="empty-state">No records to display.</div>';
    countBadge.textContent = "0 rows";
    return;
  }

  countBadge.textContent = `Showing 1-10 of ${rows.length} rows`;

  const previewRows = rows.slice(0, 10);
  const cols = state.columns;

  let html = '<table class="preview-table"><thead><tr>';
  cols.forEach(c => { html += `<th>${c}</th>`; });
  html += '</tr></thead><tbody>';

  previewRows.forEach(r => {
    html += '<tr>';
    cols.forEach(c => {
      const val = r[c] !== undefined ? r[c] : '';
      html += `<td>${val}</td>`;
    });
    html += '</tr>';
  });

  html += '</tbody></table>';
  container.innerHTML = html;
}

// Update Stats Banner
function updateDatasetStats(rows) {
  document.getElementById("stat-rows").textContent = rows.length.toLocaleString();
  document.getElementById("stat-cols").textContent = state.columns.length;

  const numericCols = state.columns.filter(c => rows.some(r => typeof r[c] === "number" || !isNaN(parseFloat(r[c]))));
  const catCols = state.columns.filter(c => !numericCols.includes(c));

  document.getElementById("stat-metrics").textContent = numericCols.length;
  document.getElementById("stat-cats").textContent = catCols.length;
}

// Populate Chart Dropdowns
function populateChartSelects() {
  const catSelect = document.getElementById("chart-cat-select");
  const metricSelect = document.getElementById("chart-metric-select");

  catSelect.innerHTML = "";
  metricSelect.innerHTML = "";

  state.columns.forEach(col => {
    const isNum = state.parsedData.some(r => typeof r[col] === "number" || !isNaN(parseFloat(r[col])));
    
    const opt = document.createElement("option");
    opt.value = col;
    opt.textContent = col;

    if (isNum) {
      metricSelect.appendChild(opt);
    } else {
      catSelect.appendChild(opt);
    }
  });

  // Default selections
  if (catSelect.options.length === 0) {
    state.columns.forEach(col => {
      const opt = document.createElement("option");
      opt.value = col; opt.textContent = col;
      catSelect.appendChild(opt);
    });
  }
  if (metricSelect.options.length === 0) {
    state.columns.forEach(col => {
      const opt = document.createElement("option");
      opt.value = col; opt.textContent = col;
      metricSelect.appendChild(opt);
    });
  }
}

// Calculate Live KPI Cards
function calculateLiveKPIs(rows) {
  document.getElementById("kpi-rec-val").textContent = rows.length.toLocaleString();

  const metricCol = document.getElementById("chart-metric-select")?.value || 
    state.columns.find(c => ["revenue", "sales", "total_revenue", "amount", "profit"].includes(c.toLowerCase())) ||
    state.columns.find(c => typeof rows[0][c] === "number");

  if (metricCol) {
    const sum = rows.reduce((acc, r) => acc + (parseFloat(r[metricCol]) || 0), 0);
    const avg = sum / (rows.length || 1);
    const isCurr = ["revenue", "sales", "price", "profit", "cost", "total"].some(k => metricCol.toLowerCase().includes(k));

    document.getElementById("kpi-rev-val").textContent = isCurr ? `$${sum.toLocaleString(undefined, {minimumFractionDigits:2, maximumFractionDigits:2})}` : sum.toLocaleString();
    document.getElementById("kpi-avg-val").textContent = isCurr ? `$${avg.toLocaleString(undefined, {minimumFractionDigits:2, maximumFractionDigits:2})}` : avg.toFixed(1);
  }

  const profCol = state.columns.find(c => c.toLowerCase().includes("profit"));
  if (profCol) {
    const profSum = rows.reduce((acc, r) => acc + (parseFloat(r[profCol]) || 0), 0);
    document.getElementById("kpi-prof-val").textContent = `$${profSum.toLocaleString(undefined, {minimumFractionDigits:2, maximumFractionDigits:2})}`;
  } else {
    document.getElementById("kpi-prof-val").textContent = "--";
  }
}

// Update Chart.js Live Interactive Preview
function updateChartPreview() {
  const canvas = document.getElementById("chart-preview-canvas");
  if (!canvas || !state.parsedData || state.parsedData.length === 0) return;

  const chartType = document.getElementById("chart-type-select").value || "bar";
  const catCol = document.getElementById("chart-cat-select").value || state.columns[0];
  const metricCol = document.getElementById("chart-metric-select").value || state.columns[1];
  const aggMethod = document.getElementById("chart-agg-select").value || "sum";

  // Aggregate Data
  const aggData = {};
  state.parsedData.forEach(r => {
    const key = r[catCol] || "Unknown";
    const val = parseFloat(r[metricCol]) || 0;
    if (!aggData[key]) aggData[key] = [];
    aggData[key].push(val);
  });

  const labels = Object.keys(aggData);
  const dataPoints = labels.map(k => {
    const arr = aggData[k];
    if (aggMethod === "avg") return arr.reduce((a, b) => a + b, 0) / arr.length;
    if (aggMethod === "count") return arr.length;
    if (aggMethod === "max") return Math.max(...arr);
    if (aggMethod === "min") return Math.min(...arr);
    return arr.reduce((a, b) => a + b, 0); // Default sum
  });

  document.getElementById("chart-preview-title").textContent = `${metricCol} by ${catCol} (${aggMethod.toUpperCase()})`;

  if (state.chartInstance) {
    state.chartInstance.destroy();
  }

  const ctx = canvas.getContext("2d");
  const isPie = ["pie", "doughnut"].includes(chartType);

  const colors = [
    "#3B82F6", "#10B981", "#F59E0B", "#8B5CF6", "#EC4899",
    "#06B6D4", "#6366F1", "#14B8A6", "#F97316", "#84CC16"
  ];

  state.chartInstance = new Chart(ctx, {
    type: chartType,
    data: {
      labels: labels,
      datasets: [{
        label: `${metricCol} (${aggMethod.toUpperCase()})`,
        data: dataPoints,
        backgroundColor: isPie ? colors.slice(0, labels.length) : "rgba(59, 130, 246, 0.8)",
        borderColor: isPie ? "#1E293B" : "#2563EB",
        borderWidth: 2,
        borderRadius: chartType === "bar" ? 6 : 0
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          display: isPie,
          labels: { color: "#CBD5E1" }
        }
      },
      scales: isPie ? {} : {
        x: { ticks: { color: "#94A3B8" }, grid: { color: "rgba(255,255,255,0.05)" } },
        y: { ticks: { color: "#94A3B8" }, grid: { color: "rgba(255,255,255,0.05)" } }
      }
    }
  });
}

// Load Built-in Sample Data
function loadSampleData() {
  parseRawDataset(SAMPLE_CSV, "sample_sales_data.csv");
}

// Execute Python Report Pipeline in Browser
async function executePythonPipeline() {
  if (!state.pyodideReady) {
    showToast("Python WebAssembly engine is still initializing...", "info");
    return;
  }

  if (!state.parsedData || state.parsedData.length === 0) {
    showToast("Please upload or load a dataset first", "error");
    return;
  }

  const runBtn = document.getElementById("run-pipeline-btn");
  const dockMsg = document.getElementById("dock-status-msg");

  runBtn.disabled = true;
  runBtn.innerHTML = '<span class="status-dot"></span> Processing in Python...';
  dockMsg.textContent = "Executing Python openpyxl engine...";
  logToConsole("Starting Python report generation pipeline...");

  try {
    // Gather UI configurations
    const themeRadio = document.querySelector('input[name="theme-color"]:checked');
    const selectedTheme = themeRadio ? themeRadio.value : "navy";

    const config = {
      report_title: document.getElementById("report-title-input").value || "Executive Business Performance Report",
      author: document.getElementById("report-author-input").value || "Shane Weickum",
      source_name: state.fileName,
      summary_sheet_name: document.getElementById("summary-sheet-name").value || "Executive Summary",
      data_sheet_name: document.getElementById("data-sheet-name").value || "Detailed Data",
      theme: selectedTheme,
      enable_charts: document.getElementById("enable-charts-toggle").checked,
      chart_type: document.getElementById("chart-type-select").value || "bar",
      chart_cat: document.getElementById("chart-cat-select").value,
      chart_metric: document.getElementById("chart-metric-select").value,
      chart_agg: document.getElementById("chart-agg-select").value || "sum"
    };

    const jsonStr = JSON.stringify(state.parsedData);
    const configStr = JSON.stringify(config);

    logToConsole(`Configuring workbook: Theme=${config.theme}, Chart=${config.chart_type}, Metric=${config.chart_metric}...`);

    // Call Python function directly
    const buildReportFunc = state.pyodide.globals.get("build_excel_report");
    const base64Xlsx = buildReportFunc(jsonStr, configStr);

    logToConsole("✓ Excel report workbook compiled successfully into binary buffer!");

    // Convert Base64 back to Blob and trigger download
    const byteChars = atob(base64Xlsx);
    const byteNumbers = new Array(byteChars.length);
    for (let i = 0; i < byteChars.length; i++) {
      byteNumbers[i] = byteChars.charCodeAt(i);
    }
    const byteArray = new Uint8Array(byteNumbers);
    const blob = new Blob([byteArray], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });

    const timestamp = new Date().toISOString().replace(/[-:T]/g, "").slice(0, 14);
    const outName = `${state.fileName.replace(/\.[^/.]+$/, "")}_Report_${timestamp}.xlsx`;

    const downloadLink = document.createElement("a");
    downloadLink.href = URL.createObjectURL(blob);
    downloadLink.download = outName;
    document.body.appendChild(downloadLink);
    downloadLink.click();
    downloadLink.remove();

    dockMsg.textContent = `Report generated: ${outName}`;
    logToConsole(`✓ Download triggered: ${outName}`);
    showToast(`Downloaded ${outName}`, "success");

    // If email is configured, notify
    if (document.getElementById("send-email-toggle").checked) {
      const recipient = document.getElementById("email-recipient").value;
      logToConsole(`[Email Dispatcher] Ready to dispatch report attachment to: ${recipient}`);
      showToast(`Email report queued for ${recipient}`, "info");
    }
  } catch (err) {
    console.error("Python pipeline execution error:", err);
    logToConsole(`[ERROR] Pipeline failed: ${err.message}`);
    showToast(`Pipeline Error: ${err.message}`, "error");
    dockMsg.textContent = "Pipeline execution failed.";
  } finally {
    runBtn.disabled = false;
    runBtn.innerHTML = `
      <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"></polygon>
      </svg>
      Generate & Download Excel (.xlsx)
    `;
  }
}

// Local Storage for Settings Persistence
function initLocalStorage() {
  const fields = ["email-recipient", "email-subject", "smtp-host", "smtp-port", "smtp-user", "smtp-pass"];
  fields.forEach(id => {
    const el = document.getElementById(id);
    if (!el) return;
    const saved = localStorage.getItem(`excel_automate_${id}`);
    if (saved !== null) el.value = saved;
    el.addEventListener("input", () => {
      localStorage.setItem(`excel_automate_${id}`, el.value);
    });
  });
}

// File Drag & Drop Handling
function initFileHandlers() {
  const dropZone = document.getElementById("drop-zone");
  const fileInput = document.getElementById("file-input");
  const browseBtn = document.getElementById("browse-btn");

  browseBtn.addEventListener("click", () => fileInput.click());
  dropZone.addEventListener("click", (e) => {
    if (e.target !== browseBtn) fileInput.click();
  });

  dropZone.addEventListener("dragover", (e) => {
    e.preventDefault();
    dropZone.classList.add("drag-over");
  });

  dropZone.addEventListener("dragleave", () => {
    dropZone.classList.remove("drag-over");
  });

  dropZone.addEventListener("drop", (e) => {
    e.preventDefault();
    dropZone.classList.remove("drag-over");
    if (e.dataTransfer.files.length > 0) {
      handleUploadedFile(e.dataTransfer.files[0]);
    }
  });

  fileInput.addEventListener("change", (e) => {
    if (e.target.files.length > 0) {
      handleUploadedFile(e.target.files[0]);
    }
  });
}

function handleUploadedFile(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    parseRawDataset(e.target.result, file.name);
  };
  reader.readAsText(file);
}

// Theme Card Selection Handler
function initThemeSelectors() {
  const themeCards = document.querySelectorAll(".theme-card");
  themeCards.forEach(card => {
    card.addEventListener("click", () => {
      themeCards.forEach(c => c.classList.remove("active"));
      card.classList.add("active");
    });
  });
}

// PWA Service Worker & Install Button
function initPWA() {
  if ("serviceWorker" in navigator) {
    navigator.serviceWorker.register("./sw.js")
      .then(() => console.log("Service Worker Registered"))
      .catch(err => console.warn("SW registration error:", err));
  }

  const installBtn = document.getElementById("install-btn");
  window.addEventListener("beforeinstallprompt", (e) => {
    e.preventDefault();
    state.deferredPrompt = e;
    installBtn.classList.remove("hidden");
  });

  installBtn.addEventListener("click", async () => {
    if (state.deferredPrompt) {
      state.deferredPrompt.prompt();
      const { outcome } = await state.deferredPrompt.userChoice;
      if (outcome === "accepted") {
        installBtn.classList.add("hidden");
      }
      state.deferredPrompt = null;
    }
  });
}

// Initialize on DOM Loaded
document.addEventListener("DOMContentLoaded", () => {
  restoreConsoleLog();
  initTabNavigation();
  initFileHandlers();
  initThemeSelectors();
  initLocalStorage();
  initPWA();

  // Chart dropdown change listeners
  ["chart-type-select", "chart-cat-select", "chart-metric-select", "chart-agg-select"].forEach(id => {
    document.getElementById(id)?.addEventListener("change", () => {
      updateChartPreview();
      calculateLiveKPIs(state.parsedData);
    });
  });

  // Action button listeners
  document.getElementById("load-sample-btn")?.addEventListener("click", loadSampleData);
  document.getElementById("run-pipeline-btn")?.addEventListener("click", executePythonPipeline);
  document.getElementById("clear-console-btn")?.addEventListener("click", () => {
    document.getElementById("console-output").textContent = "Console cleared.";
    sessionStorage.removeItem(CONSOLE_LOG_KEY);
    document.getElementById("console-error-badge")?.classList.add("hidden");
  });
  document.getElementById("copy-console-btn")?.addEventListener("click", () => {
    const text = document.getElementById("console-output")?.innerText || "";
    navigator.clipboard.writeText(text)
      .then(() => showToast("Console logs copied to clipboard", "success"))
      .catch(() => showToast("Could not copy logs", "error"));
  });

  // Password toggle
  document.getElementById("toggle-password-btn")?.addEventListener("click", () => {
    const passInput = document.getElementById("smtp-pass");
    passInput.type = passInput.type === "password" ? "text" : "password";
  });

  // Launch Pyodide WebAssembly Engine
  initPyodideEngine();
});
