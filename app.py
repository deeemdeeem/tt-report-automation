import os
import io
import tempfile
from datetime import datetime

from flask import Flask, render_template_string, request, send_file, redirect, url_for, flash
from werkzeug.utils import secure_filename

import pandas as pd
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# ---------------- Flask setup ----------------
app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "dev-key")

# PPT Template on file directory
PPT_TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "TT_report.pptx")
ALLOWED_EXCEL_EXTS = {".xlsm", ".xlsx"}

HTML = """
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>TT Report Generator</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap" rel="stylesheet">
  <style>
    :root{
      --bg:#0b1220;        /* deep navy */
      --card:#0f172a;      /* slate-900 */
      --muted:#94a3b8;     /* slate-400 */
      --text:#e2e8f0;      /* slate-200 */
      --accent:#6d28d9;    /* violet-700 */
      --accent-hover:#7c3aed;
      --border:#1f2937;    /* slate-800 */
      --good:#10b981;      /* emerald */
      --warn:#f59e0b;      /* amber */
    }
    *{ box-sizing: border-box; }
    html,body{ height:100%; }
    body{
      margin:0; background:var(--bg); color:var(--text);
      font-family: Inter, system-ui, -apple-system, Segoe UI, Roboto, Arial, sans-serif;
    }
    .header{
      position: sticky; top:0; background:rgba(15,23,42,.8);
      backdrop-filter: blur(8px);
      border-bottom:1px solid var(--border);
    }
    .header-inner{
      max-width:900px; margin:0 auto; padding:12px 20px; display:flex; align-items:center; gap:12px;
    }
    .header img{ height:28px; display:block; }
    .wrap{ max-width:900px; margin:36px auto; padding:0 20px; }
    .card{
      background:var(--card); border:1px solid var(--border);
      border-radius:16px; padding:24px; box-shadow:0 10px 30px rgba(0,0,0,.35);
    }
    h1, h2{ margin:0 0 6px; font-weight:700; letter-spacing:.2px; }
    h1{ font-size:26px; }
    h2{ font-size:20px; color:#fff; }
    p.sub{ margin:0 0 18px; color:var(--muted); }
    form{ display:flex; flex-wrap:wrap; gap:12px; align-items:center; }
    .file-wrap{
      position:relative; display:flex; align-items:center; gap:10px;
      border:1px dashed var(--border); background:rgba(2,6,23,.5);
      padding:12px 14px; border-radius:12px; min-width:320px; flex:1;
    }
    input[type=file]{ position:absolute; inset:0; opacity:0; cursor:pointer; }
    .file-label{ color:var(--muted); font-size:14px; }
    .file-name{ font-size:14px; color:var(--text); white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
    .btn{
      border:none; padding:12px 16px; border-radius:12px; font-weight:600; cursor:pointer;
      transition: transform .04s ease;
    }
    .btn:active{ transform: translateY(1px); }
    .btn-primary{ background:var(--accent); color:#fff; }
    .btn-primary:hover{ background:var(--accent-hover); }
    .btn-ghost{ background:transparent; border:1px solid var(--border); color:var(--text); }
    .hint{ margin-top:10px; color:var(--muted); font-size:12px; }
    .flash{
      background:rgba(245,158,11,.1); color:#fde68a;
      border:1px solid rgba(245,158,11,.35);
      padding:10px 12px; border-radius:10px; margin-bottom:12px;
    }
    /* --- Loading overlay + spinner --- */
    .overlay{
      position: fixed; inset:0; display:none; place-items:center;
      background: rgba(2,6,23,.55); z-index: 9999;
    }
    .spinner{
      width:56px; height:56px; border-radius:50%;
      border:4px solid rgba(255,255,255,.15);
      border-top-color: #a78bfa; /* lighter violet */
      animation: spin 0.9s linear infinite;
      box-shadow: 0 0 30px rgba(167,139,250,.45);
    }
    @keyframes spin { to { transform: rotate(360deg); } }
  </style>
</head>
<body>
  <!-- Header with TruTrade logo -->
  <div class="header">
    <div class="header-inner">
      <img alt="TruTrade" src="https://trutradebeta.alexanderbabbage.com/static/media/trutrade_logo.8ab17e85dea03ec4e762.png" />
    </div>
  </div>

  <div class="wrap">
    <!-- Step 1: Download template -->
    <div class="card" style="margin-bottom:16px;">
      <h2>Step 1: Download Template</h2>
      <p class="sub">
        Download the <strong>TruTrade Executive Report Worksheet (.xlsm)</strong>. Paste your TruTrade datasets into the template and the built-in macros will automatically analyze each tab.<br>
        <em>Note:</em> after you load each dataset, allow a short processing time while the template finishes its analysis.
      </p>
      <a class="btn btn-primary" href="{{ url_for('download_template') }}">Download .xlsm Template</a>
    </div>

    <!-- Step 2: Upload & Generate -->
    <div class="card">
      <h2>Step 2: Upload & Generate Report</h2>
      <p class="sub">
        Upload the <strong>filled worksheet (.xlsm or .xlsx)</strong> here, then click <strong>Generate</strong>. We’ll populate the PowerPoint template and download a new <strong>.pptx</strong> for you.
        <br><em>Heads up:</em> maps and charts are <strong>not</strong> auto-generated in this version.
      </p>

      {% with messages = get_flashed_messages() %}
        {% if messages %}
          {% for m in messages %}<div class="flash">{{ m }}</div>{% endfor %}
        {% endif %}
      {% endwith %}

      <form id="genForm" action="{{ url_for('generate') }}" method="post" enctype="multipart/form-data" onreset="resetName();">
        <div class="file-wrap">
          <span class="file-label">Choose File</span>
          <span id="fileName" class="file-name">No file selected</span>
          <input id="fileInput" type="file" name="xlsm" accept=".xlsm,.xlsx" required>
        </div>

        <button id="genBtn" class="btn btn-primary" type="submit">Generate</button>
        <button id="clearBtn" class="btn btn-ghost" type="reset">Clear</button>
      </form>

      <div class="hint">Template in use: <code>{{ template_name }}</code></div>
    </div>
  </div>

  <!-- Loading overlay -->
  <div id="overlay" class="overlay" aria-hidden="true">
    <div class="spinner" role="status" aria-label="Generating report..."></div>
  </div>

  <script>
    const input   = document.getElementById('fileInput');
    const nameEl  = document.getElementById('fileName');
    const form    = document.getElementById('genForm');
    const genBtn  = document.getElementById('genBtn');
    const clearBtn= document.getElementById('clearBtn');
    const overlay = document.getElementById('overlay');

    function resetName(){ nameEl.textContent = 'No file selected'; }
    input.addEventListener('change', () => {
      nameEl.textContent = input.files.length ? input.files[0].name : 'No file selected';
    });

    function showLoading(){
      overlay.style.display = 'grid';
      genBtn.disabled = true;
      clearBtn.disabled = true;
      genBtn.textContent = 'Generating…';
      genBtn.setAttribute('aria-busy','true');
    }
    function hideLoading(){
      overlay.style.display = 'none';
      genBtn.disabled = false;
      clearBtn.disabled = false;
      genBtn.textContent = 'Generate';
      genBtn.removeAttribute('aria-busy');
    }

    // Intercept submit so we can show spinner and control the download
    form.addEventListener('submit', async (e) => {
      e.preventDefault();
      if (!input.files.length) return;

      try{
        showLoading();
        const fd = new FormData(form);
        const res = await fetch(form.action, { method: 'POST', body: fd });
        if(!res.ok){
          hideLoading();
          alert('Error generating report. Please check your file and try again.');
          return;
        }
        const blob = await res.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        const ts = new Date().toISOString().slice(0,16).replace('T','_');
        a.href = url;
        a.download = `TT_report_${ts}.pptx`;
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(url);
      }catch(err){
        console.error(err);
        alert('Unexpected error. Please try again.');
      }finally{
        hideLoading();
      }
    });
  </script>
</body>
</html>

"""
# Logic 
def build_presentation(xlsm_path: str, template_path: str) -> io.BytesIO:
    # Load template PPT
    prs = Presentation(template_path)

    # Read Excel sheets (openpyxl reads .xlsm/.xlsx; macros aren’t executed)
    dfs = pd.read_excel(
        xlsm_path,
        sheet_name=[
            "LeasingInfographic", "CompetitiveMarketPosition", "ZipCodes",
            "DrawDemo", "DistanceTravelled", "Frequency", "Duration", "MileageDemo"
        ],
        engine="openpyxl",
    )
    df_leasing = dfs["LeasingInfographic"]
    df_sheet2 = dfs["CompetitiveMarketPosition"]
    df_zipcodes = dfs["ZipCodes"]
    df_drawdemos = dfs["DrawDemo"]
    df_distance = dfs["DistanceTravelled"]
    df_frequency = dfs["Frequency"]
    df_duration = dfs["Duration"]
    df_mileage = dfs["MileageDemo"]

    variable_mapping = {
        "VL10": f"{int(round(df_leasing.iloc[0, 0] * 100, 0))}%",
        "VOP08": "{:,.0f}".format(df_leasing.iloc[0, 3]),
        "LD08": f"{int(round(df_leasing.iloc[3, 3] * 100, 0))}%",
        "MT08": f"{int(round(df_leasing.iloc[6, 3] * 100, 0))}%",
        "VF08": df_leasing.iloc[11, 3],
        "HH08": f"{int(round(df_leasing.iloc[0, 7] * 100, 0))}%",
        "HHI08": "${:,.0f}".format(df_leasing.iloc[3, 7]),
        "HHIMSA08": "${:,.0f}".format(df_leasing.iloc[3, 8]),
        "CD08": f"{int(round(df_leasing.iloc[6, 7] * 100, 0))}%",
        "VC08": f"{int(round(df_leasing.iloc[9, 7] * 100, 0))}%",
        "DT08": df_leasing.iloc[11, 7],
        "ZIP1": df_leasing.iloc[3, 0],
        "ZIP2": df_leasing.iloc[4, 0],
        "ZIP3": df_leasing.iloc[5, 0],
        "ZIP4": df_leasing.iloc[6, 0],
        "ZIP5": df_leasing.iloc[7, 0],
        "ZIPANALYSIS15": df_zipcodes.iloc[0, 14],
        "DDANALYSIS12": df_drawdemos.iloc[0, 18],
        "CMPANALYSIS10": df_sheet2.iloc[0, 9],
        "MILANALYSIS11": df_mileage.iloc[0, 3],

        # Mileage Demos mapping
        "MA121": f"{df_mileage.iloc[2, 2]:,.0f}", "MB121": f"{df_mileage.iloc[2, 3]:,.0f}", "MC121": f"{df_mileage.iloc[2, 4]:,.0f}", "MD121": f"{df_mileage.iloc[2, 5]:,.0f}",
        "MA122": f"{df_mileage.iloc[3, 2]:,.0f}", "MB122": f"{df_mileage.iloc[3, 3]:,.0f}", "MC122": f"{df_mileage.iloc[3, 4]:,.0f}", "MD122": f"{df_mileage.iloc[3, 5]:,.0f}",
        "MA123": f"{df_mileage.iloc[4, 2] * 100:.1f}%",  "MB123": f"{df_mileage.iloc[4, 3] * 100:.1f}%",  "MC123": f"{df_mileage.iloc[4, 4] * 100:.1f}%",  "MD123": f"{df_mileage.iloc[4, 5] * 100:.1f}%",
        "MA124": f"{df_mileage.iloc[5, 2] * 100:.1f}%",  "MB124": f"{df_mileage.iloc[5, 3] * 100:.1f}%",  "MC124": f"{df_mileage.iloc[5, 4] * 100:.1f}%",  "MD124": f"{df_mileage.iloc[5, 5] * 100:.1f}%",
        "MA125": f"{df_mileage.iloc[6, 2] * 100:.1f}%",  "MB125": f"{df_mileage.iloc[6, 3] * 100:.1f}%",  "MC125": f"{df_mileage.iloc[6, 4] * 100:.1f}%",  "MD125": f"{df_mileage.iloc[6, 5] * 100:.1f}%",
        "MA126": f"{df_mileage.iloc[7, 2] * 100:.1f}%",  "MB126": f"{df_mileage.iloc[7, 3] * 100:.1f}%",  "MC126": f"{df_mileage.iloc[7, 4] * 100:.1f}%",  "MD126": f"{df_mileage.iloc[7, 5] * 100:.1f}%",
        "MA127": f"{df_mileage.iloc[8, 2] * 100:.1f}%",  "MB127": f"{df_mileage.iloc[8, 3] * 100:.1f}%",  "MC127": f"{df_mileage.iloc[8, 4] * 100:.1f}%",  "MD127": f"{df_mileage.iloc[8, 5] * 100:.1f}%",
        "MA128": f"{df_mileage.iloc[9, 2] * 100:.1f}%",  "MB128": f"{df_mileage.iloc[9, 3] * 100:.1f}%",  "MC128": f"{df_mileage.iloc[9, 4] * 100:.1f}%",  "MD128": f"{df_mileage.iloc[9, 5] * 100:.1f}%",
        "MA129": f"{df_mileage.iloc[10, 2]:.1f}", "MB129": f"{df_mileage.iloc[10, 3]:.1f}", "MC129": f"{df_mileage.iloc[10, 4]:.1f}", "MD129": f"{df_mileage.iloc[10, 5]:.1f}",
        "MA130": f"{df_mileage.iloc[11, 2] * 100:.1f}%", "MB130": f"{df_mileage.iloc[11, 3] * 100:.1f}%", "MC130": f"{df_mileage.iloc[11, 4] * 100:.1f}%", "MD130": f"{df_mileage.iloc[11, 5] * 100:.1f}%",
        "MA131": f"{df_mileage.iloc[12, 2] * 100:.1f}%", "MB131": f"{df_mileage.iloc[12, 3] * 100:.1f}%", "MC131": f"{df_mileage.iloc[12, 4] * 100:.1f}%", "MD131": f"{df_mileage.iloc[12, 5] * 100:.1f}%",
        "MA132": f"{df_mileage.iloc[13, 2] * 100:.1f}%", "MB132": f"{df_mileage.iloc[13, 3] * 100:.1f}%", "MC132": f"{df_mileage.iloc[13, 4] * 100:.1f}%", "MD132": f"{df_mileage.iloc[13, 5] * 100:.1f}%",
        "MA133": f"{df_mileage.iloc[14, 2] * 100:.1f}%", "MB133": f"{df_mileage.iloc[14, 3] * 100:.1f}%", "MC133": f"{df_mileage.iloc[14, 4] * 100:.1f}%", "MD133": f"{df_mileage.iloc[14, 5] * 100:.1f}%",
        "MA134": f"{df_mileage.iloc[15, 2] * 100:.1f}%", "MB134": f"{df_mileage.iloc[15, 3] * 100:.1f}%", "MC134": f"{df_mileage.iloc[15, 4] * 100:.1f}%", "MD134": f"{df_mileage.iloc[15, 5] * 100:.1f}%",
        "MA135": "${:,.0f}".format(df_mileage.iloc[16, 2]), "MB135": "${:,.0f}".format(df_mileage.iloc[16, 3]), "MC135": "${:,.0f}".format(df_mileage.iloc[16, 4]), "MD135": "${:,.0f}".format(df_mileage.iloc[16, 5]),
        "MA136": f"{df_mileage.iloc[17, 2] * 100:.1f}%", "MB136": f"{df_mileage.iloc[17, 3] * 100:.1f}%", "MC136": f"{df_mileage.iloc[17, 4] * 100:.1f}%", "MD136": f"{df_mileage.iloc[17, 5] * 100:.1f}%",
        "MA137": f"{df_mileage.iloc[18, 2] * 100:.1f}%", "MB137": f"{df_mileage.iloc[18, 3] * 100:.1f}%", "MC137": f"{df_mileage.iloc[18, 4] * 100:.1f}%", "MD137": f"{df_mileage.iloc[18, 5] * 100:.1f}%",
        "MA138": f"{df_mileage.iloc[19, 2] * 100:.1f}%", "MB138": f"{df_mileage.iloc[19, 3] * 100:.1f}%", "MC138": f"{df_mileage.iloc[19, 4] * 100:.1f}%", "MD138": f"{df_mileage.iloc[19, 5] * 100:.1f}%",
        "MA139": f"{df_mileage.iloc[20, 2] * 100:.1f}%", "MB139": f"{df_mileage.iloc[20, 3] * 100:.1f}%", "MC139": f"{df_mileage.iloc[20, 4] * 100:.1f}%", "MD139": f"{df_mileage.iloc[20, 5] * 100:.1f}%",
        "MA140": f"{df_mileage.iloc[21, 2] * 100:.1f}%", "MB140": f"{df_mileage.iloc[21, 3] * 100:.1f}%", "MC140": f"{df_mileage.iloc[21, 4] * 100:.1f}%", "MD140": f"{df_mileage.iloc[21, 5] * 100:.1f}%",
        "MA141": f"{df_mileage.iloc[22, 2] * 100:.1f}%", "MB141": f"{df_mileage.iloc[22, 3] * 100:.1f}%", "MC141": f"{df_mileage.iloc[22, 4] * 100:.1f}%", "MD141": f"{df_mileage.iloc[22, 5] * 100:.1f}%",
        "MA142": f"{df_mileage.iloc[23, 2] * 100:.1f}%", "MB142": f"{df_mileage.iloc[23, 3] * 100:.1f}%", "MC142": f"{df_mileage.iloc[23, 4] * 100:.1f}%", "MD142": f"{df_mileage.iloc[23, 5] * 100:.1f}%",
        "MA143": f"{df_mileage.iloc[24, 2] * 100:.1f}%", "MB143": f"{df_mileage.iloc[24, 3] * 100:.1f}%", "MC143": f"{df_mileage.iloc[24, 4] * 100:.1f}%", "MD143": f"{df_mileage.iloc[24, 5] * 100:.1f}%",
        "MA144": f"{df_mileage.iloc[25, 2] * 100:.1f}%", "MB144": f"{df_mileage.iloc[25, 3] * 100:.1f}%", "MC144": f"{df_mileage.iloc[25, 4] * 100:.1f}%", "MD144": f"{df_mileage.iloc[25, 5] * 100:.1f}%",
        "MA145": f"{df_mileage.iloc[26, 2] * 100:.1f}%", "MB145": f"{df_mileage.iloc[26, 3] * 100:.1f}%", "MC145": f"{df_mileage.iloc[26, 4] * 100:.1f}%", "MD145": f"{df_mileage.iloc[26, 5] * 100:.1f}%"
    }

    # Replace placeholders across shapes and tables
    for slide in prs.slides:
        for shape in slide.shapes:
            if hasattr(shape, "text_frame") and shape.text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        for key, value in variable_mapping.items():
                            if key in run.text:
                                run.text = run.text.replace(key, str(value))

            if getattr(shape, "has_table", False):
                for row in shape.table.rows:
                    for cell in row.cells:
                        for key, value in variable_mapping.items():
                            if key in cell.text:
                                cell.text = cell.text.replace(key, str(value))
                                for paragraph in cell.text_frame.paragraphs:
                                    paragraph.alignment = PP_ALIGN.CENTER
                                    for run in paragraph.runs:
                                        run.font.name = "Roboto"
                                        run.font.size = Pt(9)
                                        run.font.color.rgb = RGBColor(0, 0, 0)

    # Table copy/format rules
    slides_to_update = {
        9: df_sheet2,
        11: df_drawdemos,
        14: df_zipcodes,
        36: df_distance,
        37: df_frequency,
        38: df_duration
    }

    formatting_rules = {
        "CompetitiveMarketPosition": {
            "percent_columns": [3, 5, 6],
            "currency_columns": [2],
            "thousands_columns": [1]
        },
        "ZipCodes": {
            "percent_columns": [4, 5, 8],
            "currency_columns": [9],
            "thousands_columns": [6, 7]
        },
        "DrawDemo": {
            "percent_rows": [
                "18-24", "25-34", "35-44", "45-54", "55-64", "65+",
                "Less than $50,000", "$50,000-$74,999", "$75,000-$99,999",
                "$100,000-$149,999", "$150,000 or more", "CHILDREN IN HOUSEHOLD",
                "Less than college", "Some college", "College degree", "Post-graduate degree",
                "Caucasian/White", "African-American/Black", "Hispanic/Latino",
                "Asian", "Other"
            ],
            "currency_rows": ["HOUSEHOLD INCOME", "Average HH Income"]
        },
        "DistanceTravelled": {
            "percent_columns": [1, 2, 3, 4, 5],
            "decimal_columns": [6, 7]
        },
        "Frequency": {
            "percent_columns": [1, 2, 3],
            "decimal_columns": [4]
        },
        "Duration": {
            "percent_columns": [1, 2, 3, 4],
            "decimal_columns": [5]
        }
    }

    for slide_number, df_data in slides_to_update.items():
        slide = prs.slides[slide_number]
        sheet_name = [name for name, df in dfs.items() if df.equals(df_data)][0]
        rules = formatting_rules.get(sheet_name, {})
        table = next((s.table for s in slide.shapes if getattr(s, "has_table", False)), None)
        if not table:
            continue

        rows, cols = df_data.shape

        # Header styling where needed
        if sheet_name in ("DrawDemo", "DistanceTravelled", "Frequency", "Duration", "CompetitiveMarketPosition"):
            for col_index, col_name in enumerate(df_data.columns):
                if col_index < len(table.columns):
                    cell = table.cell(0, col_index)
                    cell.text = str(col_name)
                    for paragraph in cell.text_frame.paragraphs:
                        paragraph.alignment = PP_ALIGN.CENTER if col_index == 0 else PP_ALIGN.LEFT
                        for run in paragraph.runs:
                            run.font.name = "Roboto"
                            run.font.size = Pt(9)
                            run.font.color.rgb = RGBColor(255, 255, 255)
                            run.font.bold = True

        for row_index in range(min(rows, len(table.rows) - 1)):
            for col_index in range(min(cols, len(table.columns))):
                value = df_data.iloc[row_index, col_index]
                value = "" if pd.isna(value) else value
                row_label = str(df_data.iloc[row_index, 0])

                if isinstance(value, (int, float)):
                    if sheet_name == "DrawDemo":
                        if any(keyword in row_label for keyword in rules.get("percent_rows", [])):
                            formatted_value = f"{round(value * 100, 1)}%"
                        elif any(keyword in row_label for keyword in rules.get("currency_rows", [])):
                            formatted_value = "${:,.0f}".format(value)
                        else:
                            formatted_value = str(value)
                    else:
                        if col_index in rules.get("percent_columns", []):
                            formatted_value = f"{round(value * 100, 1)}%"
                        elif col_index in rules.get("currency_columns", []):
                            formatted_value = "${:,.0f}".format(value)
                        elif col_index in rules.get("thousands_columns", []):
                            formatted_value = "{:,.0f}".format(value)
                        elif col_index in rules.get("decimal_columns", []):
                            formatted_value = f"{value:.1f}"
                        else:
                            formatted_value = str(value)
                else:
                    formatted_value = str(value)

                cell = table.cell(row_index + 1, col_index)
                cell.text = formatted_value
                for paragraph in cell.text_frame.paragraphs:
                    for run in paragraph.runs:
                        run.font.name = "Roboto"
                        run.font.size = Pt(9)
                        if sheet_name in ("DistanceTravelled", "Frequency", "Duration") and row_index == 0:
                            run.font.color.rgb = RGBColor(255, 255, 255)
                            run.font.bold = True
                        else:
                            run.font.color.rgb = RGBColor(0, 0, 0)

    # Write PPT to memory buffer and return
    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output

# Routers
@app.route("/")
def index():
    if not os.path.exists(PPT_TEMPLATE_PATH):
        flash("Template not found: put TT_report.pptx beside app.py")
    return render_template_string(HTML, template_name=os.path.basename(PPT_TEMPLATE_PATH))

@app.route("/download-template")
def download_template():
    return send_file(
        "TT_worksheet.xlsm",
        as_attachment=True,
        download_name="TT_worksheet.xlsm",
        mimetype="application/vnd.ms-excel.sheet.macroEnabled.12"
    )

@app.route("/generate", methods=["POST"])
def generate():
    file = request.files.get("xlsm")
    if not file or file.filename == "":
        flash("Please choose an .xlsm or .xlsx file.")
        return redirect(url_for("index"))

    ext = os.path.splitext(file.filename)[1].lower()
    if ext not in ALLOWED_EXCEL_EXTS:
        flash("Unsupported file type. Upload .xlsm or .xlsx.")
        return redirect(url_for("index"))

    with tempfile.TemporaryDirectory() as tmpdir:
        safe_name = secure_filename(file.filename)
        xlsm_path = os.path.join(tmpdir, safe_name)
        file.save(xlsm_path)

        try:
            output = build_presentation(xlsm_path, PPT_TEMPLATE_PATH)
        except Exception as e:
            flash(f"Error generating PPT: {e}")
            return redirect(url_for("index"))

    ts = datetime.now().strftime("%Y-%m-%d_%H-%M")
    return send_file(
        output,
        as_attachment=True,
        download_name=f"TT_report_{ts}.pptx",
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )

# --------------- Run locally ---------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
