#!/usr/bin/env python3
"""
Elztec Dashboard Generator
Leest de Exact Online Excel export en genereert index.html
"""

import pandas as pd
import base64
import os
import sys
from pathlib import Path

# ── Configuratie ──────────────────────────────────────────────
EXCEL_PATH = "data/financieel.xlsx"
LOGO_PATH  = "data/logo.png"
OUTPUT     = "index.html"

def fmt(val):
    """Formatteer getal als euro bedrag (Nederlandse notatie)"""
    val = abs(val)
    return f"€ {val:,.0f}".replace(",", ".")

def pct(result, omzet):
    """Bereken marge percentage"""
    if omzet == 0:
        return 0
    return round((result / omzet) * 100, 1)

def extract_data(path):
    """Lees de Exact export en extraheer de kerngetallen"""
    df = pd.read_excel(path, header=None)

    # Detecteer maandkolommen (rij 11 bevat datums)
    date_row = df.iloc[11]
    months = {}
    for col in range(1, len(date_row)):
        val = date_row[col]
        if pd.notna(val) and hasattr(val, 'month'):
            months[col] = val.strftime("%B %Y")

    # Vind de kernrijen
    data = {}
    for i, row in df.iterrows():
        label = str(row[0]).strip()
        if label == "Totaal: Omzet":
            data["omzet_row"] = i
        elif label == "Totaal: Kosten":
            data["kosten_row"] = i
        elif label == "Resultaat":
            data["resultaat_row"] = i
        elif label == "Totaal: Kostprijs van de omzet":
            data["kostprijs_row"] = i
        elif label == "Totaal: Marge":
            data["marge_row"] = i

    # Extraheer per maand
    maanden = []
    for col, naam in sorted(months.items()):
        omzet    = abs(float(df.iloc[data["omzet_row"], col] or 0))
        kosten   = float(df.iloc[data["kosten_row"], col] or 0)
        resultaat = float(df.iloc[data["resultaat_row"], col] or 0)
        kostprijs = float(df.iloc[data["kostprijs_row"], col] or 0) if "kostprijs_row" in data else 0
        bruto_marge = abs(float(df.iloc[data["marge_row"], col] or 0)) if "marge_row" in data else omzet - kostprijs

        if omzet > 0 or kosten > 0 or resultaat != 0:
            maanden.append({
                "naam": naam,
                "omzet": omzet,
                "kosten": kosten,
                "resultaat": resultaat,
                "bruto_marge": bruto_marge,
                "marge_pct": pct(resultaat, omzet)
            })

    # YTD totalen (laatste kolom)
    ytd_col = max(months.keys()) + 1
    try:
        ytd_omzet     = abs(float(df.iloc[data["omzet_row"], -1] or 0))
        ytd_kosten    = float(df.iloc[data["kosten_row"], -1] or 0)
        ytd_resultaat = float(df.iloc[data["resultaat_row"], -1] or 0)
    except:
        ytd_omzet     = sum(m["omzet"] for m in maanden)
        ytd_kosten    = sum(m["kosten"] for m in maanden)
        ytd_resultaat = sum(m["resultaat"] for m in maanden)

    # Detailregels per categorie
    def get_category_rows(start_label, end_label):
        rows = []
        in_section = False
        for i, row in df.iterrows():
            label = str(row[0]).strip()
            if label == start_label:
                in_section = True
                continue
            if label == end_label:
                break
            if in_section and pd.notna(row[0]) and not label.startswith("Totaal"):
                row_data = {"naam": label}
                for col, maand in sorted(months.items()):
                    val = row[col]
                    row_data[maand] = abs(float(val)) if pd.notna(val) and val != 0 else None
                rows.append(row_data)
        return rows

    omzet_regels   = get_category_rows("Omzet", "Totaal: Omzet")
    kosten_regels  = get_category_rows("Kostprijs van de omzet", "Totaal: Kostprijs van de omzet")

    return maanden, ytd_omzet, ytd_kosten, ytd_resultaat, omzet_regels, kosten_regels

def load_logo(path):
    """Laad logo als base64 of geef lege string terug"""
    if os.path.exists(path):
        with open(path, 'rb') as f:
            ext = Path(path).suffix.lower().replace('.', '')
            if ext == 'svg':
                mime = 'image/svg+xml'
            else:
                mime = f'image/{ext}'
            return f"data:{mime};base64,{base64.b64encode(f.read()).decode()}"
    return ""

def generate_html(maanden, ytd_omzet, ytd_kosten, ytd_resultaat, omzet_regels, kosten_regels, logo_src):
    """Genereer de volledige HTML pagina"""

    ytd_marge_pct = pct(ytd_resultaat, ytd_omzet)
    now = pd.Timestamp.now()
    export_datum = f"{now.day} {now.strftime('%B %Y')}"

    # Bouw tab inhoud
    def maand_tabel(m):
        rows = ""
        for r in omzet_regels:
            val = r.get(m["naam"])
            if val:
                rows += f'<tr><td>{r["naam"]}</td><td class="c-blue">{fmt(val)}</td></tr>\n'
        rows += f'<tr class="row-sub"><td>Totaal Omzet</td><td class="c-blue">{fmt(m["omzet"])}</td></tr>\n'
        rows += '<tr class="row-cat"><td colspan="2">Kostprijs van de Omzet</td></tr>\n'
        for r in kosten_regels:
            val = r.get(m["naam"])
            if val:
                rows += f'<tr><td>{r["naam"]}</td><td class="c-red">{fmt(val)}</td></tr>\n'
        rows += f'<tr class="row-sub"><td>Totaal Kostprijs</td><td class="c-red">{fmt(sum(r.get(m["naam"]) or 0 for r in kosten_regels))}</td></tr>\n'
        rows += f'<tr class="row-sub"><td class="c-green">Bruto Marge</td><td class="c-green">{fmt(m["bruto_marge"])}</td></tr>\n'
        rows += f'<tr class="row-sub"><td>Totaal Bedrijfskosten</td><td class="c-red">{fmt(m["kosten"] - sum(r.get(m["naam"]) or 0 for r in kosten_regels))}</td></tr>\n'
        rows += f'<tr class="row-total"><td>&#10022; Nettoresultaat {m["naam"]}</td><td>{fmt(m["resultaat"])}</td></tr>\n'
        return f'''
        <table>
          <thead><tr><th>Omschrijving</th><th>Bedrag</th></tr></thead>
          <tbody>
            <tr class="row-cat"><td colspan="2">Omzet</td></tr>
            {rows}
          </tbody>
        </table>'''

    # Tabs HTML
    tabs_buttons = ""
    tabs_content = ""
    for i, m in enumerate(maanden):
        active = "active" if i == 0 else ""
        display = "" if i == 0 else 'style="display:none"'
        tabs_buttons += f'<button class="tab {active}" onclick="showTab(\'m{i}\',this)">{m["naam"]}</button>\n'
        tabs_content += f'<div id="tab-m{i}" {display}>{maand_tabel(m)}</div>\n'

    # YTD tab
    ytd_rows = ""
    for m in maanden:
        ytd_rows += f'<tr class="row-sub"><td>{m["naam"]}</td><td class="c-blue">{fmt(m["omzet"])}</td><td class="c-red">{fmt(m["kosten"])}</td><td class="c-{"green" if m["resultaat"] >= 0 else "red"}">{fmt(m["resultaat"])}</td><td style="color:#5b5fef;font-weight:700">{m["marge_pct"]}%</td></tr>\n'
    ytd_rows += f'<tr class="row-total"><td>&#10022; YTD Totaal</td><td>{fmt(ytd_omzet)}</td><td>{fmt(ytd_kosten)}</td><td>{fmt(ytd_resultaat)}</td><td>{ytd_marge_pct}%</td></tr>\n'

    tabs_buttons += '<button class="tab" onclick="showTab(\'ytd\',this)">YTD</button>\n'
    tabs_content += f'''<div id="tab-ytd" style="display:none">
      <table>
        <thead><tr><th>Maand</th><th>Omzet</th><th>Kosten</th><th>Resultaat</th><th>Marge %</th></tr></thead>
        <tbody>{ytd_rows}</tbody>
      </table></div>'''

    # KPI cards
    kpi_cards = ""
    for m in maanden:
        kpi_cards += f'''
        <div class="kpi-month">
          <div class="kpi-month-name">{m["naam"]}</div>
          <div class="kpi-val">omzet: {fmt(m["omzet"])}</div>
        </div>'''

    # Chart data
    chart_labels = str([m["naam"] for m in maanden])
    chart_omzet  = str([m["omzet"] for m in maanden])
    chart_kosten = str([m["kosten"] for m in maanden])
    chart_result = str([m["resultaat"] for m in maanden])

    # Marge bars (eerste 2 maanden)
    marge_bars = ""
    for i, m in enumerate(maanden[:3]):
        colors = ["#0e7fc2", "#5b5fef", "#0ea472"]
        color = colors[i % len(colors)]
        marge_bars += f'''
        <div class="marge-item">
          <div class="marge-row">
            <span class="marge-name">{m["naam"]}</span>
            <span class="marge-pct" style="color:{color}">{m["marge_pct"]}%</span>
          </div>
          <div class="marge-track"><div class="marge-fill" id="mf{i}" style="width:0%;background:{color}"></div></div>
          <div class="marge-note">{fmt(m["resultaat"])} op {fmt(m["omzet"])} omzet</div>
        </div>'''

    marge_js = "\n".join([f'document.getElementById("mf{i}").style.width="{m["marge_pct"]}%";'
                          for i, m in enumerate(maanden[:3])])

    logo_html = f'<img class="logo-img" src="{logo_src}" alt="Elztec">' if logo_src else \
                '<span style="font-size:22px;font-weight:800;color:#5b5fef">Elztec</span>'

    # Eerste twee maanden voor KPI cards
    m1 = maanden[0] if len(maanden) > 0 else {"naam": "-", "omzet": 0, "kosten": 0, "resultaat": 0}
    m2 = maanden[1] if len(maanden) > 1 else {"naam": "-", "omzet": 0, "kosten": 0, "resultaat": 0}

    return f"""<!DOCTYPE html>
<html lang="nl">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Elztec Financieel Dashboard</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<style>
* {{ margin:0; padding:0; box-sizing:border-box; }}
body {{ font-family:'Inter',sans-serif; background:#f4f6fb; color:#1a1b2e; min-height:100vh; }}
body::before {{ content:""; position:fixed; inset:0; background-image:linear-gradient(rgba(91,95,239,0.04) 1px,transparent 1px),linear-gradient(90deg,rgba(91,95,239,0.04) 1px,transparent 1px); background-size:40px 40px; pointer-events:none; z-index:0; }}
.wrapper {{ position:relative; z-index:1; max-width:1160px; margin:0 auto; padding:32px 24px 60px; }}
header {{ display:flex; align-items:center; justify-content:space-between; background:#fff; border:1px solid #e2e4f0; border-radius:14px; padding:20px 28px; margin-bottom:28px; box-shadow:0 2px 10px rgba(0,0,0,0.06); flex-wrap:wrap; gap:12px; }}
.logo-area {{ display:flex; align-items:center; gap:16px; }}
.logo-img {{ height:40px; width:auto; }}
.divider {{ width:1px; height:34px; background:#e2e4f0; }}
.header-text {{ display:flex; flex-direction:column; gap:3px; }}
.header-title {{ font-size:15px; font-weight:700; }}
.header-sub {{ font-family:'JetBrains Mono',monospace; font-size:10px; color:#8a8ca8; letter-spacing:1px; text-transform:uppercase; }}
.header-tags {{ display:flex; gap:8px; flex-wrap:wrap; }}
.tag {{ font-family:'JetBrains Mono',monospace; font-size:10px; padding:5px 12px; border-radius:20px; font-weight:600; letter-spacing:0.5px; text-transform:uppercase; }}
.tag-purple {{ background:rgba(91,95,239,0.1); color:#5b5fef; border:1px solid rgba(91,95,239,0.2); }}
.tag-gray {{ background:#f0f1f7; color:#8a8ca8; border:1px solid #e2e4f0; }}
.section-title {{ font-family:'JetBrains Mono',monospace; font-size:10px; letter-spacing:2px; text-transform:uppercase; color:#8a8ca8; margin-bottom:12px; }}
.kpi-row {{ display:grid; grid-template-columns:repeat(3,1fr); gap:16px; margin-bottom:24px; }}
.kpi {{ background:#fff; border:1px solid #e2e4f0; border-radius:14px; padding:22px 20px; position:relative; overflow:hidden; box-shadow:0 2px 8px rgba(0,0,0,0.04); transition:transform 0.2s,box-shadow 0.2s; }}
.kpi:hover {{ transform:translateY(-2px); box-shadow:0 6px 20px rgba(91,95,239,0.1); }}
.kpi-bar {{ position:absolute; top:0; left:0; right:0; height:3px; border-radius:14px 14px 0 0; }}
.kpi-bar.blue {{ background:linear-gradient(90deg,#0e7fc2,#60c5f7); }}
.kpi-bar.red {{ background:linear-gradient(90deg,#e83a5a,#f97090); }}
.kpi-bar.purple {{ background:linear-gradient(90deg,#5b5fef,#7c7ff5); }}
.kpi-label {{ font-size:11px; font-weight:600; color:#8a8ca8; text-transform:uppercase; letter-spacing:0.5px; margin-bottom:16px; }}
.kpi-months {{ display:flex; gap:16px; margin-bottom:14px; }}
.kpi-month {{ flex:1; }}
.kpi-month-name {{ font-family:'JetBrains Mono',monospace; font-size:9px; color:#8a8ca8; text-transform:uppercase; letter-spacing:1px; margin-bottom:4px; }}
.kpi-val {{ font-size:20px; font-weight:800; letter-spacing:-0.5px; }}
.kpi-val.blue {{ color:#0e7fc2; }} .kpi-val.red {{ color:#e83a5a; }} .kpi-val.purple {{ color:#5b5fef; }}
.kpi-sep {{ height:1px; background:#e2e4f0; margin-bottom:12px; }}
.kpi-ytd {{ display:flex; justify-content:space-between; align-items:center; }}
.kpi-ytd-label {{ font-family:'JetBrains Mono',monospace; font-size:9px; color:#8a8ca8; text-transform:uppercase; letter-spacing:1px; }}
.kpi-ytd-val {{ font-size:14px; font-weight:700; }}
.charts-row {{ display:grid; grid-template-columns:1.6fr 1fr; gap:16px; margin-bottom:24px; }}
.card {{ background:#fff; border:1px solid #e2e4f0; border-radius:14px; padding:22px; box-shadow:0 2px 8px rgba(0,0,0,0.04); }}
.card-header {{ display:flex; align-items:center; justify-content:space-between; margin-bottom:20px; gap:8px; }}
.card-title {{ font-size:13px; font-weight:700; }}
.card-badge {{ font-family:'JetBrains Mono',monospace; font-size:9px; background:#f0f1f7; color:#8a8ca8; border:1px solid #e2e4f0; padding:3px 9px; border-radius:20px; text-transform:uppercase; letter-spacing:0.5px; }}
.marge-list {{ display:flex; flex-direction:column; gap:16px; }}
.marge-item {{ display:flex; flex-direction:column; gap:5px; }}
.marge-row {{ display:flex; justify-content:space-between; align-items:baseline; }}
.marge-name {{ font-size:13px; font-weight:600; }}
.marge-pct {{ font-family:'JetBrains Mono',monospace; font-size:12px; font-weight:700; }}
.marge-track {{ height:7px; background:#f0f1f7; border-radius:4px; overflow:hidden; }}
.marge-fill {{ height:100%; border-radius:4px; transition:width 1.2s ease; }}
.marge-note {{ font-family:'JetBrains Mono',monospace; font-size:10px; color:#8a8ca8; }}
.marge-sep {{ height:1px; background:#e2e4f0; }}
.breakdown {{ background:#fff; border:1px solid #e2e4f0; border-radius:14px; padding:22px; margin-bottom:24px; box-shadow:0 2px 8px rgba(0,0,0,0.04); }}
.tabs {{ display:flex; gap:4px; background:#f0f1f7; padding:4px; border-radius:10px; width:fit-content; border:1px solid #e2e4f0; flex-wrap:wrap; }}
.tab {{ padding:6px 14px; font-size:12px; font-weight:600; border-radius:7px; cursor:pointer; border:none; background:transparent; color:#8a8ca8; transition:all 0.15s; font-family:'Inter',sans-serif; }}
.tab.active {{ background:#5b5fef; color:#fff; box-shadow:0 2px 8px rgba(91,95,239,0.35); }}
table {{ width:100%; border-collapse:collapse; font-size:13px; }}
thead th {{ font-family:'JetBrains Mono',monospace; font-size:9px; text-transform:uppercase; letter-spacing:1.5px; color:#8a8ca8; padding:8px 12px; text-align:right; border-bottom:2px solid #e2e4f0; font-weight:600; }}
thead th:first-child {{ text-align:left; }}
tbody tr:hover {{ background:#f8f9fd; }}
tbody td {{ padding:10px 12px; border-bottom:1px solid #e2e4f0; text-align:right; }}
tbody td:first-child {{ text-align:left; font-weight:500; color:#1a1b2e; }}
tbody tr:last-child td {{ border-bottom:none; }}
.row-cat td {{ background:rgba(91,95,239,0.07); color:#5b5fef !important; font-size:10px !important; font-weight:700 !important; letter-spacing:1.5px; text-transform:uppercase; padding:8px 12px !important; border-bottom:none !important; }}
.row-sub td {{ background:#f0f1f7; font-weight:700 !important; }}
.row-total td {{ background:rgba(91,95,239,0.08); font-size:14px !important; font-weight:800 !important; color:#5b5fef !important; border-top:2px solid rgba(91,95,239,0.2) !important; }}
.c-blue {{ color:#0e7fc2 !important; font-weight:600; }} .c-red {{ color:#e83a5a !important; font-weight:600; }} .c-green {{ color:#0ea472 !important; font-weight:600; }}
.footer {{ display:flex; align-items:center; gap:10px; background:#fff; border:1px solid #e2e4f0; border-left:3px solid #5b5fef; border-radius:10px; padding:14px 20px; font-family:'JetBrains Mono',monospace; font-size:11px; color:#8a8ca8; box-shadow:0 2px 8px rgba(0,0,0,0.04); }}
.footer strong {{ color:#5b5fef; }}
@media(max-width:720px) {{ .kpi-row {{ grid-template-columns:1fr; }} .charts-row {{ grid-template-columns:1fr; }} }}
</style>
</head>
<body>
<div class="wrapper">

<header>
  <div class="logo-area">
    {logo_html}
    <div class="divider"></div>
    <div class="header-text">
      <div class="header-title">Financieel Dashboard</div>
      <div class="header-sub">Boekjaar 2026 &middot; Winst &amp; Verlies</div>
    </div>
  </div>
  <div class="header-tags">
    <div class="tag tag-purple">Exact Online</div>
    <div class="tag tag-gray">{export_datum}</div>
  </div>
</header>

<div class="section-title">Kerngetallen per maand</div>
<div class="kpi-row">
  <div class="kpi">
    <div class="kpi-bar blue"></div>
    <div class="kpi-label">Omzet</div>
    <div class="kpi-months">
      <div class="kpi-month"><div class="kpi-month-name">{m1["naam"]}</div><div class="kpi-val blue">{fmt(m1["omzet"])}</div></div>
      <div class="kpi-month"><div class="kpi-month-name">{m2["naam"]}</div><div class="kpi-val blue">{fmt(m2["omzet"])}</div></div>
    </div>
    <div class="kpi-sep"></div>
    <div class="kpi-ytd"><div class="kpi-ytd-label">YTD Totaal</div><div class="kpi-ytd-val" style="color:#0e7fc2">{fmt(ytd_omzet)}</div></div>
  </div>
  <div class="kpi">
    <div class="kpi-bar red"></div>
    <div class="kpi-label">Totale Kosten</div>
    <div class="kpi-months">
      <div class="kpi-month"><div class="kpi-month-name">{m1["naam"]}</div><div class="kpi-val red">{fmt(m1["kosten"])}</div></div>
      <div class="kpi-month"><div class="kpi-month-name">{m2["naam"]}</div><div class="kpi-val red">{fmt(m2["kosten"])}</div></div>
    </div>
    <div class="kpi-sep"></div>
    <div class="kpi-ytd"><div class="kpi-ytd-label">YTD Totaal</div><div class="kpi-ytd-val" style="color:#e83a5a">{fmt(ytd_kosten)}</div></div>
  </div>
  <div class="kpi">
    <div class="kpi-bar purple"></div>
    <div class="kpi-label">Nettoresultaat</div>
    <div class="kpi-months">
      <div class="kpi-month"><div class="kpi-month-name">{m1["naam"]}</div><div class="kpi-val purple">{fmt(m1["resultaat"])}</div></div>
      <div class="kpi-month"><div class="kpi-month-name">{m2["naam"]}</div><div class="kpi-val purple">{fmt(m2["resultaat"])}</div></div>
    </div>
    <div class="kpi-sep"></div>
    <div class="kpi-ytd"><div class="kpi-ytd-label">YTD Totaal</div><div class="kpi-ytd-val" style="color:#5b5fef">{fmt(ytd_resultaat)}</div></div>
  </div>
</div>

<div class="charts-row">
  <div class="card">
    <div class="card-header"><div class="card-title">Omzet &middot; Kosten &middot; Resultaat</div><div class="card-badge">2026</div></div>
    <canvas id="barChart" height="200"></canvas>
  </div>
  <div class="card">
    <div class="card-header"><div class="card-title">Nettomarge</div><div class="card-badge">% van omzet</div></div>
    <div class="marge-list">
      {marge_bars}
    </div>
  </div>
</div>

<div class="breakdown">
  <div class="card-header">
    <div class="card-title">Winst- en Verliesrekening &mdash; detail</div>
    <div class="tabs">
      {tabs_buttons}
    </div>
  </div>
  {tabs_content}
</div>

<div class="footer">
  &#128193;&nbsp;<span>Bron: <strong>Exact Online</strong> &mdash; Automatisch gegenereerd op {export_datum} &middot; ElzTec B.V.</span>
</div>
</div>

<script>
var ctx = document.getElementById("barChart").getContext("2d");
new Chart(ctx, {{
  type:"bar",
  data:{{
    labels:{chart_labels},
    datasets:[
      {{label:"Omzet",         data:{chart_omzet},  backgroundColor:"rgba(14,127,194,0.75)", borderRadius:6, borderSkipped:false}},
      {{label:"Totale Kosten", data:{chart_kosten}, backgroundColor:"rgba(232,58,90,0.70)",  borderRadius:6, borderSkipped:false}},
      {{label:"Resultaat",     data:{chart_result}, backgroundColor:"rgba(91,95,239,0.80)",  borderRadius:6, borderSkipped:false}}
    ]
  }},
  options:{{
    responsive:true,
    plugins:{{
      legend:{{position:"bottom",labels:{{font:{{family:"Inter",size:12}},color:"#1a1b2e",padding:16,boxWidth:12,boxHeight:12,usePointStyle:true,pointStyle:"rectRounded"}}}},
      tooltip:{{backgroundColor:"#fff",borderColor:"#e2e4f0",borderWidth:1,titleColor:"#1a1b2e",bodyColor:"#1a1b2e",callbacks:{{label:function(c){{return" "+c.dataset.label+": \u20ac "+c.parsed.y.toLocaleString("nl-NL");}}}}}}
    }},
    scales:{{
      x:{{grid:{{display:false}},ticks:{{font:{{family:"Inter",size:13}},color:"#1a1b2e"}}}},
      y:{{grid:{{color:"#e2e4f0"}},ticks:{{font:{{family:"JetBrains Mono",size:11}},color:"#8a8ca8",callback:function(v){{return"\u20ac "+(v/1000).toFixed(0)+"k";}}}}}}
    }}
  }}
}});
setTimeout(function(){{ {marge_js} }}, 400);
function showTab(id,btn){{
  document.querySelectorAll('[id^="tab-"]').forEach(function(el){{el.style.display="none";}});
  document.getElementById("tab-"+id).style.display="";
  document.querySelectorAll(".tab").forEach(function(b){{b.classList.remove("active");}});
  btn.classList.add("active");
}}
</script>
</body>
</html>"""

# ── Main ──────────────────────────────────────────────────────
if __name__ == "__main__":
    excel = sys.argv[1] if len(sys.argv) > 1 else EXCEL_PATH
    logo  = sys.argv[2] if len(sys.argv) > 2 else LOGO_PATH

    print(f"Lees Excel: {excel}")
    maanden, ytd_omzet, ytd_kosten, ytd_resultaat, omzet_regels, kosten_regels = extract_data(excel)
    print(f"Gevonden maanden: {[m['naam'] for m in maanden]}")

    logo_src = load_logo(logo)
    html = generate_html(maanden, ytd_omzet, ytd_kosten, ytd_resultaat, omzet_regels, kosten_regels, logo_src)

    with open(OUTPUT, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f"Dashboard gegenereerd: {OUTPUT} ({len(html):,} bytes)")
