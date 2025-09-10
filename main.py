# -*- coding: utf-8 -*-
import re
import math
import os
from io import BytesIO
from typing import Optional

import pandas as pd
from flask import Flask, render_template, request, send_file, redirect, url_for, flash

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET", "dev-key")

# ==========================================================
# ---------------------  UI / CSS  -------------------------
# ==========================================================
CSS = r"""
<style>
:root{
  --w-s: 8.5rem;     /* étroit  */
  --w-m: 12rem;      /* moyen   */
  --w-l: 18rem;      /* large   */
  --w-xl: 26rem;     /* très large */
}

*{box-sizing:border-box}
html, body { height: 100%; }
body{
  font-family: ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif;
  margin:0;
  background:#f7fafc;
  color:#0f172a;
}

/* ------- header note ------- */
.note{
  background:#fff8e6;border:1px solid #ffd48a;
  padding:12px 14px;border-radius:10px;margin:16px 0 18px
}

/* ------- full width page ------- */
.page{
  width:100%;
  max-width: none;
  padding:24px 24px 60px;
}

/* ------- toolbar (article + file + button) ------- */
.toolbar{
  background:#fff;border:1px solid #e5e7eb;border-radius:12px;
  padding:16px; display:grid; gap:14px;
  grid-template-columns: 1fr max-content;
  align-items:end;
}
.toolbar .left{display:grid; gap:10px}
.toolbar .filebar{display:flex; align-items:center; gap:10px}
.toolbar .right{display:flex; gap:10px; align-items:center}

/* button + inputs  */
input[type="text"]{
  width:100%; height:42px; border:1px solid #cbd5e1; border-radius:9px;
  padding:0 12px; font-size:15px; background:#fff;
}
input[type="file"]{ border:1px dashed #cbd5e1; padding:8px 10px; border-radius:8px; background:#fff}
button{
  height:42px; padding:0 18px; border:0; border-radius:10px; cursor:pointer;
  background:#0ea5e9; color:#fff; font-weight:600; font-size:15px;
}
button.secondary{ background:#e5e7eb; color:#111827 }

/* ------- table viewport ------- */
.viewport{
  width:100%;
  border:1px solid #e5e7eb; background:#fff;
  border-radius:12px; margin-top:14px;
  overflow:auto;  /* scroll if needed */
  max-height: 75vh;
}

/* table */
table{width:100%; border-collapse:collapse; table-layout:fixed;}
th,td{
  border:1px solid #e5e7eb; padding:8px 10px; vertical-align:top;
  white-space:normal; word-break:normal; overflow-wrap:anywhere; hyphens:auto;
}
th{position:sticky; top:0; background:#f8fafc; z-index:2; font-weight:700; text-align:center}

/* bullets */
ul{margin:0; padding-left:1.05rem}
li{margin:0.08rem 0}
.no-bullets ul{list-style:none; padding-left:0; margin:0}

/* empty dash */
.empty{color:#9CA3AF}

/* red hit */
.hit{color:#c00; font-weight:700}

/* column width helpers */
.col-s { width:var(--w-s);  min-width:var(--w-s)}
.col-m { width:var(--w-m);  min-width:var(--w-m)}
.col-l { width:var(--w-l);  min-width:var(--w-l)}
.col-xl{ width:var(--w-xl); min-width:var(--w-xl)}

/* download row */
.actions{margin:10px 0 6px}

/* spinner overlay */
#busy{
  display:none; position:fixed; inset:0; backdrop-filter:saturate(60%) blur(1px);
  background:rgba(255,255,255,.6); align-items:center; justify-content:center; z-index:50
}
.spin{
  width:42px; height:42px; border-radius:999px; border:4px solid #38bdf8; border-top-color:transparent;
  animation:rot .9s linear infinite
}
@keyframes rot{to{transform:rotate(360deg)}}
</style>
"""

# ==========================================================
# -------------  Columns & rendering options  --------------
# ==========================================================

# Screen width classes per header (adjust freely)
WIDTH_CLASS = {
    # scalaires
    "Nom de l'intimé": "col-l",
    "Ordre professionnel": "col-l",
    "Numéro de la décision": "col-m",
    "Date de la décision rendue": "col-m",
    "Nature de la décision": "col-m",
    "Période des faits": "col-m",
    "Plaidoyer de culpabilité": "col-s",
    "Total chefs": "col-s",
    "Radiation max": "col-s",
    "Total amendes": "col-m",
    "Total réprimandes": "col-s",
    "À vérifier": "col-l",
    "Date de création": "col-m",
    "Date de mise à jour": "col-m",

    # listes
    "Résumé des faits concis": "col-xl",
    "Liste des chefs et articles en infraction": "col-xl",
    "Nbr Chefs par articles": "col-l",
    "Nbr Chefs par articles par période de radiation": "col-l",
    "Liste des sanctions imposées": "col-l",
    "Nombre de chefs par articles et total amendes": "col-l",
    "Nombre de chefs par article ayant une réprimande": "col-l",
    "Autres mesures ordonnées": "col-l",
}

# Columns rendered as bullet lists when they contain multiple items
LIST_COLUMNS = {
    "Résumé des faits concis",
    "Liste des chefs et articles en infraction",
    "Nbr Chefs par articles",
    "Nbr Chefs par articles par période de radiation",
    "Liste des sanctions imposées",
    "Nombre de chefs par articles et total amendes",
    "Nombre de chefs par article ayant une réprimande",
    "Autres mesures ordonnées",
    "À vérifier",
}

# ---- Interest columns (only these are highlighted & filtered when checkbox is on)
INTEREST_COLS = {
    "Nbr Chefs par articles",
    "Nbr Chefs par articles par période de radiation",
    "Nombre de chefs par articles et total amendes",
    "Nombre de chefs par article ayant une réprimande",
}

# columns that must never be highlighted (explicit list, your #2)
NO_HILITE_COLS = {
    "Liste des chefs et articles en infraction",
    "Liste des sanctions imposées",
}

EMPTY_SPAN = "<span class='empty'>—</span>"

# ==========================================================
# ----------------  Text utils & rendering  ----------------
# ==========================================================

def _safe_str(x) -> str:
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return ""
    return str(x).strip()

def fmt_amount(x) -> str:
    """Format 0 -> 0 $, 5000 -> 5 000 $"""
    s = _safe_str(x)
    if s == "":
        return ""
    try:
        val = float(str(s).replace(" ", "").replace("\xa0", ""))
        if abs(val) < 0.005:
            return "0 $"
        ints = f"{int(round(val)):,.0f}".replace(",", " ").replace("\xa0"," ")
        return f"{ints} $"
    except Exception:
        return s

def split_items(text: str) -> list[str]:
    """Split into items, preserving simple punctuation."""
    if not text:
        return []
    t = text.replace("•", "\n").replace("\r", "\n")
    parts = re.split(r"\n|;|\u2022| - |\t", t)
    parts = [p.strip(" •\t") for p in parts if p and p.strip(" •\t")]
    return parts if parts else [text.strip()]

def highlight_inside(text: str, pattern: re.Pattern) -> str:
    """Return HTML with <span class='hit'> around matches."""
    if not text:
        return ""
    return pattern.sub(lambda m: f'<span class="hit">{m.group(0)}</span>', text)

def to_bullets(text: str, bulletize: bool) -> str:
    """Render in <ul><li>… if bulletize=True and multiple items."""
    if not text:
        return ""
    items = split_items(text)
    if not bulletize or len(items) == 1:
        return items[0]
    lis = "\n".join(f"<li>{p}</li>" for p in items)
    return f"<ul>{lis}</ul>"

def render_cell(value: str, column_name: str, only_segment: bool, pattern: re.Pattern) -> str:
    """
    HTML rendering of a cell with the rules requested.
    - Highlight only in INTEREST_COLS (never in NO_HILITE_COLS).
    - When only_segment=True, keep only the items that contain the article,
      but only for INTEREST_COLS (your checkbox rule).
    """
    raw = _safe_str(value)

    # Money formatting
    if column_name == "Total amendes":
        raw = fmt_amount(raw)

    # Segment filtering (checkbox) only for the 4 interest columns
    if only_segment and column_name in INTEREST_COLS:
        items = split_items(raw)
        items = [x for x in items if pattern.search(x)]
        raw = "\n".join(items)

    # Highlighting:
    if column_name in NO_HILITE_COLS:
        highlighted = raw                       # never highlight in the 2 excluded columns
    elif column_name in INTEREST_COLS:
        highlighted = highlight_inside(raw, pattern)   # highlight only in the 4 interest columns
    else:
        highlighted = raw                       # other columns: no highlight

    # Bullets?
    is_list_col = column_name in LIST_COLUMNS
    html = to_bullets(highlighted, bulletize=is_list_col)

    cls = "" if is_list_col else " no-bullets"
    display = html if html else EMPTY_SPAN
    return f'<div class="{cls.strip()}">{display}</div>'

def build_html_table(df: pd.DataFrame, article: str, only_segment: bool) -> str:
    token = re.escape(article.strip())
    pattern = re.compile(rf"(?<!\d){token}(?!\d)", flags=re.IGNORECASE)

    headers = list(df.columns)

    html = [CSS, '<div class="viewport"><table>']
    html.append("<thead><tr>")
    for h in headers:
        cls = WIDTH_CLASS.get(h, "col-m")
        html.append(f'<th class="{cls}">{h}</th>')
    html.append("</tr></thead>")

    html.append("<tbody>")
    for _, row in df.iterrows():
        html.append("<tr>")
        for h in headers:
            cls = WIDTH_CLASS.get(h, "col-m")
            cell = render_cell(row.get(h, ""), h, only_segment, pattern)
            html.append(f'<td class="{cls}">{cell}</td>')
        html.append("</tr>")
    html.append("</tbody></table></div>")
    return "\n".join(html)

# ==========================================================
# ----------------  Excel export with hits  ----------------
# ==========================================================

def write_excel_with_hits(df: pd.DataFrame, article: str, only_segment: bool) -> BytesIO:
    """
    Export Excel:
      - wrap text + top alignment
      - highlight matches inside the cell ONLY in INTEREST_COLS
      - never highlight in NO_HILITE_COLS
      - apply 'only segment' logic only for INTEREST_COLS
    """
    token = re.escape(article.strip())
    pattern = re.compile(rf"(?<!\d){token}(?!\d)", flags=re.IGNORECASE)

    out = df.copy()
    if "Total amendes" in out.columns:
        out["Total amendes"] = out["Total amendes"].map(fmt_amount)

    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as xw:
        out.to_excel(xw, index=False, startrow=1, sheet_name="Résultat")
        wb  = xw.book
        ws  = xw.sheets["Résultat"]

        # header row above the table
        ws.write(0, 0, f"Article filtré : {article}")

        # formats
        wrap_top = wb.add_format({"text_wrap": True, "valign": "top"})
        hit_fmt  = wb.add_format({"color": "#cc0000", "bold": True})

        # freeze header row
        ws.freeze_panes(2, 0)

        # Write back each interest cell with rich string (so we can color only matches)
        headers = list(out.columns)
        row_start = 2  # data starts here (because to_excel with startrow=1 + header row)

        for r in range(len(out)):
            for c, col in enumerate(headers):
                txt = _safe_str(out.iat[r, c])
                # apply 'segment only' for INTEREST_COLS
                if only_segment and col in INTEREST_COLS:
                    items = split_items(txt)
                    items = [x for x in items if pattern.search(x)]
                    txt = "\n".join(items)

                # For non-interest columns: just ensure wrap/top (except we will replace anyway with rich if matches)
                # We rewrite *every* cell so that wrap/top is guaranteed consistently
                if (col in NO_HILITE_COLS) or (col not in INTEREST_COLS):
                    ws.write(row_start + r, c, txt, wrap_top)
                    continue

                # Interest columns: highlight inside cell using rich string
                if not txt:
                    ws.write(row_start + r, c, "", wrap_top)
                    continue

                parts = pattern.split(txt)
                hits  = pattern.findall(txt)

                if hits:
                    # interleave normal text and highlighted matches
                    payload = []
                    for i, part in enumerate(parts):
                        if part:
                            payload.append(part)
                        if i < len(hits):
                            payload.append(hit_fmt)
                            payload.append(hits[i])
                    ws.write_rich_string(row_start + r, c, *payload, wrap_top)
                else:
                    ws.write(row_start + r, c, txt, wrap_top)

        # Auto width (bounded)
        for col_idx, col_name in enumerate(headers):
            src = out[col_name].astype(str)
            est = int(src.map(len).max() * 1.05) if len(src) else 12
            ws.set_column(col_idx, col_idx, max(12, min(60, est)))

    bio.seek(0)
    return bio

# ==========================================================
# ----------------------  Helpers  -------------------------
# ==========================================================

def read_uploaded_excel(fp) -> pd.DataFrame:
    """Read the uploaded Excel and return the dataframe you display.
       (This assumes your workbook already has the expected columns.)
    """
    # the user already confirmed everything works here with the current file,
    # so simply read the first sheet as before:
    return pd.read_excel(fp)

# ==========================================================
# ---------------------  Flask views  ----------------------
# ==========================================================

_latest_excel: Optional[BytesIO] = None
_latest_filename: str = "resultat.xlsx"

@app.route("/", methods=["GET"])
def home():
    return render_template("index.html", table_html=None, css=CSS)

@app.route("/analyze", methods=["POST"])
def analyze():
    global _latest_excel, _latest_filename

    article = request.form.get("article", "").strip()
    only_segment = request.form.get("only_segment") == "on"
    file = request.files.get("file")

    if not article:
        flash("Veuillez saisir un article.", "error")
        return redirect(url_for("home"))
    if not file or not file.filename:
        flash("Veuillez choisir un fichier Excel (.xlsx / .xlsm).", "error")
        return redirect(url_for("home"))

    try:
        df = read_uploaded_excel(file)
    except Exception as e:
        flash(f"Erreur de lecture Excel : {e}", "error")
        return redirect(url_for("home"))

    # Build HTML table
    table_html = build_html_table(df, article=article, only_segment=only_segment)

    # Build Excel
    try:
        bio = write_excel_with_hits(df, article=article, only_segment=only_segment)
    except Exception as e:
        flash(f"Erreur lors de la création Excel : {e}", "error")
        return render_template("index.html", table_html=table_html, css=CSS)

    _latest_excel = bio
    _latest_filename = f"Article_{article}.xlsx"

    return render_template("index.html", table_html=table_html, css=CSS)

@app.route("/download", methods=["GET"])
def download():
    if _latest_excel is None:
        flash("Aucun résultat à télécharger.", "error")
        return redirect(url_for("home"))
    _latest_excel.seek(0)
    return send_file(
        _latest_excel,
        as_attachment=True,
        download_name=_latest_filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8000, debug=False)
