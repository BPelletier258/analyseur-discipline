# main.py — Filtrage par article (largeurs + puces + couleur + Excel propre)
# Date : 2025-09-08

import os
import io
import re
import time
import unicodedata
from datetime import datetime
from typing import Dict, Optional, List

import pandas as pd
from flask import Flask, request, render_template_string, send_file

app = Flask(__name__)

# ──────────────────────────────────────────────────────────────────────────────
#  STYLES & GABARIT HTML
# ──────────────────────────────────────────────────────────────────────────────

STYLE = r"""
<style>
:root{
  --w-def: 22rem;        /* largeur par défaut */
  --w-2x:  44rem;        /* double largeur        */
  --w-num: 8rem;         /* colonnes numériques   */
}
body{font-family: system-ui,-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif; margin:24px;}
h1{font-size:20px;margin:0 0 12px;}
form{display:grid;gap:10px;margin:0 0 12px;}
input[type="text"],input[type="file"]{font-size:14px}
.hint{font-size:12px;color:#666}
.note{background:#fff6e5;border:1px solid #ffd89b;padding:8px 10px;border-radius:6px;margin:10px 0}
.download{margin:10px 0}
.kbd{font-family:ui-monospace,Menlo,Consolas,monospace;background:#f3f4f6;padding:2px 4px;border-radius:4px}

.viewport{height:60vh;overflow:auto;border:1px solid #ddd}
table{border-collapse:collapse;table-layout:fixed;width:100%;font-size:13px}
th,td{border:1px solid #ddd;padding:6px 8px;vertical-align:top}
th{background:#f3f4f6;text-align:center}

ul.bullets{margin:0;padding-left:1.15rem}
.bullets li{margin:0 0 .15rem}
.dash{color:#666}
.hl{color:#d00;font-weight:600}  /* surlignage article */

.msg{white-space:pre-wrap;font-family:ui-monospace,Menlo,Consolas,monospace;font-size:12px;margin-top:12px}
.ok{color:#065f46}.err{color:#7f1d1d}
</style>
"""

HTML = r"""
<!doctype html>
<meta charset="utf-8">
<title>Analyseur Discipline – Filtrage par article</title>
{{ style|safe }}
<h1>Analyseur Discipline – Filtrage par article</h1>

<div class="note">
  Règles : détection exacte de l’article. Si la 1<sup>re</sup> cellule contient
  « <span class="kbd">Article filtré :</span> », on ignore la 1<sup>re</sup> ligne (lignes d’en-tête sur la 2<sup>e</sup>).
</div>

<form method="POST" enctype="multipart/form-data">
  <label>Article à rechercher (ex. <span class="kbd">29</span>, <span class="kbd">59(2)</span>)</label>
  <input name="article" required value="{{ article or '' }}" placeholder="ex.: 29, 59(2)">
  <label><input type="checkbox" name="segments_only" {% if segments_only %}checked{% endif %}>
         Afficher uniquement le segment contenant l’article dans les 4 colonnes d’intérêt</label>
  <label>Fichier Excel</label>
  <input type="file" name="file" accept=".xlsx,.xlsm" required>
  <button>Analyser</button>
  <div class="hint">Formats : .xlsx / .xlsm</div>
</form>

{% if table_html %}
  <div class="download">
    <a href="{{ download_url }}">Télécharger le résultat (Excel)</a>
  </div>
  <div class="viewport">{{ table_html|safe }}</div>
{% endif %}

{% if message %}
  <div class="msg {{ 'ok' if ok else 'err' }}">{{ message }}</div>
{% endif %}
"""

# ──────────────────────────────────────────────────────────────────────────────
#  OUTILS
# ──────────────────────────────────────────────────────────────────────────────

def _norm(s: str) -> str:
    if not isinstance(s, str):
        s = "" if s is None else str(s)
    s = unicodedata.normalize("NFKD", s).encode("ascii","ignore").decode("ascii")
    s = s.replace("\u00A0"," ").strip().lower()
    s = " ".join(s.split())
    return s

def _prep_text(v) -> str:
    if not isinstance(v, str):
        v = "" if v is None else str(v)
    v = v.replace("•", " ").replace("·"," ").replace("◦"," ")
    v = v.replace("\u00A0"," ").replace("\u202F"," ")
    v = v.replace("\r\n","\n").replace("\r","\n")
    v = re.sub(r"[ \t]+"," ",v)
    return v.strip()

def build_article_pattern(token: str) -> re.Pattern:
    token = (token or "").strip()
    esc = re.escape(token)
    tail = r"(?![\d.])" if token and token[-1].isdigit() else r"\b"
    return re.compile(rf"(?:\b(?:art(?:icle)?\s*[: ]*)?)({esc}){tail}", re.I)

# colonnes d’intérêt + colonnes à puces + colonnes numériques
INTEREST_COLS = {
    "Nbr Chefs par articles",
    "Nbr Chefs par articles par période de radiation",
    "Nombre de chefs par articles et total amendes",
    "Nombre de chefs par article ayant une réprimande",
}
BULLET_COLS = {
    "Résumé des faits concis",
    "Liste des chefs et articles en infraction",
    "Nbr Chefs par articles",
    "Liste des sanctions imposées",
    "Nbr Chefs par articles par période de radiation",
    "Nombre de chefs par articles et total amendes",
    "Nombre de chefs par article ayant une réprimande",
    "Autres mesures ordonnées",
    "À vérifier",
}
NUMERIC_COLS = {"Total chefs","Total amendes","Total réprimandes"}

# règles de largeur
WIDTH_2X = {
    "Résumé des faits concis",
    "Autres mesures ordonnées",
    "Date de création",
    "Date de mise à jour",
    "Numéro de la décision",
    "Nombre de chefs par articles et total amendes",
    "Nbr Chefs par articles",
    "Nbr Chefs par articles par période de radiation",
    "Nombre de chefs par article ayant une réprimande",
}
# « Liste des chefs et articles en infraction » et « Liste des sanctions imposées »
# conservent la largeur par défaut (pas dans WIDTH_2X).

def first_row_is_banner(xls_stream) -> bool:
    preview = pd.read_excel(xls_stream, header=None, nrows=1, engine="openpyxl")
    xls_stream.seek(0)
    if preview.empty: 
        return False
    v = preview.iloc[0,0]
    return isinstance(v,str) and _norm(v).startswith(_norm("Article filtré :"))

def read_dataframe(xls_stream) -> pd.DataFrame:
    if first_row_is_banner(xls_stream):
        df = pd.read_excel(xls_stream, header=0, skiprows=1, engine="openpyxl")
    else:
        df = pd.read_excel(xls_stream, header=0, engine="openpyxl")
    return df

# ──────────────────────────────────────────────────────────────────────────────
#  MISE EN FORME HTML
# ──────────────────────────────────────────────────────────────────────────────

def highlight_html(text: str, pat: re.Pattern) -> str:
    """Entoure seulement le nombre recherché d'un span.hl (évite \1)."""
    if not text: 
        return ""
    def repl(m): return f"<span class='hl'>{m.group(1)}</span>"
    return pat.sub(repl, text)

def to_bullets_html(text: str) -> str:
    if not text:
        return "<span class='dash'>—</span>"
    # on découpe “|” (nos segments) ou nouvelles lignes / points-virgules
    parts = [t.strip(" •\t") for t in re.split(r"\|+|\n|;", text) if t.strip(" •\t")]
    if not parts:
        return "<span class='dash'>—</span>"
    lis = "".join(f"<li>{t}</li>" for t in parts)
    return f"<ul class='bullets'>{lis}</ul>"

def build_table_html(df: pd.DataFrame, pat: re.Pattern, segments_only: bool) -> str:
    # — Largeurs via <colgroup> calculées sur les en-têtes —
    cols: List[str] = list(df.columns)
    col_elems = []
    for c in cols:
        if c in WIDTH_2X:
            col_elems.append("<col style='width:var(--w-2x)'>")
        elif c in NUMERIC_COLS:
            col_elems.append("<col style='width:var(--w-num)'>")
        else:
            col_elems.append("<col style='width:var(--w-def)'>")
    colgroup = "<colgroup>" + "".join(col_elems) + "</colgroup>"

    # — THEAD —
    thead = "<thead><tr>" + "".join(f"<th>{c}</th>" for c in cols) + "</tr></thead>"

    # — TBODY —
    trs = []
    for _, row in df.iterrows():
        tds = []
        for c in cols:
            val = _prep_text(row.get(c, ""))
            # segments_only : ne s’applique qu’aux 4 colonnes d’intérêt
            if segments_only and c in INTEREST_COLS:
                # on isole uniquement les segments contenant l’article
                segs = [p.strip() for p in re.split(r"[;\n\|]", val) if p.strip()]
                segs = [s for s in segs if pat.search(s)]
                val = " | ".join(segs)

            # coloration dans 4 colonnes + « Liste des chefs et articles en infraction »
            if c in INTEREST_COLS or c == "Liste des chefs et articles en infraction":
                val = highlight_html(val, pat)

            # rendu à puces
            if c in BULLET_COLS:
                html = to_bullets_html(val)
            else:
                html = val if val else "<span class='dash'>—</span>"

            tds.append(f"<td>{html}</td>")
        trs.append("<tr>" + "".join(tds) + "</tr>")
    tbody = "<tbody>" + "".join(trs) + "</tbody>"

    return f"<table>{colgroup}{thead}{tbody}</table>"

# ──────────────────────────────────────────────────────────────────────────────
#  EXPORT EXCEL
# ──────────────────────────────────────────────────────────────────────────────

def export_excel(df: pd.DataFrame, article: str) -> str:
    """
    - Ligne 1 : 'Article filtré : <article>'
    - En-têtes figées (pane à A3)
    - Puces sous forme de lignes (• …) & wrap_text
    - '-' pour cellules vides
    - Largeurs auto (bornées)
    """
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Filtre"

    # Ligne 1 : bannière
    ws.cell(1,1, f"Article filtré : {article}")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(1,len(df.columns)))
    ws["A1"].font = Font(bold=True)

    # Ligne 2 : en-têtes
    for j, col in enumerate(df.columns, start=1):
        ws.cell(2,j, col)

    # Corps (avec puces et '-')
    def excel_text(v, colname):
        v = _prep_text(v)
        if not v:
            return "-"
        if colname in BULLET_COLS:
            parts = [t.strip(" •\t") for t in re.split(r"\|+|\n|;", v) if t.strip(" •\t")]
            return "• " + "\n• ".join(parts) if parts else "-"
        return v

    for i, (_, r) in enumerate(df.iterrows(), start=3):
        for j, col in enumerate(df.columns, start=1):
            ws.cell(i,j, excel_text(r.get(col,""), col))

    # Wrap & largeur
    for col_idx, col in enumerate(df.columns, start=1):
        letter = get_column_letter(col_idx)
        for cell in ws[letter]:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        # estimation simple : +2 chars
        max_len = max([len(str(ws.cell(row=r, column=col_idx).value or "")) for r in range(1, ws.max_row+1)])
        width = min(60, max(12, max_len*0.12*10))   # heuristique douce
        ws.column_dimensions[letter].width = width

    # Figer en-têtes (pane au dessous de la ligne 2)
    ws.freeze_panes = "A3"

    # Sauvegarde
    ts = int(time.time())
    out = f"/tmp/filtrage_{ts}.xlsx"
    wb.save(out)
    return f"/download?path={out}"

# ──────────────────────────────────────────────────────────────────────────────
#  ROUTES
# ──────────────────────────────────────────────────────────────────────────────

@app.route("/", methods=["GET","POST"])
def index():
    if request.method == "GET":
        return render_template_string(HTML, style=STYLE, table_html=None, article="", segments_only=False,
                                      message=None, ok=True)

    file = request.files.get("file")
    article = (request.form.get("article") or "").strip()
    segments_only = request.form.get("segments_only") is not None

    if not file or not article:
        return render_template_string(HTML, style=STYLE, table_html=None, article=article,
                                      segments_only=segments_only,
                                      message="Erreur : fichier et article requis.", ok=False)

    name = (file.filename or "").lower()
    if not (name.endswith(".xlsx") or name.endswith(".xlsm")):
        return render_template_string(HTML, style=STYLE, table_html=None, article=article,
                                      segments_only=segments_only,
                                      message="Format non pris en charge. Fournir un .xlsx ou .xlsm.", ok=False)

    try:
        df = read_dataframe(file.stream)

        # Filtrage : on garde les lignes où AU MOINS une colonne d’intérêt contient l’article
        pat = build_article_pattern(article)
        masks = []
        for c in INTEREST_COLS:
            if c in df.columns:
                masks.append(df[c].astype(str).apply(lambda v: bool(pat.search(_prep_text(v)))))
        if not masks:
            return render_template_string(
                HTML, style=STYLE, table_html=None, article=article, segments_only=segments_only,
                message=("Aucune des 4 colonnes d’intérêt n’a été trouvée dans le fichier.\n"
                         f"Colonnes présentes : {list(df.columns)}"), ok=False)

        mask_any = masks[0]
        for m in masks[1:]:
            mask_any = mask_any | m

        df = df[mask_any].copy()
        if df.empty:
            return render_template_string(HTML, style=STYLE, table_html=None, article=article,
                                          segments_only=segments_only,
                                          message=f"Aucune ligne ne contient l’article « {article} ».", ok=True)

        # HTML : puces + surlignage + largeurs
        table_html = build_table_html(df, pat, segments_only)

        # Excel : bannière + en-têtes figées + “-” + wrap
        download_url = export_excel(df if not segments_only else df, article)

        return render_template_string(HTML, style=STYLE, table_html=table_html, article=article,
                                      segments_only=segments_only,
                                      download_url=download_url,
                                      message=f"{len(df)} ligne(s) filtrée(s).", ok=True)

    except Exception as e:
        return render_template_string(HTML, style=STYLE, table_html=None, article=article,
                                      segments_only=segments_only,
                                      message=f"Erreur inattendue : {repr(e)}", ok=False)

@app.route("/download")
def download():
    path = request.args.get("path")
    if not path or not os.path.exists(path):
        return "Fichier introuvable.", 404
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
