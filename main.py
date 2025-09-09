import os
import io
import re
import math
import time
import unicodedata
from datetime import datetime
from typing import List

import pandas as pd
from flask import Flask, request, render_template_string, send_file

app = Flask(__name__)

# --------- Helpers de normalisation ---------
def _norm(s: str) -> str:
    if not isinstance(s, str):
        s = "" if s is None else str(s)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = s.replace("\u00A0", " ").replace("\u202F", " ")
    return " ".join(s.strip().lower().split())

# --------- Mise en forme argent ---------
NARROW_NBSP = "\u202F"  # espace fine insécable
def fmt_money(v) -> str:
    """500 -> '500 $', 5000 -> '5 000 $', NaN/None -> '—'."""
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return "—"
    try:
        n = float(v)
        # Round propre pour éviter 4999.9999
        n = int(round(n))
        # groupement avec espace fine insécable
        s = f"{n:,}".replace(",", NARROW_NBSP)
        return f"{s} $"
    except Exception:
        # si déjà texte, on renvoie tel quel
        s = str(v).strip()
        return s if s else "—"

# --------- CSS / HTML ---------
STYLE = """
<style>
:root{
  --w-def: 22rem;   /* largeur par défaut */
  --w-2x:  44rem;   /* largeur x2 demandée */
  --w-num: 8rem;    /* colonnes numériques étroites */
}
body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;margin:24px}
h1{font-size:20px;margin:0 0 12px}
form{display:grid;gap:10px;margin:8px 0 14px}
input[type="text"],input[type="file"]{font-size:14px}
input[type="text"]{padding:8px}
button{padding:8px 12px;font-size:14px;cursor:pointer}
.hint{font-size:12px;color:#666}
.note{background:#fff6e5;border:1px solid #ffd89b;padding:8px 10px;border-radius:6px;margin:10px 0}
.download{margin:12px 0}
.kbd{font-family:ui-monospace,Menlo,Consolas,monospace;background:#f3f4f6;padding:2px 4px;border-radius:4px}
.table-viewport{height:60vh;overflow:auto;border:1px solid #ddd}
.table-viewport table{table-layout:fixed;width:100%}
th,td{border:1px solid #ddd;padding:6px 8px;vertical-align:top;font-size:13px}
th{background:#f3f4f6;text-align:center}
.hl{color:#d00;font-weight:600}
.msg{margin-top:10px;white-space:pre-wrap;font:12px ui-monospace,Menlo,monospace}
.ok{color:#065f46}.err{color:#7f1d1d}
</style>
"""

TPL = """
<!doctype html><html><head><meta charset="utf-8"><title>Analyseur Discipline – Filtrage par article</title>{{style|safe}}</head>
<body>
<h1>Analyseur Discipline – Filtrage par article</h1>

<div class="note">Règles : détection exacte de l’article. Si la 1<sup>re</sup> cellule contient « <span class="kbd">Article filtré :</span> », on ignore la 1<sup>re</sup> ligne (lignes d’en-tête sur la 2<sup>e</sup>).</div>

<form method="POST" enctype="multipart/form-data">
  <label>Article à rechercher (ex. <span class="kbd">29</span>, <span class="kbd">59(2)</span>)</label>
  <input type="text" name="article" value="{{article or ''}}" required>
  <label>Fichier Excel</label>
  <input type="file" name="file" accept=".xlsx,.xlsm" required>
  <button type="submit">Analyser</button>
  <div class="hint">Formats : .xlsx / .xlsm</div>
</form>

{% if table %}
  <div class="download"><a href="{{download_url}}">Télécharger le résultat (Excel)</a></div>
  <div class="table-viewport">{{table|safe}}</div>
{% endif %}

{% if msg %}<div class="msg {{'ok' if ok else 'err'}}">{{msg}}</div>{% endif %}
</body></html>
"""

# --------- Largeurs de colonnes (colgroup) ---------
WIDE_2X = {
    _norm("Résumé des faits concis"),
    _norm("Autres mesures ordonnées"),
    _norm("Date de création"),
    _norm("Date de mise à jour"),
    _norm("Numéro de la décision"),
    _norm("Nbr Chefs par articles"),
    _norm("Nbr Chefs par articles par période de radiation"),
    _norm("Nombre de chefs par article ayant une réprimande"),
    _norm("Nombre de chefs par articles et total amendes"),
}
# “On laisse tel quel” = on ne force pas à 2x (mais on garde la largeur par défaut)
KEEP_DEFAULT = {
    _norm("Liste des chefs et articles en infraction"),
    _norm("Liste des sanctions imposées"),
}

def build_colgroup(cols: List[str]) -> str:
    """Construit un <colgroup> avec largeurs stables, basées sur les titres."""
    col_styles = []
    for c in cols:
        nc = _norm(c)
        if nc in WIDE_2X:
            width = "var(--w-2x)"
        elif re.search(r"^(total|nbr|nombre|radiation|max|chefs|amendes?)", nc):
            # colonnes “numériques / courtes”
            width = "var(--w-num)"
        else:
            width = "var(--w-def)"
        col_styles.append(f'<col style="width:{width}">')
    return "<colgroup>" + "".join(col_styles) + "</colgroup>"

# --------- Lecture Excel avec bannière “Article filtré :” ---------
def read_excel(file_stream) -> pd.DataFrame:
    peek = pd.read_excel(file_stream, header=None, nrows=1, engine="openpyxl")
    file_stream.seek(0)
    banner = False
    if not peek.empty:
        first = str(peek.iloc[0,0])
        if _norm(first).startswith(_norm("Article filtré :")):
            banner = True
    if banner:
        df = pd.read_excel(file_stream, skiprows=1, header=0, engine="openpyxl")
    else:
        df = pd.read_excel(file_stream, header=0, engine="openpyxl")
    return df

# --------- Mise en forme DataFrame avant HTML & Excel ---------
def tidy_dataframe(df: pd.DataFrame, money_col_name="Total amendes") -> pd.DataFrame:
    """- Replace NaN -> '—'
       - Formate 'Total amendes'
    """
    df = df.copy()

    # 1) formater l’argent si la colonne existe
    for c in df.columns:
        if _norm(c) == _norm(money_col_name):
            df[c] = df[c].apply(fmt_money)

    # 2) remplacer NaN partout ailleurs par '—'
    df = df.where(pd.notna(df), "—")

    return df

# --------- Export Excel ---------
def to_excel_download(df: pd.DataFrame, article: str) -> str:
    path = f"/tmp/filtrage_{int(time.time())}.xlsx"
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        # on écrit à partir de la ligne 2 (startrow=1) pour garder la ligne 1 pour le bandeau
        df.to_excel(xw, index=False, sheet_name="Filtre", startrow=1)
        ws = xw.book["Filtre"]

        # Ligne 1 : "Article filtré : <article>"
        ws.cell(row=1, column=1, value=f"Article filtré : {article}")

        # Figer la première ligne d’en-têtes (en dessous du bandeau)
        ws.freeze_panes = "A3"  # la 3e ligne est la 1re ligne de données, donc bandeau + en-têtes figés

        # Largeurs auto (simples)
        from openpyxl.utils import get_column_letter
        for j, col in enumerate(df.columns, start=1):
            # maxi 60, mini 12
            max_len = max([len(str(col))] + [len(str(v)) for v in df[col].astype(str).tolist()]) + 2
            width = max(12, min(60, max_len))
            ws.column_dimensions[get_column_letter(j)].width = width

    return f"/download?path={path}"

# --------- Routes ---------
@app.route("/", methods=["GET","POST"])
def index():
    if request.method == "GET":
        return render_template_string(TPL, style=STYLE, table=None, article="", msg=None, ok=True)

    # POST
    article = (request.form.get("article") or "").strip()
    f = request.files.get("file")
    if not f or not article:
        return render_template_string(TPL, style=STYLE, table=None, article=article,
                                      msg="Fichier et article requis.", ok=False)

    # formats acceptés
    n = (f.filename or "").lower()
    if not (n.endswith(".xlsx") or n.endswith(".xlsm")):
        return render_template_string(TPL, style=STYLE, table=None, article=article,
                                      msg="Formats acceptés : .xlsx / .xlsm", ok=False)

    try:
        df = read_excel(f.stream)

        # Préparation affichage : NaN -> '—' / Total amendes formaté
        df_disp = tidy_dataframe(df)

        # HTML : to_html puis injection d’un colgroup pour forcer les largeurs
        html = df_disp.to_html(index=False, escape=False)
        colgroup = build_colgroup(list(df_disp.columns))
        # injecter juste après l'ouverture de la table
        html = html.replace("<table border=\"1\" class=\"dataframe\">",
                            f"<table border=\"1\" class=\"dataframe\">{colgroup}")

        download_url = to_excel_download(df_disp, article)

        return render_template_string(TPL, style=STYLE, table=html, article=article,
                                      download_url=download_url,
                                      msg="Aperçu généré. (Le lien Excel est disponible ci-dessus.)", ok=True)
    except Exception as e:
        return render_template_string(TPL, style=STYLE, table=None, article=article,
                                      msg=f"Erreur : {e}", ok=False)

@app.route("/download")
def download():
    path = request.args.get("path")
    if not path or not os.path.exists(path):
        return "Fichier introuvable.", 404
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
