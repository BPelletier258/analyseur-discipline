# -*- coding: utf-8 -*-
import re
import math
from io import BytesIO

import pandas as pd
from flask import Flask, request, render_template_string, send_file

# =====================================================================================
# ---------------------------------   CONFIG UI   -------------------------------------
# =====================================================================================

CSS = r"""
<style>
:root{
  --w-s: 8.5rem;     /* étroit  */
  --w-m: 12rem;      /* moyen   */
  --w-l: 18rem;      /* large   */
  --w-xl: 26rem;     /* très large */
}
*{box-sizing:border-box}
body{font-family: ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif; margin:16px}
.note{background:#fff8e6;border:1px solid #ffd48a;padding:8px 10px;border-radius:6px;margin:8px 0 14px}
form{display:grid;gap:10px;grid-auto-flow:row; margin-bottom:8px}
input[type="text"]{font-size:14px;padding:6px 8px}
input[type="file"]{font-size:13px;color:#666}
button{padding:6px 10px;font-size:14px}
.viewport{height:60vh;overflow:auto;border:1px solid #ddd}

table{width:100%; border-collapse:collapse; table-layout:fixed;}
th,td{
  border:1px solid #e5e7eb; padding:6px 8px; vertical-align:top;
  white-space:normal; word-break:normal; overflow-wrap:anywhere; hyphens:auto;
}
th{position:sticky; top:0; background:#f8fafc; z-index:1; font-weight:600; text-align:center;}

ul{margin:0; padding-left:1.05rem}
li{margin:0.1rem 0}
.no-bullets ul{list-style:none; padding-left:0; margin:0}
.empty{color:#9CA3AF;}  /* tiret gris */
.hit{color:#c00; font-weight:700}

/* Largeurs : appliquer .col-s / .col-m / .col-l / .col-xl à <th> et <td> */
.col-s { width:var(--w-s);  min-width:var(--w-s)}
.col-m { width:var(--w-m);  min-width:var(--w-m)}
.col-l { width:var(--w-l);  min-width:var(--w-l)}
.col-xl{ width:var(--w-xl); min-width:var(--w-xl)}
</style>
"""

# Classes de largeur par colonne (libre à toi d’ajuster)
WIDTH_CLASS = {
    # scalaires
    "Nom de l'intimé": "col-l",
    "Ordre professionnel": "col-l",
    "Numéro de la décision": "col-m",
    "Date de la décision rendue": "col-m",
    "Nature de la décision": "col-m",
    "Période des faits": "col-m",
    "Plaidoyer de culpabilité": "col-s",
    "Nbr Chefs par articles": "col-m",
    "Total chefs": "col-s",
    "Radiation max": "col-s",
    "Nombre de chefs par articles et total amendes": "col-l",
    "Total amendes": "col-m",
    "Nombre de chefs par article ayant une réprimande": "col-l",
    "Total réprimandes": "col-s",
    "À vérifier": "col-l",
    "Date de création": "col-m",
    "Date de mise à jour": "col-m",

    # “listes”
    "Résumé des faits concis": "col-xl",
    "Liste des chefs et articles en infraction": "col-xl",
    "Liste des sanctions imposées": "col-l",
    "Nbr Chefs par articles par période de radiation": "col-l",
    "Autres mesures ordonnées": "col-l",
}

# Colonnes rendues en liste à puces
LIST_COLUMNS = {
    "Résumé des faits concis",
    "Liste des chefs et articles en infraction",
    "Nbr Chefs par articles",
    "Nbr Chefs par articles par période de radiation",
    "Liste des sanctions imposées",
    "Nombre de chefs par article ayant une réprimande",
    "Autres mesures ordonnées",
    "À vérifier",
}

# Colonnes d’intérêt pour “segment uniquement”
INTEREST_COLS = [
    "Résumé des faits concis",
    "Liste des chefs et articles en infraction",
    "Nbr Chefs par articles",
    "Liste des sanctions imposées",
]

# Colonnes utilisées pour filtrer les lignes contenant l’article
FILTER_COLS = INTEREST_COLS

EMPTY_SPAN = "<span class='empty'>—</span>"

# =====================================================================================
# ---------------------------------  UTILITAIRES  -------------------------------------
# =====================================================================================

def _safe_str(x) -> str:
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return ""
    return str(x).strip()

def _esc(text: str) -> str:
    """Échappe <, >, & pour éviter toute interprétation HTML non voulue"""
    return (text.replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;"))

def fmt_amount(x) -> str:
    """Format 0 -> 0 $, 5000 -> 5 000 $"""
    s = _safe_str(x)
    if s == "":
        return ""
    try:
        val = float(str(s).replace(" ", "").replace("\xa0",""))
        if abs(val) < 0.005:
            return "0 $"
        ints = f"{int(round(val)):,.0f}".replace(",", " ").replace("\xa0"," ")
        return f"{ints} $"
    except Exception:
        return s

def highlight(html_escaped_text: str, pattern: re.Pattern) -> str:
    """Sur du texte DÉJÀ échappé, ajoute <span class='hit'>…</span> sur les matches."""
    if not html_escaped_text:
        return ""
    return pattern.sub(lambda m: f"<span class='hit'>{m.group(0)}</span>", html_escaped_text)

def split_items(text: str) -> list[str]:
    """Découpe léger en items; prend en compte retours ligne, puces, point-virgule, ' - '."""
    if not text:
        return []
    t = text.replace("•", "\n").replace("\r", "\n")
    parts = re.split(r"\n|;|\u2022| - ", t)
    parts = [p.strip(" •\t") for p in parts if p and p.strip(" •\t")]
    return parts if parts else [text.strip()]

def to_bullets(text: str, bulletize: bool) -> str:
    """Rend en <ul><li> si bulletize=True et qu'il y a plusieurs items."""
    if not text:
        return ""
    items = split_items(text)
    if not bulletize or len(items) == 1:
        return items[0]
    lis = "\n".join(f"<li>{i}</li>" for i in items)
    return f"<ul>{lis}</ul>"

def render_cell(value: str, column_name: str, show_only_segment: bool, pattern: re.Pattern) -> str:
    raw = _safe_str(value)

    # Excel et UI : formater les amendes
    if column_name == "Total amendes":
        raw = fmt_amount(raw)

    # Échapper avant surlignage (sécurité)
    raw_esc = _esc(raw)

    # Option "segment uniquement" pour les 4 colonnes d'intérêt
    if show_only_segment and column_name in INTEREST_COLS:
        items = split_items(raw_esc)
        items = [highlight(x, pattern) for x in items if pattern.search(x)]
        raw_esc = "\n".join(items)

    # Surlignage (même si segment_only = False)
    raw_esc = highlight(raw_esc, pattern)

    # Puces uniquement pour LIST_COLUMNS
    is_list_col = column_name in LIST_COLUMNS
    html = to_bullets(raw_esc, bulletize=is_list_col)

    # Classe pour supprimer les puces visuelles sur les scalaires
    cls = "" if is_list_col else " no-bullets"
    display = html if html else EMPTY_SPAN
    return '<div class="{}">{}</div>'.format(cls.strip(), display)

def build_html_table(df: pd.DataFrame, article: str, show_only_segment: bool) -> str:
    # Pattern strict : l’article ne doit pas être au milieu d’un nombre
    token = re.escape(article.strip())
    pattern = re.compile(rf"(?<!\d){token}(?!\d)", flags=re.IGNORECASE)

    headers = list(df.columns)

    out = [CSS, '<div class="viewport"><table>']
    # thead
    out.append("<thead><tr>")
    for h in headers:
        cls = WIDTH_CLASS.get(h, "col-m")
        out.append(f'<th class="{cls}">{_esc(h)}</th>')
    out.append("</tr></thead>")

    # tbody
    out.append("<tbody>")
    for _, row in df.iterrows():
        out.append("<tr>")
        for h in headers:
            cls = WIDTH_CLASS.get(h, "col-m")
            cell = render_cell(row.get(h, ""), h, show_only_segment=show_only_segment, pattern=pattern)
            out.append(f'<td class="{cls}">{cell}</td>')
        out.append("</tr>")
    out.append("</tbody></table></div>")
    return "\n".join(out)

# =====================================================================================
# ---------------------------------  EXCEL EXPORT  ------------------------------------
# =====================================================================================

def dataframe_to_excel(df: pd.DataFrame, article: str) -> BytesIO:
    bio = BytesIO()
    out = df.copy()
    if "Total amendes" in out.columns:
        out["Total amendes"] = out["Total amendes"].map(fmt_amount)

    with pd.ExcelWriter(bio, engine="xlsxwriter") as xw:
        out.to_excel(xw, index=False, startrow=1, sheet_name="Résultat")
        ws = xw.sheets["Résultat"]
        ws.write(0, 0, f"Article filtré : {article}")
        ws.freeze_panes(2, 0)  # fige la ligne d’en-têtes (ligne 2 réelle)

        # largeur auto simple
        for col_idx, col_name in enumerate(out.columns):
            width = max(12, min(60, int(out[col_name].astype(str).map(len).max() * 1.1)))
            ws.set_column(col_idx, col_idx, width)

    bio.seek(0)
    return bio

# =====================================================================================
# ---------------------------------   FLASK APP   -------------------------------------
# =====================================================================================

app = Flask(__name__)
_last_excel: BytesIO | None = None
_last_table: str = ""

TEMPLATE = """
<!doctype html>
<meta charset="utf-8">
<title>Analyseur Discipline – Filtrage par article</title>
{{ css|safe }}
<div class="note">
  Règles : détection exacte de l’article. Si la 1<sup>re</sup> cellule contient
  « <b>Article filtré :</b> », on ignore la 1<sup>re</sup> ligne (lignes d’en-têtes sur la 2<sup>e</sup>).
</div>

<form method="post" enctype="multipart/form-data">
  <div>
    <label>Article à rechercher (ex. <code>29, 59(2)</code>)</label><br>
    <input type="text" name="article" value="{{ article }}" style="width:140px">
  </div>

  <label>
    <input type="checkbox" name="only" {% if only %}checked{% endif %}>
    Afficher uniquement le segment contenant l’article dans les 4 colonnes d’intérêt
  </label>

  <div>
    <label>Fichier Excel</label><br>
    <input type="file" name="file">
  </div>

  <button type="submit">Analyser</button>
</form>

<p>Formats : .xlsx / .xlsm</p>

{% if table %}
  <p><a href="/download">Télécharger le résultat (Excel)</a></p>
  {{ table | safe }}
{% endif %}
"""

def _clean_first_header_row(df: pd.DataFrame) -> pd.DataFrame:
    """Si la 1ʳᵉ ligne contient 'Article filtré :', on l’ignore (ligne descriptive)."""
    if not df.empty:
        first_cell = str(df.iloc[0, 0]).strip().lower()
        if first_cell.startswith("article filtré"):
            return df.iloc[1:].reset_index(drop=True)
    return df

def _filter_on_article(df: pd.DataFrame, article: str) -> pd.DataFrame:
    if not article.strip():
        return df.copy()

    token = re.escape(article.strip())
    pat = rf"(?<!\d){token}(?!\d)"  # même logique que pour l’affichage

    present = [c for c in FILTER_COLS if c in df.columns]
    if not present:
        return df.copy()

    mask = pd.Series(False, index=df.index)
    for c in present:
        col = df[c].astype(str).fillna("")
        mask = mask | col.str.contains(pat, case=False, regex=True)

    return df[mask].copy()

@app.route("/", methods=["GET", "POST"])
def index():
    global _last_excel, _last_table

    table = ""
    article = (request.form.get("article") or "29").strip()
    only = bool(request.form.get("only"))

    if request.method == "POST" and "file" in request.files and request.files["file"].filename:
        f = request.files["file"]
        # Lis l’Excel
        df = pd.read_excel(f, engine="openpyxl")
        df = _clean_first_header_row(df)
        df = _filter_on_article(df, article)

        # Affichage HTML + Excel
        table = build_html_table(df, article=article, show_only_segment=only)
        _last_table = table
        _last_excel = dataframe_to_excel(df, article)

    # GET ou POST sans fichier : si on a déjà un tableau, on le remonte
    if not table and _last_table:
        table = _last_table

    return render_template_string(TEMPLATE, css=CSS, table=table, article=article, only=only)

@app.route("/download")
def download():
    if not _last_excel:
        return "Rien à télécharger", 400
    _last_excel.seek(0)
    return send_file(_last_excel,
                     as_attachment=True,
                     download_name="resultat.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
