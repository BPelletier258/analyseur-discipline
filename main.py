# -*- coding: utf-8 -*-
import re
import math
from io import BytesIO
from typing import Optional

import pandas as pd
from flask import Flask, render_template, request, send_file
from markupsafe import Markup  # Flask 3.x : Markup vient de markupsafe
# Assure-toi d'avoir `xlsxwriter` dans requirements.txt

app = Flask(__name__)

# =========================
# ----  CONFIG UI / CSS  ---
# =========================
CSS = r"""
<style>
:root{
  --w-s: 8.5rem;     /* étroit  */
  --w-m: 12rem;      /* moyen   */
  --w-l: 18rem;      /* large   */
  --w-xl: 26rem;     /* très large */
}
*{box-sizing:border-box}
body{font-family: ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif; color:#0f172a}

.wrap{max-width:1600px; width:98vw; margin:0 auto;}

/* Bandeau règles + formulaire pleine largeur */
.note{background:#fff8e6;border:1px solid #ffd48a;padding:12px 14px;border-radius:8px;margin:12px 0 16px}
.formbar{background:#f8fafc;border:1px solid #e5e7eb;border-radius:10px;padding:16px}
.formgrid{display:grid;grid-template-columns: 1fr auto auto;gap:12px;align-items:end}
.formcol{display:flex;flex-direction:column;gap:8px}
label{font-size:14px;color:#475569}
input[type="text"]{padding:8px 10px;border:1px solid #cbd5e1;border-radius:8px;font-size:14px}
input[type="file"]{font-size:14px}
button{background:#0ea5e9;color:white;border:none;padding:10px 16px;border-radius:10px;font-weight:600;cursor:pointer}
button:hover{background:#0284c7}

/* Table */
.viewport{height:60vh;overflow:auto;border:1px solid #e5e7eb;border-radius:10px}
table{width:100%; border-collapse:collapse; table-layout:fixed;}
th,td{border:1px solid #e5e7eb; padding:6px 8px; vertical-align:top;
      white-space:normal; word-break:normal; overflow-wrap:anywhere; hyphens:auto;}
th{position:sticky; top:0; background:#f1f5f9; z-index:1; font-weight:700; text-align:center}
ul{margin:0; padding-left:1.05rem}
li{margin:0.1rem 0}
.no-bullets ul{list-style:none; padding-left:0; margin:0}
.empty{color:#9CA3AF;}
.hit{color:#c00; font-weight:700}

/* Largeurs */
.col-s { width:var(--w-s);  min-width:var(--w-s)}
.col-m { width:var(--w-m);  min-width:var(--w-m)}
.col-l { width:var(--w-l);  min-width:var(--w-l)}
.col-xl{ width:var(--w-xl); min-width:var(--w-xl)}
</style>
"""

# =========================
# ----  PARAMÈTRES  -------
# =========================

# Colonnes rendues en listes à puces
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

# Les 4 colonnes d’intérêt (case à cocher et surlignage HTML/Excel)
INTEREST_COLS = [
    "Nbr Chefs par articles",
    "Nbr Chefs par articles par période de radiation",
    "Nombre de chefs par articles et total amendes",
    "Nombre de chefs par article ayant une réprimande",
]

# Colonnes où l’on NE met plus le rouge (même si ça matche)
NO_HTML_HILIGHT = {
    "Liste des chefs et articles en infraction",
    "Liste des sanctions imposées",
}

# Classes de largeur
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
    "Nombre de chefs par articles et total amendes": "col-l",
    "Total amendes": "col-m",
    "Total réprimandes": "col-s",
    "À vérifier": "col-l",
    "Date de création": "col-m",
    "Date de mise à jour": "col-m",

    # listes
    "Résumé des faits concis": "col-xl",
    "Liste des chefs et articles en infraction": "col-xl",
    "Liste des sanctions imposées": "col-l",
    "Nbr Chefs par articles": "col-l",
    "Nbr Chefs par articles par période de radiation": "col-l",
    "Autres mesures ordonnées": "col-l",
}
EMPTY_SPAN = "<span class='empty'>—</span>"
# ============== Utils ==============

def _safe_str(x) -> str:
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return ""
    return str(x).strip()

def fmt_amount(x) -> str:
    s = _safe_str(x)
    if s == "":
        return ""
    try:
        val = float(str(s).replace(" ", "").replace("\xa0","").replace("$","").replace(",","."))
        if abs(val) < 0.005:
            return "0 $"
        ints = f"{int(round(val)):,.0f}".replace(",", " ").replace("\xa0"," ")
        return f"{ints} $"
    except Exception:
        return s

def split_items(text: str) -> list[str]:
    if not text:
        return []
    t = text.replace("•", "\n").replace("\r", "\n")
    parts = re.split(r"\n|;|\u2022|- ", t)
    parts = [p.strip(" •\t") for p in parts if p and p.strip(" •\t")]
    return parts if parts else [text.strip()]

def to_bullets(text: str, bulletize: bool) -> str:
    if not text:
        return ""
    items = split_items(text)
    if not bulletize or len(items) == 1:
        return items[0]
    lis = "\n".join(f"<li>{p}</li>" for p in items)
    return f"<ul>{lis}</ul>"

def html_highlight(text: str, pattern: re.Pattern, column: str) -> str:
    """Surligne (HTML) sauf pour les colonnes explicitement exclues."""
    if not text:
        return ""
    if column in NO_HTML_HILIGHT:
        return text
    return pattern.sub(lambda m: f'<span class="hit">{m.group(0)}</span>', text)

def render_cell(
    value: str,
    column_name: str,
    bulletize: bool,
    show_only_segment: bool,
    pattern: re.Pattern
) -> str:
    """
    Rendu HTML d'une cellule, avec :
      - formatage des montants ('Total amendes'),
      - isolement éventuel du segment dans les colonnes d'intérêt,
      - surlignage de l'article recherché,
      - rendu en liste à puces pour les colonnes de type liste,
      - valeur vide affichée comme un tiret gris.

    Dépend de : _safe_str, fmt_amount, split_items, highlight, to_bullets,
                LIST_COLUMNS, INTEREST_COLS, EMPTY_SPAN.
    """
    # Normalisation de la valeur brute
    raw = _safe_str(value)

    # 1) Mise en forme des amendes
    if column_name == "Total amendes":
        raw = fmt_amount(raw)

    # 2) Option "n'afficher que le segment contenant l'article"
    if show_only_segment and column_name in INTEREST_COLS:
        items = split_items(raw)
        items = [x for x in items if pattern.search(x)]
        raw = "\n".join(items)

    # 3) Surlignage (après éventuel filtrage)
         raw = html_highlight(raw, pattern, column_name)

    # 4) Rendu : liste à puces seulement pour les colonnes prévues
    is_list_col = column_name in LIST_COLUMNS
    html = to_bullets(raw, bulletize=is_list_col)

    # 5) Classe CSS pour enlever les puces ailleurs
    cls = "" if is_list_col else " no-bullets"

    # 6) ✅ Pas de backslash dans l'expression d'un f-string
    display = html if html else EMPTY_SPAN
    return f'<div class="{cls.strip()}">{display}</div>'


def build_html_table(df: pd.DataFrame, article: str, show_only_segment: bool) -> str:
    token = re.escape(article.strip())
    pattern = re.compile(rf"(?<!\d){token}(?!\d)", flags=re.IGNORECASE)

    headers = list(df.columns)

    out = [CSS, '<div class="viewport"><table>']
    out.append("<thead><tr>")
    for h in headers:
        out.append(f'<th class="{WIDTH_CLASS.get(h, "col-m")}">{h}</th>')
    out.append("</tr></thead><tbody>")

    for _, row in df.iterrows():
        out.append("<tr>")
        for h in headers:
            cell = render_cell(
              value=row.get(h, ""),
              column_name=h,
              bulletize=(h in LIST_COLUMNS),
              show_only_segment=show_only_segment,
              pattern=pattern
            )

            out.append(f'<td class="{WIDTH_CLASS.get(h, "col-m")}">{cell}</td>')
        out.append("</tr>")

    out.append("</tbody></table></div>")
    return "\n".join(out)

def filter_rows_keep_if_any_interest_match(df: pd.DataFrame, article: str) -> pd.DataFrame:
    """Ne garde que les lignes où l’article apparaît dans AU MOINS UNE des 4 colonnes d’intérêt."""
    token = re.escape(article.strip())
    pattern = re.compile(rf"(?<!\d){token}(?!\d)", flags=re.IGNORECASE)

    def has_match(row) -> bool:
        for col in INTEREST_COLS:
            if col in row and pattern.search(_safe_str(row[col])):
                return True
        return False

    mask = df.apply(has_match, axis=1)
    return df[mask].reset_index(drop=True)

def export_excel(df: pd.DataFrame, article: str) -> BytesIO:
    """Excel :
       - Ligne 1: 'Article filtré : X'
       - Ligne 2: en-têtes (style)
       - Lignes suivantes: données
       - Wrap + alignement haut partout
       - Largeurs auto
       - L’article est surligné (en rouge) dans LES 4 colonnes d’intérêt (à l’intérieur des cellules)
    """
    token = re.escape(article.strip())
    pattern = re.compile(rf"(?<!\d){token}(?!\d)", flags=re.IGNORECASE)

    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as xw:
        wb  = xw.book
        ws  = wb.add_worksheet("Résultat")
        xw.sheets["Résultat"] = ws

        # Styles
        hdr_fmt = wb.add_format({
            "bold": True, "bg_color": "#e2e8f0", "align": "left", "valign": "vtop", "border": 1
        })
        top_wrap = wb.add_format({"text_wrap": True, "valign": "top", "border": 1})
        title_fmt = wb.add_format({"bold": True})

        # Ligne 1 : titre
        ws.write(0, 0, f"Article filtré : {article}", title_fmt)

        # Ligne 2 : en-têtes
        headers = list(df.columns)
        ws.write_row(1, 0, headers, hdr_fmt)

        # Écrit le DataFrame SANS en-têtes à partir de la ligne 3
        df.to_excel(xw, sheet_name="Résultat", startrow=2, startcol=0, index=False, header=False)

        # Wrap + alignement haut partout
        ncols = len(headers)
        ws.set_column(0, ncols-1, 22, top_wrap)  # largeur de base + format
        ws.freeze_panes(2, 0)  # fige la ligne d’en-têtes (après le titre)

        # Largeurs “auto” approximatives (max 60)
        for c, name in enumerate(headers):
            max_len = max(
                [len(str(name))] + [len(_safe_str(v)) for v in df.iloc[:, c].tolist()]
            )
            ws.set_column(c, c, min(60, max(12, int(max_len*1.05))), top_wrap)

        # Surlignage “dans la cellule” pour LES 4 colonnes d’intérêt uniquement
        red = wb.add_format({"font_color": "#c00000", "bold": True})
        for col_name in INTEREST_COLS:
            if col_name not in headers:
                continue
            col_idx = headers.index(col_name)
            # réécriture des cellules de cette colonne (pour injecter un format partiel)
            # Simplification : si la cellule contient le pattern, on réécrit la cellule en texte
            # en entourant l’article par un format "rich string"
            for row_idx in range(len(df)):
                txt = _safe_str(df.iat[row_idx, col_idx])
                if not txt:
                    continue
                m = list(pattern.finditer(txt))
                if not m:
                    continue
                # Construit une rich string: [avant, (rouge, match), milieu, (rouge, match2), ... , après]
                pieces = []
                last = 0
                for match in m:
                    if match.start() > last:
                        pieces.append(txt[last:match.start()])
                    pieces.append(red)
                    pieces.append(match.group(0))
                    last = match.end()
                if last < len(txt):
                    pieces.append(txt[last:])

                # Efface et réécrit la cellule avec le format wrap/haut
                ws.write_rich_string(row_idx + 2, col_idx, *pieces, top_wrap)

    bio.seek(0)
    return bio

# ============== Flask ==============

_last_excel: Optional[BytesIO] = None
_last_excel_name: str = "resultat.xlsx"

@app.route("/", methods=["GET"])
def home():
    return render_template("index.html", html_table="", css=Markup(CSS), article="", only_segment=False)

@app.route("/analyze", methods=["POST"])
def analyze():
    global _last_excel, _last_excel_name

    article = request.form.get("article", "").strip()
    only_segment = bool(request.form.get("only_segment"))

    if "file" not in request.files or article == "":
        return render_template("index.html", html_table="", css=Markup(CSS), article=article, only_segment=only_segment, error="Fichier et article requis.")

    file = request.files["file"]
    try:
        df = pd.read_excel(file)
    except Exception as e:
        return render_template("index.html", html_table="", css=Markup(CSS), article=article, only_segment=only_segment, error=f"Lecture Excel impossible : {e}")

    # Normalisations simples de colonnes $ (si présente)
    if "Total amendes" in df.columns:
        df["Total amendes"] = df["Total amendes"].map(fmt_amount)

    # 1/ On ne garde que les lignes où l’article est présent dans AU MOINS UNE des 4 colonnes d’intérêt
    df = filter_rows_keep_if_any_interest_match(df, article)

    # 2/ HTML
    html_table = build_html_table(df, article, only_segment)

    # 3/ Excel prêt à télécharger
    _last_excel = export_excel(df, article)
    _last_excel_name = f"resultat_{article}.xlsx"

    return render_template(
        "index.html",
        html_table=Markup(html_table),
        css=Markup(CSS),
        article=article,
        only_segment=only_segment,
        error=None
    )

@app.route("/download", methods=["GET"])
def download():
    if _last_excel is None:
        # rien à télécharger
        return home()
    return send_file(_last_excel, as_attachment=True, download_name=_last_excel_name, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8000, debug=False)
