# -*- coding: utf-8 -*-
import io
import re
import os
import math
import secrets
from typing import Optional, Dict

import pandas as pd
from flask import (
    Flask, render_template, request, send_file, redirect, url_for, session
)

# =========================================================
#  Flask
# =========================================================
app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", secrets.token_hex(16))
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # 20 MB

# Stockage (très simple) en mémoire des exports Excel par session
# (clé = token placé en session ; valeur = bytes du fichier)
EXCEL_STORE: Dict[str, bytes] = {}

# =========================================================
#  Paramètres d'affichage / rendu
# =========================================================

CSS = r"""
<style>
:root{
  --w-s: 8.5rem;     /* étroit  */
  --w-m: 12rem;      /* moyen   */
  --w-l: 18rem;      /* large   */
  --w-xl: 26rem;     /* très large */
}
*{box-sizing:border-box}
body{font-family: ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif;}
.note{background:#fff8e6;border:1px solid #ffd48a;padding:8px 10px;border-radius:8px;margin:8px 0 14px}
.viewport{height:60vh;overflow:auto;border:1px solid #ddd}

table{width:100%; border-collapse:collapse; table-layout:fixed;}
th,td{border:1px solid #e5e7eb; padding:6px 8px; vertical-align:top;
      white-space:normal; word-break:normal; overflow-wrap:anywhere; hyphens:auto; }

th{position:sticky; top:0; background:#f8fafc; z-index:1; font-weight:600; text-align:center;}

ul{margin:0; padding-left:1.05rem}
li{margin:0.1rem 0}
.no-bullets ul{list-style:none; padding-left:0; margin:0}
.empty{color:#9CA3AF;}  /* tiret gris */
.hit{color:#c00; font-weight:700}

/* Largeurs */
.col-s { width:var(--w-s);  min-width:var(--w-s)}
.col-m { width:var(--w-m);  min-width:var(--w-m)}
.col-l { width:var(--w-l);  min-width:var(--w-l)}
.col-xl{ width:var(--w-xl); min-width:var(--w-xl)}
</style>
"""

# Mapping des classes de largeur par intitulé d’en-tête
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

    # colonnes “listes”
    "Résumé des faits concis": "col-xl",
    "Liste des chefs et articles en infraction": "col-xl",
    "Liste des sanctions imposées": "col-l",
    "Nbr Chefs par articles par période de radiation": "col-l",
    "Autres mesures ordonnées": "col-l",
}

# Colonnes qui doivent s’afficher en listes à puces
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

# Colonnes d’intérêt pour l’option “segment uniquement”
INTEREST_COLS = [
    "Résumé des faits concis",
    "Liste des chefs et articles en infraction",
    "Nbr Chefs par articles",
    "Liste des sanctions imposées",
]

EMPTY_SPAN = "<span class='empty'>—</span>"

# =========================================================
#  Utilitaires
# =========================================================
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
        val = float(str(s).replace(" ", "").replace("\xa0",""))
        if abs(val) < 0.005:
            return "0 $"
        ints = f"{int(round(val)):,.0f}".replace(",", " ").replace("\xa0"," ")
        return f"{ints} $"
    except Exception:
        return s

def highlight(text: str, pattern: re.Pattern) -> str:
    if not text:
        return ""
    return pattern.sub(lambda m: f'<span class="hit">{m.group(0)}</span>', text)

def split_items(text: str) -> list[str]:
    """Découpe léger en items (retours, puces, points-virgules, etc.)."""
    if not text:
        return []
    t = text.replace("•", "\n").replace("\r", "\n")
    parts = re.split(r"\n|;|\u2022|- ", t)
    parts = [p.strip(" •\t") for p in parts if p and p.strip(" •\t")]
    return parts if parts else [text.strip()]

def to_bullets(text: str, bulletize: bool) -> str:
    """Rend en <ul><li> si bulletize=True et qu'il y a plusieurs items ; sinon texte brut."""
    if not text:
        return ""
    items = split_items(text)
    if not bulletize or len(items) == 1:
        return items[0]
    lis = "\n".join(f"<li>{p}</li>" for p in items)
    return f"<ul>{lis}</ul>"

def render_cell(value: str, column_name: str, show_only_segment: bool, pattern: re.Pattern) -> str:
    raw = _safe_str(value)

    # Mise en forme des amendes
    if column_name == "Total amendes":
        raw = fmt_amount(raw)

    # Option : ne garder que les segments contenant l'article dans 4 colonnes d'intérêt
    if show_only_segment and column_name in INTEREST_COLS:
        items = split_items(raw)
        items = [highlight(x, pattern) for x in items if pattern.search(x)]
        raw = "\n".join(items)

    # Surlignage (même si segment_only = False)
    raw = highlight(raw, pattern)

    is_list_col = column_name in LIST_COLUMNS
    html = to_bullets(raw, bulletize=is_list_col)

    cls = "" if is_list_col else " no-bullets"
    return f'<div class="{cls.strip()}">{html or EMPTY_SPAN}</div>'

def build_html_table(df: pd.DataFrame, article: str, show_only_segment: bool) -> str:
    # Prépare le pattern (exact, on échappe l’entrée, et on évite de capturer des chiffres adjacents)
    token = re.escape(article.strip())
    pattern = re.compile(rf"(?<!\d){token}(?!\d)", flags=re.IGNORECASE)

    headers = list(df.columns)

    html = [CSS, '<div class="viewport"><table>']
    # thead
    html.append("<thead><tr>")
    for h in headers:
        cls = WIDTH_CLASS.get(h, "col-m")
        html.append(f'<th class="{cls}">{h}</th>')
    html.append("</tr></thead>")

    # tbody
    html.append("<tbody>")
    for _, row in df.iterrows():
        html.append("<tr>")
        for h in headers:
            cls = WIDTH_CLASS.get(h, "col-m")
            cell = render_cell(row.get(h, ""), h, show_only_segment=show_only_segment, pattern=pattern)
            html.append(f'<td class="{cls}">{cell}</td>')
        html.append("</tr>")
    html.append("</tbody></table></div>")
    return "\n".join(html)

# =========================================================
#  Lecture / filtrage Excel
# =========================================================
def read_first_sheet_detect_header(b: bytes) -> pd.DataFrame:
    """Lit la première feuille. Si A1 commence par 'Article filtré :', on prend l’en-tête sur la 2e ligne."""
    bio = io.BytesIO(b)
    xls = pd.ExcelFile(bio, engine="openpyxl")
    # on inspecte brut
    tmp = pd.read_excel(xls, sheet_name=0, header=None)
    header_row = 1 if tmp.iloc[0, 0] and str(tmp.iloc[0, 0]).strip().lower().startswith("article filtré") else 0
    df = pd.read_excel(xls, sheet_name=0, header=header_row)
    # On enlève les colonnes totalement vides
    df = df.dropna(axis=1, how="all")
    return df

def filter_by_article(df: pd.DataFrame, article: str) -> pd.DataFrame:
    token = re.escape(article.strip())
    pattern = re.compile(rf"(?<!\d){token}(?!\d)", flags=re.IGNORECASE)

    # on regarde d'abord les colonnes d'intérêt ; si absentes on tombera sur toutes les colonnes texte
    cols = [c for c in INTEREST_COLS if c in df.columns]
    if not cols:
        cols = list(df.columns)

    def row_match(s):
        return any(pattern.search(_safe_str(s.get(c, ""))) for c in cols)

    mask = df.apply(row_match, axis=1)
    return df.loc[mask].reset_index(drop=True)

# =========================================================
#  Export Excel formaté
# =========================================================
def produce_excel_bytes(df: pd.DataFrame, article: str) -> bytes:
    """Construit un xlsx : ligne 1 = 'Article filtré : X', entêtes figées, wrap + valign top."""
    out = df.copy()
    if "Total amendes" in out.columns:
        out["Total amendes"] = out["Total amendes"].map(fmt_amount)

    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as xw:
        out.to_excel(xw, index=False, startrow=1, sheet_name="Résultat")
        wb  = xw.book
        ws  = xw.sheets["Résultat"]

        # Ligne titre + gel des volets
        ws.write(0, 0, f"Article filtré : {article}")
        ws.freeze_panes(2, 0)

        # Format par défaut : wrap + valign top
        wrap_top = wb.add_format({"text_wrap": True, "valign": "top"})
        ws.set_default_row(hide_unused_rows=False)  # nécessaire sur certaines versions
        # Largeurs auto (bornées) + format wrap_top
        for col_idx, col_name in enumerate(out.columns):
            # Longueur max de la colonne (en caractères) * facteur
            maxlen = max(12, int(out[col_name].astype(str).map(len).max() * 1.1))
            width = max(12, min(60, maxlen))
            ws.set_column(col_idx, col_idx, width, wrap_top)

        # Appliquer aussi wrap/valign aux lignes de données (optionnel mais robuste)
        for r in range(2, 2 + len(out)):  # 2 = ligne d’entêtes réelle (0-based) + 1
            ws.set_row(r, None, wrap_top)

    bio.seek(0)
    return bio.getvalue()

# =========================================================
#  Routes
# =========================================================
@app.route("/", methods=["GET", "POST"])
def index():
    ctx = dict(error=None, table_html=None, article=None, segment_only=False, token=None)

    if request.method == "POST":
        try:
            article = (request.form.get("article") or "").strip()
            if not article:
                ctx["error"] = "Veuillez saisir un article (ex. 29, 59(2))."
                return render_template("index.html", **ctx)

            segment_only = bool(request.form.get("segment_only"))
            f = request.files.get("file")
            if not f or not f.filename:
                ctx["error"] = "Veuillez joindre un fichier Excel (.xlsx / .xlsm)."
                return render_template("index.html", **ctx)

            raw = f.read()
            df_all = read_first_sheet_detect_header(raw)
            df = filter_by_article(df_all, article)

            if df.empty:
                ctx.update(error=f"Aucune ligne ne contient l’article « {article} » dans les colonnes cibles.",
                           article=article, segment_only=segment_only)
                return render_template("index.html", **ctx)

            # Rendu HTML
            html = build_html_table(df, article=article, show_only_segment=segment_only)
            # Export Excel mémorisé
            xls_bytes = produce_excel_bytes(df, article)
            token = secrets.token_urlsafe(16)
            EXCEL_STORE[token] = xls_bytes
            session["dl_token"] = token

            ctx.update(table_html=html, article=article, segment_only=segment_only, token=token)
            return render_template("index.html", **ctx)

        except Exception as e:
            ctx["error"] = f"Échec de l’analyse : {e}"
            return render_template("index.html", **ctx)

    # GET
    token = session.get("dl_token")
    if token and token in EXCEL_STORE:
        ctx["token"] = token
    return render_template("index.html", **ctx)

@app.route("/download")
def download():
    token = session.get("dl_token")
    if not token or token not in EXCEL_STORE:
        return redirect(url_for("index"))
    data = EXCEL_STORE[token]
    # Option : nettoyer après téléchargement
    # EXCEL_STORE.pop(token, None)
    return send_file(
        io.BytesIO(data),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name="resultat.xlsx",
    )

# =========================================================
#  Lancement local
# =========================================================
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
