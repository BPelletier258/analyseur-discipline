# -*- coding: utf-8 -*-
import os, re, time, unicodedata, math
from html import escape
from typing import Dict, Optional, Set, List

import pandas as pd
from flask import Flask, request, render_template_string, send_file
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

app = Flask(__name__)

# ──────────────────────────────────────────────────────────────────────────────
# Style & page
# ──────────────────────────────────────────────────────────────────────────────

STYLE = """
<style>
  :root{
    /* largeur par défaut */
    --w-def: 22rem;
    /* largeur ×2 demandée */
    --w-2x: 44rem;
    /* colonnes numériques étroites */
    --w-num: 8rem;
  }
  body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;margin:22px}
  h1{font-size:20px;margin:0 0 10px}
  form{display:grid;gap:10px;margin:8px 0 12px}
  input[type="text"],input[type="file"]{font-size:14px}
  .note{background:#fff8e6;border:1px solid #ffd48a;padding:8px 10px;border-radius:6px}
  .hint{font-size:12px;color:#666;margin-top:6px}
  .download{margin:12px 0}
  .kbd{font-family:ui-monospace,Menlo,Consolas,monospace;background:#f3f4f6;padding:1px 4px;border-radius:4px}
  .viewport{height:60vh;overflow:auto;border:1px solid #ddd}
  table{border-collapse:collapse;width:100%;table-layout:fixed}
  th,td{border:1px solid #ddd;padding:6px 8px;vertical-align:top}
  th{background:#f3f4f6;text-align:center}
  td{word-break:break-word}
  /* rendu listes */
  ul.bullets{list-style:disc;margin:0;padding-left:1.15rem}
  ul.bullets li{margin:0.15rem 0}
  /* surlignage article */
  .hl{color:#d00000;font-weight:600}
  /* colonnes “num” un peu plus étroites */
  td.num, th.num{width:var(--w-num); text-align:center}
  /* tu peux fixer des largeurs ciblées ci-dessous si besoin :
     exemple: la 6e colonne (Résumé) plus large
  */
  /* table.dataframe tr > *:nth-child(6){width:var(--w-2x)} */
</style>
"""

HTML = """
<!doctype html>
<meta charset="utf-8">
<title>Analyseur Discipline – Filtrage par article</title>
{{ style|safe }}
<h1>Analyseur Discipline – Filtrage par article</h1>

<p class="note">
  Règles : détection exacte de l’article. Si la 1<sup>re</sup> cellule contient
  « <span class="kbd">Article filtré :</span> », on ignore la 1<sup>re</sup> ligne (lignes d’en-tête sur la 2<sup>e</sup>).
</p>

<form method="post" enctype="multipart/form-data">
  <label>Article à rechercher (ex. <span class="kbd">29</span>, <span class="kbd">59(2)</span>)</label>
  <input type="text" name="article" value="{{ article or '' }}" required placeholder="ex.: 29, 59(2)">

  <label>
    <input type="checkbox" name="segments_only" value="1" {% if segments_only %}checked{% endif %}>
    Afficher uniquement le segment contenant l’article dans les 4 colonnes d’intérêt
  </label>

  <label>Fichier Excel</label>
  <input type="file" name="file" accept=".xlsx,.xlsm" required>

  <button>Analyser</button>
  <div class="hint">Formats : .xlsx / .xlsm</div>
</form>

{% if table_html %}
  <div class="download"><a href="{{ dl }}">Télécharger le résultat (Excel)</a></div>
  <div class="viewport">{{ table_html|safe }}</div>
{% endif %}

{% if msg %}
  <pre class="hint">{{ msg }}</pre>
{% endif %}
"""

# ──────────────────────────────────────────────────────────────────────────────
# Normalisation & entêtes
# ──────────────────────────────────────────────────────────────────────────────

def _norm(s: str) -> str:
    if not isinstance(s, str):
        s = "" if s is None else str(s)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = s.replace("\u00A0", " ")
    return " ".join(s.strip().lower().split())

HEADER_ALIASES: Dict[str, Set[str]] = {
  # 4 colonnes d’intérêt
  "articles_enfreints": {
      _norm("Nbr Chefs par articles"),
      _norm("Articles enfreints"),
      _norm("Articles en infraction"),
      _norm("Liste des chefs et articles en infraction"),
  },
  "duree_totale_radiation": {
      _norm("Nbr Chefs par articles par période de radiation"),
      _norm("Nbr Chefs par articles par periode de radiation"),
      _norm("Durée totale effective radiation"),
  },
  "article_amende_chef": {
      _norm("Nombre de chefs par articles et total amendes"),
      _norm("Article amende/chef"),
  },
  "autres_sanctions": {
      _norm("Nombre de chefs par article ayant une réprimande"),
      _norm("Autres sanctions"),
  },
  # Autres noms utiles pour la présentation
  "liste_chefs_articles": {
      _norm("Liste des chefs et articles en infraction"),
  },
  "total_amendes": {_norm("Total amendes")},
}

INTEREST_KEYS = [
    "articles_enfreints",
    "duree_totale_radiation",
    "article_amende_chef",
    "autres_sanctions",
]

def resolve_columns(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    mapping = {}
    inv = {_norm(c): c for c in df.columns}
    for canon, names in HEADER_ALIASES.items():
        hit = None
        for n in names:
            if n in inv:
                hit = inv[n]
                break
        mapping[canon] = hit
    return mapping

# ──────────────────────────────────────────────────────────────────────────────
# Lecture Excel (bannière « Article filtré : » en A1)
# ──────────────────────────────────────────────────────────────────────────────

def read_excel(file) -> pd.DataFrame:
    # on lit 2 lignes pour voir A1
    prev = pd.read_excel(file, header=None, nrows=2, engine="openpyxl")
    file.seek(0)
    first = prev.iloc[0,0] if not prev.empty else None
    skip = 1 if isinstance(first, str) and _norm(first).startswith(_norm("Article filtré :")) else 0
    return pd.read_excel(file, header=0, skiprows=skip, engine="openpyxl")

# ──────────────────────────────────────────────────────────────────────────────
# Motif exact & utils
# ──────────────────────────────────────────────────────────────────────────────

def build_article_regex(token: str) -> re.Pattern:
    token = (token or "").strip()
    if not token:
        raise ValueError("Article vide.")
    esc = re.escape(token)
    tail = r"(?![0-9.])" if token[-1].isdigit() else r"\b"
    return re.compile(rf"(?:\b(?:art(?:icle)?\s*[: ]*)?)({esc}){tail}", re.IGNORECASE)

def prep_text(v) -> str:
    if v is None or (isinstance(v, float) and math.isnan(v)):
        return ""
    s = str(v)
    return s.replace("\r\n", "\n").replace("\r", "\n")

def only_segments(text: str, pat: re.Pattern) -> str:
    if not text.strip():
        return ""
    parts = re.split(r"[;\n]", text)
    kept = [p.strip() for p in parts if pat.search(p)]
    return " | ".join(kept)

def format_amount(value) -> str:
    """Ex : 0 -> '0 $' ; 5000 -> '5 000 $' ; '5000' -> idem."""
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return "—"
    try:
        n = float(str(value).replace(" ", "").replace("$", "").replace(",", "."))
    except Exception:
        return escape(str(value))
    n = int(round(n))
    s = f"{n:,}".replace(",", " ")
    return f"{s} $"

def escape_and_highlight(s: str, pat: re.Pattern) -> str:
    esc = escape(s, quote=False)
    # remplace seulement la capture (le numéro d’article)
    return pat.sub(lambda m: m.group(0).replace(m.group(1), f'<span class="hl">{m.group(1)}</span>'), esc)

def to_ul_html(s: str, pat: re.Pattern) -> str:
    s = prep_text(s)
    if not s.strip():
        return "—"
    # on considère ; | puces, et retours à la ligne comme séparateurs
    segs = [x.strip(" •·") for x in re.split(r"(?:\n|;|\|)", s) if x and x.strip(" •·")]
    if len(segs) == 1:
        return escape_and_highlight(segs[0], pat)
    items = "".join(f"<li>{escape_and_highlight(x, pat)}</li>" for x in segs)
    return f'<ul class="bullets">{items}</ul>'

# ──────────────────────────────────────────────────────────────────────────────
# Vue HTML (DataFrame -> HTML avec listes + surlignage)
# ──────────────────────────────────────────────────────────────────────────────

def df_to_html(df: pd.DataFrame, pat: re.Pattern, colmap: Dict[str,str]) -> str:
    df2 = df.copy()

    # formatage Total amendes
    col_amount = colmap.get("total_amendes")
    if col_amount and col_amount in df2.columns:
        df2[col_amount] = df2[col_amount].apply(format_amount)

    # colonnes en listes à puces (toutes texte)
    def is_text(c):
        return str(df2[c].dtype) == "object"

    list_like: List[str] = []
    # 4 colonnes d’intérêt
    for k in INTEREST_KEYS:
        c = colmap.get(k)
        if c and c in df2.columns and is_text(c):
            list_like.append(c)
    # + la colonne « Liste des chefs et articles en infraction » si distincte
    c_extra = colmap.get("liste_chefs_articles")
    if c_extra and c_extra in df2.columns and is_text(c_extra):
        list_like.append(c_extra)

    # + quelques colonnes souvent narratives si elles existent
    for guess in ["Résumé des faits concis", "Autres mesures ordonnées", "À vérifier",
                  "Resume des faits concis", "Autres mesures ordonnees", "A verifier"]:
        if guess in df2.columns and is_text(guess):
            list_like.append(guess)

    # dédoublonnage
    list_like = list(dict.fromkeys(list_like))

    # application
    for c in df2.columns:
        if is_text(c):
            if c in list_like:
                df2[c] = df2[c].apply(lambda v: to_ul_html(v, pat))
            else:
                df2[c] = df2[c].apply(lambda v: escape_and_highlight(prep_text(v), pat) if str(v).strip() else "—")

    # petites classes “num” sur quelques colonnes connues
    classes = []
    for i, col in enumerate(df2.columns, 1):
        cls = "num" if _norm(col) in {_norm("Total chefs"), _norm("Total amendes")} else ""
        classes.append(cls)

    # rendu
    html = df2.to_html(index=False, escape=False)
    # injecter classes num dans <th> et <td> (optionnel)
    for i, cls in enumerate(classes, 1):
        if not cls:
            continue
        html = html.replace(f"<th>{df2.columns[i-1]}", f'<th class="{cls}">{df2.columns[i-1]}', 1)
    return html

# ──────────────────────────────────────────────────────────────────────────────
# Export Excel (ligne 1 “Article filtré : X”, en-tête figée)
# ──────────────────────────────────────────────────────────────────────────────

def to_excel(df: pd.DataFrame, article: str, colmap: Dict[str,str]) -> str:
    # copie export : contenu texte lisible, sans balises
    exp = df.copy()

    # total amendes en texte “5 000 $”
    c_amt = colmap.get("total_amendes")
    if c_amt and c_amt in exp.columns:
        exp[c_amt] = exp[c_amt].apply(format_amount)

    # listes -> texte multi-lignes (•)
    def to_lines(v):
        s = prep_text(v)
        if not s.strip():
            return "—"
        parts = [x.strip(" •·") for x in re.split(r"(?:\n|;|\|)", s) if x and x.strip(" •·")]
        return "• " + ("\n• ".join(parts)) if parts else s

    for c in exp.columns:
        if str(exp[c].dtype) == "object":
            exp[c] = exp[c].apply(to_lines)

    # écriture
    ts = int(time.time())
    path = f"/tmp/filtre_{ts}.xlsx"
    exp.to_excel(path, index=False, sheet_name="Filtre")
    wb = load_workbook(path)
    ws = wb.active

    # insérer la ligne 1 “Article filtré : X”
    ws.insert_rows(1)
    ws.cell(1, 1).value = f"Article filtré : {article}"
    ws.cell(1, 1).font = Font(bold=True)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ws.max_column)
    ws.cell(1, 1).alignment = Alignment(horizontal="left")

    # figer l’en-tête (ligne 2)
    ws.freeze_panes = "A3"

    # largeur auto raisonnable
    for j in range(1, ws.max_column + 1):
        let = get_column_letter(j)
        maxlen = 12
        for i in range(1, min(ws.max_row, 400) + 1):
            v = ws.cell(i, j).value
            l = 0 if v is None else len(str(v))
            if l > maxlen:
                maxlen = l
        ws.column_dimensions[let].width = min(60, max(12, maxlen + 2))

    wb.save(path)
    return f"/download?path={path}"

# ──────────────────────────────────────────────────────────────────────────────
# App
# ──────────────────────────────────────────────────────────────────────────────

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "GET":
        return render_template_string(HTML, style=STYLE, table_html=None, article="", segments_only=False, dl=None, msg=None)

    file = request.files.get("file")
    article = (request.form.get("article") or "").strip()
    segments_only = bool(request.form.get("segments_only"))

    if not file or not article:
        return render_template_string(HTML, style=STYLE, table_html=None, article=article, segments_only=segments_only,
                                      dl=None, msg="Veuillez fournir un fichier Excel et un article.")

    if not (file.filename or "").lower().endswith((".xlsx", ".xlsm")):
        return render_template_string(HTML, style=STYLE, table_html=None, article=article, segments_only=segments_only,
                                      dl=None, msg="Formats pris en charge : .xlsx / .xlsm")

    try:
        df = read_excel(file.stream)
    except Exception as e:
        return render_template_string(HTML, style=STYLE, table_html=None, article=article, segments_only=segments_only,
                                      dl=None, msg=f"Lecture Excel : {e!r}")

    colmap = resolve_columns(df)
    pat = build_article_regex(article)

    # Filtrage des lignes : si l’article apparait dans au moins une des 4 colonnes d’intérêt
    masks = []
    for key in INTEREST_KEYS:
        c = colmap.get(key)
        if c and c in df.columns:
            masks.append(df[c].astype(str).apply(lambda v: bool(pat.search(prep_text(v)))))
    if not masks:
        return render_template_string(HTML, style=STYLE, table_html=None, article=article, segments_only=segments_only,
                                      dl=None,
                                      msg="Aucune des 4 colonnes d’intérêt n’a été trouvée. Vérifie les en-têtes.")

    mask_any = masks[0]
    for m in masks[1:]:
        mask_any = mask_any | m

    df_f = df[mask_any].copy()
    if df_f.empty:
        return render_template_string(HTML, style=STYLE, table_html=None, article=article, segments_only=segments_only,
                                      dl=None, msg=f"Aucune ligne ne contient l’article « {article} ».")

    # Option : n’afficher que le SEGMENT contenant l’article dans les 4 colonnes d’intérêt
    if segments_only:
        for key in INTEREST_KEYS:
            c = colmap.get(key)
            if c and c in df_f.columns and str(df_f[c].dtype) == "object":
                df_f[c] = df_f[c].apply(lambda v: only_segments(prep_text(v), pat))

    # HTML
    table_html = df_to_html(df_f, pat, colmap)

    # Excel
    dl = to_excel(df_f, article, colmap)

    return render_template_string(HTML, style=STYLE, table_html=table_html, article=article,
                                  segments_only=segments_only, dl=dl, msg=None)

@app.route("/download")
def download():
    path = request.args.get("path")
    if not path or not os.path.exists(path):
        return "Fichier introuvable.", 404
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
