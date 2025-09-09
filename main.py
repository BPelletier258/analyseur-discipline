# === CANVAS META =============================================================
# Fichier : main.py — x4 écrans, listes à puces, nettoyage \n & surlignage rouge (09‑08)
# Objet :
#  - Barre de défilement horizontale visible, largeur de table ≈ 4 écrans.
#  - Colonnes d'intérêt rendues en listes à puces (HTML) au lieu de texte brut.
#  - Suppression des "\n" (littéraux) et vrais retours chariot dans les cellules.
#  - Surlignage en rouge de l'article recherché dans les 4 colonnes cibles.
#  - Export Excel inchangé (texte brut, pas de HTML).
# ============================================================================

import io
import os
import re
import time
import unicodedata
from datetime import datetime
from typing import Dict, Optional, Set

import pandas as pd
from flask import Flask, request, render_template_string, send_file

app = Flask(__name__)

# ────────────────────────────────────────────────────────────────────────────
# Styles (x4 écrans, entêtes centrées, listes, surlignage)
# ────────────────────────────────────────────────────────────────────────────
STYLE_BLOCK = """
<style>
  :root{--fg:#111827;--muted:#6b7280;--hit:#b91c1c}
  body{font-family:system-ui,-apple-system,"Segoe UI",Roboto,Helvetica,Arial,sans-serif;margin:24px;color:var(--fg)}
  h1{font-size:20px;margin-bottom:12px}
  form{display:grid;gap:12px;margin-bottom:16px}
  input[type="text"]{padding:8px;font-size:14px}
  input[type="file"],button{font-size:14px}
  .note{background:#fff6e5;border:1px solid #ffd89b;padding:8px 10px;border-radius:6px;margin:10px 0 16px}
  .kbd{font-family:ui-monospace,SFMono-Regular,Menlo,monospace;background:#f3f4f6;padding:2px 4px;border-radius:4px}
  .download{margin:12px 0}

  /* Table et viewport */
  .table-viewport{height:62vh;overflow:auto;border:1px solid #e5e7eb}
  .table-wide{min-width:400vw} /* ≈ 4 écrans */
  table{border-collapse:collapse;width:100%;font-size:13px}
  th,td{border:1px solid #e5e7eb;padding:6px 8px;vertical-align:top;min-width:16rem}
  th{background:#f3f4f6;text-align:center}

  /* Listes et hits */
  ul.bullets{margin:0;padding-left:1.15em}
  ul.bullets li{margin:0 0 .35em 0}
  .hit{color:var(--hit);font-weight:600}

  /* Colonnes étroites utiles (si présentes) */
  .narrow{min-width:7rem}
</style>
"""

HTML_TEMPLATE = """
<!doctype html>
<html>
<head>
<meta charset=\"utf-8\"/>
<title>Analyseur Discipline – Filtrage par article</title>
{{ style_block|safe }}
</head>
<body>
  <h1>Analyseur Discipline – Filtrage par article</h1>
  <div class=note>
    Règles : détection exacte de l’article; si la 1<sup>re</sup> cellule contient « <span class=kbd>Article filtré :</span> »,
    elle est ignorée (entêtes sur la 2<sup>e</sup> ligne). Formats : <span class=kbd>.xlsx</span> / <span class=kbd>.xlsm</span>.
  </div>

  <form method=POST enctype=multipart/form-data>
    <label>Article à rechercher (ex. <span class=kbd>29</span>, <span class=kbd>59(2)</span>)</label>
    <input type=text name=article value="{{ searched_article or '' }}" required placeholder="ex.: 29 ou 59(2)"/>
    <label>Fichier Excel</label>
    <input type=file name=file accept=.xlsx,.xlsm required/>
    <button type=submit>Analyser</button>
  </form>

  {% if table_html %}
    <div class=download><a href="{{ download_url }}">Télécharger le résultat (Excel)</a></div>
    <div class=table-viewport><div class=table-wide>{{ table_html|safe }}</div></div>
  {% endif %}

  {% if message %}
    <div style="margin-top:10px;white-space:pre-wrap;font-family:ui-monospace,SFMono-Regular,Menlo,monospace;font-size:12px;">{{ message }}</div>
  {% endif %}
</body>
</html>
"""

# ────────────────────────────────────────────────────────────────────────────
# Normalisation & alias d’en‑têtes
# ────────────────────────────────────────────────────────────────────────────

def _norm(s: str) -> str:
    """Normalise : accents→ASCII, trim, minuscule, espaces compressés."""
    if not isinstance(s, str):
        s = str(s) if s is not None else ""
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = s.replace("\u00A0", " ")  # NBSP
    s = " ".join(s.strip().lower().split())
    return s

HEADER_ALIASES: Dict[str, Set[str]] = {
    # Noms cibles
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
        _norm("Duree totale effective radiation"),
    },
    "article_amende_chef": {
        _norm("Nombre de chefs par articles et total amendes"),
        _norm("Article amende/chef"),
    },
    "autres_sanctions": {
        _norm("Nombre de chefs par article ayant une réprimande"),
        _norm("Nombre de chefs par article ayant une reprimande"),
        _norm("Autres sanctions"),
    },
}

FILTER_CANONICAL = [
    "articles_enfreints",
    "duree_totale_radiation",
    "article_amende_chef",
    "autres_sanctions",
]


def resolve_columns(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    norm_to_original = {_norm(c): c for c in df.columns}
    out: Dict[str, Optional[str]] = {}
    for canon, variants in HEADER_ALIASES.items():
        found = None
        for v in variants:
            if v in norm_to_original:
                found = norm_to_original[v]
                break
        out[canon] = found
    return out

# ────────────────────────────────────────────────────────────────────────────
# Lecture Excel (règle « Article filtré : »)
# ────────────────────────────────────────────────────────────────────────────

def read_excel_respecting_header_rule(file_stream) -> pd.DataFrame:
    df_preview = pd.read_excel(file_stream, header=None, nrows=2, engine="openpyxl")
    file_stream.seek(0)

    first_cell = df_preview.iloc[0, 0] if not df_preview.empty else None
    ignore_first = isinstance(first_cell, str) and _norm(first_cell).startswith(_norm("Article filtré :"))

    if ignore_first:
        return pd.read_excel(file_stream, skiprows=1, header=0, engine="openpyxl")
    return pd.read_excel(file_stream, header=0, engine="openpyxl")

# ────────────────────────────────────────────────────────────────────────────
# Construction du motif exact de l’article
# ────────────────────────────────────────────────────────────────────────────

def build_article_pattern(user_input: str) -> re.Pattern:
    token = (user_input or "").strip()
    if not token:
        raise ValueError("Article vide.")
    esc = re.escape(token)
    tail = r"(?![\d.])" if token[-1].isdigit() else r"\b"
    return re.compile(rf"(?:\b(?:art(?:icle)?\s*[: ]*)?)({esc}){tail}", re.IGNORECASE)

# ────────────────────────────────────────────────────────────────────────────
# Pré‑traitement texte (suppression puces, NBSP, \n, CR/LF)
# ────────────────────────────────────────────────────────────────────────────

def _prep_text(v: str) -> str:
    if not isinstance(v, str):
        v = "" if v is None else str(v)
    # Supprime diverses puces visuelles
    v = v.replace("•", " ").replace("·", " ").replace("◦", " ")
    # Espaces insécables
    v = v.replace("\u00A0", " ").replace("\u202F", " ")
    # \n littéral -> espace
    v = v.replace("\\n", " ")
    # Vrais retours chariot -> espace
    v = v.replace("\r\n", " ").replace("\n", " ").replace("\r", " ")
    # Compression
    v = " ".join(v.split())
    return v

# ────────────────────────────────────────────────────────────────────────────
# Extraction (texte brut) — joint par «  |  » pour l’export Excel
# ────────────────────────────────────────────────────────────────────────────

def extract_mentions_generic(text: str, pat: re.Pattern) -> str:
    if not isinstance(text, str) or not text.strip():
        return ""
    parts = re.split(r"[;,:]|\s\|\s|\|", text)  # tolérance
    hits = [p.strip() for p in parts if pat.search(p)]
    return " | ".join(hits)


def extract_mentions_autres_sanctions(text: str, pat: re.Pattern) -> str:
    if not isinstance(text, str) or not text.strip():
        return ""
    parts = re.split(r"[;]|\s\|\s|\|", text)
    hits = [p.strip() for p in parts if pat.search(p)]
    return " | ".join(hits)


def clean_filtered_df(df: pd.DataFrame, colmap: Dict[str, Optional[str]], pat: re.Pattern) -> pd.DataFrame:
    df = df.copy()
    for canon in FILTER_CANONICAL:
        col = colmap.get(canon)
        if not col or col not in df.columns:
            continue
        if canon == "autres_sanctions":
            df[col] = df[col].apply(lambda v: extract_mentions_autres_sanctions(_prep_text(v), pat))
        else:
            df[col] = df[col].apply(lambda v: extract_mentions_generic(_prep_text(v), pat))

    # Ne garder que les lignes où au moins une des colonnes cibles contient quelque chose
    target_cols = [c for c in (colmap.get(k) for k in FILTER_CANONICAL) if c]
    if target_cols:
        mask_any = False
        for c in target_cols:
            cur = df[c].astype(str).str.strip().ne("")
            mask_any = cur if mask_any is False else (mask_any | cur)
        df = df[mask_any]
    return df

# ────────────────────────────────────────────────────────────────────────────
# Rendu HTML : conversion «  |  » -> <ul><li>…</li></ul> + surlignage
# ────────────────────────────────────────────────────────────────────────────

def to_bullets_html(val: str, pat: re.Pattern) -> str:
    if not isinstance(val, str) or not val.strip():
        return ""
    items = [s.strip() for s in re.split(r"\s*\|\s*", val) if s.strip()]
    if not items:
        return ""
    html_items = []
    for s in items:
        s = pat.sub(r"<span class=\"hit\">\\1</span>", s)
        html_items.append(f"<li>{s}</li>")
    return "<ul class=\"bullets\">" + "".join(html_items) + "</ul>"

# ────────────────────────────────────────────────────────────────────────────
# Export Excel (texte brut)
# ────────────────────────────────────────────────────────────────────────────

def to_excel_download(df: pd.DataFrame) -> str:
    ts = int(time.time())
    out_path = f"/tmp/filtrage_{ts}.xlsx"
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Filtre")
        ws = writer.book.active
        for col_idx, col in enumerate(df.columns, start=1):
            max_len = max((len(str(x)) for x in [col] + df[col].astype(str).tolist()), default=10)
            ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = min(60, max(12, max_len + 2))
    return f"/download?path={out_path}"

# ────────────────────────────────────────────────────────────────────────────
# Routes
# ────────────────────────────────────────────────────────────────────────────

@app.route("/", methods=["GET", "POST"])
def analyze():
    if request.method == "GET":
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                                      searched_article=None, message=None, download_url=None)

    file = request.files.get("file")
    article = (request.form.get("article") or "").strip()

    if not file or not article:
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                                      searched_article=article,
                                      message="Erreur : fichier et article sont requis.", download_url=None)

    # Seuls .xlsx / .xlsm (openpyxl)
    fname = (file.filename or "").lower()
    if not (fname.endswith(".xlsx") or fname.endswith(".xlsm")):
        ext = (file.filename or "?").split(".")[-1]
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                                      searched_article=article,
                                      message=f"Format non pris en charge : .{ext}. Fournir un fichier .xlsx ou .xlsm.",
                                      download_url=None)

    try:
        df = read_excel_respecting_header_rule(file.stream)
        colmap = resolve_columns(df)
        pat = build_article_pattern(article)

        # Filtrage initial (au moins une colonne cible contient le motif brut)
        masks = []
        any_cols = False
        for canon in FILTER_CANONICAL:
            col = colmap.get(canon)
            if col and col in df.columns:
                any_cols = True
                masks.append(df[col].astype(str).apply(lambda v: bool(pat.search(_prep_text(v)))))
        if not any_cols:
            detail = "\n".join([f"  - {k}: {colmap.get(k)}" for k in FILTER_CANONICAL])
            return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                                          searched_article=article,
                                          message=("Erreur : aucune des colonnes attendues n’a été trouvée.\n" +
                                                   f"Colonnes résolues :\n{detail}"),
                                          download_url=None)
        mask_any = masks[0]
        for m in masks[1:]:
            mask_any = mask_any | m

        df_filtered = df[mask_any].copy()
        if df_filtered.empty:
            return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                                          searched_article=article,
                                          message=f"Aucune ligne ne contient l’article « {article} » dans les colonnes cibles.",
                                          download_url=None)

        # Nettoyage & extraction (texte brut pour export)
        df_clean = clean_filtered_df(df_filtered, colmap, pat)
        if df_clean.empty:
            return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                                          searched_article=article,
                                          message=("Des lignes correspondaient au motif, mais après épuration, "
                                                   "aucune mention nette n’a été conservée."),
                                          download_url=None)

        # Préparation rendu HTML : conversion des 4 colonnes en listes + surlignage
        preview = df_clean.head(200).copy()
        for canon in FILTER_CANONICAL:
            col = colmap.get(canon)
            if col and col in preview.columns:
                preview[col] = preview[col].apply(lambda v: to_bullets_html(v, pat))

        # Génération HTML (escape=False pour laisser les <ul>/<li> et le surlignage)
        table_html = preview.to_html(index=False, escape=False)

        download_url = to_excel_download(df_clean)
        msg = f"{len(df_clean)} ligne(s) après filtrage et épuration. (Aperçu limité à 200 lignes.)"
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=table_html,
                                      searched_article=article, message=msg, download_url=download_url)

    except Exception as e:
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                                      searched_article=article,
                                      message=f"Erreur inattendue : {repr(e)}", download_url=None)


@app.route("/download")
def download():
    path = request.args.get("path")
    if not path or not os.path.exists(path):
        return "Fichier introuvable ou expiré.", 404
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
