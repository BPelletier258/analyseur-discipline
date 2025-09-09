# === CANVAS META =============================================================
# Fichier : main.py — version visible + motif strict + support .xlsm
# - Normalisation des en-têtes (accents, casse, espaces insécables)
# - Aliases à jour pour les nouveaux titres
# - Règle « Article filtré : » tolérante (entêtes sur 2e ligne)
# - Filtrage EXACT sur l’article
# - Aperçu HTML : listes à puces + surlignage (3 colonnes) + largeur 4 écrans
# - Export Excel : brut, sans HTML
# ============================================================================

import io
import os
import re
import time
import unicodedata
from html import escape
from datetime import datetime
from typing import Dict, Optional, Set

import pandas as pd
from flask import Flask, request, render_template_string, send_file

app = Flask(__name__)

# -----------------------------------------------------------------------------
# Styles
# -----------------------------------------------------------------------------
STYLE_BLOCK = """
<style>
  body { font-family: system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif; margin: 24px; }
  h1 { font-size: 20px; margin-bottom: 12px; }
  form { display: grid; gap: 12px; margin-bottom: 16px; }
  input[type="text"] { padding: 8px; font-size: 14px; }
  input[type="file"] { font-size: 14px; }
  button { padding: 8px 12px; font-size: 14px; cursor: pointer; }
  .hint { font-size: 12px; color: #666; }
  .note { background: #fff6e5; border: 1px solid #ffd89b; padding: 8px 10px; border-radius: 6px; margin: 10px 0 16px; }

  table { border-collapse: collapse; width: 100%; font-size: 13px; }
  th, td { border: 1px solid #ddd; padding: 6px 8px; vertical-align: top; }
  th { background: #f3f4f6; text-align: center; }
  .msg { margin-top: 12px; white-space: pre-wrap; font-family: ui-monospace, SFMono-Regular, Menlo, monospace; font-size: 12px; }
  .ok { color: #065f46; }
  .err { color: #7f1d1d; }
  .download { margin: 12px 0; }
  .kbd { font-family: ui-monospace, SFMono-Regular, Menlo, monospace; background:#f3f4f6; padding:2px 4px; border-radius:4px; }

  /* Vue avec barre horizontale + largeur ≈ 4 écrans */
  .table-viewport{height:60vh; overflow:auto; border:1px solid #ddd;}
  .table-wide{min-width:400vw;}
  .table-viewport table{width:100%;}

  /* Listes dans les cellules */
  ul.cell-list{margin:0; padding-left:18px;}
  ul.cell-list li{margin:0 0 2px 0; list-style: disc;}
  .hit{color:#b91c1c; font-weight:600;}
</style>
"""

# -----------------------------------------------------------------------------
# Page HTML
# -----------------------------------------------------------------------------
HTML_TEMPLATE = """
<!doctype html>
<html>
<head>
<meta charset="utf-8" />
<title>Analyseur Discipline – Filtrage par article</title>
{{ style_block|safe }}
</head>
<body>
  <h1>Analyseur Discipline – Filtrage par article</h1>

  <div class="note">
    Règles : détection exacte de l’article; si la 1<sup>re</sup> cellule contient
    « <span class="kbd">Article filtré :</span> », elle est ignorée (entêtes sur la 2<sup>e</sup> ligne).
  </div>

  <form method="POST" enctype="multipart/form-data">
    <label>Article à rechercher (ex. <span class="kbd">29</span>, <span class="kbd">59(2)</span>)</label>
    <input type="text" name="article" value="{{ searched_article or '' }}" required placeholder="ex.: 29 ou 59(2)" />
    <label>Fichier Excel</label>
    <input type="file" name="file" accept=".xlsx,.xlsm" required />
    <button type="submit">Analyser</button>
    <div class="hint">Formats : .xlsx / .xlsm</div>
  </form>

  {% if table_html %}
    <div class="download">
      <a href="{{ download_url }}">Télécharger le résultat (Excel)</a>
    </div>
    <div class="table-viewport"><div class="table-wide">{{ table_html|safe }}</div></div>
  {% endif %}

  {% if message %}
    <div class="msg {{ 'ok' if message_ok else 'err' }}">{{ message }}</div>
  {% endif %}
</body>
</html>
"""

# -----------------------------------------------------------------------------
# Normalisation & alias d’en-têtes
# -----------------------------------------------------------------------------
def _norm(s: str) -> str:
    """Accents→ASCII, trim, minuscule, espaces compressés, NBSP→espace."""
    if not isinstance(s, str):
        s = str(s) if s is not None else ""
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = s.replace("\u00A0", " ")
    s = " ".join(s.strip().lower().split())
    return s

# Canoniques → alias
HEADER_ALIASES: Dict[str, Set[str]] = {
    # 1) Liste des chefs et articles en infraction
    "articles_enfreints": {
        _norm("Liste des chefs et articles en infraction"),
        _norm("Nbr Chefs par articles"),     # tolérer l’ancien pivot
        _norm("Articles enfreints"),
        _norm("Articles en infraction"),
    },
    # 2) Nbr Chefs par articles par période de radiation
    "duree_totale_radiation": {
        _norm("Nbr Chefs par articles par période de radiation"),
        _norm("Nbr Chefs par articles par periode de radiation"),
        _norm("Durée totale effective radiation"),
        _norm("Duree totale effective radiation"),
    },
    # 3) Nombre de chefs par articles et total amendes
    "article_amende_chef": {
        _norm("Nombre de chefs par articles et total amendes"),
        _norm("Article amende/chef"),
        _norm("Articles amende / chef"),
        _norm("Amendes (article/chef)"),
    },
    # 4) Nombre de chefs par article ayant une réprimande
    "autres_sanctions": {
        _norm("Nombre de chefs par article ayant une réprimande"),
        _norm("Nombre de chefs par article ayant une reprimande"),
        _norm("Autres sanctions"),
        _norm("Autres mesures ordonnées"),
        _norm("Autres sanctions / mesures"),
    },
}

FILTER_CANONICAL = [
    "articles_enfreints",
    "duree_totale_radiation",
    "article_amende_chef",
    "autres_sanctions",
]

def resolve_columns(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    """Retourne {canonique: libellé réel} si présent, sinon None."""
    norm_to_original = {_norm(c): c for c in df.columns}
    resolved: Dict[str, Optional[str]] = {}
    for canon, variants in HEADER_ALIASES.items():
        hit = None
        for v in variants:
            if v in norm_to_original:
                hit = norm_to_original[v]
                break
        resolved[canon] = hit
    return resolved

# -----------------------------------------------------------------------------
# Lecture Excel (règle « Article filtré : »)
# -----------------------------------------------------------------------------
def read_excel_respecting_header_rule(file_stream) -> pd.DataFrame:
    df_preview = pd.read_excel(file_stream, header=None, nrows=2, engine="openpyxl")
    file_stream.seek(0)

    first_cell = df_preview.iloc[0, 0] if not df_preview.empty else None
    banner = isinstance(first_cell, str) and _norm(first_cell).startswith(_norm("Article filtré :"))

    if banner:
        df = pd.read_excel(file_stream, skiprows=1, header=0, engine="openpyxl")
    else:
        df = pd.read_excel(file_stream, header=0, engine="openpyxl")
    return df

# -----------------------------------------------------------------------------
# Motif exact pour l’article (29, 59(2), 2.01 a), etc.)
# -----------------------------------------------------------------------------
def build_article_pattern(user_input: str) -> re.Pattern:
    token = (user_input or "").strip()
    if not token:
        raise ValueError("Article vide.")
    esc = re.escape(token)
    tail = r"(?![\d.])" if token[-1].isdigit() else r"\b"
    pattern = rf"(?:\b(?:art(?:icle)?\s*[: ]*)?)({esc}){tail}"
    return re.compile(pattern, flags=re.IGNORECASE)

# -----------------------------------------------------------------------------
# Nettoyages légers pour la RECHERCHE (pas pour l’affichage)
# -----------------------------------------------------------------------------
def _prep_text(v: str) -> str:
    if not isinstance(v, str):
        v = "" if v is None else str(v)
    # uniformiser les puces et espaces spéciaux
    v = v.replace("•", " ").replace("·", " ").replace("◦", " ")
    v = v.replace("\u00A0", " ").replace("\u202F", " ")
    # normaliser retours
    v = v.replace("\r\n", "\n").replace("\r", "\n")
    # compresser
    v = " ".join(v.split())
    return v

# -----------------------------------------------------------------------------
# Fabrication de listes HTML + surlignage
# -----------------------------------------------------------------------------
def _listify_html(text: str, pat: Optional[re.Pattern] = None, highlight: bool = False) -> str:
    """
    Convertit un bloc texte en <ul><li>…</li></ul>.
    Séparateurs : puces ou retours ligne. Surligne (span.hit) le groupe capturé.
    """
    if text is None:
        return ""
    s = str(text).replace("\r\n", "\n").replace("\r", "\n")
    s = s.replace("•", "\n").replace("·", "\n").replace("◦", "\n")
    raw = [t.strip(" \t•-") for t in re.split(r"\n+", s) if t and t.strip(" \t•-")]
    if not raw:
        return ""

    def _hilite(item: str) -> str:
        if not highlight or pat is None:
            return escape(item)
        out, last = [], 0
        for m in pat.finditer(item):
            out.append(escape(item[last:m.start(1)]))
            out.append(f'<span class="hit">{escape(m.group(1))}</span>')
            last = m.end(1)
        out.append(escape(item[last:]))
        return "".join(out)

    items_html = "".join(f"<li>{_hilite(it)}</li>" for it in raw)
    return f"<ul class='cell-list'>{items_html}</ul>"

# -----------------------------------------------------------------------------
# Export Excel (brut)
# -----------------------------------------------------------------------------
def to_excel_download(df: pd.DataFrame) -> str:
    ts = int(time.time())
    out_path = f"/tmp/filtrage_{ts}.xlsx"
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Filtre")
        ws = writer.book.active
        # largeur colonnes approx.
        for col_idx, col in enumerate(df.columns, start=1):
            max_len = max((len(str(x)) for x in [col] + df[col].astype(str).tolist()), default=10)
            ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = min(60, max(12, max_len + 2))
    return f"/download?path={out_path}"

# -----------------------------------------------------------------------------
# Route principale
# -----------------------------------------------------------------------------
@app.route("/", methods=["GET", "POST"])
def analyze():
    if request.method == "GET":
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK,
                                      table_html=None, searched_article=None,
                                      message=None, message_ok=True)

    file = request.files.get("file")
    article = (request.form.get("article") or "").strip()

    if not file or not article:
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK,
                                      table_html=None, searched_article=article,
                                      message="Erreur : fichier et article sont requis.",
                                      message_ok=False)

    # uniquement .xlsx/.xlsm
    fname = (file.filename or "").lower()
    if not (fname.endswith(".xlsx") or fname.endswith(".xlsm")):
        ext = (file.filename or "").split(".")[-1]
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK,
                                      table_html=None, searched_article=article,
                                      message=f"Format non pris en charge : {ext}. "
                                              f"Veuillez fournir un classeur .xlsx ou .xlsm.",
                                      message_ok=False)

    try:
        # 1) lecture + détection bannière
        df = read_excel_respecting_header_rule(file.stream)
        # 2) résolutions des colonnes
        colmap = resolve_columns(df)
        # 3) motif
        pat = build_article_pattern(article)

        # 4) filtrage des lignes : l’article doit apparaître dans ≥ 1 colonne d’intérêt
        masks = []
        any_cols = False
        for canon in FILTER_CANONICAL:
            col = colmap.get(canon)
            if col and col in df.columns:
                any_cols = True
                masks.append(df[col].astype(str).apply(lambda v: bool(pat.search(_prep_text(v)))))
        if not any_cols:
            detail = "\n".join([f"  - {k}: {colmap.get(k)}" for k in FILTER_CANONICAL])
            return render_template_string(
                HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                searched_article=article,
                message=("Erreur : aucune des colonnes attendues n’a été trouvée dans le fichier.\n"
                         "Vérifiez les en-têtes ou ajustez les alias.\n\n"
                         f"Colonnes résolues :\n{detail}\n\nColonnes disponibles :\n{list(df.columns)}"),
                message_ok=False
            )

        if not masks:
            return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK,
                                          table_html=None, searched_article=article,
                                          message="Aucune colonne exploitable pour le filtrage.",
                                          message_ok=False)

        mask_any = masks[0]
        for m in masks[1:]:
            mask_any = mask_any | m

        df_filtered = df[mask_any].copy()
        if df_filtered.empty:
            return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK,
                                          table_html=None, searched_article=article,
                                          message=f"Aucune ligne ne contient l’article « {article} » dans les colonnes cibles.",
                                          message_ok=True)

        # 5) Export : BRUT (pas de coupes)
        download_url = to_excel_download(df_filtered)

        # 6) Préparation APERÇU HTML : listes + surlignage sur 3 colonnes
        df_view = df_filtered.copy()
        col_chefs_articles = colmap.get("articles_enfreints")
        col_radiations     = colmap.get("duree_totale_radiation")
        col_amendes        = colmap.get("article_amende_chef")
        col_reprimandes    = colmap.get("autres_sanctions")

        if col_chefs_articles and col_chefs_articles in df_view.columns:
            # pas de surlignage dans cette colonne
            df_view[col_chefs_articles] = df_view[col_chefs_articles].apply(
                lambda v: _listify_html(v, pat, highlight=False)
            )
        if col_radiations and col_radiations in df_view.columns:
            df_view[col_radiations] = df_view[col_radiations].apply(
                lambda v: _listify_html(v, pat, highlight=True)
            )
        if col_amendes and col_amendes in df_view.columns:
            df_view[col_amendes] = df_view[col_amendes].apply(
                lambda v: _listify_html(v, pat, highlight=True)
            )
        if col_reprimandes and col_reprimandes in df_view.columns:
            df_view[col_reprimandes] = df_view[col_reprimandes].apply(
                lambda v: _listify_html(v, pat, highlight=True)
            )

        preview = df_view.head(200)
        table_html = preview.to_html(index=False, escape=False)

        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK,
                                      table_html=table_html, searched_article=article,
                                      download_url=download_url,
                                      message=f"{len(df_filtered)} ligne(s) après filtrage. (Aperçu limité à 200 lignes.)",
                                      message_ok=True)

    except Exception as e:
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK,
                                      table_html=None, searched_article=article,
                                      message=f"Erreur inattendue : {repr(e)}",
                                      message_ok=False)

# -----------------------------------------------------------------------------
# Téléchargement
# -----------------------------------------------------------------------------
@app.route("/download")
def download():
    path = request.args.get("path")
    if not path or not os.path.exists(path):
        return "Fichier introuvable ou expiré.", 404
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))

# -----------------------------------------------------------------------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
