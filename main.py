# -*- coding: utf-8 -*-
# Analyseur Discipline – Filtrage par article (largeurs + puces + surlignage + export)
# 2025-09-09

import os
import re
import time
import unicodedata
from typing import Dict, Optional, Set, List

import pandas as pd
from flask import Flask, request, render_template_string, send_file

app = Flask(__name__)

# ──────────────────────────────────────────────────────────────────────────────
# Styles généraux + classes de largeur
# ──────────────────────────────────────────────────────────────────────────────
STYLE_BLOCK = """
<style>
  :root{
    --base-w: 220px;         /* largeur mini par défaut */
    --w-2x:   440px;         /* largeur mini 2x */
  }
  body{font-family: system-ui,-apple-system,"Segoe UI",Roboto,Helvetica,Arial,sans-serif; margin:24px;}
  h1{font-size:20px; margin:0 0 12px;}
  form{display:grid; gap:12px; margin:0 0 16px;}
  input[type="text"]{padding:8px; font-size:14px}
  input[type="file"]{font-size:14px}
  label{font-weight:600}
  .hint{font-size:12px; color:#666}
  .note{background:#fff6e5; border:1px solid #ffd89b; padding:8px 10px; border-radius:6px; margin:10px 0 16px}
  .download{margin:12px 0}

  /* Conteneur scrollable horizontal + vertical */
  .viewport{height:60vh; overflow:auto; border:1px solid #e5e7eb; border-radius:6px}
  /* On garde la table de Pandas mais on lui ajoute une classe */
  table.dataframe.table-custom{border-collapse:collapse; width:100%; table-layout:auto; font-size:13px}
  table.dataframe.table-custom th, table.dataframe.table-custom td{
    border:1px solid #e5e7eb; padding:6px 8px; vertical-align:top; white-space:normal; word-wrap:break-word
  }
  table.dataframe.table-custom th{background:#f3f4f6; text-align:center}
  /* Surlignage de l’article */
  .hit{color:#c1121f; font-weight:700}

  /* Largeurs : on applique par nth-child (injection CSS ciblée plus bas) */
  /* Classes génériques au cas où */
  .w-base{min-width:var(--base-w)}
  .w-2x{min-width:var(--w-2x)}

  /* UL propres pour les cellules à puces */
  .cell-list{margin:0; padding-left:18px}
</style>
"""

# ──────────────────────────────────────────────────────────────────────────────
# HTML de base
# ──────────────────────────────────────────────────────────────────────────────
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
    Règles : détection exacte de l’article ; si la 1<sup>re</sup> cellule contient « <code>Article filtré :</code> », la 2<sup>e</sup> ligne sert d’en-têtes.
  </div>

  <form method="POST" enctype="multipart/form-data">
    <div>
      <label>Article à rechercher (ex. <code>29</code>, <code>59(2)</code>)</label>
      <input type="text" name="article" value="{{ article or '' }}" required placeholder="ex.: 29 ou 59(2)">
    </div>
    <div>
      <label>Fichier Excel</label>
      <input type="file" name="file" accept=".xlsx,.xlsm" required>
    </div>
    <div>
      <label style="font-weight:500">
        <input type="checkbox" name="only_hits" {% if only_hits %}checked{% endif %}>
        Afficher uniquement le segment contenant l'article dans les 4 colonnes d’intérêt
      </label>
    </div>
    <button type="submit">Analyser</button>
    <div class="hint">Formats : .xlsx / .xlsm</div>
  </form>

  {% if table_html %}
    <div class="download">
      <a href="{{ download_url }}">Télécharger le résultat (Excel)</a>
    </div>
    <div class="viewport">
      {{ table_html|safe }}
    </div>
  {% endif %}

  {% if message %}
    <pre class="hint" style="white-space:pre-wrap">{{ message }}</pre>
  {% endif %}
</body>
</html>
"""

# ──────────────────────────────────────────────────────────────────────────────
# Normalisation en-têtes
# ──────────────────────────────────────────────────────────────────────────────
def _norm(s: str) -> str:
    if not isinstance(s, str):
        s = "" if s is None else str(s)
    s = unicodedata.normalize("NFKD", s).encode("ascii","ignore").decode("ascii")
    s = s.replace("\u00A0", " ")
    return " ".join(s.strip().lower().split())

# Alias d’en-têtes (canons → variantes)
HDR: Dict[str, Set[str]] = {
  "offences_list": {
    _norm("Liste des chefs et articles en infraction"),
    _norm("Articles en infraction"), _norm("Articles enfreints")
  },
  "per_article": {
    _norm("Nbr Chefs par articles"),
    _norm("Nombre de chefs par articles")
  },
  "per_period": {
    _norm("Nbr Chefs par articles par période de radiation"),
    _norm("Nbr Chefs par articles par periode de radiation")
  },
  "amendes": {
    _norm("Nombre de chefs par articles et total amendes"),
    _norm("Article amende/chef"), _norm("Articles amende / chef"),
  },
  "repr": {
    _norm("Nombre de chefs par article ayant une réprimande"),
    _norm("Nombre de chefs par article ayant une reprimande")
  },
  "faits": {
    _norm("Résumé des faits concis"),
    _norm("Resume des faits concis")
  },
  "sanctions": {
    _norm("Liste des sanctions imposées"),
    _norm("Liste des sanctions imposees")
  },
  "autres_mesures": {
    _norm("Autres mesures ordonnées"),
    _norm("Autres mesures ordonnees")
  },
  "a_verifier": {_norm("À vérifier"), _norm("A verifier")},
  "date_creation": {_norm("Date de création"), _norm("Date de creation")},
  "date_maj": {_norm("Date de mise à jour"), _norm("Date de mise a jour")},
  "no_decision": {_norm("Numéro de la décision"), _norm("Numero de la decision"), _norm("No decision")},
}

# Colonnes d’intérêt pour le FILTRAGE + surlignage
FILTER_CANON = ["per_article", "per_period", "amendes", "repr"]

def resolve_cols(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    mapping = {}
    inv = {_norm(c): c for c in df.columns}
    for canon, variants in HDR.items():
        hit = None
        for v in variants:
            if v in inv:
                hit = inv[v]; break
        mapping[canon] = hit
    return mapping

# ──────────────────────────────────────────────────────────────────────────────
# Lecture Excel (gestion bannière "Article filtré :")
# ──────────────────────────────────────────────────────────────────────────────
def read_excel(file) -> pd.DataFrame:
    preview = pd.read_excel(file, engine="openpyxl", header=None, nrows=2)
    file.seek(0)
    first = preview.iloc[0,0] if not preview.empty else None
    if isinstance(first, str) and _norm(first).startswith(_norm("Article filtré :")):
        return pd.read_excel(file, engine="openpyxl", skiprows=1)
    return pd.read_excel(file, engine="openpyxl")

# ──────────────────────────────────────────────────────────────────────────────
# Article → motif regex exact
# ──────────────────────────────────────────────────────────────────────────────
def build_pattern(article: str) -> re.Pattern:
    token = (article or "").strip()
    if not token: raise ValueError("Article vide")
    esc = re.escape(token)
    tail = r"(?![\d.])" if token[-1].isdigit() else r"\b"
    return re.compile(rf"(?:\b(?:art(?:icle)?\s*[: ]*)?)({esc}){tail}", re.IGNORECASE)

# ──────────────────────────────────────────────────────────────────────────────
# Utilitaires de texte
# ──────────────────────────────────────────────────────────────────────────────
def prep(v) -> str:
    if not isinstance(v,str): v = "" if v is None else str(v)
    return v.replace("\u00A0"," ").replace("\u202F"," ").replace("\r","").strip()

def to_bullets(text: str) -> str:
    """Transforme un bloc en <ul><li>…</li></ul>, sinon renvoie '—' si vide."""
    t = prep(text)
    if not t: return "—"
    # on coupe sur retours ligne ou '•' ou ' | '
    parts = []
    for raw in re.split(r"\n|•|\s\|\s", t):
        p = raw.strip(" ;,")
        if p: parts.append(p)
    if not parts: return "—"
    lis = "".join(f"<li>{p}</li>" for p in parts)
    return f'<ul class="cell-list">{lis}</ul>'

def highlight_article_html(text: str, pat: re.Pattern) -> str:
    if not text: return text
    # on ne modifie pas hors HTML
    return pat.sub(r'<span class="hit">\\1</span>', text)

# Extraction des segments contenant l’article (pour l’option “only_hits”)
def keep_only_hits(text: str, pat: re.Pattern) -> str:
    t = prep(text)
    if not t: return ""
    segs = re.split(r"[;\n]", t)
    kept = [s.strip() for s in segs if pat.search(s)]
    return " | ".join(kept)

# ──────────────────────────────────────────────────────────────────────────────
# Fabrication du tableau HTML (puces, surlignage, largeurs)
# ──────────────────────────────────────────────────────────────────────────────
def build_html_table(df: pd.DataFrame,
                     colmap: Dict[str,str],
                     pat: re.Pattern,
                     only_hits: bool) -> str:
    work = df.copy()

    # 1) Option : ne garder QUE le segment dans les 4 colonnes d’intérêt
    if only_hits:
        for k in FILTER_CANON:
            col = colmap.get(k)
            if col and col in work.columns:
                work[col] = work[col].apply(lambda v: keep_only_hits(v, pat))

    # 2) Surlignage dans les 4 colonnes d’intérêt + Liste des chefs…
    for k in FILTER_CANON + ["offences_list"]:
        col = colmap.get(k)
        if col and col in work.columns:
            work[col] = work[col].astype(str).apply(lambda v: highlight_article_html(v, pat))

    # 3) Rendu en listes à puces pour les colonnes “riches”
    bullet_cols = [
        colmap.get("faits"), colmap.get("offences_list"),
        colmap.get("sanctions"), colmap.get("per_article"),
        colmap.get("per_period"), colmap.get("amendes"),
        colmap.get("repr"), colmap.get("autres_mesures"),
        colmap.get("a_verifier"),
    ]
    for c in bullet_cols:
        if c and c in work.columns:
            work[c] = work[c].apply(to_bullets)

    # 4) NaN → tiret (attention : sans puce)
    work = work.fillna("—")

    # 5) Génération HTML
    html = work.to_html(index=False, escape=False)
    # On force notre classe (pour styles/table-layout) sans perdre “dataframe”
    html = html.replace('<table border="1" class="dataframe">', '<table border="1" class="dataframe table-custom">')

    # 6) Injection CSS de largeur par nth-child, en ciblant les libellés
    #    On lit les <th> pour retrouver leurs positions
    headers = re.findall(r"<th[^>]*>(.*?)</th>", html, flags=re.S)
    headers_clean = [re.sub("<.*?>","",h).strip() for h in headers]  # on retire les tags
    # Colonnes à élargir ×2
    double_set = {
        # groupe des 5 “simples”
        _norm("Résumé des faits concis"),
        _norm("Autres mesures ordonnées"),
        _norm("Date de création"),
        _norm("Date de mise à jour"),
        _norm("Numéro de la décision"),
        # groupe des 4 (même largeur)
        _norm("Nbr Chefs par articles"),
        _norm("Nbr Chefs par articles par période de radiation"),
        _norm("Nombre de chefs par articles et total amendes"),
        _norm("Nombre de chefs par article ayant une réprimande"),
    }
    # Colonnes “laisser tel quel” → aucune règle (offences_list & sanctions)
    leave_as_is = {
        _norm("Liste des chefs et articles en infraction"),
        _norm("Liste des sanctions imposées"),
    }

    # On fabrique des règles nth-child
    rules: List[str] = []
    for idx, title in enumerate(headers_clean, start=1):
        n = _norm(title)
        if n in leave_as_is:
            continue  # aucune règle = largeur auto
        if n in double_set:
            rules.append(f'.table-custom th:nth-child({idx}), .table-custom td:nth-child({idx})' +
                         '{min-width:var(--w-2x)}')
        else:
            # largeur de base pour les autres (ça évite de redevenir trop étroit)
            rules.append(f'.table-custom th:nth-child({idx}), .table-custom td:nth-child({idx})' +
                         '{min-width:var(--base-w)}')

    width_css = "<style>\n" + "\n".join(rules) + "\n</style>"
    # On insère ce bloc juste après l’ouverture de la table
    html = html.replace("<table ", width_css + "\n<table ")

    # 7) Nettoyage "nan" résiduels au cas où
    html = re.sub(r">\\s*(?:nan|NaN)\\s*<", ">—<", html)

    return html

# ──────────────────────────────────────────────────────────────────────────────
# Export Excel (auto-fit basique + feuille unique)
# ──────────────────────────────────────────────────────────────────────────────
def to_excel_download(df: pd.DataFrame) -> str:
    ts = int(time.time())
    out_path = f"/tmp/filtrage_{ts}.xlsx"
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Filtre")
        ws = writer.book.active
        # auto-fit simple
        for c_idx, col in enumerate(df.columns, start=1):
            max_len = max((len(str(x)) for x in [col] + df[col].astype(str).tolist()), default=10)
            width = min(60, max(12, max_len + 2))
            ws.column_dimensions[ws.cell(row=1, column=c_idx).column_letter].width = width
    return f"/download?path={out_path}"

# ──────────────────────────────────────────────────────────────────────────────
# Route principale
# ──────────────────────────────────────────────────────────────────────────────
@app.route("/", methods=["GET","POST"])
def main():
    if request.method == "GET":
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK,
                                      article="", only_hits=False,
                                      table_html=None, download_url=None, message=None)

    file = request.files.get("file")
    article = (request.form.get("article") or "").strip()
    only_hits = bool(request.form.get("only_hits"))

    if not file or not article:
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK,
                                      article=article, only_hits=only_hits,
                                      table_html=None, download_url=None,
                                      message="Veuillez fournir un article et un fichier Excel (.xlsx / .xlsm).")

    # formats pris en charge
    fname = (file.filename or "").lower()
    if not (fname.endswith(".xlsx") or fname.endswith(".xlsm")):
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK,
                                      article=article, only_hits=only_hits,
                                      table_html=None, download_url=None,
                                      message="Format non pris en charge. Utilisez un .xlsx ou .xlsm.")

    try:
        df = read_excel(file.stream)
        colmap = resolve_cols(df)
        pat = build_pattern(article)

        # Filtrage : on retient les lignes où l’article apparait AU MOINS dans une des 4 colonnes d’intérêt
        masks = []
        for k in FILTER_CANON:
            col = colmap.get(k)
            if col and col in df.columns:
                m = df[col].astype(str).apply(lambda v: bool(pat.search(prep(v))))
                masks.append(m)
        if not masks:
            return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK,
                                          article=article, only_hits=only_hits,
                                          table_html=None, download_url=None,
                                          message=("Aucune des 4 colonnes d’intérêt n’a été trouvée.\n"
                                                   f"Colonnes disponibles : {list(df.columns)}"))

        mask_any = masks[0]
        for m in masks[1:]: mask_any = mask_any | m
        df_keep = df[mask_any].copy()
        if df_keep.empty:
            return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK,
                                          article=article, only_hits=only_hits,
                                          table_html=None, download_url=None,
                                          message=f"Aucune ligne ne contient l’article « {article} » dans les colonnes d’intérêt.")

        # HTML (puces + surlignage + largeurs)
        table_html = build_html_table(df_keep, colmap, pat, only_hits)
        # Export Excel
        download_url = to_excel_download(df_keep if not only_hits else df_keep.assign(**{
            # si only_hits : on pose aussi la version “segment” dans le fichier
            (colmap[k] or f"_col_{k}"): df_keep[colmap[k]].apply(lambda v: keep_only_hits(v, pat))
            for k in FILTER_CANON if colmap.get(k) in df_keep.columns
        }))

        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK,
                                      article=article, only_hits=only_hits,
                                      table_html=table_html, download_url=download_url,
                                      message=None)

    except Exception as e:
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK,
                                      article=article, only_hits=only_hits,
                                      table_html=None, download_url=None,
                                      message=f"Erreur inattendue : {repr(e)}")

# ──────────────────────────────────────────────────────────────────────────────
# Téléchargement
# ──────────────────────────────────────────────────────────────────────────────
@app.route("/download")
def download():
    path = request.args.get("path")
    if not path or not os.path.exists(path):
        return "Fichier introuvable ou expiré.", 404
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))

# ──────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=False)
