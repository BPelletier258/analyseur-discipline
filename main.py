# === CANVAS META =============================================================
# Fichier : main.py — v6 (autosheet + listes HTML complètes + surlignage) — 2025-09-08
# Note : Aperçu HTML = listes à puces (tout le contenu conservé) + mise en évidence en rouge
#        Export Excel = contenu brut (sans HTML), largeurs ajustées.
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

# ──────────────────────────────────────────────────────────────────────────────
# Styles
# ──────────────────────────────────────────────────────────────────────────────
STYLE_BLOCK = """
<style>
  body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;margin:24px}
  h1{font-size:20px;margin-bottom:12px}
  form{display:grid;gap:12px;margin-bottom:16px}
  input[type=text]{padding:8px;font-size:14px}
  input[type=file]{font-size:14px}
  button{padding:8px 12px;font-size:14px;cursor:pointer}
  .note{background:#fff6e5;border:1px solid #ffd89b;padding:8px 10px;border-radius:6px;margin:10px 0 16px}
  .kbd{font-family:ui-monospace,SFMono-Regular,Menlo,monospace;background:#f3f4f6;padding:2px 4px;border-radius:4px}
  .download{margin:12px 0}

  /* Tableau */
  .table-viewport{height:60vh;overflow:auto;border:1px solid #ddd}
  .table-wide{min-width:200vw} /* ≈ deux écrans */
  table{border-collapse:collapse;width:100%;font-size:13px}
  th,td{border:1px solid #ddd;padding:6px 8px;vertical-align:top}
  th{background:#f3f4f6;text-align:center}
  td{white-space:normal}

  /* Listes compactes dans les cellules */
  ul.ul-tight{margin:0;padding-left:18px}
  ul.ul-tight li{margin:0 0 2px 0}
  .hit{color:#b91c1c;font-weight:700}

  .msg{margin-top:12px;white-space:pre-wrap;font-family:ui-monospace,SFMono-Regular,Menlo,monospace;font-size:12px}
  .ok{color:#065f46}.err{color:#7f1d1d}
</style>
"""

HTML_TEMPLATE = """
<!doctype html>
<html>
<head>
<meta charset=\"utf-8\" />
<title>Analyseur Discipline – Filtrage par article</title>
{{ style_block|safe }}
</head>
<body>
  <h1>Analyseur Discipline – Filtrage par article</h1>
  <div class=note>
    Règles : détection exacte de l’article; si la 1<sup>re</sup> cellule contient « <span class=kbd>Article filtré :</span> », elle est ignorée (entêtes sur la 2<sup>e</sup> ligne).
  </div>
  <form method=POST enctype=multipart/form-data>
    <label>Article à rechercher (ex. <span class=kbd>29</span>, <span class=kbd>59(2)</span>)</label>
    <input type=text name=article value=""" + "{{ searched_article or '' }}" + """ required placeholder=\"ex.: 29 ou 59(2)\" />
    <label>Fichier Excel</label>
    <input type=file name=file accept=.xlsx,.xlsm required />
    <button type=submit>Analyser</button>
    <div class=hint>Formats : .xlsx / .xlsm</div>
  </form>
  {% if table_html %}
    <div class=download><a href=""" + "{{ download_url }}" + """>Télécharger le résultat (Excel)</a></div>
    <div class=table-viewport><div class=table-wide>{{ table_html|safe }}</div></div>
  {% endif %}
  {% if message %}<div class=\"msg {{ 'ok' if message_ok else 'err' }}\">{{ message }}</div>{% endif %}
</body>
</html>
"""

# ──────────────────────────────────────────────────────────────────────────────
# Normalisation & alias d’en‑têtes
# ──────────────────────────────────────────────────────────────────────────────

def _norm(s: str) -> str:
    if not isinstance(s, str):
        s = str(s) if s is not None else ""
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = s.replace("\u00A0", " ")
    s = " ".join(s.strip().lower().split())
    return s

HEADER_ALIASES: Dict[str, Set[str]] = {
    "articles_enfreints": { _norm("Nbr Chefs par articles"), _norm("Articles enfreints"), _norm("Articles en infraction"), _norm("Liste des chefs et articles en infraction") },
    "duree_totale_radiation": { _norm("Nbr Chefs par articles par période de radiation"), _norm("Nbr Chefs par articles par periode de radiation"), _norm("Durée totale effective radiation"), _norm("Duree totale effective radiation") },
    "article_amende_chef": { _norm("Nombre de chefs par articles et total amendes"), _norm("Article amende/chef"), _norm("Articles amende / chef") },
    "autres_sanctions": { _norm("Nombre de chefs par article ayant une réprimande"), _norm("Nombre de chefs par article ayant une reprimande"), _norm("Autres sanctions"), _norm("Autres mesures ordonnées") },
}

FILTER_CANONICAL = [
    "articles_enfreints",
    "duree_totale_radiation",
    "article_amende_chef",
    "autres_sanctions",
]


def resolve_columns(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    norm_to_original = { _norm(c): c for c in df.columns }
    out: Dict[str, Optional[str]] = {}
    for canon, variants in HEADER_ALIASES.items():
        hit = None
        for v in variants:
            if v in norm_to_original:
                hit = norm_to_original[v]; break
        out[canon] = hit
    return out

# ──────────────────────────────────────────────────────────────────────────────
# Lecture Excel (gestion « Article filtré : »)
# ──────────────────────────────────────────────────────────────────────────────

def read_excel_respecting_header_rule(file_stream) -> pd.DataFrame:
    df_preview = pd.read_excel(file_stream, header=None, nrows=2, engine="openpyxl"); file_stream.seek(0)
    first_cell = df_preview.iloc[0,0] if not df_preview.empty else None
    banner = isinstance(first_cell, str) and _norm(first_cell).startswith(_norm("Article filtré :"))
    if banner:
        return pd.read_excel(file_stream, skiprows=1, header=0, engine="openpyxl")
    return pd.read_excel(file_stream, header=0, engine="openpyxl")

# ──────────────────────────────────────────────────────────────────────────────
# Motif exact
# ──────────────────────────────────────────────────────────────────────────────

def build_article_pattern(user_input: str) -> re.Pattern:
    token = (user_input or "").strip()
    if not token:
        raise ValueError("Article vide.")
    esc = re.escape(token)
    ends_with_digit = token[-1].isdigit()
    tail = r"(?![\d.])" if ends_with_digit else r"\b"
    pattern = rf"(?:\b(?:art(?:icle)?\s*[: ]*)?)({esc}){tail}"
    return re.compile(pattern, flags=re.IGNORECASE)

# ──────────────────────────────────────────────────────────────────────────────
# Utilitaires contenu → listes HTML (sans perte) + surlignage
# ──────────────────────────────────────────────────────────────────────────────

def _normalize_newlines(s: str) -> str:
    return s.replace("\r\n", "\n").replace("\r", "\n")

_SPLIT_RE = re.compile(r"\n|•|\u2022|\u2023|\u25E6|;|\u00B7")


def highlight_hits(text: str, pat: re.Pattern) -> str:
    if not isinstance(text, str):
        text = "" if text is None else str(text)
    return pat.sub(r'<span class="hit">\\1</span>', text)


def as_bullet_list_html(value: str, pat: re.Pattern) -> str:
    if not isinstance(value, str) or not value.strip():
        return ""
    # Ne pas supprimer les virgules (elles font partie des items) — éviter la perte de contexte
    txt = _normalize_newlines(value)
    # Si l’unique séparateur est l’espace, on garde tel quel
    items = [x.strip() for x in _SPLIT_RE.split(txt) if x and x.strip()]
    if not items:
        # Rien d’évident à scinder : on met simplement en valeur les occurrences
        return highlight_hits(value, pat)
    lis = [f"<li>{highlight_hits(it, pat)}</li>" for it in items]
    return f"<ul class='ul-tight'>{''.join(lis)}</ul>"

# ──────────────────────────────────────────────────────────────────────────────
# Aperçu : transformer seulement pour HTML (export = brut)
# ──────────────────────────────────────────────────────────────────────────────

def dataframe_for_preview(df: pd.DataFrame, colmap: Dict[str, Optional[str]], pat: re.Pattern) -> pd.DataFrame:
    view = df.copy()
    for canon in FILTER_CANONICAL:
        col = colmap.get(canon)
        if col and col in view.columns:
            view[col] = view[col].apply(lambda v: as_bullet_list_html(v, pat))
    return view

# ──────────────────────────────────────────────────────────────────────────────
# Export Excel (brut, sans HTML)
# ──────────────────────────────────────────────────────────────────────────────

def to_excel_download(df: pd.DataFrame) -> str:
    ts = int(time.time()); out_path = f"/tmp/filtrage_{ts}.xlsx"
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Filtre")
        ws = writer.book.active
        # Ajustement simple des largeurs
        for j, col in enumerate(df.columns, start=1):
            vals = [len(str(col))] + [len(str(x)) for x in df[col].astype(str).tolist()]
            width = min(60, max(12, max(vals) + 2))
            ws.column_dimensions[ws.cell(row=1, column=j).column_letter].width = width
    return f"/download?path={out_path}"

# ──────────────────────────────────────────────────────────────────────────────
# Route principale
# ──────────────────────────────────────────────────────────────────────────────

@app.route('/', methods=['GET','POST'])
def analyze():
    if request.method == 'GET':
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                                      searched_article=None, message=None, message_ok=True)

    file = request.files.get('file'); article = (request.form.get('article') or '').strip()
    if not file or not article:
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                                      searched_article=article,
                                      message='Erreur : fichier et article sont requis.', message_ok=False)

    fname = (file.filename or '').lower()
    if not (fname.endswith('.xlsx') or fname.endswith('.xlsm')):
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                                      searched_article=article,
                                      message='Format non pris en charge. Fournissez un .xlsx ou .xlsm.',
                                      message_ok=False)
    try:
        df = read_excel_respecting_header_rule(file.stream)
        colmap = resolve_columns(df)
        pat = build_article_pattern(article)

        # Filtre des lignes (sans mutiler le contenu)
        masks = []
        any_cols = False
        for canon in FILTER_CANONICAL:
            col = colmap.get(canon)
            if col and col in df.columns:
                any_cols = True
                masks.append(df[col].astype(str).apply(lambda v: bool(pat.search(_normalize_newlines(v)))))
        if not any_cols:
            detail = "\n".join([f"  - {k}: {colmap.get(k)}" for k in FILTER_CANONICAL])
            return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                                          searched_article=article,
                                          message=("Erreur : colonnes cibles introuvables.\n\n"
                                                   f"Colonnes résolues :\n{detail}\n\nColonnes disponibles :\n{list(df.columns)}"),
                                          message_ok=False)
        mask_any = masks[0]
        for m in masks[1:]:
            mask_any = mask_any | m
        df_filtered = df[mask_any].copy()
        if df_filtered.empty:
            return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                                          searched_article=article,
                                          message=f"Aucune ligne ne contient l’article « {article} » dans les colonnes cibles.",
                                          message_ok=True)

        # Vue HTML (listes + surlignage, sans perte)
        df_view = dataframe_for_preview(df_filtered, colmap, pat)

        # Export : brut
        download_url = to_excel_download(df_filtered)

        # Rendu HTML : on n’échappe pas (escape=False) car nous contrôlons l’HTML injecté
        table_html = df_view.head(200).to_html(index=False, escape=False)
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=table_html,
                                      searched_article=article, download_url=download_url,
                                      message=f"{len(df_filtered)} ligne(s) après filtrage. (Aperçu limité à 200 lignes.)",
                                      message_ok=True)

    except Exception as e:
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                                      searched_article=article,
                                      message=f"Erreur inattendue : {repr(e)}", message_ok=False)


@app.route('/download')
def download():
    path = request.args.get('path')
    if not path or not os.path.exists(path):
        return 'Fichier introuvable ou expiré.', 404
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
