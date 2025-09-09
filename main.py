# === CANVAS META =============================================================
# Fichier : main.py — version + motif corrigé + diagnostics (09-08)
# Canvas-Stamp : 2025-09-08T15:45Z
# Cible déployée (SHA court, cf. pied de page / /version) : b4ebb8e
# ============================================================================
# main.py (MAJ titres de colonnes)
# - Normalisation des en-têtes (accents, casse, espaces insécables)
# - Aliases mis à jour pour les NOUVEAUX titres demandés
# - Règle « Article filtré : » tolérante
# - Recherche EXACTE de l’article
# - Extraction **sans tronquer** + mise en forme en listes
# - Mise en évidence de l’article en rouge dans 4 colonnes d’intérêt

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
  table { border-collapse: collapse; width: 100%; font-size: 13px; table-layout: fixed; }
  th, td { border: 1px solid #ddd; padding: 6px 8px; vertical-align: top; }
  th { background: #f3f4f6; }
  td ul{ margin: 0; padding-left: 18px; }
  td li{ margin: 0 0 2px 0; }
  .hit { color: #c1121f; font-weight: 600; } /* article surligné en rouge */
  .msg { margin-top: 12px; white-space: pre-wrap; font-family: ui-monospace, SFMono-Regular, Menlo, monospace; font-size: 12px; }
  .ok { color: #065f46; }
  .err { color: #7f1d1d; }
  .download { margin: 12px 0; }
  .kbd { font-family: ui-monospace, SFMono-Regular, Menlo, monospace; background:#f3f4f6; padding:2px 4px; border-radius:4px; }
  /* Zone scrollable (barre horizontale visible) */
  .table-viewport{height:60vh; overflow:auto; border:1px solid #ddd;}
  /* Largeur ~ 2 écrans pour éviter l’écrasement */
  .table-wide{min-width:200vw;}
  .table-viewport table{width:100%;}
</style>
"""

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
    Règles : détection exacte de l’article; si la 1<sup>re</sup> cellule contient « <span class="kbd">Article filtré :</span> », elle est ignorée (entêtes sur la 2<sup>e</sup> ligne).
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

# ──────────────────────────────────────────────────────────────────────────────
# Normalisation & alias d’en-têtes
# ──────────────────────────────────────────────────────────────────────────────

def _norm(s: str) -> str:
    if not isinstance(s, str):
        s = str(s) if s is not None else ""
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = s.replace("\u00A0", " ")  # NBSP
    s = " ".join(s.strip().lower().split())
    return s

HEADER_ALIASES: Dict[str, Set[str]] = {
    "articles_enfreints": {
        _norm("Nbr Chefs par articles"),
        _norm("Articles enfreints"),
        _norm("Articles en infraction"),
        _norm("Liste des chefs et articles en infraction"),
        _norm("Nbr   Chefs   par    articles"),
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
        _norm("Articles amende / chef"),
        _norm("Amendes (article/chef)"),
    },
    "autres_sanctions": {
        _norm("Nombre de chefs par article ayant une réprimande"),
        _norm("Nombre de chefs par article ayant une reprimande"),
        _norm("Autres sanctions"),
        _norm("Autres mesures ordonnées"),
        _norm("Autres sanctions / mesures"),
    },
    "nbr_chefs_par_articles": {
        _norm("Nbr Chefs par articles"),
        _norm("Nombre de chefs par articles"),
    },
}

# colonnes dont on forcera l’affichage en listes HTML (en plus des 4 d’intérêt)
ALWAYS_UL = {
    _norm("Résumé des faits concis"),
    _norm("Liste des sanctions imposées"),
    _norm("Autres mesures ordonnées"),
    _norm("À vérifier"),
    _norm("A verifier"),
}

FILTER_CANONICAL = [
    "articles_enfreints",          # Nbr Chefs par articles
    "duree_totale_radiation",      # Nbr Chefs par articles par période de radiation
    "article_amende_chef",         # Nombre de chefs par articles et total amendes
    "autres_sanctions",            # Nombre de chefs par article ayant une réprimande
]

def resolve_columns(df: pd.DataFrame) -> Dict[str, Optional[str]]:
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

# ──────────────────────────────────────────────────────────────────────────────
# Lecture Excel (gestion « Article filtré : »)
# ──────────────────────────────────────────────────────────────────────────────

def read_excel_respecting_header_rule(file_stream) -> pd.DataFrame:
    df_preview = pd.read_excel(file_stream, header=None, nrows=2, engine="openpyxl")
    file_stream.seek(0)
    first_cell = df_preview.iloc[0, 0] if not df_preview.empty else None
    is_first_row_banner = False
    if isinstance(first_cell, str):
        if _norm(first_cell).startswith(_norm("Article filtré :")):
            is_first_row_banner = True
    if is_first_row_banner:
        df = pd.read_excel(file_stream, skiprows=1, header=0, engine="openpyxl")
    else:
        df = pd.read_excel(file_stream, header=0, engine="openpyxl")
    return df

# ──────────────────────────────────────────────────────────────────────────────
# Motif exact pour l’article
# ──────────────────────────────────────────────────────────────────────────────

def build_article_pattern(user_input: str) -> re.Pattern:
    token = (user_input or "").strip()
    if not token:
        raise ValueError("Article vide.")
    esc = re.escape(token)
    ends_with_digit = token[-1].isdigit()
    tail = r"(?![\d.])" if ends_with_digit else r"\b"
    return re.compile(rf"(?:\b(?:art(?:icle)?\s*[: ]*)?)({esc}){tail}", re.IGNORECASE)

# ──────────────────────────────────────────────────────────────────────────────
# Pré-traitement & rendu HTML
# ──────────────────────────────────────────────────────────────────────────────

def _prep_text(v: str) -> str:
    """Normalise le texte tout en CONSERVANT la structure en lignes."""
    if not isinstance(v, str):
        v = "" if v is None else str(v)
    # puces -> retour ligne
    v = (v.replace("•", "\n").replace("·", "\n").replace("◦", "\n"))
    # normalisation espaces / sauts de ligne
    v = v.replace("\u00A0", " ").replace("\u202F", " ")
    v = v.replace("\r\n", "\n").replace("\r", "\n")
    # ; comme séparateur fort -> retour ligne
    v = re.sub(r"\s*;\s*", "\n", v)
    # comprime les espaces (mais pas les \n)
    v = re.sub(r"[ \t]+", " ", v)
    # nettoie lignes vides / 'nan'
    lines = [ln.strip(" .") for ln in v.split("\n") if ln and ln.strip().lower() != "nan"]
    return "\n".join(ln for ln in lines if ln)

def _as_ul(text: str) -> str:
    """Transforme du texte multi-lignes en <ul><li>…</li></ul> (ou chaîne vide)."""
    if not text:
        return ""
    lines = [ln.strip() for ln in text.split("\n") if ln.strip()]
    if not lines:
        return ""
    lis = "".join(f"<li>{ln}</li>" for ln in lines)
    return f"<ul>{lis}</ul>"

def _highlight(text: str, pat: re.Pattern) -> str:
    if not text:
        return ""
    return pat.sub(lambda m: f"<span class='hit'>{m.group(1)}</span>", text)

# ──────────────────────────────────────────────────────────────────────────────
# Mise en forme colonnes (sans tronquer le contenu)
# ──────────────────────────────────────────────────────────────────────────────

def format_interest_columns(df: pd.DataFrame, colmap: Dict[str, Optional[str]], pat: re.Pattern) -> pd.DataFrame:
    """Pour les 4 colonnes d’intérêt: nettoie, **met en liste** et surligne l’article."""
    df = df.copy()
    for canon in FILTER_CANONICAL:
        col = colmap.get(canon)
        if not col or col not in df.columns:
            continue
        df[col] = df[col].fillna("").astype(str).map(_prep_text).map(lambda s: _as_ul(_highlight(s, pat)))
    return df

def format_misc_lists(df: pd.DataFrame) -> pd.DataFrame:
    """Bulletise d’autres colonnes utiles (résumé, sanctions, autres mesures, à vérifier…)."""
    df = df.copy()
    for col in df.columns:
        if _norm(col) in ALWAYS_UL:
            df[col] = df[col].fillna("").astype(str).map(_prep_text).map(_as_ul)
    return df

# ──────────────────────────────────────────────────────────────────────────────
# Export Excel (même contenu que l’aperçu HTML, sans HTML)
# ──────────────────────────────────────────────────────────────────────────────

def to_excel_download(df: pd.DataFrame) -> str:
    ts = int(time.time())
    out_path = f"/tmp/filtrage_{ts}.xlsx"
    # pour Excel : on retire balises, on garde le texte brut multi-lignes
    df_x = df.copy()
    for c in df_x.columns:
        df_x[c] = (
            df_x[c]
            .astype(str)
            .str.replace(r"<span class='hit'>", "", regex=False)
            .str.replace(r"</span>", "", regex=False)
            .str.replace(r"</?ul>", "", regex=True)
            .str.replace(r"</?li>", "\n", regex=True)
            .str.replace(r"\n{2,}", "\n", regex=True)
            .str.strip()
        )
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df_x.to_excel(writer, index=False, sheet_name="Filtre")
        ws = writer.book.active
        for col_idx, col in enumerate(df_x.columns, start=1):
            max_len = max((len(str(x)) for x in [col] + df_x[col].astype(str).tolist()), default=10)
            ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = min(60, max(12, max_len + 2))
    return f"/download?path={out_path}"

# ──────────────────────────────────────────────────────────────────────────────
# Routes
# ──────────────────────────────────────────────────────────────────────────────

@app.route("/", methods=["GET", "POST"])
def analyze():
    if request.method == "GET":
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                                      searched_article=None, message=None, message_ok=True)

    file = request.files.get("file")
    article = (request.form.get("article") or "").strip()

    if not file or not article:
        return render_template_string(
            HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None, searched_article=article,
            message="Erreur : fichier et article sont requis.", message_ok=False
        )

    fname = (file.filename or "").lower()
    if not (fname.endswith(".xlsx") or fname.endswith(".xlsm")):
        return render_template_string(
            HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None, searched_article=article,
            message=("Format non pris en charge : " + (file.filename or "").split(".")[-1] +
                     ". Veuillez fournir .xlsx ou .xlsm."), message_ok=False
        )

    try:
        df = read_excel_respecting_header_rule(file.stream)
        colmap = resolve_columns(df)
        pat = build_article_pattern(article)

        # Filtrage des lignes: article présent dans AU MOINS une colonne d’intérêt
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
                HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None, searched_article=article,
                message=("Erreur : aucune des colonnes attendues n’a été trouvée dans le fichier.\n"
                         "Vérifiez les en-têtes ou ajoutez des alias dans le code.\n\n"
                         f"Colonnes résolues :\n{detail}\n\nColonnes disponibles :\n{list(df.columns)}"),
                message_ok=False
            )
        mask_any = masks[0]
        for m in masks[1:]:
            mask_any = mask_any | m
        df_filtered = df[mask_any].copy()
        if df_filtered.empty:
            return render_template_string(
                HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None, searched_article=article,
                message=f"Aucune ligne ne contient l’article « {article} » dans les colonnes cibles.", message_ok=True
            )

        # Mise en forme des colonnes (sans supprimer de contenu)
        df_fmt = format_interest_columns(df_filtered, colmap, pat)
        df_fmt = format_misc_lists(df_fmt)

        # Remplace NaN résiduels par chaînes vides (évite « nan » affiché)
        df_fmt = df_fmt.fillna("")

        download_url = to_excel_download(df_fmt)
        preview = df_fmt.head(200)
        table_html = preview.to_html(index=False, escape=False)

        return render_template_string(
            HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=table_html, searched_article=article,
            download_url=download_url,
            message=f"{len(df_fmt)} ligne(s) après filtrage. (Aperçu limité à 200 lignes.)", message_ok=True
        )

    except Exception as e:
        return render_template_string(
            HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None, searched_article=article,
            message=f"Erreur inattendue : {repr(e)}", message_ok=False
        )

@app.route("/download")
def download():
    path = request.args.get("path")
    if not path or not os.path.exists(path):
        return "Fichier introuvable ou expiré.", 404
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
