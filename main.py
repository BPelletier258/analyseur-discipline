# === CANVAS META =============================================================
# Fichier : main.py — version + motif corrigé + listes + sélection feuille (09-08b)
# Canvas-Stamp : 2025-09-08T20:15Z
# Remarques :
# - Lecture .xlsx/.xlsm avec sélection robuste de la feuille « Cumul décisions »
# - Normalisation des en-têtes (accents/casse/NBSP)
# - Extraction EXACTE de l’article, tolérant « Art: », « Article »…
# - Sortie HTML en listes à puces (avec mise en évidence rouge de l’article)
# - Export Excel sans HTML (puces texte) + largeurs de colonnes auto
# =========================================================================

import io
import os
import re
import time
import html
import unicodedata
from datetime import datetime
from typing import Dict, Optional, Set, List

import pandas as pd
from flask import Flask, request, render_template_string, send_file

app = Flask(__name__)

STYLE_BLOCK = """
<style>
  body { font-family: system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif; margin: 24px; }
  h1 { font-size: 20px; margin-bottom: 12px; }
  form { display: grid; gap: 12px; margin-bottom: 16px; }
  input[type=\"text\"] { padding: 8px; font-size: 14px; }
  input[type=\"file\"] { font-size: 14px; }
  button { padding: 8px 12px; font-size: 14px; cursor: pointer; }
  .hint { font-size: 12px; color: #666; }
  .note { background: #fff6e5; border: 1px solid #ffd89b; padding: 8px 10px; border-radius: 6px; margin: 10px 0 16px; }
  table { border-collapse: collapse; width: 100%; font-size: 13px; }
  th, td { border: 1px solid #ddd; padding: 6px 8px; vertical-align: top; }
  th { background: #f3f4f6; text-align:center; }
  .msg { margin-top: 12px; white-space: pre-wrap; font-family: ui-monospace, SFMono-Regular, Menlo, monospace; font-size: 12px; }
  .ok { color: #065f46; }
  .err { color: #7f1d1d; }
  .download { margin: 12px 0; }
  .kbd { font-family: ui-monospace, SFMono-Regular, Menlo, monospace; background:#f3f4f6; padding:2px 4px; border-radius:4px; }
  /* Zone de tableau visible avec barre horizontale */
  .table-viewport{height:64vh; overflow:auto; border:1px solid #ddd;}
  /* Largeur ≈ 4 écrans pour éviter les colonnes étriquées */
  .table-wide{min-width:400vw;}
  .table-viewport table{width:100%;}
  /* Listes + surlignage de l’article */
  ul.bullets{margin:0; padding-left:1.25rem;}
  ul.bullets li{margin:0 0 .25rem 0;}
  .hit{color:#b91c1c; font-weight:600;}
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

  <div class=\"note\"> 
    Règles : détection exacte de l’article; si la 1<sup>re</sup> cellule contient « <span class=\"kbd\">Article filtré :</span> », elle est ignorée (entêtes sur la 2<sup>e</sup> ligne).
  </div>

  <form method=\"POST\" enctype=\"multipart/form-data\">
    <label>Article à rechercher (ex. <span class=\"kbd\">29</span>, <span class=\"kbd\">59(2)</span>)</label>
    <input type=\"text\" name=\"article\" value=\"{{ searched_article or '' }}\" required placeholder=\"ex.: 29 ou 59(2)\" />
    <label>Fichier Excel</label>
    <input type=\"file\" name=\"file\" accept=\".xlsx,.xlsm\" required />
    <button type=\"submit\">Analyser</button>
    <div class=\"hint\">Formats : .xlsx / .xlsm</div>
  </form>

  {% if table_html %}
    <div class=\"download\">
      <a href=\"{{ download_url }}\">Télécharger le résultat (Excel)</a>
    </div>
    <div class=\"table-viewport\"><div class=\"table-wide\">{{ table_html|safe }}</div></div>
  {% endif %}

  {% if message %}
    <div class=\"msg {{ 'ok' if message_ok else 'err' }}\">{{ message }}</div>
  {% endif %}
</body>
</html>
"""

# ──────────────────────────────────────────────────────────────────────────────
# Normalisation & alias d’en-têtes
# ──────────────────────────────────────────────────────────────────────────────

def _norm(s: str) -> str:
    """Normalise un libellé : accents→ASCII, trim, minuscule, espaces compressés."""
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
}

FILTER_CANONICAL = [
    "articles_enfreints",
    "duree_totale_radiation",
    "article_amende_chef",
    "autres_sanctions",
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
# Lecture Excel : choisir la bonne feuille (priorité « Cumul décisions »)
# ──────────────────────────────────────────────────────────────────────────────

def _pick_sheet_name(xls: pd.ExcelFile) -> str:
    # priorité absolue à « Cumul décisions » (avec/sans accents)
    target_norm = {"cumul decisions", "cumul decision", "cumul-decisions", "cumul-decisions"}
    for name in xls.sheet_names:
        n = _norm(name)
        if n in target_norm:
            return name
    # sinon, heuristique : première feuille dont la 1re cellule contient « Article filtré : »
    for name in xls.sheet_names:
        try:
            preview = pd.read_excel(xls, sheet_name=name, header=None, nrows=2, engine="openpyxl")
            first_cell = preview.iloc[0, 0] if not preview.empty else None
            if isinstance(first_cell, str) and _norm(first_cell).startswith(_norm("Article filtré :")):
                return name
        except Exception:
            continue
    # repli : première feuille
    return xls.sheet_names[0]


def read_excel_respecting_header_rule(file_stream) -> pd.DataFrame:
    # On doit réutiliser le flux après sondage → bufferiser
    raw = file_stream.read()
    bio = io.BytesIO(raw)
    xls = pd.ExcelFile(bio, engine="openpyxl")
    sheet = _pick_sheet_name(xls)

    # règle « Article filtré : »
    df_preview = pd.read_excel(xls, sheet_name=sheet, header=None, nrows=2, engine="openpyxl")
    first_cell = df_preview.iloc[0, 0] if not df_preview.empty else None
    banner = isinstance(first_cell, str) and _norm(first_cell).startswith(_norm("Article filtré :"))

    if banner:
        df = pd.read_excel(xls, sheet_name=sheet, skiprows=1, header=0, engine="openpyxl")
    else:
        df = pd.read_excel(xls, sheet_name=sheet, header=0, engine="openpyxl")
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
    pattern = rf"(?:\b(?:art(?:icle)?\s*[: ]*)?)({esc}){tail}"
    return re.compile(pattern, flags=re.IGNORECASE)

# ──────────────────────────────────────────────────────────────────────────────
# Pré-traitement texte & utilitaires d’extraction
# ──────────────────────────────────────────────────────────────────────────────

def _prep_text_keep_lines(v: str) -> str:
    """Normalise le texte mais *conserve* les retours-lignes pour fabriquer des puces.
    - remplace NBSP/NNBSP par espace
    - unifie CRLF/CR en LF
    - convertit les puces •/·/◦ en retours-lignes
    - compresse les espaces mais garde les \n
    """
    if not isinstance(v, str):
        v = "" if v is None else str(v)
    v = v.replace("\u00A0", " ").replace("\u202F", " ")  # NBSP & NNBSP
    v = v.replace("\r\n", "\n").replace("\r", "\n")
    v = v.replace("•", "\n").replace("·", "\n").replace("◦", "\n")
    # compresse les espaces consécutifs (hors \n)
    v = re.sub(r"[ \t]+", " ", v)
    v = re.sub(r"\n+", "\n", v)
    return v.strip()


def _split_segments(s: str) -> List[str]:
    if not s:
        return []
    segs = re.split(r"[;\n]", s)
    return [seg.strip(" \t-•") for seg in segs if seg.strip(" \t-•")] 


def _highlight(text: str, pat: re.Pattern) -> str:
    def repl(m: re.Match) -> str:
        # remplace uniquement le groupe capturé (le numéro d’article), pas tout « Art: … »
        g1 = m.group(1) if m.lastindex else m.group(0)
        return m.group(0).replace(g1, f'<span class=\"hit\">{html.escape(g1)}</span>')
    return pat.sub(repl, html.escape(text))


def _items_to_html(items: List[str], pat: re.Pattern) -> str:
    if not items:
        return ""
    lis = ''.join(f"<li>{_highlight(it, pat)}</li>" for it in items)
    return f"<ul class='bullets'>{lis}</ul>"


def _items_to_excel_text(items: List[str]) -> str:
    return "\n".join(f"• {it}" for it in items)


# Extraction colonne → listes d’items

def extract_items_generic(text: str, pat: re.Pattern) -> List[str]:
    s = _prep_text_keep_lines(text)
    items = _split_segments(s)
    return [it for it in items if pat.search(it)]


def extract_items_autres_sanctions(text: str, pat: re.Pattern) -> List[str]:
    s = _prep_text_keep_lines(text)
    items = _split_segments(s)
    kept = []
    for it in items:
        if pat.search(it):
            kept.append(it)
    return kept

# ──────────────────────────────────────────────────────────────────────────────
# Nettoyage DataFrame → deux versions : HTML (listes) et Excel (texte)
# ──────────────────────────────────────────────────────────────────────────────

def build_presentations(df: pd.DataFrame, colmap: Dict[str, Optional[str]], pat: re.Pattern):
    df_html = df.copy()
    df_xlsx = df.copy()

    for canon in FILTER_CANONICAL:
        col = colmap.get(canon)
        if not col or col not in df.columns:
            continue
        extractor = extract_items_autres_sanctions if canon == "autres_sanctions" else extract_items_generic
        df_html[col] = df[col].apply(lambda v: _items_to_html(extractor(v, pat), pat))
        df_xlsx[col] = df[col].apply(lambda v: _items_to_excel_text(extractor(v, pat)))

    # filtre lignes vides (après extraction)
    subset_cols = [c for c in (colmap.get(k) for k in FILTER_CANONICAL) if c]
    if subset_cols:
        mask_any = False
        for c in subset_cols:
            cur = df_xlsx[c].astype(str).str.strip().ne("")
            mask_any = cur if mask_any is False else (mask_any | cur)
        df_html = df_html[mask_any]
        df_xlsx = df_xlsx[mask_any]

    return df_html, df_xlsx

# ──────────────────────────────────────────────────────────────────────────────
# Export Excel
# ──────────────────────────────────────────────────────────────────────────────

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
            HTML_TEMPLATE,
            style_block=STYLE_BLOCK,
            table_html=None,
            searched_article=article,
            message="Erreur : fichier et article sont requis.",
            message_ok=False
        )

    # Validation : uniquement .xlsx / .xlsm (openpyxl)
    fname = (file.filename or "").lower()
    if not (fname.endswith(".xlsx") or fname.endswith(".xlsm")):
        return render_template_string(
            HTML_TEMPLATE,
            style_block=STYLE_BLOCK,
            table_html=None,
            searched_article=article,
            message=(
                "Format non pris en charge : " + (file.filename or "").split(".")[-1] + ". "
                "Veuillez fournir un classeur Excel .xlsx ou .xlsm. Les fichiers .xls (Excel 97-2003) ne sont pas supportés."
            ),
            message_ok=False
        )

    try:
        df = read_excel_respecting_header_rule(file.stream)
        colmap = resolve_columns(df)
        pat = build_article_pattern(article)

        # préfiltrage rapide (pour performance) :
        masks = []
        any_cols = False
        for canon in FILTER_CANONICAL:
            col = colmap.get(canon)
            if col and col in df.columns:
                any_cols = True
                masks.append(df[col].astype(str).apply(lambda v: bool(pat.search(_prep_text_keep_lines(v)))))
        if not any_cols:
            detail = "\n".join([f"  - {k}: {colmap.get(k)}" for k in FILTER_CANONICAL])
            return render_template_string(
                HTML_TEMPLATE,
                style_block=STYLE_BLOCK,
                table_html=None,
                searched_article=article,
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
                HTML_TEMPLATE,
                style_block=STYLE_BLOCK,
                table_html=None,
                searched_article=article,
                message=f"Aucune ligne ne contient l’article « {article} » dans les colonnes cibles.",
                message_ok=True
            )

        # Fabrique deux présentations : HTML (listes + <span class='hit'>) et Excel (texte puces)
        df_html, df_xlsx = build_presentations(df_filtered, colmap, pat)

        if df_html.empty:
            return render_template_string(
                HTML_TEMPLATE,
                style_block=STYLE_BLOCK,
                table_html=None,
                searched_article=article,
                message=("Des lignes correspondaient au motif, mais après épuration des cellules, "
                         "aucune mention nette de l’article n’a été conservée."),
                message_ok=True
            )

        download_url = to_excel_download(df_xlsx)
        table_html = df_html.head(200).to_html(index=False, escape=False)

        return render_template_string(
            HTML_TEMPLATE,
            style_block=STYLE_BLOCK,
            table_html=table_html,
            searched_article=article,
            download_url=download_url,
            message=f"{len(df_html)} ligne(s) après filtrage et épuration. (Aperçu limité à 200 lignes.)",
            message_ok=True
        )

    except Exception as e:
        return render_template_string(
            HTML_TEMPLATE,
            style_block=STYLE_BLOCK,
            table_html=None,
            searched_article=article,
            message=f"Erreur inattendue : {repr(e)}",
            message_ok=False
        )


@app.route("/download")
def download():
    path = request.args.get("path")
    if not path or not os.path.exists(path):
        return "Fichier introuvable ou expiré.", 404
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
