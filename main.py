# === main.py — 2025-09-09 ===========================================
# - En-têtes centrés
# - Puce uniquement si contenu, sinon tiret — (sans puce)
# - Option "isoler segments" via case à cocher
# - Mise en évidence (rouge) de l'article dans 5 colonnes
# - Export Excel propre (sans HTML)
# ====================================================================

import os
import re
import time
import unicodedata
from typing import Dict, Optional, Set, List

import pandas as pd
from flask import Flask, request, render_template_string, send_file
from html import escape as html_escape

app = Flask(__name__)

# ─────────────────────────────────────────────────────────────────────
# Styles & Page
# ─────────────────────────────────────────────────────────────────────

STYLE_BLOCK = """
<style>
  :root { --hit:#c62828; }
  body { font-family: system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif; margin: 24px; }
  h1 { font-size: 20px; margin-bottom: 12px; }
  form { display: grid; gap: 10px; margin-bottom: 14px; align-items: start; }
  input[type="text"] { padding: 8px; font-size: 14px; }
  input[type="file"] { font-size: 14px; }
  button { padding: 8px 12px; font-size: 14px; cursor: pointer; }
  label.inline { display:flex; gap:.55rem; align-items:center; font-size: 14px; }
  .hint { font-size: 12px; color: #666; }
  .note { background: #fff6e5; border: 1px solid #ffd89b; padding: 8px 10px; border-radius: 6px; margin: 10px 0 14px; }
  .download { margin: 12px 0; }

  .table-viewport{height:60vh; overflow:auto; border:1px solid #ddd; border-radius:6px;}
  .table-wide{min-width:200vw;}
  table { border-collapse: collapse; width: 100%; font-size: 13px; }
  th, td { border: 1px solid #ddd; padding: 6px 8px; vertical-align: top; }
  th { background: #f3f4f6; text-align:center; }
  .msg { margin-top: 12px; white-space: pre-wrap; font-family: ui-monospace, SFMono-Regular, Menlo, monospace; font-size: 12px; }
  .ok { color: #065f46; }
  .err { color: #7f1d1d; }

  /* contenu des cellules listées */
  td ul { margin: 0; padding-left: 1.1rem; }
  td li { margin: .12rem 0; }
  .hit { color: var(--hit); font-weight: 600; }
  .kbd { font-family: ui-monospace, SFMono-Regular, Menlo, monospace; background:#f3f4f6; padding:2px 4px; border-radius:4px; }
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
    Règles : détection exacte de l’article; si la 1<sup>re</sup> cellule contient
    « <span class="kbd">Article filtré :</span> », elle est ignorée (entêtes sur la 2<sup>e</sup> ligne).
  </div>

  <form method="POST" enctype="multipart/form-data">
    <label>Article à rechercher (ex. <span class="kbd">29</span>, <span class="kbd">59(2)</span>)</label>
    <input type="text" name="article" value="{{ searched_article or '' }}" required placeholder="ex.: 29 ou 59(2)" />
    <label>Fichier Excel</label>
    <input type="file" name="file" accept=".xlsx,.xlsm" required />
    <label class="inline">
      <input type="checkbox" name="isolate" value="1" {% if isolate %}checked{% endif %}>
      Afficher seulement les segments contenant l’article (dans
      les 4 colonnes d’intérêt + « Liste des chefs et articles en infraction »)
    </label>
    <button type="submit">Analyser</button>
    <div class="hint">Formats : .xlsx / .xlsm</div>
  </form>

  {% if table_html %}
    <div class="download"><a href="{{ download_url }}">Télécharger le résultat (Excel)</a></div>
    <div class="table-viewport"><div class="table-wide">{{ table_html|safe }}</div></div>
  {% endif %}

  {% if message %}
    <div class="msg {{ 'ok' if message_ok else 'err' }}">{{ message }}</div>
  {% endif %}
</body>
</html>
"""

# ─────────────────────────────────────────────────────────────────────
# Outils de normalisation / entêtes
# ─────────────────────────────────────────────────────────────────────

def _norm(s: str) -> str:
    """Accent-insensitive, trim, lower, compaction."""
    if not isinstance(s, str):
        s = "" if s is None else str(s)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = s.replace("\u00A0", " ")
    return " ".join(s.strip().lower().split())

# Aliases d’en-têtes
HEADER_ALIASES: Dict[str, Set[str]] = {
    # Canon : articles_enfreints  (ex-"Articles enfreints" / "Liste des chefs et articles en infraction")
    "articles_enfreints": {
        _norm("Nbr Chefs par articles"),
        _norm("Articles enfreints"),
        _norm("Articles en infraction"),
        _norm("Liste des chefs et articles en infraction"),
    },
    # Canon : duree_totale_radiation
    "duree_totale_radiation": {
        _norm("Nbr Chefs par articles par période de radiation"),
        _norm("Nbr Chefs par articles par periode de radiation"),
        _norm("Durée totale effective radiation"),
        _norm("Duree totale effective radiation"),
    },
    # Canon : article_amende_chef
    "article_amende_chef": {
        _norm("Nombre de chefs par articles et total amendes"),
        _norm("Article amende/chef"),
        _norm("Articles amende / chef"),
        _norm("Amendes (article/chef)"),
    },
    # Canon : autres_sanctions  (ici = réprimandes)
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
    """Retourne {canonique: libellé_original}"""
    norm_to_orig = {_norm(c): c for c in df.columns}
    out: Dict[str, Optional[str]] = {}
    for canon, variants in HEADER_ALIASES.items():
        out[canon] = next((norm_to_orig[v] for v in variants if v in norm_to_orig), None)
    return out

# ─────────────────────────────────────────────────────────────────────
# Lecture Excel (bannière « Article filtré : » en première ligne)
# ─────────────────────────────────────────────────────────────────────

def read_excel_respecting_header_rule(fstream) -> pd.DataFrame:
    preview = pd.read_excel(fstream, header=None, nrows=2, engine="openpyxl")
    fstream.seek(0)
    first = preview.iloc[0, 0] if not preview.empty else None
    banner = isinstance(first, str) and _norm(first).startswith(_norm("Article filtré :"))
    if banner:
        return pd.read_excel(fstream, skiprows=1, header=0, engine="openpyxl")
    return pd.read_excel(fstream, header=0, engine="openpyxl")

# ─────────────────────────────────────────────────────────────────────
# Recherche article (motif exact)
# ─────────────────────────────────────────────────────────────────────

def build_article_pattern(token: str) -> re.Pattern:
    token = (token or "").strip()
    if not token:
        raise ValueError("Article vide.")
    esc = re.escape(token)
    ends_with_digit = token[-1].isdigit()
    tail = r"(?![\d.])" if ends_with_digit else r"\b"
    # capture du jeton en groupe(1) pour pouvoir le surligner
    return re.compile(rf"(?:\b(?:art(?:icle)?\s*[: ]*)?)({esc}){tail}", re.IGNORECASE)

def boundary_token_regex(token: str) -> re.Pattern:
    """Pour le surlignage dans le HTML (plus simple, juste le jeton entouré de non-chiffres)."""
    return re.compile(rf"(?<!\d){re.escape(token)}(?!\d)")

# ─────────────────────────────────────────────────────────────────────
# Pré-traitement / segmentation
# ─────────────────────────────────────────────────────────────────────

def _prep_text(v) -> str:
    if not isinstance(v, str):
        v = "" if v is None else str(v)
    # uniformisation espaces & sauts
    v = v.replace(" ", " ").replace(" ", " ")
    v = v.replace("\r\n", "\n").replace("\r", "\n")
    v = " ".join(v.split())
    return v

def _split_any(text: str) -> List[str]:
    """Segmente un texte brut en items (puces, sauts de ligne, ';')."""
    if not isinstance(text, str) or not text.strip():
        return []
    t = text.replace("\r\n", "\n").replace("\r", "\n")
    for ch in ("•", "·", "◦"):
        t = t.replace(ch, "\n")
    t = t.replace(";"," \n ")
    parts = [p.strip() for p in t.split("\n")]
    return [p for p in parts if p]

# ─────────────────────────────────────────────────────────────────────
# Extraction “segments contenant l’article”
# ─────────────────────────────────────────────────────────────────────

def extract_segments(text: str, pat: re.Pattern) -> List[str]:
    """Retourne uniquement les segments contenant l'article (via ; ou sauts/puces)."""
    if not isinstance(text, str) or not text.strip():
        return []
    # on segmente « généreusement »
    parts = _split_any(text)
    return [p for p in parts if pat.search(_prep_text(p))]

# ─────────────────────────────────────────────────────────────────────
# Rendu HTML des cellules (liste ou tiret), avec surlignage
# ─────────────────────────────────────────────────────────────────────

EMPTY_MARK = "—"

def _highlight_html(text: str, token_re: re.Pattern) -> str:
    """Sécurise en HTML puis surligne le jeton."""
    esc = html_escape(text)
    return token_re.sub(lambda m: f'<span class="hit">{m.group(0)}</span>', esc)

def render_list_html(items: List[str], token_re: Optional[re.Pattern]) -> str:
    if not items:
        return EMPTY_MARK
    if token_re:
        items = [_highlight_html(x, token_re) for x in items]
    else:
        items = [html_escape(x) for x in items]
    return "<ul>" + "".join(f"<li>{x}</li>" for x in items) + "</ul>"

# ─────────────────────────────────────────────────────────────────────
# Export Excel utilitaire
# ─────────────────────────────────────────────────────────────────────

def to_excel_download(df: pd.DataFrame) -> str:
    ts = int(time.time())
    out_path = f"/tmp/filtrage_{ts}.xlsx"
    # valeurs NaN → chaîne vide pour Excel
    df_out = df.copy()
    df_out = df_out.fillna("")
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df_out.to_excel(writer, index=False, sheet_name="Filtre")
        ws = writer.book.active
        # largeur de colonnes approximative
        for col_idx, col in enumerate(df_out.columns, start=1):
            max_len = max((len(str(x)) for x in [col] + df_out[col].astype(str).tolist()), default=10)
            ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = min(64, max(12, max_len + 2))
    return f"/download?path={out_path}"

# ─────────────────────────────────────────────────────────────────────
# Route principale
# ─────────────────────────────────────────────────────────────────────

@app.route("/", methods=["GET", "POST"])
def analyze():
    if request.method == "GET":
        return render_template_string(
            HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
            searched_article=None, isolate=False, message=None, message_ok=True
        )

    file = request.files.get("file")
    article = (request.form.get("article") or "").strip()
    isolate = bool(request.form.get("isolate"))

    if not file or not article:
        return render_template_string(
            HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
            searched_article=article, isolate=isolate,
            message="Erreur : fichier et article sont requis.", message_ok=False
        )

    fname = (file.filename or "").lower()
    if not (fname.endswith(".xlsx") or fname.endswith(".xlsm")):
        return render_template_string(
            HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
            searched_article=article, isolate=isolate,
            message=("Format non pris en charge. Veuillez fournir .xlsx ou .xlsm."), message_ok=False
        )

    try:
        df = read_excel_respecting_header_rule(file.stream)
        colmap = resolve_columns(df)
        pat = build_article_pattern(article)
        token_re = boundary_token_regex(article)

        # 1) Filtrage des lignes : l’article doit apparaître dans ≥1 des 4 colonnes cibles
        masks = []
        any_cols = False
        for canon in FILTER_CANONICAL:
            col = colmap.get(canon)
            if col and col in df.columns:
                any_cols = True
                masks.append(df[col].astype(str).apply(lambda v: bool(pat.search(_prep_text(v)))))
        if not any_cols:
            detail = "\n".join([f"- {k}: {colmap.get(k)}" for k in FILTER_CANONICAL])
            return render_template_string(
                HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                searched_article=article, isolate=isolate,
                message=("Erreur : colonnes cibles introuvables.\n" + detail + "\nColonnes du fichier :\n" + str(list(df.columns))),
                message_ok=False
            )

        mask_any = masks[0]
        for m in masks[1:]:
            mask_any = mask_any | m

        df_filtered = df[mask_any].copy()
        if df_filtered.empty:
            return render_template_string(
                HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                searched_article=article, isolate=isolate,
                message=f"Aucune ligne ne contient l’article « {article} » dans les colonnes cibles.",
                message_ok=True
            )

        # 2) Préparation des colonnes à rendre "en liste"
        #    (ajoute les colonnes textuelles affichées en listes)
        list_titles_extra = [
            "Résumé des faits concis",
            "Liste des chefs et articles en infraction",
            "Liste des sanctions imposées",
            "Autres mesures ordonnées",
            "À vérifier",
        ]
        list_cols: List[str] = []
        for canon in FILTER_CANONICAL:
            col = colmap.get(canon)
            if col and col in df_filtered.columns:
                list_cols.append(col)
        for t in list_titles_extra:
            if t in df_filtered.columns and t not in list_cols:
                list_cols.append(t)

        # 3) Jeu de données destiné à l'export (pas de HTML)
        #    - si "isolate", on remplace le contenu des 4 colonnes cibles
        #      + « Liste des chefs et articles en infraction » par les *segments* contenant l'article.
        df_export = df_filtered.copy()

        isolate_targets = set()
        for canon in FILTER_CANONICAL:
            col = colmap.get(canon)
            if col and col in df_export.columns:
                isolate_targets.add(col)
        # assurer que la colonne "Liste des chefs et articles en infraction" est incluse si elle existe
        if "Liste des chefs et articles en infraction" in df_export.columns:
            isolate_targets.add("Liste des chefs et articles en infraction")

        if isolate:
            for col in isolate_targets:
                df_export[col] = df_export[col].apply(
                    lambda v: " | ".join(extract_segments(v, pat)) if isinstance(v, str) else ""
                )

        # 4) Jeu de données destiné à l'HTML (liste à puces + surlignage)
        df_html = df_export.copy()

        # colonnes à surligner (4 cibles + « Liste des chefs et articles en infraction »)
        highlight_cols = set(isolate_targets)

        def _format_cell(value, column_name):
            # transforme la cellule en <ul>…</ul> si items, sinon tiret —
            # segmentation : si "isolate" a été appliqué à l'export, la valeur contient " | "
            # sinon on segmente à partir du texte brut.
            if column_name in list_cols:
                if isinstance(value, str) and " | " in value:
                    items = [x.strip() for x in value.split("|") if x.strip()]
                else:
                    items = _split_any("" if value is None else str(value))
                token = token_re if column_name in highlight_cols else None
                return render_list_html(items, token)
            # colonnes non listées : afficher brut échappé (ou tiret si vide)
            s = "" if value is None else str(value).strip()
            return html_escape(s) if s else EMPTY_MARK

        for c in list_cols:
            if c in df_html.columns:
                df_html[c] = df_html[c].map(lambda v, col=c: _format_cell(v, col))

        # 5) Export Excel (texte sans HTML)
        download_url = to_excel_download(df_export)

        # 6) Tableau HTML
        preview = df_html.head(200)
        table_html = preview.to_html(index=False, escape=False)

        return render_template_string(
            HTML_TEMPLATE, style_block=STYLE_BLOCK,
            table_html=table_html, searched_article=article, isolate=isolate,
            download_url=download_url,
            message=f"{len(df_export)} ligne(s) retenues. (Aperçu limité à 200.)",
            message_ok=True
        )

    except Exception as e:
        return render_template_string(
            HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
            searched_article=article, isolate=isolate,
            message=f"Erreur inattendue : {repr(e)}", message_ok=False
        )

# ─────────────────────────────────────────────────────────────────────
# Téléchargement
# ─────────────────────────────────────────────────────────────────────

@app.route("/download")
def download():
    path = request.args.get("path")
    if not path or not os.path.exists(path):
        return "Fichier introuvable ou expiré.", 404
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
