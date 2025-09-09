# === CANVAS META =============================================================
# Fichier : main.py — autosheet + alias ordonnés + puces + surlignage + diagnostics
# Stamp : 2025-09-08
# ============================================================================

import io
import os
import re
import time
import unicodedata
from datetime import datetime
from typing import Dict, List

import pandas as pd
from flask import Flask, request, render_template_string, send_file, jsonify

# ----------------------------------------------------------------------------
# App & estampille de version
# ----------------------------------------------------------------------------
app = Flask(__name__)

STARTED_AT = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC")
APP_VERSION = (
    os.environ.get("RENDER_GIT_COMMIT")
    or os.environ.get("GIT_COMMIT")
    or os.environ.get("SOURCE_VERSION")
    or datetime.utcnow().strftime("dev-%Y%m%d-%H%M%S")
)
APP_VERSION_SHORT = (APP_VERSION or "")[:7]

@app.context_processor
def inject_globals():
    return dict(app_version=APP_VERSION, app_version_short=APP_VERSION_SHORT, started_at=STARTED_AT)

# ----------------------------------------------------------------------------
# UI (template + styles)
# ----------------------------------------------------------------------------
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
  th { background: #f3f4f6; text-align: center; position: sticky; top: 0; }

  .msg { margin-top: 12px; white-space: pre-wrap; font-family: ui-monospace, SFMono-Regular, Menlo, monospace; font-size: 12px; }
  .ok { color: #065f46; } .err { color: #7f1d1d; }
  .download { margin: 12px 0; }
  .kbd { font-family: ui-monospace, SFMono-Regular, Menlo, monospace; background:#f3f4f6; padding:2px 4px; border-radius:4px; }

  /* Zone de tableau visible avec barre horizontale */
  .table-viewport { height: 60vh; overflow: auto; border: 1px solid #ddd; }
  .table-wide { min-width: 200vw; }  /* ≈ deux écrans de large */
  .table-viewport table { width: 100%; }

  /* Listes dans les cellules */
  ul.list { margin: 0; padding-left: 18px; }
  ul.list li { margin: 2px 0; }

  /* Surlignage de l'article */
  .hit { color: #b91c1c; font-weight: 600; }

  footer { margin-top: 18px; font-size: 12px; color: #6b7280; }
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

  <footer>
    Version: <strong>{{ app_version_short }}</strong> ({{ app_version }}) • Démarré: {{ started_at }} • <a href=\"/version\">/version</a> • <a href=\"/health\">/health</a>
  </footer>
</body>
</html>
"""

# ----------------------------------------------------------------------------
# Normalisation & alias d’en‑têtes (ORDONNÉS)
# ----------------------------------------------------------------------------

def _norm(s: str) -> str:
    """Normalise un libellé : accents→ASCII, trim, minuscule, espaces compressés."""
    if not isinstance(s, str):
        s = str(s) if s is not None else ""
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = s.replace("\u00A0", " ").replace("\u202F", " ")
    s = " ".join(s.strip().lower().split())
    return s

# Alias ordonnés (priorité aux NOUVEAUX titres)
HEADER_ALIASES: Dict[str, List[str]] = {
    "articles_enfreints": [
        _norm("Nbr Chefs par articles"),
        _norm("Articles enfreints"),
        _norm("Articles en infraction"),
        _norm("Liste des chefs et articles en infraction"),
    ],
    "duree_totale_radiation": [
        _norm("Nbr Chefs par articles par période de radiation"),
        _norm("Nbr Chefs par articles par periode de radiation"),
        _norm("Durée totale effective radiation"),
        _norm("Duree totale effective radiation"),
    ],
    "article_amende_chef": [
        _norm("Nombre de chefs par articles et total amendes"),
        _norm("Article amende/chef"),
    ],
    "autres_sanctions": [
        _norm("Nombre de chefs par article ayant une réprimande"),
        _norm("Nombre de chefs par article ayant une reprimande"),
        _norm("Autres sanctions"),
    ],
}

FILTER_CANONICAL = [
    "articles_enfreints",
    "duree_totale_radiation",
    "article_amende_chef",
    "autres_sanctions",
]


def resolve_columns(df: pd.DataFrame) -> Dict[str, List[str]]:
    """Retourne, pour chaque canonique, la LISTE des colonnes présentes (ordre prioritaire)."""
    norm_to_original = {_norm(c): c for c in df.columns}
    resolved: Dict[str, List[str]] = {}
    for canon, ordered_aliases in HEADER_ALIASES.items():
        hits: List[str] = []
        for a in ordered_aliases:
            if a in norm_to_original:
                hits.append(norm_to_original[a])
        resolved[canon] = hits
    return resolved

# ----------------------------------------------------------------------------
# Lecture Excel – auto‑sélection de la meilleure feuille + bannière
# ----------------------------------------------------------------------------

def read_best_sheet(file_bytes: bytes) -> (pd.DataFrame, str):
    xls = pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl")
    best_df = None
    best_sheet = None
    best_score = -1
    for sheet in xls.sheet_names:
        prev = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet, header=None, nrows=2, engine="openpyxl")
        first_cell = prev.iloc[0, 0] if not prev.empty else None
        banner = isinstance(first_cell, str) and _norm(first_cell).startswith(_norm("Article filtré :"))
        if banner:
            df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet, skiprows=1, header=0, engine="openpyxl")
        else:
            df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet, header=0, engine="openpyxl")
        colmap = resolve_columns(df)
        score = sum(1 for k in FILTER_CANONICAL if colmap.get(k))
        if score > best_score:
            best_df, best_sheet, best_score = df, sheet, score
    if best_df is None:
        best_df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0, header=0, engine="openpyxl")
        best_sheet = xls.sheet_names[0]
    return best_df, best_sheet

# ----------------------------------------------------------------------------
# Motif exact pour l’article — borne droite corrigée (59(2), 2.01 a))
# ----------------------------------------------------------------------------

def build_article_pattern(user_input: str) -> re.Pattern:
    token = (user_input or "").strip()
    if not token:
        raise ValueError("Article vide.")
    token = token.replace("\u00A0", " ").replace("\u202F", " ")
    token = " ".join(token.split())
    esc = re.escape(token)
    left_guard = r"(?<![\d.])"  # évite 2016→16
    last = token[-1]
    right_tail = r"(?![\d.])" if last.isdigit() else r"(?![A-Za-z0-9])"
    return re.compile(rf"(?:\\bart(?:icle)?\\s*[: ]*)?{left_guard}({esc}){right_tail}", re.IGNORECASE)

# ----------------------------------------------------------------------------
# Pré‑traitement texte (puces, NBSP, CR/LF)
# ----------------------------------------------------------------------------

def _prep_text(v: str) -> str:
    if not isinstance(v, str):
        v = "" if v is None else str(v)
    v = v.replace("•", " ").replace("·", " ").replace("◦", " ")
    v = v.replace("\u00A0", " ").replace("\u202F", " ")
    v = v.replace("\r\n", "\n").replace("\r", "\n")
    v = " ".join(v.split())
    return v

# ----------------------------------------------------------------------------
# Utilitaires d’affichage : surlignage + listes à puces
# ----------------------------------------------------------------------------

def highlight_html(text: str, pat: re.Pattern) -> str:
    if not isinstance(text, str) or not text:
        return text
    def _repl(m: re.Match) -> str:
        return m.group(0).replace(m.group(1), f'<span class="hit">{m.group(1)}</span>')
    return pat.sub(_repl, text)


def to_bullets_html(text: str) -> str:
    if not isinstance(text, str) or not text.strip():
        return ""
    s = text.replace("\n", "|").replace(";", "|")
    parts = [p.strip() for p in s.split("|") if p and p.strip()]
    if not parts:
        return ""
    lis = ''.join(f"<li>{p}</li>" for p in parts)
    return f"<ul class=\"list\">{lis}</ul>"

# ----------------------------------------------------------------------------
# Extraction / épuration
# ----------------------------------------------------------------------------

def extract_mentions_generic(text: str, pat: re.Pattern) -> str:
    if not isinstance(text, str) or not text.strip():
        return ""
    parts = re.split(r"[;,\n]", text)
    hits = [p.strip() for p in parts if pat.search(p)]
    return " | ".join(hits)


def extract_mentions_autres_sanctions(text: str, pat: re.Pattern) -> str:
    if not isinstance(text, str) or not text.strip():
        return ""
    candidates = []
    for seg in re.split(r"[;\n]", text):
        if pat.search(seg):
            candidates.append(seg.strip())
    return " | ".join(candidates)


def clean_filtered_df(df: pd.DataFrame, colmap: Dict[str, List[str]], pat: re.Pattern) -> pd.DataFrame:
    df = df.copy()
    all_present_cols: List[str] = []
    for canon in FILTER_CANONICAL:
        cols = [c for c in colmap.get(canon, []) if c in df.columns]
        all_present_cols.extend(cols)
        for col in cols:
            if canon == "autres_sanctions":
                df[col] = df[col].apply(lambda v: extract_mentions_autres_sanctions(_prep_text(v), pat))
            else:
                df[col] = df[col].apply(lambda v: extract_mentions_generic(_prep_text(v), pat))
    if all_present_cols:
        mask_any = False
        for c in all_present_cols:
            cur = df[c].astype(str).str.strip().ne("")
            mask_any = cur if mask_any is False else (mask_any | cur)
        df = df[mask_any]
    return df

# ----------------------------------------------------------------------------
# Export Excel
# ----------------------------------------------------------------------------

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

# ----------------------------------------------------------------------------
# Routes
# ----------------------------------------------------------------------------

@app.route("/", methods=["GET", "POST"])
def analyze():
    if request.method == "GET":
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                                      searched_article=None, message=None, message_ok=True)

    file = request.files.get("file")
    article = (request.form.get("article") or "").strip()

    if not file or not article:
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                                      searched_article=article, message="Erreur : fichier et article sont requis.",
                                      message_ok=False)

    # Validation extension
    fname = (file.filename or "").lower()
    if not (fname.endswith(".xlsx") or fname.endswith(".xlsm")):
        return render_template_string(
            HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None, searched_article=article,
            message=("Format non pris en charge : " + (file.filename or "").split(".")[-1] + ". "
                     "Veuillez fournir un classeur .xlsx ou .xlsm (pas .xls)."),
            message_ok=False
        )

    try:
        file_bytes = file.read()
        df, chosen_sheet = read_best_sheet(file_bytes)
        colmap = resolve_columns(df)
        pat = build_article_pattern(article)

        # Diagnostics : comptage des matches par colonne
        counts = []
        masks = []
        any_cols = False
        for canon in FILTER_CANONICAL:
            for col in [c for c in colmap.get(canon, []) if c in df.columns]:
                any_cols = True
                m = df[col].astype(str).apply(lambda v: bool(pat.search(_prep_text(v))))
                masks.append(m)
                counts.append((col, int(m.sum())))

        mapping_lines = [f"- {canon}: {colmap.get(canon)}" for canon in FILTER_CANONICAL]
        mapping_txt = "\n".join(mapping_lines)

        if not any_cols:
            return render_template_string(
                HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None, searched_article=article,
                message=("Erreur : aucune des colonnes attendues n’a été trouvée.\n"
                         f"Feuille choisie : {chosen_sheet}\n\n"
                         f"Colonnes détectées :\n{mapping_txt}\n\n"
                         f"Colonnes disponibles :\n{list(df.columns)}"),
                message_ok=False
            )

        mask_any = masks[0]
        for m in masks[1:]:
            mask_any = mask_any | m

        df_filtered = df[mask_any].copy()
        diag = "; ".join([f"{c}: {n}" for c, n in counts]) or "(aucune)"

        if df_filtered.empty:
            return render_template_string(
                HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None, searched_article=article,
                message=(f"Aucune ligne ne contient l’article « {article} ».\n"
                         f"Feuille : {chosen_sheet}\nColonnes :\n{mapping_txt}\nDétails : {diag}"),
                message_ok=True
            )

        df_clean = clean_filtered_df(df_filtered, colmap, pat)

        if df_clean.empty:
            return render_template_string(
                HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None, searched_article=article,
                message=("Des lignes correspondaient au motif, mais après épuration aucune mention nette n’a été conservée.\n"
                         f"Feuille : {chosen_sheet}\nColonnes :\n{mapping_txt}\nDétails : {diag}"),
                message_ok=True
            )

        download_url = to_excel_download(df_clean)

        # Prévisualisation : surlignage + listes à puces dans les 4 colonnes cibles
        preview = df_clean.head(200).copy()
        for canon in FILTER_CANONICAL:
            for col in [c for c in colmap.get(canon, []) if c in preview.columns]:
                preview[col] = preview[col].apply(lambda v: to_bullets_html(highlight_html(str(v), pat)))

        table_html = preview.to_html(index=False, escape=False)

        return render_template_string(
            HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=table_html, searched_article=article,
            download_url=download_url,
            message=(f"{len(df_clean)} ligne(s) après filtrage et épuration. (Aperçu limité à 200 lignes.)\n"
                     f"Feuille : {chosen_sheet}\nColonnes :\n{mapping_txt}\nDétails : {diag}"),
            message_ok=True
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


@app.route("/version")
def version():
    return jsonify({
        "version": APP_VERSION,
        "version_short": APP_VERSION_SHORT,
        "started_at": STARTED_AT,
        "time_utc": datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC"),
    })


@app.route("/health", methods=["GET", "HEAD"])
def health():
    return jsonify(status="ok", version=APP_VERSION_SHORT, started_at=STARTED_AT,
                   time_utc=datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC")), 200


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))












