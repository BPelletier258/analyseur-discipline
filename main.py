# === CANVAS META =============================================================
# Fichier : main.py — alias ordonnés + multi‑colonnes + diagnostics
# ============================================================================

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
  body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;margin:24px}
  h1{font-size:20px;margin-bottom:12px}
  form{display:grid;gap:12px;margin-bottom:16px}
  input[type=text]{padding:8px;font-size:14px}
  input[type=file]{font-size:14px}
  button{padding:8px 12px;font-size:14px;cursor:pointer}
  .note{background:#fff6e5;border:1px solid #ffd89b;padding:8px 10px;border-radius:6px;margin:10px 0 16px}
  .hint{font-size:12px;color:#666}
  table{border-collapse:collapse;width:100%;font-size:13px}
  th,td{border:1px solid #ddd;padding:6px 8px;vertical-align:top}
  th{background:#f3f4f6}
  .msg{margin-top:12px;white-space:pre-wrap;font-family:ui-monospace,SFMono-Regular,Menlo,monospace;font-size:12px}
  .ok{color:#065f46}.err{color:#7f1d1d}
  .kbd{font-family:ui-monospace,SFMono-Regular,Menlo,monospace;background:#f3f4f6;padding:2px 4px;border-radius:4px}
  footer{margin-top:18px;font-size:12px;color:#6b7280}
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
    <input type="text" name="article" value="{{ searched_article or '' }}" required />
    <label>Fichier Excel</label>
    <input type="file" name="file" accept=".xlsx,.xlsm" required />
    <button type="submit">Analyser</button>
    <div class="hint">Formats : .xlsx / .xlsm</div>
  </form>

  {% if table_html %}{{ table_html|safe }}{% endif %}

  {% if message %}
    <div class="msg {{ 'ok' if message_ok else 'err' }}">{{ message }}</div>
  {% endif %}

  <footer>
    Version: <strong>{{ app_version_short }}</strong> ({{ app_version }}) • Démarré: {{ started_at }} • <a href="/version">/version</a> • <a href="/health">/health</a>
  </footer>
</body>
</html>
"""

# ----------------------------------------------------------------------------
# Normalisation & alias d’en‑têtes (ORDONNÉS)
# ----------------------------------------------------------------------------

def _norm(s: str) -> str:
    if not isinstance(s, str):
        s = str(s) if s is not None else ""
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = s.replace("\u00A0", " ").replace("\u202F", " ")
    s = " ".join(s.strip().lower().split())
    return s

# IMPORTANT : listes (et non sets) pour respecter la PRIORITÉ
HEADER_ALIASES: Dict[str, List[str]] = {
    # 1) Nouvelles appellations d’abord
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
    """Retourne pour chaque canonique *toutes* les colonnes présentes (ordre prioritaire)."""
    norm_to_original = {_norm(c): c for c in df.columns}
    resolved: Dict[str, List[str]] = {}
    for canon, ordered_aliases in HEADER_ALIASES.items():
        hits: List[str] = []
        for alias in ordered_aliases:
            if alias in norm_to_original:
                hits.append(norm_to_original[alias])
        resolved[canon] = hits
    return resolved

# ----------------------------------------------------------------------------
# Lecture Excel (gestion « Article filtré : »)
# ----------------------------------------------------------------------------

def read_excel_respecting_header_rule(file_stream) -> pd.DataFrame:
    df_preview = pd.read_excel(file_stream, header=None, nrows=2, engine="openpyxl")
    file_stream.seek(0)
    first_cell = df_preview.iloc[0, 0] if not df_preview.empty else None
    banner = isinstance(first_cell, str) and _norm(first_cell).startswith(_norm("Article filtré :"))
    if banner:
        return pd.read_excel(file_stream, skiprows=1, header=0, engine="openpyxl")
    return pd.read_excel(file_stream, header=0, engine="openpyxl")

# ----------------------------------------------------------------------------
# Motif exact pour l’article (strict MAIS tolérant)
# ----------------------------------------------------------------------------

def build_article_pattern(user_input: str) -> re.Pattern:
    token = (user_input or "").strip()
    if not token:
        raise ValueError("Article vide.")
    token = token.replace("\u00A0", " ").replace("\u202F", " ")
    token = " ".join(token.split())

    esc = re.escape(token)
    ends_with_digit = token[-1].isdigit()
    right_tail = r"(?![\d.])" if ends_with_digit else r"\b"
    left_guard = r"(?<![\d.])"  # évite 2016→16
    return re.compile(rf"(?:\\bart(?:icle)?\\s*[: ]*)?{left_guard}({esc}){right_tail}", re.IGNORECASE)

# ----------------------------------------------------------------------------
# Pré‑traitement texte (puces, NBSP, CR/LF robuste)
# ----------------------------------------------------------------------------

def _prep_text(v: str) -> str:
    if not isinstance(v, str):
        v = "" if v is None else str(v)
    v = v.replace("•", " ").replace("·", " ").replace("◦", " ")
    v = v.replace("\u00A0", " ").replace("\u202F", " ")
    v = v.replace(chr(13)+chr(10), "\n").replace(chr(13), "\n")
    v = " ".join(v.split())
    return v

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

    fname = (file.filename or "").lower()
    if not (fname.endswith(".xlsx") or fname.endswith(".xlsm")):
        return render_template_string(
            HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None, searched_article=article,
            message=("Format non pris en charge : " + (file.filename or "").split(".")[-1] + ". "
                     "Veuillez fournir un classeur Excel .xlsx ou .xlsm. Les fichiers .xls ne sont pas supportés."),
            message_ok=False
        )

    try:
        df = read_excel_respecting_header_rule(file.stream)
        colmap = resolve_columns(df)
        pat = build_article_pattern(article)

        # Diagnostics : on SCANNE toutes les colonnes trouvées pour chaque canon
        counts = []
        masks = []
        any_cols = False
        for canon in FILTER_CANONICAL:
            for col in [c for c in colmap.get(canon, []) if c in df.columns]:
                any_cols = True
                m = df[col].astype(str).apply(lambda v: bool(pat.search(_prep_text(v))))
                masks.append(m)
                counts.append((col, int(m.sum())))

        if not any_cols:
            detail = "\n".join([f"  - {k}: {colmap.get(k)}" for k in FILTER_CANONICAL])
            return render_template_string(
                HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None, searched_article=article,
                message=("Erreur : aucune des colonnes attendues n’a été trouvée dans le fichier.\n"
                         "Vérifiez les en‑têtes.\n\n"
                         f"Colonnes résolues :\n{detail}\n\nColonnes disponibles :\n{list(df.columns)}"),
                message_ok=False
            )

        if not masks:
            return render_template_string(
                HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None, searched_article=article,
                message="Aucune colonne exploitable pour le filtrage.", message_ok=False
            )

        mask_any = masks[0]
        for m in masks[1:]:
            mask_any = mask_any | m

        df_filtered = df[mask_any].copy()
        diag = "; ".join([f"{c}: {n}" for c, n in counts]) or "(aucune)"

        if df_filtered.empty:
            return render_template_string(
                HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None, searched_article=article,
                message=(f"Aucune ligne ne contient l’article « {article} » dans les colonnes cibles.\n"
                         f"Détails (matches par colonne) : {diag}"),
                message_ok=True
            )

        df_clean = clean_filtered_df(df_filtered, colmap, pat)

        if df_clean.empty:
            return render_template_string(
                HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None, searched_article=article,
                message=("Des lignes correspondaient au motif, mais après épuration des cellules, "
                         "aucune mention nette de l’article n’a été conservée.\n"
                         f"Détails (matches par colonne) : {diag}"),
                message_ok=True
            )

        download_url = to_excel_download(df_clean)
        preview = df_clean.head(200)
        table_html = preview.to_html(index=False)

        return render_template_string(
            HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=table_html, searched_article=article,
            download_url=download_url,
            message=(f"{len(df_clean)} ligne(s) après filtrage et épuration. (Aperçu limité à 200 lignes.)\n"
                     f"Détails (matches par colonne) : {diag}"),
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
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)) )







