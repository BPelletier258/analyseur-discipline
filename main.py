# === CANVAS META =============================================================
# Fichier : main.py — version + listes + largeur colonnes + anti-NaN (09-08b)
# ============================================================================
import io
import os
import re
import time
import unicodedata
from datetime import datetime
from typing import Dict, Optional, Set, List

import pandas as pd
from flask import Flask, request, render_template_string, send_file

app = Flask(__name__)

STYLE_BLOCK = """
<style>
  body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;margin:24px}
  h1{font-size:20px;margin-bottom:12px}
  form{display:grid;gap:12px;margin-bottom:16px}
  input[type="text"]{padding:8px;font-size:14px}
  input[type="file"]{font-size:14px}
  button{padding:8px 12px;font-size:14px;cursor:pointer}
  .hint{font-size:12px;color:#666}
  .note{background:#fff6e5;border:1px solid #ffd89b;padding:8px 10px;border-radius:6px;margin:10px 0 16px}
  table{border-collapse:collapse;width:100%;font-size:13px;table-layout:auto}
  th,td{border:1px solid #ddd;padding:6px 8px;vertical-align:top}
  th{background:#f3f4f6;text-align:center}
  td{text-align:left}
  .msg{margin-top:12px;white-space:pre-wrap;font-family:ui-monospace,SFMono-Regular,Menlo,monospace;font-size:12px}
  .ok{color:#065f46}.err{color:#7f1d1d}
  .download{margin:12px 0}.kbd{font-family:ui-monospace,SFMono-Regular,Menlo,monospace;background:#f3f4f6;padding:2px 4px;border-radius:4px}
  /* viewport + barre horizontale */
  .table-viewport{height:60vh;overflow:auto;border:1px solid #ddd}
  .table-wide{min-width:1600px}
  .table-viewport table{width:100%}
  /* listes à puces */
  ul.bul{margin:0;padding-left:1.25em}
  ul.bul li{margin:0.20em 0}
  /* surbrillance de l’article dans les 4 colonnes d’intérêt */
  .hit{color:#b91c1c;font-weight:600}
  /* largeur confort pour les 4 colonnes d’intérêt */
  .col-wide{min-width:420px;width:420px}
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
# Normalisation de texte / en-têtes
# ──────────────────────────────────────────────────────────────────────────────

def _norm(s: str) -> str:
    if not isinstance(s, str):
        s = "" if s is None else str(s)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = s.replace("\u00A0", " ")
    s = " ".join(s.strip().lower().split())
    return s

def _prep_text(v: str) -> str:
    """Nettoyage : puces unicode, NBSP, retours chariot, séquence '\\n' littérale → LF."""
    if not isinstance(v, str):
        v = "" if v is None else str(v)
    v = v.replace("•", " ").replace("·", " ").replace("◦", " ")
    v = v.replace(" ", " ").replace(" ", " ")
    # séquences littérales "\n" → LF
    v = v.replace("\\n", "\n")
    # CRLF / CR → LF
    v = v.replace("\r\n", "\n").replace("\r", "\n")
    # espaces superflus
    v = "\n".join(" ".join(line.split()) for line in v.split("\n"))
    return v.strip()

# ──────────────────────────────────────────────────────────────────────────────
# Aliases d’en-têtes (inclut nouveaux libellés)
# ──────────────────────────────────────────────────────────────────────────────

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
        _norm("Autres mesures ordonnees"),
    },
    # colonnes à puces (affichage)
    "resume_faits": {
        _norm("Résumé des faits concis"),
        _norm("Resume des faits concis"),
        _norm("Resume des faits"),
    },
    "liste_sanctions_imposees": {
        _norm("Liste des sanctions imposées"),
        _norm("Liste des sanctions imposees"),
    },
    "autres_mesures_ordonnees": {
        _norm("Autres mesures ordonnées"),
        _norm("Autres mesures ordonnees"),
    },
    "a_verifier": {
        _norm("À vérifier"),
        _norm("A verifier"),
    },
}

FILTER_CANONICAL = [
    "articles_enfreints",
    "duree_totale_radiation",
    "article_amende_chef",
    "autres_sanctions",
]

BULLETIZE_EXTRA = [  # colonnes à puces côté affichage
    "resume_faits",
    "liste_sanctions_imposees",
    "autres_mesures_ordonnees",
    "a_verifier",
]

def resolve_columns(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    norm_to_original = {_norm(c): c for c in df.columns}
    out: Dict[str, Optional[str]] = {}
    for canon, variants in HEADER_ALIASES.items():
        hit = None
        for v in variants:
            if v in norm_to_original:
                hit = norm_to_original[v]
                break
        out[canon] = hit
    return out

# ──────────────────────────────────────────────────────────────────────────────
# Lecture Excel (règle « Article filtré : »)
# ──────────────────────────────────────────────────────────────────────────────

def read_excel_respecting_header_rule(file_stream) -> pd.DataFrame:
    df_preview = pd.read_excel(file_stream, header=None, nrows=2, engine="openpyxl")
    file_stream.seek(0)

    first_cell = df_preview.iloc[0, 0] if not df_preview.empty else None
    is_banner = isinstance(first_cell, str) and _norm(first_cell).startswith(_norm("Article filtré :"))

    if is_banner:
        return pd.read_excel(file_stream, skiprows=1, header=0, engine="openpyxl")
    return pd.read_excel(file_stream, header=0, engine="openpyxl")

# ──────────────────────────────────────────────────────────────────────────────
# Motif pour l’article + helpers d’extraction / puces
# ──────────────────────────────────────────────────────────────────────────────

def build_article_pattern(user_input: str) -> re.Pattern:
    token = (user_input or "").strip()
    if not token:
        raise ValueError("Article vide.")
    esc = re.escape(token)
    ends_with_digit = token[-1].isdigit()
    tail = r"(?![\d.])" if ends_with_digit else r"\b"
    return re.compile(rf"(?:\b(?:art(?:icle)?\s*[: ]*)?)({esc}){tail}", re.IGNORECASE)

def _split_items(text: str) -> List[str]:
    """Découpe un bloc en éléments (séparateurs ; , LF)."""
    text = _prep_text(text)
    if not text:
        return []
    parts = re.split(r"[;,\n]+", text)
    return [p.strip() for p in parts if p.strip()]

def _bul_html(items: List[str]) -> str:
    if not items:
        return ""
    return "<ul class='bul'>" + "".join(f"<li>{it}</li>" for it in items) + "</ul>"

def _colorize_items(items: List[str], pat: re.Pattern) -> List[str]:
    return [pat.sub(r"<span class='hit'>\\1</span>", it) for it in items]

# ──────────────────────────────────────────────────────────────────────────────
# Nettoyage + extraction + mise en forme (HTML)
# ──────────────────────────────────────────────────────────────────────────────

def clean_and_format(df: pd.DataFrame, colmap: Dict[str, Optional[str]], pat: re.Pattern) -> pd.DataFrame:
    df = df.copy()

    # 1) Filtrer lignes où l’article apparaît dans au moins une des 4 colonnes d’intérêt
    masks = []
    for canon in FILTER_CANONICAL:
        col = colmap.get(canon)
        if col and col in df.columns:
            masks.append(df[col].astype(str).apply(lambda v: bool(pat.search(_prep_text(v)))))
    if not masks:
        return df.iloc[0:0]  # vide
    mask_any = masks[0]
    for m in masks[1:]:
        mask_any = mask_any | m
    df = df[mask_any].copy()

    # 2) Mise en puces + surbrillance pour les 4 colonnes d’intérêt
    for canon in FILTER_CANONICAL:
        col = colmap.get(canon)
        if not col or col not in df.columns:
            continue
        df[col] = df[col].apply(lambda v: _bul_html(_colorize_items(_split_items(v), pat)))

    # 3) Colonnes supplémentaires à afficher en puces (sans coloration)
    for canon in BULLETIZE_EXTRA:
        col = colmap.get(canon)
        if col and col in df.columns:
            df[col] = df[col].apply(lambda v: _bul_html(_split_items(v)))

    # 4) Anti-NaN
    df = df.fillna("")

    return df

# ──────────────────────────────────────────────────────────────────────────────
# Export Excel
# ──────────────────────────────────────────────────────────────────────────────

def to_excel_download(df: pd.DataFrame) -> str:
    ts = int(time.time())
    out_path = f"/tmp/filtrage_{ts}.xlsx"
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Filtre")
        ws = writer.book.active
        for cidx, col in enumerate(df.columns, start=1):
            max_len = max((len(str(x)) for x in [col] + df[col].astype(str).tolist()), default=10)
            ws.column_dimensions[ws.cell(row=1, column=cidx).column_letter].width = min(60, max(12, max_len + 2))
    return f"/download?path={out_path}"

# ──────────────────────────────────────────────────────────────────────────────
# Route principale
# ──────────────────────────────────────────────────────────────────────────────

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

    # .xlsx / .xlsm seulement
    fname = (file.filename or "").lower()
    if not (fname.endswith(".xlsx") or fname.endswith(".xlsm")):
        ext = (file.filename or "").split(".")[-1]
        return render_template_string(
            HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None, searched_article=article,
            message=f"Format non pris en charge : .{ext}. Fournissez un classeur .xlsx ou .xlsm.",
            message_ok=False
        )

    try:
        df = read_excel_respecting_header_rule(file.stream)
        colmap = resolve_columns(df)
        pat = build_article_pattern(article)

        df_clean = clean_and_format(df, colmap, pat)
        if df_clean.empty:
            # diagnostic utile
            col_diag = "\n".join([f" - {k}: {colmap.get(k)}" for k in FILTER_CANONICAL + BULLETIZE_EXTRA])
            return render_template_string(
                HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None, searched_article=article,
                message=(f"Aucune ligne ne contient l’article « {article} » dans les colonnes cibles.\n\n"
                         f"Colonnes détectées:\n{col_diag}\n\nColonnes du fichier:\n{list(df.columns)}"),
                message_ok=True
            )

        # Excel téléchargeable
        download_url = to_excel_download(df_clean)

        # Mise en forme HTML avec largeurs spécifiques pour les 4 colonnes d’intérêt
        present_interest = [colmap[k] for k in FILTER_CANONICAL if colmap.get(k) in df_clean.columns]
        sty = (df_clean.style
               .hide(axis="index")
               .set_properties(subset=present_interest, **{"min-width": "420px", "width": "420px"})
               )
        table_html = sty.to_html(na_rep="", escape=False)

        return render_template_string(
            HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=table_html, download_url=download_url,
            searched_article=article,
            message=f"{len(df_clean)} ligne(s) après filtrage et mise en forme. (Aperçu limité à 200 lignes.)",
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

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
