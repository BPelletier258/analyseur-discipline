# === CANVAS META =============================================================
# Fichier : main.py — filtre par article + surlignage + puces + option "mode concentré"
# Stamp    : 2025-09-09T00:00Z
# ============================================================================
import io
import os
import re
import time
import unicodedata
from datetime import datetime
from typing import Dict, Optional, Set
from html import escape

import pandas as pd
from flask import Flask, request, render_template_string, send_file

app = Flask(__name__)

# ──────────────────────────────────────────────────────────────────────────────
# Styles & gabarit HTML
# ──────────────────────────────────────────────────────────────────────────────

STYLE_BLOCK = """
<style>
  :root{
    --hit:#c1121f;
    --th:#f3f4f6;
    --bd:#e5e7eb;
  }
  body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;margin:24px}
  h1{font-size:20px;margin:0 0 10px}
  form{display:grid;grid-template-columns:1fr;gap:10px;margin:12px 0 16px}
  input[type="text"]{padding:8px;font-size:14px}
  input[type="file"]{font-size:14px}
  label{font-weight:600}
  .hint{font-size:12px;color:#6b7280}
  .note{background:#fff7ed;border:1px solid #fed7aa;padding:10px;border-radius:8px;margin:10px 0 14px}
  .kbd{font-family:ui-monospace,SFMono-Regular,Menlo,monospace;background:#f3f4f6;padding:2px 4px;border-radius:4px}
  .options{display:flex;gap:14px;align-items:center}
  .msg{margin-top:12px;white-space:pre-wrap;font-family:ui-monospace,SFMono-Regular,Menlo,monospace;font-size:12px}
  .ok{color:#065f46}.err{color:#7f1d1d}

  /* Tableau scrollable horizontalement (barre visible) */
  .table-viewport{height:60vh;overflow:auto;border:1px solid var(--bd)}
  .table-wide{min-width:200vw} /* ≈ deux écrans ; ajuste si besoin */
  table{border-collapse:collapse;width:100%;font-size:13px}
  th,td{border:1px solid var(--bd);padding:6px 8px;vertical-align:top}
  th{background:var(--th);text-align:center}
  ul.bul{margin:0;padding-left:1.1em}
  ul.bul>li{margin:0.15em 0}
  .hit{color:var(--hit);font-weight:700}
</style>
"""

HTML_TEMPLATE = """
<!doctype html>
<html>
<head>
  <meta charset="utf-8"/>
  <title>Analyseur Discipline – Filtrage par article</title>
  {{ style_block|safe }}
</head>
<body>
  <h1>Analyseur Discipline – Filtrage par article</h1>

  <div class="note">
    Règles : détection exacte de l’article; si la 1<sup>re</sup> cellule contient
    « <span class="kbd">Article filtré :</span> », elle est ignorée (en-têtes sur la 2<sup>e</sup> ligne).
  </div>

  <form method="POST" enctype="multipart/form-data">
    <div>
      <label>Article à rechercher (ex. <span class="kbd">29</span>, <span class="kbd">59(2)</span>)</label>
      <input type="text" name="article" value="{{ searched_article or '' }}" required placeholder="ex.: 29 ou 59(2)"/>
    </div>

    <div>
      <label>Fichier Excel</label>
      <input type="file" name="file" accept=".xlsx,.xlsm" required/>
      <div class="hint">Formats : .xlsx / .xlsm</div>
    </div>

    <div class="options">
      <label>
        <input type="checkbox" name="isolate" {% if isolate %}checked{% endif %}/>
        Mode concentré : ne montrer que les segments contenant l’article <span class="hint">(dans les 4 colonnes d’intérêt)</span>
      </label>
    </div>

    <button type="submit">Analyser</button>
  </form>

  {% if table_html %}
    <p><a href="{{ download_url }}">Télécharger le résultat (Excel)</a></p>
    <div class="table-viewport"><div class="table-wide">{{ table_html|safe }}</div></div>
  {% endif %}

  {% if message %}
    <div class="msg {{ 'ok' if message_ok else 'err' }}">{{ message }}</div>
  {% endif %}
</body>
</html>
"""

# ──────────────────────────────────────────────────────────────────────────────
# Normalisation & aliases d’en-têtes
# ──────────────────────────────────────────────────────────────────────────────

def _norm(s: str) -> str:
    """Normalise : accents→ASCII, minuscule, espaces compressés."""
    if not isinstance(s, str):
        s = str(s) if s is not None else ""
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = s.replace("\u00A0", " ")
    return " ".join(s.strip().lower().split())

# Colonnes d’intérêt (canoniques) & alias
HEADER_ALIASES: Dict[str, Set[str]] = {
    # 1) Nbr Chefs par articles
    "articles_enfreints": {
        _norm("Nbr Chefs par articles"),
        _norm("Articles enfreints"),
        _norm("Articles en infraction"),
        _norm("Liste des chefs et articles en infraction"),
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
    },
    # Colonne “Liste des chefs et articles en infraction” (pour surlignage/puces)
    "liste_chefs_articles": {
        _norm("Liste des chefs et articles en infraction"),
    },
}

INTEREST_CANONICAL = ["articles_enfreints", "duree_totale_radiation", "article_amende_chef", "autres_sanctions"]
HIGHLIGHT_CANONICAL = INTEREST_CANONICAL + ["liste_chefs_articles"]

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
# Lecture Excel (gestion « Article filtré : » en 1re ligne)
# ──────────────────────────────────────────────────────────────────────────────

def read_excel_respecting_header_rule(file_stream) -> pd.DataFrame:
    df_prev = pd.read_excel(file_stream, header=None, nrows=2, engine="openpyxl")
    file_stream.seek(0)
    banner = isinstance(df_prev.iloc[0, 0] if not df_prev.empty else None, str) and \
             _norm(df_prev.iloc[0, 0]).startswith(_norm("Article filtré :"))
    if banner:
        return pd.read_excel(file_stream, skiprows=1, header=0, engine="openpyxl")
    return pd.read_excel(file_stream, header=0, engine="openpyxl")

# ──────────────────────────────────────────────────────────────────────────────
# Motif de recherche & prétraitement
# ──────────────────────────────────────────────────────────────────────────────

def build_article_pattern(user_input: str) -> re.Pattern:
    token = (user_input or "").strip()
    if not token:
        raise ValueError("Article vide.")
    esc = re.escape(token)
    tail = r"(?![\d.])" if token[-1].isdigit() else r"\b"
    # capture = groupe 1 (pour surlignage)
    return re.compile(rf"(?:\b(?:art(?:icle)?\s*[: ]*)?)({esc}){tail}", re.IGNORECASE)

def _prep_text(v: str) -> str:
    if not isinstance(v, str):
        v = "" if v is None else str(v)
    # normalise les séparateurs ; supprime NBSP/CR
    v = v.replace("•", "\n").replace("·", "\n").replace("◦", "\n")
    v = v.replace(" ", " ").replace(" ", " ")
    v = v.replace("\r\n", "\n").replace("\r", "\n")
    # compresse les espaces multiples
    v = re.sub(r"[ \t]+", " ", v)
    return v.strip()

# ──────────────────────────────────────────────────────────────────────────────
# Helpers : extraction (mode concentré), surlignage & listes à puces
# ──────────────────────────────────────────────────────────────────────────────

def _split_segments(text: str) -> list[str]:
    """Découpe un contenu multi-énoncés en segments lisibles."""
    text = _prep_text(text)
    if not text:
        return []
    # split sur sauts de ligne / point-virgule / puces converties
    parts = re.split(r"\n|;", text)
    return [p.strip(" •·◦-—\u2022").strip() for p in parts if p.strip(" •·◦-—\u2022").strip()]

def extract_only_matches(text: str, pat: re.Pattern) -> str:
    """Mode concentré : ne conserver que les segments contenant l’article."""
    segs = _split_segments(text)
    hits = [s for s in segs if pat.search(s)]
    return " | ".join(hits)

def to_bullets_with_highlight(text: str, pat: re.Pattern, highlight: bool) -> str:
    segs = _split_segments(text)
    if not segs:
        return ""
    items = []
    for s in segs:
        s_html = escape(s)
        if highlight:
            s_html = pat.sub(r'<span class="hit">\1</span>', s_html)
        items.append(f"<li>{s_html}</li>")
    return f'<ul class="bul">{"".join(items)}</ul>'

# ──────────────────────────────────────────────────────────────────────────────
# Export Excel
# ──────────────────────────────────────────────────────────────────────────────

def to_excel_download(df: pd.DataFrame) -> str:
    ts = int(time.time())
    path = f"/tmp/filtrage_{ts}.xlsx"
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Filtre")
        ws = writer.book.active
        # auto-largeur simple
        for i, col in enumerate(df.columns, start=1):
            max_len = max((len(str(x)) for x in [col] + df[col].astype(str).tolist()), default=10)
            ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = min(60, max(12, max_len + 2))
    return f"/download?path={path}"

# ──────────────────────────────────────────────────────────────────────────────
# Routes
# ──────────────────────────────────────────────────────────────────────────────

@app.route("/", methods=["GET", "POST"])
def analyze():
    if request.method == "GET":
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK,
                                      table_html=None, searched_article=None,
                                      isolate=False, message=None, message_ok=True)

    file = request.files.get("file")
    article = (request.form.get("article") or "").strip()
    isolate = bool(request.form.get("isolate"))

    if not file or not article:
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                                      searched_article=article, isolate=isolate,
                                      message="Erreur : fichier et article sont requis.", message_ok=False)

    # Vérifie extension (openpyxl)
    fname = (file.filename or "").lower()
    if not (fname.endswith(".xlsx") or fname.endswith(".xlsm")):
        ext = (file.filename or "").split(".")[-1]
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                                      searched_article=article, isolate=isolate,
                                      message=f"Format non pris en charge : .{ext}. Utilisez .xlsx ou .xlsm.",
                                      message_ok=False)

    try:
        df = read_excel_respecting_header_rule(file.stream)
        colmap = resolve_columns(df)
        pat = build_article_pattern(article)

        # 1) FILTRAGE DES LIGNES (inchangé) : article présent dans AU MOINS UNE colonne d’intérêt
        masks = []
        any_cols = False
        for canon in INTEREST_CANONICAL:
            col = colmap.get(canon)
            if col and col in df.columns:
                any_cols = True
                masks.append(df[col].astype(str).apply(lambda v: bool(pat.search(_prep_text(v)))))
        if not any_cols:
            detail = "\n".join([f"  - {k}: {colmap.get(k)}" for k in INTEREST_CANONICAL])
            return render_template_string(
                HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                searched_article=article, isolate=isolate, message_ok=False,
                message=("Aucune des colonnes d’intérêt n’a été trouvée.\n"
                         "Vérifiez les en-têtes ou les alias.\n\nColonnes résolues:\n" + detail)
            )

        mask_any = masks[0]
        for m in masks[1:]:
            mask_any = mask_any | m
        df_filtered = df[mask_any].copy()

        if df_filtered.empty:
            return render_template_string(
                HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                searched_article=article, isolate=isolate, message_ok=True,
                message=f"Aucune ligne ne contient l’article « {article} » dans les colonnes cibles."
            )

        # 2) MODE D’AFFICHAGE
        # - mode normal : on garde les cellules intactes ; on met en puces + surlignage
        # - mode concentré : dans les 4 colonnes d’intérêt, on ne garde que les segments contenant l’article
        df_display = df_filtered.copy()

        if isolate:
            for canon in INTEREST_CANONICAL:
                col = colmap.get(canon)
                if col and col in df_display.columns:
                    df_display[col] = df_display[col].apply(lambda v: extract_only_matches(v, pat))

        # 3) Formattage : listes à puces + surlignage dans les colonnes à mettre en valeur
        for canon in HIGHLIGHT_CANONICAL:
            col = colmap.get(canon)
            if col and col in df_display.columns:
                df_display[col] = df_display[col].apply(lambda v: to_bullets_with_highlight(v, pat, highlight=True))

        # (les autres colonnes restent brutes)

        # 4) EXPORT Excel
        download_url = to_excel_download(df_display if isolate else df_filtered)

        # 5) Aperçu HTML
        preview = df_display.head(200)  # limite d’affichage
        html = preview.to_html(index=False, escape=False)

        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK,
                                      table_html=html, searched_article=article,
                                      isolate=isolate, download_url=download_url,
                                      message=f"{len(df_filtered)} ligne(s) conservée(s). "
                                              f"{'Mode concentré activé.' if isolate else 'Affichage intégral.'}",
                                      message_ok=True)

    except Exception as e:
        return render_template_string(
            HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
            searched_article=article, isolate=isolate,
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
