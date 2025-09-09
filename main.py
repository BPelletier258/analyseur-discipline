# -*- coding: utf-8 -*-
# Analyseur Discipline – Filtrage par article
# (rendu listes à puces + option "segments seuls" + titres centrés + NaN→—)

import os, re, time, unicodedata
from typing import Dict, Optional, List

import pandas as pd
from flask import Flask, request, render_template_string, send_file

app = Flask(__name__)

# ──────────────────────────────────────────────────────────────────────────────
# STYLE & HTML
# ──────────────────────────────────────────────────────────────────────────────

STYLE = """
<style>
  :root { --fg:#111827; --muted:#6b7280; --hit:#c1121f; }
  * { box-sizing: border-box; }
  body { font-family: system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif; color:var(--fg); margin:24px; }
  h1 { font-size:20px; margin:0 0 12px; }
  form { display:grid; gap:10px; margin:10px 0 16px; }
  label { font-size:14px; }
  input[type="text"] { padding:8px; font-size:14px; width:280px; }
  input[type="file"] { font-size:14px; }
  .row { display:flex; align-items:center; gap:16px; flex-wrap:wrap; }
  .hint { font-size:12px; color:var(--muted); }
  .note { background:#fff6e5; border:1px solid #ffd89b; padding:8px 10px; border-radius:6px; margin:10px 0 16px; }
  .download { margin:12px 0; }
  .kbd { font-family: ui-monospace, SFMono-Regular, Menlo, monospace; background:#f3f4f6; padding:2px 4px; border-radius:4px; }
  .msg { margin-top:12px; white-space:pre-wrap; font-family:ui-monospace, Menlo, monospace; font-size:12px; }
  .ok { color:#065f46; } .err { color:#7f1d1d; }

  /* Tableau */
  .viewport { height:60vh; overflow:auto; border:1px solid #e5e7eb; border-radius:6px; }
  table { border-collapse:collapse; width:100%; min-width:2000px; font-size:13px; }
  th, td { border:1px solid #e5e7eb; padding:6px 8px; vertical-align:top; }
  th { background:#f3f4f6; text-align:center; }
  td ul { margin:0; padding-left:20px; }
  td ul li { margin:0 0 2px; }
  .hit { color:var(--hit); font-weight:600; }
</style>
"""

HTML = """
<!doctype html>
<html>
<head><meta charset="utf-8"><title>Analyseur Discipline – Filtrage par article</title>{{ style|safe }}</head>
<body>
  <h1>Analyseur Discipline – Filtrage par article</h1>

  <div class="note">
    Règles : détection exacte de l’article; si la 1<sup>re</sup> cellule contient « <span class="kbd">Article filtré :</span> », elle est ignorée (entêtes sur la 2<sup>e</sup> ligne).
  </div>

  <form method="POST" enctype="multipart/form-data">
    <div class="row">
      <label>Article à rechercher</label>
      <input type="text" name="article" value="{{ article or '' }}" required placeholder="ex.: 29, 59(2)" />
      <label>Fichier Excel</label>
      <input type="file" name="file" accept=".xlsx,.xlsm" required />
    </div>
    <div class="row">
      <label><input type="checkbox" name="segments_only" value="1" {% if segments_only %}checked{% endif %} />
        Afficher uniquement les segments contenant l’article (colonnes d’intérêt + « Liste des chefs et articles en infraction »)</label>
    </div>
    <div class="hint">Formats acceptés : .xlsx / .xlsm</div>
    <div class="row"><button type="submit">Analyser</button></div>
  </form>

  {% if table_html %}
    <div class="download"><a href="{{ dl_url }}">Télécharger le résultat (Excel)</a></div>
    <div class="viewport">{{ table_html|safe }}</div>
  {% endif %}

  {% if message %}<div class="msg {{ 'ok' if ok else 'err' }}">{{ message }}</div>{% endif %}
</body>
</html>
"""

# ──────────────────────────────────────────────────────────────────────────────
# Normalisation & regex article
# ──────────────────────────────────────────────────────────────────────────────

def _norm(s: str) -> str:
    if not isinstance(s, str):
        s = "" if s is None else str(s)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = s.replace("\u00A0", " ")
    return " ".join(s.strip().lower().split())

def build_article_pattern(token: str) -> re.Pattern:
    token = (token or "").strip()
    if not token:
        raise ValueError("Article vide.")
    esc = re.escape(token)
    tail = r"(?![\d.])" if token[-1].isdigit() else r"\b"
    # capture du numéro pour le surlignage
    return re.compile(rf"(?:\b(?:art(?:icle)?\s*[: ]*)?)({esc}){tail}", re.IGNORECASE)

# ──────────────────────────────────────────────────────────────────────────────
# Lecture Excel avec bannière "Article filtré :"
# ──────────────────────────────────────────────────────────────────────────────

def read_excel(file_stream) -> pd.DataFrame:
    p = pd.read_excel(file_stream, header=None, nrows=2, engine="openpyxl")
    file_stream.seek(0)
    banner = isinstance(p.iloc[0,0] if not p.empty else None, str) and _norm(p.iloc[0,0]).startswith(_norm("Article filtré :"))
    if banner:
        return pd.read_excel(file_stream, skiprows=1, header=0, engine="openpyxl")
    return pd.read_excel(file_stream, header=0, engine="openpyxl")

# ──────────────────────────────────────────────────────────────────────────────
# Alias d’en-têtes
# ──────────────────────────────────────────────────────────────────────────────
# ➜ IMPORTANT : on sépare bien « Liste des chefs et articles en infraction »
# des 4 colonnes d’intérêt

ALIASES: Dict[str, List[str]] = {
    # 4 colonnes d’intérêt
    "nbr_chefs_par_articles": [
        "Nbr Chefs par articles", "Nombre de chefs par articles"
    ],
    "nbr_chefs_par_articles_par_periode": [
        "Nbr Chefs par articles par période de radiation",
        "Nbr Chefs par articles par periode de radiation"
    ],
    "nb_chefs_par_articles_total_amendes": [
        "Nombre de chefs par articles et total amendes"
    ],
    "nb_chefs_par_article_reprimande": [
        "Nombre de chefs par article ayant une réprimande",
        "Nombre de chefs par article ayant une reprimande"
    ],

    # Textes longs (à afficher en listes à puces)
    "liste_chefs_articles": [
        "Liste des chefs et articles en infraction"
    ],
    "resume_faits": ["Résumé des faits concis", "Resume des faits concis"],
    "liste_sanctions": ["Liste des sanctions imposées", "Liste des sanctions imposees"],
    "autres_mesures": ["Autres mesures ordonnées", "Autres mesures ordonnees"],
    "a_verifier": ["À vérifier", "A verifier", "A vérifier", "À verifier"],

    # Utiles pour info
    "total_chefs": ["Total chefs"],
}

def resolve(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    n2o = {_norm(c): c for c in df.columns}
    out: Dict[str, Optional[str]] = {}
    for canon, variants in ALIASES.items():
        found = None
        for v in variants:
            if _norm(v) in n2o:
                found = n2o[_norm(v)]
                break
        out[canon] = found
    return out

INTEREST_KEYS = [
    "nbr_chefs_par_articles",
    "nbr_chefs_par_articles_par_periode",
    "nb_chefs_par_articles_total_amendes",
    "nb_chefs_par_article_reprimande",
]

BULLET_KEYS = INTEREST_KEYS + [
    "liste_chefs_articles", "resume_faits", "liste_sanctions", "autres_mesures", "a_verifier"
]

# ──────────────────────────────────────────────────────────────────────────────
# Helpers rendu : nettoyage / split / puces / surlignage
# ──────────────────────────────────────────────────────────────────────────────

SEP_REGEX = re.compile(r"(?:\n|•|\u2022|\u2027|\u25E6|\s\|\s)")

def _prep(v) -> str:
    if not isinstance(v, str):
        v = "" if v is None else str(v)
    # normalise retours / NBSP / puces pour la recherche
    v = v.replace("\r", "\n").replace(" ", " ").replace(" ", " ")
    return v

def split_segments(v: str) -> List[str]:
    v = _prep(v).strip()
    if not v or v.lower() in ("nan", "none"):
        return []
    # On convertit " | " en retour, puis split sur plusieurs séparateurs courants
    v = v.replace(" | ", "\n")
    parts = [p.strip(" •\t") for p in SEP_REGEX.split(v)]
    return [p for p in parts if p]

def highlight_html(s: str, pat: re.Pattern) -> str:
    return pat.sub(r'<span class="hit">\1</span>', s)

def to_html_bullets(v: str, pat: Optional[re.Pattern]=None, segments_only=False) -> str:
    segs = split_segments(v)
    if segments_only and pat is not None:
        segs = [x for x in segs if pat.search(_prep(x))]
    if not segs:
        return "—"
    if pat is not None:
        segs = [highlight_html(x, pat) for x in segs]
    return "<ul>" + "".join(f"<li>{x}</li>" for x in segs) + "</ul>"

def to_text_bullets(v: str, pat: Optional[re.Pattern]=None, segments_only=False) -> str:
    segs = split_segments(v)
    if segments_only and pat is not None:
        segs = [x for x in segs if pat.search(_prep(x))]
    if not segs:
        return "—"
    # Pas de surlignage dans l’export Excel
    return " • ".join(segs)

# ──────────────────────────────────────────────────────────────────────────────
# Filtrage des lignes (conserve toutes les colonnes, mais retient les lignes
# où l’article est présent dans AU MOINS une des 4 colonnes d’intérêt)
# ──────────────────────────────────────────────────────────────────────────────

def row_filter_mask(df: pd.DataFrame, colmap: Dict[str,str], pat: re.Pattern) -> pd.Series:
    masks = []
    has_any = False
    for key in INTEREST_KEYS:
        col = colmap.get(key)
        if col and col in df.columns:
            has_any = True
            masks.append(df[col].astype(str).apply(lambda x: bool(pat.search(_prep(x)))))
    if not has_any:
        # si le classeur ne contient pas les colonnes d’intérêt -> toutes False
        return pd.Series([False]*len(df), index=df.index)
    m = masks[0]
    for k in masks[1:]:
        m = m | k
    return m

# ──────────────────────────────────────────────────────────────────────────────
# TABLE → HTML + EXCEL
# ──────────────────────────────────────────────────────────────────────────────

def build_display_and_export(df: pd.DataFrame, colmap: Dict[str,str], pat: re.Pattern, segments_only: bool):
    # 1) Vue HTML (copie)
    view = df.copy()

    # NaN -> "—" partout par défaut
    view = view.fillna("—")

    # 2) Colonnes à puces (HTML)
    for key in BULLET_KEYS:
        col = colmap.get(key)
        if col and col in view.columns:
            # surlignage demandé aussi dans « Liste des chefs et articles en infraction »
            apply_pat = pat if (key in INTEREST_KEYS or key == "liste_chefs_articles") else None
            view[col] = view[col].apply(lambda x: to_html_bullets(x, apply_pat, segments_only))

    # 3) Construction HTML avec titres centrés (to_html + escape=False)
    table_html = view.to_html(index=False, escape=False)

    # 4) DataFrame export (texte simple)
    export = df.copy()
    export = export.fillna("—")
    for key in BULLET_KEYS:
        col = colmap.get(key)
        if col and col in export.columns:
            export[col] = export[col].apply(lambda x: to_text_bullets(x, None, segments_only))
    return table_html, export

def to_excel_download(df: pd.DataFrame) -> str:
    ts = int(time.time())
    out = f"/tmp/filtrage_{ts}.xlsx"
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Filtre")
        ws = w.book.active
        # auto-largeurs
        for i, col in enumerate(df.columns, 1):
            vals = [str(col)] + df[col].astype(str).tolist()
            width = min(60, max(12, max(len(v) for v in vals) + 2))
            ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = width
    return f"/download?path={out}"

# ──────────────────────────────────────────────────────────────────────────────
# ROUTES
# ──────────────────────────────────────────────────────────────────────────────

@app.route("/", methods=["GET","POST"])
def index():
    if request.method == "GET":
        return render_template_string(HTML, style=STYLE, table_html=None, article="", segments_only=False,
                                      message=None, ok=True)

    file = request.files.get("file")
    article = (request.form.get("article") or "").strip()
    segments_only = request.form.get("segments_only") == "1"

    if not file or not article:
        return render_template_string(HTML, style=STYLE, table_html=None, article=article, segments_only=segments_only,
                                      message="Erreur : fichier et article requis.", ok=False)

    name = (file.filename or "").lower()
    if not (name.endswith(".xlsx") or name.endswith(".xlsm")):
        return render_template_string(HTML, style=STYLE, table_html=None, article=article, segments_only=segments_only,
                                      message="Format non pris en charge. Utilisez .xlsx ou .xlsm.", ok=False)

    try:
        df = read_excel(file.stream)
        colmap = resolve(df)

        pat = build_article_pattern(article)
        mask = row_filter_mask(df, colmap, pat)
        df_filtered = df[mask].copy()

        if df_filtered.empty:
            return render_template_string(HTML, style=STYLE, table_html=None, article=article, segments_only=segments_only,
                                          message=f"Aucune ligne ne contient l’article « {article} » dans les colonnes d’intérêt.",
                                          ok=True)

        table_html, export_df = build_display_and_export(df_filtered, colmap, pat, segments_only)
        dl_url = to_excel_download(export_df)

        msg = f"{len(df_filtered)} ligne(s) retenue(s)."
        return render_template_string(HTML, style=STYLE, table_html=table_html, article=article, segments_only=segments_only,
                                      dl_url=dl_url, message=msg, ok=True)

    except Exception as e:
        return render_template_string(HTML, style=STYLE, table_html=None, article=article, segments_only=segments_only,
                                      message=f"Erreur inattendue : {repr(e)}", ok=False)

@app.route("/download")
def download():
    path = request.args.get("path")
    if not path or not os.path.exists(path):
        return "Fichier introuvable ou expiré.", 404
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
