# main.py — affichage en listes + surbrillance + option "Segments seulement"
# (09-08)

import io, os, re, time, html, unicodedata
from datetime import datetime
from typing import Dict, Optional, Set, List

import pandas as pd
from flask import Flask, request, render_template_string, send_file

app = Flask(__name__)

# ---------- UI / styles ----------
STYLE_BLOCK = """
<style>
  :root { --hit:#c1121f; --line:#e5e7eb; --th:#f3f4f6; --note:#fff6e5; }
  body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;margin:24px}
  h1{font-size:20px;margin:0 0 10px}
  form{display:grid;gap:10px;margin:14px 0 16px}
  input[type="text"]{padding:8px;font-size:14px}
  input[type="file"]{font-size:14px}
  button{padding:8px 12px;font-size:14px;cursor:pointer}
  .hint{font-size:12px;color:#6b7280}
  .note{background:var(--note);border:1px solid #ffd89b;padding:8px 10px;border-radius:6px;margin:8px 0 14px}
  .kbd{font-family:ui-monospace,SFMono-Regular,Menlo,monospace;background:#f3f4f6;padding:2px 4px;border-radius:4px}
  .download{margin:12px 0}
  .msg{margin-top:12px;white-space:pre-wrap;font-family:ui-monospace,SFMono-Regular,Menlo,monospace;font-size:12px}
  .ok{color:#065f46}.err{color:#7f1d1d}

  .table-viewport{height:60vh;overflow:auto;border:1px solid var(--line)}
  .table-viewport table{border-collapse:collapse;width:100%;font-size:13px}
  th,td{border:1px solid var(--line);padding:6px 8px;vertical-align:top}
  th{background:var(--th)}
  td ul{margin:0;padding-left:18px}
  td li{margin:2px 0}
  .hit{color:var(--hit);font-weight:700}
</style>
"""

HTML_TEMPLATE = """
<!doctype html>
<html><head><meta charset="utf-8" />
<title>Analyseur Discipline – Filtrage par article</title>
{{ style_block|safe }}</head><body>
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
  <label>
    <input type="checkbox" name="segments_only" {% if segments_only %}checked{% endif %} />
    Isoler <b>uniquement les segments</b> contenant l’article (dans les 4 colonnes d’intérêt)
  </label>
  <button type="submit">Analyser</button>
  <div class="hint">Formats : .xlsx / .xlsm</div>
</form>

{% if table_html %}
  <div class="download"><a href="{{ download_url }}">Télécharger le résultat (Excel)</a></div>
  <div class="table-viewport">{{ table_html|safe }}</div>
{% endif %}

{% if message %}<div class="msg {{ 'ok' if message_ok else 'err' }}">{{ message }}</div>{% endif %}
</body></html>
"""

# ---------- utilitaires d'en-têtes ----------
def _norm(s: str) -> str:
    if not isinstance(s, str):
        s = "" if s is None else str(s)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = s.replace("\u00A0", " ")
    return " ".join(s.strip().lower().split())

# Alias pour colonnes
HEADER_ALIASES: Dict[str, Set[str]] = {
    # 4 colonnes d'intérêt
    "articles_enfreints": {
        _norm("Nbr Chefs par articles"),
        _norm("Articles enfreints"), _norm("Articles en infraction"),
        _norm("Nbr   Chefs   par    articles")
    },
    "duree_totale_radiation": {
        _norm("Nbr Chefs par articles par période de radiation"),
        _norm("Nbr Chefs par articles par periode de radiation"),
        _norm("Durée totale effective radiation"),
        _norm("Duree totale effective radiation"),
    },
    "article_amende_chef": {
        _norm("Nombre de chefs par articles et total amendes"),
        _norm("Article amende/chef"), _norm("Articles amende / chef"),
        _norm("Amendes (article/chef)"),
    },
    "autres_sanctions": {
        _norm("Nombre de chefs par article ayant une réprimande"),
        _norm("Nombre de chefs par article ayant une reprimande"),
        _norm("Autres sanctions"),
    },
    # colonnes texte à bulletiser (en plus)
    "resume_faits": {_norm("Résumé des faits concis"), _norm("Resume des faits concis")},
    "liste_chefs_articles": {_norm("Liste des chefs et articles en infraction")},
    "liste_sanctions": {_norm("Liste des sanctions imposées"), _norm("Liste des sanctions imposees")},
    "autres_mesures": {_norm("Autres mesures ordonnées"), _norm("Autres mesures ordonnees")},
    "a_verifier": {_norm("À vérifier"), _norm("A verifier")},
}

INTEREST_KEYS = ["articles_enfreints", "duree_totale_radiation", "article_amende_chef", "autres_sanctions"]
ALWAYS_BULLET_KEYS = ["resume_faits", "liste_chefs_articles", "liste_sanctions", "autres_mesures", "a_verifier"]

def resolve_columns(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    mapping = {_norm(c): c for c in df.columns}
    out: Dict[str, Optional[str]] = {}
    for canon, variants in HEADER_ALIASES.items():
        original = None
        for v in variants:
            if v in mapping:
                original = mapping[v]; break
        out[canon] = original
    return out

# ---------- lecture Excel (règle "Article filtré :") ----------
def read_excel_respecting_header_rule(file_stream) -> pd.DataFrame:
    peek = pd.read_excel(file_stream, header=None, nrows=2, engine="openpyxl")
    file_stream.seek(0)
    banner = False
    if not peek.empty:
        first = peek.iloc[0, 0]
        if isinstance(first, str) and _norm(first).startswith(_norm("Article filtré :")):
            banner = True
    if banner:
        return pd.read_excel(file_stream, skiprows=1, header=0, engine="openpyxl")
    return pd.read_excel(file_stream, header=0, engine="openpyxl")

# ---------- recherche article ----------
def build_article_pattern(user_input: str) -> re.Pattern:
    tok = (user_input or "").strip()
    if not tok:
        raise ValueError("Article vide.")
    esc = re.escape(tok)
    tail = r"(?![\d.])" if tok[-1].isdigit() else r"\b"
    return re.compile(rf"(?:\b(?:art(?:icle)?\s*[: ]*)?)({esc}){tail}", re.IGNORECASE)

# ---------- nettoyage / segmentation ----------
def _prep_text(v: str) -> str:
    if not isinstance(v, str):
        v = "" if v is None else str(v)
    # normalise les retours ligne (important pour les \n qui restaient)
    v = v.replace("\r\n", "\n").replace("\r", "\n")
    # enlève puces ASCII si déjà présentes pour éviter doubles puces
    v = v.replace("•", "\n").replace("·", "\n").replace("◦", "\n")
    # remplace NBSP & co
    v = v.replace("\u00A0", " ").replace("\u202F", " ")
    # compacte espaces sans toucher aux \n
    v = "\n".join(" ".join(line.split()) for line in v.split("\n"))
    return v.strip()

SEG_SPLIT_RE = re.compile(r"\n|;")  # séparateurs pour faire des puces

def split_segments(text: str) -> List[str]:
    clean = _prep_text(text)
    parts = [p.strip("-–•· \t") for p in SEG_SPLIT_RE.split(clean)]
    return [p for p in parts if p]

def html_bullets(text: str, pat: re.Pattern, only_hits: bool) -> str:
    items = []
    for seg in split_segments(text):
        if only_hits and not pat.search(seg):
            continue
        seg_html = html.escape(seg)
        seg_html = pat.sub(r'<span class="hit">\1</span>', seg_html)
        items.append(f"<li>{seg_html}</li>")
    return f"<ul>{''.join(items)}</ul>" if items else ""

def text_bullets(text: str, pat: re.Pattern, only_hits: bool) -> str:
    items = []
    for seg in split_segments(text):
        if only_hits and not pat.search(seg):
            continue
        items.append(f"• {seg}")
    return "\n".join(items)

# ---------- filtrage de lignes ----------
def build_row_mask(df: pd.DataFrame, colmap: Dict[str, Optional[str]], pat: re.Pattern):
    masks = []
    for k in INTEREST_KEYS:
        col = colmap.get(k)
        if col and col in df.columns:
            masks.append(df[col].astype(str).apply(lambda v: bool(pat.search(_prep_text(v)))))
    if not masks:
        return None
    m = masks[0]
    for x in masks[1:]:
        m = m | x
    return m

# ---------- rendu HTML + Export ----------
def format_for_display_and_export(
    df: pd.DataFrame, colmap: Dict[str, Optional[str]], pat: re.Pattern, segments_only: bool
):
    df_disp = df.copy()
    df_xlsx = df.copy()

    # colonnes à bulletiser (HTML et Excel)
    bullet_cols: List[str] = []
    for key in (ALWAYS_BULLET_KEYS + INTEREST_KEYS):
        col = colmap.get(key)
        if col and col in df.columns:
            bullet_cols.append(col)

    # 1) HTML (avec surbrillance)
    for col in bullet_cols:
        only_hits = segments_only and (col in [colmap.get(k) for k in INTEREST_KEYS])
        df_disp[col] = df_disp[col].apply(lambda v: html_bullets(v, pat, only_hits))

    # surbrillance également dans les colonnes non-listes (au cas où)
    for col in df.columns:
        if col in bullet_cols:
            continue
        df_disp[col] = df_disp[col].apply(
            lambda v: pat.sub(r'<span class="hit">\1</span>', html.escape(_prep_text(v)))
        )

    # 2) Excel texte (sans HTML)
    for col in bullet_cols:
        only_hits = segments_only and (col in [colmap.get(k) for k in INTEREST_KEYS])
        df_xlsx[col] = df_xlsx[col].apply(lambda v: text_bullets(v, pat, only_hits))

    return df_disp, df_xlsx

def to_excel_download(df_xlsx: pd.DataFrame) -> str:
    ts = int(time.time())
    out_path = f"/tmp/filtrage_{ts}.xlsx"
    with pd.ExcelWriter(out_path, engine="openpyxl") as w:
        df_xlsx.to_excel(w, index=False, sheet_name="Filtre")
        ws = w.book.active
        for c in range(1, ws.max_column + 1):
            col_letter = ws.cell(1, c).column_letter
            values = [ws.cell(r, c).value or "" for r in range(1, ws.max_row + 1)]
            max_len = max(len(str(v)) for v in values) if values else 12
            ws.column_dimensions[col_letter].width = min(80, max(16, max_len * 0.9))
    return f"/download?path={out_path}"

# ---------- routes ----------
@app.route("/", methods=["GET", "POST"])
def analyze():
    if request.method == "GET":
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK,
                                      table_html=None, searched_article=None,
                                      segments_only=False, message=None, message_ok=True)

    file = request.files.get("file")
    article = (request.form.get("article") or "").strip()
    segments_only = bool(request.form.get("segments_only"))

    if not file or not article:
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK,
                                      table_html=None, searched_article=article,
                                      segments_only=segments_only,
                                      message="Erreur : fichier et article requis.", message_ok=False)

    # .xlsx / .xlsm seulement
    name = (file.filename or "").lower()
    if not (name.endswith(".xlsx") or name.endswith(".xlsm")):
        ext = (file.filename or "").split(".")[-1]
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK,
                                      table_html=None, searched_article=article,
                                      segments_only=segments_only,
                                      message=f"Format {ext} non pris en charge. Utilisez .xlsx ou .xlsm.",
                                      message_ok=False)

    try:
        df = read_excel_respecting_header_rule(file.stream)
        colmap = resolve_columns(df)
        pat = build_article_pattern(article)

        mask = build_row_mask(df, colmap, pat)
        if mask is None:
            detail = "\n".join([f"  - {k}: {colmap.get(k)}" for k in INTEREST_KEYS])
            return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK,
                                          table_html=None, searched_article=article,
                                          segments_only=segments_only,
                                          message=("Aucune des colonnes d’intérêt n’a été trouvée.\n"
                                                   f"Résolution:\n{detail}\n\nColonnes:\n{list(df.columns)}"),
                                          message_ok=False)

        df_keep = df[mask].copy()
        if df_keep.empty:
            return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK,
                                          table_html=None, searched_article=article,
                                          segments_only=segments_only,
                                          message=f"Aucune ligne ne contient l’article « {article} » dans les colonnes cibles.",
                                          message_ok=True)

        # mise en puces + surbrillance
        df_disp, df_xlsx = format_for_display_and_export(df_keep, colmap, pat, segments_only)

        download_url = to_excel_download(df_xlsx)
        html_table = df_disp.to_html(index=False, escape=False)

        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK,
                                      table_html=html_table, searched_article=article,
                                      segments_only=segments_only,
                                      download_url=download_url,
                                      message=f"{len(df_keep)} ligne(s) retenue(s).", message_ok=True)

    except Exception as e:
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK,
                                      table_html=None, searched_article=article,
                                      segments_only=segments_only,
                                      message=f"Erreur inattendue : {repr(e)}", message_ok=False)

@app.route("/download")
def download():
    path = request.args.get("path")
    if not path or not os.path.exists(path):
        return "Fichier introuvable ou expiré.", 404
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
