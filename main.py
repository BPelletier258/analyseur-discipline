# === CANVAS META =============================================================
# Fichier : main.py — version listes + surlignage + anti-"nan"
# ============================================================================

import os
import re
import time
import unicodedata
from typing import Dict, Optional, Set

import pandas as pd
from flask import Flask, request, render_template_string, send_file

app = Flask(__name__)

# ──────────────────────────────────────────────────────────────────────────────
# STYLE + PAGE
# ──────────────────────────────────────────────────────────────────────────────

STYLE_BLOCK = """
<style>
  :root{
    --fg:#111827; --muted:#6b7280; --line:#e5e7eb; --soft:#f3f4f6; --brand:#0f766e; --danger:#b91c1c;
  }
  body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;margin:24px;color:var(--fg)}
  h1{font-size:20px;margin:0 0 12px}
  form{display:grid;gap:10px;margin:0 0 14px}
  input[type="text"]{padding:8px;border:1px solid var(--line);border-radius:6px;font-size:14px}
  input[type="file"]{font-size:14px}
  label{font-weight:600}
  button{padding:8px 12px;border:1px solid var(--line);border-radius:6px;background:#fff;cursor:pointer}
  .hint{font-size:12px;color:var(--muted)}
  .note{background:#fff7ed;border:1px solid #fed7aa;padding:10px;border-radius:8px;margin:10px 0 14px}
  .download{margin:12px 0}
  .kbd{font-family:ui-monospace,Menlo,monospace;background:var(--soft);padding:2px 4px;border-radius:4px}
  .msg{margin-top:12px;white-space:pre-wrap;font-family:ui-monospace,Menlo,monospace;font-size:12px}
  .ok{color:#065f46}.err{color:#7f1d1d}

  /* Vue avec barre de défilement horizontale si besoin */
  .table-viewport{height:60vh;overflow:auto;border:1px solid var(--line)}
  .table-wide{min-width:200vw}

  table{border-collapse:collapse;width:100%;font-size:13px}
  th,td{border:1px solid var(--line);padding:6px 8px;vertical-align:top}
  th{background:var(--soft);text-align:center}  /* TITRES CENTRÉS */
  td{text-align:left}
  ul.clean{margin:0;padding-left:18px}
  .hl{color:var(--danger);font-weight:600}
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
    <div>
      <label>Article à rechercher (ex. <span class="kbd">29</span>, <span class="kbd">59(2)</span>)</label>
      <input type="text" name="article" value="{{ searched_article or '' }}" required placeholder="ex.: 29 ou 59(2)" />
    </div>

    <div>
      <label>Fichier Excel</label>
      <input type="file" name="file" accept=".xlsx,.xlsm" required />
    </div>

    <label style="font-weight:500">
      <input type="checkbox" name="only_segments" value="1" {% if only_segments %}checked{% endif %}/>
      Afficher seulement le segment contenant l’article dans les colonnes ciblées
    </label>

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
# Normalisation d’en-têtes & alias
# ──────────────────────────────────────────────────────────────────────────────

def _norm(s: str) -> str:
    if not isinstance(s, str):
        s = str(s) if s is not None else ""
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = s.replace("\u00A0", " ")
    return " ".join(s.strip().lower().split())

HEADER_ALIASES: Dict[str, Set[str]] = {
    # 4 colonnes d’intérêt + liste chefs/articles
    "articles_enfreints": {
        _norm("Nbr Chefs par articles"),
        _norm("Articles enfreints"),
        _norm("Articles en infraction"),
        _norm("Liste des chefs et articles en infraction"),  # on l’inclut aussi (utile au filtrage)
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
# Lecture Excel (règle « Article filtré : »)
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
        return pd.read_excel(file_stream, skiprows=1, header=0, engine="openpyxl")
    return pd.read_excel(file_stream, header=0, engine="openpyxl")

# ──────────────────────────────────────────────────────────────────────────────
# Article → motif exact
# ──────────────────────────────────────────────────────────────────────────────

def build_article_pattern(user_input: str) -> re.Pattern:
    token = (user_input or "").strip()
    if not token:
        raise ValueError("Article vide.")
    esc = re.escape(token)
    tail = r"(?![\d.])" if token[-1].isdigit() else r"\b"
    return re.compile(rf"(?:\b(?:art(?:icle)?\s*[: ]*)?)({esc}){tail}", flags=re.IGNORECASE)

# ──────────────────────────────────────────────────────────────────────────────
# Utils “blankish”, dash, split, list-render & highlight
# ──────────────────────────────────────────────────────────────────────────────

def _blankish(x) -> bool:
    if x is None:
        return True
    try:
        if not isinstance(x, str) and pd.isna(x):
            return True
    except Exception:
        pass
    if isinstance(x, str) and x.strip().lower() in {"", "nan", "nat", "none"}:
        return True
    return False

def _dash_if_empty_html() -> str:
    return "—"

def _highlight_article(text: str, pat: Optional[re.Pattern]) -> str:
    if not pat or _blankish(text):
        return "" if _blankish(text) else str(text)
    def repl(m): return f"<span class='hl'>{m.group(1)}</span>"
    return pat.sub(repl, str(text))

def _split_to_items(raw) -> list[str]:
    if _blankish(raw):
        return []
    t = str(raw).replace("\r", "\n")
    t = t.replace("•", "\n").replace("·", "\n")
    # On scinde sur | ; ou retours
    parts = re.split(r"\n+|\s*\|\s*|;", t)
    items = []
    for p in parts:
        s = p.strip(" \t-•·—")
        if s and s.lower() not in {"nan","nat","none"}:
            items.append(s)
    return items

def render_list_cell(value, pat: Optional[re.Pattern], only_segments: bool, highlight: bool) -> str:
    items = _split_to_items(value)
    if only_segments and pat:
        items = [it for it in items if pat.search(it)]
    if not items:
        return _dash_if_empty_html()
    if highlight and pat:
        items = [_highlight_article(it, pat) for it in items]
    # liste HTML propre
    return "<ul class='clean'>" + "".join(f"<li>{it}</li>" for it in items) + "</ul>"

# ──────────────────────────────────────────────────────────────────────────────
# Préparation texte simple (pour le filtrage)
# ──────────────────────────────────────────────────────────────────────────────

def _prep_text(v) -> str:
    if _blankish(v):
        return ""
    s = str(v)
    s = s.replace("\u00A0", " ").replace("\r\n", "\n").replace("\r", "\n")
    # On neutralise le symbole puce pour le filtrage
    s = s.replace("•", " ").replace("·", " ").replace("◦", " ")
    s = " ".join(s.split())
    return s

# ──────────────────────────────────────────────────────────────────────────────
# Excel export
# ──────────────────────────────────────────────────────────────────────────────

def to_excel_download(df: pd.DataFrame) -> str:
    ts = int(time.time())
    out_path = f"/tmp/filtrage_{ts}.xlsx"
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Filtre")
        ws = writer.book.active
        for col_idx, col in enumerate(df.columns, start=1):
            max_len = max((len(str(x)) for x in [col] + df[col].astype(str).tolist()), default=10)
            ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = min(60, max(12, max_len+2))
    return f"/download?path={out_path}"

# ──────────────────────────────────────────────────────────────────────────────
# Routes
# ──────────────────────────────────────────────────────────────────────────────

@app.route("/", methods=["GET", "POST"])
def analyze():
    if request.method == "GET":
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK,
                                      table_html=None, searched_article=None,
                                      only_segments=False, message=None, message_ok=True)

    file = request.files.get("file")
    article = (request.form.get("article") or "").strip()
    only_segments = request.form.get("only_segments") == "1"

    if not file or not article:
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK,
                                      table_html=None, searched_article=article,
                                      only_segments=only_segments,
                                      message="Erreur : fichier et article sont requis.", message_ok=False)

    fname = (file.filename or "").lower()
    if not (fname.endswith(".xlsx") or fname.endswith(".xlsm")):
        ext = (file.filename or "").split(".")[-1]
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK,
                                      table_html=None, searched_article=article,
                                      only_segments=only_segments,
                                      message=f"Format non pris en charge : {ext}. Utilisez .xlsx ou .xlsm.",
                                      message_ok=False)

    try:
        df = read_excel_respecting_header_rule(file.stream)
        colmap = resolve_columns(df)
        pat = build_article_pattern(article)

        # Filtrage : ligne conservée si l’article est présent dans AU MOINS une colonne cible
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
                HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                searched_article=article, only_segments=only_segments,
                message=("Erreur : aucune des colonnes attendues n’a été trouvée.\n"
                         f"Colonnes résolues :\n{detail}\n\nColonnes disponibles :\n{list(df.columns)}"),
                message_ok=False
            )
        mask_any = masks[0]
        for m in masks[1:]:
            mask_any = mask_any | m

        df_filtered = df[mask_any].copy()
        if df_filtered.empty:
            return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK,
                                          table_html=None, searched_article=article,
                                          only_segments=only_segments,
                                          message=f"Aucune ligne ne contient l’article « {article} » dans les colonnes cibles.",
                                          message_ok=True)

        # ── Construction d’une vue d’affichage (listes + surlignage + anti-“nan”)
        df_display = df_filtered.copy()

        # Normaliser toutes les cellules vides (pas de "nan")
        df_display = df_display.where(pd.notna(df_display), "")
        df_display = df_display.applymap(lambda x: "" if _blankish(x) else str(x))

        # Colonnes à rendre en listes (nom normalisé)
        listify_norms = {
            _norm("Résumé des faits concis"),
            _norm("Liste des chefs et articles en infraction"),
            _norm("Liste des sanctions imposées"),
            _norm("Autres mesures ordonnées"),
            _norm("À vérifier"), _norm("A verifier"),
            _norm("Nbr Chefs par articles"),
            _norm("Nbr Chefs par articles par période de radiation"),
            _norm("Nombre de chefs par articles et total amendes"),
            _norm("Nombre de chefs par article ayant une réprimande"),
        }
        # Colonnes avec surlignage rouge
        highlight_norms = {
            _norm("Liste des chefs et articles en infraction"),
            _norm("Nbr Chefs par articles"),
            _norm("Nbr Chefs par articles par période de radiation"),
            _norm("Nombre de chefs par articles et total amendes"),
            _norm("Nombre de chefs par article ayant une réprimande"),
        }

        # Rendu colonne par colonne
        for col in df_display.columns:
            n = _norm(col)
            if n in listify_norms:
                df_display[col] = df_display[col].apply(
                    lambda v: render_list_cell(v, pat if n in highlight_norms else None,
                                               only_segments if n in highlight_norms else False,
                                               highlight=(n in highlight_norms))
                )
            else:
                # Pas de liste → texte simple, mais anti-“nan” + surlignage si colonne dans highlight_norms ?
                def _simple(v):
                    if _blankish(v):
                        return _dash_if_empty_html()
                    txt = str(v)
                    return _highlight_article(txt, pat) if n in highlight_norms else txt
                df_display[col] = df_display[col].apply(_simple)

        # Lien téléchargement = on exporte le dataframe filtré “brut” (pas la version HTML)
        download_url = to_excel_download(df_filtered)

        # HTML (avec échappement désactivé pour accepter nos <ul>/<span>)
        table_html = df_display.head(200).to_html(index=False, escape=False)

        return render_template_string(
            HTML_TEMPLATE, style_block=STYLE_BLOCK,
            table_html=table_html, searched_article=article,
            only_segments=only_segments, download_url=download_url,
            message=f"{len(df_filtered)} ligne(s) après filtrage. (Aperçu limité à 200 lignes.)",
            message_ok=True
        )

    except Exception as e:
        return render_template_string(
            HTML_TEMPLATE, style_block=STYLE_BLOCK,
            table_html=None, searched_article=article,
            only_segments=only_segments,
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
