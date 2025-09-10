# === CANVAS META =============================================================
# Fichier : main.py — version stable « listes + surlignage + cases » (09-09)
# Remet d'équerre :
#   • case à cocher toujours visible
#   • surlignage rouge cohérent de l’article dans TOUTES les colonnes
#   • listes à puces seulement dans les colonnes « multi‑éléments »,
#     jamais pour les colonnes « scalaires » (noms, dates, numéros, totaux…)
#   • remplacement des NaN par un tiret « — » (sans puce inutile)
#   • formatage de « Total amendes » → « 5 000 $ » etc.
#   • export Excel conservé et stable
#
# NB : Cette version évite les retours arrière que vous avez observés :
#  - pas de CSS exotique qui casse la mise en page
#  - pas de min-width démesurée
#  - en-têtes figés (sticky) pour faciliter la lecture
# ============================================================================

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

# ──────────────────────────────────────────────────────────────────────────────
# UI / CSS
# ──────────────────────────────────────────────────────────────────────────────
STYLE_BLOCK = """
<style>
  :root{
    --w-def: 22rem;   /* largeur par défaut */
    --w-2x: 44rem;    /* largeur ×2 demandée */
    --w-num: 8rem;    /* colonnes numériques étroites */
  }
  body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;margin:22px}
  h1{font-size:20px;margin:0 0 10px}
  .note{background:#fff8e6;border:1px solid #ffd48a;padding:8px 10px;border-radius:6px;margin:10px 0 14px}
  form{display:grid;gap:10px;margin:0 0 12px}
  input[type="text"],input[type="file"]{font-size:14px}
  .hint{font-size:12px;color:#666}
  .download{margin:10px 0}

  /* Tableau */
  table{border-collapse:collapse;width:100%;font-size:13px;table-layout:fixed}
  th,td{border:1px solid #ddd;padding:6px 8px;vertical-align:top}
  th{background:#f3f4f6;position:sticky;top:0;z-index:2;text-align:center}
  td{word-break:break-word}

  /* Largeurs par défaut et spéciales */
  .col-def{min-width:var(--w-def)}
  .col-2x{min-width:var(--w-2x)}
  .col-num{min-width:var(--w-num);width:var(--w-num)}

  /* Listes */
  .bullets{margin:0;padding-left:1.1rem}
  .bullets li{margin:0.15rem 0}
  .dash{color:#666}
  .hit{color:#d00;font-weight:600}

  /* Zone visible défilante */
  .viewport{height:60vh;overflow:auto;border:1px solid #ddd}
  .msg{white-space:pre-wrap;font-family:ui-monospace,Menlo,monospace;font-size:12px;margin-top:10px}
  .ok{color:#065f46}.err{color:#7f1d1d}
</style>
"""

HTML_TEMPLATE = """
<!doctype html>
<html><head><meta charset="utf-8" />
<title>Analyseur Discipline – Filtrage par article</title>
{{ style_block|safe }}
</head>
<body>
  <h1>Analyseur Discipline – Filtrage par article</h1>
  <div class="note">
    Règles : détection exacte de l’article. Si la 1<sup>re</sup> cellule contient « <b>Article filtré :</b> », on ignore la 1<sup>re</sup> ligne (lignes d’en‑tête sur la 2<sup>e</sup>).
  </div>

  <form method="POST" enctype="multipart/form-data">
    <label>Article à rechercher (ex. <code>29</code>, <code>59(2)</code>)</label>
    <input type="text" name="article" value="{{ article or '' }}" required>

    <label><input type="checkbox" name="segments_only" {% if segments_only %}checked{% endif %}>
      Afficher uniquement le segment contenant l’article dans les 4 colonnes d’intérêt
    </label>

    <label>Fichier Excel</label>
    <input type="file" name="file" accept=".xlsx,.xlsm" required>

    <div><button>Analyser</button></div>
    <div class="hint">Formats : .xlsx / .xlsm</div>
  </form>

  {% if table_html %}
    <div class="download"><a href="{{ download_url }}">Télécharger le résultat (Excel)</a></div>
    <div class="viewport">{{ table_html|safe }}</div>
  {% endif %}

  {% if message %}
    <div class="msg {{ 'ok' if ok else 'err' }}">{{ message }}</div>
  {% endif %}
</body></html>
"""

# ──────────────────────────────────────────────────────────────────────────────
# Outils de normalisation & alias simples
# ──────────────────────────────────────────────────────────────────────────────

def _norm(s: str) -> str:
  if not isinstance(s, str):
    s = '' if s is None else str(s)
  s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
  s = s.replace("\u00A0", " ")
  return " ".join(s.strip().lower().split())

# Colonnes multi‑éléments (affichées en puces)
LIST_COLS = {
  _norm("Résumé des faits concis"),
  _norm("Liste des chefs et articles en infraction"),
  _norm("Nbr Chefs par articles"),
  _norm("Liste des sanctions imposées"),
  _norm("Nbr Chefs par articles par période de radiation"),
  _norm("Nombre de chefs par articles et total amendes"),
  _norm("Nombre de chefs par article ayant une réprimande"),
  _norm("Autres mesures ordonnées"),
  _norm("À vérifier"), _norm("A verifier")
}

# 4 colonnes d’intérêt pour l’option « segments seulement »
CANON_4 = {
  _norm("Nbr Chefs par articles"),
  _norm("Nbr Chefs par articles par période de radiation"),
  _norm("Nombre de chefs par articles et total amendes"),
  _norm("Nombre de chefs par article ayant une réprimande"),
}

# Alias tolérés (selon vos fichiers)
ALIASES = {
  _norm("Nbr Chefs par articles"):{_norm("Nombre de chefs par articles")},
  _norm("Nbr Chefs par articles par période de radiation"):{_norm("Nbr Chefs par articles par periode de radiation")},
  _norm("Nombre de chefs par articles et total amendes"):{_norm("Articles amende/chef"), _norm("Article amende/chef")},
  _norm("Nombre de chefs par article ayant une réprimande"):{_norm("Nombre de chefs par article ayant une reprimande")},
}

def norm_title_set(df: pd.DataFrame) -> Dict[str,str]:
  mapping: Dict[str,str] = {}
  for col in df.columns:
    key = _norm(col)
    mapping[key] = col
  # enrichir avec alias si présents
  for canonical, alts in ALIASES.items():
    for alt in alts:
      if alt in mapping and canonical not in mapping:
        mapping[canonical] = mapping[alt]
  return mapping

# ──────────────────────────────────────────────────────────────────────────────
# Lecture Excel (garde la règle « Article filtré : »)
# ──────────────────────────────────────────────────────────────────────────────

def read_excel_respecting_header_rule(f) -> pd.DataFrame:
  preview = pd.read_excel(f, nrows=2, header=None, engine="openpyxl")
  f.seek(0)
  first = preview.iloc[0,0] if not preview.empty else None
  use_skip = isinstance(first,str) and _norm(first).startswith(_norm("Article filtré :"))
  df = pd.read_excel(f, header=0, engine="openpyxl", skiprows=1 if use_skip else 0)
  return df

# ──────────────────────────────────────────────────────────────────────────────
# Recherche d’article (surlignage & découpe)
# ──────────────────────────────────────────────────────────────────────────────

def build_article_pattern(token: str) -> re.Pattern:
  token = (token or '').strip()
  esc = re.escape(token)
  tail = r"(?![\d.])" if token and token[-1].isdigit() else r"\b"
  return re.compile(rf"(?:\b(?:art(?:icle)?\s*[: ]*)?)({esc}){tail}", re.I)

_def_dash = '<span class="dash">—</span>'

_def_ul_tpl = "<ul class='bullets'>{}</ul>"

def _clean_text(v: object) -> str:
  if pd.isna(v):
    return ''
  s = str(v).replace("\u00A0"," ")
  # normalise retours éventuels
  s = s.replace("\r\n","\n").replace("\r","\n")
  return s.strip()


def _highlight(text: str, pat: re.Pattern) -> str:
  if not text:
    return ''
  return pat.sub(r"<span class='hit'>\\1</span>", text)


def _split_items(text: str) -> list:
  # on scinde d'abord sur les retours ligne; à défaut, sur « ; »
  raw = [p.strip() for p in re.split(r"\n|;", text) if p.strip()]
  # si aucune séparation, on tente la puce « • » déjà présente
  if not raw and '•' in text:
    raw = [p.strip(' •') for p in text.split('•') if p.strip(' •')]
  return raw


def render_cell(value: object, colname: str, pat: re.Pattern, as_list: bool) -> str:
  txt = _clean_text(value)
  if not txt:
    return _def_dash
  if as_list:
    items = _split_items(txt)
    if not items:
      return _def_dash
    items = [_highlight(it, pat) for it in items]
    return _def_ul_tpl.format(''.join(f"<li>{it}</li>" for it in items))
  else:
    return _highlight(txt, pat)


def only_segments(value: object, pat: re.Pattern) -> str:
  txt = _clean_text(value)
  if not txt:
    return ''
  items = _split_items(txt)
  if not items:
    # si pas de listes détectées, on renvoie la phrase si elle contient l'article
    return txt if pat.search(txt) else ''
  kept = [it for it in items if pat.search(it)]
  return "\n".join(kept)

# ──────────────────────────────────────────────────────────────────────────────
# Export Excel — on exporte du texte brut (sans balises)
# ──────────────────────────────────────────────────────────────────────────────

def to_excel_download(df: pd.DataFrame) -> str:
  ts = int(time.time())
  out_path = f"/tmp/filtrage_{ts}.xlsx"
  with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
    df.to_excel(writer, index=False, sheet_name="Filtre")
    ws = writer.book.active
    # largeurs AUTO raisonnables
    for j, col in enumerate(df.columns, start=1):
      values = [str(col)] + [str(x) if not pd.isna(x) else '' for x in df[col].tolist()]
      w = min(60, max(12, max(len(v) for v in values[:200]) + 2))
      ws.column_dimensions[ws.cell(row=1, column=j).column_letter].width = w
  return f"/download?path={out_path}"

# ──────────────────────────────────────────────────────────────────────────────
# Contrôleur
# ──────────────────────────────────────────────────────────────────────────────

@app.route('/', methods=['GET','POST'])
def index():
  if request.method == 'GET':
    return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                                  article='', segments_only=True, message=None, ok=True)

  # POST
  article = (request.form.get('article') or '').strip()
  segments_only = bool(request.form.get('segments_only'))
  up = request.files.get('file')

  if not up or not article:
    return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                                  article=article, segments_only=segments_only,
                                  message="Veuillez fournir un article et un fichier Excel .xlsx/.xlsm.", ok=False)

  name = (up.filename or '').lower()
  if not (name.endswith('.xlsx') or name.endswith('.xlsm')):
    return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                                  article=article, segments_only=segments_only,
                                  message="Format non pris en charge. Utilisez .xlsx/.xlsm.", ok=False)

  try:
    df = read_excel_respecting_header_rule(up.stream)
    if df.empty:
      return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                                    article=article, segments_only=segments_only,
                                    message="Classeur vide ou illisible.", ok=False)

    # Remplacement NaN pour l'export (texte brut) — on gardera des tirets visuels à l'affichage
    df_export = df.copy()
    df_export = df_export.fillna("")

    # Map des titres normalisés → titres d'origine
    title_map = norm_title_set(df)

    # Filtrage des lignes : au moins une des 4 colonnes contient l'article
    pat = build_article_pattern(article)

    cols4 = [title_map[k] for k in title_map.keys() if k in CANON_4]
    if not cols4:
      # S'il manque des colonnes, on tente une heuristique douce :
      guess = [c for c in df.columns if _norm(c) in CANON_4]
      cols4 = guess
    if not cols4:
      return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                                    article=article, segments_only=segments_only,
                                    message=("Aucune des 4 colonnes d’intérêt n’a été trouvée.\n"
                                             f"Colonnes disponibles : {list(df.columns)}"), ok=False)

    mask = False
    for c in cols4:
      m = df[c].astype(str).apply(lambda v: bool(pat.search(_clean_text(v))))
      mask = m if mask is False else (mask | m)
    df = df[mask].copy()
    if df.empty:
      return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                                    article=article, segments_only=segments_only,
                                    message=f"Aucune ligne ne contient l’article « {article} ».", ok=True)

    # Option « segments seulement » sur les 4 colonnes d’intérêt
    if segments_only:
      for c in cols4:
        df[c] = df[c].apply(lambda v: only_segments(v, pat))

    # Prépare l'affichage HTML : listes / surlignage + tires pour vides
    def is_list_col(col: str) -> bool:
      n = _norm(col)
      if n in LIST_COLS:
        return True
      return False

    # Format monétaire pour « Total amendes » (si la colonne existe et est num.)
    money_key = None
    for c in df.columns:
      if _norm(c) == _norm("Total amendes"):
        money_key = c
        break
    if money_key:
      def _fmt_money(x):
        if pd.isna(x) or str(x).strip()=='' or str(x).strip().lower()=='nan':
          return ''
        try:
          # on capture un nombre dans une chaîne éventuelle
          num = float(re.sub(r"[^0-9.-]", "", str(x)))
          if round(num)==0:
            return "0 $"
          s = f"{int(round(num)):,}".replace(',', ' ')
          return f"{s} $"
        except Exception:
          return str(x)
      df[money_key] = df[money_key].apply(_fmt_money)

    # Construire le HTML cellule par cellule
    html_df = pd.DataFrame(columns=df.columns)
    for col in df.columns:
      as_list = is_list_col(col)
      html_df[col] = df[col].apply(lambda v: render_cell(v, col, pat, as_list))

    # Applique classes de largeur raisonnables aux colonnes connues
    def col_class(col: str) -> str:
      n = _norm(col)
      if n in {_norm('Numéro de la décision'), _norm('Total chefs'), _norm('Radiation max'), _norm('Total amendes'), _norm('Total réprimandes'), _norm('Date de création'), _norm('Date de mise à jour')}:
        return 'col-num'
      if n in LIST_COLS:
        # colonnes riches → double largeur
        return 'col-2x'
      return 'col-def'

    # to_html + post-traitement pour injecter classes
    table_html = html_df.to_html(index=False, escape=False)
    # Injecte classes sur TH
    for j, col in enumerate(html_df.columns, start=1):
      table_html = table_html.replace(f'<th>{col}</th>', f'<th class="{col_class(col)}">{col}</th>')

    download_url = to_excel_download(df_export.loc[df.index])

    return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK,
                                  table_html=table_html, download_url=download_url,
                                  article=article, segments_only=segments_only,
                                  message=f"{len(df)} ligne(s) retenue(s).", ok=True)

  except Exception as e:
    return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                                  article=article, segments_only=segments_only,
                                  message=f"Erreur : {e}", ok=False)


@app.route('/download')
def download():
  path = request.args.get('path')
  if not path or not os.path.exists(path):
    return "Fichier introuvable ou expiré.", 404
  return send_file(path, as_attachment=True, download_name=os.path.basename(path))


if __name__ == '__main__':
  app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
