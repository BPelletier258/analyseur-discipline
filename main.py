# === CANVAS META =============================================================
# Fichier : main.py — largeur de colonnes + listes + Excel (09-09)
# ============================================================================
# Changements clés
# - ✅ Case à cocher « segment seul » rétablie
# - ✅ Listes à puces stables (\n ou • ⇒ <ul><li>)
# - ✅ "nan" → tiret (sans puce)
# - ✅ Surlignage rouge de l'article dans 5 colonnes (incluant « Liste des chefs… »)
# - ✅ Largeurs de colonnes figées via <colgroup> (voir WIDTH_RULES)
# - ✅ Total amendes formaté « 5 000 $ » (UI) + format Excel (# ##0 "$")
# - ✅ Export Excel : 1ʳᵉ ligne « Article filtré : X », entête figée, retour à la ligne et colonnes élargies
# - ✅ Téléchargement robuste

import io
import os
import re
import time
import unicodedata
from datetime import datetime
from typing import Dict, Optional, Set, List

import numpy as np
import pandas as pd
from flask import Flask, request, render_template_string, send_file

app = Flask(__name__)

# ──────────────────────────────────────────────────────────────────────────────
# UI / CSS
# ──────────────────────────────────────────────────────────────────────────────

STYLE_BLOCK = """
<style>
  :root{
    --w-def: 22rem;       /* largeur par défaut */
    --w-2x:  44rem;       /* largeur ×2 demandée */
    --w-num: 8rem;        /* colonnes numériques étroites */
  }
  body { font-family: system-ui,-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif; margin:22px }
  h1 { font-size:20px; margin:0 0 10px }
  .note{background:#fff8e6;border:1px solid #ffd48a;padding:8px 10px;border-radius:6px;margin:8px 0 14px}
  form{display:grid;gap:10px;margin:0 0 12px}
  input[type="text"],input[type="file"]{font-size:14px}
  .hint{font-size:12px;color:#666}
  .download{margin:12px 0}

  table{border-collapse:collapse;width:100%;font-size:13px;table-layout:fixed}
  th,td{border:1px solid #ddd;padding:6px 8px;vertical-align:top}
  th{background:#f3f4f6;text-align:center}
  .bullets{margin:0;padding-left:1.1rem}
  .bullets li{margin:0}
  .dash{color:#666}
  .hit{color:#d00;font-weight:600}

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
    <div>
      <label>Article à rechercher (ex. <code>29</code>, <code>59(2)</code>)</label><br>
      <input type="text" name="article" value="{{ searched_article or '' }}" required style="width:18rem" placeholder="ex.: 29 ou 59(2)" />
    </div>
    <label><input type="checkbox" name="segment_only" {% if segment_only %}checked{% endif %}> Afficher uniquement le segment contenant l’article dans les 4 colonnes d’intérêt</label>
    <div>
      <label>Fichier Excel</label><br>
      <input type="file" name="file" accept=".xlsx,.xlsm" required />
    </div>
    <button type="submit">Analyser</button>
    <div class="hint">Formats : .xlsx / .xlsm</div>
  </form>

  {% if table_html %}
    <div class="download"><a href="{{ download_url }}">Télécharger le résultat (Excel)</a></div>
    <div class="viewport">{{ table_html|safe }}</div>
  {% endif %}

  {% if message %}
    <div class="msg {{ 'ok' if message_ok else 'err' }}">{{ message }}</div>
  {% endif %}
</body></html>
"""

# ──────────────────────────────────────────────────────────────────────────────
# Normalisation & alias d’en‑têtes
# ──────────────────────────────────────────────────────────────────────────────

def _norm(s: str) -> str:
    if not isinstance(s, str):
        s = "" if s is None else str(s)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = s.replace("\u00A0", " ")
    return " ".join(s.strip().lower().split())

HEADER_ALIASES: Dict[str, Set[str]] = {
    "articles_enfreints": {_norm("Nbr Chefs par articles"), _norm("Articles enfreints"), _norm("Articles en infraction"), _norm("Liste des chefs et articles en infraction")},
    "duree_totale_radiation": {_norm("Nbr Chefs par articles par période de radiation"), _norm("Duree totale effective radiation")},
    "article_amende_chef": {_norm("Nombre de chefs par articles et total amendes"), _norm("Article amende/chef")},
    "autres_sanctions": {_norm("Nombre de chefs par article ayant une réprimande"), _norm("Autres sanctions")},
    "numero_decision": {_norm("Numéro de la décision"), _norm("Numero de decision"), _norm("Decision #")},
}

FILTER_CANONICAL = ["articles_enfreints", "duree_totale_radiation", "article_amende_chef", "autres_sanctions"]

TEXT_LIST_COLUMNS_HINTS = {
    # colonnes à puces (si présentes)
    _norm("Résumé des faits concis"),
    _norm("Liste des chefs et articles en infraction"),
    _norm("Liste des sanctions imposées"),
    _norm("Nbr Chefs par articles"),
    _norm("Nbr Chefs par articles par période de radiation"),
    _norm("Nombre de chefs par articles et total amendes"),
    _norm("Nombre de chefs par article ayant une réprimande"),
    _norm("Autres mesures ordonnées"),
    _norm("À vérifier"), _norm("A vérifier"), _norm("A verifier"),
}

# Largeurs désirées par entête (logique : défaut, ×2, numérique)
WIDTH_RULES = {
    # Doubler la largeur de ces colonnes
    _norm("Résumé des faits concis"): "2x",
    _norm("Autres mesures ordonnées"): "2x",
    _norm("Date de création"): "2x",
    _norm("Date de mise à jour"): "2x",
    _norm("Numéro de la décision"): "2x",
    _norm("Nombre de chefs par articles et total amendes"): "2x",
    _norm("Nbr Chefs par articles"): "2x",
    _norm("Nbr Chefs par articles par période de radiation"): "2x",
    _norm("Nombre de chefs par article ayant une réprimande"): "2x",

    # Colonnes numériques étroites
    _norm("Total chefs"): "num",
    _norm("Radiation max"): "num",
    _norm("Total amendes"): "num",
    _norm("Total réprimandes"): "num",
}

# ──────────────────────────────────────────────────────────────────────────────
# Résolution colonnes, lecture Excel
# ──────────────────────────────────────────────────────────────────────────────

def resolve_columns(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    n2o = {_norm(c): c for c in df.columns}
    out: Dict[str, Optional[str]] = {}
    for canon, variants in HEADER_ALIASES.items():
        hit = None
        for v in variants:
            if v in n2o:
                hit = n2o[v]; break
        out[canon] = hit
    return out

def read_excel_respecting_header_rule(file_stream) -> pd.DataFrame:
    prv = pd.read_excel(file_stream, header=None, nrows=2, engine="openpyxl")
    file_stream.seek(0)
    banner = prv.iloc[0,0] if not prv.empty else None
    skip = 1 if isinstance(banner,str) and _norm(banner).startswith(_norm("Article filtré :")) else 0
    return pd.read_excel(file_stream, header=0, skiprows=skip, engine="openpyxl")

# ──────────────────────────────────────────────────────────────────────────────
# Motif de détection + pré‑traitement
# ──────────────────────────────────────────────────────────────────────────────

def build_article_pattern(token: str) -> re.Pattern:
    token = (token or "").strip()
    if not token:
        raise ValueError("Article vide")
    esc = re.escape(token)
    tail = r"(?![\d.])" if token[-1].isdigit() else r"\b"
    return re.compile(rf"(?:\b(?:art(?:icle)?\s*[: ]*)?)({esc}){tail}", re.I)

def _prep_text(v) -> str:
    if pd.isna(v):
        return ""
    s = str(v)
    s = s.replace("•"," \u2022 ")
    s = s.replace("\u00A0"," ").replace("\u202F"," ")
    s = s.replace("\r\n","\n").replace("\r","\n")
    return s

# ──────────────────────────────────────────────────────────────────────────────
# Nettoyage / extraction
# ──────────────────────────────────────────────────────────────────────────────

def _extract_segments(text: str, pat: re.Pattern) -> str:
    if not text: return ""
    parts = re.split(r"[;\n]", text)
    hits = [p.strip() for p in parts if pat.search(p)]
    return " | ".join(hits)

# Mise en forme de listes UL à partir de séparateurs (\n, •, ;, |)
LIST_SPLIT_RE = re.compile(r"\n|\u2022|;|\|")

def _as_bullets_html(s: str) -> str:
    s = (s or "").strip()
    if not s:
        return '<span class="dash">—</span>'
    items = [i.strip(" -•\t") for i in LIST_SPLIT_RE.split(s) if i and i.strip(" -•\t")]
    if not items:
        return '<span class="dash">—</span>'
    if len(items) == 1:
        return items[0]
    lis = ''.join(f"<li>{i}</li>" for i in items)
    return f"<ul class=\"bullets\">{lis}</ul>"

# Surlignage rouge

def _highlight(s: str, pat: re.Pattern) -> str:
    if not s: return s
    return pat.sub(r"<span class=\"hit\">\\1</span>", s)

# Fabrique un HTML de tableau avec <colgroup> pour figer les largeurs

def dataframe_to_html_fixed(df: pd.DataFrame) -> str:
    # Remplacement NA visuel
    df2 = df.replace({np.nan: ""})
    raw = df2.to_html(index=False, escape=False)

    # Extraire les en-têtes pour appliquer les règles de largeur
    headers: List[str] = []
    m = re.search(r"<thead>.*?</thead>", raw, re.S)
    if m:
        thead = m.group(0)
        headers = re.findall(r"<th>(.*?)</th>", thead)
    widths: List[str] = []
    for h in headers:
        key = _norm(re.sub(r"<.*?>", "", h))
        rule = WIDTH_RULES.get(key, "def")
        if rule == "2x": widths.append("style=\"width:var(--w-2x)\"")
        elif rule == "num": widths.append("style=\"width:var(--w-num)\"")
        else: widths.append("style=\"width:var(--w-def)\"")
    colgroup = "<colgroup>" + "".join([f"<col {w}>" for w in widths]) + "</colgroup>"
    # Insérer juste après <table>
    html = re.sub(r"<table(.*?)>", lambda m: f"<table{m.group(1)}>{colgroup}", raw, count=1)
    return html

# ──────────────────────────────────────────────────────────────────────────────
# Export Excel (1ʳᵉ ligne, freeze panes, wrap, formats)
# ──────────────────────────────────────────────────────────────────────────────

def to_excel_download(df: pd.DataFrame, article: str) -> str:
    ts = int(time.time())
    out_path = f"/tmp/filtrage_{ts}.xlsx"
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        sheet = "Filtre"
        # On écrit à partir de la ligne 2 (on garde la 1ʳᵉ pour la bannière)
        df.to_excel(writer, index=False, sheet_name=sheet, startrow=1)
        ws = writer.book[sheet]
        # Ligne 1 : bannière
        ws.cell(row=1, column=1, value=f"Article filtré : {article}")
        # Largeurs + wrap + freeze
        from openpyxl.utils import get_column_letter
        from openpyxl.styles import Alignment, Font, PatternFill
        ws.freeze_panes = "A3"  # fige ligne d'entête
        # Mise en forme amendes (si col existe)
        numfmt = '# ##0 "$"'  # espace fine insécable + $
        for j, col in enumerate(df.columns, start=1):
            letter = get_column_letter(j)
            # largeur : heuristique 14..60
            max_len = max([len(str(col))] + [len(str(v)) for v in df[col].astype(str).tolist()])
            ws.column_dimensions[letter].width = max(14, min(60, max_len + 2))
            # wrap
            for i in range(2, 2 + len(df)):
                ws.cell(row=i+1, column=j).alignment = Alignment(wrap_text=True, vertical="top")
            if _norm(col) == _norm("Total amendes"):
                for i in range(2, 2 + len(df)):
                    cell = ws.cell(row=i+1, column=j)
                    # si numérique, appliquer le format
                    try:
                        float(cell.value)
                        cell.number_format = numfmt
                    except Exception:
                        pass
        # Style de la bannière
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(1, len(df.columns)))
        ws.cell(row=1, column=1).font = Font(bold=True)
        ws.cell(row=1, column=1).fill = PatternFill("solid", fgColor="FFFDE7")
    return f"/download?path={out_path}"

# ──────────────────────────────────────────────────────────────────────────────
# Contrôleur
# ──────────────────────────────────────────────────────────────────────────────

@app.route("/", methods=["GET","POST"])
def analyze():
    if request.method == "GET":
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                                      searched_article="", segment_only=True,
                                      message=None, message_ok=True)

    file = request.files.get("file")
    article = (request.form.get("article") or "").strip()
    segment_only = request.form.get("segment_only") is not None

    if not file or not article:
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                                      searched_article=article, segment_only=segment_only,
                                      message="Erreur : fichier et article sont requis.", message_ok=False)

    fname = (file.filename or "").lower()
    if not (fname.endswith(".xlsx") or fname.endswith(".xlsm")):
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                                      searched_article=article, segment_only=segment_only,
                                      message="Format non supporté. Fournissez un .xlsx ou .xlsm.", message_ok=False)
    try:
        df = read_excel_respecting_header_rule(file.stream)
        colmap = resolve_columns(df)
        pat = build_article_pattern(article)

        # Filtrage des lignes : l'article doit apparaître au moins dans 1 des 4 colonnes cibles
        masks = []
        any_cols = False
        for canon in FILTER_CANONICAL:
            col = colmap.get(canon)
            if col and col in df.columns:
                any_cols = True
                masks.append(df[col].astype(str).apply(lambda v: bool(pat.search(_prep_text(v)))))
        if not any_cols:
            return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                                          searched_article=article, segment_only=segment_only,
                                          message=("Aucune des colonnes cibles n'a été trouvée.\n"
                                                   f"Résolution: {colmap}\n\nColonnes: {list(df.columns)}"),
                                          message_ok=False)
        mask_any = masks[0]
        for m in masks[1:]: mask_any = mask_any | m
        df = df[mask_any].copy()
        if df.empty:
            return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                                          searched_article=article, segment_only=segment_only,
                                          message=f"Aucune ligne ne contient l’article « {article} ».",
                                          message_ok=True)

        # Préparation du rendu :
        # 1) Colonnes d'intérêt : soit segments seuls, soit texte complet avec surlignage
        for canon in FILTER_CANONICAL:
            col = colmap.get(canon)
            if not col or col not in df.columns: continue
            if segment_only:
                df[col] = df[col].apply(lambda v: _extract_segments(_prep_text(v), pat))
            else:
                df[col] = df[col].apply(lambda v: _highlight(_prep_text(v), pat))

        # 2) Surligner aussi dans « Liste des chefs et articles en infraction » si présente
        inf_col_norm = _norm("Liste des chefs et articles en infraction")
        for c in df.columns:
            if _norm(c) == inf_col_norm:
                df[c] = df[c].apply(lambda v: _highlight(_prep_text(v), pat))
                break

        # 3) Total amendes : format UI "5 000 $" (sans toucher au type Excel, traité à l'export)
        for c in df.columns:
            if _norm(c) == _norm("Total amendes"):
                def _fmt_amt(x):
                    if pd.isna(x) or str(x).strip()=="":
                        return '<span class="dash">—</span>'
                    try:
                        n = float(str(x).replace(" ", "").replace("$", ""))
                        s = f"{int(round(n)):,}".replace(","," ")
                        return f"{s} $"
                    except Exception:
                        return str(x)
                df[c] = df[c].apply(_fmt_amt)
                break

        # 4) Listes à puces sur les colonnes textuelles
        n2o = {_norm(c): c for c in df.columns}
        for norm_name in TEXT_LIST_COLUMNS_HINTS:
            col = n2o.get(norm_name)
            if not col: continue
            df[col] = df[col].apply(lambda s: _as_bullets_html(_prep_text(s)))

        # Remplacer NA résiduels par tiret
        df = df.replace({np.nan: ""})
        for c in df.columns:
            df[c] = df[c].apply(lambda v: v if (isinstance(v,str) and v.strip()!="") else '<span class="dash">—</span>')

        # HTML + largeurs fixes via <colgroup>
        table_html = dataframe_to_html_fixed(df)

        # Excel
        download_url = to_excel_download(_df_for_excel(df), article)

        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK,
                                      table_html=table_html, download_url=download_url,
                                      searched_article=article, segment_only=segment_only,
                                      message=f"{len(df)} ligne(s) retenue(s).", message_ok=True)

    except Exception as e:
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                                      searched_article=article, segment_only=True,
                                      message=f"Erreur inattendue : {repr(e)}", message_ok=False)

# Convertit le DF HTML (avec <ul>) en DF propre pour Excel (texte simple + puces)

def _df_for_excel(df_html: pd.DataFrame) -> pd.DataFrame:
    def strip_html(s: str) -> str:
        if not isinstance(s, str): return "" if s is None else str(s)
        # Remplacer listes
        s = s.replace("</li>", "\n").replace("<li>", "• ")
        s = re.sub(r"<.*?>", "", s)
        return s.strip()
    out = df_html.copy()
    for c in out.columns:
        out[c] = out[c].apply(strip_html)
    return out

@app.route("/download")
def download():
    path = request.args.get("path")
    if not path or not os.path.exists(path):
        return "Fichier introuvable ou expiré.", 404
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
