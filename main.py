# -*- coding: utf-8 -*-
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

# ───────────────────────────────── CSS ─────────────────────────────────

STYLE_BLOCK = """
<style>
  :root{
    /* largeur de base d'une colonne "normale" */
    --w-def: 22rem;
    /* variantes */
    --w-05x: 11rem;   /* 0.5x */
    --w-15x: 33rem;   /* 1.5x */
    --w-2x:  44rem;   /* 2x   */
    --w-num: 8rem;    /* colonnes numériques compactes */
  }
  body { font-family: system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif; margin: 24px; }
  h1 { font-size: 20px; margin-bottom: 12px; }
  form { display: grid; gap: 12px; margin-bottom: 16px; }
  input[type="text"] { padding: 8px; font-size: 14px; }
  input[type="file"] { font-size: 14px; }
  button { padding: 8px 12px; font-size: 14px; cursor: pointer; }
  .hint { font-size: 12px; color: #666; }
  .note { background: #fff6e5; border: 1px solid #ffd89b; padding: 8px 10px; border-radius: 6px; margin: 10px 0 16px; }
  .download { margin: 12px 0; }

  /* Tableau + viewport scrollable horizontal/vertical */
  .table-viewport{height:60vh; overflow:auto; border:1px solid #ddd;}
  .table-viewport table{border-collapse:collapse; width:100%; font-size:13px;}
  th, td { border: 1px solid #ddd; padding: 6px 8px; vertical-align: top; }
  th { background: #f3f4f6; text-align:center; }

  /* Puces */
  .bullets ul{ margin:0; padding-left:1.2rem; }
  .bullets li{ margin:0.15rem 0; }
  .dash { color:#d00; font-weight:600; } /* mise en évidence de l'article */

  /* Largeurs par <colgroup> */
  col.col-def{min-width:var(--w-def); width:var(--w-def);}
  col.col-05{min-width:var(--w-05x); width:var(--w-05x);}
  col.col-15{min-width:var(--w-15x); width:var(--w-15x);}
  col.col-2x{min-width:var(--w-2x);  width:var(--w-2x);}
  col.col-num{min-width:var(--w-num); width:var(--w-num);}
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
    Règles : détection exacte de l’article; si la 1<sup>re</sup> cellule contient « <code>Article filtré :</code> », on ignore la 1<sup>re</sup> ligne (entêtes sur la 2<sup>e</sup>).
  </div>

  <form method="POST" enctype="multipart/form-data">
    <label>Article à rechercher (ex. <code>29</code>, <code>59(2)</code>)</label>
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
    <div class="table-viewport">
      {{ table_html|safe }}
    </div>
  {% endif %}

  {% if message %}
    <pre class="hint">{{ message }}</pre>
  {% endif %}
</body>
</html>
"""

# ─────────────────────────── Normalisation titres ───────────────────────────

def _norm(s: str) -> str:
    if not isinstance(s, str):
        s = str(s) if s is not None else ""
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = s.replace("\u00A0", " ")
    return " ".join(s.strip().lower().split())

HEADER_ALIASES: Dict[str, Set[str]] = {
    "articles_enfreints": {
        _norm("Nbr Chefs par articles"),
        _norm("Articles enfreints"),
        _norm("Articles en infraction"),
        _norm("Liste des chefs et articles en infraction"),
    },
    "duree_totale_radiation": {
        _norm("Nbr Chefs par articles par période de radiation"),
        _norm("Durée totale effective radiation"),
    },
    "article_amende_chef": {
        _norm("Nombre de chefs par articles et total amendes"),
        _norm("Article amende/chef"),
    },
    "autres_sanctions": {
        _norm("Nombre de chefs par article ayant une réprimande"),
        _norm("Autres mesures ordonnées"),
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
    norm_to_original = {_norm(c): c for c in df.columns}
    out: Dict[str, Optional[str]] = {}
    for canon, choices in HEADER_ALIASES.items():
        hit = None
        for v in choices:
            if v in norm_to_original:
                hit = norm_to_original[v]
                break
        out[canon] = hit
    return out

# ─────────────────────────── Lecture Excel (bannière) ───────────────────────────

def read_excel_respecting_header_rule(file_stream) -> pd.DataFrame:
    df2 = pd.read_excel(file_stream, header=None, nrows=2, engine="openpyxl")
    file_stream.seek(0)
    first_cell = df2.iloc[0, 0] if not df2.empty else None
    banner = isinstance(first_cell, str) and _norm(first_cell).startswith(_norm("Article filtré :"))
    if banner:
        return pd.read_excel(file_stream, skiprows=1, header=0, engine="openpyxl")
    return pd.read_excel(file_stream, header=0, engine="openpyxl")

# ───────────────────────── Motif exact pour l’article ─────────────────────────

def build_article_pattern(user_input: str) -> re.Pattern:
    token = (user_input or "").strip()
    if not token:
        raise ValueError("Article vide.")
    esc = re.escape(token)
    tail = r"(?![\d.])" if token[-1].isdigit() else r"\b"
    return re.compile(rf"(?:\b(?:art(?:icle)?\s*[: ]*)?)({esc}){tail}", re.IGNORECASE)

# ───────────────────────── Helpers nettoyage / format ─────────────────────────

def _prep_text(v: str) -> str:
    if not isinstance(v, str):
        v = "" if v is None else str(v)
    # Harmonise espaces & retours
    v = v.replace("•", " ").replace("·", " ").replace("◦", " ")
    v = v.replace("\u00A0", " ").replace("\u202F", " ")
    v = v.replace("\r\n", "\n").replace("\r", "\n")
    v = " ".join(v.split())
    return v

def _highlight_article(txt: str, pat: re.Pattern) -> str:
    # met UNIQUEMENT le token en rouge, pas tout le segment
    return pat.sub(lambda m: f'<span class="dash">{m.group(1)}</span>', txt)

def _as_bullets(val: str, pat: re.Pattern) -> str:
    """Transforme 'a\nb | c' => <ul><li>a</li><li>b</li><li>c</li></ul>. Si vide -> '—'."""
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return "—"
    s = str(val).strip()
    if not s or s.lower() == "nan":
        return "—"
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    parts = [p.strip(" •\t") for p in re.split(r"(?:\n|\|)", s) if p.strip(" •\t")]
    if not parts:
        return "—"
    items = "".join(f"<li>{_highlight_article(p, pat)}</li>" for p in parts)
    return f"<ul>{items}</ul>"

def fmt_money(x) -> str:
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return "—"
    try:
        n = int(float(x))
    except Exception:
        return str(x)
    return f"{n:,}".replace(",", " ") + " $"

def bulletize_columns(df: pd.DataFrame, columns: List[str], pat: re.Pattern) -> pd.DataFrame:
    """Applique _as_bullets sur les colonnes textuelles indiquées."""
    out = df.copy()
    for col in columns:
        if col in out.columns:
            out[col] = out[col].apply(lambda v: _as_bullets(v, pat))
    return out

# ───────────────────────── Extraction nettoyée (colonnes cibles) ─────────────────────────

def _extract_generic(text: str, pat: re.Pattern) -> str:
    if not isinstance(text, str) or not text.strip():
        return ""
    parts = re.split(r"[;,\n]", text)
    hits = [p.strip() for p in parts if pat.search(p)]
    return " | ".join(hits)

def _extract_autres(text: str, pat: re.Pattern) -> str:
    if not isinstance(text, str) or not text.strip():
        return ""
    parts = [seg.strip() for seg in re.split(r"[;\n]", text) if pat.search(seg)]
    return " | ".join(parts)

def clean_filtered_df(df: pd.DataFrame, colmap: Dict[str, Optional[str]], pat: re.Pattern) -> pd.DataFrame:
    df = df.copy()
    for canon in FILTER_CANONICAL:
        col = colmap.get(canon)
        if not col or col not in df.columns:
            continue
        if canon == "autres_sanctions":
            df[col] = df[col].astype(str).apply(lambda v: _extract_autres(_prep_text(v), pat))
        else:
            df[col] = df[col].astype(str).apply(lambda v: _extract_generic(_prep_text(v), pat))
    # Garde lignes où au moins une de ces colonnes contient quelque chose
    subset_cols = [c for c in (colmap.get(k) for k in FILTER_CANONICAL) if c]
    if subset_cols:
        mask = False
        for c in subset_cols:
            cur = df[c].astype(str).str.strip().ne("")
            mask = cur if mask is False else (mask | cur)
        df = df[mask]
    return df

# ───────────────────────── Largeurs par colonne (colgroup) ─────────────────────────

# 👉👉 RÈGLE LES LARGEURS ICI :
#    'col-def' = 1× (par défaut) ; 'col-05' = 0.5× ; 'col-15' = 1.5× ; 'col-2x' = 2× ; 'col-num' = compacte
WIDTH_CLASSES = {
    _norm("Nom de l'intimé"): "col-def",
    _norm("Ordre professionnel"): "col-def",
    _norm("Numéro de la décision"): "col-15",            # ex. 1.5×
    _norm("Date de la décision rendue"): "col-15",        # ex. 1.5×
    _norm("Nature de la décision"): "col-def",
    _norm("Période des faits"): "col-def",
    _norm("Plaidoyer de culpabilité"): "col-05",          # 0.5×
    _norm("Résumé des faits concis"): "col-2x",           # 2×
    _norm("Liste des chefs et articles en infraction"): "col-2x",
    _norm("Nbr Chefs par articles"): "col-15",
    _norm("Total chefs"): "col-num",
    _norm("Liste des sanctions imposées"): "col-2x",
    _norm("Nbr Chefs par articles par période de radiation"): "col-15",
    _norm("Radiation max"): "col-05",
    _norm("Nombre de chefs par articles et total amendes"): "col-15",
    _norm("Total amendes"): "col-num",
    _norm("Nombre de chefs par article ayant une réprimande"): "col-15",
    _norm("Total réprimandes"): "col-num",
    _norm("Autres mesures ordonnées"): "col-2x",
    _norm("À vérifier"): "col-def",
    _norm("Date de création"): "col-15",
    _norm("Date de mise à jour"): "col-15",
}

def build_colgroup(cols: List[str]) -> str:
    pieces = []
    for c in cols:
        cls = WIDTH_CLASSES.get(_norm(c), "col-def")
        pieces.append(f'<col class="{cls}">')
    return "<colgroup>" + "".join(pieces) + "</colgroup>"

def inject_colgroup(table_html: str, colgroup_html: str) -> str:
    # insère le colgroup juste après le premier tag <table ...>
    return re.sub(r"(<table[^>]*>)", r"\\1" + colgroup_html, table_html, count=1, flags=re.IGNORECASE)

# ───────────────────────── Export Excel ─────────────────────────

def to_excel_download(df: pd.DataFrame, article: str) -> str:
    ts = int(time.time())
    out_path = f"/tmp/filtrage_{ts}.xlsx"
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        # Ligne 1 = bannière
        banner = pd.DataFrame([[f"Article filtré : {article}"]])
        banner.to_excel(writer, index=False, header=False, sheet_name="Filtre", startrow=0)
        # Données commencent ligne 2 (header=True)
        df.to_excel(writer, index=False, sheet_name="Filtre", startrow=1)
        ws = writer.book["Filtre"]
        # fige la ligne d'en-têtes (ligne 2 dans Excel)
        ws.freeze_panes = "A3"
        # largeur auto raisonnable
        for j, col in enumerate(df.columns, start=1):
            max_len = max((len(str(x)) for x in [col] + df[col].astype(str).tolist()), default=10)
            ws.column_dimensions[ws.cell(row=2, column=j).column_letter].width = min(60, max(12, max_len + 2))
    return f"/download?path={out_path}"

# ───────────────────────── Routes ─────────────────────────

@app.route("/", methods=["GET", "POST"])
def analyze():
    if request.method == "GET":
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK,
                                      table_html=None, searched_article=None,
                                      message=None, download_url=None)

    file = request.files.get("file")
    article = (request.form.get("article") or "").strip()

    if not file or not article:
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                                      searched_article=article, message="Erreur : fichier et article requis.",
                                      download_url=None)

    fname = (file.filename or "").lower()
    if not (fname.endswith(".xlsx") or fname.endswith(".xlsm")):
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                                      searched_article=article, message="Format non supporté (utiliser .xlsx/.xlsm).",
                                      download_url=None)

    try:
        df = read_excel_respecting_header_rule(file.stream)
        colmap = resolve_columns(df)
        pat = build_article_pattern(article)

        # Filtrage initial : au moins une colonne cible contient l'article
        masks = []
        for canon in FILTER_CANONICAL:
            col = colmap.get(canon)
            if col and col in df.columns:
                masks.append(df[col].astype(str).apply(lambda v: bool(pat.search(_prep_text(v)))))
        if not masks:
            return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                                          searched_article=article,
                                          message="Aucune des colonnes cibles n'a été trouvée.",
                                          download_url=None)

        mask_any = masks[0]
        for m in masks[1:]:
            mask_any = mask_any | m
        df_filtered = df[mask_any].copy()

        if df_filtered.empty:
            return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                                          searched_article=article,
                                          message=f"Aucune ligne ne contient l’article « {article} ».",
                                          download_url=None)

        # Nettoyage spécifique (isole les segments où l'article apparaît)
        df_clean = clean_filtered_df(df_filtered, colmap, pat)

        if df_clean.empty:
            return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                                          searched_article=article,
                                          message="Correspondances trouvées mais segments vides après nettoyage.",
                                          download_url=None)

        # Format "Total amendes" si présent
        for c in df_clean.columns:
            if _norm(c) == _norm("Total amendes"):
                df_clean[c] = df_clean[c].apply(fmt_money)

        # Colonnes à afficher sous forme de puces (tu peux en ajouter/enlever librement)
        LIST_COLS = [
            "Résumé des faits concis",
            "Liste des chefs et articles en infraction",
            "Nbr Chefs par articles",
            "Nbr Chefs par articles par période de radiation",
            "Nombre de chefs par articles et total amendes",
            "Nombre de chefs par article ayant une réprimande",
            "Liste des sanctions imposées",
            "Autres mesures ordonnées",
            "À vérifier",
        ]
        list_cols_present = [c for c in df_clean.columns if _norm(c) in {_norm(x) for x in LIST_COLS}]

        # Puces (transforme \n / | en <ul><li>…</li></ul>) et mise en évidence du token
        df_view = bulletize_columns(df_clean, list_cols_present, pat)

        # HTML + colgroup (largeurs stables) + pas de "nan"
        preview = df_view.head(200)
        base_table = preview.to_html(index=False, escape=False, na_rep="—", classes=["dataframe", "bullets"])
        colgroup = build_colgroup(list(preview.columns))
        table_html = inject_colgroup(base_table, colgroup)

        download_url = to_excel_download(df_view, article)

        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK,
                                      table_html=table_html, searched_article=article,
                                      message=f"{len(df_view)} ligne(s) – aperçu limité à 200.",
                                      download_url=download_url)

    except Exception as e:
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                                      searched_article=article, message=f"Erreur inattendue : {e}",
                                      download_url=None)

@app.route("/download")
def download():
    path = request.args.get("path")
    if not path or not os.path.exists(path):
        return "Fichier introuvable ou expiré.", 404
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
