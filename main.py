# main.py — 09-08 (checkbox permanente, surlignage, puces, Excel propre)

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

# ──────────────────────────  STYLE & HTML  ──────────────────────────

STYLE_BLOCK = """
<style>
  :root{
    --w-def: 22rem;         /* largeur par défaut */
    --w-2x: 44rem;          /* largeur x2 demandée */
    --w-num: 8rem;          /* colonnes numériques étroites */
  }
  body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;margin:22px}
  h1{font-size:20px;margin:0 0 10px}
  .note{background:#fff8e6;border:1px solid #ffd48a;padding:8px 10px;border-radius:6px;margin:8px 0 14px}
  form{display:grid;gap:10px;margin:8px 0 12px}
  input[type="text"],input[type="file"]{font-size:14px}
  .hint{font-size:12px;color:#666}
  .download{margin:10px 0}
  .msg{white-space:pre-wrap;font-family:ui-monospace,Menlo,monospace;font-size:12px}
  .ok{color:#065f46}.err{color:#7f1d1d}
  /* Tableau */
  table{border-collapse:collapse;width:100%;font-size:13px;table-layout:fixed}
  th,td{border:1px solid #ddd;padding:6px 8px;vertical-align:top}
  th{background:#f3f4f6;text-align:center}
  .hl{color:#d00;font-weight:600}
  .dash{color:#666}
  ul.bullets{margin:0;padding-left:1.1rem}
  ul.bullets li{margin:0.15rem 0}
  /* viewport scrollable horizontal + vertical */
  .viewport{height:60vh;overflow:auto;border:1px solid #ddd}
  /* colonnes à largeur fixe via colgroup */
  col.c-def{width:var(--w-def)}
  col.c-2x{width:var(--w-2x)}
  col.c-num{width:var(--w-num)}
</style>
"""

HTML_TEMPLATE = """
<!doctype html><html><head>
<meta charset="utf-8" />
<title>Analyseur Discipline – Filtrage par article</title>
{{ style_block|safe }}
</head><body>
  <h1>Analyseur Discipline – Filtrage par article</h1>

  <div class="note">
    Règles : détection exacte de l’article. Si la 1<sup>re</sup> cellule contient
    « <b>Article filtré :</b> », on ignore la 1<sup>re</sup> ligne (lignes d’en-tête sur la 2<sup>e</sup>).
  </div>

  <form method="POST" enctype="multipart/form-data">
    <div>
      <label>Article à rechercher (ex. <code>29</code>, <code>59(2)</code>)</label><br>
      <input type="text" name="article" required value="{{ searched_article or '' }}" style="width:18rem">
    </div>

    <label>
      <input type="checkbox" name="segment_only" value="1" {% if segment_only %}checked{% endif %}>
      Afficher uniquement le segment contenant l’article dans les 4 colonnes d’intérêt
    </label>

    <div>
      <label>Fichier Excel</label><br>
      <input type="file" name="file" accept=".xlsx,.xlsm" required>
    </div>

    <button type="submit">Analyser</button>
    <div class="hint">Formats : .xlsx / .xlsm</div>
  </form>

  {% if table_html %}
    <div class="download">
      <a href="{{ download_url }}">Télécharger le résultat (Excel)</a>
    </div>
    <div class="viewport">{{ table_html|safe }}</div>
  {% endif %}

  {% if message %}
    <div class="msg {{ 'ok' if message_ok else 'err' }}">{{ message }}</div>
  {% endif %}
</body></html>
"""

# ──────────────────────────  UTILITAIRES  ──────────────────────────

def _norm(s: str) -> str:
    if not isinstance(s, str):
        s = "" if s is None else str(s)
    s = unicodedata.normalize("NFKD", s).encode("ascii","ignore").decode("ascii")
    s = s.replace("\u00A0"," ").replace("\u202F"," ")
    return " ".join(s.strip().lower().split())

# Aliases d’en-têtes (vous pouvez en ajouter si besoin)
HEADER_ALIASES: Dict[str, Set[str]] = {
    # Colonnes d’intérêt (filtrage)
    "articles_enfreints": {_norm("Nbr Chefs par articles"),
                           _norm("Articles en infraction"),
                           _norm("Liste des chefs et articles en infraction")},
    "duree_totale_radiation": {_norm("Nbr Chefs par articles par période de radiation"),
                               _norm("Nbr Chefs par articles par periode de radiation")},
    "article_amende_chef": {_norm("Nombre de chefs par articles et total amendes"),
                            _norm("Article amende/chef"),
                            _norm("Total amendes par article")},
    "autres_sanctions": {_norm("Nombre de chefs par article ayant une réprimande"),
                         _norm("Nombre de chefs par article ayant une reprimande")},
    # Colonnes utiles pour formatage/Excel
    "total_amendes": {_norm("Total amendes"), _norm("Montant total des amendes")},
    "date_creation": {_norm("Date de création"), _norm("date de creation")},
    "date_maj": {_norm("Date de mise à jour"), _norm("date de mise a jour")},
}

FILTER_CANONICAL = [
    "articles_enfreints",
    "duree_totale_radiation",
    "article_amende_chef",
    "autres_sanctions",
]

def resolve_columns(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    mp = {_norm(c): c for c in df.columns}
    out: Dict[str, Optional[str]] = {}
    for canon, variants in HEADER_ALIASES.items():
        hit = None
        for v in variants:
            if v in mp:
                hit = mp[v]; break
        out[canon] = hit
    return out

def read_excel_respecting_header_rule(stream) -> pd.DataFrame:
    head = pd.read_excel(stream, header=None, nrows=2, engine="openpyxl")
    stream.seek(0)
    first = head.iloc[0,0] if not head.empty else None
    banner = isinstance(first,str) and _norm(first).startswith(_norm("Article filtré :"))
    if banner:
        return pd.read_excel(stream, skiprows=1, header=0, engine="openpyxl")
    return pd.read_excel(stream, header=0, engine="openpyxl")

def build_article_pattern(token: str) -> re.Pattern:
    t = (token or "").strip()
    if not t: raise ValueError("Article vide.")
    esc = re.escape(t)
    tail = r"(?![\d.])" if t[-1].isdigit() else r"\b"
    return re.compile(rf"(?:\b(?:art(?:icle)?\s*[: ]*)?)({esc}){tail}", re.I)

def _prep_text(v) -> str:
    if not isinstance(v,str): v = "" if v is None else str(v)
    return " ".join(
        v.replace("•"," ").replace("·"," ").replace("◦"," ")
         .replace("\u00A0"," ").replace("\u202F"," ")
         .replace("\r\n","\n").replace("\r","\n").split()
    )

# Extraction « segment uniquement » (case cochée)
def _segments(text: str, pat: re.Pattern, sep_regex=r"[;,\n]") -> List[str]:
    if not isinstance(text,str) or not text.strip(): return []
    parts = [p.strip() for p in re.split(sep_regex, text)]
    return [p for p in parts if pat.search(p)]

# Mise en forme HTML (puces + surlignage)
def _hl(html: str, pat: re.Pattern) -> str:
    return pat.sub(r'<span class="hl">\1</span>', html)

def _as_bullets(st: str) -> str:
    """Conserve de vraies puces <ul><li>…</li></ul> ; si vide → tiret « — » sans puce."""
    if not isinstance(st,str) or not st.strip():
        return '<span class="dash">—</span>'
    # on découpe sur saut de ligne ou « • »
    items = []
    tmp = st.replace("•", "\n")
    for line in [x.strip() for x in tmp.split("\n")]:
        if line:
            items.append(f"<li>{line}</li>")
    if not items:
        return '<span class="dash">—</span>'
    return f'<ul class="bullets">{"".join(items)}</ul>'

# Table HTML à largeur stable (colgroup)
def df_to_html_bulleted(df: pd.DataFrame, pat: re.Pattern, highlight_cols: Set[str]) -> str:
    # mapping grossier de largeur : numériques, x2, default
    num_like = {"Total chefs", "Total amendes"}
    x2_like = {
        "Résumé des faits concis",
        "Autres mesures ordonnées",
        "Date de création",
        "Date de mise à jour",
        "Numéro de la décision",
        "Nbr Chefs par articles",
        "Nbr Chefs par articles par période de radiation",
        "Nombre de chefs par articles et total amendes",
        "Nombre de chefs par article ayant une réprimande",
    }

    cols = list(df.columns)
    col_classes = []
    for c in cols:
        if c in num_like: col_classes.append("c-num")
        elif c in x2_like: col_classes.append("c-2x")
        else: col_classes.append("c-def")

    # en-tête + colgroup
    parts = ['<table><colgroup>']
    for cls in col_classes:
        parts.append(f'<col class="{cls}">')
    parts.append('</colgroup><thead><tr>')
    for c in cols:
        parts.append(f"<th>{c}</th>")
    parts.append('</tr></thead><tbody>')

    # corps
    for _, row in df.iterrows():
        parts.append("<tr>")
        for c in cols:
            val = row[c]
            # normalise NaN
            if pd.isna(val) or (isinstance(val,float) and pd.isna(val)):
                html = '<span class="dash">—</span>'
            else:
                s = str(val)
                # surlignage dans colonnes textuelles
                if c in highlight_cols:
                    s = _hl(s, pat)
                html = _as_bullets(s)
            parts.append(f"<td>{html}</td>")
        parts.append("</tr>")
    parts.append("</tbody></table>")
    return "".join(parts)

# Nettoyage des 4 colonnes d’intérêt (mode « segment seulement »)
def clean_segments_only(df: pd.DataFrame, colmap: Dict[str, Optional[str]], pat: re.Pattern) -> pd.DataFrame:
    df = df.copy()
    for canon in FILTER_CANONICAL:
        col = colmap.get(canon)
        if not col or col not in df.columns: continue
        df[col] = df[col].apply(lambda v: " \n ".join(_segments(_prep_text(v), pat)))
    # on filtre : garder au moins une mention
    subset = [colmap[k] for k in FILTER_CANONICAL if colmap.get(k)]
    if subset:
        m = False
        for c in subset:
            cur = df[c].astype(str).str.strip().ne("")
            m = cur if m is False else (m | cur)
        df = df[m]
    return df

# Formattage amendes : "5 000 $"
def format_money_series(ser: pd.Series) -> pd.Series:
    def _fmt(x):
        if pd.isna(x) or x == "": return "—"
        try:
            # tolère strings "5000", "5 000", "5000 $"
            s = str(x).replace(" ", "").replace("$","").replace("\u00A0","")
            n = float(s)
            return f"{int(n):,}".replace(",", " ") + " $"
        except Exception:
            return str(x)
    return ser.apply(_fmt)

# Excel export soigné
def to_excel_download(df: pd.DataFrame, article: str, colmap: Dict[str, Optional[str]]) -> str:
    ts = int(time.time())
    path = f"/tmp/filtrage_{ts}.xlsx"
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        # Ligne 1 : bannière "Article filtré : X"
        df.to_excel(writer, index=False, sheet_name="Filtre", startrow=1)
        ws = writer.book["Filtre"]
        ws.cell(1,1, f"Article filtré : {article}")
        # merge sur toute la ligne d'en-tête
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(1, df.shape[1]))
        # gel des volets : entête + bannière
        ws.freeze_panes = "A3"
        # largeurs approx
        from openpyxl.utils import get_column_letter
        for i, col in enumerate(df.columns, start=1):
            max_len = max([len(str(col))] + [len(str(x)) for x in df[col].astype(str).tolist() if x is not None])
            ws.column_dimensions[get_column_letter(i)].width = min(60, max(12, max_len + 2))
        # format monétaire si colonne présente
        ta = colmap.get("total_amendes")
        if ta and ta in df.columns:
            col_idx = list(df.columns).index(ta) + 1
            for r in range(3, df.shape[0]+3):
                ws.cell(r, col_idx).number_format = '# ##0 [$-x-currency]$'  # affichage local
    return f"/download?path={path}"

# ──────────────────────────  ROUTES  ──────────────────────────

@app.route("/", methods=["GET","POST"])
def analyze():
    if request.method == "GET":
        return render_template_string(
            HTML_TEMPLATE, style_block=STYLE_BLOCK,
            table_html=None, message=None, message_ok=True,
            searched_article="", segment_only=False
        )

    file = request.files.get("file")
    article = (request.form.get("article") or "").strip()
    segment_only = request.form.get("segment_only") == "1"

    if not file or not article:
        return render_template_string(
            HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
            message="Fichier et article requis.", message_ok=False,
            searched_article=article, segment_only=segment_only
        )

    # .xlsx / .xlsm uniquement
    low = (file.filename or "").lower()
    if not (low.endswith(".xlsx") or low.endswith(".xlsm")):
        return render_template_string(
            HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
            message="Format non pris en charge. Fournissez un .xlsx ou .xlsm.",
            message_ok=False, searched_article=article, segment_only=segment_only
        )

    try:
        df = read_excel_respecting_header_rule(file.stream)
        colmap = resolve_columns(df)
        pat = build_article_pattern(article)

        # masque de filtrage sur les 4 colonnes d’intérêt
        masks = []
        any_cols = False
        for canon in FILTER_CANONICAL:
            col = colmap.get(canon)
            if col and col in df.columns:
                any_cols = True
                masks.append(df[col].astype(str).apply(lambda v: bool(pat.search(_prep_text(v)))))
        if not any_cols:
            detail = "\n".join([f"- {k}: {colmap.get(k)}" for k in FILTER_CANONICAL])
            return render_template_string(
                HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                message=("Aucune des colonnes cibles n’a été trouvée.\n"
                         "Vérifiez les en-têtes ou ajustez les alias dans le code.\n\n"
                         "Résolution:\n"+detail+"\n\nColonnes disponibles:\n"+", ".join(df.columns)),
                message_ok=False, searched_article=article, segment_only=segment_only
            )

        mask_any = masks[0]
        for m in masks[1:]:
            mask_any = mask_any | m

        df_filtered = df[mask_any].copy()
        if df_filtered.empty:
            return render_template_string(
                HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                message=f"Aucune ligne ne contient l’article « {article} » dans les colonnes cibles.",
                message_ok=True, searched_article=article, segment_only=segment_only
            )

        # Option : segment uniquement
        if segment_only:
            df_use = clean_segments_only(df_filtered, colmap, pat)
        else:
            df_use = df_filtered.copy()

        # Remplacements NaN → —
        df_use = df_use.where(pd.notnull(df_use), "—")

        # Total amendes : format « 5 000 $ »
        ta = colmap.get("total_amendes")
        if ta and ta in df_use.columns:
            df_use[ta] = format_money_series(df_use[ta])

        # Colonnes à surligner (texte intégral)
        highlight_cols = set()
        if not segment_only:
            for canon in FILTER_CANONICAL:
                col = colmap.get(canon)
                if col and col in df_use.columns:
                    highlight_cols.add(col)

        # HTML bullets + surlignage
        html = df_to_html_bulleted(df_use, pat, highlight_cols)

        # Excel export
        download_url = to_excel_download(df_use, article, colmap)

        return render_template_string(
            HTML_TEMPLATE, style_block=STYLE_BLOCK,
            table_html=html, download_url=download_url,
            message=f"{len(df_use)} ligne(s) correspondante(s).",
            message_ok=True, searched_article=article, segment_only=segment_only
        )

    except Exception as e:
        return render_template_string(
            HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
            message=f"Erreur inattendue : {repr(e)}",
            message_ok=False, searched_article=article, segment_only=segment_only
        )

@app.route("/download")
def download():
    path = request.args.get("path")
    if not path or not os.path.exists(path):
        return "Fichier introuvable ou expiré.", 404
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
