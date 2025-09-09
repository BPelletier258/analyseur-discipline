# app.py — Analyseur Discipline (09-09, anti-nan, surlignage fiable, largeurs + Excel figé)

import os, re, unicodedata, time
from typing import Dict, Optional, Set, List
from datetime import datetime

import pandas as pd
from flask import Flask, request, render_template_string, send_file

# ──────────────────────────────────────────────────────────────────────────────
# UI / CSS
# ──────────────────────────────────────────────────────────────────────────────

STYLE = """
<style>
  :root{
    --w-def: 22rem;     /* largeur par défaut */
    --w-2x:  44rem;     /* largeur ×2 demandée */
    --w-num: 8rem;      /* colonnes numériques étroites */
  }
  body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;margin:22px}
  h1{font-size:20px;margin:0 0 10px}
  .note{background:#fff8e6;border:1px solid #ffd48a;padding:8px 10px;border-radius:6px;margin:8px 0 14px}
  form{display:grid;gap:10px;margin:0 0 12px}
  input[type="text"],input[type="file"]{font-size:14px}
  .hint{font-size:12px;color:#666}
  .download{margin:10px 0}
  table{border-collapse:collapse;width:100%;font-size:13px;table-layout:fixed}
  th,td{border:1px solid #ddd;padding:6px 8px;vertical-align:top}
  th{background:#f3f4f6;text-align:center}
  .col-def{min-width:var(--w-def)}
  .col-2x{min-width:var(--w-2x)}
  .col-num{min-width:var(--w-num);width:var(--w-num);text-align:center}
  .bullets{margin:0;padding-left:1.1rem}
  .bullets li{margin:0.15rem 0}
  .dash{color:#666}
  .hl{color:#d00;font-weight:600}
  .viewport{height:60vh;overflow:auto;border:1px solid #ddd}
  .msg{white-space:pre-wrap;font-family:ui-monospace,Menlo,monospace;font-size:12px;margin-top:10px}
  .ok{color:#065f46}.err{color:#7f1d1d}
</style>
"""

TPL = """
<!doctype html><html><head><meta charset="utf-8"><title>Analyseur Discipline</title>{{css}}</head>
<body>
  <h1>Analyseur Discipline – Filtrage par article</h1>
  <div class="note">
    Règles : détection exacte ; si la 1<sup>re</sup> cellule contient « <code>Article filtré :</code> », on ignore la 1<sup>re</sup> ligne (entêtes en 2<sup>e</sup> ligne).
  </div>
  <form method="POST" enctype="multipart/form-data">
    <label>Article à rechercher (ex. <code>29</code>, <code>59(2)</code>)</label>
    <input name="article" type="text" required value="{{article or ''}}" />
    <label><input type="checkbox" name="only_segment" {{'checked' if only_segment else ''}}>
      Afficher uniquement le segment contenant l’article dans les 4 colonnes d’intérêt
    </label>
    <label>Fichier Excel</label>
    <input name="file" type="file" accept=".xlsx,.xlsm" required />
    <button type="submit">Analyser</button>
    <div class="hint">Formats : .xlsx / .xlsm</div>
  </form>

  {% if table_html %}
  <div class="download"><a href="{{download_url}}">Télécharger le résultat (Excel)</a></div>
  <div class="viewport">{{table_html|safe}}</div>
  {% endif %}

  {% if message %}
  <div class="msg {{'ok' if ok else 'err'}}">{{message}}</div>
  {% endif %}
</body></html>
"""

app = Flask(__name__)

# ──────────────────────────────────────────────────────────────────────────────
# Normalisation des en-têtes & alias
# ──────────────────────────────────────────────────────────────────────────────

def _norm(s: str) -> str:
    if not isinstance(s, str): s = "" if s is None else str(s)
    s = unicodedata.normalize("NFKD", s).encode("ascii","ignore").decode("ascii")
    s = s.replace("\u00A0"," ")
    return " ".join(s.strip().lower().split())

ALIASES: Dict[str, Set[str]] = {
  "liste_chefs_articles_infraction": {_norm("Liste des chefs et articles en infraction")},
  "resume_faits": {_norm("Résumé des faits concis"), _norm("Resume des faits concis")},
  "liste_sanctions": {_norm("Liste des sanctions imposées"), _norm("Liste des sanctions imposees")},

  # 4 colonnes d’intérêt
  "nbr_chefs_articles": {_norm("Nbr Chefs par articles"), _norm("Nombre de chefs par articles")},
  "nbr_chefs_par_periode": {_norm("Nbr Chefs par articles par période de radiation"),
                            _norm("Nbr Chefs par articles par periode de radiation")},
  "nb_chefs_total_amendes": {_norm("Nombre de chefs par articles et total amendes")},
  "nb_chefs_reprimande": {_norm("Nombre de chefs par article ayant une réprimande"),
                          _norm("Nombre de chefs par article ayant une reprimande")},

  # Autres colonnes larges à doubler
  "autres_mesures": {_norm("Autres mesures ordonnées"), _norm("Autres mesures ordonnees")},
  "a_verifier": {_norm("À vérifier"), _norm("A verifier")},
  "num_decision": {_norm("Numéro de la décision"), _norm("Numero de la decision")},
  "date_creation": {_norm("Date de création"), _norm("Date de creation")},
  "date_maj": {_norm("Date de mise à jour"), _norm("Date de mise a jour")},

  # num petites
  "total_chefs": {_norm("Total chefs")},
  "total_amendes": {_norm("Total amendes")},
  "total_reprimandes": {_norm("Total réprimandes"), _norm("Total reprimandes")},
  "radiation_max": {_norm("Radiation max")},
}

BULLET_COL_KEYS = {
  "resume_faits","liste_chefs_articles_infraction","nbr_chefs_articles","liste_sanctions",
  "nbr_chefs_par_periode","nb_chefs_total_amendes","nb_chefs_reprimande","autres_mesures","a_verifier"
}

INTEREST_KEYS = {"nbr_chefs_articles","nbr_chefs_par_periode","nb_chefs_total_amendes","nb_chefs_reprimande"}

DOUBLE_KEYS = {"resume_faits","autres_mesures","date_creation","date_maj","num_decision"} | INTEREST_KEYS
NUM_KEYS = {"total_chefs","total_amendes","total_reprimandes","radiation_max"}

def resolve(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    m = {_norm(c): c for c in df.columns}
    out: Dict[str, Optional[str]] = {k: None for k in ALIASES}
    for key, names in ALIASES.items():
        for n in names:
            if n in m: out[key] = m[n]; break
    return out

# ──────────────────────────────────────────────────────────────────────────────
# Lecture Excel (gère « Article filtré : » en A1)
# ──────────────────────────────────────────────────────────────────────────────

def read_excel(file) -> pd.DataFrame:
    prev = pd.read_excel(file, header=None, nrows=1, engine="openpyxl")
    file.seek(0)
    skip = 1 if (len(prev.columns) and isinstance(prev.iat[0,0], str) and _norm(prev.iat[0,0]).startswith(_norm("Article filtré :"))) else 0
    return pd.read_excel(file, engine="openpyxl", header=0, skiprows=skip)

# ──────────────────────────────────────────────────────────────────────────────
# Motif article + nettoyage / surlignage
# ──────────────────────────────────────────────────────────────────────────────

def build_pattern(token: str) -> re.Pattern:
    token = token.strip()
    esc = re.escape(token)
    tail = r"(?![\d.])" if token and token[-1].isdigit() else r"\b"
    return re.compile(rf"(?:\b(?:art(?:icle)?\s*[: ]*)?)({esc}){tail}", re.I)

def prep_text(v) -> str:
    if v is None: return ""
    s = str(v)
    s = s.replace("•"," | ").replace("·"," | ").replace("◦"," | ").replace("\u00A0"," ")
    s = s.replace("\r\n","\n").replace("\r","\n").replace("\n"," | ")
    s = " ".join(s.split())
    return s

def split_items(s: str) -> List[str]:
    if not s: return []
    parts = [p.strip(" •-–—") for p in re.split(r"\s*\|\s*|;\s*", s) if p.strip()]
    return parts

def highlight_html(text: str, pat: re.Pattern) -> str:
    if not text: return ""
    out, last = [], 0
    for m in pat.finditer(text):
        a, b = m.span(1)
        out.append(text[last:a])
        out.append(f'<span class="hl">{m.group(1)}</span>')
        last = b
    out.append(text[last:])
    return "".join(out)

def as_list_html(s: str, pat: re.Pattern) -> str:
    items = split_items(s)
    if not items: return '<span class="dash">—</span>'
    lis = "".join(f"<li>{highlight_html(it, pat)}</li>" for it in items)
    return f'<ul class="bullets">{lis}</ul>'

# ──────────────────────────────────────────────────────────────────────────────
# Extraction des seuls segments (option)
# ──────────────────────────────────────────────────────────────────────────────

def keep_only_segments(s: str, pat: re.Pattern) -> str:
    segs = [p for p in split_items(s) if pat.search(p)]
    return " | ".join(segs)

# ──────────────────────────────────────────────────────────────────────────────
# Construction du HTML avec classes de largeur
# ──────────────────────────────────────────────────────────────────────────────

def build_table_html(df: pd.DataFrame, colmap: Dict[str, Optional[str]], pat: re.Pattern) -> str:
    # remplace NaN pour éviter "nan"
    df = df.copy()
    df = df.where(pd.notnull(df), "")

    # Option puces + surlignage
    name_for = {v:k for k,v in colmap.items() if v}
    for col in df.columns:
        key = name_for.get(col)
        if key in BULLET_COL_KEYS:
            df[col] = df[col].map(lambda x: as_list_html(prep_text(x), pat))
        else:
            df[col] = df[col].map(lambda x: highlight_html(prep_text(x), pat))

    # Génère le tableau (sans échapper le HTML)
    html = df.to_html(index=False, escape=False)
    # Ajoute classes de largeur via colgroup
    headers = list(df.columns)
    classes = []
    for h in headers:
        key = name_for.get(h)
        if key in NUM_KEYS: classes.append('class="col-num"')
        elif key in DOUBLE_KEYS: classes.append('class="col-2x"')
        else: classes.append('class="col-def"')
    colgroup = "<colgroup>" + "".join(f"<col {c}/>" for c in classes) + "</colgroup>"
    html = re.sub(r"(<table[^>]*>)", r"\1"+colgroup, html, count=1)
    return html

# ──────────────────────────────────────────────────────────────────────────────
# Excel : 1ʳᵉ ligne “Article filtré : …”, entêtes ligne 2, freeze A3
# ──────────────────────────────────────────────────────────────────────────────

def to_excel(df: pd.DataFrame, article: str, colmap: Dict[str, Optional[str]]) -> str:
    # Nettoyage pour Excel : contenus à puces sur lignes séparées
    df = df.copy().where(pd.notnull(df), "")
    name_for = {v:k for k,v in colmap.items() if v}

    for col in df.columns:
        key = name_for.get(col)
        if key in BULLET_COL_KEYS:
            df[col] = df[col].map(lambda s: "\n".join(split_items(prep_text(s))))
        else:
            df[col] = df[col].map(lambda s: prep_text(s))

    ts = int(time.time())
    path = f"/tmp/filtrage_{ts}.xlsx"
    with pd.ExcelWriter(path, engine="openpyxl") as wr:
        # startrow=1 pour laisser la ligne 1 libre (bannière)
        df.to_excel(wr, index=False, sheet_name="Filtre", startrow=1)
        ws = wr.book.active
        # Ligne 1 : bannière
        ws["A1"] = f"Article filtré : {article}"
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=df.shape[1])
        ws["A1"].font = ws["A2"].font.copy(bold=True)
        # Geler l'entête (ligne 2) -> curseur sur A3
        ws.freeze_panes = "A3"
        # Largeurs auto (plafonnées)
        from openpyxl.utils import get_column_letter
        for i, c in enumerate(df.columns, start=1):
            values = [str(c)] + [str(v) if v is not None else "" for v in df[c]]
            max_len = min(60, max(len(x) for x in values) + 2)
            ws.column_dimensions[get_column_letter(i)].width = max(12, max_len)
    return f"/download?path={path}"

# ──────────────────────────────────────────────────────────────────────────────
# Filtrage principal
# ──────────────────────────────────────────────────────────────────────────────

@app.route("/", methods=["GET","POST"])
def main():
    if request.method == "GET":
        return render_template_string(TPL, css=STYLE, table_html=None, article="", only_segment=False,
                                      message=None, ok=True)

    f = request.files.get("file")
    article = (request.form.get("article") or "").strip()
    only_segment = bool(request.form.get("only_segment"))

    if not f or not article:
        return render_template_string(TPL, css=STYLE, table_html=None, article=article, only_segment=only_segment,
                                      message="Fichier et article requis.", ok=False)

    name = (f.filename or "").lower()
    if not (name.endswith(".xlsx") or name.endswith(".xlsm")):
        return render_template_string(TPL, css=STYLE, table_html=None, article=article, only_segment=only_segment,
                                      message="Format non supporté. Fournir un .xlsx ou .xlsm", ok=False)

    try:
        df = read_excel(f.stream)
        colmap = resolve(df)
        # colonnes nécessaires au filtrage : les 4 d’intérêt
        needed = [colmap[k] for k in INTEREST_KEYS if colmap[k]]
        if not needed:
            return render_template_string(
                TPL, css=STYLE, table_html=None, article=article, only_segment=only_segment,
                message="Aucune des 4 colonnes d’intérêt n’a été trouvée dans le fichier.", ok=False
            )
        pat = build_pattern(article)

        # Masque : au moins une des 4 colonnes contient l’article
        mask = None
        for c in needed:
            cur = df[c].astype(str).map(lambda s: bool(pat.search(prep_text(s))))
            mask = cur if mask is None else (mask | cur)
        df = df[mask].copy()
        if df.empty:
            return render_template_string(TPL, css=STYLE, table_html=None, article=article, only_segment=only_segment,
                                          message=f"Aucune ligne ne contient l’article « {article} ».", ok=True)

        # Option "seul segment" dans les 4 colonnes d’intérêt
        if only_segment:
            for key in INTEREST_KEYS:
                c = colmap.get(key)
                if c and c in df.columns:
                    df[c] = df[c].astype(str).map(lambda s: keep_only_segments(prep_text(s), pat))

        # HTML
        table_html = build_table_html(df.head(500), colmap, pat)  # aperçu (sécurité)
        # Excel
        dl_url = to_excel(df, article, colmap)

        return render_template_string(TPL, css=STYLE, table_html=table_html, article=article,
                                      only_segment=only_segment, download_url=dl_url,
                                      message=f"{len(df)} ligne(s) retenue(s).", ok=True)
    except Exception as e:
        return render_template_string(TPL, css=STYLE, table_html=None, article=article, only_segment=only_segment,
                                      message=f"Erreur : {e!r}", ok=False)

@app.route("/download")
def download():
    path = request.args.get("path")
    if not path or not os.path.exists(path): return "Fichier introuvable.", 404
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
