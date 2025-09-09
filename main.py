# main.py  —  09-09
# - Option “segment seul” rétablie
# - Téléchargement Excel toujours présent quand un tableau est rendu
# - Largeurs de colonnes selon tes règles (injection <colgroup>)
# - En-têtes centrées, contenu laissé à gauche
# - Remplace les NaN par "—"
# - Mise en évidence (rouge) de l’article dans les 4 colonnes d’intérêt + “Liste des chefs et articles en infraction”

import os, re, time, unicodedata
from typing import Dict, Optional, Set, List
from flask import Flask, request, render_template_string, send_file
import pandas as pd

app = Flask(__name__)

# ========== UI / CSS ==========
STYLE_BLOCK = """
<style>
  body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;margin:24px}
  h1{font-size:20px;margin-bottom:12px}
  form{display:grid;gap:10px;margin-bottom:16px}
  input[type="text"],input[type="file"]{font-size:14px}
  button{padding:8px 12px;font-size:14px;cursor:pointer}
  .note{background:#fff6e5;border:1px solid #ffd89b;padding:8px 10px;border-radius:6px;margin:10px 0 14px}
  .kbd{font-family:ui-monospace,SFMono-Regular,Menlo,monospace;background:#f3f4f6;padding:2px 4px;border-radius:4px}
  .download{margin:10px 0 8px}
  .msg{white-space:pre-wrap;font-family:ui-monospace,SFMono-Regular,Menlo,monospace;font-size:12px;margin-top:10px}
  .ok{color:#065f46}.err{color:#7f1d1d}

  table{border-collapse:collapse;width:100%;font-size:13px}
  th,td{border:1px solid #ddd;padding:6px 8px;vertical-align:top}
  th{background:#f3f4f6;text-align:center}
  td{text-align:left}

  /* fenêtre défilante horizontale quand nécessaire */
  .table-viewport{height:62vh;overflow:auto;border:1px solid #ddd}
  .table-viewport table{width:100%}
</style>
"""

HTML_TEMPLATE = """
<!doctype html>
<html lang="fr">
<head><meta charset="utf-8"><title>Analyseur Discipline – Filtrage par article</title>{{ style|safe }}</head>
<body>
  <h1>Analyseur Discipline – Filtrage par article</h1>

  <div class="note">
    Règles : détection exacte de l’article. Si la 1<sup>re</sup> cellule contient « <span class="kbd">Article filtré :</span> », 
    on ignore la 1<sup>re</sup> ligne (en-têtes sur la 2<sup>e</sup> ligne).
  </div>

  <form method="POST" enctype="multipart/form-data">
    <label>Article à rechercher (ex. <span class="kbd">29</span>, <span class="kbd">59(2)</span>)</label>
    <input type="text" name="article" required value="{{ article or '' }}" placeholder="ex. 29 ou 59(2)" />
    <label><input type="checkbox" name="segment_seul" {% if segment_seul %}checked{% endif %}> 
      Afficher uniquement le segment contenant l’article dans les 4 colonnes d’intérêt
    </label>
    <label>Fichier Excel</label>
    <input type="file" name="file" accept=".xlsx,.xlsm" required />
    <button type="submit">Analyser</button>
    <div>Formats : .xlsx / .xlsm</div>
  </form>

  {% if table_html %}
    <div class="download"><a href="{{ download_url }}">Télécharger le résultat (Excel)</a></div>
    <div class="table-viewport">{{ table_html|safe }}</div>
  {% endif %}

  {% if message %}<div class="msg {{ 'ok' if ok else 'err'}}">{{ message }}</div>{% endif %}
</body></html>
"""

# ========== Normalisation des libellés ==========
def _norm(s: str) -> str:
    if not isinstance(s, str):
        s = "" if s is None else str(s)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = s.replace("\u00A0", " ")
    return " ".join(s.strip().lower().split())

# Aliases d’en-têtes → canoniques
HEADER_ALIASES: Dict[str, Set[str]] = {
  # 4 colonnes d’intérêt
  "nbr_chefs_par_articles": {
      _norm("Nbr Chefs par articles"), _norm("Nombre de chefs par articles"),
  },
  "nbr_chefs_par_articles_par_periode": {
      _norm("Nbr Chefs par articles par période de radiation"),
      _norm("Nbr Chefs par articles par periode de radiation"),
  },
  "nb_chefs_article_reprimande": {
      _norm("Nombre de chefs par article ayant une réprimande"),
      _norm("Nombre de chefs par article ayant une reprimande"),
  },
  "nb_chefs_articles_et_total_amendes": {
      _norm("Nombre de chefs par articles et total amendes"),
  },
  # autres colonnes qu’on manipule pour la mise en forme
  "liste_chefs_articles_infraction": {
      _norm("Liste des chefs et articles en infraction"),
  },
  "liste_sanctions": {_norm("Liste des sanctions imposées"), _norm("Liste des sanctions imposees")},
  "resume_faits": {_norm("Résumé des faits concis"), _norm("Resume des faits concis")},
  "autres_mesures": {_norm("Autres mesures ordonnées"), _norm("Autres mesures ordonnees")},
  "date_creation": {_norm("Date de création"), _norm("Date de creation")},
  "date_mise_jour": {_norm("Date de mise à jour"), _norm("Date de mise a jour")},
  "numero_decision": {_norm("Numéro de la décision"), _norm("Numero de la decision")},
}

INTEREST_KEYS = [
  "nbr_chefs_par_articles",
  "nbr_chefs_par_articles_par_periode",
  "nb_chefs_article_reprimande",
  "nb_chefs_articles_et_total_amendes",
]

def resolve_columns(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    norm2orig = {_norm(c): c for c in df.columns}
    out = {}
    for key, variants in HEADER_ALIASES.items():
        hit = None
        for v in variants:
            if v in norm2orig:
                hit = norm2orig[v]
                break
        out[key] = hit
    return out

# ========== Lecture Excel (règle "Article filtré :") ==========
def read_excel(file_stream) -> pd.DataFrame:
    prev = pd.read_excel(file_stream, header=None, nrows=2, engine="openpyxl")
    file_stream.seek(0)
    first = prev.iloc[0, 0] if not prev.empty else None
    if isinstance(first, str) and _norm(first).startswith(_norm("Article filtré :")):
        return pd.read_excel(file_stream, skiprows=1, header=0, engine="openpyxl")
    return pd.read_excel(file_stream, header=0, engine="openpyxl")

# ========== Recherche d’article ==========
def build_article_pattern(token: str) -> re.Pattern:
    token = (token or "").strip()
    if not token:
        raise ValueError("Article vide.")
    esc = re.escape(token)
    tail = r"(?![\d.])" if token[-1].isdigit() else r"\b"
    return re.compile(rf"(?:\b(?:art(?:icle)?\s*[: ]*)?)({esc}){tail}", re.I)

def _prep(v: str) -> str:
    if not isinstance(v, str):
        v = "" if v is None else str(v)
    return " ".join(v.replace("•"," ").replace(" "," ").replace(" "," ").replace("\r"," ").split())

def _highlight_html(text: str, pat: re.Pattern) -> str:
    def repl(m): return f'<span style="color:#c1121f;font-weight:600">{m.group(1)}</span>'
    try:
        return pat.sub(repl, text)
    except Exception:
        return text

# ========== Extraction “segment seul” ==========
def extract_segments(text: str, pat: re.Pattern) -> str:
    if not isinstance(text, str) or not text.strip():
        return ""
    parts = re.split(r"[;\n]", text)  # coupe proprement les segments usuels
    keep = [p.strip() for p in parts if pat.search(p)]
    return " • ".join(keep)

# ========== Largeurs par colonne (injection <colgroup>) ==========
def inject_colgroup(table_html: str, columns: List[str]) -> str:
    # base = 18ch ; “double” = 36ch
    base = 18
    dbl = 36
    # colonnes à élargir x2
    widen2 = {
        _norm("Résumé des faits concis"),
        _norm("Autres mesures ordonnées"),
        _norm("Date de création"),
        _norm("Date de mise à jour"),
        _norm("Numéro de la décision"),
        _norm("Nombre de chefs par articles et total amendes"),
        _norm("Nbr Chefs par articles"),
        _norm("Nbr Chefs par articles par période de radiation"),
        _norm("Nombre de chefs par article ayant une réprimande"),
    }
    widths = []
    for c in columns:
        widths.append(dbl if _norm(c) in widen2 else base)
    colgroup = "<colgroup>" + "".join([f'<col style="width:{w}ch">' for w in widths]) + "</colgroup>"
    # insère juste après l’ouverture de <table>
    return table_html.replace("<table", "<table" + ">" + colgroup, 1).replace(">>", ">")

# ========== Export Excel ==========
def to_excel_download(df: pd.DataFrame) -> str:
    ts = int(time.time())
    out = f"/tmp/filtrage_{ts}.xlsx"
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Filtre", index=False)
        ws = writer.book.active
        for i, col in enumerate(df.columns, start=1):
            maxlen = max((len(str(x)) for x in [col] + df[col].astype(str).tolist()), default=10)
            ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = min(60, max(12, maxlen + 2))
    return f"/download?path={out}"

# ========== Route principale ==========
@app.route("/", methods=["GET","POST"])
def home():
    if request.method == "GET":
        return render_template_string(HTML_TEMPLATE, style=STYLE_BLOCK, table_html=None,
                                      download_url=None, message=None, ok=True,
                                      article="", segment_seul=False)

    file = request.files.get("file")
    article = (request.form.get("article") or "").strip()
    segment_seul = bool(request.form.get("segment_seul"))

    if not file or not article:
        return render_template_string(HTML_TEMPLATE, style=STYLE_BLOCK, table_html=None,
                                      download_url=None, article=article, segment_seul=segment_seul,
                                      message="Fichier et article requis.", ok=False)

    fname = (file.filename or "").lower()
    if not (fname.endswith(".xlsx") or fname.endswith(".xlsm")):
        return render_template_string(HTML_TEMPLATE, style=STYLE_BLOCK, table_html=None,
                                      download_url=None, article=article, segment_seul=segment_seul,
                                      message="Formats supportés : .xlsx / .xlsm", ok=False)

    try:
        df = read_excel(file.stream)
        colmap = resolve_columns(df)
        pat = build_article_pattern(article)

        # Filtrage des lignes : article présent dans au moins une des 4 colonnes d’intérêt
        masks = []
        for key in INTEREST_KEYS:
            col = colmap.get(key)
            if col and col in df.columns:
                masks.append(df[col].astype(str).apply(lambda v: bool(pat.search(_prep(v)))))
        if not masks:
            return render_template_string(HTML_TEMPLATE, style=STYLE_BLOCK, table_html=None,
                                          download_url=None, article=article, segment_seul=segment_seul,
                                          message="Aucune des 4 colonnes d’intérêt n’a été trouvée dans ce fichier.",
                                          ok=False)
        mask_any = masks[0]
        for m in masks[1:]: mask_any = mask_any | m
        df = df[mask_any].copy()
        if df.empty:
            return render_template_string(HTML_TEMPLATE, style=STYLE_BLOCK, table_html=None,
                                          download_url=None, article=article, segment_seul=segment_seul,
                                          message=f"Aucune ligne ne contient l’article « {article} » dans les colonnes cibles.",
                                          ok=True)

        # Nettoyage “segment seul” (uniquement dans les 4 colonnes d’intérêt)
        if segment_seul:
            for key in INTEREST_KEYS:
                col = colmap.get(key)
                if col and col in df.columns:
                    df[col] = df[col].apply(lambda v: extract_segments(v, pat))

        # Remplace NaN par —
        df = df.fillna("—")

        # Mise en évidence (rouge) de l’article dans 4 colonnes d’intérêt + “Liste des chefs et articles en infraction”
        highlight_cols = [colmap.get(k) for k in INTEREST_KEYS] + [colmap.get("liste_chefs_articles_infraction")]
        highlight_cols = [c for c in highlight_cols if c and c in df.columns]
        for c in highlight_cols:
            df[c] = df[c].astype(str).apply(lambda s: _highlight_html(s, pat))

        # Génération HTML avec <colgroup> pour les largeurs
        preview = df.head(200)  # garde un aperçu raisonnable
        html = preview.to_html(index=False, escape=False)  # on garde le HTML de mise en évidence
        html = inject_colgroup(html, list(preview.columns))

        download_url = to_excel_download(df)
        return render_template_string(HTML_TEMPLATE, style=STYLE_BLOCK,
                                      table_html=html, download_url=download_url,
                                      article=article, segment_seul=segment_seul,
                                      message=f"{len(df)} ligne(s) — aperçu affiché limité à 200.", ok=True)

    except Exception as e:
        return render_template_string(HTML_TEMPLATE, style=STYLE_BLOCK, table_html=None,
                                      download_url=None, article=article, segment_seul=segment_seul,
                                      message=f"Erreur : {e!r}", ok=False)

# Téléchargement
@app.route("/download")
def download():
    path = request.args.get("path")
    if not path or not os.path.exists(path):
        return "Fichier introuvable.", 404
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
