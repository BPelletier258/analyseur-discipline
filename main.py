# main.py — 100% complet
import os, re, time, unicodedata
from typing import Dict, Optional, Set, List
import pandas as pd
from flask import Flask, request, render_template_string, send_file

app = Flask(__name__)

# ──────────────────────────────────────────────────────────────────────────────
# Mise en forme (CSS + gabarit HTML)
# ──────────────────────────────────────────────────────────────────────────────
STYLE = """
<style>
  :root { --hit:#d31b1b; }
  body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;margin:24px}
  h1{font-size:20px;margin:0 0 12px}
  form{display:grid;gap:10px;margin:0 0 14px}
  input[type=text]{padding:8px;font-size:14px}
  input[type=file]{font-size:14px}
  label.inline{display:flex;align-items:center;gap:.5rem;font-size:13px}
  .note{background:#fff7e6;border:1px solid #ffd48a;padding:8px 10px;border-radius:6px;margin:8px 0 14px}
  .kbd{font-family:ui-monospace,Menlo,monospace;background:#f3f4f6;padding:2px 4px;border-radius:4px}
  .download{margin:10px 0}
  .table-viewport{height:62vh;overflow:auto;border:1px solid #ddd}
  .table-wide{min-width:200vw}
  table{border-collapse:collapse;width:100%;font-size:13px}
  th,td{border:1px solid #e5e7eb;padding:6px 8px;vertical-align:top}
  th{background:#f3f4f6;text-align:center}
  ul.cell{margin:0;padding-left:18px}
  ul.cell li{margin:0 0 2px}
  .hit{color:var(--hit);font-weight:700}
  .dash{color:#111}
  .msg{white-space:pre-wrap;font-family:ui-monospace,Menlo,monospace;font-size:12px;margin-top:10px}
  .ok{color:#065f46}.err{color:#7f1d1d}
</style>
"""

HTML = """
<!doctype html>
<meta charset="utf-8">
<title>Analyseur Discipline – Filtrage par article</title>
{{ style|safe }}
<h1>Analyseur Discipline – Filtrage par article</h1>

<div class="note">
Règles : détection exacte de l’article. Si la 1<sup>re</sup> cellule contient
« <span class="kbd">Article filtré :</span> », la 2<sup>e</sup> ligne est utilisée comme en-têtes.
</div>

<form method="POST" enctype="multipart/form-data">
  <label>Article à rechercher (ex. <span class="kbd">29</span>, <span class="kbd">59(2)</span>)</label>
  <input type="text" name="article" required value="{{ article or '' }}" placeholder="ex. 29 ou 59(2)">

  <label>Fichier Excel</label>
  <input type="file" name="file" accept=".xlsx,.xlsm" required>

  <label class="inline">
    <input type="checkbox" name="segments_only" value="1" {% if segments_only %}checked{% endif %}>
    Afficher uniquement le segment contenant l’article dans les 4 colonnes d’intérêt
  </label>

  <button type="submit">Analyser</button>
</form>

{% if table_html %}
  <div class="download"><a href="{{ download_url }}">Télécharger le résultat (Excel)</a></div>
  <div class="table-viewport"><div class="table-wide">{{ table_html|safe }}</div></div>
{% endif %}

{% if message %}
  <div class="msg {{ 'ok' if ok else 'err' }}">{{ message }}</div>
{% endif %}
"""

# ──────────────────────────────────────────────────────────────────────────────
# Normalisation d’en-têtes & alias
# ──────────────────────────────────────────────────────────────────────────────
def _norm(s:str)->str:
    if not isinstance(s,str):
        s = "" if s is None else str(s)
    s = unicodedata.normalize("NFKD", s).encode("ascii","ignore").decode("ascii")
    s = s.replace("\u00A0"," ")
    return " ".join(s.strip().lower().split())

# Colonnes — canons et alias (ajoute/retire librement ici)
ALIASES: Dict[str, Set[str]] = {
  # 4 colonnes d’intérêt (filtrage)
  "nbr_chefs_articles": {
    _norm("Nbr Chefs par articles"), _norm("Articles enfreints"),
    _norm("Articles en infraction")
  },
  "nbr_chefs_par_radiation": {
    _norm("Nbr Chefs par articles par période de radiation"),
    _norm("Nbr Chefs par articles par periode de radiation"),
    _norm("Durée totale effective radiation"), _norm("Duree totale effective radiation")
  },
  "chefs_et_total_amendes": {
    _norm("Nombre de chefs par articles et total amendes"),
    _norm("Article amende/chef"), _norm("Articles amende / chef")
  },
  "nbr_chefs_avec_reprimande": {
    _norm("Nombre de chefs par article ayant une réprimande"),
    _norm("Nombre de chefs par article ayant une reprimande"),
    _norm("Autres sanctions")
  },

  # Autres colonnes qu’on met en puces
  "resume_faits": { _norm("Résumé des faits concis"), _norm("Resume des faits concis") },
  "liste_chefs_articles": { _norm("Liste des chefs et articles en infraction") },
  "liste_sanctions": { _norm("Liste des sanctions imposées"), _norm("Liste des sanctions imposees") },
  "autres_mesures": { _norm("Autres mesures ordonnées"), _norm("Autres mesures ordonnees") },
  "a_verifier": { _norm("À vérifier"), _norm("A verifier") },

  # Quelques infos (non obligatoires)
  "date_creation": { _norm("Date de création"), _norm("Date de creation") },
  "date_maj": { _norm("Date de mise à jour"), _norm("Date de mise a jour") },
}

INTEREST_CANONS = [
  "nbr_chefs_articles",
  "nbr_chefs_par_radiation",
  "chefs_et_total_amendes",
  "nbr_chefs_avec_reprimande",
]

# Colonnes à afficher en listes à puces :
BULLET_CANONS = set(INTEREST_CANONS + [
  "resume_faits","liste_chefs_articles","liste_sanctions","autres_mesures","a_verifier"
])

# Colonnes qui doivent mettre l’article en rouge :
HIGHLIGHT_CANONS = set(INTEREST_CANONS + ["liste_chefs_articles"])

def resolve_columns(df:pd.DataFrame)->Dict[str, Optional[str]]:
    m = {_norm(c): c for c in df.columns}
    out = {}
    for canon, names in ALIASES.items():
        hit = None
        for n in names:
            if n in m:
                hit = m[n]; break
        out[canon] = hit
    return out

# ──────────────────────────────────────────────────────────────────────────────
# Lecture Excel (tolère bannière « Article filtré : » en A1)
# ──────────────────────────────────────────────────────────────────────────────
def read_excel(stream)->pd.DataFrame:
    head = pd.read_excel(stream, header=None, nrows=2, engine="openpyxl")
    stream.seek(0)
    first = head.iloc[0,0] if not head.empty else ""
    if isinstance(first,str) and _norm(first).startswith(_norm("Article filtré :")):
        return pd.read_excel(stream, skiprows=1, header=0, engine="openpyxl")
    return pd.read_excel(stream, header=0, engine="openpyxl")

# ──────────────────────────────────────────────────────────────────────────────
# Article pattern (exact, pas « 29 » dans « 299 »)
# ──────────────────────────────────────────────────────────────────────────────
def article_regex(token:str)->re.Pattern:
    token = (token or "").strip()
    if not token: raise ValueError("Article vide.")
    esc = re.escape(token)
    tail = r"(?![\d.])" if token[-1].isdigit() else r"\b"
    # Groupe 1 = l’article capturé (pour la mise en rouge)
    return re.compile(rf"(?:\b(?:art(?:icle)?\s*[: ]*)?)({esc}){tail}", re.I)

# ──────────────────────────────────────────────────────────────────────────────
# Utilities de mise en forme (liste, tiret, mise en rouge)
# ──────────────────────────────────────────────────────────────────────────────
SEPS = re.compile(r"(?:\r\n|\n|\\n|;|•|·|◦)")

def _txt(x)->str:
    if x is None: return ""
    s = str(x)
    return s.replace("\u00A0"," ").strip()

def split_items(value:str)->List[str]:
    s = _txt(value)
    if not s: return []
    # normaliser les "\n" littéraux en vraies nouvelles lignes
    s = s.replace("\\n","\n")
    parts = [p.strip(" •\t\r\n") for p in SEPS.split(s)]
    return [p for p in parts if p]

def hi(text:str, pat:re.Pattern)->str:
    # met en rouge uniquement le groupe 1 du motif
    return pat.sub(lambda m: f'<span class="hit">{m.group(1)}</span>', text)

def as_bullets(value, pat=None, highlight=False)->str:
    items = split_items(value)
    if not items:
        return '<span class="dash">—</span>'
    if highlight and pat:
        items = [hi(i,pat) for i in items]
    lis = "".join(f"<li>{i}</li>" for i in items)
    return f'<ul class="cell">{lis}</ul>'

def as_text(value)->str:
    s = _txt(value)
    return s if s else '<span class="dash">—</span>'

# Extraction du seul segment contenant l’article (pour l’option « segments only »)
def keep_only_segments(value, pat)->str:
    parts = split_items(value)
    hits = [p for p in parts if pat.search(p)]
    return "\n".join(hits)

# ──────────────────────────────────────────────────────────────────────────────
# Export Excel
# ──────────────────────────────────────────────────────────────────────────────
def excel_download(df:pd.DataFrame)->str:
    path = f"/tmp/filtrage_{int(time.time())}.xlsx"
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Filtre")
        ws = w.book.active
        # largeur auto (simple)
        for j, col in enumerate(df.columns, start=1):
            vals = [len(str(col))] + [len(str(x)) for x in df[col].fillna("").astype(str)]
            ws.column_dimensions[ws.cell(1,j).column_letter].width = min(60, max(12, max(vals)+2))
    return f"/download?path={path}"

# ──────────────────────────────────────────────────────────────────────────────
# Route principale
# ──────────────────────────────────────────────────────────────────────────────
@app.route("/", methods=["GET","POST"])
def main():
    if request.method=="GET":
        return render_template_string(HTML, style=STYLE, table_html=None,
                                      article="", segments_only=False, message=None, ok=True)

    article = (request.form.get("article") or "").strip()
    segments_only = request.form.get("segments_only") == "1"
    f = request.files.get("file")

    if not article or not f:
        return render_template_string(HTML, style=STYLE, table_html=None,
                                      article=article, segments_only=segments_only,
                                      message="Article et fichier sont requis.", ok=False)

    name = (f.filename or "").lower()
    if not (name.endswith(".xlsx") or name.endswith(".xlsm")):
        ext = name.split(".")[-1] if "." in name else "?"
        return render_template_string(HTML, style=STYLE, table_html=None,
                                      article=article, segments_only=segments_only,
                                      message=f"Format non pris en charge: .{ext}. Utilise .xlsx ou .xlsm.",
                                      ok=False)

    try:
        df = read_excel(f.stream)
        colmap = resolve_columns(df)
        pat = article_regex(article)

        # 1) Filtrer les LIGNES si l’article apparaît dans au moins une des 4 colonnes d’intérêt
        masks = []
        for canon in INTEREST_CANONS:
            col = colmap.get(canon)
            if col and col in df.columns:
                masks.append(df[col].astype(str).apply(lambda v: bool(pat.search(_txt(v)))))
        if not masks:
            return render_template_string(
                HTML, style=STYLE, table_html=None, article=article, segments_only=segments_only,
                message=("Aucune des colonnes d’intérêt n’a été trouvée.\n"
                         f"Résolution colonnes:\n{colmap}\n\nColonnes Excel:\n{list(df.columns)}"),
                ok=False
            )
        mask = masks[0]
        for m in masks[1:]: mask = mask | m
        df = df[mask].copy()
        if df.empty:
            return render_template_string(HTML, style=STYLE, table_html=None,
                                          article=article, segments_only=segments_only,
                                          message=f"Aucune ligne ne contient l’article « {article} ».", ok=True)

        # 2) Préparer l’affichage : listes à puces + mise en rouge + option « segments only »
        #    On ne touche qu’aux colonnes connues; les autres restent texte simple avec tiret si vide.
        for canon, col in colmap.items():
            if not col or col not in df.columns: 
                continue
            is_bullets = canon in BULLET_CANONS
            needs_highlight = canon in HIGHLIGHT_CANONS

            series = df[col].copy()

            if segments_only and canon in INTEREST_CANONS:
                # Remplacer le contenu par le (ou les) segment(s) contenant l’article
                series = series.apply(lambda v: keep_only_segments(v, pat))

            if is_bullets:
                df[col] = series.apply(lambda v: as_bullets(v, pat, needs_highlight))
            else:
                # simple texte (tiret si vide)
                df[col] = series.apply(as_text)

        # Valeurs restantes NaN -> tiret
        for c in df.columns:
            df[c] = df[c].apply(lambda v: as_text(v) if (not isinstance(v,str) or not v.startswith("<")) else v)

        # 3) HTML + export
        download_url = excel_download(df)
        html = df.head(200).to_html(index=False, escape=False)
        return render_template_string(
            HTML, style=STYLE, table_html=html, download_url=download_url,
            article=article, segments_only=segments_only,
            message=f"{len(df)} ligne(s) retenue(s). Aperçu limité à 200 lignes.", ok=True
        )

    except Exception as e:
        return render_template_string(HTML, style=STYLE, table_html=None,
                                      article=article, segments_only=segments_only,
                                      message=f"Erreur : {repr(e)}", ok=False)

@app.route("/download")
def download():
    path = request.args.get("path")
    if not path or not os.path.exists(path): return "Fichier introuvable.", 404
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
