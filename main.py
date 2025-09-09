# -*- coding: utf-8 -*-
# ======================================================================
# Analyseur Discipline – Filtrage par article
# - Puces uniquement sur colonnes "liste"
# - Pas de puce sur colonnes mono-valeur (noms, dates, totaux…)
# - Surlignage rouge de l’article dans les colonnes d’intérêt
# - NaN -> "—" (sans puce)
# - Formatage de "Total amendes" : 5 000 $, 25 000 $, etc.
# ======================================================================

import os
import re
import time
import unicodedata
from typing import Dict, Optional, Set

import pandas as pd
from flask import Flask, request, render_template_string, send_file

app = Flask(__name__)

# ----------------------------- Styles ---------------------------------

STYLE_BLOCK = """
<style>
  :root{
    --w-def:22rem;    /* largeur par défaut */
    --w-2x: 44rem;    /* colonnes qu'on élargit x2 */
    --w-num: 8rem;    /* colonnes numériques étroites */
  }
  body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;margin:22px}
  h1{font-size:20px;margin:0 0 10px}
  .note{background:#fff8e6;border:1px solid #ffd48a;padding:8px 10px;border-radius:6px;margin:10px 0}
  form{display:grid;gap:10px;margin:0 0 12px}
  input[type="text"],input[type="file"]{font-size:14px}
  .hint{font-size:12px;color:#666}
  .download{margin:10px 0}
  table{border-collapse:collapse;width:100%;font-size:13px;table-layout:fixed}
  th,td{border:1px solid #ddd;padding:6px 8px;vertical-align:top}
  th{background:#f3f4f6;text-align:center}
  /* Colonnes : largeur de base + ajustements ciblés */
  th,td{min-width:var(--w-def)}
  .col-num, .col-tot, .col-amd{min-width:var(--w-num)}
  .col-wide{min-width:var(--w-2x)}
  /* UL/LI compacts pour les listes */
  ul{margin:0;padding-left:1.1rem}
  li{margin:0.15rem 0}
  /* Surlignage de l’article */
  .hl{color:#d00;font-weight:700}
  /* Zone scrollable horizontale si besoin */
  .viewport{overflow:auto;border:1px solid #ddd}
  .msg{white-space:pre-wrap;font-family:ui-monospace,Menlo,Consolas,monospace;font-size:12px;margin-top:8px}
  .ok{color:#065f46}.err{color:#7f1d1d}
</style>
"""

# ----------------------------- HTML -----------------------------------

HTML = """
<!doctype html><html><head><meta charset="utf-8">
<title>Analyseur Discipline — Filtrage par article</title>
{{ css|safe }}</head><body>
<h1>Analyseur Discipline – Filtrage par article</h1>

<div class="note">
  Règles : détection exacte de l’article. Si la 1<sup>re</sup> cellule contient
  « <code>Article filtré :</code> », on ignore la 1<sup>re</sup> ligne (lignes d’en-tête sur la 2<sup>e</sup>).
</div>

<form method="POST" enctype="multipart/form-data">
  <label>Article à rechercher (ex. <code>29</code>, <code>59(2)</code>)</label>
  <input type="text" name="article" value="{{ article or '' }}" required>
  <label>Fichier Excel</label>
  <input type="file" name="file" accept=".xlsx,.xlsm" required>
  <button>Analyser</button>
  <div class="hint">Formats : .xlsx / .xlsm</div>
</form>

{% if table_html %}
  <div class="download"><a href="{{ dl }}">Télécharger le résultat (Excel)</a></div>
  <div class="viewport">{{ table_html|safe }}</div>
{% endif %}

{% if msg %}<div class="msg {{ 'ok' if ok else 'err' }}">{{ msg }}</div>{% endif %}
</body></html>
"""

# ------------------------ Normalisation en-têtes ----------------------

def _norm(s: str) -> str:
    """ascii, trim, lower, espaces compressés."""
    if not isinstance(s, str):
        s = "" if s is None else str(s)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = s.replace("\u00A0", " ")
    return " ".join(s.strip().lower().split())

HEADER_ALIASES: Dict[str, Set[str]] = {
  # 4 colonnes cibles de filtrage (on garde vos alias existants)
  "articles_enfreints": {_norm("Nbr Chefs par articles"), _norm("Articles enfreints"),
                         _norm("Articles en infraction"), _norm("Liste des chefs et articles en infraction")},
  "duree_totale_radiation": {_norm("Nbr Chefs par articles par période de radiation"),
                             _norm("Nbr Chefs par articles par periode de radiation"),
                             _norm("Durée totale effective radiation"), _norm("Duree totale effective radiation")},
  "article_amende_chef": {_norm("Nombre de chefs par articles et total amendes"),
                          _norm("Article amende/chef"), _norm("Articles amende / chef")},
  "autres_sanctions": {_norm("Nombre de chefs par article ayant une réprimande"),
                       _norm("Nombre de chefs par article ayant une reprimande"),
                       _norm("Autres sanctions"), _norm("Autres mesures ordonnées")},
  # utilitaires (pas de filtrage, mais utiles au rendu)
  "resume_faits": {_norm("Résumé des faits concis"), _norm("Resume des faits concis")},
  "liste_chefs": {_norm("Liste des chefs et articles en infraction")},
  "liste_sanctions": {_norm("Liste des sanctions imposées"), _norm("Liste des sanctions imposees")},
  "a_verifier": {_norm("À vérifier"), _norm("A verifier")},
  "autres_mesures": {_norm("Autres mesures ordonnées"), _norm("Autres mesures ordonnees")},
  "total_amendes": {_norm("Total amendes")},
  "total_chefs": {_norm("Total chefs")},
  "date_creation": {_norm("Date de création"), _norm("Date de creation")},
  "date_maj": {_norm("Date de mise à jour"), _norm("Date de mise a jour")},
  "nom": {_norm("Nom de l'intimé"), _norm("Nom de l’intimé")},
  "numero": {_norm("Numéro de la décision"), _norm("Numero de la decision")},
}

FILTER_CANON = ["articles_enfreints", "duree_totale_radiation",
                "article_amende_chef", "autres_sanctions"]

def resolve_columns(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    m = {_norm(c): c for c in df.columns}
    out: Dict[str, Optional[str]] = {}
    for canon, aliases in HEADER_ALIASES.items():
        hit = None
        for a in aliases:
            if a in m:
                hit = m[a]; break
        out[canon] = hit
    return out

# ------------------------------ Excel ---------------------------------

def read_excel(file_stream) -> pd.DataFrame:
    prev = pd.read_excel(file_stream, header=None, nrows=2, engine="openpyxl")
    file_stream.seek(0)
    first = prev.iloc[0, 0] if not prev.empty else ""
    banner = isinstance(first, str) and _norm(first).startswith(_norm("Article filtré :"))
    if banner:
        return pd.read_excel(file_stream, skiprows=1, header=0, engine="openpyxl")
    return pd.read_excel(file_stream, header=0, engine="openpyxl")

# ------------------------- Recherche article --------------------------

def build_article_pattern(token: str) -> re.Pattern:
    t = (token or "").strip()
    if not t:
        raise ValueError("Article vide.")
    esc = re.escape(t)
    tail = r"(?![\d.])" if t[-1].isdigit() else r"\b"
    return re.compile(rf"(?:\b(?:art(?:icle)?\s*[: ]*)?)({esc}){tail}", re.I)

def _prep(v) -> str:
    s = "" if v is None else str(v)
    s = s.replace("•", " ").replace("·", " ").replace("◦", " ")
    s = s.replace("\u00A0", " ").replace("\u202F", " ")
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    return " ".join(s.split())

# ---------------------------- Nettoyage --------------------------------

def _extract_hits_generic(text: str, pat: re.Pattern) -> str:
    if not text.strip(): return ""
    parts = re.split(r"[;,\n]", text)
    hits = [p.strip() for p in parts if pat.search(p)]
    return " | ".join(hits)

def _extract_hits_autres(text: str, pat: re.Pattern) -> str:
    if not text.strip(): return ""
    parts = [p.strip() for p in re.split(r"[;\n]", text) if pat.search(p)]
    return " | ".join(parts)

def clean_filtered(df: pd.DataFrame, colmap: Dict[str, Optional[str]], pat: re.Pattern) -> pd.DataFrame:
    out = df.copy()
    for canon in FILTER_CANON:
        col = colmap.get(canon)
        if not col or col not in out.columns: continue
        if canon == "autres_sanctions":
            out[col] = out[col].apply(lambda v: _extract_hits_autres(_prep(v), pat))
        else:
            out[col] = out[col].apply(lambda v: _extract_hits_generic(_prep(v), pat))
    # on ne garde que les lignes où au moins une des 4 colonnes contient qqch
    cols = [colmap[c] for c in FILTER_CANON if colmap.get(c)]
    if cols:
        mask = False
        for c in cols:
            m = out[c].astype(str).str.strip().ne("")
            mask = m if mask is False else (mask | m)
        out = out[mask]
    return out

# ------------------------ Rendu : listes ciblées -----------------------

# Colonnes RENDUES EN LISTE (puces). Tout le reste reste en texte simple.
LIST_COLUMNS = {
  _norm("Résumé des faits concis"),
  _norm("Liste des chefs et articles en infraction"),
  _norm("Liste des sanctions imposées"),
  _norm("Nbr Chefs par articles"),
  _norm("Nbr Chefs par articles par période de radiation"),
  _norm("Nombre de chefs par article ayant une réprimande"),
  _norm("Autres mesures ordonnées"),
  _norm("À vérifier"),
}

def _fmt_money(val) -> str:
    """ '5000' -> '5 000 $' ; 0 -> '0 $' ; vide -> '—' """
    if val is None or (isinstance(val, float) and pd.isna(val)) or (isinstance(val, str) and not val.strip()):
        return "—"
    try:
        n = float(str(val).replace(" ", "").replace(",", "."))
        s = f"{int(round(n)):d}"
        s = re.sub(r"(?<=\d)(?=(\d{3})+$)", " ", s)  # espacements milliers
        return f"{s} $"
    except Exception:
        # si déjà un texte (ex: "0 $", "—") on renvoie tel quel
        return str(val)

def _highlight(text: str, pat: re.Pattern) -> str:
    return pat.sub(r'<span class="hl">\\1</span>', text)

def htmlize_df(df: pd.DataFrame, pat: re.Pattern) -> str:
    """Transforme le DF en HTML en appliquant :
       - puces uniquement sur colonnes LIST_COLUMNS
       - surlignage rouge de l’article
       - NaN -> '—'
       - formatage 'Total amendes'
    """
    df2 = df.copy()

    # mapping normalisé -> vrai nom de colonne pour décisions ciblées
    norm_cols = {_norm(c): c for c in df2.columns}

    # 1) Total amendes => format monnaie
    for cand in ["total amendes", "total_amendes"]:
        if cand in norm_cols:
            c = norm_cols[cand]
            df2[c] = df2[c].apply(_fmt_money)

    # 2) Remplacement NaN générique par '—' (texte simple)
    df2 = df2.fillna("—")

    # 3) Surlignage + listification ciblée
    def render_cell(col_name: str, val) -> str:
        raw = "" if val is None else str(val)
        ncol = _norm(col_name)

        # Colonnes "liste" => UL/LI (si contenu non vide et réellement segmenté)
        if ncol in LIST_COLUMNS:
            if raw.strip() in ("", "—"):
                return "—"
            # on split sur séparateurs usuels (on accepte " | " laissé par l'extraction)
            parts = [p.strip() for p in re.split(r"\|\s*|\n", raw) if p.strip()]
            if not parts:
                return "—"
            items = "".join(f"<li>{_highlight(p, pat)}</li>" for p in parts)
            return f"<ul>{items}</ul>"

        # Colonnes simples => pas de puce
        if raw.strip() == "":
            return "—"
        return _highlight(raw, pat)

    for col in df2.columns:
        df2[col] = df2[col].apply(lambda v, c=col: render_cell(c, v))

    # 4) Indices de classes pour largeur (optionnel, garde l’esthétique)
    #    -> on marque quelques colonnes numériques / étroites
    classes = []
    for c in df2.columns:
        n = _norm(c)
        if n in {"total chefs", "total amendes"}:
            classes.append("col-num")
        elif n in {"date de creation", "date de mise a jour"}:
            classes.append("col-num")
        elif n in {"resume des faits concis", "autres mesures ordonnees"}:
            classes.append("col-wide")
        else:
            classes.append("")

    html = df2.to_html(index=False, escape=False)
    # injecte classes <th> / <td> (simple et robuste)
    # pandas sort déjà un <thead><tr><th>... ; on remplace séquentiellement
    for cls in classes:
        html = html.replace("<th>", f"<th class=\"{cls}\">", 1)
    # pour les <td>, on applique par colonne (pandas met colgroup pas simple à cibler,
    # mais le rendu reste lisible même sans classe td ; on peut s’en passer)
    return html

# --------------------------- Export Excel -----------------------------

def to_excel(df: pd.DataFrame) -> str:
    ts = int(time.time())
    path = f"/tmp/filtrage_{ts}.xlsx"
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Filtre")
        ws = w.book.active
        # largeurs automatiques (simple)
        for i, col in enumerate(df.columns, start=1):
            values = [str(col)] + [str(x) for x in df[col].tolist()]
            width = min(60, max(12, max(len(v) for v in values) + 2))
            ws.column_dimensions[ws.cell(1, i).column_letter].width = width
        # fige la ligne d'en-tête
        ws.freeze_panes = "A2"
    return f"/download?path={path}"

# ------------------------------- Route --------------------------------

@app.route("/", methods=["GET", "POST"])
def main():
    if request.method == "GET":
        return render_template_string(HTML, css=STYLE_BLOCK, table_html=None,
                                      article=None, dl=None, msg=None, ok=True)

    file = request.files.get("file")
    article = (request.form.get("article") or "").strip()

    if not file or not article:
        return render_template_string(HTML, css=STYLE_BLOCK, table_html=None, article=article,
                                      dl=None, msg="Fichier et article requis.", ok=False)

    # .xlsx /.xlsm uniquement
    fn = (file.filename or "").lower()
    if not (fn.endswith(".xlsx") or fn.endswith(".xlsm")):
        return render_template_string(HTML, css=STYLE_BLOCK, table_html=None, article=article,
                                      dl=None, msg="Fournis un .xlsx ou .xlsm (pas .xls).", ok=False)

    try:
        df = read_excel(file.stream)
        colmap = resolve_columns(df)
        pat = build_article_pattern(article)

        # masque “lignes qui contiennent l’article dans au moins 1 colonne cible”
        masks = []
        any_col = False
        for canon in FILTER_CANON:
            col = colmap.get(canon)
            if col and col in df.columns:
                any_col = True
                masks.append(df[col].astype(str).apply(lambda v: bool(pat.search(_prep(v)))))
        if not any_col:
            return render_template_string(HTML, css=STYLE_BLOCK, table_html=None, article=article,
                                          dl=None, msg="Aucune des colonnes cibles n’a été trouvée.", ok=False)
        mask = masks[0]
        for m in masks[1:]:
            mask |= m
        df = df[mask]
        if df.empty:
            return render_template_string(HTML, css=STYLE_BLOCK, table_html=None, article=article,
                                          dl=None, msg=f"Aucune ligne avec l’article « {article} ».", ok=True)

        df_clean = clean_filtered(df, colmap, pat)
        if df_clean.empty:
            return render_template_string(HTML, css=STYLE_BLOCK, table_html=None, article=article,
                                          dl=None, msg="Après épuration, rien à afficher.", ok=True)

        # Rendu HTML (listes ciblées, sans puces ailleurs)
        table_html = htmlize_df(df_clean, pat)
        dl = to_excel(df_clean)

        return render_template_string(HTML, css=STYLE_BLOCK, table_html=table_html,
                                      article=article, dl=dl,
                                      msg=f"{len(df_clean)} ligne(s) après filtrage.", ok=True)

    except Exception as e:
        return render_template_string(HTML, css=STYLE_BLOCK, table_html=None, article=article,
                                      dl=None, msg=f"Erreur : {e!r}", ok=False)

@app.route("/download")
def download():
    path = request.args.get("path")
    if not path or not os.path.exists(path):
        return "Fichier introuvable.", 404
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
