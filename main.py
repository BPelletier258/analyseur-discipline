# -*- coding: utf-8 -*-
# Analyseur Discipline – Filtrage par article
# Version stabilisée — styles intégrés + options + Excel propre (A1 "Article filtré : …" + en-tête figée)

import os
import re
import io
import time
import math
import unicodedata
from typing import Dict, Optional, List, Set

import pandas as pd
from flask import Flask, request, render_template_string, send_file

app = Flask(__name__)

# ───────────────────────────────────────── CSS & HTML (intégrés, pas injectés dynamiquement)
HTML = r"""
<!doctype html>
<html lang="fr">
<head>
<meta charset="utf-8">
<title>Analyseur Discipline – Filtrage par article</title>
<style>
  :root{
    --w-def: 22rem;      /* largeur par défaut */
    --w-2x:  44rem;      /* largeur x2 pour colonnes demandées */
    --w-num: 8rem;       /* colonnes numériques */
  }
  body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;margin:22px}
  h1{font-size:20px;margin:0 0 10px}
  form{display:grid;gap:10px;margin:8px 0 14px}
  input[type="text"],input[type="file"]{font-size:14px}
  input[type="text"]{padding:8px}
  .hint{font-size:12px;color:#666}
  .note{background:#fff6e5;border:1px solid #ffd48a;padding:8px 10px;border-radius:6px;margin:10px 0}
  .download{margin:12px 0}
  .kbd{font:12px ui-monospace,SFMono-Regular,Menlo,monospace;background:#f3f4f6;border:1px solid #e5e7eb;border-radius:4px;padding:1px 4px}
  /* conteneur avec scroll horizontal */
  .viewport{height:60vh;overflow:auto;border:1px solid #ddd}
  table{border-collapse:collapse;table-layout:fixed;min-width:var(--w-2x);width:100%;font-size:13px}
  th,td{border:1px solid #ddd;padding:6px 8px;vertical-align:top}
  th{background:#f3f4f6;text-align:center}
  .center{ text-align:center }
  /* Largeurs par classes */
  .w-def{min-width:var(--w-def)}
  .w-2x{min-width:var(--w-2x)}
  .w-num{min-width:var(--w-num)}
  /* listes à puces propres */
  ul{margin:0;padding:0 0 0 1.15rem}
  li{margin:0.15rem 0}
  .dash{color:#666}
  .hit{color:#d00;font-weight:600}   /* mise en rouge de l’article repéré */
</style>
</head>
<body>
  <h1>Analyseur Discipline – Filtrage par article</h1>

  <div class="note">
    Règles : détection exacte de l’article. Si la 1<sup>re</sup> cellule contient
    <span class="kbd">Article filtré :</span>, on ignore la 1<sup>re</sup> ligne (lignes d’en-tête sur la 2<sup>e</sup>).
  </div>

  <form method="POST" enctype="multipart/form-data">
    <label>Article à rechercher (ex. <span class="kbd">29</span>, <span class="kbd">59(2)</span>)</label>
    <input name="article" required placeholder="ex. 29 ou 59(2)" value="{{ article|e }}">
    <div>
      <label>
        <input type="checkbox" name="isolate" value="1" {% if isolate %}checked{% endif %}>
        Afficher uniquement le segment contenant l’article dans les 4 colonnes d’intérêt
      </label>
    </div>
    <label>Fichier Excel</label>
    <input type="file" name="file" accept=".xlsx,.xlsm" required>
    <button>Analyser</button>
    <div class="hint">Formats : .xlsx / .xlsm</div>
  </form>

  {% if download_url %}
    <div class="download"><a href="{{ download_url }}">Télécharger le résultat (Excel)</a></div>
  {% endif %}

  {% if table %}
    <div class="viewport">{{ table|safe }}</div>
  {% endif %}

  {% if msg %}
    <pre class="hint">{{ msg }}</pre>
  {% endif %}
</body>
</html>
"""

# ───────────────────────────────────────── utilitaires texte/entêtes

def _norm(s: str) -> str:
    if not isinstance(s, str):
        s = "" if s is None else str(s)
    s = unicodedata.normalize("NFKD", s).encode("ascii","ignore").decode("ascii")
    s = s.replace("\u00A0"," ")
    return " ".join(s.strip().lower().split())

# alias d’en-têtes (beaucoup de tolérance)
ALIASES: Dict[str, Set[str]] = {
    # colonnes d’intérêt (filtrage + surlignage en rouge)
    "nbr_chefs_par_articles": {
        _norm("Nbr Chefs par articles"),
        _norm("Nombre de chefs par articles"),
        _norm("Articles enfreints"),
        _norm("Articles en infraction"),
    },
    "nbr_chefs_par_articles_par_periode": {
        _norm("Nbr Chefs par articles par période de radiation"),
        _norm("Nbr Chefs par articles par periode de radiation"),
        _norm("Durée totale effective radiation"),
        _norm("Duree totale effective radiation"),
    },
    "nbre_chefs_par_articles_total_amendes": {
        _norm("Nombre de chefs par articles et total amendes"),
        _norm("Article amende/chef"),
        _norm("Articles amende / chef"),
        _norm("Amendes (article/chef)"),
    },
    "nbre_chefs_par_article_reprimande": {
        _norm("Nombre de chefs par article ayant une réprimande"),
        _norm("Nombre de chefs par article ayant une reprimande"),
        _norm("Autres sanctions"),
    },

    # surlignage en plus ici :
    "liste_chefs_articles_infraction": {
        _norm("Liste des chefs et articles en infraction"),
    },

    # autres colonnes (pour largeur/cosmétique)
    "resume_faits": {_norm("Résumé des faits concis"), _norm("Resume des faits concis")},
    "liste_sanctions": {_norm("Liste des sanctions imposées"), _norm("Liste des sanctions imposees")},
    "autres_mesures": {_norm("Autres mesures ordonnées"), _norm("Autres mesures ordonnees")},
    "a_verifier": {_norm("À vérifier"), _norm("A verifier")},
    "date_creation": {_norm("Date de création"), _norm("Date de creation")},
    "date_maj": {_norm("Date de mise à jour"), _norm("Date de mise a jour")},
    "numero_decision": {_norm("Numéro de la décision"), _norm("Numero de la decision")},
    "total_amendes": {_norm("Total amendes")},
    "total_reprimandes": {_norm("Total réprimandes"), _norm("Total reprimandes")},
    "total_chefs": {_norm("Total chefs")},
}

INTEREST_KEYS = [
    "nbr_chefs_par_articles",
    "nbr_chefs_par_articles_par_periode",
    "nbre_chefs_par_articles_total_amendes",
    "nbre_chefs_par_article_reprimande",
]

def resolve_columns(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    m = {_norm(c): c for c in df.columns}
    out: Dict[str, Optional[str]] = {}
    for key, variants in ALIASES.items():
        out[key] = next((m[v] for v in variants if v in m), None)
    return out

# ───────────────────────────────────────── lecture Excel (gère bannière "Article filtré : ...")

def read_excel(stream) -> pd.DataFrame:
    head = pd.read_excel(stream, header=None, nrows=1, engine="openpyxl")
    stream.seek(0)
    if not head.empty and isinstance(head.iloc[0,0], str) and _norm(head.iloc[0,0]).startswith(_norm("Article filtré :")):
        df = pd.read_excel(stream, skiprows=1, engine="openpyxl")
    else:
        df = pd.read_excel(stream, engine="openpyxl")
    return df

# ───────────────────────────────────────── recherche & surlignage

def build_article_regex(token: str) -> re.Pattern:
    token = token.strip()
    if not token:
        raise ValueError("Article vide.")
    esc = re.escape(token)
    tail = r"(?![\d.])" if token[-1].isdigit() else r"\b"
    # groupe 1 = le numéro exact, pour entourer de <span class="hit">
    return re.compile(rf"(?:\b(?:art(?:icle)?\s*[: ]*)?)({esc}){tail}", re.IGNORECASE)

def prep_text(v) -> str:
    if not isinstance(v, str):
        v = "" if (v is None or (isinstance(v,float) and math.isnan(v))) else str(v)
    # normalisation douce des séparateurs
    v = v.replace("•", "\n").replace("·", "\n").replace("◦", "\n").replace("  ", " ")
    v = v.replace("\u00A0", " ").replace("\u202F", " ")
    return v.strip()

def split_segments(text: str) -> List[str]:
    if not text:
        return []
    # éclate sur retours ligne, point-virgule ou " • " résiduels
    parts = re.split(r"[;\n]+", text)
    return [p.strip(" •\t\r ") for p in parts if p.strip(" •\t\r ")]

def keep_only_hits(text: str, rx: re.Pattern) -> List[str]:
    return [seg for seg in split_segments(text) if rx.search(seg)]

def highlight(text: str, rx: re.Pattern) -> str:
    # remplace groupe(1) par <span class="hit">…</span> sans toucher le reste
    return rx.sub(r'<span class="hit">\1</span>', text)

# ───────────────────────────────────────── rendu HTML (puces + classes largeur)

# mapping « clé alias résolue » => classe largeur
WIDTH_CLASS = {
    # laisses tel quel : "liste_chefs_articles_infraction", "liste_sanctions" (w-def)
    "resume_faits": "w-2x",
    "autres_mesures": "w-2x",
    "date_creation": "w-2x",
    "date_maj": "w-2x",
    "numero_decision": "w-2x",
    "nbre_chefs_par_articles_total_amendes": "w-2x",
    "nbr_chefs_par_articles": "w-2x",
    "nbr_chefs_par_articles_par_periode": "w-2x",
    "nbre_chefs_par_article_reprimande": "w-2x",
    "total_chefs": "w-num",
    "total_amendes": "w-num",
    "total_reprimandes": "w-num",
}

def key_for_column(label: str) -> Optional[str]:
    n = _norm(label)
    for key, variants in ALIASES.items():
        if any(n == v for v in variants):
            return key
    return None

def cell_to_html(label: str, raw: str, rx: re.Pattern, isolate: bool, red_in_cols: Set[str]) -> str:
    key = key_for_column(label) or ""
    txt = prep_text(raw)

    # isole uniquement dans les 4 colonnes d’intérêt (si coché)
    if isolate and key in set(INTEREST_KEYS):
        segs = keep_only_hits(txt, rx)
    else:
        segs = split_segments(txt)

    # si aucune donnée → tiret sans puce
    if not segs:
        return '<span class="dash">—</span>'

    # mise en évidence de l’article pour colonnes ciblées (4 colonnes d’intérêt + liste chefs/infraction)
    if key in red_in_cols:
        segs = [highlight(s, rx) for s in segs]

    # rendu à puces
    items = "".join(f"<li>{s}</li>" for s in segs)
    return f"<ul>{items}</ul>"

def render_table(df: pd.DataFrame, rx: re.Pattern, isolate: bool) -> str:
    # quelles colonnes reçoivent la coloration rouge ?
    red_cols = set(INTEREST_KEYS + ["liste_chefs_articles_infraction"])

    # classes de largeur par nom *résolu* ; fallback w-def
    ths: List[str] = []
    classes: List[str] = []
    for col in df.columns:
        key = key_for_column(col) or ""
        cls = WIDTH_CLASS.get(key, "w-def")
        ths.append(f'<th class="{cls}">{col}</th>')
        classes.append(cls)

    # lignes
    rows_html: List[str] = []
    for _, row in df.iterrows():
        tds = []
        for (col, cls) in zip(df.columns, classes):
            tds.append(f'<td class="{cls}">{cell_to_html(col, row[col], rx, isolate, red_cols)}</td>')
        rows_html.append("<tr>" + "".join(tds) + "</tr>")

    table = "<table><thead><tr>" + "".join(ths) + "</tr></thead><tbody>" + "".join(rows_html) + "</tbody></table>"
    return table

# ───────────────────────────────────────── export Excel

def export_excel(df: pd.DataFrame, article: str) -> str:
    # remplace NaN par vide pour éviter "nan"
    df_x = df.copy().fillna("")
    # pour Excel on garde des sauts de ligne simples
    for c in df_x.columns:
        df_x[c] = df_x[c].map(lambda v: "\n".join(split_segments(prep_text(v))) if isinstance(v, str) else "")

    ts = int(time.time())
    path = f"/tmp/filtrage_{ts}.xlsx"

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        # on démarre à la ligne 2 (startrow=1) pour mettre A1 = "Article filtré : …"
        df_x.to_excel(writer, index=False, sheet_name="Filtre", startrow=1)
        ws = writer.book["Filtre"]
        ws["A1"] = f"Article filtré : {article}"
        # fige la ligne d’en-tête (A2 => entêtes visibles au scroll)
        ws.freeze_panes = "A2"
        # largeur colonnes
        from openpyxl.utils import get_column_letter
        for i, col in enumerate(df_x.columns, start=1):
            # largeur heuristique (max 60)
            values = [str(col)] + [str(v) for v in df_x[col].tolist()]
            width = min(60, max(12, max(len(s.split("\n")[0]) for s in values) + 2))
            ws.column_dimensions[get_column_letter(i)].width = width

    return f"/download?path={path}"

# ───────────────────────────────────────── route

@app.route("/", methods=["GET", "POST"])
def index():
    article = (request.form.get("article") or "").strip()
    isolate = request.form.get("isolate") == "1"

    if request.method == "GET":
        return render_template_string(HTML, article="", isolate=False, table=None, download_url=None, msg=None)

    file = request.files.get("file")
    if not file or not article:
        return render_template_string(HTML, article=article, isolate=isolate, table=None,
                                      download_url=None, msg="Fichier et article requis.")

    # formats acceptés
    fname = (file.filename or "").lower()
    if not (fname.endswith(".xlsx") or fname.endswith(".xlsm")):
        return render_template_string(HTML, article=article, isolate=isolate, table=None, download_url=None,
                                      msg="Veuillez fournir un .xlsx ou .xlsm (les .xls ne sont pas supportés).")

    try:
        df = read_excel(file.stream)
    except Exception as e:
        return render_template_string(HTML, article=article, isolate=isolate, table=None, download_url=None,
                                      msg=f"Erreur lecture Excel : {e!r}")

    if df.empty:
        return render_template_string(HTML, article=article, isolate=isolate, table=None, download_url=None,
                                      msg="Classeur vide.")

    colmap = resolve_columns(df)

    # vérifie qu'au moins une des 4 colonnes d’intérêt existe
    target_cols = [colmap[k] for k in INTEREST_KEYS if colmap.get(k) in df.columns]
    if not target_cols:
        return render_template_string(HTML, article=article, isolate=isolate, table=None, download_url=None,
            msg="Aucune des 4 colonnes d’intérêt n’a été trouvée. Vérifie les en-têtes (onglet « Cumul décisions »).")

    rx = build_article_regex(article)

    # masque « la ligne contient l’article dans au moins une des 4 colonnes d’intérêt »
    mask = False
    for c in target_cols:
        cur = df[c].astype(str).map(prep_text).map(lambda s: bool(rx.search(s)))
        mask = cur if mask is False else (mask | cur)

    df = df[mask].copy()
    if df.empty:
        return render_template_string(HTML, article=article, isolate=isolate, table=None, download_url=None,
                                      msg=f"Aucune ligne ne contient l’article « {article} » dans les colonnes d’intérêt.")

    # Rendu HTML (puces + surlignage + largeurs)
    html_table = render_table(df, rx, isolate)

    # Export Excel prêt
    dl = export_excel(df, article)

    return render_template_string(HTML, article=article, isolate=isolate, table=html_table, download_url=dl, msg=None)

@app.route("/download")
def download():
    path = request.args.get("path")
    if not path or not os.path.exists(path):
        return "Fichier introuvable.", 404
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))
