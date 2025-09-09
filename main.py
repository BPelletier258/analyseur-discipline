# -*- coding: utf-8 -*-
# Analyseur Discipline — rendu HTML en listes + surlignage sûr

import io, os, re, time, unicodedata
from typing import Dict, Optional, Set, Iterable

import pandas as pd
from flask import Flask, request, render_template_string, send_file

app = Flask(__name__)

# ---------- Styles ----------
STYLE_BLOCK = """
<style>
  :root { --hit:#d0021b; }
  body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;margin:24px}
  h1{font-size:20px;margin:0 0 12px}
  .note{background:#fff6e5;border:1px solid #ffd89b;padding:8px 10px;border-radius:6px;margin:10px 0 16px}
  form{display:grid;gap:12px;margin:0 0 16px}
  input[type="text"],input[type="file"],button{font-size:14px}
  table{border-collapse:collapse;font-size:13px}
  th,td{border:1px solid #ddd;padding:6px 8px;vertical-align:top}
  th{background:#f3f4f6;text-align:center}
  .bul{list-style:disc;margin:0;padding-left:18px}
  .hit{color:var(--hit);font-weight:700}
  .viewport{height:60vh;overflow:auto;border:1px solid #ddd}
  .wide{min-width:200vw}
</style>
"""

# ---------- HTML ----------
HTML_TEMPLATE = """
<!doctype html><html><head><meta charset="utf-8"/>
<title>Analyseur Discipline – Filtrage par article</title>
{{ style_block|safe }}</head><body>
<h1>Analyseur Discipline – Filtrage par article</h1>
<div class="note">
Règles : détection exacte de l’article; si la 1<sup>re</sup> cellule contient « <b>Article filtré :</b> », elle est ignorée (en-têtes sur la 2<sup>e</sup> ligne).
</div>

<form method="POST" enctype="multipart/form-data">
  <label>Article à rechercher (ex. <b>29</b>, <b>59(2)</b>)</label>
  <input type="text" name="article" value="{{ searched_article or '' }}" required>
  <label>Fichier Excel</label>
  <input type="file" name="file" accept=".xlsx,.xlsm" required>
  <button type="submit">Analyser</button>
  <div style="font-size:12px;color:#666">Formats : .xlsx / .xlsm</div>
</form>

{% if table_html %}
  <p><a href="{{ download_url }}">Télécharger le résultat (Excel)</a></p>
  <div class="viewport"><div class="wide">{{ table_html|safe }}</div></div>
{% endif %}

{% if message %}<pre style="white-space:pre-wrap">{{ message }}</pre>{% endif %}
</body></html>
"""

# ---------- Utilitaires ----------
def _norm(s: str) -> str:
    """Normalise un libellé (accents→ASCII, espaces insécables, casse, espaces)."""
    if not isinstance(s, str):
        s = "" if s is None else str(s)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = s.replace("\u00A0", " ").replace("\u202F", " ")
    return " ".join(s.strip().lower().split())

def _std_newlines(txt: str) -> str:
    """Uniformise tous les retours à \\n et convertit le \\n LITTÉRAL aussi."""
    txt = txt.replace("\r\n", "\n").replace("\r", "\n")
    txt = txt.replace("\\n", "\n")            # littéral \n venu d'Excel/ETL
    return txt

def _split_items(txt: str) -> list[str]:
    """Découpe un champ texte en éléments (puces, ;, retours, etc.)."""
    if not isinstance(txt, str):
        return []
    t = _std_newlines(txt)
    # remplace les diverses puces par \n pour simplifier
    t = re.sub(r"[•▪●·◦]", "\n", t)
    parts = re.split(r"\n|;", t)
    items = [p.strip(" \t•·◦;").rstrip(".").strip() for p in parts if str(p).strip()]
    # vire les 'nan' / 'NaN'
    items = [i for i in items if _norm(i) not in ("nan", "none")]
    return items

def _ul(items: Iterable[str]) -> str:
    items = [i for i in items if i]
    if not items: 
        return ""
    return "<ul class='bul'>" + "".join(f"<li>{i}</li>" for i in items) + "</ul>"

# ---------- Colonnes & alias ----------
HEADER_ALIASES: Dict[str, Set[str]] = {
    "articles_enfreints": {
        _norm("Nbr Chefs par articles"),
        _norm("Articles enfreints"),
        _norm("Articles en infraction"),
        _norm("Liste des chefs et articles en infraction"),
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
    },
    # colonnes à afficher aussi en liste (présentation)
    "resume": {_norm("Résumé des faits concis"), _norm("Resume des faits concis")},
    "liste_chefs": {_norm("Liste des chefs et articles en infraction")},
    "sanctions_imposees": {_norm("Liste des sanctions imposées"), _norm("Liste des sanctions imposees")},
    "autres_mesures": {_norm("Autres mesures ordonnées"), _norm("Autres mesures ordonnees")},
    "a_verifier": {_norm("À vérifier"), _norm("A verifier"), _norm("A vérifier")},
}

INTEREST_KEYS = ("articles_enfreints", "duree_totale_radiation", "article_amende_chef", "autres_sanctions")

def resolve_columns(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    by_norm = {_norm(c): c for c in df.columns}
    out: Dict[str, Optional[str]] = {}
    for key, aliases in HEADER_ALIASES.items():
        hit = None
        for a in aliases:
            if a in by_norm:
                hit = by_norm[a]; break
        out[key] = hit
    return out

# ---------- Lecture Excel avec bannière « Article filtré : » ----------
def read_excel_respecting_header_rule(stream) -> pd.DataFrame:
    preview = pd.read_excel(stream, header=None, nrows=2, engine="openpyxl")
    stream.seek(0)
    banner = False
    if not preview.empty and isinstance(preview.iloc[0,0], str):
        banner = _norm(preview.iloc[0,0]).startswith(_norm("Article filtré :"))
    if banner:
        return pd.read_excel(stream, skiprows=1, header=0, engine="openpyxl")
    return pd.read_excel(stream, header=0, engine="openpyxl")

# ---------- Motif exact pour l’article ----------
def build_article_pattern(user_input: str) -> re.Pattern:
    token = (user_input or "").strip()
    if not token:
        raise ValueError("Article vide.")
    esc = re.escape(token)
    tail = r"(?![\d.])" if token[-1].isdigit() else r"\b"
    # capture = groupe 1 pour surlignage
    return re.compile(rf"(?:\b(?:art(?:icle)?\s*[: ]*)?)({esc}){tail}", re.IGNORECASE)

# ---------- Rendu listes + surlignage ----------
def to_bullets_html(value: str, pat: Optional[re.Pattern] = None, highlight: bool = False) -> str:
    items = _split_items(value)
    if not items:
        return ""
    if highlight and pat is not None:
        def repl(m: re.Match) -> str:
            return f"<span class='hit'>{m.group(1)}</span>"
        items = [re.sub(pat, repl, it) for it in items]
    return _ul(items)

# ---------- Filtrage + préparation rendu ----------
def prepare_display_df(df: pd.DataFrame, colmap: Dict[str, Optional[str]], pat: re.Pattern) -> pd.DataFrame:
    disp = df.copy()

    # 1) Colonnes d’intérêt : filtrage au sens strict (on ne garde que les items qui contiennent l’article)
    for key in INTEREST_KEYS:
        col = colmap.get(key)
        if col and col in disp.columns:
            def only_hits(v: str) -> str:
                # garde UNIQUEMENT les segments contenant l'article
                hits = [seg for seg in _split_items(v) if re.search(pat, seg or "")]
                return _ul(re.sub(pat, lambda m: f"<span class='hit'>{m.group(1)}</span>", h) for h in hits)
            disp[col] = disp[col].apply(only_hits)

    # 2) Autres colonnes présentées en listes (sans filtrage)
    present_as_lists = ("resume", "liste_chefs", "sanctions_imposees", "autres_mesures", "a_verifier")
    for key in present_as_lists:
        col = colmap.get(key)
        if col and col in disp.columns:
            disp[col] = disp[col].apply(lambda v: to_bullets_html(v, None, False))

    # 3) Supprime les lignes devenues totalement vides dans les 4 colonnes d’intérêt
    interest_cols = [c for c in (colmap.get(k) for k in INTEREST_KEYS) if c]
    if interest_cols:
        mask_any = False
        for c in interest_cols:
            nonempty = disp[c].astype(str).str.strip().ne("")
            mask_any = nonempty if mask_any is False else (mask_any | nonempty)
        disp = disp[mask_any]

    return disp

# ---------- Export Excel ----------
def to_excel_download(df: pd.DataFrame) -> str:
    path = f"/tmp/filtrage_{int(time.time())}.xlsx"
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Filtre")
        ws = w.book.active
        for i, col in enumerate(df.columns, start=1):
            vals = [str(col)] + [str(x) for x in df[col].fillna("")]
            width = min(70, max(12, max(len(v) for v in vals) + 2))
            ws.column_dimensions[ws.cell(1, i).column_letter].width = width
    return f"/download?path={path}"

# ---------- Routes ----------
@app.route("/", methods=["GET","POST"])
def root():
    if request.method == "GET":
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None, searched_article=None, message=None)

    f = request.files.get("file")
    token = (request.form.get("article") or "").strip()
    if not f or not token:
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None, searched_article=token, message="Fichier et article requis.")

    # garde .xlsx/.xlsm uniquement
    name = (f.filename or "").lower()
    if not (name.endswith(".xlsx") or name.endswith(".xlsm")):
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None, searched_article=token, message="Veuillez fournir un .xlsx ou .xlsm")

    try:
        df = read_excel_respecting_header_rule(f.stream)
        colmap = resolve_columns(df)
        pat = build_article_pattern(token)

        # masque initial: l'article apparaît dans AU MOINS l'une des 4 colonnes d’intérêt
        masks = []
        for k in INTEREST_KEYS:
            c = colmap.get(k)
            if c and c in df.columns:
                masks.append(df[c].astype(str).apply(lambda v: bool(re.search(pat, v or ""))))
        if not masks:
            detail = "\n".join(f"- {k}: {colmap.get(k)}" for k in INTEREST_KEYS)
            return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None, searched_article=token,
                                          message=f"Aucune des colonnes d’intérêt n’a été trouvée.\n{detail}\nColonnes disponibles:\n{list(df.columns)}")

        m_any = masks[0]
        for m in masks[1:]: m_any = m_any | m
        df = df[m_any].copy()
        if df.empty:
            return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None, searched_article=token,
                                          message=f"Aucune ligne ne contient l’article « {token} » dans les colonnes cibles.")

        disp = prepare_display_df(df, colmap, pat)
        if disp.empty:
            return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None, searched_article=token,
                                          message="Des lignes correspondaient, mais aucun item clair après nettoyage.")

        # rendu HTML (on n’échappe pas le HTML pour conserver <ul> et le surlignage)
        html = disp.to_html(index=False, escape=False, na_rep="")

        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=html, searched_article=token,
                                      download_url=to_excel_download(disp),
                                      message=f"{len(disp)} ligne(s) après filtrage & formatage (aperçu non tronqué).")
    except Exception as e:
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None, searched_article=token,
                                      message=f"Erreur: {e!r}")

@app.route("/download")
def download():
    path = request.args.get("path") or ""
    if not path or not os.path.exists(path): return "Fichier introuvable", 404
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
