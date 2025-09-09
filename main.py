# === CANVAS META =============================================================
# Fichier : main.py — rendu HTML à puces + surlignage + largeurs (09-08b)
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
# Styles (table large + listes + surlignage + colonnes élargies)
# ──────────────────────────────────────────────────────────────────────────────
STYLE_BLOCK = """
<style>
  :root{
    --w-wide: 520px;     /* colonnes d'intérêt + colonnes textuelles larges */
    --w-date: 220px;     /* dates + À vérifier */
  }
  body{font-family:system-ui,-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;margin:24px}
  h1{font-size:20px;margin-bottom:12px}
  form{display:grid;gap:12px;margin-bottom:16px}
  input[type="text"]{padding:8px;font-size:14px}
  input[type="file"]{font-size:14px}
  button{padding:8px 12px;font-size:14px;cursor:pointer}
  .note{background:#fff6e5;border:1px solid #ffd89b;padding:8px 10px;border-radius:6px;margin:10px 0 16px}
  .kbd{font-family:ui-monospace,SFMono-Regular,Menlo,monospace;background:#f3f4f6;padding:2px 4px;border-radius:4px}

  /* Conteneur scrollable horizontal/vertical */
  .table-viewport{height:60vh;overflow:auto;border:1px solid #ddd}
  .table-viewport table{border-collapse:collapse;width:100%;font-size:13px;table-layout:auto}
  th,td{border:1px solid #ddd;padding:6px 8px;vertical-align:top;white-space:normal}
  th{background:#f3f4f6}

  /* Listes à puces propres dans les cellules */
  ul.cell-list{margin:0;padding-left:18px}
  ul.cell-list li{margin:0 0 4px 0}

  /* Surlignage de l'article dans les 4 colonnes d'intérêt */
  .hit{color:#c1121f;font-weight:600}

  /* Espaceur invisible pour forcer les largeurs de colonne */
  .wspacer{display:inline-block;height:0;line-height:0;font-size:0;overflow:hidden;vertical-align:top}
  .w-wide{width:var(--w-wide)}
  .w-date{width:var(--w-date)}
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
    Règles : détection exacte de l’article; si la 1<sup>re</sup> cellule contient « <span class="kbd">Article filtré :</span> », elle est ignorée (entêtes sur la 2<sup>e</sup> ligne).
  </div>

  <form method="POST" enctype="multipart/form-data">
    <label>Article à rechercher (ex. <span class="kbd">29</span>, <span class="kbd">59(2)</span>)</label>
    <input type="text" name="article" value="{{ searched_article or '' }}" required placeholder="ex.: 29 ou 59(2)" />
    <label>Fichier Excel</label>
    <input type="file" name="file" accept=".xlsx,.xlsm" required />
    <button type="submit">Analyser</button>
    <div class="hint">Formats : .xlsx / .xlsm</div>
  </form>

  {% if table_html %}
    <div class="download"><a href="{{ download_url }}">Télécharger le résultat (Excel)</a></div>
    <div class="table-viewport">{{ table_html|safe }}</div>
  {% endif %}

  {% if message %}
    <pre class="msg">{{ message }}</pre>
  {% endif %}
</body>
</html>
"""

# ──────────────────────────────────────────────────────────────────────────────
# Normalisation libellés
# ──────────────────────────────────────────────────────────────────────────────
def _norm(s: str) -> str:
    if not isinstance(s, str):
        s = "" if s is None else str(s)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = s.replace("\u00A0", " ")
    s = " ".join(s.strip().lower().split())
    return s

# Alias d’en-têtes (canoniques → variantes)
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
        _norm("Autres mesures ordonnées"),
        _norm("Autres sanctions / mesures"),
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
    for canon, variants in HEADER_ALIASES.items():
        out[canon] = next((norm_to_original[v] for v in variants if v in norm_to_original), None)
    return out

# ──────────────────────────────────────────────────────────────────────────────
# Lecture Excel (respect “Article filtré :” en 1re cellule)
# ──────────────────────────────────────────────────────────────────────────────
def read_excel_respecting_header_rule(file_stream) -> pd.DataFrame:
    df_preview = pd.read_excel(file_stream, header=None, nrows=2, engine="openpyxl")
    file_stream.seek(0)

    first = df_preview.iloc[0, 0] if not df_preview.empty else None
    banner = isinstance(first, str) and _norm(first).startswith(_norm("Article filtré :"))
    if banner:
        return pd.read_excel(file_stream, skiprows=1, header=0, engine="openpyxl")
    return pd.read_excel(file_stream, header=0, engine="openpyxl")

# ──────────────────────────────────────────────────────────────────────────────
# Motif exact pour l’article (capture du n° pour surlignage)
# ──────────────────────────────────────────────────────────────────────────────
def build_article_pattern(user_input: str) -> re.Pattern:
    token = (user_input or "").strip()
    if not token:
        raise ValueError("Article vide.")
    esc = re.escape(token)
    tail = r"(?![\d.])" if token[-1].isdigit() else r"\b"
    return re.compile(rf"(?:\b(?:art(?:icle)?\s*[: ]*)?)({esc}){tail}", flags=re.IGNORECASE)

# ──────────────────────────────────────────────────────────────────────────────
# Nettoyage texte + rendu HTML en listes
# ──────────────────────────────────────────────────────────────────────────────
def _isna(v) -> bool:
    try:
        return pd.isna(v)
    except Exception:
        return False

def _prep_text(v) -> str:
    """Nettoie pour la recherche ET l’affichage (supprime puces, NBSP, CR/LF)."""
    if _isna(v):
        return ""
    if not isinstance(v, str):
        v = str(v)
    v = v.replace("•", " ").replace("·", " ").replace("◦", " ")
    v = v.replace("\u00A0", " ").replace(" ", " ")
    v = v.replace("\r\n", "\n").replace("\r", "\n")
    v = " ".join(v.split())
    return v

def _highlight(text: str, pat: re.Pattern) -> str:
    if not text:
        return ""
    # Important : utiliser une fonction pour éviter les "\1" littéraux.
    return pat.sub(lambda m: f'<span class="hit">{m.group(1)}</span>', text)

def _to_bullets(raw: str, *, highlight=False, pat: Optional[re.Pattern]=None) -> str:
    """Transforme une chaîne en <ul><li>…</li></ul> propre.
       Séparateurs gérés : ' | ' (pipeline), ; , nouvelle ligne, puces.
    """
    s = _prep_text(raw)
    if not s:
        return ""
    # Si le texte vient déjà de l’étape d’extraction, il est séparé par " | ".
    s = (
        s.replace(" | ", "\n")
         .replace("|", "\n")
         .replace("•", "\n")
         .replace(";", "\n")
    )
    parts = [p.strip(" -.;") for p in s.split("\n") if p and p.strip().lower() not in {"nan", "none"}]
    if not parts:
        return ""
    if highlight and pat is not None:
        parts = [_highlight(p, pat) for p in parts]
    lis = "".join(f"<li>{p}</li>" for p in parts)
    return f"<ul class='cell-list'>{lis}</ul>"

# ──────────────────────────────────────────────────────────────────────────────
# Extraction “pipe” des seules mentions contenant l’article
# ──────────────────────────────────────────────────────────────────────────────
def extract_mentions_generic(text: str, pat: re.Pattern) -> str:
    s = _prep_text(text)
    if not s:
        return ""
    pieces = re.split(r"[;,\n]", s)
    hits = [p.strip() for p in pieces if pat.search(p)]
    return " | ".join(hits)

def extract_mentions_autres_sanctions(text: str, pat: re.Pattern) -> str:
    s = _prep_text(text)
    if not s:
        return ""
    hits = [seg.strip() for seg in re.split(r"[;\n]", s) if pat.search(seg)]
    return " | ".join(hits)

def clean_filtered_df(df: pd.DataFrame, colmap: Dict[str, Optional[str]], pat: re.Pattern) -> pd.DataFrame:
    df = df.copy()
    for canon in FILTER_CANONICAL:
        col = colmap.get(canon)
        if not col or col not in df.columns:
            continue
        if canon == "autres_sanctions":
            df[col] = df[col].apply(lambda v: extract_mentions_autres_sanctions(v, pat))
        else:
            df[col] = df[col].apply(lambda v: extract_mentions_generic(v, pat))

    # Garder les lignes où au moins 1 des 4 colonnes a un contenu non vide
    targets = [c for c in (colmap.get(k) for k in FILTER_CANONICAL) if c]
    if targets:
        mask_any = False
        for c in targets:
            cur = df[c].astype(str).str.strip().ne("")
            mask_any = cur if mask_any is False else (mask_any | cur)
        df = df[mask_any]
    return df

# ──────────────────────────────────────────────────────────────────────────────
# Export Excel (auto-largeur)
# ──────────────────────────────────────────────────────────────────────────────
def to_excel_download(df: pd.DataFrame) -> str:
    ts = int(time.time())
    out_path = f"/tmp/filtrage_{ts}.xlsx"
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Filtre")
        ws = writer.book.active
        for j, col in enumerate(df.columns, start=1):
            vals = [str(x) for x in df[col].fillna("").tolist()]
            max_len = max([len(col)] + [len(s) for s in vals] + [12])
            ws.column_dimensions[ws.cell(row=1, column=j).column_letter].width = min(60, max_len + 2)
    return f"/download?path={out_path}"

# ──────────────────────────────────────────────────────────────────────────────
# Route principale
# ──────────────────────────────────────────────────────────────────────────────
@app.route("/", methods=["GET", "POST"])
def analyze():
    if request.method == "GET":
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                                      searched_article=None, message=None, download_url=None)

    # Validations
    file = request.files.get("file")
    article = (request.form.get("article") or "").strip()
    if not file or not article:
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                                      searched_article=article,
                                      message="Erreur : fichier et article sont requis.",
                                      download_url=None)

    fname = (file.filename or "").lower()
    if not (fname.endswith(".xlsx") or fname.endswith(".xlsm")):
        ext = os.path.splitext(file.filename or "")[-1]
        return render_template_string(
            HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None, searched_article=article, download_url=None,
            message=f"Format non pris en charge : {ext}. Veuillez fournir un .xlsx ou .xlsm."
        )

    try:
        # Lecture + mapping d'en-têtes
        df = read_excel_respecting_header_rule(file.stream)
        colmap = resolve_columns(df)
        pat = build_article_pattern(article)

        # Pré-filtrage : lignes où l’article apparaît dans AU MOINS une des 4 colonnes d’intérêt
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
                HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None, searched_article=article, download_url=None,
                message=("Aucune des colonnes attendues n'a été trouvée.\n\n"
                         f"Résolution des colonnes:\n{detail}\n\nColonnes du fichier:\n{list(df.columns)}")
            )

        mask_any = masks[0]
        for m in masks[1:]:
            mask_any = mask_any | m
        df = df[mask_any].copy()
        if df.empty:
            return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
                                          searched_article=article, download_url=None,
                                          message=f"Aucune ligne ne contient l’article « {article} » dans les colonnes cibles.")

        # Extraction pipeline + nettoyage
        df_clean = clean_filtered_df(df, colmap, pat)
        if df_clean.empty:
            return render_template_string(
                HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None, searched_article=article, download_url=None,
                message=("Des lignes correspondaient, mais après épuration des cellules, "
                         "aucune mention nette n’a été conservée.")
            )

        # Rendu HTML : transformer en listes + surligner l’article dans les 4 colonnes d’intérêt
        display = df_clean.copy()

        # Colonnes d’intérêt (original labels)
        c_art = colmap.get("articles_enfreints")
        c_rad = colmap.get("duree_totale_radiation")
        c_amd = colmap.get("article_amende_chef")
        c_rep = colmap.get("autres_sanctions")

        # Autres colonnes textuelles à rendre en listes
        def find_col(label: str) -> Optional[str]:
            tgt = _norm(label)
            for c in display.columns:
                if _norm(c) == tgt:
                    return c
            return None

        c_resume = find_col("Résumé des faits concis")
        c_list_chefs = find_col("Liste des chefs et articles en infraction")
        c_sanctions = find_col("Liste des sanctions imposées")
        c_autres_mesures = find_col("Autres mesures ordonnées")
        c_a_verifier = find_col("À vérifier") or find_col("A verifier")
        c_nbr_par_articles = c_art  # même colonne, affichage en liste
        # Colonnes dates
        c_date_cre = find_col("Date de création") or find_col("Date de creation")
        c_date_maj = find_col("Date de mise à jour") or find_col("Date de mise a jour")

        # Appliquer listes (avec surlignage uniquement sur les 4 colonnes d'intérêt)
        for col in [c_resume, c_list_chefs, c_sanctions, c_autres_mesures, c_a_verifier]:
            if col:
                display[col] = display[col].apply(lambda v: _to_bullets(v, highlight=False))

        for col in [c_art, c_rad, c_amd, c_rep]:
            if col:
                display[col] = display[col].apply(lambda v: _to_bullets(v, highlight=True, pat=pat))

        # Forcer les largeurs via un "spacer" invisible au début des cellules des colonnes larges
        wide_cols = [c for c in [c_resume, c_list_chefs, c_sanctions, c_art, c_rad, c_amd, c_rep] if c]
        for col in wide_cols:
            display[col] = display[col].apply(lambda v: f"<span class='wspacer w-wide'></span>{'' if _isna(v) else v}")

        date_cols = [c for c in [c_date_cre, c_date_maj, c_a_verifier] if c]
        for col in date_cols:
            display[col] = display[col].apply(lambda v: f"<span class='wspacer w-date'></span>{'' if _isna(v) else _to_bullets(v) if col==c_a_verifier else ('' if _isna(v) else str(v))}")

        # Convertir en HTML (ne pas échapper pour laisser nos listes/surlignages)
        preview = display.head(200)
        table_html = preview.to_html(index=False, escape=False)

        # Lien Excel (contenu non transformé en HTML)
        download_url = to_excel_download(df_clean)

        msg = f"{len(df_clean)} ligne(s) après filtrage. (Aperçu limité à 200 lignes.)"
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=table_html,
                                      searched_article=article, download_url=download_url, message=msg)

    except Exception as e:
        return render_template_string(
            HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=None,
            searched_article=article, download_url=None,
            message=f"Erreur inattendue : {repr(e)}"
        )

# ──────────────────────────────────────────────────────────────────────────────
# Téléchargement
# ──────────────────────────────────────────────────────────────────────────────
@app.route("/download")
def download():
    path = request.args.get("path")
    if not path or not os.path.exists(path):
        return "Fichier introuvable ou expiré.", 404
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
