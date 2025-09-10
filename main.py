# -*- coding: utf-8 -*-
import re
import math
import pandas as pd
from io import BytesIO

# =========================
# ----  CONFIG LARGEURS ---
# =========================
# Quatre classes de largeur, faciles à ajuster
CSS = r"""
<style>
:root{
  --w-s: 8.5rem;     /* étroit  */
  --w-m: 12rem;      /* moyen   */
  --w-l: 18rem;      /* large   */
  --w-xl: 26rem;     /* très large */
}
*{box-sizing:border-box}
body{font-family: ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif;}
.note{background:#fff8e6;border:1px solid #ffd48a;padding:8px 10px;border-radius:6px;margin:8px 0 14px}
.viewport{height:60vh;overflow:auto;border:1px solid #ddd}

table{width:100%; border-collapse:collapse; table-layout:fixed;}
th,td{border:1px solid #e5e7eb; padding:6px 8px; vertical-align:top;
      white-space:normal; /* << évite les retours lettre par lettre */
      word-break:normal;  /* << pas de cassure agressive */
      overflow-wrap:anywhere; hyphens:auto; }

th{position:sticky; top:0; background:#f8fafc; z-index:1; font-weight:600; text-align:center;}

ul{margin:0; padding-left:1.05rem}
li{margin:0.1rem 0}
.no-bullets ul{list-style:none; padding-left:0; margin:0}
.empty{color:#9CA3AF;}  /* tiret gris */
.hit{color:#c00; font-weight:700}

/* Largeurs :
   Applique .col-s / .col-m / .col-l / .col-xl à <th> et <td> */
.col-s { width:var(--w-s);  min-width:var(--w-s)}
.col-m { width:var(--w-m);  min-width:var(--w-m)}
.col-l { width:var(--w-l);  min-width:var(--w-l)}
.col-xl{ width:var(--w-xl); min-width:var(--w-xl)}
</style>
"""

# Choix des largeurs par en-tête (tu peux ajuster)
WIDTH_CLASS = {
    # scalaires
    "Nom de l'intimé": "col-l",
    "Ordre professionnel": "col-l",
    "Numéro de la décision": "col-m",
    "Date de la décision rendue": "col-m",
    "Nature de la décision": "col-m",
    "Période des faits": "col-m",
    "Plaidoyer de culpabilité": "col-s",
    "Nbr Chefs par articles": "col-m",
    "Total chefs": "col-s",
    "Radiation max": "col-s",
    "Nombre de chefs par articles et total amendes": "col-l",
    "Total amendes": "col-m",
    "Nombre de chefs par article ayant une réprimande": "col-l",
    "Total réprimandes": "col-s",
    "À vérifier": "col-l",
    "Date de création": "col-m",
    "Date de mise à jour": "col-m",

    # colonnes “listes”
    "Résumé des faits concis": "col-xl",
    "Liste des chefs et articles en infraction": "col-xl",
    "Liste des sanctions imposées": "col-l",
    "Nbr Chefs par articles par période de radiation": "col-l",
    "Autres mesures ordonnées": "col-l",
}

# Colonnes qui doivent être rendues sous forme de liste à puces (si le texte comporte plusieurs items)
LIST_COLUMNS = {
    "Résumé des faits concis",
    "Liste des chefs et articles en infraction",
    "Nbr Chefs par articles",
    "Nbr Chefs par articles par période de radiation",
    "Liste des sanctions imposées",
    "Nombre de chefs par article ayant une réprimande",
    "Autres mesures ordonnées",
    "À vérifier",
}

# Colonnes d’intérêt pour l’option “segment uniquement”
INTEREST_COLS = [
    "Résumé des faits concis",
    "Liste des chefs et articles en infraction",
    "Nbr Chefs par articles",
    "Liste des sanctions imposées",
]

# =========================
# ----  UTILITAIRES   -----
# =========================
EMPTY_SPAN = "<span class='empty'>—</span>"

def render_cell(value: str, column_name: str, bulletize: bool, show_only_segment: bool, pattern: re.Pattern) -> str:
    raw = _safe_str(value)

    if column_name == "Total amendes":
        raw = fmt_amount(raw)

    if show_only_segment and column_name in INTEREST_COLS:
        items = split_items(raw)
        items = [highlight(x, pattern) for x in items if pattern.search(x)]
        raw = "\n".join(items)

    raw = highlight(raw, pattern)

    is_list_col = column_name in LIST_COLUMNS
    html = to_bullets(raw, bulletize=is_list_col)

    cls = "" if is_list_col else " no-bullets"

    # ✅ calcule d’abord ce qu’on insère, sans backslash dans l’expression du f-string
    display = html if html else EMPTY_SPAN
    return f'<div class="{cls.strip()}">{display}</div>'

def _safe_str(x) -> str:
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return ""
    return str(x).strip()

def fmt_amount(x) -> str:
    """Format 0 -> 0 $, 5000 -> 5 000 $"""
    s = _safe_str(x)
    if s == "":
        return ""
    try:
        val = float(str(s).replace(" ", "").replace("\xa0",""))
        if abs(val) < 0.005:
            return "0 $"
        ints = f"{int(round(val)):,.0f}".replace(",", " ").replace("\xa0"," ")
        return f"{ints} $"
    except Exception:
        return s

def highlight(text: str, pattern: re.Pattern) -> str:
    if not text:
        return ""
    return pattern.sub(lambda m: f'<span class="hit">{m.group(0)}</span>', text)

def split_items(text: str) -> list[str]:
    """Découpage léger en items; garde la ponctuation de base."""
    if not text:
        return []
    # Normalise les séparateurs fréquents
    t = text.replace("•", "\n").replace("\r", "\n")
    # On coupe sur retours, point-virgule et puces textuelles
    parts = re.split(r"\n|;|\u2022|- ", t)
    # Nettoyage
    parts = [p.strip(" •\t") for p in parts if p and p.strip(" •\t")]
    return parts if parts else [text.strip()]

def to_bullets(text: str, bulletize: bool) -> str:
    """Rend en <ul><li> si bulletize=True et qu'il y a plusieurs items ; sinon renvoie le texte brut."""
    if not text:
        return ""
    items = split_items(text)
    if not bulletize or len(items) == 1:
        # mono-élément
        return items[0]
    lis = "\n".join(f"<li>{p}</li>" for p in items)
    return f"<ul>{lis}</ul>"

def render_cell(value: str, column_name: str, bulletize: bool, show_only_segment: bool, pattern: re.Pattern) -> str:
    """Rendu HTML d'une cellule selon le type de colonne et les options."""
    raw = _safe_str(value)

    # Mise en forme des amendes
    if column_name == "Total amendes":
        raw = fmt_amount(raw)

    # Si segment_only et colonne d’intérêt : on ne garde que les items contenant l’article
    if show_only_segment and column_name in INTEREST_COLS:
        items = split_items(raw)
        items = [highlight(x, pattern) for x in items if pattern.search(x)]
        raw = "\n".join(items)

    # Surlignage (même si segment_only est False)
    raw = highlight(raw, pattern)

    # Décide si on met des puces
    is_list_col = column_name in LIST_COLUMNS
    html = to_bullets(raw, bulletize=is_list_col)

    # Classe pour supprimer les puces si nécessaire
    cls = "" if is_list_col else " no-bullets"
    return f'<div class="{cls.strip()}">{html or "<span class=\'empty\'>—</span>"}</div>'

def build_html_table(df: pd.DataFrame, article: str, show_only_segment: bool) -> str:
    # Prépare le pattern (exact, on échappe l’entrée)
    token = re.escape(article.strip())
    pattern = re.compile(rf"(?<!\d){token}(?!\d)", flags=re.IGNORECASE)

    # Construction entêtes
    headers = list(df.columns)

    # Start HTML
    html = [CSS, '<div class="viewport"><table>']
    # thead
    html.append("<thead><tr>")
    for h in headers:
        cls = WIDTH_CLASS.get(h, "col-m")
        html.append(f'<th class="{cls}">{h}</th>')
    html.append("</tr></thead>")

    # tbody
    html.append("<tbody>")
    for _, row in df.iterrows():
        html.append("<tr>")
        for h in headers:
            cls = WIDTH_CLASS.get(h, "col-m")
            cell = render_cell(row.get(h, ""), h, bulletize=True, show_only_segment=show_only_segment, pattern=pattern)
            html.append(f'<td class="{cls}">{cell}</td>')
        html.append("</tr>")
    html.append("</tbody></table></div>")
    return "\n".join(html)

# =========================
# ----  EXEMPLE D’USAGE ---
# =========================
def produce_view_and_excel(df_source: pd.DataFrame, article: str, show_only_segment: bool):
    """
    df_source : DataFrame déjà filtré sur l’article (lignes pertinentes)
    article   : '29' ou '59(2)' etc.
    show_only_segment : True = isoler les segments dans les 4 colonnes d’intérêt
    """
    # 1) Rendu HTML
    html = build_html_table(df_source, article=article, show_only_segment=show_only_segment)

    # 2) Export Excel (en-tête "Article filtré : X" ligne 1)
    bio = BytesIO()
    out = df_source.copy()
    # Format amendes à l’export
    if "Total amendes" in out.columns:
        out["Total amendes"] = out["Total amendes"].map(fmt_amount)
    with pd.ExcelWriter(bio, engine="xlsxwriter") as xw:
        out.to_excel(xw, index=False, startrow=1, sheet_name="Résultat")
        ws = xw.sheets["Résultat"]
        ws.write(0, 0, f"Article filtré : {article}")
        ws.freeze_panes(2, 0)  # fige la ligne d’en-têtes (ligne 2 réelle)
        # largeur auto simple
        for col_idx, col_name in enumerate(out.columns):
            width = max(12, min(60, int(out[col_name].astype(str).map(len).max() * 1.1)))
            ws.set_column(col_idx, col_idx, width)
    bio.seek(0)
    return html, bio  # html pour l’affichage, bio pour "Télécharger le résultat (Excel)"
