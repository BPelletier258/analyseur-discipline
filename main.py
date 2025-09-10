# -*- coding: utf-8 -*-
import re
import math
from io import BytesIO

import pandas as pd
from flask import Flask, render_template, request, send_file
from markupsafe import Markup

app = Flask(__name__)

# In-memory last Excel file (simple single-user cache)
LAST_XLSX = None

# =========================
# ----  UI / STYLES    ----
# =========================
CSS = r"""
<style>
:root{
  --w-s: 8.5rem;   /* étroit */
  --w-m: 12rem;    /* moyen  */
  --w-l: 18rem;    /* large  */
  --w-xl: 26rem;   /* très large */
}
*{box-sizing:border-box}
body{font-family: ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif;}

/* container + header note */
.note{background:#fff8e6;border:1px solid #ffd48a;padding:10px 12px;border-radius:8px;margin:0 0 14px 0}
.wrap {max-width: 1280px; margin: 24px auto; padding: 0 16px;}
/* make table area span full width */
.full { width: 100%; }

/* form layout */
.formcard{
  background:#f7fafc; border:1px solid #e5e7eb; border-radius:12px; padding:16px;
  display:grid; gap:12px; grid-template-columns: 1fr auto; align-items:end;
}
.form-left{ display:grid; grid-template-columns:1fr auto; gap:10px; align-items:center; }
.form-row{ display:block; }
.form-row label{ display:block; font-weight:600; margin-bottom:6px; color:#374151; }
.form-input{ width:100%; height:40px; padding:8px 10px; border:1px solid #d1d5db; border-radius:8px; }
.form-check{ display:flex; align-items:center; gap:8px; color:#374151; }
.file-row{ display:flex; align-items:center; gap:10px; }

.btn{
  height:40px; padding:0 18px; border-radius:10px; border:0;
  color:#fff; background:#0ea5e9; cursor:pointer; font-weight:700;
}
.btn:disabled{ opacity:.7; cursor:wait; }

.download { margin:10px 0 12px; }
.download a{
  display:inline-block; background:#2563eb; color:#fff; padding:10px 16px;
  border-radius:10px; text-decoration:none; font-weight:700;
}

/* Table viewport */
.viewport{height:60vh; overflow:auto; border:1px solid #e5e7eb; border-radius:10px;}

/* TABLE */
table{width:100%; border-collapse:collapse; table-layout:fixed;}
th,td{
  border:1px solid #e5e7eb; padding:6px 8px; vertical-align:top;
  white-space:normal; word-break:normal; overflow-wrap:anywhere; hyphens:auto;
}
th{position:sticky; top:0; background:#f8fafc; z-index:2; font-weight:700; text-align:center;}

ul{margin:0; padding-left:1.05rem}
li{margin:0.1rem 0}
.no-bullets ul{list-style:none; padding-left:0; margin:0}
.empty{color:#9CA3AF;}  /* tiret gris */
.hit{color:#c00; font-weight:700}

/* width classes */
.col-s { width:var(--w-s);  min-width:var(--w-s)}
.col-m { width:var(--w-m);  min-width:var(--w-m)}
.col-l { width:var(--w-l);  min-width:var(--w-l)}
.col-xl{ width:var(--w-xl); min-width:var(--w-xl)}

/* spinner overlay */
#overlay {
  position:fixed; inset:0; background:rgba(255,255,255,.75);
  display:none; align-items:center; justify-content:center; z-index:9999;
}
.spinner { width:52px; height:52px; border:6px solid #e5e7eb; border-top-color:#0ea5e9;
  border-radius:50%; animation:spin 1s linear infinite; }
@keyframes spin { to { transform: rotate(360deg);} }
</style>
"""

# =========================
# ----  CONFIG COLS   -----
# =========================

# 4 colonnes d’intérêt
INTEREST_COLS = [
    "Nbr Chefs par articles",
    "Nbr Chefs par articles par période de radiation",
    "Nombre de chefs par articles et total amendes",
    "Nombre de chefs par article ayant une réprimande",
]

# Colonnes "listes" (rendues avec puces si >1 item)
LIST_COLUMNS = {
    "Résumé des faits concis",
    "Liste des chefs et articles en infraction",
    "Nbr Chefs par articles",
    "Nbr Chefs par articles par période de radiation",
    "Liste des sanctions imposées",
    "Nombre de chefs par articles et total amendes",
    "Nombre de chefs par article ayant une réprimande",
    "Autres mesures ordonnées",
    "À vérifier",
}

# Colonnes à ne jamais surligner (même si l’article est présent)
NO_HIGHLIGHT_COLS = {
    "Liste des chefs et articles en infraction",
    "Liste des sanctions imposées",
}

# Largeurs par colonne
WIDTH_CLASS = {
    # scalaires
    "Nom de l'intimé": "col-l",
    "Ordre professionnel": "col-l",
    "Numéro de la décision": "col-m",
    "Date de la décision rendue": "col-m",
    "Nature de la décision": "col-m",
    "Période des faits": "col-m",
    "Plaidoyer de culpabilité": "col-s",
    "Total chefs": "col-s",
    "Radiation max": "col-s",
    "Total amendes": "col-m",
    "Total réprimandes": "col-s",
    "À vérifier": "col-l",
    "Date de création": "col-m",
    "Date de mise à jour": "col-m",

    # colonnes “listes”
    "Résumé des faits concis": "col-xl",
    "Liste des chefs et articles en infraction": "col-xl",
    "Liste des sanctions imposées": "col-l",
    "Nbr Chefs par articles": "col-l",
    "Nbr Chefs par articles par période de radiation": "col-l",
    "Nombre de chefs par articles et total amendes": "col-l",
    "Nombre de chefs par article ayant une réprimande": "col-l",
    "Autres mesures ordonnées": "col-l",
}

EMPTY_SPAN = "<span class='empty'>—</span>"

# =========================
# ----  HELPERS       -----
# =========================

def compile_article_pattern(article: str) -> re.Pattern:
    token = re.escape(article.strip())
    return re.compile(rf"(?<!\d){token}(?!\d)", flags=re.IGNORECASE)

def keep_only_rows_with_article_in_core(df: pd.DataFrame, pattern: re.Pattern) -> pd.DataFrame:
    core_col = "Nbr Chefs par articles"
    if core_col not in df.columns:
        return df
    mask = df[core_col].astype(str).apply(lambda s: bool(pattern.search(s)))
    return df[mask].reset_index(drop=True)

def _safe_str(x) -> str:
    if x is None or (isinstance(x, float) and math.isnan(x)):
        return ""
    return str(x).strip()

def fmt_amount(x) -> str:
    s = _safe_str(x)
    if s == "":
        return ""
    try:
        val = float(str(s).replace(" ", "").replace("\xa0", ""))
        if abs(val) < 0.005:
            return "0 $"
        ints = f"{int(round(val)):,.0f}".replace(",", " ").replace("\xa0"," ")
        return f"{ints} $"
    except Exception:
        return s

def split_items(text: str) -> list[str]:
    """Découpe léger en items (retours / puces / ; / '- ')"""
    if not text:
        return []
    t = text.replace("•", "\n").replace("\r", "\n")
    parts = re.split(r"\n|;|\u2022|- ", t)
    parts = [p.strip(" •\t") for p in parts if p and p.strip(" •\t")]
    return parts if parts else [text.strip()]

def only_segment_with_token(text: str, pattern: re.Pattern) -> str:
    """Ne garde que les items contenant le token (pour les 4 colonnes d’intérêt)."""
    if not text:
        return ""
    items = split_items(text)
    keep = [i for i in items if pattern.search(i)]
    return "\n".join(keep)

def to_bullets(text: str, bulletize: bool) -> str:
    if not text:
        return ""
    items = split_items(text)
    if not bulletize or len(items) == 1:
        return items[0]
    lis = "\n".join(f"<li>{p}</li>" for p in items)
    return f"<ul>{lis}</ul>"

def highlight_html(text: str, pattern: re.Pattern) -> str:
    if not text:
        return ""
    return pattern.sub(lambda m: f'<span class="hit">{m.group(0)}</span>', text)

def render_cell(value: str, column_name: str, pattern: re.Pattern, show_only_segment: bool) -> str:
    raw = _safe_str(value)

    # Format monétaire
    if column_name == "Total amendes":
        raw = fmt_amount(raw)

    # Si case cochée et colonne d’intérêt => ne garder que les segments qui contiennent l’article
    if show_only_segment and column_name in INTEREST_COLS:
        raw = only_segment_with_token(raw, pattern)

    # Surlignage seulement dans les colonnes d’intérêt (pas dans NO_HIGHLIGHT_COLS)
    will_highlight = (column_name in INTEREST_COLS) and (column_name not in NO_HIGHLIGHT_COLS)
    if will_highlight:
        raw = highlight_html(raw, pattern)

    # Bullets ?
    is_list_col = column_name in LIST_COLUMNS
    html = to_bullets(raw, bulletize=is_list_col)
    cls = "" if is_list_col else " no-bullets"
    display = html if html else EMPTY_SPAN
    return f'<div class="{cls.strip()}">{display}</div>'

def build_html_table(df: pd.DataFrame, article: str, show_only_segment: bool) -> str:
    pattern = compile_article_pattern(article)
    headers = list(df.columns)
    out = [CSS, '<div class="viewport full"><table>']
    out.append("<thead><tr>")
    for h in headers:
        out.append(f'<th class="{WIDTH_CLASS.get(h, "col-m")}">{h}</th>')
    out.append("</tr></thead><tbody>")
    for _, row in df.iterrows():
        out.append("<tr>")
        for h in headers:
            cell_html = render_cell(row.get(h, ""), h, pattern, show_only_segment)
            out.append(f'<td class="{WIDTH_CLASS.get(h, "col-m")}">{cell_html}</td>')
        out.append("</tr>")
    out.append("</tbody></table></div>")
    return "\n".join(out)

# ---------- Excel helpers (rich in-cell highlight for interest columns) ----------

def rich_segments(text: str, pattern: re.Pattern):
    """
    Build segments for xlsxwriter write_rich_string:
    returns list like ["pre", fmt_hit, "match", "mid", fmt_hit, "match2", "post"]
    """
    if not text:
        return [""]
    parts = []
    last = 0
    for m in pattern.finditer(text):
        if m.start() > last:
            parts.append(text[last:m.start()])
        parts.append(("HIT", m.group(0)))  # mark hit
        last = m.end()
    if last < len(text):
        parts.append(text[last:])
    if not parts:
        return [text]
    # ensure string first/last as required by write_rich_string
    seq = []
    if isinstance(parts[0], tuple):
        seq.append("")  # leading empty normal
    for p in parts:
        if isinstance(p, tuple):
            # placeholder; will map to format later
            seq.append(("HIT", p[1]))
        else:
            seq.append(p)
    if isinstance(seq[-1], tuple):
        seq.append("")  # trailing empty normal
    return seq

def write_cell_with_highlight(ws, row, col, text, pattern, fmt_text, fmt_hit):
    """Write text with in-cell highlight for matches (xlsxwriter)."""
    segs = rich_segments(text, pattern)
    # If no ("HIT", ...) present, just write normally
    has_hit = any(isinstance(s, tuple) and s[0] == "HIT" for s in segs)
    if not has_hit:
        ws.write(row, col, text, fmt_text)
        return
    # Build args for write_rich_string
    args = []
    for s in segs:
        if isinstance(s, tuple) and s[0] == "HIT":
            args += [fmt_hit, s[1]]
        else:
            args.append(s)
    ws.write_rich_string(row, col, *args, fmt_text)

def produce_view_and_excel(df_source: pd.DataFrame, article: str, show_only_segment: bool):
    """Return (html, BytesIO)"""
    pattern = compile_article_pattern(article)

    # 1) HTML
    html = build_html_table(df_source, article=article, show_only_segment=show_only_segment)

    # 2) Excel
    bio = BytesIO()
    df = df_source.copy()

    # Isoler segments dans Excel aussi (4 colonnes d’intérêt)
    if show_only_segment:
        for c in INTEREST_COLS:
            if c in df.columns:
                df[c] = df[c].astype(str).apply(lambda s: only_segment_with_token(s, pattern))

    # Format monétaire
    if "Total amendes" in df.columns:
        df["Total amendes"] = df["Total amendes"].map(fmt_amount)

    with pd.ExcelWriter(bio, engine="xlsxwriter") as xw:
        sheet = "Résultat"
        df.to_excel(xw, index=False, startrow=1, sheet_name=sheet)
        ws = xw.sheets[sheet]

        # En-tête "Article filtré : X"
        ws.write(0, 0, f"Article filtré : {article}")

        # Formats
        wb = xw.book
        fmt_text = wb.add_format({"valign": "top", "text_wrap": True})
        fmt_hit  = wb.add_format({"valign": "top", "text_wrap": True, "bold": True, "font_color": "#cc0000"})

        # fige la ligne d’en-têtes (ligne 2 visuelle)
        ws.freeze_panes(2, 0)

        # Inject rich highlight only in interest cols
        for r in range(len(df)):
            for c, colname in enumerate(df.columns):
                v = _safe_str(df.iat[r, c])
                if colname in INTEREST_COLS:
                    write_cell_with_highlight(ws, r + 1, c, v, pattern, fmt_text, fmt_hit)
                else:
                    ws.write(r + 1, c, v, fmt_text)

        # Adjust widths
        for c, colname in enumerate(df.columns):
            # crude auto-width
            maxlen = max([len(str(x)) for x in [colname] + df[colname].astype(str).tolist()] + [12])
            ws.set_column(c, c, min(60, int(maxlen * 1.1)))

    bio.seek(0)
    return html, bio

# =========================
# ----  READ FILE      ----
# =========================

def read_input_excel(file_storage) -> pd.DataFrame:
    """
    - If first cell contains 'Article filtré :', header is on 2nd row.
    - Otherwise, first row is the header.
    """
    # First pass
    df0 = pd.read_excel(file_storage, header=None, engine="openpyxl")
    if df0.shape[0] >= 2 and df0.iloc[0, 0] and isinstance(df0.iloc[0, 0], str) and "Article filtré" in df0.iloc[0, 0]:
        headers = df0.iloc[1].tolist()
        df = df0.iloc[2:].copy()
        df.columns = headers
    else:
        file_storage.seek(0)
        df = pd.read_excel(file_storage, engine="openpyxl")
    # Normalize col names
    df.columns = [str(c).strip() for c in df.columns]
    return df

# =========================
# ----  ROUTES         ----
# =========================

@app.route("/", methods=["GET", "POST"])
def index():
    global LAST_XLSX
    html_table = None
    error = None

    if request.method == "POST":
        try:
            article = request.form.get("article", "").strip()
            show_only_segment = bool(request.form.get("segment_only"))

            file = request.files.get("file")
            if not article:
                raise ValueError("Veuillez entrer un numéro d’article.")
            if not file:
                raise ValueError("Veuillez choisir un fichier Excel (.xlsx/.xlsm).")

            df = read_input_excel(file)

            # ❶ keep only rows that have the article in *Nbr Chefs par articles*
            pattern = compile_article_pattern(article)
            df = keep_only_rows_with_article_in_core(df, pattern)

            if df.empty:
                html_table = Markup(CSS + '<div class="note">Aucune ligne ne contient l’article recherché dans « Nbr Chefs par articles ».</div>')
            else:
                html, bio = produce_view_and_excel(df, article, show_only_segment)
                html_table = Markup(html)
                LAST_XLSX = bio

        except Exception as e:
            error = f"Échec de l’analyse : {e}"

    return render_template("index.html",
                           css=Markup(CSS),
                           table_html=html_table,
                           error=error)

@app.route("/download")
def download():
    global LAST_XLSX
    if not LAST_XLSX:
        # no current export
        return "Aucun résultat disponible.", 404
    LAST_XLSX.seek(0)
    return send_file(LAST_XLSX,
                     as_attachment=True,
                     download_name="resultat.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8000, debug=False)
