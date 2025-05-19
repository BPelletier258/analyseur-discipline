
import re
from io import BytesIO
from flask import Flask, request, render_template_string, send_file
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows

app = Flask(__name__)

# ――― Paramètres ―――
TARGET_ARTICLE = r"\bArt[.]\s*14\b"  # vous pourrez remplacer 14 dynamiquement si besoin
CELL_HIGHLIGHT = PatternFill(fill_type="solid", fgColor="FFCCCC")  # rose pâle
RED_FONT = Font(color="FF0000")
LINK_FONT = Font(underline="single", color="0000FF")

HTML_TEMPLATE = """
<!doctype html>
<html lang="fr">
<head>
  <meta charset="utf-8">
  <title>Analyse Discipliniare</title>
  <style>
    body { font-family: sans-serif; padding: 1em; }
    .table-container { overflow-x: auto; -webkit-overflow-scrolling: touch; }
    table { border-collapse: collapse; width: 100%; }
    th, td { border: 1px solid #ccc; padding: 4px 8px; white-space: nowrap; }
    th { background: #f5f5f5; }
  </style>
</head>
<body>
  <h1>Analyse Discipliniare</h1>
  <form method="POST" enctype="multipart/form-data">
    <label>Fichier Excel: <input type="file" name="excel"></label><br><br>
    <label>Article à filtrer: <input name="article" value="{{article}}"></label><br><br>
    <button>Analyser</button>
  </form>
  {% if table_html %}
    <hr>
    <div class="table-container">
      {{ table_html|safe }}
    </div>
    <p><a href="{{ download_url }}">⬇️ Télécharger le fichier Excel formaté</a></p>
  {% endif %}
</body>
</html>
"""

@app.route("/", methods=["GET", "POST"])
def analyze():
    table_html = download_url = None
    article = "14"
    if request.method == "POST":
        f = request.files.get("excel")
        article = request.form.get("article", article).strip()
        regex = re.compile(rf"\bArt[.]\s*{article}\b")
        df = pd.read_excel(f)
        # normalise noms de colonnes, etc., si nécessaire
        # filtrage strict
        mask = df.apply(lambda row: row.astype(str).str.contains(regex).any(), axis=1)
        filtered = df[mask].copy()
        # ajoute la colonne Résumé en hyperlien
        filtered["Résumé"] = filtered["Résumé"].apply(
            lambda url: f'<a href="{url}">Résumé</a>'
        )
        # colorisation Excel : on fera dans la génération .xlsx
        # conversion en table HTML
        table_html = filtered.to_html(index=False, escape=False)
        # sauvegarde en mémoire du fichier Excel
        bio = BytesIO()
        export_excel(filtered, bio, regex)
        bio.seek(0)
        download_url = "/download"
        # stocker en session ou global pour /download
        app.config["EXCEL_BIO"] = bio

    return render_template_string(
        HTML_TEMPLATE,
        table_html=table_html,
        download_url=download_url,
        article=article
    )

@app.route("/download")
def download():
    bio = app.config.get("EXCEL_BIO")
    return send_file(
        bio,
        as_attachment=True,
        download_name=f"décisions_article_{request.args.get('article','14')}_formaté.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

def export_excel(df, stream, regex):
    wb = Workbook()
    ws = wb.active
    ws.title = "Décisions filtrées"

    # écrire l’entête
    ws.append(list(df.columns))
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = PatternFill(fill_type="solid", fgColor="DDDDDD")

    # écrire les lignes et colorer
    for row_idx, row in enumerate(dataframe_to_rows(df, index=False, header=False), start=2):
        for col_idx, val in enumerate(row, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if isinstance(val, str) and regex.search(val):
                cell.font = RED_FONT
            # pour la colonne Résumé (dernière), on transforme en lien
            if df.columns[col_idx-1] == "Résumé":
                cell.value = "Résumé"
                cell.hyperlink = row[col_idx-1]
                cell.font = LINK_FONT

    # ajuster largeur et retour à la ligne
    for col in ws.columns:
        max_length = max(len(str(c.value or "")) for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(max_length*1.1, 50)

    wb.save(stream)

if __name__ == "__main__":
    app.run(debug=True)














































































































