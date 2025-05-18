import re
import sys
import pandas as pd
from flask import Flask, request, render_template_string, send_file
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils.dataframe import dataframe_to_rows

app = Flask(__name__)

# --- INSTRUCTIONS ---
# L'utilisateur uploade un fichier Excel et saisit le numéro d'article à filtrer.
# Le code génère un fichier Excel formaté et affiche un tableau HTML stylé.

INDEX_HTML = '''
<!doctype html>
<html lang="fr">
<head>
  <meta charset="utf-8">
  <title>Analyse Disciplinaire</title>
  <style>
    table {border-collapse: collapse; width: 100%; table-layout: fixed;}
    th, td {border: 1px solid #ccc; padding: 6px; text-align: left; overflow: hidden; white-space: nowrap; text-overflow: ellipsis;}
    th {background: #f5f5f5;}
    .container {max-width: 100%; overflow-x: auto;}
  </style>
</head>
<body>
  <h1>Analyse Disciplinaire</h1>
  <form action="/analyze" method="post" enctype="multipart/form-data">
    <label>Fichier Excel: <input type="file" name="file" required></label><br><br>
    <label>Article à filtrer: <input type="text" name="article" value="14" size="4"></label><br><br>
    <button type="submit">Analyser</button>
  </form>
  <hr>
  {% if table_html %}
    <a href="/download">⬇️ Télécharger le fichier Excel formaté</a>
    <div class="container">{{ table_html|safe }}</div>
  {% endif %}
</body>
</html>
'''

def make_regex(article):
    # correspondance exacte (Art. 14, pas Art.149 ou Art.140)
    return re.compile(rf"\bArt\.\s*{int(article)}(?!\d)", re.IGNORECASE)

@app.route('/', methods=['GET'])
def index():
    return render_template_string(INDEX_HTML)

@app.route('/analyze', methods=['POST'])
def analyze():
    upload = request.files.get('file')
    article = request.form.get('article', '14')
    pattern = make_regex(article)

    df = pd.read_excel(upload)
    cols = [c for c in df.columns if 'articles enfreints' in c.lower()]
    if not cols:
        return 'Colonne "articles enfreints" introuvable', 400
    col_name = cols[0]

    mask = df[col_name].astype(str).str.contains(pattern)
    filtered = df.loc[mask].copy()

    # supprime colonnes de résumé/hyperliens
    drop_cols = [c for c in filtered.columns if 'résumé' in c.lower()]
    filtered.drop(columns=drop_cols, inplace=True)

    # création Excel formaté
    wb = Workbook()
    ws = wb.active
    for row in dataframe_to_rows(filtered, index=False, header=True):
        ws.append(row)
    red_font = Font(color='FFFF0000')
    for r in ws.iter_rows(min_row=2):
        for cell in r:
            if isinstance(cell.value, str) and pattern.search(cell.value):
                cell.font = red_font

    output_path = 'filtered_output.xlsx'
    wb.save(output_path)
    app.config['LAST_OUTPUT'] = output_path

    table_html = filtered.to_html(index=False)
    return render_template_string(INDEX_HTML, table_html=table_html)

@app.route('/download')
def download():
    path = app.config.get('LAST_OUTPUT', 'filtered_output.xlsx')
    return send_file(path, download_name=path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)



























































































