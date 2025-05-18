
import re
import sys
import pandas as pd
from flask import Flask, request, render_template_string, send_file
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils.dataframe import dataframe_to_rows

app = Flask(__name__)

# --- Page HTML ---
INDEX_HTML = '''
<!doctype html>
<html lang="fr">
<head>
  <meta charset="utf-8">
  <title>Analyse Disciplinaire</title>
  <style>
    body { font-family: Arial, sans-serif; padding: 20px; }
    form { margin-bottom: 20px; }
    table { border-collapse: collapse; width: 100%; table-layout: auto; }
    th, td { border: 1px solid #ccc; padding: 8px; text-align: left; white-space: normal; word-break: break-word; }
    th { background: #f0f0f0; }
    .container { overflow-x: auto; }
    a.download { display: inline-block; margin: 10px 0; }
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
    <a href="/download" class="download">⬇️ Télécharger le fichier Excel formaté</a>
    <div class="container">{{ table_html|safe }}</div>
  {% endif %}
</body>
</html>
'''

# Crée une regex stricte pour Art. XX (pas XXX)
def make_regex(article):
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
    # repère la colonne 'articles enfreints'
    cols = [c for c in df.columns if 'articles enfreints' in c.lower()]
    if not cols:
        return 'Colonne "articles enfreints" introuvable', 400
    col = cols[0]

    # filtre strict
    mask = df[col].astype(str).str.contains(pattern)
    filtered = df.loc[mask].copy()

    # supprime toute colonne 'résumé'
    drop = [c for c in filtered.columns if 'résumé' in c.lower()]
    filtered.drop(columns=drop, inplace=True)

    # génère Excel formaté
    wb = Workbook()
    ws = wb.active
    for row in dataframe_to_rows(filtered, index=False, header=True):
        ws.append(row)
    red = Font(color='FFFF0000')
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            if isinstance(cell.value, str) and pattern.search(cell.value):
                cell.font = red

    out = 'filtered_output.xlsx'
    wb.save(out)
    app.config['LAST'] = out

    # HTML
    html = filtered.to_html(index=False)
    return render_template_string(INDEX_HTML, table_html=html)

@app.route('/download')
def download():
    return send_file(app.config['LAST'], download_name='decisions_filtrees.xlsx', as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)






























































































