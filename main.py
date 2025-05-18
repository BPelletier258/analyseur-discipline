import re
import sys
import pandas as pd
from flask import Flask, request, render_template_string, send_file
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
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
    .container { overflow-x: auto; }
    table { border-collapse: collapse; width: 100%; table-layout: auto; }
    th, td { border: 1px solid #ccc; padding: 6px; text-align: left; white-space: normal; word-break: break-word; }
    th { background: #f0f0f0; }
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

# Crée une regex stricte pour Art. XX sans capter XXX
def make_regex(article):
    art = int(article)
    return re.compile(rf"\bArt\.\s*{art}(?!\d)", re.IGNORECASE)

@app.route('/', methods=['GET'])
def index():
    return render_template_string(INDEX_HTML)

@app.route('/analyze', methods=['POST'])
def analyze():
    upload = request.files.get('file')
    article = request.form.get('article', '14')
    pattern = make_regex(article)

    df = pd.read_excel(upload)

    # filtre pour toute cellule contenant le bon article
    mask = df.applymap(lambda v: bool(pattern.search(str(v))))
    row_mask = mask.any(axis=1)
    filtered = df.loc[row_mask].copy()

    # supprime colonne résumé
    drop_cols = [c for c in filtered.columns if 'résumé' in c.lower()]
    filtered.drop(columns=drop_cols, inplace=True)

    # format Excel
    wb = Workbook()
    ws = wb.active
    # entêtes
    for r_idx, row in enumerate(dataframe_to_rows(filtered, index=False, header=True), start=1):
        ws.append(row)
        if r_idx == 1:
            # header bold
            for cell in ws[r_idx]:
                cell.font = Font(bold=True)
    # style wrap & rouge
    red = Font(color='FFFF0000')
    for col in ws.columns:
        max_length = 10
        for cell in col:
            cell.alignment = Alignment(wrap_text=True)
            val = str(cell.value or '')
            if pattern.search(val):
                cell.font = red
            max_length = max(max_length, len(val))
        # colonne automatique
        ws.column_dimensions[col[0].column_letter].width = min(max_length * 1.1, 50)

    out = 'filtered_output.xlsx'
    wb.save(out)
    app.config['LAST_FILE'] = out

    # HTML table avec classes
    html = filtered.to_html(index=False, classes='discipline')
    return render_template_string(INDEX_HTML, table_html=html)

@app.route('/download')
def download():
    return send_file(app.config.get('LAST_FILE'),
                     download_name='decisions_filtrees.xlsx',
                     as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)































































































