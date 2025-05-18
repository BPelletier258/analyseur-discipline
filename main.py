import re
import glob
import sys
import pandas as pd
from flask import Flask, request, render_template_string, send_file
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows

app = Flask(__name__)

# --- Paramètres ---
TARGET_ARTICLE = r"Art\.\s*14"

# Page d'accueil avec formulaire d'upload
INDEX_HTML = '''
<!doctype html>
<html lang="fr">
<head><meta charset="utf-8"><title>Analyse Disciplinaire</title>
<style>
  table {border-collapse: collapse; width: 100%;}
  th, td {border: 1px solid #ccc; padding: 4px; text-align: left;}
  th {background: #f0f0f0;}
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
  <a href="/download">⬇️ Télécharger le fichier Excel formaté</a><br><br>
  {{ table_html|safe }}
{% endif %}
</body>
</html>
'''

@app.route('/', methods=['GET'])
def index():
    return render_template_string(INDEX_HTML)

@app.route('/analyze', methods=['POST'])
def analyze():
    # Lecture du fichier uploadé
    f = request.files['file']
    article = request.form.get('article', '14')
    pattern = re.compile(rf"Art\.\s*{article}")
    df = pd.read_excel(f)

    # Détection colonne articles enfreints
    col = [c for c in df.columns if 'articles enfreints' in c.lower()]
    if not col:
        return 'Colonne "articles enfreints" introuvable.', 400
    col = col[0]

    # Filtrage rigoureux
    mask = df[col].astype(str).str.contains(pattern)
    filtered = df.loc[mask].copy()
    # Supprimer colonne hyperliens « Résumé »
    drop_cols = [c for c in filtered.columns if 'résumé' in c.lower()]
    filtered.drop(columns=drop_cols, inplace=True)

    # Mise en forme Excel
    wb = Workbook()
    ws = wb.active
    for r in dataframe_to_rows(filtered, index=False, header=True):
        ws.append(r)
    # Surligner en rouge les cellules contenant l'article
    red_font = Font(color='00FF0000')
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            if isinstance(cell.value, str) and pattern.search(cell.value):
                cell.font = red_font

    out = 'filtered_output.xlsx'
    wb.save(out)

    # Génération HTML table
    table_html = filtered.to_html(index=False)
    # Stocker pour envoi
    request.environ['output_file'] = out
    return render_template_string(INDEX_HTML, table_html=table_html)

@app.route('/download')
def download():
    out = request.environ.get('output_file', 'filtered_output.xlsx')
    return send_file(out, download_name=out)

if __name__ == '__main__':
    app.run(debug=True)


























































































