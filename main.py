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
    table.discipline { border-collapse: collapse; width: 100%; min-width: 800px; table-layout: fixed; }
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
    <div class="container">
      {{ table_html|safe }}
    </div>
  {% endif %}
</body>
</html>
'''

# Regex strict : Art. XX ou Art: XX sans capter Art XXX
def make_regex(article):
    art = int(article)
    return re.compile(rf"\bArt[\.:]\s*{art}(?!\d)", re.IGNORECASE)

@app.route('/', methods=['GET'])
def index():
    return render_template_string(INDEX_HTML)

@app.route('/analyze', methods=['POST'])
def analyze():
    upload = request.files.get('file')
    article = request.form.get('article', '14')
    pattern = make_regex(article)

    # Lecture du fichier uploadé
    df = pd.read_excel(upload)

    # Détection et insertion de la colonne Résumé à la fin
    summary_cols = [c for c in df.columns if 'résumé' in c.lower() or 'resume' in c.lower()]
    if summary_cols:
        url_col = summary_cols[0]
        df['Résumé'] = df[url_col].apply(
            lambda url: f'<a href="{url}" target="_blank">Résumé</a>' if pd.notna(url) else ''
        )
        df.drop(columns=summary_cols, inplace=True)

    # Filtrer les lignes contenant l'article exact dans n'importe quelle colonne
    mask = df.applymap(lambda v: bool(pattern.search(str(v))))
    filtered = df.loc[mask.any(axis=1)].copy()

    # --- Génération du fichier Excel formaté ---
    wb = Workbook()
    ws = wb.active
    # Insérer toutes les lignes (en-tête + données)
    for r_idx, row in enumerate(dataframe_to_rows(filtered, index=False, header=True), start=1):
        ws.append(row)
        if r_idx == 1:
            for cell in ws[r_idx]:
                cell.font = Font(bold=True)
    # Style et surlignage rouge partout où l'article apparaît
    red = Font(color='FFFF0000')
    for col in ws.columns:
        max_len = 0
        for cell in col:
            cell.alignment = Alignment(wrap_text=True)
            text = str(cell.value or '')
            if pattern.search(text):
                cell.font = red
            max_len = max(max_len, len(text))
        ws.column_dimensions[col[0].column_letter].width = min(max_len * 1.1, 50)

    out_file = 'filtered_output.xlsx'
    wb.save(out_file)
    app.config['LAST_FILE'] = out_file

    # --- Génération HTML avec hyperlinks et affichage scrollable ---
    html = filtered.to_html(index=False, classes='discipline', escape=False)
    return render_template_string(INDEX_HTML, table_html=html)

@app.route('/download')
def download():
    return send_file(app.config.get('LAST_FILE'), download_name='decisions_filtrees.xlsx', as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)


































































































