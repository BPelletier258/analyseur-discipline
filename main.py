import re
import pandas as pd
from flask import Flask, request, render_template_string, send_file
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

app = Flask(__name__)

HTML_TEMPLATE = '''
<!doctype html>
<html lang="fr">
<head>
  <meta charset="utf-8">
  <title>Analyse Disciplinaire</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 20px; }
    h1 { font-size: 2em; margin-bottom: 0.5em; }
    .form-container { background: #f9f9f9; padding: 15px; border-radius: 5px; max-width: 500px; }
    label { display: block; margin: 10px 0 5px; font-weight: bold; font-size: 1.1em; }
    input[type=text], input[type=file] { width: 100%; padding: 8px; font-size: 1.1em; }
    button { margin-top: 15px; padding: 10px 20px; font-size: 1.1em; }
    .article-label { margin-top: 20px; font-size: 1.2em; font-weight: bold; }
    .table-container { overflow-x: auto; margin-top: 20px; }
    table { border-collapse: collapse; width: 100%; table-layout: fixed; }
    th, td { border: 1px solid #333; padding: 8px; vertical-align: top; word-wrap: break-word; }
    th { background: #e0e0e0; font-weight: bold; }
  </style>
</head>
<body>
  <h1>Analyse Disciplinaire</h1>
  <div class="form-container">
    <form method="post" enctype="multipart/form-data">
      <label for="file">Fichier Excel</label>
      <input type="file" id="file" name="file" required>
      <label for="article">Article à filtrer</label>
      <input type="text" id="article" name="article" value="14" required>
      <button type="submit">Analyser</button>
    </form>
  </div>
  <hr>
  {% if searched_article %}
    <div class="article-label">Article recherché : {{ searched_article }}</div>
  {% endif %}
  {% if table_html %}
    <a href="/download">⬇️ Télécharger le fichier Excel formaté</a>
    <div class="table-container">
      <h2>Tableau complet</h2>
      {{ table_html|safe }}
    </div>
    {% if filtered_html %}
    <div class="table-container">
      <h2>Détails pour l'article {{ searched_article }}</h2>
      {{ filtered_html|safe }}
    </div>
    {% endif %}
  {% endif %}
</body>
</html>
'''

# regex builder to match exact article (avoid hundreds)
def build_pattern(article):
    # escape dots
    art = re.escape(article)
    # match Art. <article>([^0-9]|\b)
    return rf"Art\.\s*{art}(?![0-9])"

# style for Excel
header_fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
red_font = Font(color="FF0000")
border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

@app.route('/', methods=['GET', 'POST'])
def analyze():
    table_html = filtered_html = None
    if request.method == 'POST':
        file = request.files['file']
        article = request.form['article'].strip()
        df = pd.read_excel(file)
        # remove summary URL column if exists
        url_cols = [c for c in df.columns if 'Résumé' in c or 'résumé' in c]
        for c in url_cols:
            df.drop(columns=c, inplace=True)
        # build pattern
        pat = build_pattern(article)
        # highlight cells in Excel: record positions
        highlights = []
        for row_idx, row in df.iterrows():
            for col in df.columns:
                if isinstance(row[col], str) and re.search(pat, row[col]):
                    highlights.append((row_idx, col))
        # filtered df for HTML highlight
        mask = df.apply(lambda r: any(re.search(pat, str(v)) for v in r), axis=1)
        df_filtered = df[mask]
        # full and filtered HTML
        table_html = df.to_html(classes='', index=False, border=1)
        filtered_html = df_filtered.to_html(classes='', index=False, border=1)
        # generate Excel in memory
        output = BytesIO()
        wb = Workbook()
        ws = wb.active
        ws.title = f"Article_{article}"
        # write header
        for col_idx, col in enumerate(df.columns, 1):
            cell = ws.cell(row=1, column=col_idx, value=col)
            cell.fill = header_fill
            cell.font = Font(bold=True)
            cell.border = border
        # write data
        for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=False), start=2):
            for c_idx, val in enumerate(row, 1):
                cell = ws.cell(row=r_idx, column=c_idx, value=val)
                cell.border = border
        # apply highlights
        for (r_idx, col) in highlights:
            c_idx = list(df.columns).index(col) + 1
            cell = ws.cell(row=r_idx+2, column=c_idx)
            cell.font = red_font
        wb.save(output)
        output.seek(0)
        # store in session
        request.environ['excel_bytes'] = output.read()
        # render
        return render_template_string(HTML_TEMPLATE,
                                      table_html=table_html,
                                      filtered_html=filtered_html,
                                      searched_article=article)
    return render_template_string(HTML_TEMPLATE)

@app.route('/download')
def download():
    data = request.environ.get('excel_bytes')
    return send_file(BytesIO(data),
                     download_name="decisions_formatees.xlsx",
                     as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)




























































































































