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
    h1 { font-size: 1.8em; margin-bottom: 0.5em; }
    .form-container { background: #f9f9f9; padding: 20px; border-radius: 5px; max-width: 600px; }
    label { display: block; margin: 15px 0 5px; font-weight: bold; font-size: 1.2em; }
    input[type=text], input[type=file] { width: 100%; padding: 10px; font-size: 1.2em; }
    button { margin-top: 20px; padding: 12px 24px; font-size: 1.2em; }
    .article-label { margin-top: 25px; font-size: 1.3em; font-weight: bold; }
    .table-container { overflow-x: auto; margin-top: 30px; }
    table { border-collapse: collapse; width: 100%; table-layout: fixed; }
    th, td { border: 1px solid #444; padding: 10px; vertical-align: top; word-wrap: break-word; }
    th { background: #ddd; font-weight: bold; font-size: 1.1em; }
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
      {{ table_html|safe }}
    </div>
  {% endif %}
</body>
</html>
'''

# strict regex to match only exact article

def build_pattern(article):
    art = re.escape(article)
    return rf"Art\.\s*{art}(?![0-9])"

# Excel styles
header_fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
red_font = Font(color="FF0000")
border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

@app.route('/', methods=['GET','POST'])
def analyze():
    table_html = None
    if request.method == 'POST':
        file = request.files['file']
        article = request.form['article'].strip()
        df = pd.read_excel(file)
        # drop any summary URL columns
        url_cols = [c for c in df.columns if 'Résumé' in c or 'résumé' in c]
        df.drop(columns=url_cols, inplace=True, errors='ignore')
        # build pattern and find highlights
        pat = build_pattern(article)
        highlights = []
        for r_idx, row in df.iterrows():
            for col in df.columns:
                val = str(row[col])
                if re.search(pat, val):
                    highlights.append((r_idx, col))
        # filter rows
        mask = df.apply(lambda row: any(re.search(pat, str(v)) for v in row), axis=1)
        df_filtered = df[mask]
        # HTML table
        table_html = df_filtered.to_html(index=False, border=1)
        # generate Excel
        output = BytesIO()
        wb = Workbook()
        ws = wb.active
        # header
        for c_idx, col in enumerate(df.columns, start=1):
            cell = ws.cell(row=1, column=c_idx, value=col)
            cell.fill = header_fill
            cell.font = Font(bold=True)
            cell.border = border
        # data
        for r, row in enumerate(dataframe_to_rows(df, index=False, header=False), start=2):
            for c, val in enumerate(row, start=1):
                cell = ws.cell(row=r, column=c, value=val)
                cell.border = border
        # apply highlights
        for r_idx, col in highlights:
            c_idx = df.columns.get_loc(col) + 1
            cell = ws.cell(row=r_idx+2, column=c_idx)
            cell.font = red_font
        wb.save(output)
        output.seek(0)
        request.environ['excel_bytes'] = output.read()
        return render_template_string(HTML_TEMPLATE, table_html=table_html, searched_article=article)
    return render_template_string(HTML_TEMPLATE)

@app.route('/download')
def download():
    data = request.environ.get('excel_bytes')
    return send_file(BytesIO(data), as_attachment=True, download_name='decisions_filtrees.xlsx')

if __name__ == '__main__':
    app.run(debug=True)





























































































































