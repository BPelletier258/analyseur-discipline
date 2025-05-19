import re
import pandas as pd
from flask import Flask, request, render_template_string, send_file, redirect, url_for
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

app = Flask(__name__)

# stockage du dernier Excel généré
last_excel = None

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

# construction of strict regex pattern

def build_pattern(article):
    art = re.escape(article)
    return rf"\b{art}(?![0-9])"

# Excel styles
grey_fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
red_font = Font(color="FF0000")
border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
wrap_alignment = Alignment(wrap_text=True, vertical='top')

@app.route('/', methods=['GET','POST'])
def analyze():
    global last_excel
    if request.method == 'POST':
        file = request.files['file']
        article = request.form['article'].strip()
        df = pd.read_excel(file)
        # conserver la colonne Résumé, supprimer uniquement la colonne brute "resume"
        df.drop(columns=['resume'], inplace=True, errors='ignore')

        # repérer les cellules à colorer
        pat = build_pattern(article)
        highlights = []
        for idx, row in df.iterrows():
            for col in df.columns:
                if re.search(pat, str(row[col])):
                    highlights.append((idx, col))
        # filtrer lignes
        mask = df.apply(lambda r: any(re.search(pat, str(v)) for v in r), axis=1)
        df_filtered = df[mask].copy()

        # HTML avec colonne Résumé
        table_html = df_filtered.to_html(index=False, escape=False)

        # Excel
        output = BytesIO()
        wb = Workbook()
        ws = wb.active
        # titre article
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(df_filtered.columns))
        top = ws.cell(row=1, column=1, value=f"Article filtré : {article}")
        top.font = Font(size=14, bold=True)

        # en-têtes
        for i, col in enumerate(df_filtered.columns, start=1):
            c = ws.cell(row=2, column=i, value=col)
            c.fill = grey_fill
            c.font = Font(size=12, bold=True)
            c.border = border
            c.alignment = wrap_alignment

        # données
        for r, row in enumerate(dataframe_to_rows(df_filtered, index=False, header=False), start=3):
            for c_idx, val in enumerate(row, start=1):
                cell = ws.cell(row=r, column=c_idx, value=val)
                cell.border = border
                cell.alignment = wrap_alignment
        # coloration
        for r_idx, col in highlights:
            if mask.iloc[r_idx]:
                col_idx = df_filtered.columns.get_loc(col) + 1
                row_number = list(df_filtered.index).index(r_idx) + 3
                ws.cell(row=row_number, column=col_idx).font = red_font

        # largeur fixe
        for i in range(1, len(df_filtered.columns)+1):
            letter = get_column_letter(i)
            ws.column_dimensions[letter].width = 20

        wb.save(output)
        output.seek(0)
        last_excel = output.getvalue()
        return render_template_string(HTML_TEMPLATE, table_html=table_html, searched_article=article)
    return render_template_string(HTML_TEMPLATE)

@app.route('/download')
def download():
    global last_excel
    if not last_excel:
        return redirect(url_for('analyze'))
    return send_file(BytesIO(last_excel), as_attachment=True, download_name='decisions_filtrees.xlsx')

if __name__ == '__main__':
    app.run(debug=True)
































































































































