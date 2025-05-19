import re
import pandas as pd
from flask import Flask, request, render_template_string, send_file
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows

app = Flask(__name__)

# HTML template with horizontal scrolling, styled form, display of searched article, and fixed header styles
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
    .table-container { overflow-x: auto; margin-top: 20px; }
    table { border-collapse: collapse; width: 100%; table-layout: fixed; }
    th, td { border: 1px solid #333; padding: 8px; vertical-align: top; word-wrap: break-word; }
    th { background: #e0e0e0; font-weight: bold; }
    .summary-col, .summary-col th, .summary-col td { text-align: center; }
    .article-label { margin-top: 20px; font-size: 1.2em; font-weight: bold; }
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

def process(df, target):
    # identify and extract any URL columns (résumé links)
    url_col = next((c for c in df.columns if re.search(r'r[eé]sum', c, re.I)), None)
    urls = df[url_col] if url_col else pd.Series(['']*len(df), index=df.index)
    # drop all URL columns to avoid raw links in output
    all_urls = [c for c in df.columns if df[c].astype(str).str.startswith('http').any()]
    df = df.drop(columns=all_urls, errors='ignore')

    # strict regex: match "Art:" or "Art." or "Article" plus exact number, not prefixes of larger numbers
    t = re.escape(target)
    pat = re.compile(rf"\bArt(?:icle)?[\.:]?\s*{t}(?![0-9])", re.I)

    # filter rows where any cell contains the target pattern
    mask = df.apply(lambda row: any(isinstance(v, str) and pat.search(v) for v in row), axis=1)
    filtered = df[mask].copy()

    # attach the original URL values for later hyperlink
    filtered['_url'] = urls[filtered.index]
    # prepare HTML "Résumé" link
    filtered['Résumé'] = filtered['_url'].apply(lambda v: f'<a class="summary-col" href="{v}" target="_blank">Résumé</a>' if isinstance(v, str) and v.startswith('http') else '')
    return filtered


def to_excel(df, target):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Décisions'
    # headers
    headers = [c for c in df.columns if c!='_url']
    ws.append(headers)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center', wrapText=True)
    # fill rows and hyperlinks
    for r_idx, row in enumerate(dataframe_to_rows(df[headers], index=False, header=False), start=2):
        for c_idx, val in enumerate(row, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.alignment = Alignment(wrapText=True)
        url = df['_url'].iloc[r_idx-2]
        hl_cell = ws.cell(row=r_idx, column=len(headers), value='Résumé')
        if isinstance(url, str) and url.startswith('http'):
            hl_cell.hyperlink = url
            hl_cell.style = 'Hyperlink'
    # highlight all cells matching the pattern in red
    pat = re.compile(rf"\bArt(?:icle)?[\.:]?{re.escape(target)}(?![0-9])", re.I)
    for col in ws.columns:
        max_len = 0
        for cell in col:
            txt = str(cell.value or '')
            max_len = max(max_len, len(txt))
            if cell.row>1 and pat.search(txt):
                cell.font = Font(color='FF0000')
        col_letter = col[0].column_letter
        ws.column_dimensions[col_letter].width = min(max_len + 4, 40)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

app.config['EXCEL_BUF'] = None

@app.route('/', methods=['GET','POST'])
def analyze():
    table_html = None
    searched = None
    if request.method=='POST':
        searched = request.form['article'].strip()
        df = pd.read_excel(request.files['file'])
        filtered = process(df, searched)
        html_df = filtered.drop(columns=['_url'])
        table_html = html_df.to_html(index=False, escape=False)
        app.config['EXCEL_BUF'] = to_excel(filtered, searched)
    return render_template_string(HTML_TEMPLATE,
                                  table_html=table_html,
                                  searched_article=searched)

@app.route('/download')
def download():
    return send_file(app.config['EXCEL_BUF'], download_name='decisions_filtrees.xlsx', as_attachment=True)

if __name__=='__main__':
    app.run()























































































































