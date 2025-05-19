import re
import pandas as pd
from flask import Flask, request, render_template_string, send_file
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows

app = Flask(__name__)

# HTML template with horizontal scrolling and fixed header styles
HTML_TEMPLATE = '''
<!doctype html>
<title>Analyse Disciplinaire</title>
<style>
  .table-container { overflow-x: auto; margin-top: 10px; }
  table { border-collapse: collapse; width: 100%; table-layout: fixed; }
  th, td { border: 1px solid #333; padding: 8px; vertical-align: top; word-wrap: break-word; }
  th { background: #f0f0f0; font-weight: bold; }
  .summary-col { white-space: nowrap; }
</style>
<h1>Analyse Disciplinaire</h1>
<form method="post" enctype="multipart/form-data">
  Fichier Excel: <input type="file" name="file" required><br>
  Article à filtrer: <input type="text" name="article" value="14" required><br>
  <button type="submit">Analyser</button>
</form>
<hr>
{% if table_html %}
  <a href="/download">⬇️ Télécharger le fichier Excel formaté</a>
  <div class="table-container">
    {{ table_html|safe }}
  </div>
{% endif %}
'''

def process(df, target):
    # drop any URL columns
    url_cols = [c for c in df.columns if df[c].astype(str).str.match(r'https?://').any()]
    df = df.drop(columns=url_cols, errors='ignore')

    # strict regex for target article (not matching hundreds)
    pat = re.compile(rf"\b(?:Art(?:icle)?\.?\s*){target}(?!\d)\b", re.I)

    # filter rows containing target anywhere
    mask = df.apply(lambda row: any(isinstance(v, str) and pat.search(v) for v in row), axis=1)
    filtered = df[mask].copy()

    # build Résumé column with hyperlink
    summary_col = next((c for c in df.columns if re.search(r'r[eé]sum', c, re.I)), None)
    def make_link(url):
        return f'<a href="{url}" target="_blank">Résumé</a>' if isinstance(url, str) and url.startswith('http') else ''
    if summary_col:
        links = filtered.pop(summary_col).apply(make_link)
    else:
        links = pd.Series([''] * len(filtered), index=filtered.index)
    filtered['Résumé'] = links

    return filtered


def to_excel(df, target):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Décisions'

    # write headers
    ws.append(list(df.columns))
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center', wrapText=True)

    # write data
    for row in dataframe_to_rows(df, index=False, header=False):
        ws.append(row)

    # style and auto-width
    pat = re.compile(rf"\b(?:Art(?:icle)?\.?\s*){target}(?!\d)\b", re.I)
    for col in ws.columns:
        max_len = 0
        for cell in col:
            text = str(cell.value or '')
            max_len = max(max_len, len(text))
            cell.alignment = Alignment(wrapText=True)
            if pat.search(text):
                cell.font = Font(color='FF0000')
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 40)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

app.config['EXCEL_BUF'] = None

@app.route('/', methods=['GET', 'POST'])
def analyze():
    table_html = None
    if request.method == 'POST':
        uploaded = request.files['file']
        target = request.form['article'].strip()
        df = pd.read_excel(uploaded)
        filtered = process(df, target)
        # HTML
        table_html = filtered.to_html(index=False, escape=False)
        # Excel
        app.config['EXCEL_BUF'] = to_excel(filtered, target)
    return render_template_string(HTML_TEMPLATE, table_html=table_html)

@app.route('/download')
def download():
    buf = app.config.get('EXCEL_BUF')
    return send_file(buf, download_name='decisions_filtrees.xlsx', as_attachment=True)

if __name__ == '__main__':
    app.run()




















































































































