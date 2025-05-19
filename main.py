import re
import pandas as pd
from flask import Flask, request, render_template_string, send_file
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows

app = Flask(__name__)

# HTML template with horizontal scrolling
HTML_TEMPLATE = '''
<!doctype html>
<title>Analyse Disciplinaire</title>
<h1>Analyse Disciplinaire</h1>
<form method="post" enctype="multipart/form-data">
  Fichier Excel: <input type="file" name="file"><br>
  Article à filtrer: <input type="text" name="article" value="14"><br>
  <button type="submit">Analyser</button>
</form>
<hr>
{% if table_html %}
  <a href="/download">⬇️ Télécharger le fichier Excel formaté</a>
  <div style="overflow-x:auto; white-space:nowrap; margin-top:10px;">
    {{ table_html | safe }}
  </div>
{% endif %}
'''

def process(df, target):
    # drop any column containing raw URLs (http/https)
    url_cols = [c for c in df.columns if df[c].astype(str).str.match(r'https?://').any()]
    df = df.drop(columns=url_cols)

    # find existing résumé column (by containing 'résumé' or 'resume')
    summary_col = next((c for c in df.columns if re.search(r'r[eé]sum', c, re.I)), None)
    # drop any other columns with 'resum' in name
    drop_cols = [c for c in df.columns if re.search(r'r[eé]sum', c, re.I) and c != summary_col]
    df = df.drop(columns=drop_cols)

    # compile regex for Art/Article target as distinct number (avoid 3-digit)
    pat = re.compile(rf"\b(?:Art(?:icle)?\.?\s*){target}(?!\d)\b", re.I)

    # filter rows where any cell contains the exact target article
    mask = df.apply(lambda row: any(isinstance(v, str) and pat.search(v) for v in row), axis=1)
    filtered = df[mask].copy()

    # add clean Résumé column with link
    if summary_col:
        filtered['Résumé'] = filtered[summary_col].fillna('').apply(
            lambda u: f'<a href="{u}">Résumé</a>' if u.startswith('http') else '')
        filtered = filtered.drop(columns=[summary_col])
    else:
        filtered['Résumé'] = ''

    return filtered


def to_excel(df):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Décisions'

    # write headers
    ws.append(list(df.columns))
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center', wrapText=True)

    # write data rows
    for r in dataframe_to_rows(df, index=False, header=False):
        ws.append(r)

    # coloring and sizing
    target = app.config.get('TARGET', '')
    pat_cell = re.compile(rf"\b(?:Art(?:icle)?\.?\s*){target}(?!\d)\b", re.I)

    for col in ws.columns:
        max_len = 0
        for cell in col:
            text = str(cell.value or '')
            max_len = max(max_len, len(text))
            cell.alignment = Alignment(wrapText=True)
            if pat_cell.search(text):
                cell.font = Font(color='FF0000')
        # fixed width with cap
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 40)

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


@app.route('/', methods=['GET', 'POST'])
def analyze():
    table_html = None
    if request.method == 'POST':
        file = request.files['file']
        target = request.form['article'].strip()
        app.config['TARGET'] = target

        df = pd.read_excel(file)
        filtered = process(df, target)

        # HTML table (no escape to allow links)
        table_html = filtered.to_html(index=False, escape=False)
        # prepare Excel download buffer
        app.config['EXCEL_BUF'] = to_excel(filtered)

    return render_template_string(HTML_TEMPLATE, table_html=table_html)


@app.route('/download')
def download():
    buf = app.config.get('EXCEL_BUF')
    return send_file(buf, download_name='decisions_filtrees.xlsx', as_attachment=True)


if __name__ == '__main__':
    app.run()



















































































































