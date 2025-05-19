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
  .summary-col, .summary-col th, .summary-col td { text-align: center; }
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
    # find URL column for résumé
    url_col = next((c for c in df.columns if re.search(r'r[eé]sum', c, re.I)), None)
    urls = df[url_col] if url_col else pd.Series(['']*len(df), index=df.index)
    # drop all URL columns
    all_urls = [c for c in df.columns if df[c].astype(str).str.startswith('http').any()]
    df = df.drop(columns=all_urls, errors='ignore')

    # strict regex allowing Art: or Art.
    t = re.escape(target)
    pat = re.compile(rf"\bArt(?:icle)?[\.:]?\s*{t}(?![\d])", re.I)

    # filter rows containing target in any text cell
    mask = df.apply(lambda row: any(isinstance(v, str) and pat.search(v) for v in row), axis=1)
    filtered = df[mask].copy()

    # attach raw URLs for Excel hyperlink
    filtered['_url'] = urls[filtered.index]
    # build HTML link column
    def mk(v): return f'<a class="summary-col" href="{v}" target="_blank">Résumé</a>' if isinstance(v,str) else ''
    filtered['Résumé'] = filtered['_url'].apply(mk)
    return filtered


def to_excel(df, target):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Décisions'
    # write headers
    headers = [c for c in df.columns if c!='_url']
    ws.append(headers)
    for cell in ws[1]: cell.font, cell.alignment = Font(bold=True), Alignment(horizontal='center', wrapText=True)
    # write rows, set hyperlinks
    for r_idx, row in enumerate(dataframe_to_rows(df[headers], index=False, header=False), start=2):
        for c_idx, val in enumerate(row, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=val)
            cell.alignment = Alignment(wrapText=True)
        # hyperlink in last column Résumé
        url = df['_url'].iloc[r_idx-2]
        hl_cell = ws.cell(row=r_idx, column=len(headers), value='Résumé')
        if url:
            hl_cell.hyperlink, hl_cell.style = url, 'Hyperlink'
    # color cells matching pattern
    t = re.escape(target); pat = re.compile(rf"\bArt(?:icle)?[\.:]?\s*{t}(?![\d])", re.I)
    for col in ws.columns:
        max_len = 0
        for cell in col:
            txt = str(cell.value or '')
            max_len = max(max_len, len(txt))
            if cell.row>1 and pat.search(txt):
                cell.font = Font(color='FF0000')
        ws.column_dimensions[col[0].column_letter].width = min(max_len+4, 40)
    buf = BytesIO(); wb.save(buf); buf.seek(0)
    return buf

app.config['EXCEL_BUF'] = None

@app.route('/', methods=['GET','POST'])
def analyze():
    table_html = None
    if request.method=='POST':
        tgt = request.form['article'].strip()
        df = pd.read_excel(request.files['file'])
        filtered = process(df, tgt)
        # HTML: display all except raw _url
        html_df = filtered.drop(columns=['_url'])
        table_html = html_df.to_html(index=False, escape=False)
        app.config['EXCEL_BUF'] = to_excel(filtered, tgt)
    return render_template_string(HTML_TEMPLATE, table_html=table_html)

@app.route('/download')
def download():
    return send_file(app.config['EXCEL_BUF'], download_name='decisions_filtrees.xlsx', as_attachment=True)

if __name__=='__main__':
    app.run()






















































































































