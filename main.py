import re
import sys
import unicodedata
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
from flask import Flask, request, send_file, render_template_string, abort

# -----------------------------------------------------------------------------
# CONFIGURATION
# -----------------------------------------------------------------------------
output_file = "decisions_article_formate.xlsx"
app = Flask(__name__)

# simple HTML template with upload form and article input
INDEX_HTML = '''
<!doctype html>
<title>Analyse Disciplinaire</title>
<h3>Analyse Disciplinaire</h3>
<form action="/analyze" method=post enctype=multipart/form-data>
  <p>Fichier Excel: <input type=file name=file accept=".xls,.xlsx"></p>
  <p>Article à filtrer: <input type=text name=article value="14" size=3></p>
  <p><input type=submit value="Analyser"></p>
</form>
'''  

# -----------------------------------------------------------------------------
# UTILS
# -----------------------------------------------------------------------------
def normalize(col):
    text = unicodedata.normalize('NFKD', str(col))
    return ''.join(c for c in text if not unicodedata.combining(c)).lower().strip()

# -----------------------------------------------------------------------------
# EXTRACTION & FILTRAGE
# -----------------------------------------------------------------------------
def filter_dataframe(stream, article_n):
    df = pd.read_excel(stream)
    df.columns = [normalize(c) for c in df.columns]
    # dynamic article regex
    target = rf"Art[\.:]\s*{re.escape(article_n)}(?=\D|$)"
    # rename columns tolerantly
    if 'nom de lintime' in df.columns:
        df.rename(columns={'nom de lintime': "nom de l'intime"}, inplace=True)
    for col in df.columns:
        if 'article' in col and 'enf' in col:
            df.rename(columns={col: 'articles enfreints'}, inplace=True)
            break
    required = ['numero de decision', "nom de l'intime", 'articles enfreints']
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise KeyError(f"Colonnes manquantes : {missing}")
    mask = df['articles enfreints'].astype(str).str.contains(target, regex=True)
    return df.loc[mask].copy(), target

# -----------------------------------------------------------------------------
# BUILD EXCEL
# -----------------------------------------------------------------------------
def build_excel(filtered, urls, target):
    wb = Workbook()
    ws = wb.active
    ws.title = f'Article_{target}'
    for r in dataframe_to_rows(filtered, index=False, header=True):
        ws.append(r)
    # style header
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')
    # adjust columns
    for col in ws.columns:
        max_len = max((len(str(c.value)) for c in col), default=0)
        ws.column_dimensions[col[0].column_letter].width = min(max_len+2, 40)
    # color cells and hyperlinks
    header_index = {cell.value: idx+1 for idx, cell in enumerate(ws[1])}
    for i, row in enumerate(ws.iter_rows(min_row=2), start=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True)
            if isinstance(cell.value, str) and re.search(target, cell.value):
                cell.font = Font(color='FF0000')
        if 'Résumé' in header_index:
            link_cell = ws.cell(row=i, column=header_index['Résumé'])
            link_cell.hyperlink = urls[i-2]
            link_cell.style = 'Hyperlink'
    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio

# -----------------------------------------------------------------------------
# ROUTES
# -----------------------------------------------------------------------------
@app.route('/')
def index():
    return render_template_string(INDEX_HTML)

@app.route('/analyze', methods=['POST'])
def analyze():
    file = request.files.get('file')
    article_n = request.form.get('article', '').strip()
    if not file or not article_n:
        abort(400, 'Fichier ou article manquant.')
    try:
        filtered, target = filter_dataframe(file.stream, article_n)
    except Exception as e:
        abort(400, str(e))
    # re-read full DF for URLs
    file.stream.seek(0)
    df_all = pd.read_excel(file.stream)
    urls = df_all.get('resume', ['']*len(df_all)).tolist()
    # build excel
    bio = build_excel(filtered.assign(Résumé=['Résumé']*len(filtered)), urls, target)
    md = filtered.to_markdown(index=False)
    return send_file(bio, as_attachment=True, download_name=output_file), '<pre>'+md+'</pre>'

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)






















































































