import re
import sys
import unicodedata
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
from flask import Flask, request, render_template_string, abort
import base64

# -----------------------------------------------------------------------------
# CONFIGURATION
# -----------------------------------------------------------------------------
output_file = "decisions_article_formate.xlsx"
app = Flask(__name__)

# HTML template with form and results
HTML = '''<!doctype html>
<title>Analyse Disciplinaire</title>
<h3>Analyse Disciplinaire</h3>
<form action="/analyze" method=post enctype=multipart/form-data>
  <p>Fichier Excel: <input type=file name=file accept=".xls,.xlsx"></p>
  <p>Article à filtrer: <input type=text name=article value="14" size=3></p>
  <p><input type=submit value="Analyser"></p>
</form>
{% if download_link and markdown %}
<hr>
<a download="{{filename}}" href="{{download_link}}">⬇️ Télécharger le fichier Excel formaté</a>
<pre>{{markdown}}</pre>
{% endif %}
'''

# normalize text

def normalize(col):
    text = unicodedata.normalize('NFKD', str(col))
    return ''.join(c for c in text if not unicodedata.combining(c)).lower().strip()

# filter logic

def filter_dataframe(stream, article_n):
    df = pd.read_excel(stream)
    df.columns = [normalize(c) for c in df.columns]
    target = rf"Art[\.:]\s*{re.escape(article_n)}(?=\D|$)"
    # rename tolerances
    if 'nom de lintime' in df.columns:
        df.rename(columns={'nom de lintime': "nom de l'intime"}, inplace=True)
    for col in df.columns:
        if 'article' in col and 'enf' in col:
            df.rename(columns={col: 'articles enfreints'}, inplace=True)
            break
    req = ['numero de decision', "nom de l'intime", 'articles enfreints']
    missing = [c for c in req if c not in df.columns]
    if missing:
        raise KeyError(f"Colonnes manquantes : {missing}")
    return df[df['articles enfreints'].astype(str).str.contains(target, regex=True)].copy(), target

# build excel in memory

def build_excel(filtered, urls, target):
    wb = Workbook()
    ws = wb.active
    ws.title = f'Article_{target}'
    for r in dataframe_to_rows(filtered, index=False, header=True):
        ws.append(r)
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')
    for col in ws.columns:
        max_len = max((len(str(c.value)) for c in col), default=0)
        ws.column_dimensions[col[0].column_letter].width = min(max_len+2, 40)
    hdr = {cell.value: i+1 for i, cell in enumerate(ws[1])}
    for i, row in enumerate(ws.iter_rows(min_row=2), start=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True)
            if isinstance(cell.value, str) and re.search(target, cell.value):
                cell.font = Font(color='FF0000')
        if 'Résumé' in hdr:
            link_cell = ws.cell(row=i, column=hdr['Résumé'])
            link_cell.hyperlink = urls[i-2]
            link_cell.style = 'Hyperlink'
    bio = BytesIO()
    wb.save(bio)
    b = bio.getvalue()
    b64 = base64.b64encode(b).decode()
    link = f"data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}"
    return link

# routes

@app.route('/')
def index():
    return render_template_string(HTML, download_link=None, markdown=None)

@app.route('/analyze', methods=['POST'])
def analyze():
    file = request.files.get('file')
    article_n = request.form.get('article','').strip()
    if not file or not article_n:
        abort(400,'Fichier ou article manquant.')
    try:
        filtered, target = filter_dataframe(file.stream, article_n)
    except Exception as e:
        abort(400, str(e))
    file.stream.seek(0)
    df_all = pd.read_excel(file.stream)
    urls = df_all.get('resume', ['']*len(df_all)).tolist()
    link = build_excel(filtered.assign(Résumé=['Résumé']*len(filtered)), urls, target)
    md = filtered.to_markdown(index=False)
    return render_template_string(HTML, download_link=link, filename=output_file, markdown=md)

if __name__=='__main__':
    app.run(host='0.0.0.0', port=5000)























































































