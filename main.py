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

# normalize text/key for comparison

def normalize(text):
    t = unicodedata.normalize('NFKD', str(text))
    t = ''.join(c for c in t if not unicodedata.combining(c))
    return t.lower().strip()

# filter logic

def filter_dataframe(stream, article_n):
    df = pd.read_excel(stream)
    # map normalized headers
    col_map = {c: normalize(c) for c in df.columns}
    df.rename(columns={orig: normalized for orig, normalized in col_map.items()}, inplace=True)
    # rename key columns
    if 'nom de lintime' in df.columns:
        df.rename(columns={'nom de lintime': "nom de l'intime"}, inplace=True)
    if 'articles enfreints' not in df.columns:
        for c in df.columns:
            if 'article' in c and 'enf' in c:
                df.rename(columns={c: 'articles enfreints'}, inplace=True)
                break
    # ensure required
    req = ['numero de decision', "nom de l'intime", 'articles enfreints']
    missing = [c for c in req if c not in df.columns]
    if missing:
        raise KeyError(f"Colonnes manquantes : {missing}")
    # build regex
    target = rf"Art[\.:]\s*{re.escape(article_n)}(?=\D|$)"
    filtered = df[df['articles enfreints'].astype(str).str.contains(target, regex=True)].copy()
    return filtered, target, df

# build excel in memory

def build_excel(filtered, df_all, target):
    # prepare URLs list from normalized 'resume' or raw 'Résumé'
    url_col = None
    for c in df_all.columns:
        if normalize(c) == 'resume' or c.lower() == 'résumé':
            url_col = c
            break
    urls = df_all[url_col].fillna('').tolist() if url_col else ['']*len(filtered)

    wb = Workbook()
    ws = wb.active
    ws.title = f'Article_{article_n}'
    # add header & rows
    for r in dataframe_to_rows(filtered.assign(Résumé=['Résumé']*len(filtered)), index=False, header=True):
        ws.append(r)
    # style header
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')
    # auto-width
    for col in ws.columns:
        length = max(len(str(c.value)) for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(length+2, 40)
    hdr = {cell.value: i+1 for i, cell in enumerate(ws[1])}
    # style cells
    for i, row in enumerate(ws.iter_rows(min_row=2), start=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True)
            if isinstance(cell.value, str) and re.search(target, cell.value):
                cell.font = Font(color='FF0000')
        # add hyperlink
        if 'Résumé' in hdr and urls and i-2 < len(urls):
            c = ws.cell(row=i, column=hdr['Résumé'])
            c.hyperlink = urls[i-2]
            c.style = 'Hyperlink'
    # export to base64
    bio = BytesIO()
    wb.save(bio)
    b64 = base64.b64encode(bio.getvalue()).decode()
    return f"data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}"

# routes

@app.route('/')
def index():
    return render_template_string(HTML, download_link=None, markdown=None)

@app.route('/analyze', methods=['POST'])
def analyze():
    f = request.files.get('file')
    article_n = request.form.get('article','').strip()
    if not f or not article_n:
        abort(400, 'Fichier ou article manquant.')
    try:
        filtered, target, df_all = filter_dataframe(f.stream, article_n)
    except Exception as e:
        abort(400, str(e))
    # reset stream to read full excel
    f.stream.seek(0)
    try:
        df_full = pd.read_excel(f.stream)
    except Exception:
        df_full = df_all
    link = build_excel(filtered, df_full, target)
    md = filtered.to_markdown(index=False)
    return render_template_string(HTML, download_link=link, filename=output_file, markdown=md)

if __name__=='__main__':
    app.run(host='0.0.0.0', port=5000)
























































































