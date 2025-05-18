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

# -----------------------------------------
# CONFIGURATION
# -----------------------------------------
output_file = "decisions_article_formate.xlsx"
app = Flask(__name__)

# HTML form and result template
template = '''<!doctype html>
<title>Analyse Disciplinaire</title>
<h3>Analyse Disciplinaire</h3>
<form action="/analyze" method="post" enctype="multipart/form-data">
  <p>Fichier Excel: <input type="file" name="file" accept=".xls,.xlsx"></p>
  <p>Article à filtrer: <input type="text" name="article" value="14" size="3"></p>
  <p><input type="submit" value="Analyser"></p>
</form>
{% if download_link and markdown %}
<hr>
<a download="{{filename}}" href="{{download_link}}">⬇️ Télécharger le fichier Excel formaté</a>
<pre>{{markdown}}</pre>
{% endif %}
'''

# normalize strings for matching
def normalize(s):
    s = unicodedata.normalize('NFKD', str(s))
    return ''.join(c for c in s if not unicodedata.combining(c)).lower().strip()

# filter dataframe by article

def filter_df(stream, article):
    df = pd.read_excel(stream)
    # normalize headers
    namemap = {col: normalize(col) for col in df.columns}
    df.rename(columns=namemap, inplace=True)
    # remap specific keys
    if 'nom de lintime' in df.columns:
        df.rename(columns={'nom de lintime': "nom de l'intime"}, inplace=True)
    # find articles column
    if 'articles enfreints' not in df.columns:
        for col in df.columns:
            if 'article' in col and 'enf' in col:
                df.rename(columns={col: 'articles enfreints'}, inplace=True)
                break
    required = ['numero de decision', "nom de l'intime", 'articles enfreints']
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise KeyError(f"Colonnes manquantes: {missing}")
    # build pattern
    pattern = rf"Art[\.:]\s*{re.escape(article)}(?=\D|$)"
    mask = df['articles enfreints'].astype(str).str.contains(pattern, regex=True)
    return df[mask].copy(), df, pattern

# create excel and highlight

def build_excel(filtered, full_df, article, pattern):
    # locate URL column
    urlcol = None
    for col in full_df.columns:
        if normalize(col) == 'resume' or normalize(col) == 'résumé':
            urlcol = col
            break
    links = full_df[urlcol].fillna('') if urlcol else pd.Series([''] * len(filtered))

    wb = Workbook()
    ws = wb.active
    ws.title = f"Article_{article}"
    # add header and data
    filtered['Résumé'] = 'Résumé'
    for row in dataframe_to_rows(filtered, index=False, header=True):
        ws.append(row)
    # style header
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')
    # auto column widths
    for col in ws.columns:
        maxlen = max(len(str(c.value)) for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(maxlen+2, 40)
    # content styling
    hdr = {c.value: i+1 for i,c in enumerate(ws[1])}
    for i, row in enumerate(ws.iter_rows(min_row=2), start=0):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True)
            if isinstance(cell.value, str) and re.search(pattern, cell.value):
                cell.font = Font(color='FF0000')
        # hyperlink
        if 'Résumé' in hdr:
            col_idx = hdr['Résumé']
            link_cell = ws.cell(row=i+2, column=col_idx)
            link_cell.value = 'Résumé'
            if links.iat[i]:
                link_cell.hyperlink = links.iat[i]
            link_cell.font = Font(color='0000FF', underline='single')
    # output
    bio = BytesIO()
    wb.save(bio)
    return base64.b64encode(bio.getvalue()).decode()

# routes

@app.route('/')
def index():
    return render_template_string(template)

@app.route('/analyze', methods=['POST'])
def analyze():
    f = request.files.get('file')
    art = request.form.get('article','').strip()
    if not f or not art:
        abort(400, 'Fichier ou article manquant.')
    try:
        filtered, full, pat = filter_df(f.stream, art)
    except Exception as e:
        abort(400, str(e))
    f.stream.seek(0)
    # build excel
    b64 = build_excel(filtered, full, art, pat)
    link = f"data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}"
    md = filtered.to_markdown(index=False)
    return render_template_string(template, download_link=link, filename=output_file, markdown=md)

if __name__=='__main__':
    app.run(debug=True)

























































































