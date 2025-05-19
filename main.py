import re
import pandas as pd
from flask import Flask, request, render_template_string, send_file
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows

app = Flask(__name__)

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
  <div style="overflow-x:auto; margin-top:10px; width:100%;">
    {{ table_html | safe }}
  </div>
{% endif %}
'''  

def process(df, target):
    # identifier colonne résumé brute
    summary_col = next((c for c in df.columns if 'résumé' in c.lower()), None)
    # supprimer toutes colonnes raw contenant 'resume'
    drop_cols = [c for c in df.columns if re.search(r'resum', c, re.I) and c != summary_col]
    df.drop(columns=drop_cols, inplace=True)

    # pattern exact Article n (éviter 149, 140, etc.)
    pat = re.compile(rf"Article\s+{target}\b", re.I)

    # filtrer lignes où l'article apparaît
    mask = df.apply(lambda row: any(isinstance(v, str) and pat.search(v) for v in row), axis=1)
    filtered = df[mask].copy()

    # colonne Résumé hyperlien
    if summary_col:
        filtered['Résumé'] = filtered[summary_col].fillna('').apply(
            lambda u: f'<a href="{u}">Résumé</a>' if u else '')
    else:
        filtered['Résumé'] = ''

    return filtered


def to_excel(df):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Décisions'

    # écrire en-têtes
    ws.append(list(df.columns))
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center', wrapText=True)

    # écrire données
    for r in dataframe_to_rows(df, index=False, header=False):
        ws.append(r)

    # style colonnes: largeur fixe, retour à la ligne, rouge si article présent
    target = app.config.get('TARGET', '')
    pat_cell = re.compile(rf"Article\s+{target}\b", re.I)
    for col in ws.columns:
        max_len = 0
        for cell in col:
            cell.alignment = Alignment(wrapText=True)
            text = str(cell.value or '')
            max_len = max(max_len, len(text))
            if pat_cell.search(text):
                cell.font = Font(color='FF0000')
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
        # HTML
        table_html = filtered.to_html(index=False, escape=False)
        # préparer download
        app.config['EXCEL_BUF'] = to_excel(filtered)

    return render_template_string(HTML_TEMPLATE, table_html=table_html)

@app.route('/download')
def download():
    buf = app.config.get('EXCEL_BUF')
    return send_file(buf, download_name='decisions_filtrees.xlsx', as_attachment=True)

if __name__ == '__main__':
    app.run()

















































































































