import re
import pandas as pd
from flask import Flask, request, render_template_string, send_file
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

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
  <div style="overflow-x:auto; margin-top:10px;">
    {{ table_html | safe }}
  </div>
{% endif %}
'''  

# lecture et filtrage
def process(df, target):
    # recensement des colonnes URL de résumé
    summary_col = next((c for c in df.columns if 'résumé' in c.lower()), None)

    # nettoyage colonnes inutile
    # on supprime toute ancienne url brute s'appelant 'resume' ou similaire
    url_cols = [c for c in df.columns if re.search(r'resum', c, re.I)]
    if summary_col:
        url_cols.remove(summary_col)
    for c in url_cols:
        df.drop(columns=c, inplace=True)

    # pattern exact article: art. 14 (word-boundary)
    pat = re.compile(rf"\b{target}\b")

    # filtrage des lignes contenant le target dans n'importe quelle colonne texte
    mask = df.apply(lambda row: any(isinstance(v, str) and pat.search(v) for v in row), axis=1)
    filtered = df[mask].copy()

    # création colonne 'Résumé' hyperlien
    if summary_col:
        filtered['Résumé'] = filtered[summary_col].apply(
            lambda u: f'<a href="{u}">Résumé</a>' if pd.notna(u) else '')

    return filtered

# génération Excel formaté
def to_excel(df):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Filtered'

    # écriture des en-têtes
    ws.append(list(df.columns))
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')

    # écriture données
    for r in dataframe_to_rows(df, index=False, header=False):
        ws.append(r)

    # style: auto width + wrap text + rouge pour target
    for col in ws.columns:
        max_len = 0
        for cell in col:
            cell.alignment = Alignment(wrapText=True)
            val = str(cell.value or '')
            max_len = max(max_len, len(val))
            # surlignage rouge si contient target
            if re.search(rf"\b{app.config['TARGET']}\b", val):
                cell.font = Font(color='FF0000')
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 50)

    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio

@app.route('/', methods=['GET', 'POST'])
def analyze():
    table_html = None
    if request.method == 'POST':
        file = request.files['file']
        target = request.form['article']
        app.config['TARGET'] = target

        df = pd.read_excel(file)
        filtered = process(df, target)
        # conversion HTML
        table_html = filtered.to_html(index=False, escape=False)
        # sauvegarde pour téléchargement
        app.config['EXCEL_BUF'] = to_excel(filtered)

    return render_template_string(HTML_TEMPLATE, table_html=table_html)

@app.route('/download')
def download():
    buf = app.config.get('EXCEL_BUF')
    return send_file(buf, download_name='decisions_filtrees.xlsx', as_attachment=True)

if __name__ == '__main__':
    app.run()















































































































