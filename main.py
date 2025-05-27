
import re
import pandas as pd
from flask import Flask, request, render_template_string, send_file, redirect, url_for
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
last_excel = None
last_article = None

HTML_TEMPLATE = '''
<!doctype html>
<html lang="fr">
<head>
  <meta charset="utf-8">
  <title>Analyse Disciplinaire</title>
  <style>
    body { font-family: Arial, sans-serif; margin: 20px; }
    h1 { font-size: 1.8em; margin-bottom: 0.5em; }
    .form-container { background: #f9f9f9; padding: 20px; border-radius: 5px; max-width: 600px; }
    label { display: block; margin: 15px 0 5px; font-weight: bold; font-size: 1.2em; }
    input[type=text], input[type=file] { width: 100%; padding: 10px; font-size: 1.2em; }
    button { margin-top: 20px; padding: 12px 24px; font-size: 1.2em; }
    .article-label { margin-top: 25px; font-size: 1.3em; font-weight: bold; }
    .table-container { overflow-x: auto; margin-top: 30px; }
    table { border-collapse: collapse; width: 100%; table-layout: auto; }
    th, td { border: 1px solid #444; padding: 10px; vertical-align: top; word-wrap: break-word; white-space: normal; }
    th { background: #ddd; font-weight: bold; font-size: 1.1em; text-align: center; }
    td { min-width: 25ch; }
    td.detailed { min-width: 50ch; }
    .highlight { color: red; font-weight: bold; }
    a.summary-link { color: #00e; text-decoration: underline; }
  </style>
</head>
<body>
  <h1>Analyse Disciplinaire</h1>
  <div class="form-container">
    <form method="post" enctype="multipart/form-data">
      <label for="file">Fichier Excel</label>
      <input type="file" id="file" name="file" required>
      <label for="article">Article à filtrer (ex : 2.01 ou 59(2))</label>
      <input type="text" id="article" name="article" value="14" required>
      <button type="submit">Analyser</button>
    </form>
  </div>
  <hr>
  {% if searched_article %}
    <div class="article-label">Article recherché : <span class="highlight">{{ searched_article }}</span></div>
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

# strict pattern: match only Art. or Art: followed by article
ART_PATTERN = r'(?:(?<=Art\.\s)|(?<=Art:\s))({})(?![0-9.])'
DETAILED_COLS = { 'articles enfreints', 'durée totale effective radiation', 'article amende/chef', 'autres sanctions' }

@app.route('/', methods=['GET','POST'])
def analyze():
    global last_excel, last_article
    if request.method == 'POST':
        file = request.files['file']
        article = request.form['article'].strip()
        last_article = article
        df = pd.read_excel(file)
        summary_col = next((c for c in df.columns if c.lower()=='résumé'), None)
        num = re.escape(article)
        pat = re.compile(ART_PATTERN.format(num))
        mask = df['Articles enfreints'].astype(str).apply(lambda x: bool(pat.search(x)))
        filtered = df[mask].copy()
        html_df = filtered.copy()
        if summary_col:
            html_df[summary_col] = html_df[summary_col].apply(lambda u: f'<a href="{u}" class="summary-link" target="_blank">Résumé</a>' if pd.notna(u) else '')
            cols = [c for c in html_df.columns if c!=summary_col] + [summary_col]
            html_df = html_df[cols]
        def highlight(val, col):
            if pd.isna(val): return ''
            text = str(val)
            if col.lower() in DETAILED_COLS and pat.search(text):
                return f'<span class="highlight">{text}</span>'
            return text
        styled = html_df.copy()
        for col in styled.columns:
            styled[col] = styled[col].apply(lambda v: highlight(v, col))
        table_html = styled.to_html(index=False, escape=False, classes='', table_id='', border=1)
        return render_template_string(HTML_TEMPLATE, table_html=table_html, searched_article=article)
    return render_template_string(HTML_TEMPLATE)

@app.route('/download')
def download():
    if not last_excel or not last_article:
        return redirect(url_for('analyze'))
    fname = f"decisions_filtrees_{last_article}.xlsx"
    return send_file(BytesIO(last_excel), as_attachment=True, download_name=fname)

if __name__=='__main__':
    app.run(debug=True)






























































































































































