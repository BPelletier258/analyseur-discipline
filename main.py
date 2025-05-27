
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
    {{ style_block|safe }}
    
    body { font-family: Arial, sans-serif; margin: 20px; }
    h1 { font-size: 1.8em; margin-bottom: 0.5em; }
    .form-container { background: #f9f9f9; padding: 20px; border-radius: 5px; max-width: 600px; }
    label { display: block; margin: 15px 0 5px; font-weight: bold; font-size: 1.2em; }
    input[type=text], input[type=file] { width: 100%; padding: 10px; font-size: 1.2em; }
    button { margin-top: 20px; padding: 12px 24px; font-size: 1.2em; }
    .article-label { margin-top: 25px; font-size: 1.3em; font-weight: bold; }
    .table-container { overflow-x: auto; margin-top: 30px; }
    table { border-collapse: collapse; width: 100%; table-layout: auto; }
    /* default cell width 25ch */
    table td, table th { min-width: 25ch; }
    
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

# match only if prefixed by 'Art. ' or 'Art:'
ART_PATTERN = r'(?:(?<=Art\. )|(?<=Art: ))({})(?![0-9])'

# columns to highlight in HTML
DETAILED_COLS = {
    'articles enfreints',
    'durée totale effective radiation',
    'article amende/chef',
    'autres sanctions'
}

@app.route('/', methods=['GET','POST'])
def analyze():
    global last_excel, last_article
    if request.method == 'POST':
        file = request.files['file']
        article = request.form['article'].strip()
        last_article = article
        df = pd.read_excel(file)
        # identify summary and internal comments columns
        summary_col = next((c for c in df.columns if c.lower()=='résumé'), None)
        comment_col = next((c for c in df.columns if 'commentaire' in c.lower()), None)

        # regex to find exact article with prefix
        num = re.escape(article)
        pat = re.compile(ART_PATTERN.format(num))

        # filter rows based on Articles enfreints column only
        mask = df['Articles enfreints'].astype(str).apply(lambda x: bool(pat.search(x)))
        filtered = df[mask].copy()

        # prepare HTML table
        html_df = filtered.copy()
        # render summary as link
        if summary_col:
            html_df[summary_col] = html_df[summary_col].apply(
                lambda u: f'<a href="{u}" class="summary-link" target="_blank">Résumé</a>' if pd.notna(u) else ''
            )
        # blank internal comments
        if comment_col:
            html_df[comment_col] = html_df[comment_col].fillna('')

        # reorder summary/comment at end
        cols = [c for c in html_df.columns if c not in {summary_col, comment_col}] + [comment_col or '', summary_col or '']
        html_df = html_df[cols]

                # highlight only matching text in the 4 detail columns
        def apply_highlight(text, col):
            if col.lower() in DETAILED_COLS and pd.notna(text):
                return pat.sub(r'<span class="highlight"># highlight only matching text in the 4 detail columns
        def apply_highlight(text, col):
            if col.lower() in DETAILED_COLS and pd.notna(text):
                return pat.sub(r'<span class="highlight">Art. \1</span>', str(text))
            return str(text) if pd.notna(text) else ''

        for col in html_df.columns:
            html_df[col] = html_df[col].apply(lambda v: apply_highlight(v, col))

        # add class to detail columns for width</span>', str(text))
            return str(text) if pd.notna(text) else ''

        for col in html_df.columns:
            html_df[col] = html_df[col].apply(lambda v: apply_highlight(v, col))

        # generate style overrides for detailed columns widths
        detailed_indices = [i+1 for i,c in enumerate(html_df.columns) if c.lower() in DETAILED_COLS]
        # build CSS rules
        style_block = ""
        for idx in detailed_indices:
            style_block += f"table td:nth-child({idx}), table th:nth-child({idx}) {{ min-width: 50ch; }}
        "

        # render HTML
        html = html_df.to_html(index=False, escape=False)
        table_html = html
        return render_template_string(HTML_TEMPLATE, table_html=table_html, searched_article=article, style_block=style_block)
        html = html_df.to_html(index=False, escape=False)
        # inject td class
        for col in DETAILED_COLS:
            html = html.replace(f'<td>', f'<td class="detailed">',  filtered.shape[0])

        table_html = html
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
































































































































































