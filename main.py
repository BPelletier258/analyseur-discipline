
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
    table td, table th { min-width: 25ch; }
    th, td { border: 1px solid #444; padding: 10px; vertical-align: top; word-wrap: break-word; white-space: normal; }
    th { background: #ddd; font-weight: bold; font-size: 1.1em; text-align: center; }
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

# regex matches only prefix Art. or Art: before article
ART_PATTERN = r'(?:(?<=Art\. )|(?<=Art: ))({})(?![0-9])'

# detail columns for HTML highlighting and width
DETAILED_COLS = {
    'articles enfreints',
    'durée totale effective radiation',
    'article amende/chef',
    'autres sanctions'
}

# columns overall for HTML width
ALL_COLS_WIDTH = {'default': 25, 'detailed': 50}
 or Art: before article
ART_PATTERN = r'(?:(?<=Art\. )|(?<=Art: ))({})(?![0-9])'

# detail columns for HTML highlighting and width
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
        # find summary and internal comments columns (case-insensitive)
        summary_col = next((c for c in df.columns if c.lower() == 'résumé'), None)
        comment_col = next((c for c in df.columns if 'commentaire' in c.lower()), None)

        # compile pattern
        num = re.escape(article)
        pat = re.compile(ART_PATTERN.format(num))

        # filter by articles enfreints only
        mask = df['Articles enfreints'].astype(str).apply(lambda x: bool(pat.search(x)))
        filtered = df[mask].copy()

        # prepare HTML DataFrame
        html_df = filtered.copy()
        # render summary column as link
        if summary_col:
            html_df[summary_col] = html_df[summary_col].apply(
                lambda u: f'<a href="{u}" class="summary-link" target="_blank">Résumé</a>' if pd.notna(u) else ''
            )
        # fill internal comments blanks
        if comment_col:
            html_df[comment_col] = html_df[comment_col].fillna('')

        # reorder to put comment and summary at end
        cols = [c for c in html_df.columns if c not in {summary_col, comment_col}]
        if comment_col: cols.append(comment_col)
        if summary_col: cols.append(summary_col)
        html_df = html_df[cols]

        # highlight matches in detail columns
        def highlight_text(txt, col):
            if col.lower() in DETAILED_COLS and txt:
                return pat.sub(r'<span class="highlight">\1</span>', str(txt))
            return str(txt) if pd.notna(txt) else ''

        for c in html_df.columns:
            html_df[c] = html_df[c].apply(lambda v: highlight_text(v, c))

        # build style block for detailed columns (nth-child indices)
        detailed_indices = [i+1 for i, c in enumerate(html_df.columns) if c.lower() in DETAILED_COLS]
        style_block = ''
        for idx in detailed_indices:
            style_block += f"table td:nth-child({idx}), table th:nth-child({idx}) {{ min-width: 50ch; }}\n"

        # generate HTML
        html = html_df.to_html(index=False, escape=False)
        # inject 'detailed' class on td for detailed columns
        for idx in detailed_indices:
            html = html.replace(f'<td>', '<td class="detailed">', filtered.shape[0])
        return render_template_string(HTML_TEMPLATE,
                                      table_html=html,
                                      searched_article=article,
                                      style_block=style_block)
    return render_template_string(HTML_TEMPLATE, style_block='', table_html=None, searched_article='')

@app.route('/download')
def download():
    global last_excel, last_article
    if not last_article:
        return redirect(url_for('analyze'))
    fname = f"decisions_filtrees_{last_article}.xlsx"
    return send_file(BytesIO(last_excel), as_attachment=True, download_name=fname)

if __name__=='__main__':
    app.run(debug=True)




































































































































































