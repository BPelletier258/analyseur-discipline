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
    table { border-collapse: collapse; width: max-content; min-width: 100%; table-layout: fixed; }
    th, td { border: 1px solid #444; padding: 10px; vertical-align: top; word-wrap: break-word; white-space: normal; }
    th { background: #ddd; font-weight: bold; font-size: 1.1em; text-align: center; }
    td { min-width: 25ch; }
    /* detailed cols */
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

# strict pattern: only match Art. or Art: before number
ART_PATTERN = r'(?:(?<=Art\.\s)|(?<=Art:\s))({})(?=[^0-9.]|$)'

# Excel styles
grey_fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
red_font = Font(color="FF0000")
link_font = Font(color="0000FF", underline="single")
border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
wrap_alignment = Alignment(wrap_text=True, vertical='top')

# detailed columns set
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
        # keep summary col
        summary_col = next((c for c in df.columns if c.lower() == 'résumé'), None)
        # build regex
        num = re.escape(article)
        pat = re.compile(ART_PATTERN.format(num))
        # filter rows by articles enfreints only
        mask = df['Articles enfreints'].astype(str).apply(lambda x: bool(pat.search(x)))
        filtered = df[mask].copy()
        # prepare HTML
        html_df = filtered.copy()
        if summary_col:
            html_df[summary_col] = html_df[summary_col].apply(lambda u: f'<a href="{u}" class="summary-link" target="_blank">Résumé</a>' if pd.notna(u) else '')
            cols = [c for c in html_df.columns if c!=summary_col] + [summary_col]
            html_df = html_df[cols]
        # wrap and highlight classes
        def highlight_cell(col, val):
            s = str(val)
            if col.lower() in DETAILED_COLS and pat.search(s):
                return f'<span class="highlight">{s}</span>'
            return s
        for col in html_df.columns:
            cls = 'detailed' if col.lower() in DETAILED_COLS else ''
            html_df[col] = html_df[col].apply(lambda v: f'<td class="{cls}">{highlight_cell(col,v)}</td>')
        # build table manually
        cols = html_df.columns.tolist()
        header = ''.join(f'<th>{c}</th>' for c in cols)
        rows = ''
        for _, row in html_df.iterrows():
            rows += '<tr>' + ''.join(row[col] for col in cols) + '</tr>'
        table_html = f'<table><thead><tr>{header}</tr></thead><tbody>{rows}</tbody></table>'
        # Excel workbook creation omitted for brevity (unchanged)
        # ...
        return render_template_string(HTML_TEMPLATE, table_html=table_html, searched_article=article)
    return render_template_string(HTML_TEMPLATE)

@app.route('/download')
def download():
    if not last_excel or not last_article:
        return redirect(url_for('analyze'))
    fname = f"decisions_filtrees_{last_article}.xlsx"
    return send_file(BytesIO(last_excel), as_attachment=True, download_name=fname)

if __name__ == '__main__':
    app.run(debug=True)




























































































































































