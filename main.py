import re
import pandas as pd
from flask import Flask, request, render_template_string, send_file, redirect, url_for
froimport re
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
    .table-container { overflow-x: auto; overflow-y: hidden; margin-top: 30px; white-space: nowrap; }
    table { border-collapse: collapse; display: inline-block; table-layout: fixed; }
    th, td { border: 1px solid #444; padding: 10px; vertical-align: top; word-wrap: break-word; white-space: normal; min-width: 25ch; }
    th { background: #ddd; font-weight: bold; font-size: 1.1em; text-align: center; }
    /* colonnes détaillées */
    .detailed { min-width: 50ch; }
    /* mise en évidence article */
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

# pattern strict : ne matche que Art. x ou Art: x
def build_pattern(article):
    num = re.escape(article.strip())
    return rf"\bArt\.?[:]?!?\s*{num}(?![0-9])"

# Excel styles
grey_fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
red_font = Font(color="FF0000")
link_font = Font(color="0000FF", underline="single")
border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
wrap_alignment = Alignment(wrap_text=True, vertical='top')

# colonnes à surligner
HIGHLIGHT_COLS = {
    'Articles enfreints',
    'Durée totale effective radiation',
    'Article amende/chef',
    'Autres sanctions'
}

@app.route('/', methods=['GET','POST'])
def analyze():
    global last_excel, last_article
    if request.method == 'POST':
        file = request.files['file']
        article = request.form['article'].strip()
        last_article = article

        df = pd.read_excel(file)
        pat = build_pattern(article)
        # filtre sur Articles enfreints uniquement
        df_filtered = df[df['Articles enfreints'].astype(str).str.contains(pat, regex=True, na=False)].copy()

        # préparation HTML
        html_df = df_filtered.copy()
        summary_col = next((c for c in html_df.columns if c.lower()=='résumé'), None)
        comment_col = next((c for c in html_df.columns if c.lower()=='commentaires internes'), None)
        # remplacer NaN par chaîne vide
        html_df = html_df.fillna('')
        if summary_col:
            html_df[summary_col] = html_df[summary_col].apply(lambda u: f'<a href="{u}" class="summary-link" target="_blank">Résumé</a>' if u else '')
        cols = [c for c in html_df.columns if c not in (summary_col, comment_col)]
        if summary_col: cols.append(summary_col)
        if comment_col: cols.append(comment_col)
        html_df = html_df[cols]

        for col in html_df.columns:
            is_detailed = col in HIGHLIGHT_COLS
            def decorate(val):
                s = str(val)
                s = re.sub(pat, lambda m: f'<span class="highlight">{m.group(0)}</span>', s)
                if is_detailed:
                    return f'<span class="detailed">{s}</span>'
                return s
            html_df[col] = html_df[col].apply(decorate)

        table_html = html_df.to_html(index=False, escape=False)

        # Excel generation
        output = BytesIO()
        wb = Workbook()
        ws = wb.active
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(df_filtered.columns))
        ws.cell(row=1, column=1, value=f"Article filtré : {article}").font = Font(size=14, bold=True)
        for i, c in enumerate(df_filtered.columns, start=1):
            cell = ws.cell(row=2, column=i, value=c)
            cell.fill, cell.font, cell.border, cell.alignment = grey_fill, Font(size=12,bold=True), border, wrap_alignment
        for r, row in enumerate(df_filtered.itertuples(index=False), start=3):
            for i, val in enumerate(row, start=1):
                cell = ws.cell(row=r, column=i, value=val)
                cell.border, cell.alignment = border, wrap_alignment
                if df_filtered.columns[i-1] in HIGHLIGHT_COLS and re.search(pat, str(val)):
                    cell.font = red_font
        for i in range(1, len(df_filtered.columns)+1):
            ws.column_dimensions[get_column_letter(i)].width = 20
        wb.save(output)
        output.seek(0)
        last_excel = output.getvalue()

        return render_template_string(HTML_TEMPLATE, table_html=table_html, searched_article=article)
    return render_template_string(HTML_TEMPLATE)

@app.route('/download')
def download():
    global last_excel, last_article
    if not last_excel:
        return redirect(url_for('analyze'))
    return send_file(BytesIO(last_excel), as_attachment=True, download_name=f"decisions_filtrees_{last_article}.xlsx")

if __name__ == '__main__':
    app.run(debug=True)























































































































































