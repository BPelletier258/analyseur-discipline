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

# Inline CSS for HTML layout, widths, and scroll
STYLE_BLOCK = '''
.table-container { overflow-x: auto; margin-top: 30px; }
table { border-collapse: collapse; width: max-content; }
th, td { border: 1px solid #444; padding: 10px; vertical-align: top; }
th { background: #ddd; font-weight: bold; font-size: 1.1em; text-align: center; }
.highlight { color: red; font-weight: bold; }
.summary-link { color: #00e; text-decoration: underline; }
/* default narrow columns */
td, th { width: 25ch; }
/* wide columns: Résumé des faits, Articles enfreints, Durée totale effective radiation, Article amende/chef, Autres sanctions */
th:nth-child(n+8):nth-child(-n+12), th:nth-child(7) { width: 50ch; }
td:nth-child(n+8):nth-child(-n+12), td:nth-child(7) { width: 50ch; }
'''

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

# Build regex matching only in Articles enfreints prefixed by Art. or Art:
def build_pattern(article):
    art = re.escape(article)
    prefixes = [r'Art\.\s*', r'Art\s*:\s*']
    suffix = r'(?![0-9])'
    return rf'(?:{"|".join(prefixes)}){art}{suffix}'

# Excel styles
grey_fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
red_font = Font(color="FF0000")
link_font = Font(color="0000FF", underline="single")
border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
wrap = Alignment(wrap_text=True, vertical='top')

# Columns eligible for highlighting
HIGHLIGHT_COLS = {
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
        # identify columns
        col_inf = next((c for c in df.columns if c.lower()=='articles enfreints'), None)
        summary_col = next((c for c in df.columns if c.lower()=='résumé'), None)
        # filter
        pat = build_pattern(article)
        mask = df[col_inf].astype(str).apply(lambda v: bool(re.search(pat, v)))
        filtered = df[mask].copy()
        # prepare HTML table
        html_df = filtered.copy().fillna('')
        # highlight
        for col in html_df.columns:
            html_df[col] = html_df[col].astype(str).apply(
                lambda v, c=col: re.sub(pat, r'<span class="highlight">\g<0></span>', v) if c.lower() in HIGHLIGHT_COLS else v
            )
        # summary links
        if summary_col:
            html_df[summary_col] = html_df[summary_col].apply(
                lambda u: f'<a href="{u}" class="summary-link" target="_blank">Résumé</a>' if u else ''
            )
        table_html = html_df.to_html(index=False, escape=False)
        # build Excel
        out = BytesIO()
        wb = Workbook(); ws = wb.active
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(filtered.columns))
        c0 = ws.cell(row=1, column=1, value=f"Article filtré : {article}"); c0.font = Font(size=14, bold=True)
        # headers
        for i, col in enumerate(filtered.columns, start=1):
            c = ws.cell(row=2, column=i, value=col)
            c.fill=grey_fill; c.font=Font(size=12,bold=True); c.border=border; c.alignment=wrap
        # data
        for r,(_,row) in enumerate(filtered.iterrows(),start=3):
            for i,col in enumerate(filtered.columns, start=1):
                cell = ws.cell(row=r, column=i)
                cell.alignment=wrap; cell.border=border
                if summary_col and col==summary_col:
                    cell.value='Résumé'; cell.hyperlink=str(row[col]); cell.font=link_font
                else:
                    cell.value=row[col]
                if col.lower() in HIGHLIGHT_COLS and re.search(pat, str(row[col])):
                    cell.font=red_font
        # column widths
        wide, narrow = 50,25
        for i,col in enumerate(filtered.columns, start=1):
            low = col.lower()
            if low in HIGHLIGHT_COLS or low=='résumé des faits':
                ws.column_dimensions[get_column_letter(i)].width=wide
            else:
                ws.column_dimensions[get_column_letter(i)].width=narrow
        wb.save(out); out.seek(0)
        last_excel = out.getvalue()
        return render_template_string(HTML_TEMPLATE, table_html=table_html, searched_article=article, style_block=STYLE_BLOCK)
    return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK)

@app.route('/download')
def download():
    global last_excel, last_article
    if not last_excel or not last_article:
        return redirect(url_for('analyze'))
    fname = f"decisions_filtrees_{last_article}.xlsx"
    return send_file(BytesIO(last_excel), as_attachment=True, download_name=fname)

if __name__=='__main__':
    app.run(debug=True)














































































































































































