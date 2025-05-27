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
.table-container { overflow-x: auto; margin-top: 20px; }
table { border-collapse: collapse; width: max-content; }
th, td { border: 1px solid #444; padding: 8px; vertical-align: top; }
th { background: #ddd; font-weight: bold; font-size: 1.1em; text-align: center; }
.highlight { color: red; font-weight: bold; }
.summary-link { color: #00e; text-decoration: underline; }

/* default narrow columns */
th, td { width: 25ch; }
/* wide columns: Résumé des faits col8, Articles enfreints col9, Durée totale col10, Article amende/chef col11, Autres sanctions col13 */
th:nth-child(8), td:nth-child(8),
th:nth-child(9), td:nth-child(9),
th:nth-child(10), td:nth-child(10),
th:nth-child(11), td:nth-child(11),
th:nth-child(13), td:nth-child(13) {
  width: 50ch;
}
'''

# Full HTML template
HTML_TEMPLATE = '''
<!doctype html>
<html lang="fr">
<head>
  <meta charset="utf-8">
  <title>Analyse Disciplinaire</title>
  <style>
    {{ style_block|safe }}
  </style>
</head>
<body>
  <h1>Analyse Disciplinaire</h1>
  <form method="post" enctype="multipart/form-data">
    <label>Fichier Excel: <input type="file" name="file" required></label>
    <label>Article à filtrer: <input type="text" name="article" required></label>
    <button type="submit">Analyser</button>
  </form>
  <hr>
  {% if searched_article %}
    <h2>Article recherché : <span class="highlight">{{ searched_article }}</span></h2>
    <a href="/download">⬇️ Télécharger le fichier Excel formaté</a>
  {% endif %}
  {% if table_html %}
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
    prefixes_join = '|'.join([r'Art\.\s*', r'Art\s*:\s*'])
    suffix = r'(?![0-9])'
    # pattern: (?:Art\.\s*|Art\s*:\s*)<article>(?![0-9])
    return rf"(?:{prefixes_join}){art}{suffix}"

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
        # find columns
        col_inf = next((c for c in df.columns if c.lower()=='articles enfreints'), None)
        summary_col = next((c for c in df.columns if c.lower()=='résumé'), None)
        pat = build_pattern(article)
        # filter only on Articles enfreints
        mask = df[col_inf].astype(str).apply(lambda v: bool(re.search(pat, v)))
        filtered = df[mask].copy()

        # prepare HTML table
        html_df = filtered.fillna('').astype(str)
        for col in html_df.columns:
            if col.lower() in HIGHLIGHT_COLS:
                html_df[col] = html_df[col].apply(lambda v: re.sub(pat, r'<span class="highlight">\g<0></span>', v))
        if summary_col:
            html_df[summary_col] = html_df[summary_col].apply(lambda u: f'<a href="{u}" class="summary-link" target="_blank">Résumé</a>' if u else '')
        table_html = html_df.to_html(index=False, escape=False)

        # build Excel workbook
        out = BytesIO()
        wb = Workbook(); ws = wb.active
        # header row merged
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(filtered.columns))
        hdr = ws.cell(row=1, column=1, value=f"Article filtré : {article}")
        hdr.font = Font(size=14, bold=True)
        # column headers
        for i, col in enumerate(filtered.columns, start=1):
            c = ws.cell(row=2, column=i, value=col)
            c.fill = grey_fill; c.font = Font(size=12,bold=True); c.border=border; c.alignment=wrap
        # data rows
        for r, (_, row) in enumerate(filtered.iterrows(), start=3):
            for i, col in enumerate(filtered.columns, start=1):
                val = row[col]
                cell = ws.cell(row=r, column=i, value=val)
                cell.border=border; cell.alignment=wrap
                if col.lower() in HIGHLIGHT_COLS and re.search(pat, str(val)):
                    cell.font = red_font
                if summary_col and col==summary_col and val:
                    cell.value='Résumé'; cell.hyperlink=val; cell.font=link_font
        # set column widths
        wide_cols = {'résumé des faits','articles enfreints','durée totale effective radiation','article amende/chef','autres sanctions'}
        for i, col in enumerate(filtered.columns, start=1):
            ws.column_dimensions[get_column_letter(i)].width = 50 if col.lower() in wide_cols else 25
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





















































































































































































