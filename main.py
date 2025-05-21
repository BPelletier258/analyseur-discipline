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
    .form-container { background: #f9f9f9; padding: 20px; border-radius: 5px; max-width: 800px; }
    label { display: block; margin: 15px 0 5px; font-weight: bold; font-size: 1.2em; }
    input[type=text], input[type=file] { width: 100%; padding: 10px; font-size: 1.2em; }
    button { margin-top: 20px; padding: 12px 24px; font-size: 1.2em; }
    .article-label { margin-top: 25px; font-size: 1.4em; font-weight: bold; }
    .table-container { overflow-x: auto; margin-top: 30px; }
    table { border-collapse: collapse; width: 100%; table-layout: fixed; }
    th, td { border: 1px solid #444; padding: 10px; vertical-align: top; word-wrap: break-word; width: 25ch; text-align: center; }
    /* agrandir colonnes détaillées */
    th:nth-child(8), td:nth-child(8),
    th:nth-child(9), td:nth-child(9),
    th:nth-child(10), td:nth-child(10),
    th:nth-child(12), td:nth-child(12) {
      width: 50ch;
    }
    th { background: #ddd; font-weight: bold; font-size: 1.1em; }
    a.summary-link { color: #00e; text-decoration: underline; }
    span.highlight { color: red; font-weight: bold; }
  </style>
</head>
<body>
  <h1>Analyse Disciplinaire</h1>
  <div class="form-container">
    <form method="post" enctype="multipart/form-data">
      <label for="file">Fichier Excel</label>
      <input type="file" id="file" name="file" required>
      <label for="article">Numéro d'article</label>
      <input type="text" id="article" name="article" placeholder="ex: 14 ou 59(2)" required>
      <button type="submit">Analyser</button>
    </form>
  </div>
  <hr>
  {% if searched_article %}
    <div class="article-label">Résultats pour l'article : <span class="highlight">{{ searched_article }}</span></div>
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

# compile regex to match only 'Art.' or 'Art:' prefix
def html_pattern(article):
    art = re.escape(article)
    return rf"(?:(?:Art\.|Art:))\s*{art}(?![0-9A-Za-z])"

# Excel styling
grey_fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
red_font = Font(color="FF0000")
link_font = Font(color="0000FF", underline="single")
border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
wrap_alignment = Alignment(wrap_text=True, vertical='top')

# columns eligible for HTML highlighting
HIGHLIGHT_COLS = {
    'articles enfreints',
    'durée totale effective radiation',
    'article amende/chef',
    'autres sanctions'
}

@app.route('/', methods=['GET', 'POST'])
def analyze():
    global last_excel, last_article
    table_html = None
    if request.method == 'POST':
        file = request.files['file']
        article = request.form['article'].strip()
        last_article = article
        df_raw = pd.read_excel(file)
        # filter strictly on 'Articles enfreints' column with prefix
        target_col = next((c for c in df_raw.columns if c.lower() == 'articles enfreints'), None)
        if target_col:
            pat = re.compile(html_pattern(article))
            df = df_raw[df_raw[target_col].astype(str).apply(lambda s: bool(pat.search(s)))].copy()
        else:
            df = df_raw.copy()
        # prepare HTML DataFrame
        summary_col = next((c for c in df.columns if c.lower() == 'résumé'), None)
        html_df = df.copy()
        if summary_col:
            html_df[summary_col] = html_df[summary_col].apply(
                lambda u: f'<a href="{u}" class="summary-link" target="_blank">Résumé</a>' if pd.notna(u) else ''
            )
            cols = [c for c in html_df.columns if c != summary_col] + [summary_col]
            html_df = html_df[cols]
        # highlight in HTML columns
        pat_h = re.compile(html_pattern(article))
        for col in html_df.columns:
            if col.lower() in HIGHLIGHT_COLS:
                html_df[col] = html_df[col].astype(str).apply(
                    lambda s: pat_h.sub(lambda m: f'<span class="highlight">{m.group(0)}</span>', s)
                )
        # render HTML table
        table_html = html_df.to_html(index=False, escape=False)
        # build Excel
        output = BytesIO()
        wb = Workbook()
        ws = wb.active
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(df.columns))
        header_cell = ws.cell(row=1, column=1, value=f"Article filtré : {article}")
        header_cell.font = Font(size=14, bold=True)
        # headers
        for idx, col in enumerate(df.columns, start=1):
            c = ws.cell(row=2, column=idx, value=col)
            c.fill = grey_fill
            c.font = Font(size=12, bold=True)
            c.border = border
            c.alignment = wrap_alignment
        # rows
        for r_idx, (_, row) in enumerate(df.iterrows(), start=3):
            for c_idx, col in enumerate(df.columns, start=1):
                cell = ws.cell(row=r_idx, column=c_idx)
                cell.border = border
                cell.alignment = wrap_alignment
                if summary_col and col == summary_col:
                    cell.value = 'Résumé'
                    cell.hyperlink = row[col]
                    cell.font = link_font
                else:
                    cell.value = row[col]
                if col.lower() in HIGHLIGHT_COLS and pat_h.search(str(row[col])):
                    cell.font = red_font
        # set uniform column widths
        for i in range(1, len(df.columns) + 1):
            ws.column_dimensions[get_column_letter(i)].width = 20
        wb.save(output)
        output.seek(0)
        last_excel = output.getvalue()
    return render_template_string(HTML_TEMPLATE, table_html=table_html, searched_article=last_article)

@app.route('/download')
def download():
    global last_excel, last_article
    if not last_excel:
        return redirect(url_for('analyze'))
    fname = f"decisions_filtrees_{last_article}.xlsx"
    return send_file(BytesIO(last_excel), as_attachment=True, download_name=fname)

if __name__ == '__main__':
    app.run(debug=True)














































































































































