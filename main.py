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
    th, td { border: 1px solid #444; padding: 10px; vertical-align: top; word-wrap: break-word; }
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
      <input type="text" id="article" name="article" value="" placeholder="ex: 14 ou 59(2)" required>
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

# build regex for strict HTML highlighting: must be prefixed Art. or Art:
def html_pattern(article):
    a = re.escape(article)
    return rf"(Art[\.:]\s*{a})(?![0-9A-Za-z])"

# Excel styles (unchanged)
grey_fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
red_font = Font(color="FF0000")
link_font = Font(color="0000FF", underline="single")
border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
wrap_alignment = Alignment(wrap_text=True, vertical='top')

# columns eligible for HTML highlighting
disp_cols = {
    'articles enfreints',
    'périodes de radiation',
    'amendes',
    'autres sanctions'
}

@app.route('/', methods=['GET','POST'])
def analyze():
    global last_excel, last_article
    table_html = None
    if request.method == 'POST':
        file = request.files['file']
        article = request.form['article'].strip()
        last_article = article
        df_raw = pd.read_excel(file)
        # find summary col
        summary_col = next((c for c in df_raw.columns if c.lower()=='résumé'), None)
        # filter rows containing article anywhere
        pat = re.compile(rf'(?<![0-9]){re.escape(article)}(?![0-9])')
        mask = df_raw.apply(lambda r: any(pat.search(str(v)) for v in r), axis=1)
        df = df_raw[mask].copy()
        # build HTML df
        html_df = df.copy()
        # convert summary to link
        if summary_col:
            html_df[summary_col] = html_df[summary_col].apply(lambda u: f'<a href="{u}" class="summary-link" target="_blank">Résumé</a>' if pd.notna(u) else '')
            # move summary last
            cols = [c for c in html_df.columns if c!=summary_col] + [summary_col]
            html_df = html_df[cols]
        # apply HTML highlighting on specific columns
        hpat = re.compile(html_pattern(article))
        for col in html_df.columns:
            if col.lower() in disp_cols:
                html_df[col] = html_df[col].astype(str).apply(lambda s: hpat.sub(r'<span class="highlight">\1</span>', s))
        table_html = html_df.to_html(index=False, escape=False)
        # create Excel as before (unchanged)
        output = BytesIO()
        wb = Workbook()
        ws = wb.active
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(df.columns))
        cell = ws.cell(row=1, column=1, value=f"Article filtré : {article}")
        cell.font = Font(size=14, bold=True)
        # headers row
        for idx, col in enumerate(df.columns, start=1):
            c = ws.cell(row=2, column=idx, value=col)
            c.fill = grey_fill
            c.font = Font(size=12, bold=True)
            c.border = border
            c.alignment = wrap_text=True, vertical='top'
        # data rows
        for r_idx, (_, row) in enumerate(df.iterrows(), start=3):
            for c_idx, col in enumerate(df.columns, start=1):
                ce = ws.cell(row=r_idx, column=c_idx)
                ce.border = border
                ce.alignment = wrap_text=True, vertical='top'
                if summary_col and col==summary_col:
                    url = row[col]
                    ce.value = 'Résumé'
                    ce.hyperlink = url
                    ce.font = link_font
                else:
                    ce.value = row[col]
                if col.lower() in disp_cols and re.search(rf'(?<![0-9]){re.escape(article)}(?![0-9])', str(row[col])):
                    ce.font = red_font
        # column widths
        for idx in range(1, len(df.columns)+1):
            ws.column_dimensions[get_column_letter(idx)].width = 20
        wb.save(output)
        output.seek(0)
        last_excel = output.getvalue()
    return render_template_string(HTML_TEMPLATE, table_html=table_html, searched_article=last_article)

@app.route('/download')
def download():
    global last_excel, last_article
    if not last_excel or not last_article:
        return redirect(url_for('analyze'))
    fname = f"decisions_filtrees_{last_article}.xlsx"
    return send_file(BytesIO(last_excel), as_attachment=True, download_name=fname)

if __name__=='__main__':
    app.run(debug=True)







































































































































