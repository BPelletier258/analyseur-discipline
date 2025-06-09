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

# Inline CSS and HTML template
STYLE_BLOCK = '''
body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 20px; background: #f5f7fa; }
h1 { font-size: 1.65em; margin-bottom: 0.5em; color: #333; }
.form-container { background: #fff; padding: 15px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); max-width: 750px; }
form { display: flex; flex-wrap: wrap; gap: 1rem; align-items: flex-end; }
label { font-weight: bold; font-size: 1.05em; color: #444; display: flex; flex-direction: column; }
input[type=file], input[type=text] { padding: 0.6em; font-size: 1.05em; border: 1px solid #ccc; border-radius: 4px; }
button { padding: 0.6em 1.2em; font-size: 1.05em; font-weight: bold; background: #007bff; color: #fff; border: none; border-radius: 4px; cursor: pointer; transition: background 0.3s ease; }
button:hover { background: #0056b3; }
.table-container { width: 100%; overflow-x: scroll; overflow-y: hidden; scrollbar-gutter: stable both-edges; -webkit-overflow-scrolling: touch; margin-top: 30px; }
table { border-collapse: collapse; width: max-content; background: #fff; display: inline-block; }
th, td { border: 1px solid #888; padding: 8px; vertical-align: top; word-wrap: break-word; }
th { background: #e2e3e5; font-weight: bold; font-size: 1em; text-align: center; }
/* default width 25ch */
th, td { width: 25ch; }
/* detail cols 50ch */
th:nth-child(9), td:nth-child(9),
th:nth-child(10), td:nth-child(10),
th:nth-child(11), td:nth-child(11),
th:nth-child(12), td:nth-child(12),
th:nth-child(14), td:nth-child(14) { width: 50ch; }
.highlight { color: #d41e26; font-weight: bold; }
.summary-link { color: #0066cc; text-decoration: underline; }
'''

HTML_TEMPLATE = '''
<!doctype html>
<html lang="fr">
<head>
  <meta charset="utf-8">
  <title>Analyse Disciplinaire</title>
  <style>{{ style_block }}</style>
</head>
<body>
  <h1>Analyse Disciplinaire</h1>
  <div class="form-container">
    <form method="post" enctype="multipart/form-data">
      <label>Fichier Excel:<input type="file" name="file" required></label>
      <label>Article à filtrer:<input type="text" name="article" placeholder="ex: 14 ou 59(2)" required></label>
      <button type="submit">Analyser</button>
    </form>
  </div>
  <hr>
  {% if searched_article %}
    <div><strong>Article recherché : <span class="highlight">{{ searched_article }}</span></strong></div>
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

# build pattern with strict prefixes and optional source

def build_pattern(article):
    art = re.escape(article)
    space = r'(?:\s|\u00A0)*'
    prefixes = [r'Art\.' + space, r'Art:' + space, r'Art' + space + r':' + space]
    pref = '|'.join(prefixes)
    if '(' in article:
        return rf'(?:{pref}){art}'
    else:
        return rf'(?:{pref}){art}(?![0-9])'

# styles
grey_fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
red_font    = Font(color="FF0000")
link_font   = Font(color="0000FF", underline="single")
border      = Border(left=Side('thin'), right=Side('thin'), top=Side('thin'), bottom=Side('thin'))
wrap_align  = Alignment(wrap_text=True, vertical='top')

HIGHLIGHT_COLS = {'Articles enfreints', 'Durée totale effective radiation', 'Article amende/chef', 'Autres sanctions'}

@app.route('/', methods=['GET','POST'])
def analyze():
    global last_excel, last_article
    if request.method == 'POST':
        file = request.files['file']
        art  = request.form['article'].strip()
        # strict validation allowing source
        if not re.match(r'^[0-9]+(?:\.[0-9]+)*(?:\([^)]+\))?(?:\s+[A-Za-z0-9]+)?$', art):
            return "Format d'article non valide. Exemple : 14, 59(2), 2.01a) R15", 400
        last_article = art
        # auto header detect
        preview = pd.read_excel(file, nrows=2, header=None, engine='openpyxl')
        file.seek(0)
        if isinstance(preview.iloc[0,0], str) and preview.iloc[0,0].startswith("Article filtré :"):
            df = pd.read_excel(file, skiprows=1, header=0, engine='openpyxl')
        else:
            df = pd.read_excel(file, header=0, engine='openpyxl')
        pat = build_pattern(art)
        mask = df['Articles enfreints'].astype(str).apply(lambda v: bool(re.search(pat, v)))
        df_f = df[mask].fillna('')
        # highlight in HTML
        for col in HIGHLIGHT_COLS & set(df_f.columns):
            df_f[col] = df_f[col].astype(str).str.replace(pat, lambda m: f"<span class='highlight'>{m.group(0)}</span>", regex=True)
        if 'Résumé' in df_f:
            df_f['Résumé'] = df_f['Résumé'].apply(lambda u: f'<a href="{u}" class="summary-link" target="_blank">Résumé</a>' if u else '')
        table_html = df_f.to_html(index=False, escape=False)
        # Excel
        buf = BytesIO()
        wb = Workbook()
        ws = wb.active
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(df_f.columns))
        c = ws.cell(row=1, column=1, value=f"Article filtré : {art}")
        c.font = Font(size=14, bold=True)
        # headers
        for i, col in enumerate(df_f.columns, start=1):
            cell = ws.cell(row=2, column=i, value=col)
            cell.fill, cell.font, cell.border, cell.alignment = grey_fill, Font(bold=True), border, wrap_align
        # data
        for r, row in enumerate(df_f.itertuples(index=False), start=3):
            for c_idx, col in enumerate(df_f.columns, start=1):
                val = getattr(row, col)
                cell = ws.cell(row=r, column=c_idx, value=val)
                cell.border, cell.alignment = border, wrap_align
                if col in HIGHLIGHT_COLS and re.search(pat, str(val)):
                    cell.font = red_font
        # widths
        for i in range(1, len(df_f.columns)+1): ws.column_dimensions[get_column_letter(i)].width = 25
        for j in (9,10,11,12,14):
            if j <= len(df_f.columns): ws.column_dimensions[get_column_letter(j)].width = 50
        wb.save(buf); buf.seek(0)
        last_excel = buf.getvalue()
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=table_html, searched_article=art)
    return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK)

@app.route('/download')
def download():
    if not last_excel or not last_article: return redirect(url_for('analyze'))
    return send_file(BytesIO(last_excel), as_attachment=True, download_name=f"decisions_filtrees_{last_article}.xlsx")

if __name__ == '__main__':
    app.run(debug=True)























































































































































































































