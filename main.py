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
.table-container { overflow-x: auto; margin-top: 30px; }
table { border-collapse: collapse; width: max-content; background: #fff; }
th, td { border: 1px solid #888; padding: 8px; vertical-align: top; }
th { background: #e2e3e5; font-weight: bold; font-size: 1em; text-align: center; }
.highlight { color: #d41e26; font-weight: bold; }
.summary-link { color: #0066cc; text-decoration: underline; }

/* default narrow columns */
th, td { width: 25ch; }
/* wide columns */
th:nth-child(8), td:nth-child(8), th:nth-child(9), td:nth-child(9), th:nth-child(10), td:nth-child(10), th:nth-child(11), td:nth-child(11), th:nth-child(13), td:nth-child(13) { width: 50ch; }
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
      <label>Article à filtrer:<input type="text" name="article" value="" placeholder="ex: 14 ou 59(2)" required></label>
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

# Styles for Excel
grey_fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
red_font = Font(color="FF0000")
link_font = Font(color="0000FF", underline="single")
border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
wrap_alignment = Alignment(wrap_text=True, vertical='top')

# Columns eligible for highlight in Excel
HIGHLIGHT_COLS = {'articles enfreints','durée totale effective radiation','article amende/chef','autres sanctions'}

# Build regex matching only in Articles enfreints prefixed by Art. or Art:
def build_pattern(article):
    art = re.escape(article)
    prefixes = [r'Art\.\s*', r'Art\s*:\s*']
    return rf"(?:{'|'.join(prefixes)}){art}(?![0-9])"

@app.route('/', methods=['GET','POST'])
def analyze():
    global last_excel, last_article
    if request.method == 'POST':
        file = request.files['file']
        article = request.form['article'].strip()
        last_article = article
        df = pd.read_excel(file)
        # locate summary and comments
        sum_col = next((c for c in df.columns if c.lower()=='résumé'), None)
        com_col = next((c for c in df.columns if c.lower()=='commentaires internes'), None)
        # filter
        pat = build_pattern(article)
        mask = df['Articles enfreints'].astype(str).apply(lambda v: bool(re.search(pat, v)))
        df_f = df[mask].copy()
        # fill comments
        if com_col: df_f[com_col] = df_f[com_col].fillna('')
        # HTML table
        html_df = df_f.fillna('')
        if sum_col:
            html_df[sum_col] = html_df[sum_col].apply(lambda u: f'<a href="{u}" class="summary-link" target="_blank">Résumé</a>' if u else '')
            cols = [c for c in html_df.columns if c!=sum_col] + [sum_col]
            html_df = html_df[cols]
        for col in ['Articles enfreints','Durée totale effective radiation','Article amende/chef','Autres sanctions']:
            if col in html_df:
                html_df[col] = html_df[col].astype(str).str.replace(pat, lambda m: f"<span class='highlight'>{m.group(0)}</span>", regex=True)
        table_html = html_df.to_html(index=False, escape=False)
        # Excel
        out = BytesIO(); wb=Workbook(); ws=wb.active
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(df_f.columns))
        ws.cell(1,1,f"Article filtré : {article}").font=Font(size=14,bold=True)
        for i,col in enumerate(df_f.columns,1):
            c=ws.cell(2,i,col);c.fill=grey_fill;c.font=Font(size=12,bold=True);c.border=border; c.alignment=wrap_alignment
        for r,row in enumerate(df_f.itertuples(index=False),3):
            for i,val in enumerate(row,1):
                cell=ws.cell(r,i,val if not (sum_col and df_f.columns[i-1]==sum_col) else '')
                cell.border=border;cell.alignment=wrap_alignment
                if sum_col and df_f.columns[i-1]==sum_col and getattr(row,i-1):
                    cell.value='Résumé'; cell.hyperlink=getattr(row,i-1); cell.font=link_font
                if df_f.columns[i-1].lower() in HIGHLIGHT_COLS and re.search(pat,str(getattr(row,i-1))): cell.font=red_font
        # widths
        narrow,wide=25,50
        wide_list=['Résumé des faits','Articles enfreints','Durée totale effective radiation','Article amende/chef','Autres sanctions']
        for i,col in enumerate(df_f.columns,1): ws.column_dimensions[get_column_letter(i)].width = wide if col in wide_list else narrow
        wb.save(out); out.seek(0); last_excel=out.getvalue()
        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=table_html, searched_article=article)
    return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK)

@app.route('/download')
def download():
    if not last_excel or not last_article: return redirect(url_for('analyze'))
    return send_file(BytesIO(last_excel), as_attachment=True, download_name=f"decisions_filtrees_{last_article}.xlsx")

if __name__=='__main__':
    app.run(debug=True)



























































































































































































