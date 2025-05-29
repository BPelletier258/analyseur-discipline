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

# Inline CSS for HTML layout, widths, scroll bar, and enhanced form styling
STYLE_BLOCK = '''
body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 20px; background: #f5f7fa; }
h1 { font-size: 1.65em; margin-bottom: 0.5em; color: #333; }
form { display: flex; flex-wrap: wrap; gap: 1rem; align-items: flex-end; background: #fff; padding: 15px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); max-width: 750px; }
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
/* wide columns (detailed info) */
th:nth-child(8), td:nth-child(8),   /* Résumé des faits */
th:nth-child(9), td:nth-child(9),   /* Articles enfreints */
th:nth-child(10), td:nth-child(10), /* Durée totale effective radiation */
th:nth-child(11), td:nth-child(11), /* Article amende/chef */
th:nth-child(13), td:nth-child(13)  /* Autres sanctions */ {
  width: 50ch;
}
'''

# Build regex matching only in Articles enfreints prefixed by Art. or Art:
def build_pattern(article):
    art = re.escape(article)
    prefixes = [r'Art\.\s*', r'Art\s*:\s*']
    pat = rf"(?:{'|'.join(prefixes)}){art}(?![0-9])"
    return pat

# Excel styling constants
grey_fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
red_font = Font(color="FF0000")
link_font = Font(color="0000FF", underline="single")
border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
wrap_alignment = Alignment(wrap_text=True, vertical='top')

# Columns eligible for Excel highlight
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
        df_raw = pd.read_excel(file)
        summary_col = next((c for c in df_raw.columns if c.lower() == 'résumé'), None)
        comment_col = next((c for c in df_raw.columns if c.lower() == 'commentaires internes'), None)
        pat = build_pattern(article)
        # filter only by Articles enfreints column
        mask = df_raw['Articles enfreints'].astype(str).apply(lambda v: bool(re.search(pat, v)))
        df_filtered = df_raw[mask].copy()

        # fillna for comments
        if comment_col:
            df_filtered[comment_col] = df_filtered[comment_col].fillna('')

        # build HTML table
        html_df = df_filtered.copy().fillna('')
        if summary_col:
            html_df[summary_col] = html_df[summary_col].apply(lambda u: f'<a href="{u}" class="summary-link" target="_blank">Résumé</a>' if u else '')
            cols = [c for c in html_df.columns if c != summary_col] + [summary_col]
            html_df = html_df[cols]
        # highlight searched article in HTML only in detail cols
        detail_cols = ['Articles enfreints','Durée totale effective radiation','Article amende/chef','Autres sanctions']
        for col in detail_cols:
            if col in html_df:
                html_df[col] = html_df[col].astype(str).str.replace(pat, lambda m: f"<span class='highlight'>{m.group(0)}</span>", regex=True)
        table_html = html_df.to_html(index=False, escape=False)

        # build Excel workbook
        output = BytesIO()
        wb = Workbook()
        ws = wb.active
        # title row
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(df_filtered.columns))
        tcell = ws.cell(row=1, column=1, value=f"Article filtré : {article}")
        tcell.font = Font(size=14, bold=True)
        # header row
        for idx, col in enumerate(df_filtered.columns, start=1):
            c = ws.cell(row=2, column=idx, value=col)
            c.fill = grey_fill; c.font = Font(size=12, bold=True); c.border = border; c.alignment = wrap_alignment
        # data rows
        for r_idx, (_, row) in enumerate(df_filtered.iterrows(), start=3):
            for c_idx, col in enumerate(df_filtered.columns, start=1):
                cell = ws.cell(row=r_idx, column=c_idx)
                cell.border = border; cell.alignment = wrap_alignment
                if summary_col and col == summary_col and row[col]:
                    cell.value = 'Résumé'; cell.hyperlink = row[col]; cell.font = link_font
                else:
                    cell.value = row[col]
                if col.lower() in HIGHLIGHT_COLS and re.search(pat, str(row[col])):
                    cell.font = red_font
        # set column widths
        narrow, wide = 25, 50
        wide_cols = ['Résumé des faits','Articles enfreints','Durée totale effective radiation','Article amende/chef','Autres sanctions']
        for idx, col in enumerate(df_filtered.columns, start=1):
            ws.column_dimensions[get_column_letter(idx)].width = wide if col in wide_cols else narrow
        wb.save(output)
        output.seek(0)
        last_excel = output.getvalue()

        return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK, table_html=table_html, searched_article=article)
    # GET
    return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK)

@app.route('/download')
def download():
    global last_excel, last_article
    if not last_excel or not last_article:
        return redirect(url_for('analyze'))
    fname = f"decisions_filtrees_{last_article}.xlsx"
    return send_file(BytesIO(last_excel), as_attachment=True, download_name=fname)

if __name__ == '__main__':
    app.run(debug=True)


























































































































































































