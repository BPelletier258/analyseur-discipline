import re
import sys
import pandas as pd
import os
from flask import Flask, request, render_template, send_file, redirect, url_for
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows

app = Flask(__name__)

# --- Helper: build strict regex for exact article X or X.Y ---
def make_regex(article):
    escaped = re.escape(article)
    # match Art. X or Art. X.Y without catching longer numbers
    return re.compile(rf"\bArt\.?\s*{escaped}(?![\d.])\b", re.IGNORECASE)

# --- Highlight cells containing the pattern ---
def highlight_article(ws, pattern):
    red = Font(color="FF0000")
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            if cell.value and pattern.search(str(cell.value)):
                cell.font = red

# --- Wrap text and auto-size columns ---
def autosize_and_wrap(ws):
    for col in ws.columns:
        letter = col[0].column_letter
        max_length = 0
        for cell in col:
            cell.alignment = Alignment(wrapText=True)
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[letter].width = min(max_length * 1.1, 50)

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/analyze', methods=['POST'])
def analyze():
    f = request.files.get('file')
    article = request.form.get('article', '').strip()
    if not f or not article:
        return redirect(url_for('index'))

    pattern = make_regex(article)
    df = pd.read_excel(f, engine='openpyxl')

    mask = df.applymap(lambda v: bool(pattern.search(str(v))) if pd.notna(v) else False)
    filtered = df[mask.any(axis=1)].copy()
    # keep 'Résumé' last if exists
    if 'Résumé' in filtered.columns:
        urls = filtered.pop('Résumé')
        filtered['Résumé'] = urls.apply(lambda u: f"=HYPERLINK(\"{u}\", \"Résumé\")")

    # save Excel
    wb = Workbook()
    ws = wb.active
    ws.title = f"Art_{article}"
    for r in dataframe_to_rows(filtered, index=False, header=True):
        ws.append(r)
    highlight_article(ws, pattern)
    autosize_and_wrap(ws)
    output_file = f"decisions_filtrees_{article}.xlsx"
    wb.save(output_file)

    # HTML and Markdown
    html_table = f"<div style='overflow-x:auto;width:100%;'>{filtered.to_html(index=False, escape=False)}</div>"
    markdown_table = filtered.to_markdown(index=False)

    return render_template('results.html', html_table=html_table,
                           markdown_table=markdown_table,
                           excel_file=output_file)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)), debug=True)













































































































