import re
import sys
import pandas as pd
import os
from flask import Flask, request, render_template, send_file, redirect, url_for
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils.dataframe import dataframe_to_rows

app = Flask(__name__)

# --- Helper: build strict regex for exact article ---
def make_regex(article):
    # ensure numeric part
    num = int(article)
    return re.compile(rf"\bArt\.?\s*{num}(?!\d)\b", re.IGNORECASE)

# --- Highlight cells in workbook containing the pattern ---
def highlight_article(ws, pattern):
    red = Font(color="FF0000")
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            if cell.value and pattern.search(str(cell.value)):
                cell.font = red

# --- Home / upload form ---
@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

# --- Main analysis route ---
@app.route('/analyze', methods=['POST'])
def analyze():
    # load inputs
    excel_file = request.files.get('file')
    article = request.form.get('article', '').strip()
    if not excel_file or not article:
        return redirect(url_for('index'))

    # compile pattern
    pattern = make_regex(article)

    # read Excel
    df = pd.read_excel(excel_file, engine='openpyxl')

    # filter rows containing article in any column
    mask = df.applymap(lambda v: bool(pattern.search(str(v))) if pd.notna(v) else False)
    filtered = df[mask.any(axis=1)].copy()

    # prepare hyperlink column if exists
    if 'Résumé' in filtered.columns:
        filtered['Résumé'] = filtered['Résumé'].apply(lambda url: f"=HYPERLINK(\"{url}\", \"Résumé\")")

    # build Excel output
    wb = Workbook()
    ws = wb.active
    ws.title = f"Article_{article}"
    for r in dataframe_to_rows(filtered, index=False, header=True):
        ws.append(r)
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].auto_size = True
    highlight_article(ws, pattern)

    output_path = f"filtered_output_{article}.xlsx"
    wb.save(output_path)

    # generate tables
    md = filtered.to_markdown(index=False)
    html_table = (
        "<div class='table-container' style='overflow-x:auto;'>"
        + filtered.to_html(classes='table', index=False, escape=False)
        + "</div>"
    )

    return render_template('results.html', markdown_table=md,
                           html_table=html_table, excel_file=output_path)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))











































































































