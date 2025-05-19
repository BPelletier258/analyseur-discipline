import re
import glob
import sys
import pandas as pd
from flask import Flask, request, render_template, send_file
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows

# --- Flask app ---
app = Flask(__name__)

# --- Helper: build strict regex for exact article ---
def make_regex(article):
    num = int(article)
    # match 'Art. <num>' or 'Article <num>' with word boundaries, not part of larger numbers
    return re.compile(rf"\bArt\.?\s*{num}(?!\d)\b", re.IGNORECASE)

# --- Highlight cells in workbook containing the pattern ---
def highlight_article(ws, pattern):
    red = Font(color="FF0000")
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            if cell.value and pattern.search(str(cell.value)):
                cell.font = red

# --- Main analysis route ---
@app.route('/analyze', methods=['POST'])
def analyze():
    excel_file = request.files['file']
    article = request.form['article']
    pattern = make_regex(article)

    # read
    df = pd.read_excel(excel_file, engine='openpyxl')

    # filter rows containing the article in any column
    mask = df.applymap(lambda v: bool(pattern.search(str(v))) if pd.notna(v) else False)
    filtered = df[mask.any(axis=1)].copy()

    # prepare resumo hyperlinks
    filtered['Résumé'] = filtered['Résumé'].apply(lambda url: f"=HYPERLINK(\"{url}\", \"Résumé\")")

    # build Excel
    wb = Workbook()
    ws = wb.active
    ws.title = f"Article_{article}"
    for r in dataframe_to_rows(filtered, index=False, header=True):
        ws.append(r)
    # format
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].auto_size = True
    highlight_article(ws, pattern)

    output_path = f"filtered_output_{article}.xlsx"
    wb.save(output_path)

    # markdown table
    md = filtered.to_markdown(index=False)
    html_table = f"<div class='table-container'>{filtered.to_html(classes='table', index=False, escape=False)}</div>"

    return render_template('results.html', markdown_table=md, html_table=html_table, excel_file=output_path)

if __name__ == '__main__':
    app.run(debug=True)










































































































