import pandas as pd
import re
import unicodedata
from io import BytesIO
from bs4 import BeautifulSoup
from flask import Flask, request, jsonify, send_file, flash, redirect, render_template
from werkzeug.utils import secure_filename
import os

app = Flask(__name__)
app.secret_key = 'secret'
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def normalize_column(col_name):
    if isinstance(col_name, str):
        col_name = unicodedata.normalize('NFKD', col_name).encode('ASCII', 'ignore').decode('utf-8')
        col_name = col_name.lower().strip()
        col_name = col_name.replace("’", "'")
        col_name = re.sub(r'\s+', ' ', col_name)
    return col_name

def highlight_article(text, article):
    pattern = rf'(Art[\.:]?\s*{re.escape(article)}(?=[\s\W]|$))'
    return re.sub(pattern, r'**\1**', text, flags=re.IGNORECASE)

def remove_html_tags(text):
    if isinstance(text, str):
        return BeautifulSoup(text, "html.parser").get_text()
    return text

def build_markdown_table(df, article):
    headers = ["Statut", "Numéro de décision", "Nom de l'intime", "Articles enfreints",
               "Périodes de radiation", "Amendes", "Autres sanctions", "Résumé"]
    rows = []
    for _, row in df.iterrows():
        resume_link = row.get('resume', '')
        resume_md = f"[Résumé]({resume_link})" if pd.notna(resume_link) and resume_link else ""
        ligne = [
            str(row.get("statut", "")),
            str(row.get("numero de decision", "")),
            str(row.get("nom de l'intime", "")),
            highlight_article(str(row.get("articles enfreints", "")), article),
            highlight_article(str(row.get("duree totale effective radiation", "")), article),
            highlight_article(str(row.get("article amende/chef", "")), article),
            highlight_article(str(row.get("autres sanctions", "")), article),
            resume_md
        ]
        rows.append("| " + " | ".join(ligne) + " |")

    header_row = "| " + " | ".join(headers) + " |"
    separator = "|" + " --- |" * len(headers)
    return "\n".join([header_row, separator] + rows)

def build_excel_result(df, article):
    from openpyxl import Workbook
    from openpyxl.utils.dataframe import dataframe_to_rows
    from openpyxl.styles import Font, Alignment, PatternFill

    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Résultats"

    df_copy = df.copy()
    for i, row in df_copy.iterrows():
        lien = row.get('resume', '')
        if pd.notna(lien) and lien:
            df_copy.at[i, 'resume'] = f'=HYPERLINK("{lien}", "Résumé")'
        else:
            df_copy.at[i, 'resume'] = ''

    ordered_columns = [
        'statut', 'numero de decision', "nom de l'intime", 'articles enfreints',
        'duree totale effective radiation', 'article amende/chef', 'autres sanctions', 'resume'
    ]
    df_copy = df_copy[ordered_columns]
    df_copy = df_copy.dropna(how='all')
    for col in df_copy.columns:
        df_copy[col] = df_copy[col].apply(remove_html_tags)

    for r in dataframe_to_rows(df_copy, index=False, header=True):
        ws.append(r)

    header_font = Font(bold=True)
    fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = fill

    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(wrap_text=True)

    for col in ws.columns:
        col_letter = col[0].column_letter
        ws.column_dimensions[col_letter].width = 30

    wb.save(output)
    output.seek(0)
    return output

def analyse_article(df, article):
    required_cols = [
        'articles enfreints', 'duree totale effective radiation', 'article amende/chef',
        'autres sanctions', "nom de l'intime", 'numero de decision'
    ]
    for col in required_cols:
        if col not in df.columns:
            raise ValueError("Erreur : Le fichier est incomplet. Merci de vérifier la structure.")

    pattern_explicit = rf'Art[\.:]?\s*{re.escape(article)}(?=[\s\W]|$)'
    mask = df['articles enfreints'].astype(str).str.contains(pattern_explicit, na=False, flags=re.IGNORECASE)
    result = df[mask].copy()

    if result.empty:
        raise ValueError(f"Erreur : Aucun intime trouvé pour l'article {article} demandé.")

    result['statut'] = 'Conforme'
    return result

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/analyse', methods=['POST'])
def analyse():
    article = request.form.get("article")
    fichier = request.files.get("file")

    if not article or not fichier:
        flash("Veuillez fournir un article et un fichier Excel.")
        return redirect('/')

    try:
        df = pd.read_excel(fichier)
        df = df.rename(columns=lambda c: normalize_column(c))
        result = analyse_article(df, article)

        markdown = build_markdown_table(result, article)
        excel_bytes = build_excel_result(result, article)

        markdown_html = f"""
        <html><head>
        <style>
        body {{ font-family: Arial; line-height: 1.6; }}
        table {{ border-collapse: collapse; width: 100%; }}
        td, th {{ border: 1px solid #ccc; padding: 8px; }}
        </style>
        </head><body>
        <h2>Tableau des sanctions pour l'article {article}</h2>
        {markdown.replace('\n', '<br>')}
        <br><br>
        <form method='get' action='/download'>
        <button type='submit'>Télécharger le fichier Excel</button>
        </form>
        </body></html>
        """
        with open("last_output.xlsx", "wb") as f:
            f.write(excel_bytes.read())

        return markdown_html

    except Exception as e:
        flash(str(e))
        return redirect('/')

@app.route('/download')
def download():
    return send_file("last_output.xlsx", as_attachment=True)

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)










































