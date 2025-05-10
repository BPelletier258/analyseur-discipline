
from flask import Flask, request, render_template, redirect, url_for, send_file, flash
import pandas as pd
import re
import unicodedata
import os
from io import BytesIO
from flask import Markup

app = Flask(__name__)
app.secret_key = 'secret_key'

def normalize_column(col_name):
    if isinstance(col_name, str):
        col_name = unicodedata.normalize('NFKD', col_name).encode('ASCII', 'ignore').decode('utf-8')
        col_name = col_name.replace("’", "'")
        col_name = col_name.lower()
        col_name = re.sub(r'\s+', ' ', col_name).strip()
    return col_name

def highlight_article(text, article):
    pattern = rf'Art[\.:]?\s*{re.escape(article)}(?=[\s\W]|$)'
    return re.sub(pattern, r'**Art. ' + article + r'**', str(text), flags=re.IGNORECASE)

def generate_markdown_table(df, article_number):
    headers = df.columns.tolist()
    md = "| " + " | ".join(headers) + " |\n"
    md += "| " + " | ".join(["---"] * len(headers)) + " |\n"
    for _, row in df.iterrows():
        line = []
        for col in headers:
            cell = str(row[col])
            if col.lower() != 'resume':
                cell = highlight_article(cell, article_number)
            elif cell.startswith("http"):
                cell = f"[Résumé]({cell})"
            line.append(cell)
        md += "| " + " | ".join(line) + " |\n"
    return md

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/analyse', methods=['POST'])
def analyse():
    try:
        article_number = request.form.get('article')
        file = request.files.get('file')
        if not article_number or not file:
            flash("Fichier Excel ou numéro d'article manquant.")
            return redirect(url_for('index'))

        df = pd.read_excel(file)
        df.columns = [normalize_column(col) for col in df.columns]

        required_cols = [
            "articles enfreints", "duree totale effective radiation",
            "article amende/chef", "autres sanctions",
            "nom de l'intime", "numero de decision"
        ]

        for col in required_cols:
            if col not in df.columns:
                flash("Le fichier est incomplet. Merci de vérifier la structure.")
                return redirect(url_for('index'))

        if 'resume' not in df.columns:
            df['resume'] = ""

        pattern = rf'Art[\.:]?\s*{re.escape(article_number)}(?=[\s\W]|$)'
        mask = df["articles enfreints"].astype(str).str.contains(pattern, flags=re.IGNORECASE, na=False)
        filtered = df[mask].copy()

        if filtered.empty:
            flash(f"Aucun intime trouvé pour l'article {article_number}")
            return redirect(url_for('index'))

        filtered.insert(0, 'Statut', 'Conforme')

        markdown = generate_markdown_table(filtered[[
            "Statut", "numero de decision", "nom de l'intime", "articles enfreints",
            "duree totale effective radiation", "article amende/chef", "autres sanctions", "resume"
        ]], article_number)

        excel_output = BytesIO()
        with pd.ExcelWriter(excel_output, engine='xlsxwriter') as writer:
            filtered.to_excel(writer, sheet_name='Résultats', index=False)
            ws = writer.sheets['Résultats']
            for col_num, width in enumerate([30]*filtered.shape[1]):
                ws.set_column(col_num, col_num, width)
            fmt = writer.book.add_format({'text_wrap': True, 'align': 'top'})
            ws.set_row(0, None, writer.book.add_format({'bold': True, 'bg_color': '#D9D9D9'}))
            for row_num in range(1, len(filtered)+1):
                ws.set_row(row_num, None, fmt)

        excel_output.seek(0)
        return render_template("resultats.html",
                               markdown_table=Markup(markdown),
                               excel_download=True)

    except Exception as e:
        return render_template('index.html', erreur=str(e))

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(debug=True, host="0.0.0.0", port=port)










































