
from flask import Flask, request, render_template, send_file
import pandas as pd
import re
import unicodedata
from io import BytesIO
import xlsxwriter
import markdown2

app = Flask(__name__)

def normalize_column(col_name):
    col_name = unicodedata.normalize('NFKD', str(col_name)).encode('ASCII', 'ignore').decode('utf-8')
    col_name = col_name.replace("’", "'")
    col_name = col_name.lower()
    col_name = re.sub(r'\s+', ' ', col_name).strip()
    return col_name

def highlight_article(text, article):
    if not isinstance(text, str):
        return text
    pattern = re.compile(rf'(Art[\.:]?\s*{re.escape(article)}(?=[\s\W]|$))', flags=re.IGNORECASE)
    return pattern.sub(r'**\1**', text)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/analyse', methods=['POST'])
def analyse():
    if 'file' not in request.files or 'article' not in request.form:
        return render_template('index.html', erreur="Fichier ou article manquant.")

    file = request.files['file']
    article = request.form['article'].strip()

    try:
        df = pd.read_excel(file)
        df.columns = [normalize_column(col) for col in df.columns]

        required_columns = [
            "articles enfreints", "duree totale effective radiation",
            "article amende/chef", "autres sanctions",
            "nom de l'intime", "numero de decision"
        ]
        for col in required_columns:
            if col not in df.columns:
                return render_template('index.html', erreur=f"Colonne manquante : {col}")

        pattern = re.compile(rf'\bArt[\.:]?\s*{re.escape(article)}(?=[\s\W]|$)', re.IGNORECASE)
        filtered = df[df["articles enfreints"].astype(str).str.contains(pattern)]

        if filtered.empty:
            return render_template('index.html', erreur=f"Aucun intime trouvé pour l'article {article}.")

        filtered = filtered.copy()
        filtered['Statut'] = "Conforme"

        for col in ["articles enfreints", "duree totale effective radiation", "article amende/chef", "autres sanctions"]:
            filtered[col] = filtered[col].apply(lambda x: highlight_article(x, article))

        filtered = filtered[[
            "Statut", "numero de decision", "nom de l'intime",
            "articles enfreints", "duree totale effective radiation",
            "article amende/chef", "autres sanctions", "resume"
        ]]

        # Génération du tableau Markdown
        def build_markdown_table(df):
            headers = df.columns.tolist()
            lines = ["| " + " | ".join(headers) + " |",
                     "| " + " | ".join(["---"] * len(headers)) + " |"]
            for _, row in df.iterrows():
                line = []
                for col in headers:
                    value = row[col]
                    if col.lower() == "resume" and pd.notnull(value):
                        line.append(f"[Résumé]({value})")
                    else:
                        line.append(str(value).replace("\n", " "))
                lines.append("| " + " | ".join(line) + " |")
            return "\n".join(lines)

        markdown_table = build_markdown_table(filtered)
        html_markdown = markdown2.markdown(markdown_table)

        # Création fichier Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            filtered.to_excel(writer, index=False, sheet_name='Résultats')
            workbook = writer.book
            worksheet = writer.sheets['Résultats']
            wrap = workbook.add_format({'text_wrap': True, 'valign': 'top'})
            red_format = workbook.add_format({'font_color': 'red', 'text_wrap': True, 'valign': 'top'})
            header_format = workbook.add_format({'bold': True, 'bg_color': '#D9D9D9', 'align': 'center'})

            for i, col in enumerate(filtered.columns):
                worksheet.set_column(i, i, 30, wrap)
                worksheet.write(0, i, col, header_format)
                for j, val in enumerate(filtered[col], start=1):
                    fmt = red_format if '**' in str(val) else wrap
                    worksheet.write(j, i, str(val).replace('**', ''), fmt)

        output.seek(0)
        return render_template('resultats.html',
                               table_html=html_markdown,
                               fichier_excel=output)

    except Exception as e:
        return render_template('index.html', erreur=str(e))

if __name__ == "__main__":
    app.run(debug=True)










































