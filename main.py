
import pandas as pd
import re
import unicodedata
from flask import Flask, request, render_template, send_file
from io import BytesIO
import xlsxwriter

app = Flask(__name__)

def normalize_column(col_name):
    if isinstance(col_name, str):
        col_name = unicodedata.normalize('NFKD', col_name).encode('ASCII', 'ignore').decode('utf-8')
        col_name = col_name.replace("’", "'")
        col_name = col_name.lower()
        col_name = re.sub(r'\s+', ' ', col_name).strip()
    return col_name

def style_article(cell, article):
    if not isinstance(cell, str):
        return cell
    pattern = re.compile(rf'(Art[\.\:]?\s*{re.escape(article)})', re.IGNORECASE)
    return pattern.sub(r'<span style="color:red;font-weight:bold">\1</span>', cell)

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/analyse', methods=['POST'])
def analyse():
    try:
        file = request.files['file']
        article = request.form['article'].strip()

        if not file or not article:
            return render_template('index.html', erreur="Veuillez fournir un fichier Excel et un article.")

        df = pd.read_excel(file)
        df = df.rename(columns=lambda c: normalize_column(c))

        required = [
            "articles enfreints",
            "duree totale effective radiation",
            "article amende/chef",
            "autres sanctions",
            "nom de l'intime",
            "numero de decision"
        ]
        for col in required:
            if col not in df.columns:
                return render_template('index.html', erreur="Le fichier est incomplet. Merci de vérifier la structure.")

        pattern_explicit = rf'\bArt[\.\:]?\s*{re.escape(article)}\b'
        mask = df['articles enfreints'].astype(str).str.contains(pattern_explicit, na=False, flags=re.IGNORECASE)
        conformes = df[mask].copy()

        if conformes.empty:
            return render_template('index.html', erreur=f"Aucun intime trouvé pour l'article {article} demandé.")

        conformes['Statut'] = 'Conforme'

        # Appliquer la mise en forme HTML dans toutes les colonnes pertinentes
        colonnes_cibles = [
            "articles enfreints", "duree totale effective radiation",
            "article amende/chef", "autres sanctions"
        ]
        for col in colonnes_cibles:
            conformes[col] = conformes[col].apply(lambda x: style_article(x, article))

        # Créer une table HTML Markdown simplifiée sans la colonne Résumé
        display_df = conformes[[
            'Statut',
            'numero de decision',
            "nom de l'intime",
            "articles enfreints",
            "duree totale effective radiation",
            "article amende/chef",
            "autres sanctions"
        ]]
        markdown_table = display_df.to_markdown(index=False)

        # Générer le fichier Excel avec mise en forme
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet('Résultats')

        # Formats
        wrap = workbook.add_format({'text_wrap': True, 'valign': 'top'})
        header = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3'})
        rouge = workbook.add_format({'bg_color': '#FFC7CE', 'text_wrap': True})

        # Entêtes
        for col_num, col_name in enumerate(display_df.columns):
            worksheet.write(0, col_num, col_name, header)
            worksheet.set_column(col_num, col_num, 30)

        # Lignes
        for row_num, row in enumerate(display_df.values, 1):
            for col_num, value in enumerate(row):
                cell_str = str(value)
                if re.search(pattern_explicit, cell_str, flags=re.IGNORECASE):
                    worksheet.write(row_num, col_num, cell_str, rouge)
                else:
                    worksheet.write(row_num, col_num, cell_str, wrap)

        workbook.close()
        output.seek(0)

        return render_template("resultats.html",
            table_markdown=markdown_table,
            fichier_excel=output.read(),
            filename=f"resultats_article_{article}.xlsx"
        )

    except Exception as e:
        return render_template('index.html', erreur=str(e))

if __name__ == '__main__':
    app.run(debug=True)










































