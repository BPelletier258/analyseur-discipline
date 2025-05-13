import pandas as pd
import re
import unicodedata
from flask import Flask, request, render_template
from io import BytesIO
import xlsxwriter

app = Flask(__name__)

def normalize_column(col_name):
    if isinstance(col_name, str):
        col = unicodedata.normalize('NFKD', col_name)
        col = col.encode('ASCII', 'ignore').decode('utf-8')
        col = col.replace("’", "'")
        return re.sub(r'\s+', ' ', col).strip().lower()
    return col_name

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/analyse', methods=['POST'])
def analyse():
    file = request.files.get('file') or request.files.get('fichier_excel')
    article = request.form.get('article', '').strip()

    if not file or not article:
        return render_template('index.html', erreur="Veuillez fournir un fichier Excel et un article.")

    # Lecture et normalisation des colonnes
    df = pd.read_excel(file)
    df.columns = [normalize_column(c) for c in df.columns]
    df = df.loc[:, [c for c in df.columns if c and not c.startswith('unnamed')]]

    # Définir les colonnes obligatoires
    required = [
        'numero de decision', 'nom de lintime', 'articles enfreints',
        'duree totale effective radiation', 'article amende/chef', 'autres sanctions'
    ]
    missing = [c for c in required if c not in df.columns]
    if missing:
        return render_template('index.html', erreur=f"Colonnes manquantes : {', '.join(missing)}")

    # Regex strict pour l'article (évite 114, 149.1, etc.)
    pattern = re.compile(rf'(?<![\d\.])Art[\.:]?\s*{re.escape(article)}(?=[\W]|$)', re.IGNORECASE)

    def match_art(cell):
        text = unicodedata.normalize('NFKD', str(cell))
        return bool(pattern.search(text))

    # Filtrer les décisions pertinentes
    df_filtered = df[df['articles enfreints'].apply(match_art)].reset_index(drop=True)
    if df_filtered.empty:
        return render_template('index.html', erreur=f"Aucun résultat pour l'article {article}.")

    # Génération du tableau Markdown (colonnes requises)
    md_df = df_filtered[required]
    markdown_table = md_df.to_markdown(index=False)

    # Génération du fichier Excel (colonnes requises uniquement)
    excel_df = df_filtered[required]

    output = BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet('Résultats')

    # Formats
    wrap_fmt = workbook.add_format({'text_wrap': True, 'valign': 'top'})
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3'})
    red_fmt = workbook.add_format({'font_color': '#FF0000', 'text_wrap': True, 'valign': 'top'})

    # Écrire les en-têtes
    for col_idx, col_name in enumerate(excel_df.columns):
        worksheet.write(0, col_idx, col_name, header_fmt)
        worksheet.set_column(col_idx, col_idx, 30)

    # Écrire les données
    for row_idx, row in enumerate(excel_df.itertuples(index=False), start=1):
        for col_idx, value in enumerate(row):
            text = str(value)
            if match_art(text):
                worksheet.write(row_idx, col_idx, text, red_fmt)
            else:
                worksheet.write(row_idx, col_idx, text, wrap_fmt)

    workbook.close()
    output.seek(0)

    return render_template(
        'resultats.html',
        table_markdown=markdown_table,
        fichier_excel=output.read(),
        filename=f"resultats_article_{article}.xlsx"
    )

if __name__ == '__main__':
    app.run(debug=True)



























































