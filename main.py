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
        col = col.lower()
        col = re.sub(r'\s+', ' ', col).strip()
        return col
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
    # Supprimer colonnes vides ou non nommées
    df = df.loc[:, df.columns.astype(bool)]

    # Vérification des colonnes pour le Markdown
    required = [
        'numero de decision',
        'nom de lintime',
        'articles enfreints',
        'duree totale effective radiation',
        'article amende/chef',
        'autres sanctions'
    ]
    missing = [c for c in required if c not in df.columns]
    if missing:
        return render_template('index.html', erreur=f"Le fichier est incomplet. Colonnes manquantes : {', '.join(missing)}")

    # Regex strict pour éviter 114, 149.1...
    pattern = re.compile(
        rf'(?<!\d)Art[\.:]?\s*{re.escape(article)}(?=[\s\W]|$)',
        flags=re.IGNORECASE
    )

    # Filtrage précis sur la colonne articles enfreints
    mask = df['articles enfreints'].astype(str).str.contains(pattern, na=False)
    conformes = df[mask].copy()
    if conformes.empty:
        return render_template('index.html', erreur=f"Aucun intime trouvé pour l'article {article} demandé.")

    # Construction du Markdown (ordre et colonnes strictes)
    display_df = conformes[required].reset_index(drop=True)
    markdown_table = display_df.to_markdown(index=False)

    # Préparation Excel : conserver toutes les colonnes, supprimer doublons et colonnes vides
    excel_df = conformes.reset_index(drop=True)
    # Supprimer colonnes vides
    excel_df = excel_df.loc[:, excel_df.columns.astype(bool)]
    # Si double colonne 'resume', ne garder qu'une
    if excel_df.columns.tolist().count('resume') > 1:
        cols = [c for c in excel_df.columns if c != 'resume'] + ['resume']
        excel_df = excel_df[cols]

    output = BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet('Résultats')

    # Formats
    wrap_fmt = workbook.add_format({'text_wrap': True, 'valign': 'top'})
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3'})
    red_text = workbook.add_format({'font_color': '#FF0000', 'text_wrap': True})

    # Écriture des en-têtes
    for idx, col in enumerate(excel_df.columns):
        worksheet.write(0, idx, col, header_fmt)
        worksheet.set_column(idx, idx, 30)

    # Écriture des données
    for r, row in enumerate(excel_df.itertuples(index=False), start=1):
        for c, val in enumerate(row):
            txt = str(val)
            col_name = excel_df.columns[c]
            if col_name == 'resume':
                # Lien uniformisé
                worksheet.write_url(r, c, txt, string='Résumé', cell_format=red_text)
            elif pattern.search(unicodedata.normalize('NFKD', txt)):
                # Coloration du texte seulement
                worksheet.write(r, c, txt, red_text)
            else:
                worksheet.write(r, c, txt, wrap_fmt)

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
















































