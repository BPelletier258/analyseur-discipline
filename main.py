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

    # Lecture et normalisation des colonnes (évite index implicite)
    df = pd.read_excel(file, index_col=None)
    df.columns = [normalize_column(c) for c in df.columns]
    # Retirer colonnes vides ou générées (unnamed)
    df = df.loc[:, [c for c in df.columns if c and not c.startswith('unnamed')]]

    # Assurer une seule colonne resume
    # normalisation produit 'resume' pour 'Résumé'
    if df.columns.tolist().count('resume') > 1:
        cols = [c for c in df.columns if c != 'resume'] + ['resume']
        df = df[cols]

    # Colonnes obligatoires pour Markdown
    required = [
      'numero de decision', 'nom de lintime', 'articles enfreints',
      'duree totale effective radiation', 'article amende/chef', 'autres sanctions'
    ]
    missing = [c for c in required if c not in df.columns]
    if missing:
        return render_template('index.html', erreur=f"Le fichier est incomplet. Colonnes manquantes : {', '.join(missing)}")

    # Regex strict pour l'article
    pattern = re.compile(rf'(?<!\d)Art[\.:]?\s*{re.escape(article)}(?=[\s\W]|$)', re.IGNORECASE)

    # Filtrage précis
    mask = df['articles enfreints'].astype(str).apply(lambda x: bool(pattern.search(x)))
    conformes = df[mask].reset_index(drop=True)
    if conformes.empty:
        return render_template('index.html', erreur=f"Aucun intime trouvé pour l'article {article} demandé.")

        # Markdown ordonné: sélectionner colonnes requises et retirer noms vides
    display_df = conformes[required].copy().reset_index(drop=True)
    # Éliminer toute colonne vide dans Markdown
    display_df = display_df.loc[:, display_df.columns.astype(bool)]
    markdown_table = display_df.to_markdown(index=False)

    # Préparation Excel complet
    excel_df = conformes.copy()

    output = BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet('Résultats')

    # Formats
    wrap_fmt = workbook.add_format({'text_wrap': True, 'valign': 'top'})
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3'})
    red_fmt = workbook.add_format({'font_color': '#FF0000', 'text_wrap': True})

    # En-têtes
    for i, col in enumerate(excel_df.columns):
        worksheet.write(0, i, col, header_fmt)
        worksheet.set_column(i, i, 30)

    # Lignes
    for r, row in enumerate(excel_df.itertuples(index=False), start=1):
        for c, val in enumerate(row):
            txt = str(val)
            col_name = excel_df.columns[c]
            if col_name == 'resume':
                worksheet.write_url(r, c, txt, string='Résumé', cell_format=red_fmt)
            elif bool(pattern.search(txt)):
                worksheet.write(r, c, txt, red_fmt)
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


















































