import pandas as pd
import re
import unicodedata
from flask import Flask, request, render_template
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
    # style span without altering text color
    pattern = re.compile(rf'(Art[\.:]?\s*{re.escape(article)})(?=[\s\W]|$)', re.IGNORECASE)
    return pattern.sub(r'<span style="background-color:#FFC7CE; font-weight:bold">\1</span>', cell)

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/analyse', methods=['POST'])
def analyse():
    file = request.files.get('file') or request.files.get('fichier_excel')
    article = request.form.get('article', '').strip()

    if not file or not article:
        return render_template('index.html', erreur="Veuillez fournir un fichier Excel et un article.")

    df = pd.read_excel(file)
    # normaliser entêtes et supprimer colonnes non nommées
    df.columns = [normalize_column(c) for c in df.columns]
    df = df.loc[:, [c for c in df.columns if not c.startswith('unnamed')]]

    required = [
        "numero de decision",
        "nom de l'intime",
        "articles enfreints",
        "duree totale effective radiation",
        "article amende/chef",
        "autres sanctions"
    ]
    missing = [col for col in required if col not in df.columns]
    if missing:
        return render_template('index.html', erreur=f"Le fichier est incomplet. Colonnes manquantes : {', '.join(missing)}")

    # recherche précise de l'article
    pattern_explicit = rf'Art[\.:]?\s*{re.escape(article)}(?=[\s\W]|$)'
    mask = df['articles enfreints'].astype(str).str.contains(pattern_explicit, na=False, flags=re.IGNORECASE)
    conformes = df.loc[mask, :].copy()

    if conformes.empty:
        return render_template('index.html', erreur=f"Aucun intime trouvé pour l'article {article} demandé.")

    # préparation Markdown
    display_df = conformes[required].copy().reset_index(drop=True)
    markdown_table = display_df.to_markdown(index=False)

    # préparation Excel
    excel_columns = required.copy()
    if 'resume' in df.columns:
        excel_columns.append('resume')
    excel_df = conformes[excel_columns].copy().reset_index(drop=True)

    output = BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet('Résultats')

    # formats
    wrap = workbook.add_format({'text_wrap': True, 'valign': 'top'})
    header = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3'})
    red_bg = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#000000', 'text_wrap': True})

    # écriture en-têtes et largeur
    for idx, col in enumerate(excel_df.columns):
        worksheet.write(0, idx, col, header)
        width = 60 if 2 <= idx <= 5 else 30
        worksheet.set_column(idx, idx, width)

    # écriture lignes
    for r, row in enumerate(excel_df.itertuples(index=False), start=1):
        for c, val in enumerate(row):
            txt = str(val)
            if re.search(pattern_explicit, unicodedata.normalize('NFKD', txt), flags=re.IGNORECASE):
                if df.columns[c] == 'resume':
                    # lien hypertexte uniforme
                    worksheet.write_url(r, c, txt, string='Résumé', cell_format=red_bg)
                else:
                    worksheet.write(r, c, txt, red_bg)
            else:
                worksheet.write(r, c, txt, wrap)

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













































