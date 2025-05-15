import pandas as pd
import re
import unicodedata
import logging
from flask import Flask, request, render_template
from io import BytesIO
import xlsxwriter

# Logger pour debugging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

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
    # Récupération du fichier et de l'article
    uploaded = request.files.get('file') or request.files.get('fichier_excel')
    article = request.form.get('article', '').strip()
    if not uploaded or not article:
        return render_template('index.html', erreur="Veuillez fournir un fichier Excel et un article.")

    # Lecture et nettoyage
    df = pd.read_excel(uploaded)
    df.columns = [normalize_column(c) for c in df.columns]
    # Supprimer colonnes parasites
    df = df.drop(columns=[c for c in df.columns if c.startswith('unnamed')], errors='ignore')

    # Colonnes attendues
    required = [
        'numero de decision', 'nom de lintime', 'articles enfreints',
        'duree totale effective radiation', 'article amende/chef', 'autres sanctions'
    ]
    missing = [c for c in required if c not in df.columns]
    if missing:
        return render_template('index.html', erreur=f"Colonnes manquantes : {', '.join(missing)}")

    logger.debug("Lignes totales avant filtre : %d", len(df))

    # Pattern strict pour article précis
    pattern = re.compile(rf'(?<![\d\.])Art[\.:]?\s*{re.escape(article)}(?=[\W]|$)', re.IGNORECASE)
    def match_art(cell):
        txt = unicodedata.normalize('NFKD', str(cell))
        return bool(pattern.search(txt))

    # Filtrage
    filtered = df[df['articles enfreints'].apply(match_art)].reset_index(drop=True)
    logger.debug("Lignes après filtre : %d", len(filtered))
    if filtered.empty:
        return render_template('index.html', erreur=f"Aucun résultat pour l'article {article}.")

    # Markdown
    md_df = filtered[required]
    markdown_table = md_df.to_markdown(index=False)

    # Préparation Excel
    output = BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})
    ws = wb.add_worksheet('Résultats')

    wrap = wb.add_format({'text_wrap': True, 'valign': 'top'})
    header = wb.add_format({'bold': True, 'bg_color': '#D3D3D3'})
    red_cell = wb.add_format({'font_color': '#FF0000', 'text_wrap': True, 'valign': 'top'})

    # Écriture entêtes
    for col_num, col in enumerate(required):
        ws.write(0, col_num, col, header)
        ws.set_column(col_num, col_num, 30)

    # Écriture données
    for row_num, row in enumerate(filtered[required].values, start=1):
        for col_num, val in enumerate(row):
            fmt = red_cell if match_art(val) else wrap
            ws.write(row_num, col_num, val, fmt)

    wb.close()
    output.seek(0)

    return render_template('resultats.html',
        table_markdown=markdown_table,
        fichier_excel=output.read(),
        filename=f"resultats_article_{article}.xlsx"
    )

if __name__ == '__main__':
    app.run(debug=True)






























































