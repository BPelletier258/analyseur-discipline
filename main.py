
import pandas as pd
import re
import unicodedata
import logging
from flask import Flask, request, render_template
from io import BytesIO
import xlsxwriter

# Logger pour debugging\logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

app = Flask(__name__)

def normalize_column(col_name):
    if isinstance(col_name, str):
        col = unicodedata.normalize('NFKD', col_name)
        col = col.encode('ASCII', 'ignore').decode('utf-8')
        col = col.replace("’", "'")
        col = re.sub(r"\s+", " ", col).strip().lower()
        return col
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

    # Lecture et normalisation des colonnes
    df = pd.read_excel(uploaded)
    df.columns = [normalize_column(c) for c in df.columns]
    # Suppression de toute colonne contenant "resume" ou "unnamed"
    df = df.loc[:, ~df.columns.str.contains(r'resum|unnamed', case=False)]

    # Rename pour cohérence
    df = df.rename(columns={'nom de lintime': "nom de l'intime"})

    # Vérification des colonnes essentielles
    required = [
        'numero de decision', "nom de l'intime", 'articles enfreints',
        'duree totale effective radiation', 'article amende/chef', 'autres sanctions'
    ]
    missing = [c for c in required if c not in df.columns]
    if missing:
        return render_template('index.html', erreur=f"Colonnes manquantes : {', '.join(missing)}")

    logger.debug("Total lignes avant filtre : %d", len(df))

    # Pattern strict pour l'article (évite 114, 149.1, etc.)
    pattern = re.compile(rf'(?<![\d\.])Art\.?\s*{re.escape(article)}(?=\D|$)', re.IGNORECASE)
    mask = df['articles enfreints'].astype(str).apply(lambda cell: bool(pattern.search(unicodedata.normalize('NFKD', cell))))
    filtered = df.loc[mask, required].reset_index(drop=True)
    logger.debug("Lignes après filtre : %d", len(filtered))
    if filtered.empty:
        return render_template('index.html', erreur=f"Aucun résultat pour l'article {article}.")

    # Génération Markdown (GitHub)
    markdown_table = filtered.to_markdown(index=False, tablefmt='github')

    # Préparation du fichier Excel
    output = BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})
    ws = wb.add_worksheet('Résultats')

    # Formats
    header_fmt = wb.add_format({'bold': True, 'bg_color': '#D3D3D3', 'text_wrap': True, 'align': 'center'})
    wrap_fmt = wb.add_format({'text_wrap': True, 'valign': 'top'})
    red_fmt = wb.add_format({'font_color': '#FF0000', 'text_wrap': True, 'valign': 'top'})

    # Écriture en-têtes
    for col_idx, col in enumerate(required):
        ws.write(0, col_idx, col.title(), header_fmt)
        ws.set_column(col_idx, col_idx, 25)

    # Écriture des données filtrées uniquement
    for row_idx, row in filtered.iterrows():
        for col_idx, col in enumerate(required):
            val = row[col]
            # Coloration si l'article présent dans la cellule
            fmt = red_fmt if pattern.search(str(val)) else wrap_fmt
            ws.write(row_idx+1, col_idx, val, fmt)

    wb.close()
    output.seek(0)

    return render_template('resultats.html',
        table_markdown=markdown_table,
        excel_bytes=output.read(),
        filename=f"resultats_article_{article}.xlsx"
    )

if __name__ == '__main__':
    app.run(debug=True)



































































