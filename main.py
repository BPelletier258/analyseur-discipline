import pandas as pd
import re
import unicodedata
import logging
from flask import Flask, request, render_template
from io import BytesIO
import xlsxwriter

# Configuration du logger pour voir les debug dans les logs Render
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
    uploaded = request.files.get('file') or request.files.get('fichier_excel')
    article = request.form.get('article', '').strip()
    if not uploaded or not article:
        return render_template('index.html', erreur="Veuillez fournir un fichier Excel et un article.")

    # Lecture et normalisation des colonnes
    df = pd.read_excel(uploaded)
    df.columns = [normalize_column(c) for c in df.columns]
    # Supprimer les colonnes vides ou parasites
    df = df.drop(columns=[c for c in df.columns if c.startswith('unnamed') or c == 'resume'], errors='ignore')

    # Colonnes obligatoires
    required = [
        'numero de decision', 'nom de lintime', 'articles enfreints',
        'duree totale effective radiation', 'article amende/chef', 'autres sanctions'
    ]
    missing = [c for c in required if c not in df.columns]
    if missing:
        return render_template('index.html', erreur=f"Colonnes manquantes : {', '.join(missing)}")

    # Debug avant filtre
    logger.debug("Total lignes avant filtrage: %d", len(df))

    # Regex strict pour l'article
    pattern = re.compile(rf'(?<![\d\.])Art[\.:]?\s*{re.escape(article)}(?=[\W]|$)', re.IGNORECASE)
    def match_art(cell):
        text = unicodedata.normalize('NFKD', str(cell))
        return bool(pattern.search(text))

    # Filtrage précis
    df_filtered = df[df['articles enfreints'].apply(match_art)].reset_index(drop=True)
    logger.debug("Lignes retenues: %d", len(df_filtered))
    logger.debug("Décisions filtrées: %s", df_filtered['numero de decision'].tolist())
    if df_filtered.empty:
        return render_template('index.html', erreur=f"Aucun résultat pour l'article {article}.")

    # Génération du Markdown (colonnes requises)
    md_df = df_filtered[required]
    markdown_table = md_df.to_markdown(index=False)

    # Préparation du fichier Excel (colonnes requises)
    excel_df = df_filtered[required]

    output = BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})
    ws = wb.add_worksheet('Résultats')

    # Formats
    wrap_fmt = wb.add_format({'text_wrap': True, 'valign': 'top'})
    hdr_fmt  = wb.add_format({'bold': True, 'bg_color': '#D3D3D3'})
    red_fmt  = wb.add_format({'font_color': '#FF0000', 'text_wrap': True, 'valign': 'top'})

    # Entêtes
    for idx, col in enumerate(excel_df.columns):
        ws.write(0, idx, col, hdr_fmt)
        ws.set_column(idx, idx, 30)

    # Valeurs
    for r, row in enumerate(excel_df.itertuples(index=False), start=1):
        for c, val in enumerate(row):
            txt = str(val)
            fmt = red_fmt if match_art(txt) else wrap_fmt
            ws.write(r, c, txt, fmt)

    wb.close()
    output.seek(0)

    return render_template(
        'resultats.html',
        table_markdown=markdown_table,
        fichier_excel=output.read(),
        filename=f"resultats_article_{article}.xlsx"
    )

if __name__ == '__main__':
    app.run(debug=True)





























































