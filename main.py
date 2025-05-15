import pandas as pd
import re
import unicodedata
import logging
from flask import Flask, request, render_template
from io import BytesIO
import xlsxwriter

# Logger pour debug
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

app = Flask(__name__)


def normalize_column(name):
    if isinstance(name, str):
        col = unicodedata.normalize('NFKD', name)
        col = col.encode('ASCII', 'ignore').decode('utf-8')
        col = col.replace("’", "'")
        col = re.sub(r"\s+", " ", col).strip().lower()
        return col
    return name


@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')


@app.route('/analyse', methods=['POST'])
def analyse():
    # fichier + article
    uploaded = request.files.get('file') or request.files.get('fichier_excel')
    article = request.form.get('article', '').strip()
    if not uploaded or not article:
        return render_template('index.html', erreur="Veuillez fournir un fichier Excel et un article.")

    # lecture + normalisation
    df = pd.read_excel(uploaded)
    df.columns = [normalize_column(c) for c in df.columns]

    # supprime colonnes inutiles
    drop_mask = df.columns.str.contains(r'resum|unnamed', case=False)
    df = df.loc[:, ~drop_mask]

    # renames
    mapping = {
        'nom de lintime': "nom de l'intime",
        'numero de decision': 'numero de decision',
        'ordre professionnel': 'ordre professionnel'
    }
    df.rename(columns=mapping, inplace=True)

    # colonnes obligatoires
    required = [
        'numero de decision', "nom de l'intime", 'ordre professionnel',
        'articles enfreints', 'duree totale effective radiation',
        'article amende/chef', 'autres sanctions'
    ]
    missing = [c for c in required if c not in df.columns]
    if missing:
        return render_template('index.html', erreur=f"Colonnes manquantes : {', '.join(missing)}")

    logger.debug("Total initial rows: %d", len(df))
    # filtre exact article
    pat = re.compile(rf'(?<![\d\.])Art\.?\s*{re.escape(article)}(?=\D|$)', re.IGNORECASE)
    mask = df['articles enfreints'].astype(str).apply(lambda x: bool(pat.search(x)))
    filtered = df.loc[mask, required].reset_index(drop=True)
    logger.debug("Filtered rows count: %d", len(filtered))
    if filtered.empty:
        return render_template('index.html', erreur=f"Aucun résultat pour l'article {article}.")

    # Markdown unique
    md = filtered.to_markdown(index=False, tablefmt='github')

    # Excel
    out = BytesIO()
    wb = xlsxwriter.Workbook(out, {'in_memory': True})
    ws = wb.add_worksheet('Résultats')

    header = wb.add_format({'bold': True, 'bg_color': '#D3D3D3', 'text_wrap': True, 'align': 'center'})
    wrap = wb.add_format({'text_wrap': True, 'valign': 'top'})
    red = wb.add_format({'font_color': '#FF0000', 'text_wrap': True, 'valign': 'top'})

    # entêtes
    for idx, col in enumerate(required):
        ws.write(0, idx, col.title(), header)
        ws.set_column(idx, idx, 30)

    # écriture
    for r, row in filtered.iterrows():
        for c, col in enumerate(required):
            val = row[col]
            fmt = red if pat.search(str(val)) else wrap
            ws.write(r+1, c, val, fmt)

    wb.close()
    out.seek(0)

    return render_template('resultats.html',
        table_markdown=md,
        excel_bytes=out.read(),
        filename=f"resultats_article_{article}.xlsx"
    )

if __name__ == '__main__':
    app.run(debug=True)




































































