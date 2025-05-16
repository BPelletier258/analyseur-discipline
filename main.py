import pandas as pd
import re
import unicodedata
import logging
from flask import Flask, request, render_template
from io import BytesIO
import xlsxwriter

# Logger pour debug détaillé
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

app = Flask(__name__)


def normalize_column(name):
    """Standardise un nom de colonne: ascii, minuscules, sans accents."""
    if isinstance(name, str):
        col = unicodedata.normalize('NFKD', name)
        col = col.encode('ASCII', 'ignore').decode('utf-8')
        col = col.replace("’", "'")
        col = re.sub(r"\s+", " ", col)
        col = col.strip().lower()
        logger.debug("Normalized column '%s' -> '%s'", name, col)
        return col
    return name


@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')


@app.route('/analyse', methods=['POST'])
def analyse():
    fichier = request.files.get('file')
    article = request.form.get('article', '').strip()
    if not fichier or not article:
        return render_template('index.html', erreur="Veuillez fournir un fichier Excel et un numéro d'article.")

    df = pd.read_excel(fichier)
    # Normalize columns
    df.columns = [normalize_column(c) for c in df.columns]
    logger.debug("Colonnes après normalisation: %s", df.columns.tolist())

    # Drop unwanted summary or blank columns
    drop_cols = [c for c in df.columns if re.search(r"^unnamed|resume", c)]
    if drop_cols:
        logger.debug("Suppression colonnes: %s", drop_cols)
        df = df.drop(columns=drop_cols)

    # Standardize expected names
    mapping = {
        'nom de lintime': "nom de l'intime",
        'numero de decision': 'numero de decision',
        'ordre professionnel': 'ordre professionnel',
        'articles enfreints': 'articles enfreints',
        'duree totale effective radiation': 'duree totale effective radiation',
        'article amende/chef': 'article amende/chef',
        'total des amendes': 'total des amendes',
        'autres sanctions': 'autres sanctions'
    }
    df = df.rename(columns=mapping)
    logger.debug("Colonnes après renommage: %s", df.columns.tolist())

    req = [
        'numero de decision', "nom de l'intime", 'articles enfreints',
        'duree totale effective radiation', 'article amende/chef',
        'total des amendes', 'autres sanctions'
    ]
    missing = [c for c in req if c not in df.columns]
    if missing:
        logger.error("Colonnes manquantes: %s", missing)
        return render_template('index.html', erreur=f"Colonnes manquantes : {', '.join(missing)}")

    # Filtrage strict de l'article 14
    pattern = re.compile(rf'(?<!\d)Art\.?\s*{re.escape(article)}(?=\D|$)', re.IGNORECASE)
    mask = df['articles enfreints'].fillna('').apply(lambda x: bool(pattern.search(str(x))))
    df_filt = df.loc[mask, req].reset_index(drop=True)
    logger.debug("Lignes après filtrage: %d", len(df_filt))
    if df_filt.empty:
        return render_template('index.html', erreur=f"Aucun résultat pour l'article {article}.")

    # Markdown
    markdown = df_filt.to_markdown(index=False, tablefmt='github')

    # Excel
    output = BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})
    ws = wb.add_worksheet('Résultats')

    fmt_hdr = wb.add_format({'bold': True, 'bg_color': '#F0F0F0', 'align': 'center', 'text_wrap': True})
    fmt_wrap = wb.add_format({'text_wrap': True, 'valign': 'top'})
    fmt_high = wb.add_format({'text_wrap': True, 'valign': 'top', 'font_color': 'red'})

    # Ecrire en-têtes
    for i, col in enumerate(req):
        ws.write(0, i, col.title(), fmt_hdr)
        ws.set_column(i, i, 25)

    # Ecrire lignes filtrées
    for r, row in df_filt.iterrows():
        for c, col in enumerate(req):
            val = row[col]
            fmt = fmt_high if pattern.search(str(val)) else fmt_wrap
            ws.write(r+1, c, val, fmt)

    wb.close()
    output.seek(0)

    return render_template(
        'resultats.html',
        table_markdown=markdown,
        excel_bytes=output.read(),
        filename=f"decisions_article_{article}.xlsx"
    )

if __name__ == '__main__':
    app.run(debug=True)







































































