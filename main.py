
import pandas as pd
import re
import unicodedata
import logging
from flask import Flask, request, render_template
from io import BytesIO
import xlsxwriter

# Logger et sorties console pour debug
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

app = Flask(__name__)


def normalize_column(name):
    """Standardise les noms de colonnes: unicode -> ascii, minuscules, sans accents."""
    if isinstance(name, str):
        col = unicodedata.normalize('NFKD', name)
        col = col.encode('ASCII', 'ignore').decode('utf-8')
        col = col.replace("’", "'")
        col = re.sub(r"\s+", " ", col).strip().lower()
        logger.debug("Normalized column: %s -> %s", name, col)
        return col
    return name


@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')


@app.route('/analyse', methods=['POST'])
def analyse():
    uploaded = request.files.get('file') or request.files.get('fichier_excel')
    article = request.form.get('article', '').strip()
    if not uploaded or not article:
        return render_template('index.html', erreur="Veuillez fournir un fichier Excel et un article.")

    df = pd.read_excel(uploaded)
    # Normalisation colonnes
    df.columns = [normalize_column(c) for c in df.columns]
    logger.debug("Colonnes brutes: %s", df.columns.tolist())

    # Suppression colonnes de résumé ou non nommées
    drop_cols = [c for c in df.columns if re.search(r'resum|unnamed', c, re.IGNORECASE)]
    if drop_cols:
        logger.debug("Suppression colonnes: %s", drop_cols)
        df.drop(columns=drop_cols, inplace=True)
    logger.debug("Colonnes après drop résumé: %s", df.columns.tolist())

    # Correspondance des noms attendus
    mapping = {
        'nom de lintime': "nom de l'intime",
        'numero de decision': 'numero de decision',
        'ordre professionnel': 'ordre professionnel'
    }
    df.rename(columns=mapping, inplace=True)
    logger.debug("Colonnes après renommage: %s", df.columns.tolist())

    # Vérification présence colonnes essentielles
    required = [
        'numero de decision', "nom de l'intime", 'ordre professionnel',
        'articles enfreints', 'duree totale effective radiation',
        'article amende/chef', 'autres sanctions'
    ]
    missing = [c for c in required if c not in df.columns]
    if missing:
        logger.error("Colonnes manquantes: %s", missing)
        return render_template('index.html', erreur=f"Colonnes manquantes : {', '.join(missing)}")

    logger.debug("Total lignes initiales: %d", len(df))

    # Construction filtre exact sur l'article
    pat = re.compile(rf'(?<![\d\.])Art\.?\s*{re.escape(article)}(?=\D|$)', re.IGNORECASE)
    mask = df['articles enfreints'].astype(str).apply(lambda x: bool(pat.search(x)))
    print("MASK:", mask.tolist())
    logger.debug("Mask values: %s", mask.tolist())

    # Application filtre
    filtered = df.loc[mask, required].reset_index(drop=True)
    print("Filtered rows:", filtered.index.tolist())
    logger.debug("Nombre lignes filtrées: %d", len(filtered))
    if filtered.empty:
        return render_template('index.html', erreur=f"Aucun résultat pour l'article {article}.")

    # Génération Markdown
    md = filtered.to_markdown(index=False, tablefmt='github')
    logger.debug("Markdown généré")

    # Génération Excel annoté
    out = BytesIO()
    wb = xlsxwriter.Workbook(out, {'in_memory': True})
    ws = wb.add_worksheet('Résultats')

    fmt_header = wb.add_format({'bold': True, 'bg_color': '#D3D3D3', 'text_wrap': True, 'align': 'center'})
    fmt_wrap = wb.add_format({'text_wrap': True, 'valign': 'top'})
    fmt_highlight = wb.add_format({'font_color': '#FF0000', 'text_wrap': True, 'valign': 'top'})

    # Écrire en-têtes (sans index)
    for idx, col in enumerate(required):
        ws.write(0, idx, col.title(), fmt_header)
        ws.set_column(idx, idx, 30)

    # Écriture des données
    for r, row in filtered.iterrows():
        for c, col in enumerate(required):
            val = row[col]
            cell_fmt = fmt_highlight if pat.search(str(val)) else fmt_wrap
            ws.write(r+1, c, val, cell_fmt)

    wb.close()
    out.seek(0)

    logger.debug("Fichier Excel prêt avec %d lignes", len(filtered))
    return render_template('resultats.html',
        table_markdown=md,
        excel_bytes=out.read(),
        filename=f"resultats_article_{article}.xlsx"
    )


if __name__ == '__main__':
    app.run(debug=True)






































































