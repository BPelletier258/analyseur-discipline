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

    # Lecture et nettoyage des colonnes
    df = pd.read_excel(file)
    df.columns = [normalize_column(c) for c in df.columns]
    df = df.loc[:, [c for c in df.columns if c and not c.startswith('unnamed')]]

    # Colonnes obligatoires
    required = [
        'numero de decision', 'nom de lintime', 'articles enfreints',
        'duree totale effective radiation', 'article amende/chef', 'autres sanctions'
    ]
    missing = [c for c in required if c not in df.columns]
    if missing:
        return render_template('index.html', erreur=f"Colonnes manquantes : {', '.join(missing)}")

    # Regex strict pour l'article
    pattern = re.compile(
        rf'(?<![\d\.])Art[\.:]?\s*{re.escape(article)}(?=[\W]|$)',
        re.IGNORECASE
    )

    def match_art(x):
        txt = unicodedata.normalize('NFKD', str(x))
        return bool(pattern.search(txt))

    # Filtrage des décisions
    df_filtered = df[df['articles enfreints'].apply(match_art)].reset_index(drop=True)
    if df_filtered.empty:
        return render_template('index.html', erreur=f"Aucun résultat pour l'article {article}.")

    # Génération du Markdown
    md_df = df_filtered[required].copy().reset_index(drop=True)
    markdown_table = md_df.to_markdown(index=False)

    # Préparation du fichier Excel
    cols = required + (['resume'] if 'resume' in df_filtered.columns else [])
    excel_df = df_filtered[cols].copy().reset_index(drop=True)

    output = BytesIO()
    wb = xlsxwriter.Workbook(output, {'in_memory': True})
    ws = wb.add_worksheet('Résultats')

    # Formats
    fmt_wrap = wb.add_format({'text_wrap': True, 'valign': 'top'})
    fmt_hdr = wb.add_format({'bold': True, 'bg_color': '#D3D3D3'})
    fmt_red = wb.add_format({'font_color': '#FF0000', 'text_wrap': True, 'valign': 'top'})

    # Écriture des en-têtes
    for i, col in enumerate(excel_df.columns):
        ws.write(0, i, col, fmt_hdr)
        ws.set_column(i, i, 30)

    # Écriture des données
    for r, row in enumerate(excel_df.itertuples(index=False), start=1):
        for c, val in enumerate(row):
            txt = str(val)
            col_name = excel_df.columns[c]
            if col_name == 'resume':
                ws.write_url(r, c, txt, string='Résumé', cell_format=fmt_wrap)
            elif col_name in required and match_art(txt):
                ws.write(r, c, txt, fmt_red)
            else:
                ws.write(r, c, txt, fmt_wrap)

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


























































