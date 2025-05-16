import re
import glob
import sys
import pandas as pd
import unicodedata
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
from flask import Flask, send_file, abort

# -----------------------------------------------------------------------------
# CONFIGURATION
# -----------------------------------------------------------------------------
target_article = r"Art[.:]\s*14(?=\D|$)"  # filtre strict pour Art. 14
output_file = "decisions_article_14_formate.xlsx"
app = Flask(__name__)

# -----------------------------------------------------------------------------
# UTILS
# -----------------------------------------------------------------------------
def normalize(col):
    text = unicodedata.normalize('NFKD', str(col))
    return ''.join(c for c in text if not unicodedata.combining(c)).lower().strip()

# -----------------------------------------------------------------------------
# EXTRACTION & FILTRAGE
# -----------------------------------------------------------------------------
def get_filtered_df():
    # trouver le fichier Excel
    files = glob.glob("*.xls*")
    if not files:
        raise FileNotFoundError("Aucun fichier Excel trouvé dans le dossier courant.")
    input_file = files[0]

    # lecture
    try:
        df = pd.read_excel(input_file)
    except:
        df = pd.read_excel(input_file, sheet_name=0)

    # normalisation des colonnes
    df.columns = [normalize(c) for c in df.columns]

    # renommage tolérant pour nom de l'intimé
    if 'nom de lintime' in df.columns:
        df.rename(columns={'nom de lintime': "nom de l'intime"}, inplace=True)

    # renommage tolérant pour articles enfreints: détecter variante
    for col in df.columns:
        if 'article' in col and 'enf' in col:
            df.rename(columns={col: 'articles enfreints'}, inplace=True)
            break

    required = ['numero de decision', "nom de l'intime", 'articles enfreints']
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise KeyError(f"Colonnes manquantes : {missing}")

    # filtrage strict
    mask = df['articles enfreints'].astype(str).str.contains(target_article, regex=True)
    filtered = df.loc[mask].copy()

    # extraction et remplacement de la colonne résumé
    urls = filtered.get('resume', pd.Series([''] * len(filtered))).fillna('')
    if 'resume' in filtered.columns:
        filtered.drop(columns=['resume'], inplace=True)
    filtered['Résumé'] = ['Résumé'] * len(filtered)

    return filtered, urls

# -----------------------------------------------------------------------------
# CRÉATION DU FICHIER EXCEL
# -----------------------------------------------------------------------------
def build_excel(filtered, urls):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Article_14'

    # écrire les données
    for r in dataframe_to_rows(filtered, index=False, header=True):
        ws.append(r)

    # style des en-têtes
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')

    # ajuster largeur
    for col in ws.columns:
        max_len = max((len(str(c.value)) for c in col), default=0)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 40)

    # mise en forme des cellules
    header_index = {cell.value: idx + 1 for idx, cell in enumerate(ws[1])}
    for i, row in enumerate(ws.iter_rows(min_row=2), start=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True)
            if isinstance(cell.value, str) and re.search(target_article, cell.value):
                cell.font = Font(color='FF0000')
        # créer lien hypertexte
        url = urls.iat[i - 2]
        if url:
            link_cell = ws.cell(row=i, column=header_index['Résumé'])
            link_cell.hyperlink = url
            link_cell.style = 'Hyperlink'

    wb.save(output_file)

# -----------------------------------------------------------------------------
# ENDPOINTS HTTP
# -----------------------------------------------------------------------------
@app.route('/')
def root():
    return ('<h3>Analyse Disciplinaire</h3>'
            '<p>/generate &lt;renvoie le fichier Excel&gt; et /table &lt;renvoie le Markdown&gt;</p>')

@app.route('/generate')
def generate():
    try:
        filtered, urls = get_filtered_df()
        build_excel(filtered, urls)
    except Exception as e:
        abort(500, description=str(e))
    return send_file(output_file, as_attachment=True)

@app.route('/table')
def table():
    try:
        filtered, _ = get_filtered_df()
    except Exception as e:
        abort(500, description=str(e))
    return '<pre>' + filtered.to_markdown(index=False) + '</pre>'

# -----------------------------------------------------------------------------
# MODE CLI
# -----------------------------------------------------------------------------
if __name__ == '__main__':
    try:
        filtered, urls = get_filtered_df()
        print(filtered.to_markdown(index=False))
        build_excel(filtered, urls)
        print(f"🎉 Fichier généré : {output_file}")
    except Exception as err:
        print(f"❌ Erreur : {err}")
        sys.exit(1)




















































































