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
target_article = r"Art\.\s*14(?=\D|$)"  # filtre strict pour Art. 14
output_file = "decisions_article_14_formate.xlsx"
app = Flask(__name__)

# -----------------------------------------------------------------------------
# UTILS
# -----------------------------------------------------------------------------
def normalize(col):
    text = unicodedata.normalize('NFKD', str(col))
    return ''.join(c for c in text if not unicodedata.combining(c)).lower().strip()

# -----------------------------------------------------------------------------
# TRAITEMENT EXCEL
# -----------------------------------------------------------------------------
def process_excel():
    # trouver le fichier source (.xls ou .xlsx)
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
    df.rename(columns={
        'nom de lintime': "nom de l'intime",
        'numero de decision': 'numero de decision',
        'articles enfreints': 'articles enfreints'
    }, inplace=True)

    # vérification des colonnes obligatoires
    required = ['numero de decision', "nom de l'intime", 'articles enfreints']
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise KeyError(f"Colonnes manquantes : {missing}")

    # filtrage article 14
    mask = df['articles enfreints'].astype(str).str.contains(target_article, regex=True)
    filtered = df.loc[mask].copy()

    # extraire les URLs puis supprimer la colonne brute
    urls = filtered.get('resume', pd.Series(['']*len(filtered))).fillna('')
    filtered.drop(columns=[col for col in ['resume'] if col in filtered.columns], inplace=True)

    # préparer colonne Résumé (hyperlien)
    filtered['Résumé'] = ['Résumé'] * len(filtered)

    # création du classeur
    wb = Workbook()
    ws = wb.active
    ws.title = 'Article_14'

    # écriture des données
    for r in dataframe_to_rows(filtered, index=False, header=True):
        ws.append(r)

    # style en-têtes
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')

    # ajuster largeur colonnes
    for col in ws.columns:
        max_len = max((len(str(c.value)) for c in col), default=0)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 40)

    # mise en forme et hyperliens
    header_index = {cell.value: idx+1 for idx, cell in enumerate(ws[1])}
    for i, row in enumerate(ws.iter_rows(min_row=2), start=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True)
            if isinstance(cell.value, str) and re.search(target_article, cell.value):
                cell.font = Font(color='FF0000')
        # hyperlien
        url = urls.iat[i-2]
        if url:
            link_cell = ws.cell(row=i, column=header_index['Résumé'])
            link_cell.hyperlink = url
            link_cell.style = 'Hyperlink'

    wb.save(output_file)
    return output_file

# -----------------------------------------------------------------------------
# HTTP ENDPOINT
# -----------------------------------------------------------------------------
@app.route('/')
def root():
    return ('<h3>Service Analyse Disciplinaires</h3>'
            '<p>Rendez-vous sur <a href="/generate">/generate</a> pour générer le fichier.</p>')

@app.route('/generate')
def generate():
    try:
        fname = process_excel()
    except Exception as e:
        abort(500, description=str(e))
    return send_file(fname, as_attachment=True)

# -----------------------------------------------------------------------------
# CLI
# -----------------------------------------------------------------------------
if __name__ == '__main__':
    try:
        out = process_excel()
        print(f"🎉 Fichier Excel généré : {out}")
    except Exception as err:
        print(f"❌ Erreur : {err}")
        sys.exit(1)
















































































