import re
import glob
import sys
import pandas as pd
import unicodedata
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows
from flask import Flask, send_file, abort

# --------------------------------------------------------------------------------
# CONFIGURATION
# --------------------------------------------------------------------------------
target_article = r"Art\.\s*14(?=\D|$)"  # strict match 14, not 149.1
output_file = "decisions_article_14_formate.xlsx"

app = Flask(__name__)

# --------------------------------------------------------------------------------
# FONCTIONS UTILES
# --------------------------------------------------------------------------------
def normalize(col):
    s = unicodedata.normalize('NFKD', str(col))
    s = ''.join(ch for ch in s if not unicodedata.combining(ch))
    return s.lower().strip()


def process_excel():
    # trouver fichier source
    files = glob.glob("*.xls*")
    if not files:
        raise FileNotFoundError("Aucun fichier Excel trouvé dans le dossier courant.")
    input_file = files[0]

    # lecture
    try:
        df = pd.read_excel(input_file)
    except:
        df = pd.read_excel(input_file, sheet_name=0)

    # normalisation noms colonnes
    df.columns = [normalize(c) for c in df.columns]
    df.rename(columns={
        'nom de lintime': "nom de l'intime",
        'numero de decision': 'numero de decision',
        'articles enfreints': 'articles enfreints'
    }, inplace=True)

    # vérification colonnes
    req = ["numero de decision", "nom de l'intime", "articles enfreints"]
    missing = [c for c in req if c not in df.columns]
    if missing:
        raise KeyError(f"Colonnes manquantes : {missing}")

    # filtrage article 14 strict
    mask = df['articles enfreints'].astype(str).str.contains(target_article, regex=True)
    filtered = df.loc[mask].copy()

    # création colonne Résumé et conserver URL
    urls = None
    if 'resume' in filtered.columns:
        urls = filtered['resume'].fillna("")
    filtered['Résumé'] = ['Résumé'] * len(filtered)

    # construction Excel
    wb = Workbook()
    ws = wb.active
    ws.title = 'Article_14'
    # écrire en-tête + données
    for r in dataframe_to_rows(filtered, index=False, header=True):
        ws.append(r)

    # style en-tête
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')
    # ajuster largeur
    for col in ws.columns:
        length = max(len(str(cell.value)) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = min(length+2, 40)

    # mise en forme et hyperliens
    col_index = {cell.value: idx+1 for idx, cell in enumerate(ws[1])}
    for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=0):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True)
            if isinstance(cell.value, str) and re.search(target_article, cell.value):
                cell.font = Font(color='FF0000')
        # ajouter hyperlink pour Résumé
        if urls is not None:
            url = urls.iloc[row_idx]
            if url:
                c = ws.cell(row=row_idx+2, column=col_index['Résumé'])
                c.hyperlink = url
                c.style = 'Hyperlink'

    wb.save(output_file)
    return output_file

# --------------------------------------------------------------------------------
# ENDPOINT HTTP
# --------------------------------------------------------------------------------
@app.route('/')
def root():
    return ("<h3>Service Analyse Disciplinaires</h3>"
            "<p>Accédez à <a href=\"/generate\">/generate</a> pour lancer le traitement Excel.</p>")

@app.route('/generate')
def generate():
    try:
        xlsx = process_excel()
    except Exception as e:
        abort(500, description=str(e))
    return send_file(xlsx, as_attachment=True)

# --------------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------------
if __name__ == '__main__':
    try:
        file_created = process_excel()
        print(f"🎉 Fichier Excel généré : {file_created}")
    except Exception as err:
        print(f"❌ Erreur : {err}")
        sys.exit(1)















































































