
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
target_article = r"Art\.\s*14"
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
    files = glob.glob("*.xls*")
    if not files:
        raise FileNotFoundError("Aucun fichier Excel trouvé dans le dossier courant.")
    input_file = files[0]

    # lecture
    try:
        df = pd.read_excel(input_file)
    except Exception:
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

    # filtrage article
    mask = df['articles enfreints'].astype(str).str.contains(target_article)
    filtered = df.loc[mask].copy()

    # résumé
    if 'resume' in filtered.columns:
        filtered['Résumé'] = filtered['resume'].apply(lambda x: 'Résumé' if pd.notna(x) else '')
        filtered.drop(columns=['resume'], inplace=True)
    else:
        filtered['Résumé'] = ''

    # écriture Excel
    wb = Workbook()
    ws = wb.active
    ws.title = 'Article_14'
    for r in dataframe_to_rows(filtered, index=False, header=True):
        ws.append(r)
    # styles
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal='center')
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            if isinstance(cell.value, str) and re.search(target_article, cell.value):
                cell.font = Font(color='FF0000')
            cell.alignment = Alignment(wrap_text=True)
    for col in ws.columns:
        length = max(len(str(cell.value)) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = min(length+2, 40)
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
    # exécution locale
    try:
        file_created = process_excel()
        print(f"🎉 Fichier Excel généré : {file_created}")
    except Exception as err:
        print(f"❌ Erreur : {err}")
        sys.exit(1)














































































