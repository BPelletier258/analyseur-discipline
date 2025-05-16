import re
import glob
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows

# --- Paramètres ---
target_article = r"Art\.\s*14"
output_file = "decisions_article_14_formate.xlsx"

# --- Détection automatique du fichier Excel ---
# Cherche un fichier .xls ou .xlsx dans le répertoire courant
files = glob.glob("*.xls*")
if not files:
    raise FileNotFoundError("Aucun fichier Excel trouvé dans le dossier courant.")
input_file = files[0]
print(f"Fichier d'entrée détecté : {input_file}")

# --- Lecture du contenu ---
# Lecture sécurisée de la première feuille
try:
    df = pd.read_excel(input_file)
except Exception:
    df = pd.read_excel(input_file, sheet_name=0)

# --- Vérification des colonnes obligatoires ---
required = ["numero de decision", "nom de lintime", "articles enfreints"]
missing = [c for c in required if c not in df.columns]
if missing:
    raise KeyError(f"Colonnes manquantes : {missing}")

# --- Filtrage strict des décisions contenant l'article 14 ---
mask = df['articles enfreints'].astype(str).str.contains(target_article)
filtered = df.loc[mask, :].copy()
print(f"Décisions trouvées pour l'article 14 : {len(filtered)}")

# --- Ajout de la colonne Résumé et suppression des liens bruts ---
if 'résumé' in filtered.columns:
    filtered['Résumé'] = filtered['résumé'].apply(lambda x: 'Résumé' if pd.notna(x) else '')
    filtered.drop(columns=['résumé'], inplace=True)
else:
    filtered['Résumé'] = ''

# --- Normalisation des noms de colonnes pour Excel ---
filtered.columns = [c.strip().replace(' ', '_') for c in filtered.columns]

# --- Export vers Excel avec stylisation ---
wb = Workbook()
ws = wb.active
ws.title = 'Art_14'

# Remplissage
for r in dataframe_to_rows(filtered, index=False, header=True):
    ws.append(r)

# Style en-têtes
for cell in ws[1]:
    cell.font = Font(bold=True)
    cell.alignment = Alignment(horizontal='center')

# Mise en forme conditionnelle et retour à la ligne
for row in ws.iter_rows(min_row=2):
    for cell in row:
        if isinstance(cell.value, str) and re.search(target_article, cell.value):
            cell.font = Font(color='FF0000')
        cell.alignment = Alignment(wrap_text=True)

# Ajustement automatique des colonnes
for col in ws.columns:
    max_len = max(len(str(cell.value)) for cell in col)
    ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 40)

# Sauvegarde
wb.save(output_file)
print(f"Fichier Excel généré : {output_file}")










































































