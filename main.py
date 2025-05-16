import re
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows

# --- Paramètres ---
input_file = "Tableau_cumulatif_decisions_disciplinaires_14b_Avril.xls"
output_file = "decisions_article_14_formate.xlsx"
target_article = r"Art\. 14"

# --- Lecture du fichier Excel ---
# On force pandas à lire toutes les feuilles ou la première
try:
    df = pd.read_excel(input_file)
except Exception:
    df = pd.read_excel(input_file, sheet_name=0)

# --- Filtrage strict des décisions contenant l'article 14 ---
mask = df['articles enfreints'].astype(str).str.contains(target_article)
filtered = df[mask].copy()

# --- Préparation du nouveau classeur ---
wb = Workbook()
ws = wb.active
ws.title = 'Décisions Art 14'

# Supprimer la colonne resume longue
if 'resume' in filtered.columns:
    filtered.drop(columns=['resume'], inplace=True)

# Renommer la colonne 'résumé' pour s'assurer de son orthographe
data_cols = [c for c in filtered.columns]
# Enregistrer une colonne de liens intitulée Résumé
filtered['Résumé'] = df['résumé'].fillna("")

# Mise en forme du DataFrame (sous forme de camelcase colonne si besoin)
filtered.columns = [c.strip().replace(' ', '_') for c in filtered.columns]

# --- Export vers Excel avec styles ---
for r in dataframe_to_rows(filtered, index=False, header=True):
    ws.append(r)
# Style en-têtes
for cell in ws[1]:
    cell.font = Font(bold=True)
    cell.alignment = Alignment(horizontal='center')

# Mise en forme conditionnelle : mettre en rouge le texte contenant Art. 14
def highlight_article14(cell):
    if isinstance(cell.value, str) and re.search(target_article, cell.value):
        cell.font = Font(color='FF0000')
        cell.alignment = Alignment(wrap_text=True)

for row in ws.iter_rows(min_row=2):
    for cell in row:
        highlight_article14(cell)

# Ajustement auto de la largeur des colonnes
for col in ws.columns:
    max_length = max((len(str(cell.value)) for cell in col), default=0)
    col_letter = col[0].column_letter
    ws.column_dimensions[col_letter].width = min(max_length + 2, 50)

# Sauvegarde
wb.save(output_file)
print("Fichier Excel généré :", output_file)









































































