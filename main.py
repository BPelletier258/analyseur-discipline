import re
import glob
import sys
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows

# --------------------------------------------------------------------------------
# INSTRUCTIONS POUR LE DÉPLOIEMENT
# --------------------------------------------------------------------------------
# 1. AVANT de lancer le déploiement (manual deploy) sur Render :
#    • Ouvrez l'interface de déploiement de votre service Render.
#    • Dans la section "Files / Static", ajoutez votre fichier Excel source
#      (p. ex. "Tableau_cumulatif_decisions_disciplinaires.xls" ou ".xlsx").
#    • Vérifiez que le fichier apparaisse bien à la racine du dossier de votre service.
# 2. Lancez ensuite le manual deploy : Render récupérera automatiquement ce fichier.
# 3. Enfin, appelez l’endpoint ou exécutez localement main.py :
#      > python src/main.py
# --------------------------------------------------------------------------------

# --- Paramètres ---
target_article = r"Art\.\s*14"
output_file = "decisions_article_14_formate.xlsx"

# --- Détection automatique du fichier Excel ---
files = glob.glob("*.xls*")
if not files:
    print("❌ Erreur : Aucun fichier Excel trouvé dans le dossier courant.")
    print("➡️ Assurez-vous de téléverser votre fichier via l'UI Render AVANT le déploiement.")
    print("   Exemple : 'Tableau_cumulatif_decisions_disciplinaires.xls'")
    sys.exit(1)
input_file = files[0]
print(f"📁 Fichier d'entrée détecté : {input_file}")

# --- Lecture du fichier ---
try:
    df = pd.read_excel(input_file)
except Exception:
    df = pd.read_excel(input_file, sheet_name=0)

# --- Nettoyage des colonnes sans nom ---
df = df.loc[:, [c for c in df.columns if isinstance(c, str) and c.strip()]]

# --- Vérification et renommage des colonnes obligatoires ---
col_map = {'nom de lintime': "nom de l'intime"}

df.rename(columns={k: v for k, v in col_map.items() if k in df.columns}, inplace=True)
required = ["numero de decision", "nom de l'intime", "articles enfreints"]
missing = [c for c in required if c not in df.columns]
if missing:
    print(f"❌ Erreur : Colonnes manquantes -> {missing}")
    sys.exit(1)

# --- Filtrage strict des décisions mentionnant l'article 14 ---
mask = df['articles enfreints'].astype(str).str.contains(target_article, regex=True)
filtered = df.loc[mask].copy()
print(f"✅ Décisions trouvées pour l'article 14 : {len(filtered)}")

# --- Gestion de la colonne 'Résumé' ---
if 'resume' in filtered.columns:
    filtered['Résumé'] = filtered['resume'].apply(lambda x: 'Résumé' if pd.notna(x) else '')
    filtered.drop(columns=['resume'], inplace=True)
else:
    filtered['Résumé'] = ''

# --- Export vers Excel avec mise en forme finale ---
wb = Workbook()
ws = wb.active
ws.title = 'Article_14'

# Ajout des données
for r in dataframe_to_rows(filtered, index=False, header=True):
    ws.append(r)

# Style des en-têtes
for cell in ws[1]:
    cell.font = Font(bold=True)
    cell.alignment = Alignment(horizontal='center')

# Surlignage de l'article 14 en rouge et retour à la ligne
for row in ws.iter_rows(min_row=2):
    for cell in row:
        if isinstance(cell.value, str) and re.search(target_article, cell.value):
            cell.font = Font(color='FF0000')
        cell.alignment = Alignment(wrap_text=True)

# Ajustement automatique de la largeur des colonnes
for col in ws.columns:
    max_length = max(len(str(cell.value)) for cell in col)
    ws.column_dimensions[col[0].column_letter].width = min(max_length + 2, 40)

# Sauvegarde du fichier final
wb.save(output_file)
print(f"🎉 Fichier Excel généré : {output_file}")












































































