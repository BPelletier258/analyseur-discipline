import re
import glob
import sys
import pandas as pd
import unicodedata
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils.dataframe import dataframe_to_rows

# --------------------------------------------------------------------------------
# INSTRUCTIONS POUR LE DÉPLOIEMENT
# --------------------------------------------------------------------------------
# 1. AVANT de lancer le déploiement (manual deploy) sur Render :
#    • Assurez-vous que votre fichier Excel « Tableau_cumulatif_decisions_disciplinaires.xls(x) »
#      est engagé dans le dépôt à la racine (à côté de main.py), ou
#      téléversé via GitHub UI (Add file → Upload file), puis commit & push.
#    • Dans Render, faites ensuite un Manual Deploy : le build prendra en compte
#      automatiquement tous les fichiers du repo, y compris l’Excel.
# 2. Après déploiement réussi, appelez l’endpoint ou exécutez localement :
#      > python src/main.py
# --------------------------------------------------------------------------------

# --- Paramètres ---
target_article = r"Art\.\s*14"
output_file = "decisions_article_14_formate.xlsx"

# --- Détection automatique du fichier Excel ---
files = glob.glob("*.xls*")
if not files:
    print("❌ Erreur : Aucun fichier Excel trouvé dans le dossier courant.")
    print("➡️ Vérifiez que vous avez bien engagé ou téléversé le fichier via GitHub avant le déploiement.")
    print("   Exemple de nom attendu : Tableau_cumulatif_decisions_disciplinaires.xls")
    sys.exit(1)
input_file = files[0]
print(f"📁 Fichier d'entrée détecté : {input_file}")

# --- Lecture du fichier ---
try:
    df = pd.read_excel(input_file)
except Exception:
    df = pd.read_excel(input_file, sheet_name=0)

# --- Nettoyage des colonnes vides et normalisation des noms ---
def normalize(col):
    s = unicodedata.normalize('NFKD', str(col))
    s = ''.join(ch for ch in s if not unicodedata.combining(ch))
    return s.lower().strip()
df.columns = [normalize(c) for c in df.columns]

# --- Renommage ponctuel ---
col_map = {
    'nom de lintime': "nom de l'intime",
    'numero de decision': 'numero de decision',
    'articles enfreints': 'articles enfreints'
}
df.rename(columns=col_map, inplace=True)

# --- Vérification des colonnes obligatoires ---
required = ["numero de decision", "nom de l'intime", "articles enfreints"]
missing = [c for c in required if c not in df.columns]
if missing:
    print(f"❌ Colonnes manquantes : {missing}")
    sys.exit(1)

# --- Filtrage strict pour l'article 14 ---
mask = df['articles enfreints'].astype(str).str.contains(target_article, regex=True)
filtered = df.loc[mask].copy()
print(f"✅ Décisions filtrées pour l'article 14 : {len(filtered)}")

# --- Colonne 'Résumé' ---
if 'resume' in filtered.columns:
    filtered['Résumé'] = filtered['resume'].apply(lambda x: 'Résumé' if pd.notna(x) else '')
    filtered.drop(columns=['resume'], inplace=True)
else:
    filtered['Résumé'] = ''

# --- Export Excel avec style ---
wb = Workbook()
ws = wb.active
ws.title = 'Article_14'

# Ajout des données
for r in dataframe_to_rows(filtered, index=False, header=True):
    ws.append(r)

# Style en-têtes
title_row = next(ws.iter_rows(min_row=1, max_row=1))
for cell in title_row:
    cell.font = Font(bold=True)
    cell.alignment = Alignment(horizontal='center')

# Surlignage de l'article et retour à la ligne
for row in ws.iter_rows(min_row=2):
    for cell in row:
        if isinstance(cell.value, str) and re.search(target_article, cell.value):
            cell.font = Font(color='FF0000')
        cell.alignment = Alignment(wrap_text=True)

# Ajustement largeur colonnes
for col in ws.columns:
    max_length = max(len(str(cell.value)) for cell in col)
    ws.column_dimensions[col[0].column_letter].width = min(max_length+2, 40)

# Sauvegarde
try:
    wb.save(output_file)
    print(f"🎉 Fichier Excel généré : {output_file}")
except Exception as e:
    print(f"❌ Erreur lors de la sauvegarde : {e}")
    sys.exit(1)













































































