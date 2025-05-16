
import re
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment

# --- Lecture du fichier Excel ---
input_file = "input.xlsx"
df = pd.read_excel(input_file)

# --- Standardisation des noms de colonnes ---
# On accepte 'nom_de_lintime' ou 'nom de lintime' comme équivalent à 'nom de l'intime'
col_map = {
    'nom de lintime': 'nom_de_lintime',
    'nom_de_lintime': 'nom_de_lintime'
}
# Renommer les colonnes si nécessaire
df.rename(columns={k: v for k, v in col_map.items() if k in df.columns}, inplace=True)

# --- Vérification rapide des colonnes obligatoires ---
required = ['numero_de_decision', 'nom_de_lintime', 'articles_enfreints',
            'duree_totale_effective_radiation', 'article_amende/chef', 'autres sanctions']
missing = [c for c in required if c not in df.columns]
if missing:
    raise ValueError(f"Colonnes manquantes : {missing}")

# --- Filtrage strict de l'article 14 ---
pattern = r"\bArt\.? *14\b"
mask = df['articles_enfreints'].astype(str).str.contains(pattern, regex=True)
filtered = df.loc[mask].copy()

# --- Suppression des colonnes superflues ---
# Colonne de résumé brute et colonnes 'Unnamed'
drop_cols = [c for c in filtered.columns if c.lower().startswith('unnamed') or c.lower() == 'resume']
filtered.drop(columns=drop_cols, inplace=True)

# --- Export temporaire pour mise en forme ---
temp_file = "filtered_temp.xlsx"
filtered.to_excel(temp_file, index=False)

# --- Mise en forme Excel ---
wb = load_workbook(temp_file)
ws = wb.active

for row in ws.iter_rows(min_row=2, min_col=1, max_col=ws.max_column):
    for cell in row:
        # Retour à la ligne auto
        cell.alignment = Alignment(wrap_text=True)
        # Rouge si 'Art. 14' présent
        if cell.value and re.search(pattern, str(cell.value)):
            cell.font = Font(color="FF0000")

# --- Sauvegarde finale ---
output_file = "decisions_article14_formate.xlsx"
wb.save(output_file)
print(f"Fichier généré : {output_file}")








































































