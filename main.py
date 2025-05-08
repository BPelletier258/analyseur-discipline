import pandas as pd
import re
import unicodedata
from io import BytesIO

def normalize\_column(col\_name):
if isinstance(col\_name, str):
col\_name = unicodedata.normalize('NFKD', col\_name).encode('ASCII', 'ignore').decode('utf-8')
col\_name = col\_name.lower().strip()
col\_name = col\_name.replace("’", "'")
col\_name = re.sub(r'\s+', ' ', col\_name)
return col\_name

def highlight\_article(text, article):
pattern = rf'(Art\[.:]?\s\*{re.escape(article)}(?=\[\s\W]|\$))'
return re.sub(pattern, r'**\1**', text, flags=re.IGNORECASE)

def build\_markdown\_table(df, article):
headers = \["Statut", "Numéro de décision", "Nom de l'intime", "Articles enfreints",
"Périodes de radiation", "Amendes", "Autres sanctions", "Résumé"]
rows = \[]
for \_, row in df.iterrows():
resume\_link = row\.get('resume', '')
resume\_md = f"[Résumé]({resume_link})" if pd.notna(resume\_link) and resume\_link else ""
ligne = \[
str(row\.get("statut", "")),
str(row\.get("numero de decision", "")),
str(row\.get("nom de l'intime", "")),
highlight\_article(str(row\.get("articles enfreints", "")), article),
highlight\_article(str(row\.get("duree totale effective radiation", "")), article),
highlight\_article(str(row\.get("article amende/chef", "")), article),
highlight\_article(str(row\.get("autres sanctions", "")), article),
resume\_md
]
rows.append("| " + " | ".join(ligne) + " |")

```
header_row = "| " + " | ".join(headers) + " |"
separator = "|" + " --- |" * len(headers)
return "\n".join([header_row, separator] + rows)
```

def build\_excel\_result(df, article):
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe\_to\_rows
from openpyxl.styles import Font, Alignment, PatternFill

```
output = BytesIO()
wb = Workbook()
ws = wb.active
ws.title = "Résultats"

df_copy = df.copy()

# Préparer les liens
for i, row in df_copy.iterrows():
    lien = row.get('resume', '')
    if pd.notna(lien) and lien:
        df_copy.at[i, 'resume'] = f'=HYPERLINK("{lien}", "Résumé")'
    else:
        df_copy.at[i, 'resume'] = ''

# Réorganiser les colonnes
ordered_columns = [
    'statut', 'numero de decision', "nom de l'intime", 'articles enfreints',
    'duree totale effective radiation', 'article amende/chef', 'autres sanctions', 'resume'
]
df_copy = df_copy[ordered_columns]

for r in dataframe_to_rows(df_copy, index=False, header=True):
    ws.append(r)

# Style de l'en-tête
header_font = Font(bold=True)
fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
for cell in ws[1]:
    cell.font = header_font
    cell.fill = fill

# Appliquer style général
for row in ws.iter_rows():
    for cell in row:
        cell.alignment = Alignment(wrap_text=True)

for col in ws.columns:
    col_letter = col[0].column_letter
    ws.column_dimensions[col_letter].width = 30

wb.save(output)
output.seek(0)
return output
```

def analyse\_article(df, article):
required\_cols = \[
'articles enfreints', 'duree totale effective radiation', 'article amende/chef',
'autres sanctions', "nom de l'intime", 'numero de decision'
]
for col in required\_cols:
if col not in df.columns:
raise ValueError("Erreur : Le fichier est incomplet. Merci de vérifier la structure.")

```
pattern_explicit = rf'Art[\.:]?\s*{re.escape(article)}(?=[\s\W]|$)'
mask = df['articles enfreints'].astype(str).str.contains(pattern_explicit, na=False, flags=re.IGNORECASE)
result = df[mask].copy()

if result.empty:
    raise ValueError(f"Erreur : Aucun intime trouvé pour l'article {article} demandé.")

result['statut'] = 'Conforme'
return result
```




























