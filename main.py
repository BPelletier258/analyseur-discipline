"""
Analyseur disciplinaire – Flask App principale
"""

import pandas as pd
import re
import unicodedata
from io import BytesIO
from bs4 import BeautifulSoup
from flask import Flask, request, jsonify, send\_file, flash, redirect, render\_template
from werkzeug.utils import secure\_filename
import os

app = Flask(**name**)
app.secret\_key = 'secret'
UPLOAD\_FOLDER = 'uploads'
os.makedirs(UPLOAD\_FOLDER, exist\_ok=True)
app.config\['UPLOAD\_FOLDER'] = UPLOAD\_FOLDER

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

def remove\_html\_tags(text):
if isinstance(text, str):
return BeautifulSoup(text, "html.parser").get\_text()
return text

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
for i, row in df_copy.iterrows():
    lien = row.get('resume', '')
    if pd.notna(lien) and lien:
        df_copy.at[i, 'resume'] = f'=HYPERLINK("{lien}", "Résumé")'
    else:
        df_copy.at[i, 'resume'] = ''

ordered_columns = [
    'statut', 'numero de decision', "nom de l'intime", 'articles enfreints',
    'duree totale effective radiation', 'article amende/chef', 'autres sanctions', 'resume'
]
df_copy = df_copy[ordered_columns]
df_copy = df_copy.dropna(how='all')
for col in df_copy.columns:
    df_copy[col] = df_copy[col].apply(remove_html_tags)

for r in dataframe_to_rows(df_copy, index=False, header=True):
    ws.append(r)

header_font = Font(bold=True)
fill = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
for cell in ws[1]:
    cell.font = header_font
    cell.fill = fill

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

@app.route('/')
def home():
return render\_template('index.html')

@app.route('/analyse', methods=\['POST'])
def analyse():
article = request.form.get("article")
fichier = request.files.get("file")

```
if not article or not fichier:
    flash("Veuillez fournir un article et un fichier Excel.")
    return redirect('/')

try:
    df = pd.read_excel(fichier)
    df = df.rename(columns=lambda c: normalize_column(c))
    result = analyse_article(df, article)

    markdown = build_markdown_table(result, article)
    excel_bytes = build_excel_result(result, article)

    markdown_html = f"""
    <html><head>
    <style>
    body {{ font-family: Arial; line-height: 1.6; }}
    table {{ border-collapse: collapse; width: 100%; }}
    td, th {{ border: 1px solid #ccc; padding: 8px; }}
    </style>
    </head><body>
    <h2>Tableau des sanctions pour l'article {article}</h2>
    {markdown.replace('\n', '<br>')}
    <br><br>
    <form method='get' action='/download'>
    <button type='submit'>Télécharger le fichier Excel</button>
    </form>
    </body></html>
    """
    with open("last_output.xlsx", "wb") as f:
        f.write(excel_bytes.read())

    return markdown_html

except Exception as e:
    flash(str(e))
    return redirect('/')
```

@app.route('/download')
def download():
return send\_file("last\_output.xlsx", as\_attachment=True)

if **name** == '**main**':
port = int(os.environ.get("PORT", 5000))
app.run(host="0.0.0.0", port=port)










































