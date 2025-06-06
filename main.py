import pandas as pd
import re
import unicodedata
from flask import Flask, request, render\_template
from io import BytesIO
import xlsxwriter

app = Flask(**name**)

# Normalisation des noms de colonnes (sans accents, minuscule)

def normalize\_column(col\_name):
if isinstance(col\_name, str):
col\_name = unicodedata.normalize('NFKD', col\_name).encode('ASCII', 'ignore').decode('utf-8')
col\_name = col\_name.replace("’", "'")
col\_name = col\_name.lower()
col\_name = re.sub(r'\s+', ' ', col\_name).strip()
return col\_name

# Validation stricte du format d'article, incluant source (CD, CP, R, etc.)

def validate\_article\_format(article):
\# Exemples valides : 14, 59(2), 2.01a) R15, 3.02.08
pattern = r'^\[0-9]+(?:.\[0-9]+)*(?:$[^)]*$)?(?:\s*\[A-Za-z0-9]+)?\$'
return bool(re.match(pattern, article))

# Applique mise en forme HTML rouge pour occurrences d'article dans une cellule

def style\_article(cell, article):
if not isinstance(cell, str):
return cell
\# Échappe l'article, puis autorise espaces variables
esc = re.escape(article)
art\_pattern = esc.replace(r'\ ', r'\s\*')
prefix = r'(?\:Art.|Art:|Art\s\*:)' + r'\s\*'
\# Si parenthèse dans l'article, ne pas ajouter lookahead
if '(' in article:
regex = re.compile(rf"{prefix}{art\_pattern}", re.IGNORECASE)
else:
regex = re.compile(rf"{prefix}{art\_pattern}(?!\[0-9])", re.IGNORECASE)
return regex.sub(lambda m: f"<span style='color:red;font-weight:bold'>{m.group(0)}</span>", cell)

@app.route('/', methods=\['GET'])
def index():
return render\_template('index.html')

@app.route('/analyse', methods=\['POST'])
def analyse():
try:
file = request.files.get('file') or request.files.get('fichier\_excel')
article = request.form.get('article', '').strip()

```
    if not file or not article:
        return render_template('index.html', erreur="Veuillez fournir un fichier Excel et un article.")

    # Validation stricte du format d'article
    if not validate_article_format(article):
        return render_template('index.html', erreur="Format d'article non valide. Exemple : 14, 59(2), 2.01a) R15, 3.02.08")

    df = pd.read_excel(file)
    df = df.rename(columns=lambda c: normalize_column(c))

    required = [
        "articles enfreints",
        "duree totale effective radiation",
        "article amende/chef",
        "autres sanctions",
        "nom de l'intime",
        "numero de decision"
    ]
    missing = [col for col in required if col not in df.columns]
    if missing:
        return render_template('index.html', erreur=f"Le fichier est incomplet. Colonnes manquantes : {', '.join(missing)}")

    # Construction du motif strict permettant espaces variables
    esc = re.escape(article)
    art_pattern = esc.replace(r'\ ', r'\s*')
    prefix = r'(?:Art\.|Art:|Art\s*:)' + r'\s*'
    if '(' in article:
        pat_str = rf"(?:{prefix}){art_pattern}"
    else:
        pat_str = rf"(?:{prefix}){art_pattern}(?![0-9])"
    pattern_explicit = re.compile(pat_str, re.IGNORECASE)

    # Filtrage sur la colonne 'articles enfreints'
    mask = df['articles enfreints'].astype(str).apply(lambda v: bool(pattern_explicit.search(v)))
    conformes = df[mask].copy()

    if conformes.empty:
        return render_template('index.html', erreur=f"Aucun intime trouvé pour l'article {article}.")

    # Colonnes pour affichage Markdown
    md_columns = [
        'numero de decision',
        "nom de l'intime",
        "articles enfreints",
        "duree totale effective radiation",
        "article amende/chef",
        "autres sanctions"
    ]
    display_df = conformes[md_columns]
    markdown_table = display_df.to_markdown(index=False)

    # Préparation du DataFrame pour Excel
    excel_columns = md_columns.copy()
    if 'resume' in conformes.columns:
        excel_columns.append('resume')
    excel_df = conformes[excel_columns].copy()

    # Appliquer coloration HTML brute pour Excel
    for col in ["articles enfreints", "duree totale effective radiation", "article amende/chef", "autres sanctions"]:
        excel_df[col] = excel_df[col].apply(lambda x: style_article(x, article))

    # Génération du fichier Excel
    output = BytesIO()
    workbook = xlsxwriter.Workbook(output, {'in_memory': True})
    worksheet = workbook.add_worksheet('Résultats')

    # Formats
    wrap = workbook.add_format({'text_wrap': True, 'valign': 'top'})
    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3'})
    red_fmt = workbook.add_format({'font_color': '#FF0000', 'text_wrap': True, 'valign': 'top'})

    # Écriture des en-têtes avec largeur ajustée
    for idx, col_name in enumerate(excel_df.columns):
        worksheet.write(0, idx, col_name, header_fmt)
        width = 60 if 2 <= idx <= 5 else 30
        worksheet.set_column(idx, idx, width)

    # Écriture des données, coloration selon motif sur texte normalisé
    for r, row in enumerate(excel_df.itertuples(index=False), start=1):
        for c, val in enumerate(row):
            txt = '' if pd.isna(val) else str(val)
            plain = unicodedata.normalize('NFKD', re.sub(r'<[^>]+>', '', txt))
            if pattern_explicit.search(plain):
                worksheet.write(r, c, txt, red_fmt)
            else:
                worksheet.write(r, c, txt, wrap)

    workbook.close()
    output.seek(0)
    excel_data = output.read()

    return render_template('resultats.html', table_markdown=markdown_table, fichier_excel=excel_data, filename=f"resultats_article_{article}.xlsx")

except Exception as e:
    return render_template('index.html', erreur=str(e))
```

if **name** == '**main**':
app.run(debug=True)





















































































































































































































