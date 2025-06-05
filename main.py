import re
import pandas as pd
from flask import Flask, request, render_template_string, send_file, redirect, url_for
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
last_excel = None
last_article = None

# --- Bloc CSS en ligne et template HTML ---
STYLE_BLOCK = '''
body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; margin: 20px; background: #f5f7fa; }
h1 { font-size: 1.65em; margin-bottom: 0.5em; color: #333; }
.form-container { background: #fff; padding: 15px; border-radius: 8px; box-shadow: 0 2px 8px rgba(0,0,0,0.1); max-width: 750px; }
form { display: flex; flex-wrap: wrap; gap: 1rem; align-items: flex-end; }
label { font-weight: bold; font-size: 1.05em; color: #444; display: flex; flex-direction: column; }
input[type=file], input[type=text] { padding: 0.6em; font-size: 1.05em; border: 1px solid #ccc; border-radius: 4px; }
button { padding: 0.6em 1.2em; font-size: 1.05em; font-weight: bold; background: #007bff; color: #fff; border: none; border-radius: 4px; cursor: pointer; transition: background 0.3s ease; }
button:hover { background: #0056b3; }
.table-container {
  width: 100%;
  overflow-x: scroll;            /* barre horizontale toujours visible */
  overflow-y: hidden;
  scrollbar-gutter: stable both-edges;
  -webkit-overflow-scrolling: touch;
  margin-top: 30px;
}
table { border-collapse: collapse; width: max-content; background: #fff; display: inline-block; }
th, td { border: 1px solid #888; padding: 8px; vertical-align: top; }
/* En-têtes centrés */
th { background: #e2e3e5; font-weight: bold; font-size: 1em; text-align: center; }
/* Largeur par défaut : 25ch */
th, td { width: 25ch; }
/* Colonnes « détaillées » en 50ch (Réservé aux indices 8,9,10,11,13) */
th:nth-child(8), td:nth-child(8),
th:nth-child(9), td:nth-child(9),
th:nth-child(10), td:nth-child(10),
th:nth-child(11), td:nth-child(11),
th:nth-child(13), td:nth-child(13)
{ width: 50ch; }
.highlight { color: #d41e26; font-weight: bold; }
.summary-link { color: #0066cc; text-decoration: underline; }
'''

HTML_TEMPLATE = '''
<!doctype html>
<html lang="fr">
<head>
  <meta charset="utf-8">
  <title>Analyse Disciplinaire</title>
  <style>{{ style_block }}</style>
</head>
<body>
  <h1>Analyse Disciplinaire</h1>
  <div class="form-container">
    <form method="post" enctype="multipart/form-data">
      <label>Fichier Excel:<input type="file" name="file" required></label>
      <label>Article à filtrer:<input type="text" name="article" placeholder="ex: 14 ou 59(2)" required></label>
      <button type="submit">Analyser</button>
    </form>
  </div>
  <hr>
  {% if searched_article %}
    <div><strong>Article recherché : <span class="highlight">{{ searched_article }}</span></strong></div>
    <a href="/download">⬇️ Télécharger le fichier Excel formaté</a>
  {% endif %}
  {% if table_html %}
    <div class="table-container">
      {{ table_html|safe }}
    </div>
  {% endif %}
</body>
</html>
'''

def build_pattern(article):
    """
    Construit un pattern qui matche uniquement
    - 'Art. X', 'Art: X' ou 'Art : X' 
    - où X peut contenir des points (ex. 2.04, 3.02.08, 59(1), etc.)
    - et qui n'est pas suivi immédiatement d'un chiffre (pour éviter 591 si on cherche '59').
    """
    art_escaped = re.escape(article)  # échappe les points, parenthèses, etc.

    # On autorise zéro ou plusieurs espaces (normaux \s ou insécables \u00A0) autour des deux-points.
    space = r'(?:\s|\u00A0)*'
    prefixes = [
        r'Art\.' + space,
        r'Art:'  + space,
        r'Art' + space + r':' + space
    ]
    pref = '|'.join(prefixes)

    # Si l'article contient explicitement une parenthèse, on désactive le lookahead de chiffres
    if '(' in article:
        return rf'(?:{pref}){art_escaped}'
    else:
        return rf'(?:{pref}){art_escaped}(?![0-9])'

grey_fill    = PatternFill(start_color="DDDDDD", end_color="DDDDDD", fill_type="solid")
red_font     = Font(color="FF0000")
link_font    = Font(color="0000FF", underline="single")
border_style = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)
wrap_alignment = Alignment(wrap_text=True, vertical='top')

# Colonnes à mettre en rouge dans le HTML et Excel (si l'article y est présent)
HIGHLIGHT_COLS = {
    'Articles enfreints',
    'Durée totale effective radiation',
    'Article amende/chef',
    'Autres sanctions'
}

@app.route('/', methods=['GET','POST'])
def analyze():
    global last_excel, last_article

    if request.method == 'POST':
        # 1) Lecture du fichier Excel et construction du regex
        file    = request.files['file']
        article = request.form['article'].strip()
        last_article = article

        df = pd.read_excel(file)
        pat = build_pattern(article)

        # 2) Filtrage : on ne cherche que dans la colonne "Articles enfreints"
        #    Convertit en string pour éviter les NaN, etc.
        mask = df['Articles enfreints'].astype(str).apply(lambda v: bool(re.search(pat, v)))
        df_f = df[mask].copy()

        # 3) Préparation du HTML
        html_df = df_f.fillna('')  # remplace NaN par chaîne vide pour l'affichage

        # Pour chaque colonne détaillée à colorer, on entoure l'article recherché d'un <span class="highlight">
        for col in (HIGHLIGHT_COLS & set(html_df.columns)):
            html_df[col] = html_df[col].astype(str).str.replace(
                pat,
                lambda m: f"<span class='highlight'>{m.group(0)}</span>",
                regex=True
            )

        # Si "Résumé" existe, on remplace chaque URL par un lien cliquable
        if 'Résumé' in html_df.columns:
            html_df['Résumé'] = html_df['Résumé'].apply(
                lambda url: f'<a href="{url}" class="summary-link" target="_blank">Résumé</a>' if isinstance(url, str) and url else ''
            )

        table_html = html_df.to_html(index=False, escape=False)

        # 4) Construction du fichier Excel à télécharger
        buf = BytesIO()
        wb  = Workbook()
        ws  = wb.active

        # Entête fusionnée
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(df_f.columns))
        top_cell = ws.cell(row=1, column=1, value=f"Article filtré : {article}")
        top_cell.font = Font(size=14, bold=True)

        # Ligne des noms de colonnes
        for idx, col in enumerate(df_f.columns, start=1):
            c = ws.cell(row=2, column=idx, value=col)
            c.fill      = grey_fill
            c.font      = Font(bold=True)
            c.border    = border_style
            c.alignment = wrap_alignment

        # Données
        for r_idx, row in enumerate(df_f.itertuples(index=False), start=3):
            for c_idx, col in enumerate(df_f.columns, start=1):
                val = getattr(row, row._fields[c_idx-1])

                # Si c'est la colonne "Résumé" et que c'est une chaîne non vide, on crée un lien
                if col == 'Résumé' and isinstance(val, str) and val:
                    cell = ws.cell(row=r_idx, column=c_idx, value='Résumé')
                    cell.hyperlink = val
                    cell.font = link_font
                else:
                    # Pour toutes les autres colonnes, on colle la valeur (vide si None ou NaN)
                    cell = ws.cell(row=r_idx, column=c_idx, value=(val if val else ''))

                cell.border    = border_style
                cell.alignment = wrap_alignment

                # Si c'est l'une des colonnes à surligner et que le pattern y match, on colore en rouge
                if col in HIGHLIGHT_COLS and re.search(pat, str(val)):
                    cell.font = red_font

        # Réglage des largeurs : 25ch (par défaut) ou 50ch pour les colonnes détaillées
        for i in range(1, len(df_f.columns) + 1):
            ws.column_dimensions[get_column_letter(i)].width = 25
        # Indices 8,9,10,11,13 doivent être élargis à 50
        for j in (8, 9, 10, 11, 13):
            if j <= len(df_f.columns):
                ws.column_dimensions[get_column_letter(j)].width = 50

        wb.save(buf)
        buf.seek(0)
        last_excel = buf.getvalue()

        # Affichage du résultat HTML
        return render_template_string(
            HTML_TEMPLATE,
            style_block=STYLE_BLOCK,
            table_html=table_html,
            searched_article=article
        )

    # Si GET, on affiche uniquement le formulaire
    return render_template_string(HTML_TEMPLATE, style_block=STYLE_BLOCK)


@app.route('/download')
def download():
    global last_excel, last_article
    if not last_excel:
        return redirect(url_for('analyze'))
    return send_file(
        BytesIO(last_excel),
        as_attachment=True,
        download_name=f"decisions_filtrees_{last_article}.xlsx"
    )


if __name__ == '__main__':
    app.run(debug=True)













































































































































































































