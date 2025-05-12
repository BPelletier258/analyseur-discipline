import pandas as pd
import re
import unicodedata
from flask import Flask, request, render_template
from io import BytesIO
import xlsxwriter

app = Flask(__name__)

def normalize_column(col_name):
    if isinstance(col_name, str):
        col_name = unicodedata.normalize('NFKD', col_name).encode('ASCII', 'ignore').decode('utf-8')
        col_name = col_name.replace("’", "'")
        col_name = col_name.lower()
        col_name = re.sub(r'\s+', ' ', col_name).strip()
    return col_name

def style_article(cell, article):
    if not isinstance(cell, str):
        return cell
    pattern = re.compile(rf'(Art[\.\:]?\s*{re.escape(article)})', re.IGNORECASE)
    return pattern.sub(r'<span style="color:red">\1</span>', cell)

@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

@app.route('/analyse', methods=['POST'])
def analyse():
    try:
        file = request.files.get('file') or request.files.get('fichier_excel')
        article = request.form.get('article', '').strip()

        if not file or not article:
            return render_template('index.html', erreur="Veuillez fournir un fichier Excel et un article.")

        df = pd.read_excel(file)
        df = df.rename(columns=lambda c: normalize_column(c))

        required = [
            "articles enfreints",
            "duree totale effective radiation",
            "article amande/chef",
            "autres sanctions",
            "nom de l'intime",
            "numero de decision"
        ]
        missing = [col for col in required if col not in df.columns]
        if missing:
            return render_template('index.html', erreur=f"Le fichier est incomplet. Colonnes manquantes : {', '.join(missing)}")

        pattern_explicit = rf'\bArt[\.\:]?\s*{re.escape(article)}\b'
        mask = df['articles enfreints'].astype(str).str.contains(pattern_explicit, na=False, flags=re.IGNORECASE)
        conformes = df[mask].copy()

        if conformes.empty:
            return render_template('index.html', erreur=f"Aucun intime trouvé pour l'article {article} demandé.")

        md_columns = [
            'numero de decision',
            'nom de l\'intime',
            'articles enfreints',
            'duree totale effective radiation',
            'article amende/chef',
            'autres sanctions'
        ]
        display_df = conformes[md_columns]
        markdown_table = display_df.to_markdown(index=False)

        excel_columns = md_columns.copy()
        if 'resume' in conformes.columns:
            excel_columns.append('resume')
        excel_df = conformes[excel_columns].copy()

        for col in ['articles enfreints', 'duree totale effective radiation', 'article amende/chef', 'autres sanctions']:
            excel_df[col] = excel_df[col].apply(lambda x: style_article(x, article))

        output = BytesIO()
        workbook = xlsxwriter.Workbook(output, {'in_memory': True})
        worksheet = workbook.add_worksheet('Résultats')

        wrap = workbook.add_format({'text_wrap': True, 'valign': 'top'})
        header = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3'})
        rouge = workbook.add_format({'bg_color': '#FFC7CE', 'font_color': '#000000', 'text_wrap': True})

        for idx, col_name in enumerate(excel_df.columns):
            worksheet.write(0, idx, col_name, header)
            width = 60 if 2 <= idx <= 5 else 30
            worksheet.set_column(idx, idx, width)

        for row_num, row in enumerate(excel_df.itertuples(index=False), start=1):
            for col_num, value in enumerate(row):
                text = str(value)
                if re.search(pattern_explicit, unicodedata.normalize('NFKD', text), flags=re.IGNORECASE):
                    worksheet.write(row_num, col_num, text, rouge)
                else:
                    worksheet.write(row_num, col_num, text, wrap)

        workbook.close()
        output.seek(0)

        return render_template('resultats.html', table_markdown=markdown_table, fichier_excel=output.read(), filename=f"resultats_article_{article}.xlsx")

    except Exception as e:
        return render_template('index.html', erreur=str(e))

if __name__ == '__main__':
    app.run(debug=True)












































