
from flask import Flask, request, jsonify, send_file, render_template
import pandas as pd
import unicodedata
import re
import os
from io import BytesIO
from markupsafe import Markup
import markdown2

app = Flask(__name__)

def normalize_column(col_name):
    if isinstance(col_name, str):
        col_name = unicodedata.normalize('NFKD', col_name).encode('ASCII', 'ignore').decode('utf-8')
        col_name = col_name.replace("’", "'").lower()
        col_name = re.sub(r'\s+', ' ', col_name).strip()
    return col_name

def surligner_article(texte, article):
    pattern = re.compile(rf"(Art[\.\:]?\s*{re.escape(article)}(?=[\s\W]|$))", re.IGNORECASE)
    return pattern.sub(r"<span style='color:red;font-weight:bold'>\1</span>", str(texte))

def marquer_article_markdown(texte, article):
    pattern = re.compile(rf"(Art[\.\:]?\s*{re.escape(article)}(?=[\s\W]|$))", re.IGNORECASE)
    return pattern.sub(r"**\1**", str(texte))

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/analyse', methods=['POST'])
def analyse_article():
    article = request.form.get('article', '').strip()
    file = request.files.get('file')

    if not article or not file:
        return render_template('index.html', erreur="Article ou fichier manquant")

    try:
        df = pd.read_excel(file)
        df = df.rename(columns=lambda c: normalize_column(c))

        colonnes_requises = [
            "articles enfreints", "duree totale effective radiation",
            "article amende/chef", "autres sanctions",
            "nom de l'intime", "numero de decision"
        ]
        for col in colonnes_requises:
            if col not in df.columns:
                return render_template('index.html', erreur=f"Colonne manquante: {col}")

        pattern = rf"Art[.:]?\s*{re.escape(article)}(?=[\s\W]|$)"
        masque = df["articles enfreints"].astype(str).str.contains(pattern, case=False, na=False, regex=True)
        resultats = df[masque].copy()

        if resultats.empty:
            return render_template('index.html', erreur=f"Aucun intime trouvé pour l'article {article}")

        resultats['Statut'] = 'Conforme'

        colonnes_affichees = [
            "Statut", "numero de decision", "nom de l'intime", "articles enfreints",
            "duree totale effective radiation", "article amende/chef", "autres sanctions"
        ]
        if 'resume' in resultats.columns:
            colonnes_affichees.append("resume")

        resultats = resultats[colonnes_affichees]

        for col in ["articles enfreints", "duree totale effective radiation", "article amende/chef", "autres sanctions"]:
            resultats[col] = resultats[col].apply(lambda x: surligner_article(x, article))

        markdown_table = resultats.copy()
        for col in ["articles enfreints", "duree totale effective radiation", "article amende/chef", "autres sanctions"]:
            markdown_table[col] = markdown_table[col].apply(lambda x: marquer_article_markdown(x, article))

        if 'resume' in markdown_table.columns:
            markdown_table['resume'] = markdown_table['resume'].apply(
                lambda x: f"[Résumé]({x})" if pd.notna(x) else ""
            )

        markdown_table = markdown_table.rename(columns={
            "numero de decision": "Numéro de décision",
            "nom de l'intime": "Nom de l’intime",
            "articles enfreints": "Articles enfreints",
            "duree totale effective radiation": "Périodes de radiation",
            "article amende/chef": "Amendes",
            "autres sanctions": "Autres sanctions",
            "resume": "Résumé"
        })

        markdown_str = markdown_table.to_markdown(index=False)
        markdown_html = markdown2.markdown(markdown_str)

        output_excel = BytesIO()
        with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
            resultats.to_excel(writer, sheet_name='Résultats', index=False)
            feuille = writer.sheets['Résultats']
            format_gras = writer.book.add_format({'bold': True, 'bg_color': '#DDDDDD'})
            for i, col in enumerate(resultats.columns):
                feuille.set_column(i, i, 30)
                feuille.write(0, i, col, format_gras)

        output_excel.seek(0)
        encoded_excel = output_excel.read()

        return render_template('resultats.html', tableau_html=Markup(markdown_html), fichier_excel=encoded_excel)

    except Exception as e:
        return render_template('index.html', erreur=f"Erreur : {str(e)}")

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)










































