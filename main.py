from flask import Flask, request, render_template
import pandas as pd
import unicodedata
import re
import os

app = Flask(__name__)

# 🔧 Fonction de normalisation
def normalize_column(col_name):
    if isinstance(col_name, str):
        col_name = unicodedata.normalize('NFKD', col_name).encode('ASCII', 'ignore').decode('utf-8')
        col_name = col_name.replace("’", "'")
        col_name = col_name.lower()
        col_name = re.sub(r'\s+', ' ', col_name).strip()
    return col_name

# 🧠 Fonction d’analyse de l’article dans Excel
def analyse_article(df, article_number):
    pattern_explicit = rf'\bArt\.?\s*{re.escape(article_number)}\b'
    mask_articles_enfreints = df['articles enfreints'].astype(str).str.contains(pattern_explicit, na=False, flags=re.IGNORECASE)
    conformes = df[mask_articles_enfreints].copy()
    conformes['Statut'] = "Conforme"

    if conformes.empty:
        return f"<p>Aucun intime trouvé pour l'article {article_number} demandé.</p>"

    result = conformes[[
        "nom de l'intime",
        "articles enfreints",
        "duree totale effective radiation",
        "article amende/chef",
        "autres sanctions"
    ]].copy()

    result.columns = [
        "Nom de l’intime",
        f"Articles enfreints (Art {article_number})",
        f"Périodes de radiation (Art {article_number})",
        f"Amendes (Art {article_number})",
        f"Autres sanctions (Art {article_number})"
    ]

    result_html = result.to_html(index=False)
    return f"<h2>Résultats pour l'article {article_number}</h2>" + result_html

# 📥 Route d’accueil
@app.route('/')
def index():
    return render_template('index.html')

# 🔍 Route d’analyse POST
@app.route('/analyse', methods=['POST'])
def analyser():
    try:
        if 'fichier' not in request.files or 'article' not in request.form:
            raise ValueError("Fichier ou article manquant.")

        fichier = request.files['fichier']
        article = request.form['article']

        df = pd.read_excel(fichier)
        df = df.rename(columns=lambda c: normalize_column(c))

        required_columns = [
            "articles enfreints",
            "duree totale effective radiation",
            "article amende/chef",
            "autres sanctions",
            "nom de l'intime"
        ]
        if not all(col in df.columns for col in required_columns):
            return "<p>❌ Le fichier est incomplet. Vérifiez la structure du tableau Excel.</p>"

        html_result = analyse_article(df, article)
        return html_result

    except Exception as e:
        return f"<p>⚠️ Erreur : {str(e)}</p>"

# 🚀 Lancement local (non utilisé sur Render)
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)












