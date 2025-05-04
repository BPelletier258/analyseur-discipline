import os
import re
import unicodedata
import pandas as pd
from flask import Flask, request, render_template
from werkzeug.utils import secure_filename

app = Flask(__name__)

def normalize_column(col_name):
    if isinstance(col_name, str):
        col_name = unicodedata.normalize('NFKD', col_name).encode('ASCII', 'ignore').decode('utf-8')
        col_name = col_name.replace("’", "'")
        col_name = col_name.lower()
        col_name = re.sub(r'\s+', ' ', col_name).strip()
    return col_name

def analyse_article_articles_enfreints_only(df, article_number):
    pattern_explicit = rf'(?<!\w)[Aa]rt\.?\s*{re.escape(article_number)}(?!\w)'
    mask_articles_enfreints = df['articles enfreints'].astype(str).str.contains(pattern_explicit, na=False, flags=re.IGNORECASE)
    conformes = df[mask_articles_enfreints].copy()
    conformes['Statut'] = "Conforme"

    if conformes.empty:
        raise ValueError(f"Aucun intime trouvé pour l'article {article_number} demandé.")

    conformes['articles enfreints'] = conformes['articles enfreints'].astype(str).str.replace(
        pattern_explicit,
        r'<span class="highlight">\g<0></span>',
        flags=re.IGNORECASE,
        regex=True
    )

    return conformes[[
        "numero de decision",
        "nom de l'intime",
        "articles enfreints",
        "duree totale effective radiation",
        "article amende/chef",
        "autres sanctions",
        "Statut"
    ]]

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/analyse', methods=['POST'])
def analyser():
    try:
        if 'fichier' not in request.files or 'article' not in request.form:
            raise ValueError("Veuillez fournir un fichier et un numéro d'article.")

        fichier = request.files['fichier']
        article = request.form['article'].strip()

        if fichier.filename == '':
            raise ValueError("Aucun fichier sélectionné.")

        df = pd.read_excel(fichier)
        df.columns = [normalize_column(c) for c in df.columns]

        required_cols = [
            "numero de decision",
            "articles enfreints",
            "duree totale effective radiation",
            "article amende/chef",
            "autres sanctions",
            "nom de l'intime"
        ]

        if not all(col in df.columns for col in required_cols):
            raise ValueError("Le fichier est incomplet. La colonne 'Numéro de décision' est requise.")

        resultats = analyse_article_articles_enfreints_only(df, article)
        html_table = resultats.to_dict(orient='records')
        return render_template('index.html', resultats=html_table)

    except ValueError as ve:
        return render_template('index.html', erreur=str(ve))
    except Exception as e:
        return render_template('index.html', erreur=str(e))

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
















