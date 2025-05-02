from flask import Flask, render_template, request, jsonify
import pandas as pd
import re
import unicodedata
import os

app = Flask(__name__)

def normalize_column(col_name):
    if isinstance(col_name, str):
        col_name = unicodedata.normalize('NFKD', col_name).encode('ASCII', 'ignore').decode('utf-8')
        col_name = col_name.replace("’", "'")
        col_name = col_name.lower()
        col_name = re.sub(r'\s+', ' ', col_name).strip()
    return col_name

def analyse_article_articles_enfreints_only(df, article_number):
    pattern_explicit = rf'\bArt\.?\s*{re.escape(article_number)}\b'
    mask_articles_enfreints = df['articles enfreints'].astype(str).str.contains(pattern_explicit, na=False, flags=re.IGNORECASE)
    conformes = df[mask_articles_enfreints].copy()
    conformes['Statut'] = "Conforme"

    if conformes.empty:
        raise ValueError(f"Aucun intime trouvé pour l'article {article_number} demandé.")

    result = conformes[[
        "Statut",
        "numero de decision",
        "nom de l'intime",
        "articles enfreints",
        "duree totale effective radiation",
        "article amende/chef",
        "autres sanctions"
    ]]

    return result.to_dict(orient="records")

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/analyse', methods=['POST'])
def analyser():
    try:
        file = request.files['fichier']
        article = request.form['article'].strip()

        if not file:
            return jsonify({'error': 'Aucun fichier reçu.'}), 400

        df = pd.read_excel(file)
        df = df.rename(columns=lambda c: normalize_column(c))

        required_cols = [
            "articles enfreints",
            "duree totale effective radiation",
            "article amende/chef",
            "autres sanctions",
            "nom de l'intime",
            "numero de decision"
        ]
        if not all(col in df.columns for col in required_cols):
            return jsonify({'error': 'Le fichier est incomplet. Merci de vérifier la structure.'}), 400

        resultats = analyse_article_articles_enfreints_only(df, article)
        return jsonify(resultats)

    except ValueError as ve:
        return jsonify({'error': str(ve)}), 400
    except Exception as e:
    return render_template('index.html', erreur=str(e))

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)


