import os
import re
import pandas as pd
import unicodedata
from flask import Flask, request, render_template, jsonify

app = Flask(__name__)

def normalize_column(col_name):
    if isinstance(col_name, str):
        col_name = unicodedata.normalize('NFKD', col_name).encode('ASCII', 'ignore').decode('utf-8')
        col_name = col_name.replace("’", "'")
        col_name = col_name.lower()
        col_name = re.sub(r'\s+', ' ', col_name).strip()
    return col_name

def analyse_article_articles_enfreints_only(df, article_number):
    pattern_explicit = rf'\b[Aa]rt\.?\s*{re.escape(article_number)}\b'
    mask_articles_enfreints = df['articles enfreints'].astype(str).str.contains(pattern_explicit, na=False, flags=re.IGNORECASE)
    conformes = df[mask_articles_enfreints].copy()
    conformes['Statut'] = "Conforme"

    if conformes.empty:
        raise ValueError(f"Aucun intime trouvé pour l'article {article_number} demandé.")

    result = pd.DataFrame({
        'Nom de l’intime': conformes["nom de l'intime"],
        f'Articles enfreints (Art {article_number})': conformes['articles enfreints'],
        f'Périodes de radiation (Art {article_number})': conformes['duree totale effective radiation'],
        f'Amendes (Art {article_number})': conformes['article amende/chef'],
        f'Autres sanctions (Art {article_number})': conformes['autres sanctions'],
        'Statut': conformes['Statut']
    })
    return result

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/analyse/<article_number>', methods=['POST'])
def analyse(article_number):
    try:
        fichier = request.files['fichier']
        if not fichier:
            return render_template('index.html', erreur="Aucun fichier soumis.")

        df = pd.read_excel(fichier)
        df = df.rename(columns=lambda c: normalize_column(c))

        required_cols = [
            "articles enfreints",
            "duree totale effective radiation",
            "article amende/chef",
            "autres sanctions",
            "nom de l'intime"
        ]

        if not all(col in df.columns for col in required_cols):
            return render_template('index.html', erreur="Le fichier est incomplet. Merci de vérifier la structure.")

        resultats = analyse_article_articles_enfreints_only(df, article_number)
        return render_template('resultats.html', tables=[resultats.to_html(classes='data', index=False)], article=article_number)

    except ValueError as ve:
        return render_template('index.html', erreur=str(ve))
    except Exception as e:
        return render_template('index.html', erreur=str(e))

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=True)




