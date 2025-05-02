from flask import Flask, render_template, request, redirect, jsonify
import pandas as pd
import unicodedata
import re
import os

app = Flask(__name__)

# 🔧 Normalisation des noms de colonnes
def normalize_column(col_name):
    if isinstance(col_name, str):
        col_name = unicodedata.normalize('NFKD', col_name).encode('ASCII', 'ignore').decode('utf-8')
        col_name = col_name.replace("’", "'").lower()
        col_name = re.sub(r'\s+', ' ', col_name).strip()
    return col_name

# 🔍 Analyse stricte uniquement sur "articles enfreints"
def analyse_article_articles_enfreints_only(df, article_number):
    # Accepte format 14, 59(2), 2.01 a), etc.
    pattern_explicit = rf'\b[Aa]rt\.?\s*{re.escape(article_number)}\b'
    mask_articles_enfreints = df['articles enfreints'].astype(str).str.contains(pattern_explicit, na=False, flags=re.IGNORECASE)
    conformes = df[mask_articles_enfreints].copy()
    conformes['Statut'] = "Conforme"

    if conformes.empty:
        raise ValueError(f"Aucun intime trouvé pour l'article {article_number} demandé.")

    result = pd.DataFrame({
        "Nom de l’intime": conformes["nom de l'intime"],
        f"Articles enfreints (Art {article_number})": conformes['articles enfreints'],
        f"Périodes de radiation (Art {article_number})": conformes['duree totale effective radiation'],
        f"Amendes (Art {article_number})": conformes['article amende/chef'],
        f"Autres sanctions (Art {article_number})": conformes['autres sanctions'],
        "Statut": conformes['Statut']
    })
    return result

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/analyse', methods=['POST'])
def analyser():
    try:
        fichier = request.files['fichier_excel']
        article = request.form['numero_article'].strip()

        if not fichier or not article:
            raise ValueError("Le fichier ou le numéro d'article est manquant.")

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

        resultats = analyse_article_articles_enfreints_only(df, article)
        return render_template('resultats.html', tables=[resultats.to_html(classes='table table-striped', index=False)], article=article)

    except ValueError as ve:
        return render_template('index.html', erreur=f"Erreur de traitement: {str(ve)}")
    except Exception as e:
        return render_template('index.html', erreur=f"Erreur technique: {str(e)}")

# 🚀 Pour Render
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)





