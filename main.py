import pandas as pd
import re
import unicodedata
from flask import Flask, render_template, request

app = Flask(__name__)

def normalize_column(col_name):
    if isinstance(col_name, str):
        col_name = unicodedata.normalize('NFKD', col_name).encode('ASCII', 'ignore').decode('utf-8')
        col_name = col_name.replace("’", "'").lower()
        col_name = re.sub(r'\s+', ' ', col_name).strip()
    return col_name

def analyse_article_articles_enfreints_only(df, article_number):
    pattern_explicit = rf'\bArt\.?\s*{re.escape(article_number)}\b'

    mask_articles_enfreints = df['articles enfreints'].astype(str).str.contains(
        pattern_explicit, na=False, flags=re.IGNORECASE
    )

    conformes = df[mask_articles_enfreints].copy()
    conformes['Statut'] = "Conforme"

    if conformes.empty:
        raise ValueError(f"Aucun intime trouvé pour l'article {article_number} demandé.")

    def highlight_article(cell):
        return re.sub(
            pattern_explicit,
            rf"<span style='color:red; font-weight:bold;'>Art. {article_number}</span>",
            cell,
            flags=re.IGNORECASE
        )

    conformes['articles enfreints'] = conformes['articles enfreints'].astype(str).apply(highlight_article)

    result = pd.DataFrame({
        'Nom de l’intime': conformes["nom de l'intime"],
        f'Articles enfreints (Art {article_number})': conformes['articles enfreints'],
        f'Périodes de radiation (Art {article_number})': conformes['duree totale effective radiation'],
        f'Amendes (Art {article_number})': conformes['article amende/chef'],
        f'Autres sanctions (Art {article_number})': conformes['autres sanctions'],
        'Statut': conformes['Statut']
    })

    return result.to_html(index=False, escape=False)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/analyse', methods=['POST'])
def analyser():
    try:
        fichier = request.files['fichier']
        article = request.form['article']
        print(">>> Requête reçue avec article =", article)

        df = pd.read_excel(fichier)
        df = df.rename(columns=lambda c: normalize_column(c))

        required = [
            "articles enfreints",
            "duree totale effective radiation",
            "article amende/chef",
            "autres sanctions",
            "nom de l'intime"
        ]
        if not all(col in df.columns for col in required):
            raise ValueError("Le fichier est incomplet. Merci de vérifier la structure.")

        tableau_html = analyse_article_articles_enfreints_only(df, article)
        return render_template('index.html', table=tableau_html, article=article)

    except Exception as e:
        return render_template('index.html', erreur=str(e))

if __name__ == '__main__':
    print(">>> Analyseur lancé – version déployée le 3 mai 2025 ✅")
    app.run(host='0.0.0.0', port=5000)














