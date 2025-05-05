import pandas as pd
import re
import unicodedata
from flask import Flask, render_template, request

app = Flask(__name__)

# --- Utilitaires ---
def normalize_column(col_name):
    if isinstance(col_name, str):
        col_name = unicodedata.normalize('NFKD', col_name).encode('ASCII', 'ignore').decode('utf-8')
        col_name = col_name.replace("’", "'")
        col_name = col_name.lower()
        col_name = re.sub(r'\s+', ' ', col_name).strip()
    return col_name

def surligner_article(texte, article):
    if pd.isna(texte):
        return ""
    pattern = re.escape(article)
    return re.sub(rf'\b(Art\.?\s*{pattern})\b', r'<span class="rouge">\1</span>', str(texte), flags=re.IGNORECASE)

# --- Route principale ---
@app.route('/', methods=['GET'])
def index():
    return render_template("index.html")

@app.route('/analyse', methods=['POST'])
def analyser():
    try:
        fichier = request.files['fichier']
        article = request.form['article'].strip()
        print(f">>> Requete recue avec article = {article}")

        df = pd.read_excel(fichier)
        df = df.rename(columns=lambda c: normalize_column(c))

        required_columns = [
            "articles enfreints", "duree totale effective radiation",
            "article amende/chef", "autres sanctions", "nom de l'intime"
        ]
        if not all(col in df.columns for col in required_columns):
            raise ValueError("Le fichier est incomplet. Merci de verifier la structure.")

        pattern_explicit = rf'\bArt\.?\s*{re.escape(article)}(?=\b|[^a-zA-Z0-9])'
        mask = df['articles enfreints'].astype(str).str.contains(pattern_explicit, na=False, flags=re.IGNORECASE)
        conformes = df[mask].copy()

        if conformes.empty:
            raise ValueError(f"Aucun intime trouvé pour l'article {article} demandé.")

        conformes['Articles enfreints'] = conformes['articles enfreints'].apply(lambda x: surligner_article(x, article))
        conformes['Périodes de radiation'] = conformes['duree totale effective radiation'].apply(lambda x: surligner_article(x, article))
        conformes['Amendes'] = conformes['article amende/chef'].apply(lambda x: surligner_article(x, article))
        conformes['Autres sanctions'] = conformes['autres sanctions'].apply(lambda x: surligner_article(x, article))
        conformes['Nom de l’intime'] = conformes["nom de l'intime"]
        conformes['Statut'] = "Conforme"

        colonnes = ['Nom de l’intime', 'Articles enfreints', 'Périodes de radiation', 'Amendes', 'Autres sanctions', 'Statut']
        tableau_html = conformes[colonnes].to_html(classes='table table-striped', escape=False, index=False)

        return render_template("index.html", tableau_html=tableau_html, article=article)

    except Exception as e:
        return render_template("index.html", erreur=str(e))

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)



















