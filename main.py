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
    try:
        article_regex = re.escape(article)
        pattern = rf'(Art[\.:]?\s*{article_regex})(?=[\s\W]|$)'
        return re.sub(pattern, r'<span class="rouge">\1</span>', str(texte), flags=re.IGNORECASE)
    except Exception:
        return str(texte)

# --- Route principale ---
@app.route('/', methods=['GET'])
def index():
    return render_template("index.html")

@app.route('/analyse', methods=['POST'])
def analyser():
    try:
        fichier = request.files['fichier']
        article = request.form['article'].strip()
        print(f">>> Requête reçue avec article = {article} (test 3 mai)")

        df = pd.read_excel(fichier)
        df = df.rename(columns=lambda c: normalize_column(c))

        required_columns = [
            "articles enfreints", "duree totale effective radiation",
            "article amende/chef", "autres sanctions",
            "nom de l'intime", "numero de decision"
        ]
        if not all(col in df.columns for col in required_columns):
            raise ValueError("Le fichier est incomplet. Merci de vérifier la structure.")

        pattern_explicit = rf'Art[\.:]?\s*{re.escape(article)}(?=[\s\W]|$)'
        mask = df['articles enfreints'].astype(str).str.contains(pattern_explicit, na=False, flags=re.IGNORECASE)
        conformes = df[mask].copy()

        if conformes.empty:
            raise ValueError(f"Aucun intime trouvé pour l'article {article} demandé.")

        colonnes_a_surligner = {
            'Articles enfreints': 'articles enfreints',
            'Périodes de radiation': 'duree totale effective radiation',
            'Amendes': 'article amende/chef',
            'Autres sanctions': 'autres sanctions'
        }
        for nouvelle_col, source_col in colonnes_a_surligner.items():
            conformes[nouvelle_col] = conformes[source_col].astype(str).apply(lambda x: surligner_article(x, article))

        conformes['Numéro de décision'] = conformes['numero de decision']
        conformes['Nom de l’intime'] = conformes["nom de l'intime"]
        conformes['Statut'] = "Conforme"

        colonnes = ['Statut', 'Numéro de décision', 'Nom de l’intime'] + list(colonnes_a_surligner.keys())
        tableau_html = conformes[colonnes].to_html(classes='table table-striped', escape=False, index=False)

        return render_template("index.html", tableau_html=tableau_html, article=article)

    except Exception as e:
        return render_template("index.html", erreur=str(e))

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)


























