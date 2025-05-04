from flask import Flask, request, render_template
import pandas as pd
import re
import unicodedata
from io import BytesIO

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

    # Mise en gras + rouge de l’article détecté
    def surligner_article(val):
        return re.sub(
            pattern_explicit,
            lambda m: f"<b style='color: red'>{m.group(0)}</b>",
            str(val),
            flags=re.IGNORECASE
        )

    conformes['articles enfreints'] = conformes['articles enfreints'].apply(surligner_article)

    return pd.DataFrame({
        "Nom de l’intime": conformes["nom de l'intime"],
        f"Articles enfreints (Art {article_number})": conformes["articles enfreints"],
        f"Périodes de radiation (Art {article_number})": conformes['duree totale effective radiation'],
        f"Amendes (Art {article_number})": conformes['article amende/chef'],
        f"Autres sanctions (Art {article_number})": conformes['autres sanctions'],
        "Statut": conformes['Statut']
    })

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/analyse', methods=['POST'])
def analyser():
    try:
        print(">>> Analyseur lancé – version 3 mai 2025 ✅")  # ← Confirmation de version visible dans logs Render

        if 'fichier' not in request.files or 'article' not in request.form:
            raise ValueError("Veuillez fournir un fichier et un numéro d'article.")

        fichier = request.files['fichier']
        article = request.form['article'].strip()

        if not fichier or fichier.filename == '':
            raise ValueError("Aucun fichier sélectionné.")

        df = pd.read_excel(BytesIO(fichier.read()))
        df = df.rename(columns=lambda c: normalize_column(c))

        required_columns = [
            "articles enfreints",
            "duree totale effective radiation",
            "article amende/chef",
            "autres sanctions",
            "nom de l'intime"
        ]
        if not all(col in df.columns for col in required_columns):
            raise ValueError("Le fichier est incomplet. Merci de vérifier la structure.")

        resultats = analyse_article_articles_enfreints_only(df, article)

        return render_template("resultats.html", resultats=resultats.to_html(classes='table table-bordered table-striped', escape=False, index=False), erreur=None)

    except Exception as e:
        return render_template("index.html", erreur=str(e))

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)













