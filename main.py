import os
import re
import pandas as pd
import unicodedata
from flask import Flask, request, render_template, jsonify
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'

# Assurez-vous que ce dossier existe
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

def normalize_column(col_name):
    if isinstance(col_name, str):
        col_name = unicodedata.normalize('NFKD', col_name).encode('ASCII', 'ignore').decode('utf-8')
        col_name = col_name.replace("’", "'").lower()
        col_name = re.sub(r'\s+', ' ', col_name).strip()
    return col_name

def analyse_article_articles_enfreints_only(df, article_number):
    pattern_explicit = rf'\b[Aa]rt\.?\s*{re.escape(article_number)}\b'
    mask_articles_enfreints = df['articles enfreints'].astype(str).str.contains(
        pattern_explicit, na=False, flags=re.IGNORECASE)
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

@app.route("/", methods=["GET"])
def home():
    return render_template("index.html")

@app.route("/analyse", methods=["POST"])
def analyser():
    try:
        if "fichier_excel" not in request.files or "numero_article" not in request.form:
            raise ValueError("Formulaire incomplet.")

        fichier = request.files["fichier_excel"]
        article = request.form["numero_article"].strip()

        if not fichier.filename:
            raise ValueError("Aucun fichier fourni.")

        filename = secure_filename(fichier.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        fichier.save(file_path)

        df = pd.read_excel(file_path)
        df = df.rename(columns=lambda c: normalize_column(c))

        required_cols = [
            "articles enfreints",
            "duree totale effective radiation",
            "article amende/chef",
            "autres sanctions",
            "nom de l'intime"
        ]

        if not all(col in df.columns for col in required_cols):
            raise ValueError("Le fichier est incomplet. Merci de vérifier la structure.")

        resultats = analyse_article_articles_enfreints_only(df, article)
        return resultats.to_html(index=False)

    except ValueError as ve:
        return render_template("index.html", erreur=str(ve)), 400
    except Exception as e:
        return render_template("index.html", erreur=str(e))

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)







