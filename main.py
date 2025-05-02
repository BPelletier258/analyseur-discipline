from flask import Flask, request, render_template, jsonify
import pandas as pd
import re
import os
import unicodedata

app = Flask(__name__)

def normalize_column(col_name):
    if isinstance(col_name, str):
        col_name = unicodedata.normalize('NFKD', col_name).encode('ASCII', 'ignore').decode('utf-8')
        col_name = col_name.replace("’", "'")
        col_name = col_name.lower()
        col_name = re.sub(r'\s+', ' ', col_name).strip()
    return col_name

def analyse_article_articles_enfreints_only(df, article_number):
    # expression régulière plus permissive (ex. Art. 2.01 a), 59(2), 3.01.02, etc.)
    pattern_explicit = rf'\b[Aa]rt\.?\s*{re.escape(article_number)}\b'
    mask_articles_enfreints = df['articles enfreints'].astype(str).str.contains(pattern_explicit, na=False, flags=re.IGNORECASE)
    conformes = df[mask_articles_enfreints].copy()
    conformes['Statut'] = "Conforme"

    if conformes.empty:
        raise ValueError(f"Aucun intime trouvé pour l'article {article_number} demandé.")

    return pd.DataFrame({
        'Nom de l’intime': conformes["nom de l'intime"],
        f'Articles enfreints (Art {article_number})': conformes['articles enfreints'],
        f'Périodes de radiation (Art {article_number})': conformes['duree totale effective radiation'],
        f'Amendes (Art {article_number})': conformes['article amende/chef'],
        f'Autres sanctions (Art {article_number})': conformes['autres sanctions'],
        'Statut': conformes['Statut']
    })

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")

@app.route("/analyse", methods=["POST"])
def analyser():
    try:
        if "fichier_excel" not in request.files or "numero_article" not in request.form:
            return render_template("index.html", erreur="Veuillez fournir un fichier Excel et un numéro d'article.")

        fichier = request.files["fichier_excel"]
        article = request.form["numero_article"].strip()

        if not fichier or fichier.filename == "":
            return render_template("index.html", erreur="Aucun fichier sélectionné.")

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
            return render_template("index.html", erreur="Le fichier est incomplet. Merci de vérifier la structure.")

        resultats = analyse_article_articles_enfreints_only(df, article)
        return render_template("résultats.html", tables=[resultats.to_html(classes="table table-bordered", index=False)], article=article)

    except ValueError as ve:
        return render_template("index.html", erreur=str(ve))
    except Exception as e:
        return render_template("index.html", erreur=str(e))

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)






