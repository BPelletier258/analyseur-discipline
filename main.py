import pandas as pd
import unicodedata
import re
from flask import Flask, request, render_template
from datetime import datetime

app = Flask(__name__)

def normalize_column(col_name):
    if isinstance(col_name, str):
        col_name = unicodedata.normalize('NFKD', col_name).encode('ASCII', 'ignore').decode('utf-8')
        col_name = col_name.replace("’", "'").lower()
        col_name = re.sub(r'\s+', ' ', col_name).strip()
    return col_name

def analyse_article_articles_enfreints_only(df, article_number):
    # Préparation du pattern compatible avec des parenthèses, points, lettres
    pattern_explicit = rf'\bArt\.\s*{re.escape(article_number)}\b'
    print("[DEBUG] Regex utilisée:", pattern_explicit)

    mask_articles_enfreints = df['articles enfreints'].astype(str).str.contains(
        pattern_explicit, na=False, flags=re.IGNORECASE
    )
    conformes = df[mask_articles_enfreints].copy()
    print(f"[DEBUG] Correspondances trouvées : {conformes.shape[0]}")

    if conformes.empty:
        raise ValueError(f"Aucun intime trouvé pour l'article {article_number} demandé.")

    conformes['Statut'] = "Conforme"
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
    return render_template("index.html")

@app.route('/analyse', methods=['POST'])
def analyser():
    try:
        fichier = request.files.get('fichier')
        article = request.form.get('article', '').strip()

        print("\n>>> Requête reçue avec article =", article)
        if not fichier or not article:
            return render_template('index.html', erreur="Fichier ou article manquant.")

        df = pd.read_excel(fichier)
        df = df.rename(columns=lambda c: normalize_column(c))

        colonnes_attendues = [
            "articles enfreints",
            "duree totale effective radiation",
            "article amende/chef",
            "autres sanctions",
            "nom de l'intime"
        ]
        if not all(col in df.columns for col in colonnes_attendues):
            return render_template('index.html', erreur="Le fichier est incomplet. Merci de vérifier la structure.")

        resultats = analyse_article_articles_enfreints_only(df, article)
        print("[INFO] Analyse effectuée le", datetime.now().strftime("%d %B %Y"))

        # Mise en forme HTML de l'article recherché (gras et rouge)
        def surligner_article(val):
            if isinstance(val, str):
                val = re.sub(rf'(Art\.\s*{re.escape(article)})',
                             r'<b style="color:red">\1</b>', val, flags=re.IGNORECASE)
            return val

        colonnes_cibles = [col for col in resultats.columns if article in col or 'Articles enfreints' in col]
        for col in colonnes_cibles:
            resultats[col] = resultats[col].apply(surligner_article)

        return render_template('index.html', tables=[resultats.to_html(classes='data', escape=False, index=False)],
                               article=article)

    except Exception as e:
        print("[ERREUR]", str(e))
        return render_template('index.html', erreur=f"Erreur de traitement : {str(e)}")

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)


















