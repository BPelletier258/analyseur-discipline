from flask import Flask, request, render_template
import pandas as pd
import unicodedata
import re

app = Flask(__name__)

def normalize_column(col_name):
    if isinstance(col_name, str):
        col_name = unicodedata.normalize('NFKD', col_name).encode('ASCII', 'ignore').decode('utf-8')
        col_name = col_name.replace("’", "'").lower()
        col_name = re.sub(r'\s+', ' ', col_name).strip()
    return col_name

def analyse_article_articles_enfreints_only(df, article_number):
    pattern_explicit = rf'\b[Aa]rt\.?\s*{re.escape(article_number)}\b'
    mask_articles_enfreints = df['articles enfreints'].astype(str).str.contains(pattern_explicit, na=False, flags=re.IGNORECASE)
    conformes = df[mask_articles_enfreints].copy()
    conformes['Statut'] = "Conforme"

    if conformes.empty:
        raise ValueError(f"Aucun intime trouvé pour l'article {article_number} demandé.")

    # Mise en gras/rouge uniquement sur l'article ciblé
    def surligner_article(c):
        return re.sub(
            rf'(\b[Aa]rt\.?\s*{re.escape(article_number)}\b)',
            r'<b style="color:red">\1</b>',
            str(c),
            flags=re.IGNORECASE
        )

    conformes['articles enfreints'] = conformes['articles enfreints'].apply(surligner_article)

    result = pd.DataFrame({
        'Numéro de décision': conformes.get('numero de decision', ''),
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

@app.route('/analyse', methods=['POST'])
def analyser():
    try:
        if not request.files.get('fichier') or not request.form.get('article'):
            raise ValueError("Veuillez fournir un fichier et un numéro d'article.")

        fichier_excel = request.files['fichier']
        article_recherche = request.form['article'].strip()

        df = pd.read_excel(fichier_excel)
        df.columns = [normalize_column(col) for col in df.columns]

        required_columns = [
            "articles enfreints",
            "duree totale effective radiation",
            "article amende/chef",
            "autres sanctions",
            "nom de l'intime"
        ]
        if not all(col in df.columns for col in required_columns):
            raise ValueError("Le fichier est incomplet. Merci de vérifier la structure.")

        # Ajouter la colonne 'numero de decision' si elle est présente
        if 'numero de decision' not in df.columns:
            df['numero de decision'] = ''

        resultats = analyse_article_articles_enfreints_only(df, article_recherche)
        return render_template('index.html', resultats=resultats.to_html(escape=False, index=False))

    except Exception as e:
        return render_template('index.html', erreur=str(e))

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)











