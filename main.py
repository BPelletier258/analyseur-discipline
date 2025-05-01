from flask import Flask, request, jsonify, render_template
import pandas as pd
import re
import unicodedata

app = Flask(__name__)

def normalize_column(col_name):
    if isinstance(col_name, str):
        col_name = unicodedata.normalize('NFKD', col_name).encode('ASCII', 'ignore').decode('utf-8')
        col_name = col_name.replace("’", "'")
        col_name = col_name.lower()
        col_name = re.sub(r'\s+', ' ', col_name).strip()
    return col_name

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/analyse', methods=['POST'])
def analyse():
    article_number = request.form.get('article')
    file = request.files.get('file')

    if not article_number or not file:
        return jsonify({'error': 'Fichier Excel ou numéro d’article manquant'}), 400

    try:
        df = pd.read_excel(file)
        df = df.rename(columns=lambda c: normalize_column(c))

        required_cols = [
            "articles enfreints",
            "duree totale effective radiation",
            "article amende/chef",
            "autres sanctions",
            "nom de l'intime",
            "numero de decision"
        ]

        if not all(col in df.columns for col in required_cols):
            return jsonify({"error": f"Colonnes manquantes. Colonnes trouvées : {list(df.columns)}"}), 422

        pattern = rf'Art[\.]?\s*{re.escape(article_number)}\b'
        mask = df["articles enfreints"].astype(str).str.contains(pattern, na=False, flags=re.IGNORECASE)
        filtered = df[mask].copy()

        if filtered.empty:
            return jsonify({"error": f"Aucun intime trouvé pour l'article {article_number}"}), 404

        filtered['Statut'] = 'Conforme'

        result = filtered[[
            "Statut",
            "numero de decision",
            "nom de l'intime",
            "articles enfreints",
            "duree totale effective radiation",
            "article amende/chef",
            "autres sanctions"
        ]]

        return result.to_dict(orient="records")

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host="0.0.0.0", port=5000)
