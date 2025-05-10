import pandas as pd
import re
import unicodedata
from flask import Flask, request, jsonify
import io
import xlsxwriter

app = Flask(__name__)

def normalize_column(col_name):
    if isinstance(col_name, str):
        col_name = unicodedata.normalize('NFKD', col_name).encode('ASCII', 'ignore').decode('utf-8')
        col_name = col_name.replace("’", "'")
        col_name = col_name.lower()
        col_name = re.sub(r'\s+', ' ', col_name).strip()
    return col_name

def highlight_article(text, article):
    pattern = rf'(Art[\.:]?\s*{re.escape(article)}(?=[\s\W]|$))'
    return re.sub(pattern, r'**\\1**', text, flags=re.IGNORECASE)

@app.route('/analyse', methods=['POST'])
def analyse_article():
    article = request.form.get('article')
    file = request.files.get('file')
    if not article or not file:
        return jsonify({"error": "Fichier ou article manquant."}), 400

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
        for col in required_cols:
            if col not in df.columns:
                return jsonify({"error": f"Colonne requise manquante: {col}"}), 400

        pattern = rf'\bArt[\.:]?\s*{re.escape(article)}(?=[\s\W]|$)'
        mask = df['articles enfreints'].astype(str).str.contains(pattern, flags=re.IGNORECASE, na=False)
        filtered = df[mask].copy()

        if filtered.empty:
            return jsonify({"error": f"Aucun intime trouvé pour l'article {article}."}), 404

        filtered['Statut'] = 'Conforme'

        for col in ["articles enfreints", "duree totale effective radiation", "article amende/chef", "autres sanctions"]:
            filtered[col] = filtered[col].astype(str).apply(lambda x: highlight_article(x, article))

        def create_link(link):
            return f"[Résumé]({link})" if pd.notna(link) and str(link).startswith("http") else ""

        if 'resume' in filtered.columns:
            filtered['Résumé'] = filtered['resume'].apply(create_link)

        filtered = filtered[[
            'Statut',
            'numero de decision',
            "nom de l'intime",
            "articles enfreints",
            "duree totale effective radiation",
            "article amende/chef",
            "autres sanctions",
            "Résumé"
        ]]

        markdown_table = filtered.to_markdown(index=False)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            filtered.drop(columns=["Résumé"], errors='ignore').to_excel(writer, index=False, sheet_name='Résultats')
            worksheet = writer.sheets['Résultats']
            for i, col in enumerate(filtered.columns):
                worksheet.set_column(i, i, 30)
        output.seek(0)

        return jsonify({
            "markdown": markdown_table,
            "excel_file": output.getvalue().hex()
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == "__main__":
    app.run(debug=True)










































