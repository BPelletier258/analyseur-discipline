from flask import Flask, render_template, request, jsonify
import pandas as pd
import re
import os
import markdown2

app = Flask(__name__)

def normalize_column(col_name):
    return (
        col_name.lower()
        .replace("’", "'")
        .replace("é", "e")
        .replace("è", "e")
        .replace("ê", "e")
        .replace("à", "a")
        .replace("ç", "c")
        .strip()
    )

def highlight_article(text, article):
    if pd.isna(text):
        return ""
    article_escaped = re.escape(article)
    pattern = rf"(Art[\.]?\s*{article_escaped})"
    return re.sub(pattern, r"<span class='rouge'><strong>\1</strong></span>", str(text), flags=re.IGNORECASE)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/analyse', methods=['POST'])
def analyse():
    try:
        file = request.files['file']
        article = request.form['article'].strip()
        df = pd.read_excel(file)

        df.columns = [normalize_column(c) for c in df.columns]

        required_cols = [
            "articles enfreints",
            "duree totale effective radiation",
            "article amende/chef",
            "autres sanctions",
            "nom de l'intime",
            "numero de decision"
        ]

        if not all(col in df.columns for col in required_cols):
            return render_template('index.html', erreur="Colonnes manquantes dans le fichier.")

        pattern = rf"Art[\.]?\s*{re.escape(article)}(?=\W|$)"
        masque = df["articles enfreints"].astype(str).str.contains(pattern, case=False, na=False)

        if not masque.any():
            return render_template('index.html', erreur=f"Aucun intime trouvé pour l'article {article}")

        donnees = df[masque].copy()
        donnees["Statut"] = "Conforme"

        for col in ["articles enfreints", "duree totale effective radiation", "article amende/chef", "autres sanctions"]:
            donnees[col] = donnees[col].apply(lambda x: highlight_article(x, article))

        if "resume" in donnees.columns:
            donnees["Résumé"] = donnees["resume"].apply(
                lambda x: f"[Résumé]({x})" if pd.notna(x) else ""
            )
            final_cols = ["Statut", "numero de decision", "nom de l'intime", "articles enfreints",
                          "duree totale effective radiation", "article amende/chef", "autres sanctions", "Résumé"]
        else:
            final_cols = ["Statut", "numero de decision", "nom de l'intime", "articles enfreints",
                          "duree totale effective radiation", "article amende/chef", "autres sanctions"]

        markdown_table = donnees[final_cols].to_markdown(index=False)
        markdown_html = "<html><head><style>body { font-family: Arial; line-height: 1.8; } .rouge { color: red; }</style></head><body><pre>{}</pre></body></html>".format(markdown_table)

        return markdown_html

    except Exception as e:
        return render_template('index.html', erreur=str(e))

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)










































