
import os
import re
import pandas as pd
from flask import Flask, request, render_template, send_file, redirect, flash
from werkzeug.utils import secure_filename
from io import BytesIO
import markdown

app = Flask(__name__)
app.secret_key = "secret"
UPLOAD_FOLDER = "uploads"
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def normalize_column(col_name):
    import unicodedata
    if isinstance(col_name, str):
        col_name = unicodedata.normalize('NFKD', col_name).encode('ASCII', 'ignore').decode('utf-8')
        col_name = col_name.replace("’", "'")
        col_name = col_name.lower().strip()
        col_name = re.sub(r'\s+', ' ', col_name)
    return col_name

def highlight_article(text, article):
    if not isinstance(text, str):
        return text
    pattern = re.compile(rf"(Art[\.:]?\s*{re.escape(article)})", flags=re.IGNORECASE)
    return pattern.sub(r"<span style='color:red'><b>\1</b></span>", text)

@app.route("/", methods=["GET", "POST"])
def index():
    return render_template("index.html")

@app.route("/analyse", methods=["POST"])
def analyse():
    try:
        file = request.files["fichier"]
        article = request.form["article"].strip()

        if not file or not article:
            flash("Veuillez fournir un fichier et un numéro d'article.")
            return redirect("/")

        df = pd.read_excel(file)
        df = df.rename(columns=lambda c: normalize_column(c))

        required_columns = [
            "articles enfreints",
            "duree totale effective radiation",
            "article amende/chef",
            "autres sanctions",
            "nom de l'intime",
            "numero de decision"
        ]

        for col in required_columns:
            if col not in df.columns:
                return render_template("index.html", erreur="Le fichier est incomplet. Merci de vérifier la structure.")

        pattern = rf"\bArt[\.:]?\s*{re.escape(article)}(?=[\s\W]|$)"
        mask = df["articles enfreints"].astype(str).str.contains(pattern, na=False, flags=re.IGNORECASE)

        if not mask.any():
            return render_template("index.html", erreur=f"Aucun intime trouvé pour l'article {article} demandé.")

        result = df[mask].copy()
        result["Statut"] = "Conforme"

        # Mise en forme HTML
        columns_order = [
            "Statut",
            "numero de decision",
            "nom de l'intime",
            "articles enfreints",
            "duree totale effective radiation",
            "article amende/chef",
            "autres sanctions",
        ]
        if "resume" in result.columns:
            columns_order.append("resume")

        for col in columns_order:
            if col in result.columns:
                result[col] = result[col].apply(lambda x: highlight_article(x, article))

        if "resume" in result.columns:
            result["resume"] = result["resume"].apply(lambda x: f'<a href="{x}" target="_blank">Résumé</a>' if pd.notna(x) else "")

        # Génération Markdown
        table_md = result[columns_order].to_markdown(index=False)

        table_html = markdown.markdown(f"\n\n{table_md}\n\n", extensions=["tables"])

        return render_template("resultats.html", tableau_html=table_html)

    except Exception as e:
        return render_template("index.html", erreur=str(e))

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)










































