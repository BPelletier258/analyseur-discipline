
import os
import pandas as pd
import re
from flask import Flask, request, render_template, send_file
from io import BytesIO
from markupsafe import Markup

app = Flask(__name__)

def normalize_column(col):
    col = col.strip().lower()
    col = re.sub(r"['’]", "'", col)
    col = re.sub(r"[^a-z0-9àâçéèêëîïôûùüÿñæœ ._-]", "", col)
    return col

def clean_column_name(col):
    return normalize_column(str(col)).replace(" ", "_")

def highlight_article(text, article):
    pattern = re.compile(rf"(Art\.?\s*{re.escape(article)})", re.IGNORECASE)
    return pattern.sub(r"**\1**", str(text))

@app.route("/", methods=["GET", "POST"])
def analyse():
    if request.method == "POST":
        article = request.form.get("article", "").strip()
        file = request.files.get("file")
        if not article or not file:
            return render_template("index.html", erreur="Veuillez fournir un article et un fichier.")

        try:
            df = pd.read_excel(file)
            df.columns = [clean_column_name(c) for c in df.columns]

            required = [
                "articles_enfreints",
                "duree_totale_effective_radiation",
                "article_amende/chef",
                "autres_sanctions",
                "nom_de_l'intime",
                "numero_de_decision"
            ]

            if not all(col in df.columns for col in required):
                return render_template("index.html", erreur="Colonnes manquantes après nettoyage.")

            pattern = re.compile(rf"Art\.?\s*{re.escape(article)}(?=[\s\W]|$)", re.IGNORECASE)
            mask = df["articles_enfreints"].astype(str).str.contains(pattern)
            result = df[mask].copy()

            if result.empty:
                return render_template("index.html", erreur=f"Aucun intime trouvé pour l'article {article}.")

            result["Statut"] = "Conforme"

            # Mise en gras et liens markdown
            for col in ["articles_enfreints", "duree_totale_effective_radiation", "article_amende/chef", "autres_sanctions"]:
                result[col] = result[col].apply(lambda x: highlight_article(x, article))

            if "resume" in result.columns:
                result["resume"] = result["resume"].apply(lambda x: f"[Résumé]({x})" if pd.notna(x) else "")

            result = result[[
                "Statut", "numero_de_decision", "nom_de_l'intime", "articles_enfreints",
                "duree_totale_effective_radiation", "article_amende/chef", "autres_sanctions"
            ] + (["resume"] if "resume" in result.columns else [])]

            # Création Excel
            output = BytesIO()
            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                result.to_excel(writer, index=False, sheet_name="Résultats")
                sheet = writer.sheets["Résultats"]
                for i, width in enumerate([30]*len(result.columns)):
                    sheet.set_column(i, i, width)

            output.seek(0)
            markdown_table = result.to_markdown(index=False)
            return render_template("index.html", table=Markup(markdown_table), fichier=output)

        except Exception as e:
            return render_template("index.html", erreur=str(e))

    return render_template("index.html")

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)










































