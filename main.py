
import os
import pandas as pd
import unicodedata
import re
from flask import Flask, request, render_template, send_file
from io import BytesIO
import xlsxwriter

app = Flask(__name__)

# Fonction pour normaliser les noms de colonnes
def normalize_column(col_name):
    if isinstance(col_name, str):
        col_name = unicodedata.normalize('NFKD', col_name).encode('ASCII', 'ignore').decode('utf-8')
        col_name = col_name.replace("’", "'").lower()
        col_name = re.sub(r'\s+', ' ', col_name).strip()
    return col_name

# Fonction pour mettre en rouge les cellules contenant l'article
def detect_article_cell(text, article):
    pattern = rf"art[\.:]?\s*{re.escape(article)}(?=[\s\W]|$)"
    return bool(re.search(pattern, str(text), flags=re.IGNORECASE))

@app.route("/", methods=["GET"])
def index():
    return render_template("index.html")

@app.route("/analyse", methods=["POST"])
def analyse():
    try:
        fichier = request.files["fichier"]
        article = request.form["article"].strip()

        if not fichier or not article:
            return render_template("index.html", erreur="Merci de fournir un fichier et un article.")

        df = pd.read_excel(fichier)
        df.columns = [normalize_column(col) for col in df.columns]

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
                return render_template("index.html", erreur="Le fichier est incomplet. Merci de vérifier la structure.")

        # Filtrage strict de l'article
        pattern = rf"\bArt[\.:]?\s*{re.escape(article)}(?=[\s\W]|$)"
        mask = df["articles enfreints"].astype(str).str.contains(pattern, na=False, flags=re.IGNORECASE)
        conformes = df[mask].copy()
        conformes["Statut"] = "Conforme"

        if conformes.empty:
            return render_template("index.html", erreur=f"Aucun intime trouvé pour l'article {article} demandé.")

        colonnes_finales = [
            "Statut",
            "numero de decision",
            "nom de l'intime",
            "articles enfreints",
            "duree totale effective radiation",
            "article amende/chef",
            "autres sanctions"
        ]

        colonnes_finales = [col for col in colonnes_finales if col in conformes.columns]
        resultat = conformes[colonnes_finales].copy()

        # Création du fichier Excel avec format
        output = BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            resultat.to_excel(writer, index=False, sheet_name="Résultats")
            workbook = writer.book
            worksheet = writer.sheets["Résultats"]

            rouge_format = workbook.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})
            header_format = workbook.add_format({"bold": True, "bg_color": "#D9D9D9", "border": 1})
            wrap_format = workbook.add_format({"text_wrap": True})
            worksheet.set_row(0, None, header_format)

            for i, col in enumerate(resultat.columns):
                worksheet.set_column(i, i, 30, wrap_format)
                for j, value in enumerate(resultat[col]):
                    if detect_article_cell(value, article):
                        worksheet.write(j + 1, i, value, rouge_format)

        output.seek(0)

        # Génération HTML simplifiée (Markdown-like)
        table_md = "| " + " | ".join(colonnes_finales) + " |
"
        table_md += "| " + " | ".join(["---"] * len(colonnes_finales)) + " |
"
        for _, row in resultat.iterrows():
            ligne = [str(cell) for cell in row]
            table_md += "| " + " | ".join(ligne) + " |
"

        return render_template("resultats.html", table_md=table_md, table_excel=output)

    except Exception as e:
        return render_template("index.html", erreur=str(e))

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port)










































