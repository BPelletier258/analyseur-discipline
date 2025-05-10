from flask import Flask, request, render_template, send_file
import pandas as pd
import re
import unicodedata
from io import BytesIO
import xlsxwriter

app = Flask(__name__)

def normalize_column(col_name):
    if isinstance(col_name, str):
        col_name = unicodedata.normalize('NFKD', col_name).encode('ASCII', 'ignore').decode('utf-8')
        col_name = col_name.replace("’", "'").lower()
        col_name = re.sub(r'\s+', ' ', col_name).strip()
    return col_name

def color_cells(val, article):
    article_pattern = rf"Art[\.\:]?\s*{re.escape(article)}(?=[\s\W]|$)"
    if pd.notna(val) and re.search(article_pattern, str(val), flags=re.IGNORECASE):
        return {'bg_color': '#FFC7CE'}
    return {}

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/analyse", methods=["POST"])
def analyse():
    try:
        fichier = request.files["file"]
        article = request.form["article"].strip()

        df = pd.read_excel(fichier)
        df.columns = [normalize_column(c) for c in df.columns]

        colonnes_requises = [
            "articles enfreints", "duree totale effective radiation", "article amende/chef",
            "autres sanctions", "nom de l'intime", "numero de decision"
        ]
        for col in colonnes_requises:
            if col not in df.columns:
                raise ValueError("Le fichier est incomplet. Merci de vérifier la structure.")

        pattern_explicit = rf"Art[\.\:]?\s*{re.escape(article)}(?=[\s\W]|$)"
        mask = df["articles enfreints"].astype(str).str.contains(pattern_explicit, na=False, flags=re.IGNORECASE)
        resultats = df[mask].copy()

        if resultats.empty:
            raise ValueError(f"Aucun intime trouvé pour l'article {article} demandé.")

        resultats["Statut"] = "Conforme"

        colonnes_finales = [
            "Statut", "numero de decision", "nom de l'intime",
            "articles enfreints", "duree totale effective radiation",
            "article amende/chef", "autres sanctions"
        ]

        if "resume" in resultats.columns:
            colonnes_finales.append("resume")

        resultats = resultats[colonnes_finales]

        # Export Excel
        output = BytesIO()
        writer = pd.ExcelWriter(output, engine='xlsxwriter')
        resultats.to_excel(writer, sheet_name='Résultats', index=False)
        workbook = writer.book
        worksheet = writer.sheets['Résultats']
        header_format = workbook.add_format({'bold': True, 'bg_color': '#DDDDDD'})
        for col_num, value in enumerate(resultats.columns.values):
            worksheet.write(0, col_num, value, header_format)
            worksheet.set_column(col_num, col_num, 30)

        for row_num, row in resultats.iterrows():
            for col_num, col in enumerate(resultats.columns):
                val = row[col]
                fmt = color_cells(val, article)
                if fmt:
                    cell_format = workbook.add_format(fmt)
                    worksheet.write(row_num + 1, col_num, val, cell_format)

        writer.close()
        output.seek(0)

        return send_file(output, as_attachment=True, download_name="resultats.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        return render_template("index.html", erreur=str(e))










































