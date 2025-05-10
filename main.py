
import os
import re
import pandas as pd
from flask import Flask, request, render_template, send_file, redirect, flash
from werkzeug.utils import secure_filename
from io import BytesIO
import xlsxwriter

app = Flask(__name__)
app.secret_key = 'secret'
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def normalize_column(col_name):
    import unicodedata
    if isinstance(col_name, str):
        col_name = unicodedata.normalize('NFKD', col_name).encode('ASCII', 'ignore').decode('utf-8')
        col_name = col_name.replace("’", "'").lower()
        col_name = re.sub(r'\s+', ' ', col_name).strip()
    return col_name

def highlight_article(cell, article_pattern):
    if isinstance(cell, str) and re.search(article_pattern, cell, flags=re.IGNORECASE):
        return True
    return False

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files.get('fichier')
        article = request.form.get('article', '').strip()

        if not file or not article:
            flash("Veuillez fournir un fichier et un numéro d'article.")
            return redirect(request.url)

        try:
            df = pd.read_excel(file)
            df.columns = [normalize_column(col) for col in df.columns]

            required_cols = [
                "articles enfreints", "duree totale effective radiation", "article amende/chef",
                "autres sanctions", "nom de l'intime", "numero de decision"
            ]
            for col in required_cols:
                if col not in df.columns:
                    return render_template("index.html", erreur=f"Colonne manquante : {col}")

            article_pattern = rf'Art[\.:\s]*{re.escape(article)}(?=[\s\W]|$)'
            mask = df["articles enfreints"].astype(str).str.contains(article_pattern, na=False, flags=re.IGNORECASE)
            conformes = df[mask].copy()
            conformes['Statut'] = "Conforme"

            if conformes.empty:
                return render_template("index.html", erreur=f"Aucun intime trouvé pour l'article {article} demandé.")

            # Ajout du lien cliquable (Résumé) s'il existe
            if 'resume' in conformes.columns:
                conformes['Résumé'] = conformes['resume'].apply(
                    lambda url: f"[Résumé]({url})" if pd.notna(url) and isinstance(url, str) else ""
                )

            # Création du tableau Markdown
            colonnes_affichees = [
                "Statut", "numero de decision", "nom de l'intime", "articles enfreints",
                "duree totale effective radiation", "article amende/chef", "autres sanctions"
            ]
            if 'Résumé' in conformes.columns:
                colonnes_affichees.append('Résumé')

            markdown_table = conformes[colonnes_affichees].copy()
            markdown = markdown_table.to_markdown(index=False)

            # Création du fichier Excel
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                conformes[colonnes_affichees].to_excel(writer, index=False, sheet_name='Résultats')
                workbook = writer.book
                worksheet = writer.sheets['Résultats']
                red_format = workbook.add_format({'bg_color': '#FFC7CE'})

                for row_idx, row in enumerate(conformes[colonnes_affichees].itertuples(index=False), start=1):
                    for col_idx, value in enumerate(row):
                        if isinstance(value, str) and re.search(article_pattern, value, flags=re.IGNORECASE):
                            worksheet.write(row_idx, col_idx, value, red_format)
                        else:
                            worksheet.write(row_idx, col_idx, value)
                worksheet.set_row(0, 30)
                for i in range(len(colonnes_affichees)):
                    worksheet.set_column(i, i, 30)

            output.seek(0)

            return render_template("resultats.html", table_html=markdown, article=article,
                                   fichier_excel='resultats.xlsx', fichier_data=output.read())

        except Exception as e:
            return render_template("index.html", erreur=str(e))

    return render_template("index.html")

@app.route('/telecharger')
def telecharger():
    fichier_data = request.args.get('fichier_data')
    if not fichier_data:
        return "Fichier manquant", 400
    output = BytesIO(fichier_data.encode('latin1'))
    return send_file(output, download_name="resultats.xlsx", as_attachment=True)

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)










































