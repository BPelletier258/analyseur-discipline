from flask import Flask, request, render_template, send_file
from io import BytesIO
import pandas as pd, re, unicodedata, xlsxwriter

app = Flask(__name__)

def normalize(col):
    s = unicodedata.normalize('NFKD', str(col))
    return ''.join(ch for ch in s if not unicodedata.combining(ch)).lower().strip()

def style_and_filter(df, article):
    # Normalisation des noms de colonnes
    df.columns = [normalize(c) for c in df.columns]
    df.rename(columns={'nom de lintime':"nom de l'intime"}, inplace=True)
    # Filtrage strict
    pattern = re.compile(rf'Art[.:]\s*{re.escape(article)}(?=\D|$)', re.IGNORECASE)
    mask = df['articles enfreints'].astype(str).apply(lambda x: bool(pattern.search(unicodedata.normalize('NFKD', x))))
    filtered = df.loc[mask].copy()
    return filtered, pattern

@app.route('/', methods=['GET', 'POST'])
def analyse():
    if request.method == 'POST':
        file = request.files.get('file')
        article = request.form.get('article','').strip()
        if not file or not article:
            return render_template('index.html', erreur="Veuillez fournir un fichier et un article.")
        # Lecture du Excel en mémoire
        df = pd.read_excel(file)
        filtered, pattern = style_and_filter(df, article)
        if filtered.empty:
            return render_template('index.html', erreur=f"Aucun résultat pour l'article {article}.")
        # Génération Markdown
        markdown_table = filtered[['numero de decision','nom de l\'intime','articles enfreints']].to_markdown(index=False)
        # Génération Excel en mémoire
        output = BytesIO()
        wb = xlsxwriter.Workbook(output, {'in_memory':True})
        ws = wb.add_worksheet('Résultats')
        # Formats
        wrap = wb.add_format({'text_wrap':True})
        red = wb.add_format({'font_color':'#FF0000','text_wrap':True})
        # Header
        cols = filtered.columns.tolist()
        for i, col in enumerate(cols):
            ws.write(0,i,col, wb.add_format({'bold':True}))
            ws.set_column(i,i,30)
        # Data
        for r, row in enumerate(filtered.itertuples(index=False), start=1):
            for c, val in enumerate(row):
                txt = str(val)
                fmt = red if pattern.search(txt) else wrap
                ws.write(r, c, txt, fmt)
        wb.close()
        output.seek(0)
        return render_template('index.html',
            markdown=markdown_table,
            excel_data=output.read(),
            filename=f"resultats_{article}.xlsx"
        )
    # GET
    return render_template('index.html')

@app.route('/download/<filename>')
def download(filename):
    data = request.args.get('data')
    # non nécessaire si on encode directement dans le template via data URI
    return send_file(BytesIO(data.encode('latin1')), attachment_filename=filename)

if __name__=='__main__':
    app.run()



















































































