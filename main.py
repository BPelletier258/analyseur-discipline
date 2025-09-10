<!doctype html>
<html lang="fr">
<head>
  <meta charset="utf-8">
  <title>Analyseur Discipline – Filtrage par article</title>
  {{ css|safe }}
</head>
<body>
<div class="wrap">
  <h1>Analyseur Discipline – Filtrage par article</h1>

  <div class="note">
    <strong>Règles :</strong> détection exacte de l’article. Si la 1<sup>re</sup> cellule contient
    « <code>Article filtré : </code> », on ignore la 1<sup>re</sup> ligne (lignes d’en-têtes sur la 2<sup>e</sup>).
  </div>

  {% if error %}
    <div class="note" style="border-color:#fecaca;background:#fff1f2;color:#991b1b;">
      <strong>Erreur :</strong> {{ error }}
    </div>
  {% endif %}

  <form method="post" enctype="multipart/form-data" class="formcard" onsubmit="showSpinner()">
    <div class="form-left">
      <div class="form-row">
        <label for="article">Article à rechercher (ex. <b>29</b>, <b>59(2)</b>)</label>
        <input class="form-input" id="article" name="article" type="text" placeholder="29, 59(2)" required>
      </div>
      <div class="form-row">
        <label>&nbsp;</label>
        <label class="form-check" title="Isoler le segment contenant l’article dans les 4 colonnes d’intérêt">
          <input type="checkbox" name="segment_only" value="1"> Afficher uniquement le segment contenant l’article dans les 4 colonnes d’intérêt
        </label>
      </div>
    </div>

    <div class="file-row">
      <div>
        <label for="file">Fichier Excel</label><br>
        <input id="file" name="file" type="file" accept=".xlsx,.xlsm" required>
      </div>
      <div style="align-self:end;">
        <button class="btn" type="submit" id="analyzeBtn">Analyser</button>
      </div>
    </div>

    <div style="grid-column:1 / -1; color:#6b7280; margin-top:-6px;">
      Formats : <code>.xlsx</code> / <code>.xlsm</code>
    </div>
  </form>

  {% if table_html %}
    <div class="download">
      <a href="{{ url_for('download') }}">Télécharger le résultat (Excel)</a>
    </div>
    {{ table_html|safe }}
  {% endif %}
</div>

<!-- Spinner overlay -->
<div id="overlay"><div class="spinner"></div></div>
<script>
  function showSpinner(){
    document.getElementById('overlay').style.display='flex';
    const btn = document.getElementById('analyzeBtn');
    if(btn){ btn.disabled = true; }
  }
</script>
</body>
</html>
