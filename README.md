# 📊 Analyseur de Décisions Disciplinaires

Ce projet propose une application web **Analyseur de Décisions Disciplinaires** permettant à l’utilisateur de rechercher un **article spécifique** dans un fichier Excel de décisions disciplinaires et d’en extraire un tableau HTML et un fichier Excel formaté.

---

## ✅ Fonctionnalités

* **Upload** d’un fichier Excel (.xlsx) contenant les colonnes obligatoires.
* **Saisie** de l’article à filtrer (ex : `14`, `59(2)`, `2.01`).
* **Filtrage strict** : l’article recherché est mis en évidence uniquement dans les colonnes pertinentes.
* **Affichage HTML** des résultats :

  * Statut (si présent)
  * Numéro de décision
  * Nom de l’intimé
  * Articles enfreints (article surligné en rouge)
  * Durée totale effective radiation
  * Article amende/chef
  * Autres sanctions
  * Résumé (lien cliquable)
* **Export Excel** formaté :

  * Ligne titre indiquant l’article filtré
  * En‑têtes en gras sur fond gris
  * Colonnes ajustées, retours à la ligne automatiques
  * Texte en rouge dans les cellules des colonnes « Articles enfreints », « Durée… », « Article amende/chef », « Autres sanctions » lorsque l’article apparaît
  * Colonne « Résumé » avec un libellé « Résumé » et un lien hypertexte
* **Téléchargement** automatique d’un fichier nommé `decisions_filtrees_<ARTICLE>.xlsx` (ex : `decisions_filtrees_59(2).xlsx`).

---

## 📁 Structure des fichiers

* `main.py` : code principal Flask
* `requirements.txt` : dépendances Python
* `index.html` : template HTML si séparé (ou `render_template_string` embarqué)
* `render.yaml` : configuration Render pour le déploiement
* `README.md` : documentation

\-----------------------------------------------|---------------------------------------------|
\| `numero de decision`                          | Numéro unique de la décision                |
\| `nom de l’intime`                             | Nom de la personne sanctionnée              |
\| `articles enfreints`                          | Liste des articles enfreints                |
\| `duree totale effective radiation`            | Durée de radiation                          |
\| `article amende/chef`                         | Montants d’amende ou chefs de sanction      |
\| `autres sanctions`                            | Autres mesures disciplinaires               |
\| **optionnel** `résumé`                        | URL vers le résumé de la décision            |

---

## 🛠 Installation et exécution locale

1. Cloner le dépôt :

   ```bash
   git clone https://github.com/<votre-utilisateur>/analyseur-discipline.git
   cd analyseur-discipline
   ```

2. Installer les dépendances :

   ```bash
   pip install -r requirements.txt
   ```

3. Lancer l’application :

   ```bash
   python main.py
   ```

4. Ouvrir dans votre navigateur : [http://127.0.0.1:5000](http://127.0.0.1:5000)

---

## ☁️ Déploiement sur Render

1. Créer un projet sur [Render.com](https://render.com) et connecter votre dépôt GitHub.
2. Définir :

   * **Build Command** : *(laisser vide)*
   * **Start Command** : `gunicorn main:app`
3. Ajouter un `render.yaml` (optionnel) ou configurer via l’UI.
4. Pousser vos modifications ; Render déploie automatiquement.

---

## 📬 Utilisation

1. **Uploader** votre fichier Excel.
2. **Saisir** l’article recherché.
3. Cliquer **Analyser**.
4. **Voir** le tableau HTML et **télécharger** le fichier Excel formaté.

**URL de production** : [https://analyseur-discipline.onrender.com](https://analyseur-discipline.onrender.com)

---

## 🧑‍💻 Auteurs et crédits

* Développé par l’Assistant GPT & Utilisateur (2025)

---

<small>Licence MIT – Voir le fichier LICENSE pour plus d’informations.</small>


