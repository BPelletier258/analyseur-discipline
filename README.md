# 📊 Analyseur de Décisions Disciplinaires

Ce projet propose une application web **Analyseur de Décisions Disciplinaires** permettant à l’utilisateur de rechercher un **article spécifique** dans un fichier Excel de décisions disciplinaires et d’en extraire :

* Un **tableau HTML** interactif
* Un **fichier Excel** formaté et téléchargeable

---

## ✅ Fonctionnalités principales

* **Upload** d’un fichier Excel (.xlsx) contenant les colonnes obligatoires.
* **Saisie** de l’article à filtrer (ex : `14`, `59(2)`, `2.01`).
* **Filtrage strict** : l’article recherché est mis en évidence en rouge **uniquement** dans quatre colonnes ciblées.
* **Affichage HTML** des résultats :

  * Statut (si présent)
  * Numéro de décision
  * Nom de l’intimé
  * **Articles enfreints**
  * **Durée totale effective radiation**
  * **Article amende/chef**
  * **Autres sanctions**
  * **Résumé** (libellé « Résumé » cliquable)
* **Export Excel** formaté :

  * Ligne titre indiquant l’article filtré
  * En‑têtes en **gras** sur fond gris pâle, bordures conservées
  * Colonnes ajustées, **retours à la ligne automatiques**
  * **Coloration rouge** dans les cellules des quatre colonnes suivantes lorsque l’article apparaît :

    1. Articles enfreints
    2. Durée totale effective radiation
    3. Article amende/chef
    4. Autres sanctions
  * Colonne **Résumé** (libellé « Résumé ») avec lien hypertexte
  * Nom du fichier : `decisions_filtrees_<ARTICLE>.xlsx` (ex : `decisions_filtrees_59(2).xlsx`)

---

## 📁 Structure des fichiers

```
analyseur-discipline/
├─ templates/
│  ├─ index.html       # Page HTML du formulaire et rendu du tableau
│  └─ résultats.html   # Template des résultats (tableaux HTML)
├─ main.py             # Application Flask principale
├─ requirements.txt    # Dépendances Python
├─ render.yaml         # Configuration Render pour le déploiement
├─ README.md           # Documentation du projet
└─ LICENSE             # Licence MIT
```

\---------------------------|----------------------------------------------|
\| `numero de decision`      | Numéro unique de la décision                 |
\| `nom de l’intime`         | Nom de la personne sanctionnée               |
\| `articles enfreints`      | Liste des articles enfreints                 |
\| `duree totale effective radiation` | Durée de radiation                    |
\| `article amende/chef`     | Montants d’amende ou chefs de sanction       |
\| `autres sanctions`        | Autres mesures disciplinaires                |
\| **optionnel** `résumé`    | URL vers le résumé de la décision            |

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

1. Connecter le dépôt GitHub à [Render.com](https://render.com).
2. Paramétrer :

   * **Build Command** : laisser vide
   * **Start Command** : `gunicorn main:app`
3. Pousser vos modifications : Render déploie automatiquement.

---

## 📬 Utilisation de l’interface web

1. **Uploader** votre fichier Excel.
2. **Saisir** l’article recherché.
3. Cliquer sur **Analyser**.
4. **Voir** le tableau HTML et **télécharger** le fichier Excel formaté.

**URL de production** : [https://analyseur-discipline.onrender.com](https://analyseur-discipline.onrender.com)

---

## 🧑‍💻 Auteurs et crédits

* Développé par **Assistant GPT** & **Utilisateur** (2025)

<sub>Licence MIT – Voir le fichier LICENSE pour plus de détails.</sub>


<sub>Licence MIT – Voir le fichier LICENSE pour plus de détails.</sub>



