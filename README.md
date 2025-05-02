# 🧾 Analyseur d'Article Disciplinaire

Cette application Flask permet d’analyser un fichier Excel disciplinaire afin d’extraire uniquement les décisions où un article spécifique a été **explicitement** mentionné dans la colonne **« articles enfreints »**.

---

## 🚀 Démonstration en ligne

🔗 [Application hébergée sur Render](https://analyseur-discipline.onrender.com)

---

## 📂 Fonctionnalité principale

- 📤 Import d’un fichier Excel disciplinaire (`.xlsx`)
- 🔍 Recherche stricte d’un article (ex: `14`, `59(2)`, `2.01 a)`, `149.1`, etc.)
- 📊 Affichage du tableau des résultats :
  - **Numéro de décision**
  - **Nom de l’intimé**
  - **Articles enfreints** (avec mise en évidence en rouge et gras de l’article recherché)
  - **Période de radiation**
  - **Amendes**
  - **Autres sanctions**

---

## 🧪 Articles compatibles

Voici quelques formats valides :
14
59(2)
59.2
2.01 a)
149.1
2.01 b)
3.01.02

---

## ⚙️ Lancer en local

### 1. Cloner le dépôt

```bash
git clone https://github.com/<votre-utilisateur>/analyseur-discipline.git
cd analyseur-discipline
2. Créer un environnement virtuel (optionnel mais recommandé)
python -m venv venv
source venv/bin/activate  # ou venv\Scripts\activate sur Windows
3. Installer les dépendances
pip install -r requirements.txt
4. Lancer l’application localement
python main.py
Ensuite, aller sur http://localhost:5000 dans votre navigateur.

🧾 Fichiers importants

main.py → logique principale (Flask, parsing, filtrage strict)
templates/index.html → interface web
requirements.txt → dépendances (flask, pandas, openpyxl, gunicorn)
render.yaml → déploiement Render
README.md → ce fichier
📌 Remarques

L’article recherché doit être strictement présent dans la colonne "articles enfreints"
Les articles sont identifiés même s’ils sont écrits avec des variations (Art. 59(2), ART 59.2, art 59(2)…)
📤 Déploiement

L’app est compatible avec Render, grâce au fichier render.yaml :

startCommand: "python main.py"
📩 Pour toute question

Ce projet a été guidé par une logique stricte d’analyse textuelle avec validation étape par étape.
Pour tout commentaire ou suggestion, ouvrez un Issue GitHub.
