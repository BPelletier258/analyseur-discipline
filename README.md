# 📊 Analyseur de Décisions Disciplinaires

Cette application web permet de rechercher un article précis dans un fichier Excel contenant des décisions disciplinaires, et de générer dynamiquement un tableau HTML avec les résultats filtrés.

## ✅ Fonctionnalités

- Chargement d’un fichier Excel (.xlsx)
- Recherche stricte d’un **article** dans la colonne `Articles enfreints`
- Affichage dynamique des colonnes suivantes :
  - **Statut**
  - **Numéro de décision**
  - **Nom de l’intimé**
  - **Articles enfreints** (avec l’article surligné en rouge)
  - **Périodes de radiation**
  - **Amendes**
  - **Autres sanctions**
  - **Résumé** (hyperlien cliquable vers le résumé de la décision)
- Mise en surbrillance rouge de l’article recherché dans toutes les colonnes où il est mentionné.
- Requête possible via l’interface ou par API externe (ex: Postman)

## 📁 Structure des fichiers attendus

Les colonnes suivantes doivent être présentes dans le fichier Excel :

- `nom de l'intime`
- `numero de decision`
- `articles enfreints`
- `duree totale effective radiation`
- `article amende/chef`
- `autres sanctions`
- *(optionnel)* `resume` (URL vers le résumé)

## 🛠 Déploiement

L'application peut être déployée sur [Render.com](https://render.com) ou exécutée localement avec Flask.

### Exemple de commande locale

```bash
python main.py
### Dépendances (voir `requirements.txt`)

- Flask
- pandas
- openpyxl

## 📬 Utilisation de l’interface

1. Uploade un fichier Excel valide.
2. Entre un article (ex: `14`, `59(2)`, `2.01 a)`).
3. Clique sur **Analyser**.
4. Le tableau HTML s’affiche avec les résultats filtrés.

## 🧠 Auteurs et crédits

Assistant GPT + Collaboration utilisateur – 2025

