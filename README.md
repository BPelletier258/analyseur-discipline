# Analyseur Discipline – Filtrage par article

Application Flask qui lit un fichier Excel (.xlsx / .xlsm), filtre les décisions disciplinaires par *article* et affiche le résultat en HTML avec possibilité d’export en Excel.

---

## Fonctionnalités

- **Recherche d’un article** (ex. `29`, `59(2)`).
- **Filtrage strict des lignes** : on **ne garde que** les lignes où l’article recherché apparaît **dans au moins une des 4 colonnes d’intérêt** (voir ci-dessous).
- **Option “Afficher uniquement le segment contenant l’article dans les 4 colonnes d’intérêt”** :
  - Si cochée, dans ces 4 colonnes on **isole uniquement** les items contenant l’article (au lieu d’afficher le texte complet de la cellule).
- **Mise en valeur (HTML & Excel)** :
  - HTML : l’article est surligné (rouge) **uniquement** dans les 4 colonnes d’intérêt.
  - Excel : l’article est surligné **à l’intérieur des cellules** (rich text) dans les 4 colonnes d’intérêt.
- **Colonnes rendues en listes à puces** (lisibilité accrue).
- **Formatage automatique des montants (“Total amendes”)** : `500` → `500 $`, `5000` → `5 000 $`.
- **Export Excel propre** :
  - Ligne 1 : `Article filtré : X`
  - Ligne 2 : en-têtes stylés
  - Fige la ligne d’en-têtes
  - Retours à la ligne + alignement vertical en haut
  - Largeurs ajustées automatiquement (max ~60)

---

## Colonnes d’intérêt

Les **4 colonnes d’intérêt** (ciblées par l’option et par le surlignage) :

1. `Nbr Chefs par articles`
2. `Nbr Chefs par articles par période de radiation`
3. `Nombre de chefs par articles et total amendes`
4. `Nombre de chefs par article ayant une réprimande`

> Le filtrage des lignes utilise **exclusivement ces 4 colonnes** : si l’article recherché n’apparaît dans **aucune** d’elles, la ligne est **exclue**.

### Colonnes explicitement *non* surlignées (HTML)

Même si l’article est présent, **pas** de surlignage rouge dans :
- `Liste des chefs et articles en infraction`
- `Liste des sanctions imposées`

---

## Colonnes rendues en puces

Les cellules des colonnes ci-dessous sont rendues en **liste à puces** (HTML) lorsque plusieurs items sont détectés :

- `Résumé des faits concis`
- `Liste des chefs et articles en infraction`
- `Nbr Chefs par articles`
- `Nbr Chefs par articles par période de radiation`
- `Liste des sanctions imposées`
- `Nombre de chefs par article ayant une réprimande`
- `Autres mesures ordonnées`
- `À vérifier`

---

## Détection / entêtes

- Si la **1re cellule** (A1) d’un fichier Excel contient `Article filtré :`, la **1re ligne** est considérée comme un bandeau (titre) et **ignorée** : les en-têtes sont donc sur la **2e ligne**.
- Formats pris en charge : `.xlsx` et `.xlsm`.

---

## Installation locale

**Prérequis** : Python 3.11+

```bash
python -m venv .venv
source .venv/bin/activate     # (Windows: .venv\Scripts\activate)
pip install -r requirements.txt
