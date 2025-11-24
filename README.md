# Exemples d’automatisation Excel VBA

Ce dépôt contient plusieurs exemples de macros Excel/VBA prêtes à l’emploi pour illustrer des scénarios d’automatisation courants en entreprise.

Les exemples sont simples, pédagogiques et peuvent servir de base pour vos propres développements.

---

## 📂 Contenu

Le classeur `automation-examples.xlsm` contient 3 macros principales :

1. **Fusionner plusieurs fichiers Excel d’un dossier**  
   - Parcourt tous les fichiers `.xlsx` d’un dossier
   - Copie les données dans une feuille unique `Fusion`
   - Ajoute automatiquement à la première ligne vide

2. **Nettoyer un tableau de données**  
   - Supprime les lignes complètement vides
   - Supprime les espaces superflus (TRIM)
   - Met en forme le texte (optionnel)

3. **Générer un petit reporting à partir d’un tableau de ventes**  
   - Lit les données d’une feuille `Ventes`
   - Calcule le total par commercial
   - Écrit une synthèse dans la feuille `Reporting`

---

## ✅ 1. Fusionner plusieurs fichiers Excel

- Macro : `FusionnerFichiersDossier`
- Feuille cible : `Fusion`
- Hypothèses :
  - chaque fichier source contient des données sur la feuille 1
  - les en-têtes se trouvent sur la première ligne

---

## ✅ 2. Nettoyer un tableau

- Macro : `NettoyerTableau`
- Feuille cible : `Données`
- Hypothèses :
  - les en-têtes sont sur la ligne 1
  - les données commencent en ligne 2

---

## ✅ 3. Générer un reporting simple

- Macro : `GenererReporting`
- Feuille source : `Ventes`
- Feuille cible : `Reporting`
- Hypothèses :
  - colonnes :
    - A : Commercial
    - B : Montant

---

## 🧰 Prérequis

- Microsoft Excel (version desktop)
- Macros activées (`.xlsm`)

---

## ▶️ Utilisation

1. Ouvrir `automation-examples.xlsm`
2. Activer les macros
3. Ouvrir l’éditeur VBA (`Alt + F11`) pour consulter le code
4. Depuis Excel : `Outils > Macro > Macros…`  
   - Sélectionner la macro souhaitée  
   - Cliquer sur **Exécuter**

---

## ⚠️ Avertissement

Ces macros sont fournies à titre **pédagogique**.  
Elles ne contiennent aucun code client réel et doivent être adaptées à vos propres fichiers avant utilisation en production.

---

## 👤 Auteur

**Matthieu Chenal – DOPHIS**  
Expert Excel & VBA – Développement d’outils métier, automatisation & optimisation de fichiers.
