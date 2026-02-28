# Portfolio-Reporting-L3

## Description du Projet

Ce projet contient un **outil d'analyse de portefeuille**. L’outil permet de suivre et d’analyser la performance et les mesures de risque d’un portefeuille composé d’actions (30 titres).

### Fonctionnalités Principales
L’outil offre les fonctionnalités suivantes :
1. **Suivi des performances du portefeuille :**
   - Implémentation d'indicateurs clés de performance : Moyenne des rendements annualisés, ratio de Sharpe, ration d'information (IR).
   - Performance sur une période de 20 ans.

3. **Analyse des risques :**
   - Implémentation d'indicateurs clés de risque : Volatilité annualisée, VaR, TrackingError (TE).
   - Volatilité des actions individuelles et du portefeuille global.

5. **Synthèse globale :**
   - Vue d’ensemble de la performance et des risques du portefeuille.

6. **Base de données :**
   - Importation de fichier CSV créé à partir de Python
   - Création d'une base de données sur Excel à partir des données brutes fournies.

---

## Architecture du Projet

Le projet est organisé en deux parties principales :

### 1. Base de Données (*fichier Importation*)
- **Objectif :** Centralisation et structuration des données pour une utilisation efficace.
- **Étapes :**
  - Un script VBA traite les fichiers bruts et génère des fichiers nettoyés.
- **Résultats attendus :**
  - Deux bases fonctionnelles contenant les actions et du portefeuille et un benchmark.

### 2. Outil de d'Analyse  (*fichier Outil*)
- **Caractéristiques :**
  - Interaction avec les données nettoyées ou la base de données.
  - Implémentation d’une des calculs de performances et de risques
  - Génération d'un rapport 

---

## Contraintes
1. **Données :**
   - Les données des actions sont extraites des fichiers CSV fournis.
   - Le benchmark est donné (CAC 40).

---
