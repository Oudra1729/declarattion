# Application de Gestion de Déclarations de Transport

## 📋 Description

Cette application web permet de gérer et générer des déclarations de transport pour des produits réglementés. L'application utilise des fichiers Excel (.xlsx) comme source unique de données, stockés dans le dossier `data/` du projet.

## 🚀 Fonctionnalités Principales

### 1. Gestion des Données de Base

L'application gère quatre types de données principales :

- **Clients** : Informations sur les clients (nom, destination, antenne, itinéraire)
- **Conducteurs** : Informations sur les conducteurs et leurs véhicules (nom, CIN, téléphone, matricule, modèle)
- **Convoyeurs** : Informations sur les convoyeurs (nom, CIN, téléphone, CCE)
- **Produits** : Liste des produits transportables (nom, unité de mesure)

### 2. Création de Déclarations

L'application permet de créer des déclarations de transport complètes incluant :

- Informations client (destination, antenne, itinéraire)
- Informations conducteur et véhicule
- Informations convoyeur
- Liste des produits transportés avec quantités
- Numéro de document (auto-incrémenté)
- Dates (date de déclaration et date de départ)
- Numéro de passavant et date d'expiration
- Bon de livraison (optionnel)

### 3. Gestion des Fichiers Excel

#### Structure des Fichiers

Tous les fichiers Excel sont stockés dans le dossier `data/` :

- `clients.xlsx` : Liste des clients
- `drivers.xlsx` : Liste des conducteurs
- `convoyeurs.xlsx` : Liste des convoyeurs
- `products.xlsx` : Liste des produits
- `history.xlsx` : Historique des déclarations

#### Fonctionnement

1. **Chargement des Données** :
   - Priorité 1 : Données depuis `localStorage` (pour fonctionnement hors ligne)
   - Priorité 2 : Chargement depuis les fichiers Excel dans `data/`

2. **Sauvegarde des Données** :
   - Lors de l'ajout d'une nouvelle entité (client, conducteur, etc.), la donnée est :
     - Sauvegardée dans `localStorage`
     - Ajoutée comme nouvelle ligne dans le fichier Excel correspondant dans `data/`
   - Les données existantes sont préservées (pas de remplacement)

3. **Export/Import** :
   - Export de tous les fichiers Excel pour sauvegarde
   - Import depuis des fichiers Excel pour restaurer ou fusionner des données
   - Fusion de données (évite la perte de données existantes)

## 💾 Stockage des Données

### localStorage

Les données sont d'abord stockées dans le `localStorage` du navigateur pour :
- Fonctionnement hors ligne
- Accès rapide aux données
- Synchronisation avec les fichiers Excel

### Fichiers Excel

Les fichiers Excel dans `data/` sont la source de vérité principale :
- Format : .xlsx (Excel)
- Emplacement : `data/` dans le projet
- Sauvegarde directe : Utilise File System Access API (Chrome/Edge) pour sauvegarder directement dans les fichiers

## 🔧 Utilisation

### Première Utilisation

1. Ouvrir `index.html` dans un navigateur (Chrome ou Edge recommandé)
2. Si c'est la première fois, sélectionner le dossier `data/` lorsque demandé
3. Les données seront chargées depuis les fichiers Excel existants

### Ajouter des Données

1. Utiliser le bouton **"➕ Ajouter Rapide"** en haut de la page
2. Ou utiliser les boutons **"＋"** à côté de chaque champ de sélection
3. Remplir le formulaire et cliquer sur **"Enregistrer"**
4. La donnée sera automatiquement ajoutée au fichier Excel correspondant

### Créer une Déclaration

1. Remplir les informations client (sélectionner depuis la liste)
2. Remplir les informations conducteur et véhicule
3. Remplir les informations convoyeur
4. Ajouter les produits transportés
5. Remplir les informations de passavant
6. Le "Bon de Livraison" est optionnel
7. Cliquer sur **"🎉 Générer la Déclaration"**
8. La déclaration sera générée et sauvegardée dans `history.xlsx`

### Gestion des Données

Utiliser le bouton **"💾 Gestion Données"** pour :
- Exporter tous les fichiers Excel
- Importer des données depuis Excel
- Fusionner des données (pour partager entre machines)

## 📁 Structure du Projet

```
project mvp/
├── index.html          # Page principale de l'application
├── declaration.html    # Page d'affichage de la déclaration générée
├── script.js          # Logique principale de l'application
├── style.css          # Styles CSS
├── data/              # Dossier contenant les fichiers Excel
│   ├── clients.xlsx
│   ├── drivers.xlsx
│   ├── convoyeurs.xlsx
│   ├── products.xlsx
│   └── history.xlsx
└── README.md          # Ce fichier
```

## 🌐 Compatibilité Navigateurs

- **Chrome/Edge** (recommandé) : Support complet du File System Access API pour sauvegarde directe
- **Autres navigateurs** : Fonctionne mais télécharge les fichiers au lieu de sauvegarder directement

## ⚙️ Technologies Utilisées

- HTML5 / CSS3
- JavaScript (ES6+)
- SheetJS (XLSX) : Pour la manipulation des fichiers Excel
- File System Access API : Pour la sauvegarde directe des fichiers (Chrome/Edge)

## 📝 Notes Importantes

1. **Sauvegarde Automatique** : Les données sont automatiquement sauvegardées dans les fichiers Excel lors de l'ajout
2. **Pas de Base de Données** : L'application utilise uniquement Excel comme source de données
3. **Hors Ligne** : L'application fonctionne hors ligne grâce à `localStorage`
4. **Partage de Données** : Utiliser Export/Import pour partager des données entre machines
5. **Bon de Livraison** : Ce champ est optionnel et peut être laissé vide

## 🔄 Synchronisation

L'application synchronise automatiquement :
- `localStorage` ↔ Fichiers Excel dans `data/`
- Les données ajoutées sont immédiatement disponibles dans les listes déroulantes
- L'historique des déclarations est sauvegardé automatiquement

## 🆘 Dépannage

### Les données ne se chargent pas
- Vérifier que les fichiers Excel existent dans `data/`
- Utiliser "Gestion Données" → "Import Excel" pour charger les données

### Les données ne se sauvegardent pas
- Vérifier que vous utilisez Chrome ou Edge
- Sélectionner le dossier `data/` lorsque demandé
- Vérifier les permissions du navigateur

### Erreur CORS
- L'application doit être ouverte via un serveur local ou utiliser Chrome/Edge avec File System Access API
- Les données sont chargées depuis `localStorage` en priorité, donc l'application fonctionne même avec cette limitation

---

**Version** : 1.0  
**Dernière mise à jour** : 2025

