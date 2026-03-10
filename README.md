# Planification Nutcache — Add-in Excel

## Description

Add-in Excel qui permet d'importer les rapports Nutcache (WorkedHoursSummary) dans le classeur maître de planification. L'outil détecte automatiquement les projets nouveaux et existants, crée les onglets manquants et met à jour les données.

---

## Installation (étape par étape)

### Prérequis

- **Excel** : Microsoft 365 (desktop Windows ou Mac)
- **Node.js** : version 18+ (https://nodejs.org)

### 1. Installer les dépendances

Ouvrir un terminal dans le dossier `add-in-planification/` :

```bash
npm install
```

### 2. Générer les certificats HTTPS

Excel exige un serveur HTTPS. Générez les certificats locaux :

```bash
npm run generate-certs
```

Cela installe un certificat de développement approuvé par votre machine.

### 3. Démarrer le serveur local

```bash
npm start
```

Le serveur démarre sur `https://localhost:3000`.

### 4. Charger l'Add-in dans Excel

**Option A — Sideload automatique :**
```bash
npm run sideload
```

**Option B — Chargement manuel macOS :**
1. Ouvrir le **Finder** → `Cmd + Shift + G`
2. Coller : `/Users/VOTRE_NOM/Library/Containers/com.microsoft.Excel/Data/Documents/wef`
3. Si le dossier `wef` n'existe pas → le créer
4. Copier `manifest.xml` dans ce dossier
5. Fermer et rouvrir Excel
6. Onglet **Insertion** → **Mes compléments** → l'Add-in apparaît sous **Developer Add-ins**
7. ⚠️ Sur Mac, il faut réactiver l'Add-in à chaque redémarrage d'Excel

**Option C — Chargement manuel Windows :**
1. Ouvrir Excel → onglet **Insertion** → **Compléments** → **Autres compléments**
2. Cliquer **Charger mon complément** → sélectionner `manifest.xml`

**Option D — Dossier réseau partagé (déploiement équipe, Windows) :**
1. Placer `manifest.xml` dans un dossier réseau partagé
2. Excel → **Fichier** → **Options** → **Centre de gestion de la confidentialité**
3. → **Catalogues de compléments approuvés** → ajouter le chemin réseau

### 5. Utiliser l'Add-in

1. L'onglet **"Planification"** apparaît dans le ruban Excel
2. Cliquer le bouton **"Importer Nutcache"** → le panneau latéral s'ouvre
3. Sélectionner le rapport Nutcache (.xlsx)
4. Vérifier les projets détectés (badges NOUVEAU / MISE À JOUR)
5. Cliquer **"Importer dans le classeur"**

---

## Structure des fichiers

```
add-in-planification/
├── manifest.xml       ← Déclaration de l'Add-in pour Excel
├── taskpane.html      ← Interface utilisateur (panneau latéral)
├── taskpane.css       ← Styles de l'interface
├── taskpane.js        ← (intégré dans taskpane.html)
├── parser.js          ← Parse les rapports Nutcache
├── updater.js         ← Logique de création/MAJ des onglets
├── package.json       ← Dépendances et scripts npm
├── README.md          ← Ce fichier
└── assets/
    ├── icon-16.png    ← Icônes pour le ruban (à fournir)
    ├── icon-32.png
    ├── icon-64.png
    └── icon-80.png
```

---

## Icônes personnalisées

Fournir 4 fichiers PNG (fond transparent recommandé) :
- `icon-16.png` — 16×16 px
- `icon-32.png` — 32×32 px
- `icon-64.png` — 64×64 px
- `icon-80.png` — 80×80 px

Placer dans le dossier `assets/`.

---

## Format attendu du rapport Nutcache

Le fichier doit être un export **"Sommaire des heures travaillées par projet, membre et service"** de Nutcache, au format `.xlsx`.

Structure reconnue :
```
Ligne projet :  "Client - Code5chiffres - Nom du projet"
Ligne membre :  "NomMembre"
Ligne header :  "Service | Heures travaillées | Coûts MO | ..."
Lignes données: "NomService | heures | coûts | ..."
Ligne total :   "Total pour le membre | heures | coûts | ..."
Ligne total :   "Total pour le projet | heures | coûts | ..."
```

---

## Comportement de l'import

### Projet existant (code présent dans l'Index Projets + onglet existant)
- ✅ Met à jour : heures réelles, coûts, taux horaire déduit
- ✅ Met à jour : détail par phase/service
- 🔒 Ne touche PAS : budget, heures planifiées, dates, CP, notes

### Nouveau projet (code absent)
- ✅ Duplique l'onglet `_TEMPLATE`
- ✅ Remplit : code, nom, client, membres, phases
- ✅ Ajoute à l'Index Projets (statut "En cours", priorité "Moyenne")
- ✅ Ajoute au Tableau de bord (formules dynamiques)
- ⚠️ À compléter manuellement : budget, dates, CP, heures planifiées

---

## Dépannage

| Problème | Solution |
|----------|----------|
| L'onglet "Planification" n'apparaît pas | Vérifier que le serveur tourne (`npm start`) et que le manifest est bien chargé |
| Erreur de certificat | Relancer `npm run generate-certs` et accepter le certificat |
| Fichier non reconnu | Vérifier le format : doit être un export Nutcache "WorkedHoursSummary" |
| Onglet _TEMPLATE introuvable | S'assurer que le classeur maître contient l'onglet `_TEMPLATE` (masqué) |
