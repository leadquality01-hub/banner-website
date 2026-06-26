# Scraper Cabinets d'Avocat - Paris Métropole

Base de données de 500 cabinets d'avocat parisiens pour prospection email.

## Structure des fichiers

```
scraper/
├── scrape_cabinets_avocat.py   ← Script principal
├── test_scraper.py              ← Test rapide avant de lancer
├── cabinets_avocat_paris.csv    ← Résultat (créé après le scraping)
└── cabinets_avocat_paris.json   ← Résultat JSON (créé après le scraping)
```

## Installation

```bash
pip install requests beautifulsoup4 pandas lxml
```

## Utilisation

### Étape 1 : Tester d'abord
```bash
cd scraper
python test_scraper.py
```
Si le test affiche des cabinets → tout est OK, passe à l'étape 2.

### Étape 2 : Lancer le scraping complet
```bash
python scrape_cabinets_avocat.py
```
Durée estimée : **30 à 60 minutes** (pauses intégrées pour respecter le site).

## Colonnes du CSV produit

| Colonne | Description |
|---|---|
| `nom` | Nom du cabinet |
| `adresse` | Adresse complète |
| `telephone` | Numéro de téléphone |
| `email` | Email (quand disponible) |
| `site_web` | URL du site web |
| `specialite` | Domaine de spécialité |
| `zone` | Arrondissement / commune |
| `url_fiche` | Lien vers la fiche Pages Jaunes |
| `statut_contact` | "à contacter" par défaut (à mettre à jour manuellement) |
| `date_contact` | Date du dernier contact |
| `notes` | Notes libres |

## Conseil pour les emails

Les emails sont rarement affichés sur Pages Jaunes. Pour enrichir la base :
1. Utilise l'URL `site_web` pour visiter leur site et trouver la page Contact
2. Outils d'enrichissement email : **Hunter.io**, **Dropcontact**, **Kaspr** (versions gratuites disponibles)
3. Recherche manuelle : `"nom du cabinet" + "contact" site:.fr`

## Respect des CGU

- Le scraper inclut des délais aléatoires entre les requêtes (1.5–6 secondes)
- Usage strictement personnel / prospection B2B
- Ne pas lancer plusieurs instances en parallèle
