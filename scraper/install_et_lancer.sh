#!/bin/bash
# =========================================================
#  Installation et lancement du scraper - Cabinets Avocat
#  Lance ce script sur TA MACHINE LOCALE (pas dans le cloud)
# =========================================================

echo ""
echo "================================================"
echo "  Scraper Cabinets Avocat - Paris Métropole"
echo "================================================"
echo ""

# Vérifier Python 3
if ! command -v python3 &> /dev/null; then
    echo "❌ Python 3 non trouvé. Installe-le depuis https://python.org"
    exit 1
fi

echo "✅ Python $(python3 --version) détecté"

# Installer les dépendances
echo ""
echo "📦 Installation des bibliothèques..."
pip3 install requests beautifulsoup4 pandas lxml 2>&1 | grep -E "Successfully|already|ERROR"

echo ""
echo "🔍 Test de connexion à Pages Jaunes..."
python3 test_scraper.py

echo ""
echo "================================================"
echo "  Pour lancer le scraping complet (500 cabinets)"
echo "  Durée estimée : 30-60 minutes"
echo ""
echo "  COMMANDE : python3 scrape_cabinets_avocat.py"
echo "================================================"
echo ""
