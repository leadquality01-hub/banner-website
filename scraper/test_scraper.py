"""
Test rapide du scraper - vérifie que Pages Jaunes répond correctement
avant de lancer le scraping complet (qui prend 30-60 min).
"""

import requests
from bs4 import BeautifulSoup
from scrape_cabinets_avocat import get_page, parse_listings, get_total_pages

def test_connexion():
    print("🔍 Test 1 : Connexion à Pages Jaunes...")
    soup = get_page("avocat", "Paris 8eme (75008)", page=1)
    if soup:
        print("   ✅ Connexion OK")
        return soup
    else:
        print("   ❌ Échec de connexion")
        return None

def test_parsing(soup):
    print("\n🔍 Test 2 : Extraction des cabinets...")
    listings = parse_listings(soup)
    print(f"   → {len(listings)} cabinets trouvés sur la page")
    if listings:
        print("   ✅ Parsing OK")
        print(f"\n   Exemple (1er résultat) :")
        for k, v in listings[0].items():
            if v:
                print(f"     {k:15} : {v[:80]}")
    else:
        print("   ⚠️  Aucun résultat - les sélecteurs CSS ont peut-être changé")
        print("   → Affichage du HTML pour diagnostic :")
        print(soup.prettify()[:2000])
    return listings

def test_pagination(soup):
    print("\n🔍 Test 3 : Détection de la pagination...")
    pages = get_total_pages(soup)
    print(f"   → {pages} page(s) détectée(s)")
    print("   ✅ OK" if pages > 0 else "   ⚠️  Pagination non détectée (valeur par défaut utilisée)")

if __name__ == "__main__":
    print("=" * 50)
    print("  TEST DU SCRAPER - Cabinets Avocat Paris")
    print("=" * 50)
    soup = test_connexion()
    if soup:
        listings = test_parsing(soup)
        test_pagination(soup)
        print(f"\n{'='*50}")
        if listings:
            print("✅ Tout fonctionne - lance : python scrape_cabinets_avocat.py")
        else:
            print("⚠️  Scraper à ajuster - voir diagnostic ci-dessus")
        print("=" * 50)
