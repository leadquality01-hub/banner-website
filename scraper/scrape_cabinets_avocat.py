"""
Scraper de cabinets d'avocat à Paris - Pages Jaunes
Objectif : constituer une base de 500 cabinets pour prospection par email
"""

import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
import random
import json
import re
import os
from datetime import datetime

# ─── Configuration ────────────────────────────────────────────────────────────

BASE_URL = "https://www.pagesjaunes.fr/annuaire/chercher"
OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_CSV = os.path.join(OUTPUT_DIR, "cabinets_avocat_paris.csv")
OUTPUT_JSON = os.path.join(OUTPUT_DIR, "cabinets_avocat_paris.json")
TARGET = 500

# Zones géographiques Paris métropole (codes postaux + communes)
ZONES = [
    ("avocat", "Paris 1er (75001)"),
    ("avocat", "Paris 2eme (75002)"),
    ("avocat", "Paris 3eme (75003)"),
    ("avocat", "Paris 4eme (75004)"),
    ("avocat", "Paris 5eme (75005)"),
    ("avocat", "Paris 6eme (75006)"),
    ("avocat", "Paris 7eme (75007)"),
    ("avocat", "Paris 8eme (75008)"),
    ("avocat", "Paris 9eme (75009)"),
    ("avocat", "Paris 10eme (75010)"),
    ("avocat", "Paris 11eme (75011)"),
    ("avocat", "Paris 12eme (75012)"),
    ("avocat", "Paris 13eme (75013)"),
    ("avocat", "Paris 14eme (75014)"),
    ("avocat", "Paris 15eme (75015)"),
    ("avocat", "Paris 16eme (75016)"),
    ("avocat", "Paris 17eme (75017)"),
    ("avocat", "Paris 18eme (75018)"),
    ("avocat", "Paris 19eme (75019)"),
    ("avocat", "Paris 20eme (75020)"),
    ("avocat", "Boulogne-Billancourt (92100)"),
    ("avocat", "Neuilly-sur-Seine (92200)"),
    ("avocat", "Levallois-Perret (92300)"),
    ("avocat", "Nanterre (92000)"),
    ("avocat", "Vincennes (94300)"),
    ("avocat", "Saint-Denis (93200)"),
    ("avocat", "Montreuil (93100)"),
    ("avocat", "Versailles (78000)"),
]

HEADERS_LIST = [
    {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
        "Accept-Language": "fr-FR,fr;q=0.9",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    },
    {
        "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15",
        "Accept-Language": "fr-FR,fr;q=0.9",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    },
]


# ─── Fonctions de scraping ────────────────────────────────────────────────────

def get_page(quoi: str, ou: str, page: int = 1) -> BeautifulSoup | None:
    """Récupère une page de résultats Pages Jaunes."""
    params = {
        "quoiqui": quoi,
        "ou": ou,
        "page": page,
    }
    headers = random.choice(HEADERS_LIST)
    try:
        resp = requests.get(BASE_URL, params=params, headers=headers, timeout=15)
        if resp.status_code == 200:
            return BeautifulSoup(resp.text, "lxml")
        print(f"  ⚠️  Status {resp.status_code} pour page {page} - {ou}")
        return None
    except requests.RequestException as e:
        print(f"  ❌ Erreur réseau : {e}")
        return None


def extract_email_from_text(text: str) -> str:
    """Extrait un email depuis un bloc de texte."""
    match = re.search(r"[\w.\-+]+@[\w.\-]+\.[a-zA-Z]{2,}", text)
    return match.group(0) if match else ""


def parse_listings(soup: BeautifulSoup) -> list[dict]:
    """Extrait les cabinets depuis une page de résultats."""
    results = []

    # Pages Jaunes utilise des articles avec class "bi-item" ou "bi-pro"
    cards = soup.select("ul#listResults li.bi-item, ul#listResults li.bi-pro, article.bi-item")

    if not cards:
        # Sélecteur de secours plus large
        cards = soup.select("[class*='bi-item'], [class*='bi-pro']")

    for card in cards:
        cabinet = {}

        # Nom du cabinet
        name_tag = card.select_one("a.bi-denomination, h3.bi-denomination a, .denomination-links a")
        cabinet["nom"] = name_tag.get_text(strip=True) if name_tag else ""

        # Adresse
        addr_tag = card.select_one("address, .bi-address, [class*='address']")
        cabinet["adresse"] = addr_tag.get_text(" ", strip=True) if addr_tag else ""

        # Téléphone
        tel_tag = card.select_one("[class*='phone'], [class*='tel'], a[href^='tel:']")
        if tel_tag:
            cabinet["telephone"] = tel_tag.get("href", "").replace("tel:", "").strip() or tel_tag.get_text(strip=True)
        else:
            cabinet["telephone"] = ""

        # Site web
        web_tag = card.select_one("a[href*='http'][class*='site'], a[class*='web']")
        cabinet["site_web"] = web_tag.get("href", "") if web_tag else ""

        # Email (rarement affiché directement)
        email_raw = card.get_text()
        cabinet["email"] = extract_email_from_text(email_raw)

        # Spécialité / description
        desc_tag = card.select_one(".bi-activite, [class*='activite'], [class*='activites']")
        cabinet["specialite"] = desc_tag.get_text(" ", strip=True) if desc_tag else ""

        # URL de la fiche détaillée
        link_tag = card.select_one("a.bi-denomination, h3 a, a[class*='denomination']")
        cabinet["url_fiche"] = "https://www.pagesjaunes.fr" + link_tag["href"] if link_tag and link_tag.get("href", "").startswith("/") else ""

        if cabinet["nom"]:
            results.append(cabinet)

    return results


def get_total_pages(soup: BeautifulSoup) -> int:
    """Récupère le nombre total de pages de résultats."""
    pagination = soup.select_one("[class*='pagination'] [class*='last'], .pagination li:last-child a")
    if pagination:
        try:
            return int(re.search(r"\d+", pagination.get_text()).group())
        except Exception:
            pass
    # Cherche aussi dans les métadonnées
    meta = soup.select_one("[data-nb-results], [data-total]")
    if meta:
        total = meta.get("data-nb-results") or meta.get("data-total")
        if total:
            return max(1, int(total) // 20)
    return 5  # Valeur par défaut conservative


# ─── Scraper de fiche détaillée (pour récupérer l'email) ─────────────────────

def scrape_detail(url: str) -> dict:
    """Visite la fiche d'un cabinet pour récupérer email et infos supplémentaires."""
    if not url:
        return {}
    headers = random.choice(HEADERS_LIST)
    try:
        resp = requests.get(url, headers=headers, timeout=15)
        if resp.status_code != 200:
            return {}
        soup = BeautifulSoup(resp.text, "lxml")
        detail = {}

        # Email
        email_tag = soup.select_one("a[href^='mailto:']")
        if email_tag:
            detail["email"] = email_tag["href"].replace("mailto:", "").strip()
        else:
            detail["email"] = extract_email_from_text(soup.get_text())

        # Site web (parfois masqué en listing)
        site_tag = soup.select_one("a[class*='site-web'], a[class*='website']")
        if site_tag:
            detail["site_web"] = site_tag.get("href", "")

        return detail
    except Exception:
        return {}


# ─── Boucle principale ────────────────────────────────────────────────────────

def run_scraper():
    all_cabinets = []
    seen_names = set()

    print(f"\n{'='*60}")
    print(f"  Scraper Cabinets d'Avocat - Paris Métropole")
    print(f"  Objectif : {TARGET} cabinets")
    print(f"  Démarrage : {datetime.now().strftime('%H:%M:%S')}")
    print(f"{'='*60}\n")

    for quoi, ou in ZONES:
        if len(all_cabinets) >= TARGET:
            break

        print(f"\n📍 Zone : {ou}")
        soup = get_page(quoi, ou, page=1)
        if not soup:
            time.sleep(3)
            continue

        total_pages = get_total_pages(soup)
        print(f"   → {total_pages} page(s) détectée(s)")

        for page_num in range(1, min(total_pages + 1, 15)):  # Max 15 pages par zone
            if len(all_cabinets) >= TARGET:
                break

            if page_num > 1:
                soup = get_page(quoi, ou, page=page_num)
                if not soup:
                    break
                time.sleep(random.uniform(2.0, 4.5))  # Pause respectueuse

            listings = parse_listings(soup)

            new_count = 0
            for cab in listings:
                name_key = cab["nom"].lower().strip()
                if name_key and name_key not in seen_names:
                    seen_names.add(name_key)
                    cab["zone"] = ou
                    all_cabinets.append(cab)
                    new_count += 1

            print(f"   Page {page_num} : +{new_count} cabinets | Total : {len(all_cabinets)}/{TARGET}")

            # Pause entre pages
            time.sleep(random.uniform(1.5, 3.0))

        # Pause entre zones
        time.sleep(random.uniform(3.0, 6.0))

    print(f"\n{'='*60}")
    print(f"  Scraping terminé : {len(all_cabinets)} cabinets collectés")
    print(f"{'='*60}\n")

    # ─── Phase 2 : enrichissement via fiches détaillées ──────────────────────
    print("📧 Enrichissement des emails via fiches détaillées...")
    enriched = 0
    for i, cab in enumerate(all_cabinets):
        if cab.get("email") or not cab.get("url_fiche"):
            continue
        detail = scrape_detail(cab["url_fiche"])
        if detail.get("email"):
            cab["email"] = detail["email"]
            enriched += 1
        if detail.get("site_web") and not cab.get("site_web"):
            cab["site_web"] = detail["site_web"]
        if i % 50 == 0 and i > 0:
            print(f"   → {i}/{len(all_cabinets)} fiches visitées, {enriched} emails trouvés")
        time.sleep(random.uniform(1.0, 2.5))

    print(f"   ✅ {enriched} emails enrichis\n")

    return all_cabinets


# ─── Export ───────────────────────────────────────────────────────────────────

def export_results(cabinets: list[dict]):
    if not cabinets:
        print("⚠️  Aucun résultat à exporter.")
        return

    df = pd.DataFrame(cabinets, columns=[
        "nom", "adresse", "telephone", "email", "site_web", "specialite", "zone", "url_fiche"
    ])

    # Nettoyage
    df = df.drop_duplicates(subset=["nom"]).reset_index(drop=True)
    df["nom"] = df["nom"].str.strip()
    df["email"] = df["email"].str.lower().str.strip()

    # Colonne statut pour le suivi des envois
    df["statut_contact"] = "à contacter"
    df["date_contact"] = ""
    df["notes"] = ""

    # Export CSV (encodage UTF-8 pour Excel)
    df.to_csv(OUTPUT_CSV, index=False, encoding="utf-8-sig")

    # Export JSON
    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(cabinets, f, ensure_ascii=False, indent=2)

    # Statistiques
    total = len(df)
    avec_email = df["email"].notna() & (df["email"] != "")
    avec_tel = df["telephone"].notna() & (df["telephone"] != "")
    avec_site = df["site_web"].notna() & (df["site_web"] != "")

    print(f"{'='*60}")
    print(f"  RÉSULTATS EXPORTÉS")
    print(f"{'='*60}")
    print(f"  📊 Total cabinets        : {total}")
    print(f"  📧 Avec email            : {avec_email.sum()} ({avec_email.mean()*100:.0f}%)")
    print(f"  📞 Avec téléphone        : {avec_tel.sum()} ({avec_tel.mean()*100:.0f}%)")
    print(f"  🌐 Avec site web         : {avec_site.sum()} ({avec_site.mean()*100:.0f}%)")
    print(f"  💾 CSV  → {OUTPUT_CSV}")
    print(f"  💾 JSON → {OUTPUT_JSON}")
    print(f"{'='*60}\n")

    # Aperçu
    print("Aperçu des 5 premiers résultats :")
    print(df[["nom", "adresse", "telephone", "email", "zone"]].head().to_string(index=False))


# ─── Point d'entrée ───────────────────────────────────────────────────────────

if __name__ == "__main__":
    cabinets = run_scraper()
    export_results(cabinets)
