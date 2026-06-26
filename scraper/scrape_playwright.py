"""
Scraper Avocats Paris avec Playwright (vrai navigateur Chrome)
Contourne tous les anti-bots car c'est un vrai navigateur.

INSTALLATION (1 seule fois) :
  pip install playwright
  playwright install chromium

LANCEMENT :
  python scrape_playwright.py
"""

import asyncio
import re
import os
import json
import random
import time
from datetime import datetime

try:
    from playwright.async_api import async_playwright
except ImportError:
    print("❌ Playwright non installé.")
    print("   Lance ces 2 commandes puis relance le script :")
    print("   pip install playwright")
    print("   playwright install chromium")
    exit(1)

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    EXCEL_OK = True
except ImportError:
    EXCEL_OK = False

OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_XLSX = os.path.join(OUTPUT_DIR, "avocats_paris.xlsx")
OUTPUT_CSV  = os.path.join(OUTPUT_DIR, "avocats_paris.csv")
TARGET = 500

ZONES = [
    "Paris 1er (75001)", "Paris 2eme (75002)", "Paris 3eme (75003)",
    "Paris 4eme (75004)", "Paris 5eme (75005)", "Paris 6eme (75006)",
    "Paris 7eme (75007)", "Paris 8eme (75008)", "Paris 9eme (75009)",
    "Paris 10eme (75010)", "Paris 11eme (75011)", "Paris 12eme (75012)",
    "Paris 13eme (75013)", "Paris 14eme (75014)", "Paris 15eme (75015)",
    "Paris 16eme (75016)", "Paris 17eme (75017)", "Paris 18eme (75018)",
    "Paris 19eme (75019)", "Paris 20eme (75020)",
    "Boulogne-Billancourt (92100)", "Neuilly-sur-Seine (92200)",
    "Levallois-Perret (92300)", "Nanterre (92000)", "Courbevoie (92400)",
    "Vincennes (94300)", "Saint-Denis (93200)", "Versailles (78000)",
]

def extraire_email(texte):
    m = re.search(r"[\w.\-+]+@[\w.\-]+\.[a-zA-Z]{2,}", texte or "")
    return m.group(0).lower() if m else ""

async def pause(mini=2, maxi=5):
    await asyncio.sleep(random.uniform(mini, maxi))

async def scraper_zone(page, zone, tous, noms_vus):
    """Scrape une zone géographique sur Pages Jaunes."""
    nouveaux_total = 0

    for num_page in range(1, 12):
        if len(tous) >= TARGET:
            break

        url = f"https://www.pagesjaunes.fr/annuaire/chercher?quoiqui=avocat&ou={zone.replace(' ', '+')}&page={num_page}"
        try:
            await page.goto(url, wait_until="domcontentloaded", timeout=30000)
            await pause(2, 4)

            # Fermer cookie popup si présent
            try:
                btn = page.locator("#didomi-notice-agree-button, .didomi-continue-without-agreeing, button[aria-label*='accepter']")
                if await btn.count() > 0:
                    await btn.first.click()
                    await pause(1, 2)
            except Exception:
                pass

            # Vérifier qu'on a des résultats
            contenu = await page.content()
            if "Aucun résultat" in contenu or "aucun professionnel" in contenu.lower():
                break

            # Extraire les fiches
            cards = await page.query_selector_all("li.bi-item, li.bi-pro, article.bi-item, [class*='bi-item']")

            if not cards:
                # Essai sélecteur alternatif
                cards = await page.query_selector_all(".results-list li, #listResults li")

            if not cards:
                print(f"    ⚠️  Pas de résultats page {num_page}")
                break

            nouveaux_page = 0
            for card in cards:
                try:
                    # Nom
                    nom_el = await card.query_selector("a.bi-denomination, h3 a, .denomination-links a")
                    nom = (await nom_el.inner_text()).strip() if nom_el else ""
                    if not nom or nom.lower() in noms_vus:
                        continue

                    # Adresse
                    adr_el = await card.query_selector("address, [class*='address']")
                    adresse = (await adr_el.inner_text()).strip() if adr_el else ""

                    # Téléphone
                    tel_el = await card.query_selector("a[href^='tel:']")
                    telephone = ""
                    if tel_el:
                        telephone = (await tel_el.get_attribute("href") or "").replace("tel:", "").strip()

                    # Email (rarement visible en listing)
                    texte_card = await card.inner_text()
                    email = extraire_email(texte_card)

                    # Lien fiche détail
                    lien_el = await card.query_selector("a.bi-denomination, h3 a")
                    href = await lien_el.get_attribute("href") if lien_el else ""
                    fiche_url = "https://www.pagesjaunes.fr" + href if href and href.startswith("/") else ""

                    noms_vus.add(nom.lower())
                    tous.append({
                        "nom": nom,
                        "adresse": adresse.replace("\n", " "),
                        "telephone": telephone,
                        "email": email,
                        "site_web": "",
                        "specialite": "",
                        "zone": zone,
                        "fiche_url": fiche_url,
                    })
                    nouveaux_page += 1

                except Exception:
                    continue

            nouveaux_total += nouveaux_page
            print(f"    Page {num_page} : +{nouveaux_page} | Total : {len(tous)}/{TARGET}")

            if nouveaux_page == 0:
                break

            await pause(3, 6)

        except Exception as e:
            print(f"    Erreur page {num_page} : {e}")
            await pause(5, 10)
            break

    return nouveaux_total

async def enrichir_email(page, cab):
    """Visite la fiche pour trouver email et site web."""
    if not cab.get("fiche_url") or cab.get("email"):
        return
    try:
        await page.goto(cab["fiche_url"], wait_until="domcontentloaded", timeout=20000)
        await pause(1, 2)

        # Email
        email_el = await page.query_selector("a[href^='mailto:']")
        if email_el:
            href = await email_el.get_attribute("href")
            cab["email"] = href.replace("mailto:", "").strip().lower()

        if not cab["email"]:
            texte = await page.inner_text("body")
            cab["email"] = extraire_email(texte)

        # Site web
        site_el = await page.query_selector("a[class*='site-web'], a[class*='website']")
        if site_el:
            cab["site_web"] = await site_el.get_attribute("href") or ""

    except Exception:
        pass

def exporter(donnees):
    if not donnees:
        print("⚠️  Aucune donnée à exporter.")
        return

    # CSV (toujours)
    import csv
    with open(OUTPUT_CSV, "w", newline="", encoding="utf-8-sig") as f:
        champs = ["nom", "adresse", "telephone", "email", "site_web", "specialite", "zone"]
        w = csv.DictWriter(f, fieldnames=champs, extrasaction="ignore")
        w.writeheader()
        w.writerows(donnees)
    print(f"  CSV  → {OUTPUT_CSV}")

    # Excel
    if not EXCEL_OK:
        print("  (openpyxl non installé — seulement CSV créé)")
        return

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Avocats Paris"
    BLEU = "1F3864"
    thin = Side(border_style="thin", color="CCCCCC")
    brd = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.merge_cells("A1:H1")
    c = ws["A1"]
    c.value = f"AVOCATS PARIS — {len(donnees)} contacts | {datetime.now().strftime('%d/%m/%Y')}"
    c.font = Font(bold=True, size=13, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=BLEU)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    entetes = ["N°", "Nom", "Adresse", "Téléphone", "Email", "Site web", "Zone", "Statut"]
    for i, h in enumerate(entetes, 1):
        c = ws.cell(row=2, column=i, value=h)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=BLEU)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = brd
    ws.row_dimensions[2].height = 18
    ws.freeze_panes = "A3"

    for idx, d in enumerate(donnees, 1):
        r = idx + 2
        vals = [idx, d.get("nom",""), d.get("adresse",""), d.get("telephone",""),
                d.get("email",""), d.get("site_web",""), d.get("zone",""), "À contacter"]
        bg = "D6E4F0" if idx % 2 == 0 else "FFFFFF"
        for col, val in enumerate(vals, 1):
            c = ws.cell(row=r, column=col, value=val)
            c.font = Font(size=10)
            c.border = brd
            c.alignment = Alignment(horizontal="center" if col in (1,4,8) else "left", vertical="center")
            if col == 5 and val:
                c.fill = PatternFill("solid", fgColor="E8F5E9")
            elif col == 5:
                c.fill = PatternFill("solid", fgColor="FFF3E0")
            else:
                c.fill = PatternFill("solid", fgColor=bg)
        ws.row_dimensions[r].height = 16

    for i, w in enumerate([4, 35, 38, 16, 34, 30, 26, 14], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.auto_filter.ref = f"A2:H{len(donnees)+2}"
    wb.save(OUTPUT_XLSX)
    print(f"  XLSX → {OUTPUT_XLSX}")

async def main():
    print("\n" + "═"*55)
    print("  SCRAPER AVOCATS PARIS (Playwright + Chrome)")
    print(f"  Démarrage : {datetime.now().strftime('%H:%M:%S')}")
    print(f"  Objectif  : {TARGET} avocats")
    print("═"*55)

    tous = []
    noms_vus = set()

    async with async_playwright() as pw:
        # Lancer Chrome en mode visible (headless=False) pour éviter détection
        browser = await pw.chromium.launch(headless=False, slow_mo=100)
        context = await browser.new_context(
            viewport={"width": 1280, "height": 800},
            locale="fr-FR",
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
        )
        page = await context.new_page()

        # Visite d'abord la page d'accueil pour avoir les cookies
        print("\n  Ouverture de Pages Jaunes...")
        await page.goto("https://www.pagesjaunes.fr", wait_until="domcontentloaded", timeout=30000)
        await pause(3, 5)

        # Scraping par zone
        for zone in ZONES:
            if len(tous) >= TARGET:
                break
            print(f"\n📍 {zone}")
            await scraper_zone(page, zone, tous, noms_vus)
            await pause(4, 8)

        # Enrichissement emails (sur les 200 premiers sans email)
        sans_email = [c for c in tous if not c.get("email") and c.get("fiche_url")][:200]
        if sans_email:
            print(f"\n📧 Recherche d'emails sur {len(sans_email)} fiches...")
            for i, cab in enumerate(sans_email):
                await enrichir_email(page, cab)
                if (i+1) % 20 == 0:
                    emails = sum(1 for c in tous if c.get("email"))
                    print(f"    {i+1}/{len(sans_email)} fiches — {emails} emails trouvés")
                await pause(1, 3)

        await browser.close()

    # Résumé
    emails = sum(1 for c in tous if c.get("email"))
    print(f"\n{'═'*55}")
    print(f"  RÉSULTAT FINAL")
    print(f"  Avocats collectés : {len(tous)}")
    print(f"  Avec email        : {emails} ({emails*100//max(len(tous),1)}%)")
    print("═"*55)

    exporter(tous)

    print(f"\n✅ TERMINÉ — fichiers créés dans :")
    print(f"   {OUTPUT_DIR}")

asyncio.run(main())
