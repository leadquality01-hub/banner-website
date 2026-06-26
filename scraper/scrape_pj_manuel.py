"""
Scraper Pages Jaunes - Avocats Paris
Stratégie : tu résous le CAPTCHA Cloudflare une fois manuellement,
ensuite le script scrape automatiquement.

INSTALLATION :
  pip install playwright openpyxl
  playwright install chromium

LANCEMENT :
  python scrape_pj_manuel.py
"""

import asyncio
import re
import os
import csv
import random
from datetime import datetime

try:
    from playwright.async_api import async_playwright
except ImportError:
    print("Lance d'abord :")
    print("  pip install playwright openpyxl")
    print("  playwright install chromium")
    exit(1)

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    EXCEL_OK = True
except ImportError:
    EXCEL_OK = False

OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_XLSX = os.path.join(OUTPUT_DIR, "avocats_pagesjaunes.xlsx")
OUTPUT_CSV  = os.path.join(OUTPUT_DIR, "avocats_pagesjaunes.csv")
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
    "Vincennes (94300)", "Versailles (78000)",
]

EMAIL_RE = re.compile(r"[\w.\-+]{2,}@[\w.\-]+\.[a-zA-Z]{2,6}")

async def pause(a=2, b=5):
    await asyncio.sleep(random.uniform(a, b))

def extraire_email(texte):
    m = EMAIL_RE.search(texte or "")
    return m.group(0).lower() if m else ""

async def est_cloudflare(page):
    contenu = await page.content()
    return "cloudflare" in contenu.lower() or "challenge" in contenu.lower() or "Sécurité" in contenu

async def attendre_resolution_captcha(page):
    """Attend que l'utilisateur résolve le CAPTCHA manuellement."""
    print("\n" + "="*55)
    print("  ⚠️  CAPTCHA CLOUDFLARE DÉTECTÉ")
    print("="*55)
    print("  👉 Dans la fenêtre Chrome qui est ouverte :")
    print("     1. Clique sur la case du CAPTCHA")
    print("     2. Suis les instructions")
    print("     3. Attends que la page des résultats s'affiche")
    print("  Le script reprend automatiquement ensuite.")
    print("="*55)

    # Attendre max 3 minutes que le CAPTCHA soit résolu
    for _ in range(180):
        await asyncio.sleep(1)
        if not await est_cloudflare(page):
            print("  ✅ CAPTCHA résolu ! Scraping en cours...")
            return True
    print("  ❌ Timeout - CAPTCHA non résolu en 3 minutes.")
    return False

async def scraper_zone(page, zone, tous, noms_vus):
    nouveaux_total = 0

    for num_page in range(1, 15):
        if len(tous) >= TARGET:
            break

        url = f"https://www.pagesjaunes.fr/annuaire/chercher?quoiqui=avocat&ou={zone.replace(' ', '+').replace('(','%28').replace(')','%29')}&page={num_page}"

        await page.goto(url, wait_until="domcontentloaded", timeout=30000)
        await pause(2, 4)

        # Vérifier Cloudflare
        if await est_cloudflare(page):
            resolu = await attendre_resolution_captcha(page)
            if not resolu:
                return nouveaux_total
            await pause(2, 3)

        contenu = await page.content()
        if "aucun résultat" in contenu.lower() or "aucun professionnel" in contenu.lower():
            break

        # Extraire les fiches avec les vrais sélecteurs Pages Jaunes 2024/2025
        cards = await page.query_selector_all("article.bi-item, li.bi-item, [class*='bi-item']")

        if not cards:
            # Sélecteurs alternatifs
            cards = await page.query_selector_all(".bi-pro, [data-bi-id]")

        if not cards:
            # Dump HTML pour debug
            html = await page.content()
            # Chercher n'importe quelle liste de résultats
            cards = await page.query_selector_all("#listResults > li, .results > li, ul.bi-list > li")

        if not cards:
            print(f"    ⚠️  Page {num_page} : aucune fiche détectée (fin de zone ou structure inconnue)")
            break

        nouveaux_page = 0
        for card in cards:
            try:
                # Nom - essayer plusieurs sélecteurs
                nom = ""
                for sel in ["a.bi-denomination", "h3.bi-denomination a", ".denomination-links a",
                            "[class*='denomination'] a", "h2 a", "h3 a"]:
                    el = await card.query_selector(sel)
                    if el:
                        nom = (await el.inner_text()).strip()
                        break

                if not nom or nom.lower() in noms_vus:
                    continue

                # Adresse
                adresse = ""
                for sel in ["address", "[class*='address']", "[class*='adresse']"]:
                    el = await card.query_selector(sel)
                    if el:
                        adresse = (await el.inner_text()).replace("\n", " ").strip()
                        break

                # Téléphone
                telephone = ""
                for sel in ["a[href^='tel:']", "[class*='phone']", "[class*='tel']"]:
                    el = await card.query_selector(sel)
                    if el:
                        href = await el.get_attribute("href") or ""
                        telephone = href.replace("tel:", "").strip() or (await el.inner_text()).strip()
                        break

                # Email direct sur la fiche listing
                texte_card = await card.inner_text()
                email = extraire_email(texte_card)

                # Email via lien mailto
                if not email:
                    mailto = await card.query_selector("a[href^='mailto:']")
                    if mailto:
                        href = await mailto.get_attribute("href") or ""
                        email = href.replace("mailto:", "").strip().lower()

                # URL fiche détail
                fiche_url = ""
                for sel in ["a.bi-denomination", "h3 a", "[class*='denomination'] a"]:
                    el = await card.query_selector(sel)
                    if el:
                        href = await el.get_attribute("href") or ""
                        if href.startswith("/"):
                            fiche_url = "https://www.pagesjaunes.fr" + href
                        break

                noms_vus.add(nom.lower())
                tous.append({
                    "nom": nom,
                    "adresse": adresse,
                    "telephone": telephone,
                    "email": email,
                    "fiche_url": fiche_url,
                    "zone": zone,
                })
                nouveaux_page += 1

            except Exception:
                continue

        nouveaux_total += nouveaux_page
        print(f"    Page {num_page} : +{nouveaux_page} fiches | Total : {len(tous)}/{TARGET}")

        if nouveaux_page == 0:
            break

        await pause(3, 6)

    return nouveaux_total

async def enrichir_emails(page, tous):
    """Visite les fiches détail pour récupérer les emails manquants."""
    sans_email = [c for c in tous if not c.get("email") and c.get("fiche_url")]
    if not sans_email:
        return

    print(f"\n📧 Recherche emails sur {len(sans_email)} fiches individuelles...")
    trouves = 0

    for i, cab in enumerate(sans_email):
        if i % 30 == 0 and i > 0:
            print(f"    {i}/{len(sans_email)} fiches visitées — {trouves} emails trouvés")
        try:
            await page.goto(cab["fiche_url"], wait_until="domcontentloaded", timeout=20000)
            await pause(1, 2)

            if await est_cloudflare(page):
                await attendre_resolution_captcha(page)

            # Email via mailto
            el = await page.query_selector("a[href^='mailto:']")
            if el:
                href = await el.get_attribute("href") or ""
                cab["email"] = href.replace("mailto:", "").split("?")[0].strip().lower()
                trouves += 1
                continue

            # Email dans le texte
            texte = await page.inner_text("body")
            email = extraire_email(texte)
            if email:
                cab["email"] = email
                trouves += 1

            # Site web → aller chercher l'email dessus
            if not cab.get("email"):
                site_el = await page.query_selector("a[class*='site'], a[class*='web'][href*='http']")
                if site_el:
                    site_url = await site_el.get_attribute("href") or ""
                    if site_url and "pagesjaunes" not in site_url:
                        try:
                            await page.goto(site_url, wait_until="domcontentloaded", timeout=12000)
                            await pause(1, 2)
                            texte = await page.inner_text("body")
                            email = extraire_email(texte)
                            if email:
                                cab["email"] = email
                                cab["site_web"] = site_url
                                trouves += 1
                        except Exception:
                            pass

        except Exception:
            pass

        await pause(1, 2)

    print(f"    ✅ {trouves} emails récupérés sur les fiches")

def exporter(donnees):
    import csv
    if not donnees:
        print("⚠️  Aucune donnée.")
        return

    with open(OUTPUT_CSV, "w", newline="", encoding="utf-8-sig") as f:
        champs = ["nom", "adresse", "telephone", "email", "zone", "fiche_url"]
        w = csv.DictWriter(f, fieldnames=champs, extrasaction="ignore")
        w.writeheader()
        w.writerows(donnees)
    print(f"  CSV  → {OUTPUT_CSV}")

    if not EXCEL_OK:
        return

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Avocats Pages Jaunes"
    BLEU = "1F3864"
    thin = Side(border_style="thin", color="CCCCCC")
    brd = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.merge_cells("A1:G1")
    c = ws["A1"]
    c.value = f"AVOCATS PARIS — Pages Jaunes — {len(donnees)} contacts | {datetime.now().strftime('%d/%m/%Y')}"
    c.font = Font(bold=True, size=13, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=BLEU)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    entetes = ["N°", "Nom / Cabinet", "Adresse", "Téléphone", "Email", "Zone", "Statut"]
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
                d.get("email",""), d.get("zone",""), "À contacter"]
        bg = "D6E4F0" if idx % 2 == 0 else "FFFFFF"
        for col, val in enumerate(vals, 1):
            c = ws.cell(row=r, column=col, value=val)
            c.font = Font(size=10)
            c.border = brd
            c.alignment = Alignment(horizontal="center" if col in (1,4,7) else "left", vertical="center")
            if col == 5 and val:
                c.fill = PatternFill("solid", fgColor="E8F5E9")
                c.font = Font(size=10, bold=True, color="1B5E20")
            elif col == 5:
                c.fill = PatternFill("solid", fgColor="FFF3E0")
            else:
                c.fill = PatternFill("solid", fgColor=bg)
        ws.row_dimensions[r].height = 16

    for i, w in enumerate([4, 38, 38, 16, 34, 26, 14], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.auto_filter.ref = f"A2:G{len(donnees)+2}"
    wb.save(OUTPUT_XLSX)
    print(f"  XLSX → {OUTPUT_XLSX}")

async def main():
    print("\n" + "═"*55)
    print("  SCRAPER PAGES JAUNES — AVOCATS PARIS")
    print(f"  Démarrage : {datetime.now().strftime('%H:%M:%S')}")
    print(f"  Objectif  : {TARGET} avocats")
    print()
    print("  ⚠️  IMPORTANT : quand le CAPTCHA apparaît,")
    print("  résous-le manuellement dans Chrome.")
    print("  Le script continue tout seul après.")
    print("═"*55)

    tous = []
    noms_vus = set()

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=False, slow_mo=100)
        context = await browser.new_context(
            viewport={"width": 1280, "height": 900},
            locale="fr-FR",
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
        )
        page = await context.new_page()

        # Ouvrir Pages Jaunes et résoudre le CAPTCHA initial
        print("\n  Ouverture de Pages Jaunes...")
        await page.goto("https://www.pagesjaunes.fr/annuaire/chercher?quoiqui=avocat&ou=Paris", wait_until="domcontentloaded", timeout=30000)
        await pause(2, 3)

        if await est_cloudflare(page):
            resolu = await attendre_resolution_captcha(page)
            if not resolu:
                print("❌ Impossible de continuer sans résoudre le CAPTCHA.")
                await browser.close()
                return

        # Scraping zone par zone
        for zone in ZONES:
            if len(tous) >= TARGET:
                break
            print(f"\n📍 {zone}")
            await scraper_zone(page, zone, tous, noms_vus)
            await pause(4, 8)

        # Enrichissement emails via fiches détail
        await enrichir_emails(page, tous)

        await browser.close()

    emails = sum(1 for c in tous if c.get("email"))
    print(f"\n{'═'*55}")
    print(f"  TOTAL    : {len(tous)} avocats")
    print(f"  EMAILS   : {emails} ({emails*100//max(len(tous),1)}%)")
    print("═"*55)

    exporter(tous)
    print(f"\n✅ Terminé ! Fichier dans : {OUTPUT_DIR}")

asyncio.run(main())
