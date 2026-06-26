"""
Scraper Avocats Paris - Sources multiples
Sources : Barreau de Paris + avocat.fr + jurisques.com
Lance : python scrape_barreau_paris.py
Résultat : avocats_paris.xlsx dans le même dossier
"""

import requests
from bs4 import BeautifulSoup
import pandas as pd
import time
import random
import re
import os
import json
from datetime import datetime

OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_XLSX = os.path.join(OUTPUT_DIR, "avocats_paris.xlsx")
TARGET = 500

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Accept-Language": "fr-FR,fr;q=0.9,en;q=0.8",
    "Accept-Encoding": "gzip, deflate, br",
    "Connection": "keep-alive",
    "Upgrade-Insecure-Requests": "1",
    "Sec-Fetch-Dest": "document",
    "Sec-Fetch-Mode": "navigate",
    "Sec-Fetch-Site": "none",
    "Sec-Fetch-User": "?1",
}

def creer_session(base_url):
    s = requests.Session()
    s.headers.update(HEADERS)
    try:
        s.get(base_url, timeout=15)
        time.sleep(random.uniform(2, 4))
    except Exception:
        pass
    return s

def extraire_email(texte):
    match = re.search(r"[\w.\-+]+@[\w.\-]+\.[a-zA-Z]{2,}", texte)
    return match.group(0).lower() if match else ""


# ═══════════════════════════════════════════════════════════════
#  SOURCE 1 : BARREAU DE PARIS (annuaire officiel)
#  URL : https://annuaire.avocatparis.org
# ═══════════════════════════════════════════════════════════════

def scraper_barreau_paris():
    """
    L'annuaire du Barreau de Paris charge les résultats via une API JSON.
    On interroge directement l'API pour récupérer les données proprement.
    """
    print("\n" + "═"*55)
    print("  SOURCE 1 : Annuaire du Barreau de Paris")
    print("═"*55)

    resultats = []
    session = creer_session("https://annuaire.avocatparis.org")

    # L'annuaire utilise une API REST
    api_url = "https://annuaire.avocatparis.org/api/search"
    page = 1
    par_page = 20

    while len(resultats) < TARGET:
        params = {
            "q": "",
            "localisation": "Paris",
            "page": page,
            "per_page": par_page,
        }
        try:
            resp = session.get(api_url, params=params, timeout=20)
            print(f"  Page {page} — Status {resp.status_code}")

            if resp.status_code == 200:
                try:
                    data = resp.json()
                    # Adapter selon la structure JSON retournée
                    avocats = data.get("results", data.get("data", data.get("avocats", [])))
                    if not avocats:
                        break
                    for a in avocats:
                        resultats.append({
                            "nom": a.get("nom", "") + " " + a.get("prenom", ""),
                            "cabinet": a.get("cabinet", a.get("denomination", "")),
                            "adresse": a.get("adresse", a.get("address", "")),
                            "telephone": a.get("telephone", a.get("phone", "")),
                            "email": a.get("email", a.get("courriel", "")),
                            "site_web": a.get("site_web", a.get("website", "")),
                            "specialite": a.get("specialite", a.get("domaines", "")),
                            "source": "Barreau de Paris",
                        })
                    print(f"    → {len(avocats)} avocats récupérés | Total : {len(resultats)}")
                    if len(avocats) < par_page:
                        break
                    page += 1
                    time.sleep(random.uniform(2, 4))
                except json.JSONDecodeError:
                    # Pas une API JSON → scraping HTML
                    print("    → Pas d'API JSON, passage au scraping HTML...")
                    resultats = scraper_barreau_html(session)
                    break
            elif resp.status_code == 404:
                # L'URL d'API n'existe pas → scraping HTML
                print("    → API non trouvée, passage au scraping HTML...")
                resultats = scraper_barreau_html(session)
                break
            else:
                print(f"    → Bloqué ({resp.status_code}), pause...")
                time.sleep(15)
                page += 1
                if page > 5:
                    break
        except Exception as e:
            print(f"    Erreur : {e}")
            break

    return resultats

def scraper_barreau_html(session):
    """Scraping HTML de l'annuaire du Barreau de Paris."""
    resultats = []
    base = "https://annuaire.avocatparis.org"

    for page in range(1, 50):
        if len(resultats) >= TARGET:
            break
        url = f"{base}/avocats?page={page}"
        try:
            resp = session.get(url, timeout=20)
            if resp.status_code != 200:
                print(f"  Page {page} bloquée ({resp.status_code})")
                time.sleep(10)
                continue
            soup = BeautifulSoup(resp.text, "lxml")

            # Sélecteurs possibles pour l'annuaire du Barreau
            cards = (
                soup.select(".avocat-item, .lawyer-card, .result-item") or
                soup.select("[class*='avocat'], [class*='lawyer']") or
                soup.select("article, .card, li.result")
            )

            if not cards:
                print(f"  Page {page} : aucun résultat (fin ou structure changée)")
                break

            for card in cards:
                nom_tag = card.select_one("h2, h3, .nom, .name, a[href*='avocat']")
                nom = nom_tag.get_text(strip=True) if nom_tag else ""
                if not nom:
                    continue
                adr_tag = card.select_one("address, .adresse, .address")
                adresse = adr_tag.get_text(" ", strip=True) if adr_tag else ""
                email = extraire_email(card.get_text())
                tel_tag = card.select_one("a[href^='tel:'], .telephone, .phone")
                tel = tel_tag.get("href","").replace("tel:","") if tel_tag else ""
                resultats.append({
                    "nom": nom, "cabinet": "", "adresse": adresse,
                    "telephone": tel.strip(), "email": email,
                    "site_web": "", "specialite": "", "source": "Barreau de Paris",
                })

            print(f"  Page {page} : +{len(cards)} | Total : {len(resultats)}")
            time.sleep(random.uniform(3, 6))

        except Exception as e:
            print(f"  Erreur page {page} : {e}")
            break

    return resultats


# ═══════════════════════════════════════════════════════════════
#  SOURCE 2 : AVOCAT.FR
# ═══════════════════════════════════════════════════════════════

def scraper_avocat_fr():
    print("\n" + "═"*55)
    print("  SOURCE 2 : avocat.fr")
    print("═"*55)

    resultats = []
    session = creer_session("https://www.avocat.fr")
    departements = ["75001","75002","75003","75004","75005","75006","75007",
                    "75008","75009","75010","75011","75012","75013","75014",
                    "75015","75016","75017","75018","75019","75020"]

    for cp in departements:
        if len(resultats) >= TARGET:
            break
        for page in range(1, 6):
            url = f"https://www.avocat.fr/annuaire/avocat-{cp}.htm"
            if page > 1:
                url = f"https://www.avocat.fr/annuaire/avocat-{cp}-{page}.htm"
            try:
                resp = session.get(url, timeout=20)
                if resp.status_code != 200:
                    break
                soup = BeautifulSoup(resp.text, "lxml")

                cards = (
                    soup.select(".avocat-block, .avocat-item, .result-avocat") or
                    soup.select("[class*='avocat'][class*='block'], [class*='avocat'][class*='item']") or
                    soup.select(".fiche, .listing-item, article")
                )

                if not cards:
                    break

                for card in cards:
                    nom_tag = card.select_one("h2 a, h3 a, .nom-avocat, a.avocat-nom")
                    nom = nom_tag.get_text(strip=True) if nom_tag else ""
                    if not nom or "avocat" in nom.lower():
                        continue
                    cab_tag = card.select_one(".cabinet, .denomination")
                    cabinet = cab_tag.get_text(strip=True) if cab_tag else ""
                    adr_tag = card.select_one("address, .adresse")
                    adresse = adr_tag.get_text(" ", strip=True) if adr_tag else ""
                    email_tag = card.select_one("a[href^='mailto:']")
                    email = email_tag["href"].replace("mailto:","").strip() if email_tag else extraire_email(card.get_text())
                    tel_tag = card.select_one("a[href^='tel:'], .tel")
                    tel = tel_tag.get("href","").replace("tel:","") if tel_tag else ""
                    spec_tag = card.select_one(".specialite, .domaine, .expertise")
                    spec = spec_tag.get_text(strip=True) if spec_tag else ""
                    resultats.append({
                        "nom": nom, "cabinet": cabinet, "adresse": adresse or cp,
                        "telephone": tel.strip(), "email": email,
                        "site_web": "", "specialite": spec, "source": "avocat.fr",
                    })

                print(f"  {cp} page {page} : +{len(cards)} | Total : {len(resultats)}")
                time.sleep(random.uniform(2, 5))

            except Exception as e:
                print(f"  Erreur {cp} page {page} : {e}")
                break

        time.sleep(random.uniform(2, 4))

    return resultats


# ═══════════════════════════════════════════════════════════════
#  SOURCE 3 : JURISQUES.COM
# ═══════════════════════════════════════════════════════════════

def scraper_jurisques():
    print("\n" + "═"*55)
    print("  SOURCE 3 : jurisques.com")
    print("═"*55)

    resultats = []
    session = creer_session("https://www.jurisques.com")

    for page in range(1, 60):
        if len(resultats) >= TARGET:
            break
        url = f"https://www.jurisques.com/avocats-paris-75/" if page == 1 else f"https://www.jurisques.com/avocats-paris-75/page/{page}/"
        try:
            resp = session.get(url, timeout=20)
            if resp.status_code != 200:
                print(f"  Page {page} : {resp.status_code}")
                if page == 1:
                    break
                time.sleep(8)
                continue
            soup = BeautifulSoup(resp.text, "lxml")

            cards = (
                soup.select(".fiche-avocat, .avocat, .lawyer-item") or
                soup.select("article, .entry, .post") or
                soup.select("[class*='avocat'], [class*='lawyer']")
            )

            if not cards:
                break

            for card in cards:
                nom_tag = card.select_one("h2, h3, .title, a[href*='avocat']")
                nom = nom_tag.get_text(strip=True) if nom_tag else ""
                if not nom:
                    continue
                email_tag = card.select_one("a[href^='mailto:']")
                email = email_tag["href"].replace("mailto:","").strip() if email_tag else extraire_email(card.get_text())
                tel_tag = card.select_one("a[href^='tel:']")
                tel = tel_tag.get("href","").replace("tel:","") if tel_tag else ""
                adr_tag = card.select_one("address, .adresse, .address, p")
                adresse = adr_tag.get_text(" ", strip=True) if adr_tag else ""
                resultats.append({
                    "nom": nom, "cabinet": "", "adresse": adresse,
                    "telephone": tel.strip(), "email": email,
                    "site_web": "", "specialite": "", "source": "jurisques.com",
                })

            print(f"  Page {page} : +{len(cards)} | Total : {len(resultats)}")
            time.sleep(random.uniform(3, 6))

        except Exception as e:
            print(f"  Erreur page {page} : {e}")
            break

    return resultats


# ═══════════════════════════════════════════════════════════════
#  EXPORT EXCEL
# ═══════════════════════════════════════════════════════════════

def exporter_excel(donnees):
    try:
        import openpyxl
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        from openpyxl.utils import get_column_letter

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Avocats Paris"

        BLEU = "1F3864"
        thin = Side(border_style="thin", color="CCCCCC")
        bordure = Border(left=thin, right=thin, top=thin, bottom=thin)

        ws.merge_cells("A1:J1")
        c = ws["A1"]
        c.value = f"AVOCATS PARIS MÉTROPOLE — {len(donnees)} contacts | Généré le {datetime.now().strftime('%d/%m/%Y')}"
        c.font = Font(bold=True, size=13, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=BLEU)
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 26

        entetes = ["N°", "Nom", "Cabinet", "Adresse", "Téléphone", "Email", "Site web", "Spécialité", "Source", "Statut"]
        for i, h in enumerate(entetes, 1):
            c = ws.cell(row=2, column=i, value=h)
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor=BLEU)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = bordure
        ws.row_dimensions[2].height = 18
        ws.freeze_panes = "A3"

        for idx, d in enumerate(donnees, 1):
            row = idx + 2
            vals = [idx, d.get("nom",""), d.get("cabinet",""), d.get("adresse",""),
                    d.get("telephone",""), d.get("email",""), d.get("site_web",""),
                    d.get("specialite",""), d.get("source",""), "À contacter"]
            bg = "D6E4F0" if idx % 2 == 0 else "FFFFFF"
            for col, val in enumerate(vals, 1):
                c = ws.cell(row=row, column=col, value=val)
                c.font = Font(size=10)
                c.border = bordure
                c.alignment = Alignment(horizontal="center" if col in (1,5,9,10) else "left", vertical="center")
                if col == 6 and val:
                    c.fill = PatternFill("solid", fgColor="E8F5E9")  # vert = email présent
                elif col == 6:
                    c.fill = PatternFill("solid", fgColor="FFF3E0")  # orange = email manquant
                else:
                    c.fill = PatternFill("solid", fgColor=bg)
            ws.row_dimensions[row].height = 16

        for i, w in enumerate([4, 30, 28, 35, 16, 34, 28, 28, 14, 14], 1):
            ws.column_dimensions[get_column_letter(i)].width = w

        ws.auto_filter.ref = f"A2:J{len(donnees)+2}"
        wb.save(OUTPUT_XLSX)
        return True
    except Exception as e:
        print(f"Erreur Excel : {e}")
        pd.DataFrame(donnees).to_csv(OUTPUT_XLSX.replace(".xlsx",".csv"), index=False, encoding="utf-8-sig")
        return False


# ═══════════════════════════════════════════════════════════════
#  PROGRAMME PRINCIPAL
# ═══════════════════════════════════════════════════════════════

print("\n" + "═"*55)
print("  SCRAPER AVOCATS PARIS — SOURCES MULTIPLES")
print(f"  Démarrage : {datetime.now().strftime('%H:%M:%S')}")
print(f"  Objectif  : {TARGET} avocats")
print("═"*55)

tous = []
noms_vus = set()

def ajouter_sans_doublon(nouveaux):
    ajouts = 0
    for d in nouveaux:
        cle = d["nom"].lower().strip()
        if cle and cle not in noms_vus:
            noms_vus.add(cle)
            tous.append(d)
            ajouts += 1
    return ajouts

# Source 1 : Barreau de Paris (priorité)
r1 = scraper_barreau_paris()
n = ajouter_sans_doublon(r1)
print(f"\n✅ Barreau de Paris : {n} avocats ajoutés | Total : {len(tous)}")

# Source 2 : avocat.fr (si besoin de plus)
if len(tous) < TARGET:
    r2 = scraper_avocat_fr()
    n = ajouter_sans_doublon(r2)
    print(f"\n✅ avocat.fr : {n} avocats ajoutés | Total : {len(tous)}")

# Source 3 : jurisques.com (si besoin de plus)
if len(tous) < TARGET:
    r3 = scraper_jurisques()
    n = ajouter_sans_doublon(r3)
    print(f"\n✅ jurisques.com : {n} avocats ajoutés | Total : {len(tous)}")

# Export
print(f"\n{'═'*55}")
print(f"  TOTAL : {len(tous)} avocats collectés")
emails = sum(1 for d in tous if d.get("email"))
print(f"  EMAILS : {emails} ({emails*100//max(len(tous),1)}%)")
print("═"*55)

ok = exporter_excel(tous)

print(f"\n{'═'*55}")
if ok:
    print(f"  FICHIER CRÉÉ : avocats_paris.xlsx")
else:
    print(f"  FICHIER CRÉÉ : avocats_paris.csv")
print(f"  Emplacement : {OUTPUT_DIR}")
print("═"*55)
