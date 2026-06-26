"""
Script de diagnostic - capture le HTML réel de Pages Jaunes
Lance : python debug_pagesjaunes.py
Crée : debug_page.html et debug_screenshot.png
"""
import asyncio
import os
from playwright.async_api import async_playwright

OUTPUT_DIR = os.path.dirname(os.path.abspath(__file__))

async def main():
    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=False, slow_mo=200)
        context = await browser.new_context(
            viewport={"width": 1280, "height": 800},
            locale="fr-FR",
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
        )
        page = await context.new_page()

        print("Ouverture Pages Jaunes...")
        await page.goto("https://www.pagesjaunes.fr", timeout=30000)
        await asyncio.sleep(3)

        # Accepter cookies si popup
        try:
            for sel in ["#didomi-notice-agree-button", "button:has-text('Accepter')", "button:has-text('Tout accepter')"]:
                btn = page.locator(sel)
                if await btn.count() > 0:
                    await btn.first.click()
                    print(f"Cookies acceptés ({sel})")
                    await asyncio.sleep(2)
                    break
        except Exception:
            pass

        print("Recherche avocat Paris 8eme...")
        await page.goto("https://www.pagesjaunes.fr/annuaire/chercher?quoiqui=avocat&ou=Paris+8eme+%2875008%29", timeout=30000)
        await asyncio.sleep(4)

        # Screenshot
        screenshot_path = os.path.join(OUTPUT_DIR, "debug_screenshot.png")
        await page.screenshot(path=screenshot_path, full_page=False)
        print(f"Screenshot : {screenshot_path}")

        # HTML complet
        html = await page.content()
        html_path = os.path.join(OUTPUT_DIR, "debug_page.html")
        with open(html_path, "w", encoding="utf-8") as f:
            f.write(html)
        print(f"HTML sauvegardé : {html_path}")

        # Afficher les 50 premières classes CSS trouvées
        classes = await page.evaluate("""
            () => {
                const all = document.querySelectorAll('*');
                const classes = new Set();
                all.forEach(el => el.classList.forEach(c => classes.add(c)));
                return [...classes].slice(0, 100);
            }
        """)
        print("\nClasses CSS trouvées sur la page :")
        print([c for c in classes if any(k in c.lower() for k in ['result','bi-','listing','item','card','fiche','avocat'])])

        # Chercher les éléments contenant des noms
        h_tags = await page.evaluate("""
            () => {
                const tags = [...document.querySelectorAll('h2, h3, h4')];
                return tags.slice(0, 20).map(t => ({tag: t.tagName, class: t.className, text: t.innerText.trim().substring(0, 80)}));
            }
        """)
        print("\nBalises h2/h3/h4 trouvées :")
        for h in h_tags:
            print(f"  <{h['tag']} class='{h['class']}'> : {h['text']}")

        await asyncio.sleep(3)
        await browser.close()

    print("\n✅ Diagnostic terminé.")
    print(f"   Envoie le fichier debug_screenshot.png pour que je vois la vraie structure.")

asyncio.run(main())
