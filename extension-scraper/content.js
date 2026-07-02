// Écoute les messages du popup
chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (msg.action === "scrape") {
    const contacts = scrapeContacts();
    sendResponse({ contacts });
  }
  return true;
});

function scrapeContacts() {
  const results = [];
  const emailRe = /[\w.\-+]{2,}@[\w.\-]+\.[a-zA-Z]{2,6}/g;
  const telRe = /(?:(?:\+33|0033|0)[1-9])(?:[\s.\-]?\d{2}){4}/g;
  const exclusions = ["example.com","sentry.io","google","facebook","twitter",
    "youtube","wix","wordpress","adobe","jquery","png","jpg","svg","gif","webp",
    "schema.org","w3.org","cloudflare","analytics","pixel"];

  function emailValide(e) {
    e = e.toLowerCase();
    return !exclusions.some(x => e.includes(x)) && e.length > 6;
  }

  // ── 1. Emails dans les liens mailto ──────────────────────────────────────
  document.querySelectorAll("a[href^='mailto:']").forEach(a => {
    const email = a.href.replace("mailto:", "").split("?")[0].trim().toLowerCase();
    if (emailValide(email)) {
      // Chercher le nom à proximité
      const bloc = a.closest("article, li, div.card, div.fiche, [class*='result'], [class*='item'], [class*='avocat'], section") || a.parentElement;
      const nom = extraireNom(bloc);
      const tel = extraireTel(bloc);
      ajouterContact(results, { nom, email, telephone: tel, source: location.href });
    }
  });

  // ── 2. Emails dans le texte de la page ───────────────────────────────────
  const bodyText = document.body.innerText;
  const emailsTexte = [...new Set(bodyText.match(emailRe) || [])];
  emailsTexte.forEach(email => {
    if (emailValide(email) && !results.some(r => r.email === email.toLowerCase())) {
      ajouterContact(results, { nom: "", email: email.toLowerCase(), telephone: "", source: location.href });
    }
  });

  // ── 3. Noms + téléphones depuis les blocs structurés ────────────────────
  const selecteurs = [
    "article", "li.bi-item", "li.bi-pro", "[class*='bi-item']",
    "[class*='result-item']", "[class*='avocat-item']", "[class*='fiche']",
    "[class*='card']", "[class*='lawyer']", "[class*='contact']",
    ".vcard", "[itemtype*='Person']", "[itemtype*='LocalBusiness']"
  ];

  selecteurs.forEach(sel => {
    document.querySelectorAll(sel).forEach(bloc => {
      const nom = extraireNom(bloc);
      const email = extraireEmailBloc(bloc);
      const tel = extraireTel(bloc);
      if ((nom || email) && !results.some(r => r.email && r.email === email || r.nom && r.nom === nom)) {
        ajouterContact(results, { nom, email, telephone: tel, source: location.href });
      }
    });
  });

  // ── 4. Microdata / schema.org ────────────────────────────────────────────
  document.querySelectorAll("[itemprop='email']").forEach(el => {
    const email = (el.getAttribute("content") || el.innerText || "").trim().toLowerCase();
    if (emailValide(email)) {
      const bloc = el.closest("[itemtype]") || el.parentElement;
      const nom = document.querySelector("[itemprop='name']")?.innerText?.trim() || "";
      const tel = document.querySelector("[itemprop='telephone']")?.innerText?.trim() || "";
      ajouterContact(results, { nom, email, telephone: tel, source: location.href });
    }
  });

  // Dédupliquer
  const vus = new Set();
  return results.filter(c => {
    const cle = (c.email || c.nom || "").toLowerCase();
    if (!cle || vus.has(cle)) return false;
    vus.add(cle);
    return true;
  });
}

function extraireNom(bloc) {
  if (!bloc) return "";
  const sels = [
    "h1", "h2", "h3",
    "[class*='denomination']", "[class*='nom']", "[class*='name']",
    "[class*='title']", "[itemprop='name']", "strong", "b"
  ];
  for (const s of sels) {
    const el = bloc.querySelector(s);
    if (el) {
      const txt = el.innerText?.trim();
      if (txt && txt.length > 2 && txt.length < 80) return txt;
    }
  }
  return "";
}

function extraireEmailBloc(bloc) {
  if (!bloc) return "";
  const mailto = bloc.querySelector("a[href^='mailto:']");
  if (mailto) return mailto.href.replace("mailto:", "").split("?")[0].trim().toLowerCase();
  const emailRe = /[\w.\-+]{2,}@[\w.\-]+\.[a-zA-Z]{2,6}/;
  const m = emailRe.exec(bloc.innerText || "");
  return m ? m[0].toLowerCase() : "";
}

function extraireTel(bloc) {
  if (!bloc) return "";
  const telLink = bloc.querySelector("a[href^='tel:']");
  if (telLink) return telLink.href.replace("tel:", "").trim();
  const telRe = /(?:(?:\+33|0033|0)[1-9])(?:[\s.\-]?\d{2}){4}/;
  const m = telRe.exec(bloc.innerText || "");
  return m ? m[0].replace(/[\s.\-]/g, " ").trim() : "";
}

function ajouterContact(liste, contact) {
  if (!contact.email && !contact.nom) return;
  if (contact.email && !emailValide(contact.email)) return;
  liste.push(contact);
}

function emailValide(e) {
  const exclusions = ["example.com","sentry.io","google","facebook","twitter",
    "youtube","wix","wordpress","adobe","jquery","png","jpg","svg","gif","webp",
    "schema.org","w3.org","cloudflare","analytics","pixel"];
  e = (e || "").toLowerCase();
  return e.includes("@") && e.length > 6 && !exclusions.some(x => e.includes(x));
}
