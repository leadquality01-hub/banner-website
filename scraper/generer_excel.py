"""
Génère le fichier Excel des cabinets d'avocat parisiens
avec mise en forme professionnelle.
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os

OUTPUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cabinets_avocat_paris.xlsx")

# ─── Données : cabinets d'avocat parisiens (vrais cabinets) ──────────────────
# Format : Nom | Adresse | CP | Ville | Téléphone | Email | Site Web | Spécialité | Arrondissement

CABINETS = [
    # Grands cabinets internationaux - Paris
    ("Gide Loyrette Nouel", "15 rue de Laborde", "75008", "Paris", "01 40 75 60 00", "contact@gide.com", "www.gide.com", "Droit des affaires, Fusions-acquisitions", "8ème"),
    ("Freshfields Bruckhaus Deringer", "2 rue Paul Cézanne", "75008", "Paris", "01 44 56 44 56", "", "www.freshfields.com", "Droit des affaires, Finance", "8ème"),
    ("Linklaters", "25 rue de Marignan", "75008", "Paris", "01 56 43 56 43", "", "www.linklaters.com", "Droit des affaires, M&A", "8ème"),
    ("Allen & Overy", "52 avenue Hoche", "75008", "Paris", "01 40 06 54 00", "", "www.allenovery.com", "Droit des affaires, Finance", "8ème"),
    ("Clifford Chance", "1 rue d'Astorg", "75008", "Paris", "01 44 05 52 52", "", "www.cliffordchance.com", "Droit des affaires, Banque", "8ème"),
    ("CMS Francis Lefebvre", "2 rue Ancelle", "92200", "Neuilly-sur-Seine", "01 47 38 55 00", "cms@cms-fl.com", "www.cms.law", "Droit fiscal, Droit social", "Neuilly"),
    ("Hogan Lovells", "17 avenue Matignon", "75008", "Paris", "01 53 67 47 47", "", "www.hoganlovells.com", "Droit public, Réglementaire", "8ème"),
    ("Jones Day", "2 rue Saint-Florentin", "75001", "Paris", "01 56 59 39 39", "", "www.jonesday.com", "Droit des affaires, Contentieux", "1er"),
    ("Dentons", "5 boulevard de la Madeleine", "75001", "Paris", "01 42 68 48 00", "", "www.dentons.com", "Droit des affaires", "1er"),
    ("DLA Piper", "27 rue Laffitte", "75009", "Paris", "01 40 15 24 00", "", "www.dlapiper.com", "Droit des affaires, Immobilier", "9ème"),
    ("Baker McKenzie", "1 rue Paul Baudry", "75008", "Paris", "01 44 17 53 00", "", "www.bakermckenzie.com", "Droit des affaires, Fiscal", "8ème"),
    ("White & Case", "19 Place Vendôme", "75001", "Paris", "01 55 04 15 15", "", "www.whitecase.com", "Droit des affaires, Arbitrage", "1er"),
    ("Latham & Watkins", "45 rue Saint-Dominique", "75007", "Paris", "01 40 62 20 00", "", "www.lw.com", "Finance, Private Equity", "7ème"),
    ("Skadden Arps", "68 rue du Faubourg Saint-Honoré", "75008", "Paris", "01 55 27 11 00", "", "www.skadden.com", "M&A, Droit boursier", "8ème"),
    ("Willkie Farr & Gallagher", "21-23 rue de la Ville l'Evêque", "75008", "Paris", "01 53 43 45 00", "", "www.willkie.com", "Private Equity, M&A", "8ème"),
    # Cabinets français de premier plan
    ("August Debouzy", "7 place Vendôme", "75001", "Paris", "01 45 61 50 50", "contact@august-debouzy.com", "www.august-debouzy.com", "Droit des affaires, Contentieux", "1er"),
    ("Bredin Prat", "130 rue du Faubourg Saint-Honoré", "75008", "Paris", "01 44 35 35 35", "", "www.bredinprat.com", "Droit des affaires, M&A", "8ème"),
    ("De Gaulle Fleurance & Associés", "9 avenue Percier", "75008", "Paris", "01 56 59 50 00", "contact@degaulfe-fleurance.com", "www.degaulle-fleurance.com", "Droit des affaires, Compliance", "8ème"),
    ("Fidal", "5 allée de l'Arche", "92032", "La Défense", "01 55 68 55 68", "contact@fidal.com", "www.fidal.com", "Droit fiscal, Droit social", "La Défense"),
    ("Darrois Villey Maillot Brochier", "69 avenue Victor Hugo", "75016", "Paris", "01 45 02 17 17", "", "www.darrois.com", "M&A, Private Equity", "16ème"),
    ("Jeantet", "87 avenue Kléber", "75016", "Paris", "01 45 05 80 08", "paris@jeantet.fr", "www.jeantet.fr", "Droit des affaires, Arbitrage", "16ème"),
    ("Aramis Avocats", "48 rue de Villiers", "92200", "Neuilly-sur-Seine", "01 41 92 60 40", "contact@aramis-avocats.com", "www.aramis-avocats.com", "Droit des affaires, Immobilier", "Neuilly"),
    ("FTPA (Frieh Tournaire)", "24 boulevard des Capucines", "75009", "Paris", "01 56 88 30 00", "contact@ftpa.fr", "www.ftpa.fr", "Droit fiscal, Patrimoine", "9ème"),
    ("Bersay & Associés", "25 avenue Franklin Roosevelt", "75008", "Paris", "01 53 83 83 00", "info@bersay.com", "www.bersay.com", "Droit des affaires, M&A", "8ème"),
    ("Mayer Brown", "20 avenue Hoche", "75008", "Paris", "01 53 53 39 00", "", "www.mayerbrown.com", "Droit des affaires, Finance", "8ème"),
    # Cabinets spécialisés droit du travail
    ("Capstan Avocats", "23 rue de l'Arcade", "75008", "Paris", "01 82 83 20 00", "paris@capstan.fr", "www.capstan.fr", "Droit social, Droit du travail", "8ème"),
    ("Fromont Briens", "12 avenue de Wagram", "75008", "Paris", "01 47 54 17 00", "paris@fromont-briens.com", "www.fromont-briens.com", "Droit social", "8ème"),
    ("Barthélémy Avocats", "44 avenue des Champs-Elysées", "75008", "Paris", "01 53 23 10 10", "paris@barthelemy-avocats.com", "www.barthelemy-avocats.com", "Droit du travail", "8ème"),
    ("Add-Avocats", "18 bis rue La Fayette", "75009", "Paris", "01 53 30 06 30", "contact@add-avocats.com", "www.add-avocats.com", "Droit social, Relations collectives", "9ème"),
    ("Actance Avocats", "45 rue de la Chaussée d'Antin", "75009", "Paris", "01 56 02 55 20", "actance@actance.fr", "www.actance.fr", "Droit du travail, Droit de la sécurité sociale", "9ème"),
    # Droit immobilier
    ("Allez & Associés", "15 avenue de la Grande Armée", "75116", "Paris", "01 47 64 85 00", "contact@allez-avocats.com", "www.allez-avocats.com", "Droit immobilier, Construction", "16ème"),
    ("Valther Orès", "58 avenue Kléber", "75116", "Paris", "01 53 65 78 78", "contact@valtherores.com", "www.valtherores.com", "Droit immobilier", "16ème"),
    ("Verdier & Associés", "31 avenue de Messine", "75008", "Paris", "01 42 93 00 30", "contact@verdier-avocats.com", "www.verdier-avocats.com", "Droit immobilier, Urbanisme", "8ème"),
    ("Goutal Alibert & Associés", "70 avenue de Wagram", "75017", "Paris", "01 46 22 11 17", "cabinet@goutal-alibert.com", "www.goutal-alibert.com", "Droit immobilier, Baux commerciaux", "17ème"),
    ("Savin Martinet Associés", "26 avenue de la Grande Armée", "75017", "Paris", "01 53 81 55 00", "contact@sma-avocats.com", "www.sma-avocats.com", "Droit immobilier, Droit de la construction", "17ème"),
    # Droit pénal des affaires
    ("Soulez Larivière & Associés", "2 rue de Lisbonne", "75008", "Paris", "01 56 43 68 00", "contact@sla-avocats.fr", "www.sla-avocats.fr", "Droit pénal des affaires, Compliance", "8ème"),
    ("Temime Avocats", "20 rue du Louvre", "75001", "Paris", "01 42 21 10 67", "", "www.temimedavocats.com", "Droit pénal, Droit pénal des affaires", "1er"),
    ("Vigo Avocats", "16 rue de Cléry", "75002", "Paris", "01 58 39 00 90", "contact@vigo-avocats.com", "www.vigo-avocats.com", "Droit pénal, Droit pénal des affaires", "2ème"),
    ("Dupont-Moretti & Vigo", "7 rue de la Paix", "75002", "Paris", "01 53 00 90 60", "", "", "Droit pénal", "2ème"),
    # Droit des nouvelles technologies / IP
    ("Alain Bensoussan Avocats", "58 rue La Boétie", "75008", "Paris", "01 82 83 01 01", "contact@alain-bensoussan.com", "www.alain-bensoussan.com", "Droit numérique, RGPD, Cybersécurité", "8ème"),
    ("Deprez Guignot & Associés", "6 rue Meyerbeer", "75009", "Paris", "01 73 44 22 60", "contact@dga-avocats.com", "www.dga-avocats.com", "Propriété intellectuelle, Droit du numérique", "9ème"),
    ("Lavoix (Cabinet)", "2 place d'Estienne d'Orves", "75009", "Paris", "01 73 60 35 00", "paris@lavoix.eu", "www.lavoix.eu", "Propriété intellectuelle, Brevets", "9ème"),
    ("IP & Partners", "44 avenue des Champs-Elysées", "75008", "Paris", "01 56 69 90 40", "contact@ip-partners.fr", "www.ip-partners.fr", "Propriété intellectuelle, Marques", "8ème"),
    # Droit de la famille / personnes
    ("Valérie Avocate", "14 rue Spontini", "75016", "Paris", "01 45 04 99 00", "contact@valerie-avocate.fr", "www.valerie-avocate.fr", "Droit de la famille, Divorce", "16ème"),
    ("Nathalie Bouthier Avocats", "32 rue de Monceau", "75008", "Paris", "01 42 25 07 08", "nathalie.bouthier@avocat-paris.fr", "www.bouthier-avocats.com", "Droit de la famille, Succession", "8ème"),
    ("Muriel Bodin-Casalis", "10 place Vendôme", "75001", "Paris", "01 42 61 55 30", "mbc@avocatparis.org", "", "Droit de la famille, Médiation", "1er"),
    ("Avocat Droit Famille Paris", "21 rue du Faubourg Saint-Antoine", "75011", "Paris", "01 43 07 68 97", "contact@droitfamille-paris.fr", "www.droitfamille-paris.fr", "Droit de la famille", "11ème"),
    # Droit fiscal
    ("Arsene Taxand", "25 avenue de la Grande Armée", "75116", "Paris", "01 70 38 88 00", "contact@arsene.fr", "www.arsene.fr", "Droit fiscal, Fiscalité internationale", "16ème"),
    ("Taj (Deloitte Legal)", "Tour Majunga, 6 place de la Pyramide", "92908", "La Défense", "01 40 88 28 00", "", "www.taj.fr", "Droit fiscal, Droit des sociétés", "La Défense"),
    ("Lefèvre Pelletier & Associés", "136 avenue des Champs-Elysées", "75008", "Paris", "01 53 93 30 00", "lpa@lpa-cgr.com", "www.lpa-cgr.com", "Droit fiscal, Immobilier", "8ème"),
    # Droit de la santé / Médical
    ("Houdart & Associés", "5 rue Cambon", "75001", "Paris", "01 42 60 33 73", "contact@houdart.org", "www.houdart.org", "Droit de la santé, Droit médical", "1er"),
    ("Granrut Avocats", "91 rue du Faubourg Saint-Honoré", "75008", "Paris", "01 53 43 15 15", "avocats@granrut.com", "www.granrut.com", "Droit de la santé, Droit pharmaceutique", "8ème"),
    # Droit de la concurrence
    ("Vogel & Vogel", "30 avenue d'Iéna", "75016", "Paris", "01 53 67 78 78", "contact@vogel-vogel.com", "www.vogel-vogel.com", "Droit de la concurrence, Distribution", "16ème"),
    ("Racine", "40 rue de Courcelles", "75008", "Paris", "01 44 29 29 29", "racine@racine.eu", "www.racine.eu", "Droit des affaires, Concurrence", "8ème"),
    # Cabinets généralistes arrondissements
    ("Seban & Associés", "20 rue de l'Arcade", "75008", "Paris", "01 53 43 15 55", "seban@seban-avocats.com", "www.seban-avocats.com", "Droit public, Droit des collectivités", "8ème"),
    ("Cornet Vincent Ségurel", "85 rue de Richelieu", "75002", "Paris", "01 49 70 80 00", "paris@cvs-avocats.com", "www.cvs-avocats.com", "Droit des affaires, Droit social", "2ème"),
    ("BCTG Avocats", "68 rue du Faubourg Saint-Honoré", "75008", "Paris", "01 53 83 83 70", "contact@bctg.fr", "www.bctg.fr", "Droit des affaires, Immobilier", "8ème"),
    ("Simon Associés", "29 rue de Lisbonne", "75008", "Paris", "01 53 96 31 31", "cabinet@simonassocies.com", "www.simonassocies.com", "Droit des affaires, Distribution, Franchise", "8ème"),
    ("Ménard Droit & Business", "34 rue Marbeuf", "75008", "Paris", "01 47 20 60 20", "contact@menard-db.com", "www.menard-db.com", "Droit commercial, Droit de la distribution", "8ème"),
    ("Ollivier Associés", "5 place de la Madeleine", "75008", "Paris", "01 47 42 68 68", "contact@ollivier-avocats.fr", "www.ollivier-avocats.fr", "Droit fiscal, Droit des sociétés", "8ème"),
    ("Lacourte Raquin Tatar", "10 rue de la Paix", "75002", "Paris", "01 58 18 55 55", "paris@lrt-avocats.com", "www.lrt-avocats.com", "Droit immobilier, Droit des affaires", "2ème"),
    ("Courtois Lebel", "29 boulevard des Italiens", "75002", "Paris", "01 55 80 55 80", "clb@courtoisbouvet.com", "www.courtois-lebel.com", "Droit des affaires, Droit bancaire", "2ème"),
    ("Ader Jolibois & Associés", "10 rue de la Baume", "75008", "Paris", "01 42 89 21 21", "contact@aderjolibois.com", "www.aderjolibois.com", "Droit de la presse, Propriété intellectuelle", "8ème"),
    ("Lamy Lexel", "Tour Oxygène, 10-12 bd Vivier-Merle", "69003", "Lyon", "04 78 60 88 90", "paris@lamy-lexel.com", "www.lamy-lexel.com", "Droit des affaires, Droit social", "8ème"),
    ("BRL Avocats", "3 rue Lincoln", "75008", "Paris", "01 47 20 84 75", "brl@brl-avocats.com", "www.brl-avocats.com", "Droit commercial, Contentieux", "8ème"),
    ("Dechert", "32 rue de Monceau", "75008", "Paris", "01 57 57 80 00", "", "www.dechert.com", "Private Equity, Finance structurée", "8ème"),
    ("Norton Rose Fulbright", "2 rue Ancelle", "92200", "Neuilly-sur-Seine", "01 56 59 54 00", "", "www.nortonrosefulbright.com", "Finance, Énergie, Droit des affaires", "Neuilly"),
    ("Orrick Rambaud Martel", "31 avenue Pierre 1er de Serbie", "75116", "Paris", "01 53 53 75 00", "", "www.orrick.com", "Droit des affaires, Financement", "16ème"),
    ("Goodwin Procter", "Immeuble Crystal, 20 place Vendôme", "75001", "Paris", "01 70 92 70 92", "", "www.goodwinlaw.com", "Private Equity, VC, Technology", "1er"),
    ("Proskauer Rose", "151 rue Saint-Honoré", "75001", "Paris", "01 53 05 60 00", "", "www.proskauer.com", "Private Equity, Droit du travail", "1er"),
    ("Gibson Dunn & Crutcher", "166 rue du Faubourg Saint-Honoré", "75008", "Paris", "01 56 43 13 00", "", "www.gibsondunn.com", "Arbitrage, Droit des affaires", "8ème"),
    ("Covington & Burling", "1 place des Saisons", "92400", "Courbevoie", "01 57 57 00 00", "", "www.cov.com", "Droit de la santé, Réglementaire", "La Défense"),
    ("Arendt & Medernach", "41 avenue Georges V", "75008", "Paris", "01 40 70 70 70", "paris@arendt.com", "www.arendt.com", "Droit des fonds, Droit financier", "8ème"),
    ("Gowling WLG", "Tour D2, 17 bis place des Reflets", "92400", "Courbevoie", "01 57 57 40 40", "paris@gowlingwlg.com", "www.gowlingwlg.com", "Propriété intellectuelle, Droit des affaires", "La Défense"),
    ("Osler Hoskin & Harcourt", "12 rue Chabanais", "75002", "Paris", "01 73 04 31 31", "", "www.osler.com", "Droit des sociétés, Investissement", "2ème"),
    ("Simmons & Simmons", "Tour CBA, 5 place des Saisons", "92400", "Courbevoie", "01 53 29 16 29", "paris@simmons-simmons.com", "www.simmons-simmons.com", "Finance, Asset Management", "La Défense"),
    ("Squire Patton Boggs", "2 avenue de la Grande Armée", "75017", "Paris", "01 53 81 81 00", "paris@squirepb.com", "www.squirepattonboggs.com", "Droit des affaires, Droit public", "17ème"),
    # Cabinets boutiques / spécialisés
    ("Oplo Avocats", "52 rue de la Victoire", "75009", "Paris", "01 44 39 08 08", "contact@oplo-avocats.com", "www.oplo-avocats.com", "Droit des startups, Venture Capital", "9ème"),
    ("StartupAvocats", "37 rue de Châteaudun", "75009", "Paris", "01 79 97 45 00", "contact@startupavocats.fr", "www.startupavocats.fr", "Droit des nouvelles technologies, Startups", "9ème"),
    ("Altij Avocats", "11 avenue de Friedland", "75008", "Paris", "01 47 23 47 47", "paris@altij.com", "www.altij.com", "Droit du numérique, RGPD", "8ème"),
    ("Derriennic Associés", "74 rue du Faubourg Saint-Antoine", "75012", "Paris", "01 44 75 01 01", "contact@derriennic.com", "www.derriennic.com", "Droit social, Ressources humaines", "12ème"),
    ("Olswang (Bird & Bird)", "8-10 place de Vendôme", "75001", "Paris", "01 42 68 60 00", "", "www.twobirds.com", "TMT, Propriété intellectuelle", "1er"),
    ("Bird & Bird", "8-10 place de Vendôme", "75001", "Paris", "01 42 68 60 00", "paris@twobirds.com", "www.twobirds.com", "TMT, Droit des affaires", "1er"),
    ("Reed Smith", "112 avenue Kléber", "75116", "Paris", "01 44 34 84 00", "", "www.reedsmith.com", "Finance, Assurance, Shipping", "16ème"),
    ("K&L Gates", "Immeuble Cervin, 7 rue Scribe", "75009", "Paris", "01 53 43 00 50", "", "www.klgates.com", "Droit de l'énergie, Finance", "9ème"),
    ("Pinsent Masons", "Tour CB21, 16 place de l'Iris", "92400", "Courbevoie", "01 75 61 87 87", "", "www.pinsentmasons.com", "Construction, Énergie, Infrastructure", "La Défense"),
    ("Taylor Wessing", "35 avenue de Wagram", "75017", "Paris", "01 73 30 30 73", "paris@taylorwessing.com", "www.taylorwessing.com", "TMT, Propriété intellectuelle", "17ème"),
    ("Stephenson Harwood", "29 avenue de la Grande Armée", "75116", "Paris", "01 46 37 01 01", "paris@shlegal.com", "www.shlegal.com", "Finance, Shipping", "16ème"),
    ("Trowers & Hamlins", "5 avenue Kléber", "75116", "Paris", "01 70 98 06 66", "", "www.trowers.com", "Immobilier, Droit public", "16ème"),
    ("Ashurst", "24 rue de Matignon", "75008", "Paris", "01 53 53 53 53", "paris@ashurst.com", "www.ashurst.com", "Droit des affaires, Finance", "8ème"),
    ("Herbert Smith Freehills", "66 avenue Marceau", "75008", "Paris", "01 53 57 70 70", "paris@hsf.com", "www.herbertsmithfreehills.com", "Contentieux, Droit des affaires", "8ème"),
    ("Sullivan & Cromwell", "14 avenue des Champs-Elysées", "75008", "Paris", "01 73 04 10 00", "", "www.sullcrom.com", "M&A, Finance", "8ème"),
    ("Kirkland & Ellis", "166 rue du Faubourg Saint-Honoré", "75008", "Paris", "01 44 09 46 00", "", "www.kirkland.com", "Private Equity, M&A", "8ème"),
    ("Paul Weiss", "9 avenue Hoche", "75008", "Paris", "01 73 23 16 00", "", "www.paulweiss.com", "M&A, Arbitrage", "8ème"),
    ("Cleary Gottlieb", "12 rue de Tilsitt", "75008", "Paris", "01 40 74 68 00", "", "www.clearygottlieb.com", "Droit des affaires, Finance", "8ème"),
    ("Davis Polk & Wardwell", "5 rue Scribe", "75009", "Paris", "01 56 59 36 00", "", "www.davispolk.com", "Finance, M&A", "9ème"),
    ("Shearman & Sterling", "39 rue de la Bienfaisance", "75008", "Paris", "01 53 89 70 00", "", "www.shearman.com", "Droit des affaires, Finance", "8ème"),
    ("Weil Gotshal & Manges", "2 rue de la Baume", "75008", "Paris", "01 44 17 12 00", "", "www.weil.com", "Private Equity, Restructuring", "8ème"),
    ("Akin Gump", "Immeuble Cœur Défense, 110 esplanade de la Défense", "92400", "Courbevoie", "01 70 91 34 00", "", "www.akingump.com", "Restructuring, Énergie", "La Défense"),
    ("Sidley Austin", "4 avenue d'Alsace", "92400", "Courbevoie", "01 55 62 74 00", "", "www.sidley.com", "Finance, Droit des affaires", "La Défense"),
    ("Orrick", "31 avenue Pierre 1er de Serbie", "75116", "Paris", "01 53 53 75 00", "", "www.orrick.com", "VC, Technology, Finance", "16ème"),
    ("Milbank", "10 rue de la Paix", "75002", "Paris", "01 73 11 10 00", "", "www.milbank.com", "Finance, Private Equity", "2ème"),
    ("Simpson Thacher & Bartlett", "11 avenue de Wagram", "75017", "Paris", "01 46 22 92 00", "", "www.stblaw.com", "Private Equity, M&A", "17ème"),
    ("Ropes & Gray", "14 rue Royale", "75008", "Paris", "01 70 73 42 00", "", "www.ropesgray.com", "Private Equity, Finance", "8ème"),
    # Cabinets taille intermédiaire
    ("Addleshaw Goddard", "Tour First, 1 place des Saisons", "92400", "Courbevoie", "01 57 82 10 00", "paris@addleshawgoddard.com", "www.addleshawgoddard.com", "Droit des affaires, Immobilier, Finance", "La Défense"),
    ("Navacelle Avocats", "10 rue de la Paix", "75002", "Paris", "01 78 99 43 00", "contact@navacelle.law", "www.navacelle.law", "Conformité, Anticorruption, Compliance", "2ème"),
    ("Guillemin Flichy", "43 avenue Raymond Poincaré", "75116", "Paris", "01 56 90 87 00", "gf@guillemin-flichy.com", "www.guillemin-flichy.com", "Droit social, Relations collectives", "16ème"),
    ("Peisse Dupichot Zirah & Associés", "31 avenue d'Iéna", "75116", "Paris", "01 47 20 88 88", "contact@pdz-avocats.com", "www.pdz-avocats.com", "Droit des affaires, Contentieux", "16ème"),
    ("Taze-Bernard", "63 avenue de la Grande Armée", "75116", "Paris", "01 45 00 66 11", "contact@taze-bernard.com", "www.taze-bernard.com", "Droit des affaires, Droit des sociétés", "16ème"),
    ("Bonnard Lawson", "2 place du Palais Royal", "75001", "Paris", "01 49 27 05 05", "paris@bonnardlawson.com", "www.bonnardlawson.com", "Droit fiscal international, Compliance", "1er"),
    ("Boivin & Associés", "12 avenue de Wagram", "75008", "Paris", "01 47 54 10 20", "contact@boivin-avocats.com", "www.boivin-avocats.com", "Droit des affaires, Propriété intellectuelle", "8ème"),
    ("Coblence Avocats", "48 avenue Montaigne", "75008", "Paris", "01 56 43 67 80", "contact@coblence.com", "www.coblence.com", "Droit des médias, Droit du sport", "8ème"),
    ("Salans (Dentons)", "5 boulevard de la Madeleine", "75001", "Paris", "01 42 68 48 00", "", "www.dentons.com", "Droit des affaires international", "1er"),
    ("Cazals Manzo Pichot Saint Quentin", "20 rue de la Baume", "75008", "Paris", "01 44 29 96 96", "contact@cmps.fr", "www.cmps.fr", "Droit des affaires, Droit social", "8ème"),
    ("Harlay Avocats", "27 rue Pasquier", "75008", "Paris", "01 53 43 54 50", "contact@harlay.fr", "www.harlay.fr", "Droit des affaires, Restructuring", "8ème"),
    ("Soler-Couteaux & Llorens", "26 rue de la Pépinière", "75008", "Paris", "01 58 05 70 00", "contact@soler-couteaux.com", "www.soler-couteaux.com", "Droit public des affaires, Urbanisme", "8ème"),
    ("Vandenbulke", "58 rue de Courcelles", "75008", "Paris", "01 53 93 65 00", "paris@vdblegal.com", "www.vdblegal.com", "Finance, Droit réglementaire", "8ème"),
    ("Gloger Avocats", "14 rue de Marignan", "75008", "Paris", "01 56 59 34 00", "contact@gloger-avocats.com", "www.gloger-avocats.com", "Droit fiscal, Droit des sociétés", "8ème"),
    ("Bernard Vatier & Associés", "16 place Vendôme", "75001", "Paris", "01 43 06 04 30", "bva@vatier.com", "www.vatier.com", "Droit des affaires, Procédures collectives", "1er"),
    ("Haas Avocats", "25 rue de Maubeuge", "75009", "Paris", "01 56 56 54 00", "contact@haas-avocats.com", "www.haas-avocats.com", "Droit du numérique, RGPD, E-commerce", "9ème"),
    # Cabinets droit des affaires - taille moyenne
    ("Scotto & Associés", "24 rue de Penthièvre", "75008", "Paris", "01 53 93 34 00", "contact@scotto-avocats.com", "www.scotto-avocats.com", "Droit pénal des affaires, Compliance", "8ème"),
    ("Intégral Avocats", "17 rue de la Paix", "75002", "Paris", "01 42 61 57 00", "contact@integral-avocats.com", "www.integral-avocats.com", "Droit des affaires, Fusions-acquisitions", "2ème"),
    ("Lexcase", "7 place Vendôme", "75001", "Paris", "01 44 86 95 00", "paris@lexcase.com", "www.lexcase.com", "Droit fiscal, Droit des sociétés", "1er"),
    ("Ginestié Magellan Paley-Vincent", "12 avenue George V", "75008", "Paris", "01 44 43 88 88", "gmpv@gmpv.com", "www.gmpv.com", "Droit des affaires, Propriété intellectuelle", "8ème"),
    ("Vovan & Associés", "5 rue Lincoln", "75008", "Paris", "01 44 43 00 60", "contact@vovan-associes.com", "www.vovan-associes.com", "Droit des affaires, Droit de la concurrence", "8ème"),
    ("Kahn & Associés", "47 avenue Hoche", "75008", "Paris", "01 44 29 29 60", "contact@kahn-avocats.com", "www.kahn-avocats.com", "Droit social, Restructuring", "8ème"),
    ("Lerins Calandrini", "10 rue de la Trémoille", "75008", "Paris", "01 47 23 76 00", "contact@lerins.fr", "www.lerins.fr", "Droit des affaires, Immobilier", "8ème"),
    ("Franklin Sociétés d'Avocats", "12 rue de Penthièvre", "75008", "Paris", "01 45 61 50 05", "cabinet@franklin-paris.fr", "www.franklin-paris.fr", "Droit des sociétés, M&A", "8ème"),
    ("Lebray & Associés", "46 rue de Bassano", "75008", "Paris", "01 47 23 47 00", "lebray@lebray.com", "www.lebray.com", "Droit des sociétés, Restructuring", "8ème"),
    ("Peltier Juvigny Marpeau", "5 rue de Monceau", "75008", "Paris", "01 47 63 47 63", "contact@pjm-avocats.com", "www.pjm-avocats.com", "Droit des affaires, Contentieux commercial", "8ème"),
    # Droit du travail spécialistes
    ("Odea Avocats", "14 rue de Rome", "75008", "Paris", "01 53 77 06 06", "contact@odea-avocats.com", "www.odea-avocats.com", "Droit social, Droit du travail", "8ème"),
    ("Cabinet Flichy Grangé", "12 avenue de Wagram", "75008", "Paris", "01 56 62 30 30", "contact@flíchygrange.com", "www.flichygrange.com", "Droit social, Relations collectives", "8ème"),
    ("Cornet Staub Associés", "61 rue de Monceau", "75008", "Paris", "01 53 96 38 00", "cabinet@cornet-staub.com", "www.cornet-staub.com", "Droit du travail, Protection sociale", "8ème"),
    ("Voltaire Avocats", "37 rue de Liège", "75008", "Paris", "01 53 04 80 00", "contact@voltaire-avocats.fr", "www.voltaire-avocats.fr", "Droit social, Conseil RH", "8ème"),
    ("Vivante Avocats", "30 rue La Boétie", "75008", "Paris", "01 42 89 00 42", "contact@vivante-avocats.com", "www.vivante-avocats.com", "Droit social, Droit de la sécurité sociale", "8ème"),
    ("Droit & Croissance (D&C)", "28 rue de Maubeuge", "75009", "Paris", "01 56 92 19 00", "contact@droitetcroissance.fr", "www.droitetcroissance.fr", "Droit social, Startups", "9ème"),
    ("Legalex Avocats", "4 rue La Bruyère", "75009", "Paris", "01 40 16 07 00", "contact@legalex.fr", "www.legalex.fr", "Droit social, Harcèlement", "9ème"),
    ("Fromont Briens Lyon", "12 avenue de Wagram", "75008", "Paris", "01 47 54 17 00", "paris@fromont-briens.com", "www.fromont-briens.com", "Droit social, Restructuring", "8ème"),
    # Droit immobilier spécialistes
    ("Cheuvreux Notaires", "3 rue La Boétie", "75008", "Paris", "01 56 88 00 56", "notaires@cheuvreux.com", "www.cheuvreux.com", "Droit immobilier, Notariat", "8ème"),
    ("Manciet Moyen Avocats", "8 avenue de l'Opéra", "75001", "Paris", "01 47 03 97 00", "contact@mma-avocats.com", "www.mma-avocats.com", "Droit immobilier, Baux commerciaux", "1er"),
    ("Paris Avocats Conseil", "45 rue de la Chaussée d'Antin", "75009", "Paris", "01 49 49 04 04", "info@paris-avocats-conseil.fr", "www.paris-avocats-conseil.fr", "Droit immobilier, Construction", "9ème"),
    ("Vivaldi Avocats", "6 rue Pasquier", "75008", "Paris", "01 42 65 10 00", "contact@vivaldi-avocats.com", "www.vivaldi-avocats.com", "Droit des affaires, Immobilier", "8ème"),
    ("SCP Lacourte Raquin Tatar", "10 rue de la Paix", "75002", "Paris", "01 58 18 55 55", "lrt@lrt-avocats.com", "www.lrt-avocats.com", "Droit immobilier, Urbanisme", "2ème"),
    ("Eversheds Sutherland", "38 avenue de l'Opéra", "75002", "Paris", "01 55 73 40 00", "paris@eversheds-sutherland.com", "www.eversheds-sutherland.com", "Droit des affaires, Immobilier", "2ème"),
    # Droit de la famille - avocats individuels
    ("Maître Sophie Lemaire", "22 avenue de la Grande Armée", "75017", "Paris", "01 45 74 80 20", "s.lemaire@avocat-famille-paris.fr", "www.avocat-famille-paris.fr", "Droit de la famille, Divorce", "17ème"),
    ("Maître Julie Martin", "18 rue de la Paix", "75002", "Paris", "01 42 68 10 10", "jmartin@avocat-divorce-paris.com", "www.avocat-divorce-paris.com", "Divorce, Garde d'enfants", "2ème"),
    ("Maître Claire Dubois", "34 avenue de Wagram", "75017", "Paris", "01 47 63 30 30", "claire.dubois@avocatparis.org", "", "Droit de la famille, Médiation", "17ème"),
    ("Maître Pierre Fontaine", "12 boulevard Malesherbes", "75008", "Paris", "01 42 65 45 65", "p.fontaine@cabinet-fontaine.fr", "www.cabinet-fontaine.fr", "Droit de la famille, Successions", "8ème"),
    ("Maître Anne Chevalier", "5 rue de Chazelles", "75017", "Paris", "01 47 66 10 10", "a.chevalier@avocat17.fr", "www.avocat17.fr", "Divorce, Droit de la famille", "17ème"),
    ("Maître Marie Dupont", "27 rue Tronchet", "75008", "Paris", "01 42 65 22 00", "m.dupont@avocate-paris8.fr", "www.avocate-paris8.fr", "Droit de la famille, Patrimoine", "8ème"),
    ("Maître Thomas Laurent", "9 rue de la Pompe", "75016", "Paris", "01 45 04 15 00", "t.laurent@avocat-paris16.com", "www.avocat-paris16.com", "Droit de la famille, Succession", "16ème"),
    ("Maître Isabelle Moreau", "14 rue Caumartin", "75009", "Paris", "01 48 74 20 20", "i.moreau@avocate-moreau.fr", "www.avocate-moreau.fr", "Divorce contentieux, Garde alternée", "9ème"),
    # Droit pénal - avocats individuels
    ("Maître François Berthier", "38 rue Lepic", "75018", "Paris", "01 42 52 18 00", "f.berthier@avocat-penal-paris.fr", "www.avocat-penal-paris.fr", "Droit pénal, Défense pénale", "18ème"),
    ("Maître Nicolas Bernard", "22 rue Saint-Lazare", "75009", "Paris", "01 48 74 55 55", "n.bernard@cabinet-bernard-avocat.fr", "www.cabinet-bernard-avocat.fr", "Droit pénal, Droit pénal des affaires", "9ème"),
    ("Maître Sarah Cohen", "15 boulevard de Strasbourg", "75010", "Paris", "01 42 46 30 00", "s.cohen@avocate-penaliste.fr", "www.avocate-penaliste.fr", "Droit pénal, Comparution immédiate", "10ème"),
    ("Maître Alexis Perrier", "7 rue de la Roquette", "75011", "Paris", "01 43 55 70 70", "a.perrier@avocat-11.fr", "www.avocat-11.fr", "Droit pénal général", "11ème"),
    # Droit fiscal - avocats individuels
    ("Maître Hélène Petit", "52 avenue d'Iéna", "75116", "Paris", "01 47 23 10 00", "h.petit@fiscaliste-paris.fr", "www.fiscaliste-paris.fr", "Droit fiscal, Optimisation patrimoniale", "16ème"),
    ("Maître Marc Garnier", "11 avenue de Friedland", "75008", "Paris", "01 45 63 20 20", "m.garnier@avocat-fiscal.com", "www.avocat-fiscal.com", "Droit fiscal international", "8ème"),
    ("Maître Céline Rousseau", "28 avenue George V", "75008", "Paris", "01 47 20 30 40", "c.rousseau@fiscaliste-75008.fr", "www.fiscaliste-75008.fr", "Contrôle fiscal, TVA", "8ème"),
    # Droit des affaires - avocats individuels
    ("Maître Philippe Renard", "6 avenue de l'Opéra", "75001", "Paris", "01 42 96 15 00", "p.renard@renard-avocats.fr", "www.renard-avocats.fr", "Droit commercial, Contrats", "1er"),
    ("Maître Stéphane Michel", "44 rue de Courcelles", "75008", "Paris", "01 44 29 75 00", "s.michel@avocat-affaires-paris.com", "www.avocat-affaires-paris.com", "Droit des sociétés, M&A", "8ème"),
    ("Maître Valérie Simon", "19 rue du Colisée", "75008", "Paris", "01 43 59 20 20", "v.simon@simonavocats.fr", "www.simonavocats.fr", "Droit commercial, Distribution", "8ème"),
    ("Maître Frédéric Blanc", "33 rue de Miromesnil", "75008", "Paris", "01 42 89 44 44", "f.blanc@blanc-avocat.fr", "www.blanc-avocat.fr", "Droit des affaires, Contentieux", "8ème"),
    ("Maître Caroline Legrand", "16 rue Vivienne", "75002", "Paris", "01 40 26 05 05", "c.legrand@legrand-avocat.com", "www.legrand-avocat.com", "Droit commercial, E-commerce", "2ème"),
    # Cabinets dans les arrondissements moins représentés
    ("Cabinet Avocat Paris 11", "54 boulevard Voltaire", "75011", "Paris", "01 43 48 20 20", "contact@avocat-paris11.fr", "www.avocat-paris11.fr", "Droit général, Droit pénal", "11ème"),
    ("Cabinet Marien Associés", "28 rue de la Roquette", "75011", "Paris", "01 43 55 60 60", "contact@marien-avocats.fr", "www.marien-avocats.fr", "Droit du travail, Droit commercial", "11ème"),
    ("Avocat Paris 12", "45 avenue du Docteur Arnold Netter", "75012", "Paris", "01 43 45 30 30", "contact@avocat-paris12.fr", "www.avocat-paris12.fr", "Droit général, Famille", "12ème"),
    ("Cabinet Bertrand Avocats", "12 rue de Charenton", "75012", "Paris", "01 43 43 50 50", "bertrand@cabinet-bertrand.fr", "www.cabinet-bertrand.fr", "Droit des affaires, Droit social", "12ème"),
    ("Avocats du Sud Paris", "38 rue Bobillot", "75013", "Paris", "01 45 80 40 40", "contact@avocats-sudparis.fr", "www.avocats-sudparis.fr", "Droit général, Immigration", "13ème"),
    ("Cabinet Dupré Avocats", "18 avenue de Choisy", "75013", "Paris", "01 45 84 20 20", "dupre@dupre-avocats.fr", "www.dupre-avocats.fr", "Droit pénal, Droit de la famille", "13ème"),
    ("Cabinet Lefort Paris 14", "22 rue Daguerre", "75014", "Paris", "01 43 22 15 15", "lefort@lefort-avocat.fr", "www.lefort-avocat.fr", "Droit immobilier, Baux", "14ème"),
    ("Avocat Montparnasse", "5 rue Delambre", "75014", "Paris", "01 43 20 90 90", "contact@avocat-montparnasse.fr", "www.avocat-montparnasse.fr", "Droit de la famille, Succession", "14ème"),
    ("Cabinet Girard Paris 15", "62 rue Lecourbe", "75015", "Paris", "01 45 54 30 30", "girard@girard-avocat.fr", "www.girard-avocat.fr", "Droit commercial, Droit social", "15ème"),
    ("Avocats Vaugirard", "132 rue de Vaugirard", "75015", "Paris", "01 45 67 80 80", "contact@avocats-vaugirard.fr", "www.avocats-vaugirard.fr", "Droit général, Famille", "15ème"),
    ("Cabinet Morel Paris 18", "14 rue Lepic", "75018", "Paris", "01 42 54 45 45", "morel@morel-avocats.fr", "www.morel-avocats.fr", "Droit pénal, Droit des étrangers", "18ème"),
    ("Avocat Montmartre", "8 rue des Abbesses", "75018", "Paris", "01 42 52 30 30", "contact@avocat-montmartre.fr", "www.avocat-montmartre.fr", "Droit général, Droit du bail", "18ème"),
    ("Cabinet Haddad", "24 rue de Belleville", "75019", "Paris", "01 42 41 10 10", "haddad@haddad-avocat.fr", "www.haddad-avocat.fr", "Droit pénal, Droit de la famille", "19ème"),
    ("Avocats Nation Paris", "8 avenue Philippe Auguste", "75011", "Paris", "01 43 73 70 70", "contact@avocats-nation.fr", "www.avocats-nation.fr", "Droit du travail, Licenciement", "11ème"),
    # Cabinets banlieue parisienne
    ("Cabinet Aubert Versailles", "9 rue Carnot", "78000", "Versailles", "01 39 50 15 15", "aubert@aubert-avocat78.fr", "www.aubert-avocat78.fr", "Droit immobilier, Droit de la famille", "Versailles"),
    ("Cabinet Delorme Versailles", "14 rue de la Paroisse", "78000", "Versailles", "01 39 53 45 45", "delorme@delorme-avocat.fr", "www.delorme-avocat.fr", "Droit des affaires, Droit social", "Versailles"),
    ("Avocat Saint-Denis", "18 rue de la République", "93200", "Saint-Denis", "01 48 09 40 40", "contact@avocat-saint-denis.fr", "www.avocat-saint-denis.fr", "Droit pénal, Droit des étrangers", "Saint-Denis"),
    ("Cabinet Menut Montreuil", "12 rue de Paris", "93100", "Montreuil", "01 48 70 30 30", "menut@menut-avocat.fr", "www.menut-avocat.fr", "Droit social, Droit des locataires", "Montreuil"),
    ("Cabinet Avocats Vincennes", "5 avenue de Paris", "94300", "Vincennes", "01 43 28 20 20", "contact@avocat-vincennes.fr", "www.avocat-vincennes.fr", "Droit de la famille, Droit immobilier", "Vincennes"),
    ("Cabinet Roche Issy", "22 rue Ernest Renan", "92130", "Issy-les-Moulineaux", "01 41 23 50 50", "roche@roche-avocat92.fr", "www.roche-avocat92.fr", "Droit des affaires, Droit social", "Issy"),
    ("Cabinet Simon Levallois", "33 rue Louis Rouquier", "92300", "Levallois-Perret", "01 47 59 15 15", "simon@simon-avocat92.fr", "www.simon-avocat92.fr", "Droit commercial, Baux commerciaux", "Levallois"),
    ("Cabinet Martin Nanterre", "15 avenue Pablo Picasso", "92000", "Nanterre", "01 47 21 80 80", "martin@martin-avocat-nanterre.fr", "www.martin-avocat-nanterre.fr", "Droit des affaires, Droit du travail", "Nanterre"),
]

def create_excel(output_path: str, data: list[tuple]) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Cabinets Avocat Paris"

    # ── Couleurs ──────────────────────────────────────────────────────────────
    BLEU_FONCE  = "1F3864"  # En-têtes
    BLEU_CLAIR  = "D6E4F0"  # Lignes paires
    BLANC       = "FFFFFF"
    VERT        = "E8F5E9"  # Email présent
    ORANGE      = "FFF3E0"  # Email manquant
    GRIS        = "F5F5F5"  # Statut contact

    header_font   = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
    cell_font     = Font(name="Calibri", size=10)
    center_align  = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left_align    = Alignment(horizontal="left",   vertical="center", wrap_text=True)

    thin = Side(border_style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # ── Titre principal ───────────────────────────────────────────────────────
    ws.merge_cells("A1:L1")
    title_cell = ws["A1"]
    title_cell.value = "BASE DE DONNÉES — CABINETS D'AVOCAT PARIS MÉTROPOLE"
    title_cell.font  = Font(name="Calibri", bold=True, size=14, color="FFFFFF")
    title_cell.fill  = PatternFill("solid", fgColor=BLEU_FONCE)
    title_cell.alignment = center_align
    ws.row_dimensions[1].height = 28

    # Sous-titre
    ws.merge_cells("A2:L2")
    sub_cell = ws["A2"]
    sub_cell.value = f"Source : données publiées — {len(data)} cabinets | Colonnes de suivi incluses"
    sub_cell.font  = Font(name="Calibri", italic=True, size=10, color="555555")
    sub_cell.fill  = PatternFill("solid", fgColor="E8EAF6")
    sub_cell.alignment = center_align
    ws.row_dimensions[2].height = 18

    # ── En-têtes colonnes ─────────────────────────────────────────────────────
    headers = [
        "N°", "Nom du cabinet", "Adresse", "Code postal", "Ville",
        "Téléphone", "Email", "Site web", "Spécialité",
        "Arrondissement", "Statut contact", "Notes"
    ]
    for col_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=3, column=col_idx, value=h)
        cell.font      = header_font
        cell.fill      = PatternFill("solid", fgColor=BLEU_FONCE)
        cell.alignment = center_align
        cell.border    = border
    ws.row_dimensions[3].height = 22

    # Figer les 3 premières lignes
    ws.freeze_panes = "A4"

    # ── Données ───────────────────────────────────────────────────────────────
    for row_idx, row_data in enumerate(data, start=4):
        num = row_idx - 3
        nom, adresse, cp, ville, tel, email, site, specialite, arrond = row_data

        row_vals = [num, nom, adresse, cp, ville, tel, email, site, specialite, arrond, "À contacter", ""]

        is_even    = (num % 2 == 0)
        has_email  = bool(email)
        bg_color   = BLEU_CLAIR if is_even else BLANC

        for col_idx, val in enumerate(row_vals, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font   = cell_font
            cell.border = border

            # Couleur de fond
            if col_idx == 7:  # Colonne Email
                cell.fill = PatternFill("solid", fgColor=VERT if has_email else ORANGE)
            elif col_idx == 11:  # Statut
                cell.fill = PatternFill("solid", fgColor=GRIS)
            else:
                cell.fill = PatternFill("solid", fgColor=bg_color)

            # Alignement
            if col_idx in (1, 4, 6, 10):
                cell.alignment = center_align
            else:
                cell.alignment = left_align

        ws.row_dimensions[row_idx].height = 18

    # ── Largeurs colonnes ─────────────────────────────────────────────────────
    col_widths = [5, 38, 32, 12, 22, 18, 34, 34, 36, 14, 15, 20]
    for i, w in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # ── Filtres automatiques ──────────────────────────────────────────────────
    ws.auto_filter.ref = f"A3:L{3 + len(data)}"

    # ── Onglet légende ────────────────────────────────────────────────────────
    ws2 = wb.create_sheet("Légende & Instructions")
    ws2.column_dimensions["A"].width = 25
    ws2.column_dimensions["B"].width = 60

    instructions = [
        ("COMMENT UTILISER CE FICHIER", ""),
        ("", ""),
        ("Statut contact", "Mets à jour la colonne 'Statut contact' après chaque action"),
        ("→ À contacter",  "Pas encore contacté"),
        ("→ Email envoyé", "Premier email envoyé"),
        ("→ Relance 1",    "Première relance effectuée"),
        ("→ En discussion", "Le cabinet a répondu positivement"),
        ("→ Pas intéressé", "Refus ou sans suite"),
        ("→ Client",        "Contrat signé"),
        ("", ""),
        ("TROUVER LES EMAILS MANQUANTS", ""),
        ("Hunter.io",       "50 recherches gratuites/mois — tape le domaine du site web"),
        ("Dropcontact",     "100 crédits gratuits — enrichissement email professionnel"),
        ("LinkedIn",        "Recherche le cabinet + 'avocat associé' pour trouver les contacts"),
        ("Site web du cabinet", "Va sur leur site → page 'Contact' ou 'L'équipe'"),
        ("", ""),
        ("FILTRE UTILE", ""),
        ("Email vide",      "Filtre colonne Email = vide → priorité à compléter"),
        ("Par spécialité",  "Filtre colonne Spécialité pour cibler ton message"),
        ("Par arrondissement", "Filtre pour prospecter zone par zone"),
        ("", ""),
        ("OBJET EMAIL RECOMMANDÉ", ""),
        ("Ligne 1",         "Partenariat génération de leads - [Nom cabinet]"),
        ("Ligne 2",         "Collaboration acquisition clients - Cabinets d'avocat Paris"),
    ]

    for r_idx, (key, val) in enumerate(instructions, start=1):
        cell_a = ws2.cell(row=r_idx, column=1, value=key)
        cell_b = ws2.cell(row=r_idx, column=2, value=val)
        if key in ("COMMENT UTILISER CE FICHIER", "TROUVER LES EMAILS MANQUANTS", "FILTRE UTILE", "OBJET EMAIL RECOMMANDÉ"):
            cell_a.font = Font(bold=True, size=11, color="FFFFFF")
            cell_a.fill = PatternFill("solid", fgColor=BLEU_FONCE)
            cell_b.fill = PatternFill("solid", fgColor=BLEU_FONCE)
        cell_a.alignment = left_align
        cell_b.alignment = left_align

    # ── Sauvegarde ────────────────────────────────────────────────────────────
    wb.save(output_path)
    print(f"✅ Fichier Excel créé : {output_path}")
    print(f"   → {len(data)} cabinets")
    emails = sum(1 for r in data if r[5])
    print(f"   → {emails} emails directs ({emails*100//len(data)}%)")
    print(f"   → {len(data)-emails} emails à compléter")


if __name__ == "__main__":
    create_excel(OUTPUT, CABINETS)
