#!/usr/bin/env python3
"""
Extraction d'établissements wellness (hôtels, centres thermaux, parcs aquatiques)
depuis Google Maps, SANS API — via automatisation de navigateur (Playwright).
Recherche par VILLE (voir scrape_wellness_gmaps_PORTUGAL.py pour la version par département).

Entrée  : villes_portugal.csv         (colonne "ville" ; une colonne "statut" est ajoutée/mise à jour)
Sortie  : resultats_villes_portugal.csv (écriture en temps réel, une ligne par établissement trouvé)

⚠️ Important :
- Ceci automatise un navigateur pour naviguer sur Google Maps comme un humain.
  Ça ne passe pas par l'API officielle, donc pas de clé/coût, mais c'est plus
  fragile (Google modifie régulièrement le HTML) et à utiliser modérément
  (délais volontaires entre les actions pour rester raisonnable).
- Si le script "casse" un jour, c'est probablement qu'un sélecteur CSS a changé
  côté Google : il faudra le remettre à jour (cf. commentaires REPÈRE ci-dessous).

Installation :
    pip install playwright requests
    playwright install chromium

Lancement :
    python scrape_wellness_gmaps_villes.py
"""

import csv
import os
import re
import sys
import time
from datetime import datetime
import requests
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

# ============================================================
# CONFIGURATION
# ============================================================

VILLES_CSV = "villes_portugal.csv"
RESULTATS_CSV = "resultats_villes_portugal.csv"

REQUETES = [
    ("hôtel spa",                  "Hôtel avec spa"),
    ("centre thermal",             "Centre thermal"),
    ("parc aquatique",             "Parc aquatique"),
    ("spa bien-être",              "Spa / bien-être"),
    ("centre de thalassothérapie", "Thalassothérapie"),
]

COLONNES_RESULTATS = [
    "Nom de l'établissement", "Adresse", "Ville", "Pays", "Activité",
    "Type d'établissement", "Nombre d'étoiles / Classement",
    "Téléphone", "Site Web", "Email", "LinkedIn", "URL Google Maps",
]

HEADLESS = True          # passe à False pour voir le navigateur travailler (utile pour déboguer)
MAX_ETABLISSEMENTS_PAR_REQUETE = 25   # limite raisonnable de scroll par recherche
PAUSE_COURTE = (1.0, 2.0)             # pause aléatoire entre actions (min, max) secondes
PAUSE_ENTRE_VILLES = (3.0, 6.0)
TIMEOUT_SITE_WEB = 6      # secondes, pour ne pas bloquer sur un site lent/mort (requests)

EMAIL_REGEX = re.compile(r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}")
EMAIL_DOMAINES_IGNORES = ("sentry.io", "wixpress.com", "example.com", "domain.com")
LINKEDIN_REGEX = re.compile(r"https?://(?:[a-z]{2,3}\.)?linkedin\.com/(?:company|showcase)/[a-zA-Z0-9\-_%]+/?", re.IGNORECASE)


def pause(bornes):
    import random
    time.sleep(random.uniform(*bornes))


# ============================================================
# Lecture / écriture villes_portugal.csv
# ============================================================

def read_text_any_encoding(path):
    """Essaie plusieurs encodages courants (Excel/Windows enregistre souvent
    en cp1252 ou latin-1, pas en UTF-8)."""
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            with open(path, encoding=enc) as f:
                return f.read(), enc
        except UnicodeDecodeError:
            continue
    raise ValueError(f"Impossible de décoder {path} avec les encodages testés.")


def load_cities(path):
    content, enc_used = read_text_any_encoding(path)
    print(f"[i] {path} lu avec l'encodage : {enc_used}")
    reader = csv.DictReader(content.splitlines())
    rows = list(reader)
    fieldnames = reader.fieldnames

    col_ville = None
    for c in fieldnames:
        if c.strip().lower() in ("ville", "city", "villes"):
            col_ville = c
            break
    if col_ville is None:
        col_ville = fieldnames[0]

    if "statut" not in fieldnames:
        fieldnames = list(fieldnames) + ["statut"]
        for r in rows:
            r["statut"] = ""

    return rows, fieldnames, col_ville


def save_cities(path, rows, fieldnames):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def mark_done(rows, fieldnames, col_ville, city, path):
    for r in rows:
        if r.get(col_ville, "").strip() == city:
            r["statut"] = f"traité ({datetime.now().strftime('%Y-%m-%d %H:%M')})"
    save_cities(path, rows, fieldnames)


# ============================================================
# Extraction Google Maps
# ============================================================

def clean_text(text):
    """Google Maps insère parfois des sauts de ligne invisibles avant le texte
    réel dans ses boutons (icônes, labels cachés) — on réduit tous les blancs
    consécutifs (espaces, retours à la ligne, tabulations) à un seul espace."""
    if not text:
        return ""
    return re.sub(r"\s+", " ", text).strip()


def extract_listing_details(page):
    """Extrait les infos du panneau de détail d'un établissement actuellement ouvert."""
    data = {
        "nom": "", "adresse": "", "telephone": "", "site_web": "",
        "note": "", "type_etab": "", "url": page.url,
    }

    try:
        data["nom"] = clean_text(page.locator("h1").first.inner_text(timeout=3000))
    except PWTimeout:
        pass

    # Note (étoiles)
    try:
        note_el = page.locator("div.F7nice span[aria-hidden='true']").first
        data["note"] = clean_text(note_el.inner_text(timeout=1500))
    except Exception:
        pass

    # Type d'établissement (ex: "Hôtel", "Spa")
    try:
        type_el = page.locator("button.DkEaL").first
        data["type_etab"] = clean_text(type_el.inner_text(timeout=1500))
    except Exception:
        pass

    # Blocs d'info (adresse, téléphone, site web) — REPÈRE : Google utilise des
    # boutons avec data-item-id, c'est la partie la plus susceptible de changer.
    try:
        buttons = page.locator("button[data-item-id], a[data-item-id]").all()
        for b in buttons:
            item_id = b.get_attribute("data-item-id") or ""
            try:
                text = clean_text(b.inner_text(timeout=1000))
            except Exception:
                continue
            if item_id.startswith("address"):
                data["adresse"] = text
            elif item_id.startswith("phone"):
                data["telephone"] = text
            elif item_id.startswith("authority") or "website" in item_id:
                href = b.get_attribute("href") or text
                data["site_web"] = href
    except Exception:
        pass

    return data


def guess_country_from_address(address):
    if not address:
        return ""
    parts = [p.strip() for p in address.split(",") if p.strip()]
    return parts[-1] if parts else ""


def extract_linkedin(html):
    m = LINKEDIN_REGEX.search(html)
    return m.group(0).rstrip("/") if m else ""


def extract_email(html):
    mailto_matches = re.findall(r'mailto:([a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})', html)
    for m in mailto_matches:
        if not any(dom in m.lower() for dom in EMAIL_DOMAINES_IGNORES):
            return m
    matches = EMAIL_REGEX.findall(html)
    for m in matches:
        if not any(dom in m.lower() for dom in EMAIL_DOMAINES_IGNORES):
            return m
    return ""


def fetch_site_info(website_url):
    """Récupère email + LinkedIn via une simple requête HTTP (pas de navigateur
    complet nécessaire pour lire du HTML) — beaucoup plus rapide que
    d'ouvrir un onglet Playwright pour chaque site externe."""
    result = {"email": "", "linkedin": ""}
    if not website_url:
        return result

    headers = {"User-Agent": "Mozilla/5.0 (compatible; DataExtractBot/1.0)"}

    try:
        resp = requests.get(website_url, timeout=TIMEOUT_SITE_WEB, headers=headers)
        html = resp.text
        result["linkedin"] = extract_linkedin(html)
        result["email"] = extract_email(html)
    except Exception:
        pass

    if not result["email"] or not result["linkedin"]:
        for suffix in ("/contact", "/contact-us", "/nous-contacter"):
            if result["email"] and result["linkedin"]:
                break
            try:
                resp2 = requests.get(website_url.rstrip("/") + suffix, timeout=5, headers=headers)
                html2 = resp2.text
                if not result["linkedin"]:
                    result["linkedin"] = extract_linkedin(html2)
                if not result["email"]:
                    result["email"] = extract_email(html2)
            except Exception:
                continue

    return result


def search_and_scrape(page, query, city, activite, out_writer, out_file, seen):
    search_text = f"{query} à {city}"
    url = f"https://www.google.com/maps/search/{search_text.replace(' ', '+')}/?hl=fr"
    print(f"  -> {activite} : {search_text}")

    try:
        page.goto(url, timeout=30000)
    except PWTimeout:
        print("     [!] timeout au chargement, on passe")
        return

    pause(PAUSE_COURTE)

    try:
        results_panel = page.locator('div[role="feed"]').first
        results_panel.wait_for(timeout=8000)
    except PWTimeout:
        print("     [!] pas de panneau de résultats trouvé (0 résultat ou page inattendue)")
        return

    collected_links = []
    previous_count = 0
    stagnant_rounds = 0
    while len(collected_links) < MAX_ETABLISSEMENTS_PAR_REQUETE and stagnant_rounds < 3:
        cards = results_panel.locator("a.hfpxzc").all()
        collected_links = list({c.get_attribute("href") for c in cards if c.get_attribute("href")})
        if len(collected_links) == previous_count:
            stagnant_rounds += 1
        else:
            stagnant_rounds = 0
        previous_count = len(collected_links)

        results_panel.evaluate("(el) => el.scrollBy(0, 800)")
        pause((0.8, 1.5))

    print(f"     {len(collected_links)} établissement(s) repérés")

    for href in collected_links[:MAX_ETABLISSEMENTS_PAR_REQUETE]:
        if href in seen:
            continue
        seen.add(href)
        try:
            page.goto(href, timeout=20000)
            page.wait_for_selector("h1", timeout=8000)
        except PWTimeout:
            continue

        pause((0.6, 1.2))
        details = extract_listing_details(page)
        if not details["nom"]:
            continue

        country = guess_country_from_address(details["adresse"])
        site_info = fetch_site_info(details["site_web"])

        row = [
            details["nom"], details["adresse"], city, country, activite,
            details["type_etab"], details["note"], details["telephone"],
            details["site_web"], site_info["email"], site_info["linkedin"], href,
        ]
        out_writer.writerow(row)
        out_file.flush()  # écriture immédiate


def main():
    if not os.path.exists(VILLES_CSV):
        print(f"[!] Fichier introuvable : {VILLES_CSV}")
        sys.exit(1)

    rows, fieldnames, col_ville = load_cities(VILLES_CSV)
    villes_a_traiter = [r[col_ville].strip() for r in rows
                         if r[col_ville].strip() and not r.get("statut", "").startswith("traité")]

    print(f"{len(villes_a_traiter)} ville(s) à traiter sur {len(rows)} au total")

    file_exists = os.path.exists(RESULTATS_CSV)
    seen = set()

    with open(RESULTATS_CSV, "a", newline="", encoding="utf-8-sig") as out_f:
        writer = csv.writer(out_f)
        if not file_exists:
            writer.writerow(COLONNES_RESULTATS)
            out_f.flush()

        with sync_playwright() as p:
            browser = p.chromium.launch(headless=HEADLESS)
            context = browser.new_context(locale="fr-FR")
            page = context.new_page()

            try:
                page.goto("https://www.google.com/maps?hl=fr", timeout=20000)
                consent_btn = page.locator("button:has-text('Tout accepter')").first
                if consent_btn.is_visible(timeout=3000):
                    consent_btn.click()
            except Exception:
                pass

            for city in villes_a_traiter:
                print(f"\n=== {city} ===")
                for query, activite in REQUETES:
                    search_and_scrape(page, query, city, activite, writer, out_f, seen)

                mark_done(rows, fieldnames, col_ville, city, VILLES_CSV)
                print(f"  [OK] {city} marquée 'traité' dans {VILLES_CSV}")
                pause(PAUSE_ENTRE_VILLES)

            browser.close()

    print(f"\nTerminé. Résultats dans {RESULTATS_CSV}")


if __name__ == "__main__":
    main()
