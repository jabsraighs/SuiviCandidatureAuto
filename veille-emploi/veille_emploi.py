#!/usr/bin/env python3
"""
Veille automatique d'offres d'emploi.

Deux sources combinées :
  1. Alertes email (Indeed/HelloWork) — source PRINCIPALE, fiable, basée
     sur des alertes que tu crées toi-même sur ces sites.
  2. Flux RSS Indeed — source de SECOURS, best-effort : si Indeed bloque
     ou change son flux, ça n'interrompt pas le script, ça log juste.

⚠️ Ce script NE POSTULE À RIEN automatiquement. Il se contente de
récupérer les nouvelles offres correspondant à tes recherches et de te
les lister dans un fichier. C'est ensuite à toi de cliquer et postuler,
manuellement, avec un CV/lettre adaptés à chaque offre.

Utilisation :
    python3 veille_emploi.py
"""

import json
import os
from datetime import datetime
from urllib.parse import quote

import feedparser
import requests

from config import RECHERCHES
from gmail_alertes_emploi import recuperer_offres_email

# Les recherches se configurent dans config.py, pas ici.

# Dossier de données persistantes (monté en volume Docker sur /app/data).
# En local sans Docker, ça crée simplement un sous-dossier "data" à côté du script.
DATA_DIR = os.environ.get("DATA_DIR", os.path.join(os.path.dirname(os.path.abspath(__file__)), "data"))
os.makedirs(DATA_DIR, exist_ok=True)

# Fichier où l'on mémorise les offres déjà vues, pour ne pas les
# reproposer à chaque exécution.
HISTORIQUE_PATH = os.path.join(DATA_DIR, "offres_vues.json")

# Fichier de sortie lisible, régénéré à chaque passage.
SORTIE_PATH = os.path.join(DATA_DIR, "nouvelles_offres.md")


def charger_historique():
    if os.path.exists(HISTORIQUE_PATH):
        with open(HISTORIQUE_PATH, "r", encoding="utf-8") as f:
            return set(json.load(f))
    return set()


def sauver_historique(liens_vus):
    with open(HISTORIQUE_PATH, "w", encoding="utf-8") as f:
        json.dump(sorted(liens_vus), f, ensure_ascii=False, indent=2)


def recuperer_offres_rss(mot_cle, ville):
    """Source de secours, best-effort. Ne lève jamais d'exception bloquante."""
    url = f"https://rss.indeed.com/rss?q={quote(mot_cle)}&l={quote(ville)}"
    try:
        resp = requests.get(url, timeout=15, headers={"User-Agent": "Mozilla/5.0"})
        resp.raise_for_status()
    except requests.RequestException as e:
        print(f"[RSS][!] '{mot_cle}' / '{ville}' indisponible ({e}). Source de secours, on continue.")
        return []

    feed = feedparser.parse(resp.content)
    offres = []
    for entree in feed.entries:
        offres.append(
            {
                "titre": entree.get("title", "Titre inconnu"),
                "lien": entree.get("link", ""),
                "resume": entree.get("summary", "")[:200],
                "date": entree.get("published", ""),
                "recherche": f"RSS — {mot_cle} / {ville}",
            }
        )
    return offres


def main():
    liens_vus = charger_historique()
    toutes_nouvelles = []

    # --- Source principale : alertes email ---
    try:
        offres_email = recuperer_offres_email()
    except Exception as e:
        print(f"[EMAIL][!] Erreur inattendue : {e}")
        offres_email = []

    for offre in offres_email:
        if offre["lien"] and offre["lien"] not in liens_vus:
            toutes_nouvelles.append(offre)
            liens_vus.add(offre["lien"])

    # --- Source de secours : flux RSS ---
    for mot_cle, ville in RECHERCHES:
        offres = recuperer_offres_rss(mot_cle, ville)
        for offre in offres:
            if offre["lien"] and offre["lien"] not in liens_vus:
                toutes_nouvelles.append(offre)
                liens_vus.add(offre["lien"])

    sauver_historique(liens_vus)

    horodatage = datetime.now().strftime("%Y-%m-%d %H:%M")
    with open(SORTIE_PATH, "w", encoding="utf-8") as f:
        f.write(f"# Nouvelles offres — {horodatage}\n\n")
        if not toutes_nouvelles:
            f.write("Aucune nouvelle offre depuis le dernier passage.\n")
        else:
            for o in toutes_nouvelles:
                f.write(f"## {o['titre']}\n")
                f.write(f"- Recherche : {o['recherche']}\n")
                if o["date"]:
                    f.write(f"- Publiée : {o['date']}\n")
                f.write(f"- Lien : {o['lien']}\n")
                if o["resume"]:
                    f.write(f"- Résumé : {o['resume']}...\n")
                f.write("\n")

    print(f"{len(toutes_nouvelles)} nouvelle(s) offre(s) trouvée(s). Voir {SORTIE_PATH}")


if __name__ == "__main__":
    main()