# -*- coding: utf-8 -*-
"""
gmail_alertes_emploi.py — Lit les emails d'ALERTES d'offres (Indeed, HelloWork...)
dans ta boîte Gmail et en extrait les offres, sans rien poster nulle part.

Prérequis (une seule fois, manuellement, hors de ce script) :
  1. Crée une alerte email sur Indeed et sur HelloWork pour chacune de tes
     recherches (mêmes mots-clés que dans config.py), envoyée vers l'adresse
     Gmail surveillée par ton tracker de candidatures.
  2. Place ton fichier credentials.json (OAuth Google Cloud) dans DATA_DIR.
     Tu peux réutiliser le même credentials.json que pour
     gmail_candidature_tracker.py si le projet Google Cloud est le même.
  3. Premier lancement : autorisation interactive via le navigateur, sur le
     port exposé (voir docker-compose.yml, 8080:8080). Un token.json est
     ensuite sauvegardé dans DATA_DIR pour les lancements suivants.

Ce module expose une seule fonction utile : recuperer_offres_email(),
appelée depuis veille_emploi.py.
"""

import base64
import http.server
import json
import os
import re
import urllib.parse
from datetime import datetime, timedelta

from bs4 import BeautifulSoup
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build

from config import EXPEDITEURS_ALERTES, JOURS_HISTORIQUE_INITIAL_EMAIL, LIENS_A_IGNORER

SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]

DATA_DIR = os.environ.get("DATA_DIR", os.path.join(os.path.dirname(os.path.abspath(__file__)), "data"))
os.makedirs(DATA_DIR, exist_ok=True)

CREDENTIALS_PATH = os.path.join(DATA_DIR, "credentials.json")
TOKEN_PATH = os.path.join(DATA_DIR, "token_alertes.json")
DERNIER_PASSAGE_PATH = os.path.join(DATA_DIR, "dernier_passage_email.json")
OAUTH_PORT = int(os.environ.get("OAUTH_PORT", "8080"))


def _log(msg):
    print(f"[gmail_alertes] {msg}", flush=True)


def _obtenir_credentials_via_navigateur():
    """
    Ouvre un mini-serveur HTTP pour récupérer le code OAuth de Google.

    On bind le serveur sur 0.0.0.0 (nécessaire pour que le port-forwarding
    Docker fonctionne), mais on annonce localhost comme redirect_uri à
    Google (obligatoire : Google refuse 0.0.0.0, seuls localhost/127.0.0.1
    sont acceptés pour une redirection en boucle locale).
    """
    flow = Flow.from_client_secrets_file(
        CREDENTIALS_PATH, scopes=SCOPES, redirect_uri=f"http://localhost:{OAUTH_PORT}/"
    )
    auth_url, _ = flow.authorization_url(access_type="offline", prompt="consent")
    print(
        "\n"
        "=== Autorisation Gmail requise ===\n"
        f"Ouvre cette URL dans ton navigateur :\n{auth_url}\n"
        "===================================\n",
        flush=True,
    )

    code_recu = {}

    class Handler(http.server.BaseHTTPRequestHandler):
        def do_GET(self):
            params = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
            self.send_response(200)
            self.send_header("Content-type", "text/html; charset=utf-8")
            self.end_headers()
            if "code" in params:
                code_recu["code"] = params["code"][0]
                self.wfile.write(
                    "<html><body><h1>Autorisation réussie, tu peux fermer cette fenêtre.</h1></body></html>".encode(
                        "utf-8"
                    )
                )
            else:
                self.wfile.write("<html><body><h1>Erreur : pas de code reçu.</h1></body></html>".encode("utf-8"))

        def log_message(self, format, *args):
            return  # silence les logs HTTP par défaut

    httpd = http.server.HTTPServer(("0.0.0.0", OAUTH_PORT), Handler)
    while "code" not in code_recu:
        httpd.handle_request()
    httpd.server_close()

    flow.fetch_token(code=code_recu["code"])
    return flow.credentials


def _authentifier():
    if not os.path.exists(CREDENTIALS_PATH):
        _log(
            f"Fichier absent : {CREDENTIALS_PATH}. "
            "Copie ton credentials.json (Google Cloud OAuth) dans le dossier data/ "
            "pour activer la lecture des alertes email."
        )
        return None

    creds = None
    if os.path.exists(TOKEN_PATH):
        creds = Credentials.from_authorized_user_file(TOKEN_PATH, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            creds = _obtenir_credentials_via_navigateur()
        with open(TOKEN_PATH, "w", encoding="utf-8") as f:
            f.write(creds.to_json())

    return creds


def _charger_date_dernier_passage():
    if os.path.exists(DERNIER_PASSAGE_PATH):
        with open(DERNIER_PASSAGE_PATH, "r", encoding="utf-8") as f:
            return datetime.fromisoformat(json.load(f)["date"])
    return datetime.now() - timedelta(days=JOURS_HISTORIQUE_INITIAL_EMAIL)


def _sauver_date_dernier_passage():
    with open(DERNIER_PASSAGE_PATH, "w", encoding="utf-8") as f:
        json.dump({"date": datetime.now().isoformat()}, f)


def _construire_requete_gmail(depuis):
    expediteurs = " OR ".join(f"from:{d}" for d in EXPEDITEURS_ALERTES)
    date_str = depuis.strftime("%Y/%m/%d")
    return f"({expediteurs}) after:{date_str}"


def _extraire_corps_html(payload):
    """Parcourt récursivement les parties du message pour trouver le corps HTML."""
    if payload.get("mimeType") == "text/html" and "data" in payload.get("body", {}):
        data = payload["body"]["data"]
        return base64.urlsafe_b64decode(data).decode("utf-8", errors="ignore")

    for part in payload.get("parts", []):
        html = _extraire_corps_html(part)
        if html:
            return html
    return None


def _lien_pertinent(texte, href):
    if not texte or len(texte.strip()) < 8:
        return False
    texte_lower = texte.lower()
    if any(ignore in texte_lower for ignore in LIENS_A_IGNORER):
        return False
    # On ne garde que les liens qui pointent vers une fiche de poste
    # (pattern large : Indeed, HelloWork, et la plupart des jobboards).
    if not re.search(r"(indeed\.com|hellowork\.com|viewjob|/emplois/|/jobs/|/job/)", href, re.IGNORECASE):
        return False
    return True


def recuperer_offres_email():
    """Retourne une liste d'offres extraites des emails d'alerte non encore vus."""
    creds = _authentifier()
    if not creds:
        return []

    try:
        service = build("gmail", "v1", credentials=creds)
        depuis = _charger_date_dernier_passage()
        requete = _construire_requete_gmail(depuis)

        resultat = service.users().messages().list(userId="me", q=requete, maxResults=50).execute()
        messages = resultat.get("messages", [])

        offres = []
        for m in messages:
            msg = service.users().messages().get(userId="me", id=m["id"], format="full").execute()
            headers = {h["name"]: h["value"] for h in msg["payload"].get("headers", [])}
            expediteur = headers.get("From", "")

            html = _extraire_corps_html(msg["payload"])
            if not html:
                continue

            soup = BeautifulSoup(html, "html.parser")
            for a in soup.find_all("a", href=True):
                texte = a.get_text(strip=True)
                href = a["href"]
                if _lien_pertinent(texte, href):
                    offres.append(
                        {
                            "titre": texte,
                            "lien": href,
                            "resume": "",
                            "date": headers.get("Date", ""),
                            "recherche": f"Alerte email — {expediteur}",
                        }
                    )

        _sauver_date_dernier_passage()
        return offres

    except Exception as e:
        _log(f"Erreur lors de la lecture des alertes email : {e}")
        return []