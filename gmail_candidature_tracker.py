#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gmail_candidature_tracker.py
=============================
Scanne la boîte Gmail de l'utilisateur, détecte les emails liés à une
recherche d'emploi (candidature envoyée, accusé de réception, entretien,
refus, offre) et met à jour un fichier Excel (candidatures.xlsx) avec le
suivi.

Conçu pour être lancé automatiquement (cron / Tâches planifiées Windows) :
chaque exécution ne traite QUE les nouveaux emails depuis la dernière
exécution, et n'ajoute jamais de doublon.

Configuration : voir config.py (mots-clés, dossier de sortie, etc.)
Installation & planification : voir README.md
"""

import os
import re
import sys
import json
import base64
import unicodedata
from datetime import datetime, timedelta

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

import config

# ---------------------------------------------------------------------------
# Chemins
# ---------------------------------------------------------------------------
# DATA_DIR regroupe tous les fichiers persistants (credentials, token, état,
# fichier Excel) dans un seul dossier : pratique pour ne monter qu'un seul
# volume Docker et retrouver facilement le fichier Excel depuis Windows.
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.environ.get("DATA_DIR", BASE_DIR)
os.makedirs(DATA_DIR, exist_ok=True)

TOKEN_PATH = os.path.join(DATA_DIR, "token.json")
CREDENTIALS_PATH = os.path.join(DATA_DIR, "credentials.json")
STATE_PATH = os.path.join(DATA_DIR, "state.json")
OUTPUT_PATH = os.path.join(DATA_DIR, config.OUTPUT_FILENAME)

# Port utilisé pour l'autorisation OAuth (fixe, pour pouvoir le publier
# avec -p 8080:8080 quand le script tourne dans Docker).
OAUTH_PORT = int(os.environ.get("OAUTH_PORT", "8080"))

SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]

HEADERS = [
    "Date candidature", "Entreprise", "Poste", "Statut actuel",
    "Dernière mise à jour", "Historique statuts",
    "Objet du dernier email", "Expéditeur", "Lien Gmail (dernier)", "ID messages",
]


def normaliser(texte):
    """Normalise un texte pour comparaison (minuscule, sans accents/ponctuation)."""
    texte = texte or ""
    texte = unicodedata.normalize("NFKD", texte).encode("ascii", "ignore").decode()
    texte = re.sub(r"[^a-z0-9]+", "", texte.lower())
    return texte


# ---------------------------------------------------------------------------
# Authentification Gmail
# ---------------------------------------------------------------------------
def get_gmail_service():
    creds = None
    if os.path.exists(TOKEN_PATH):
        creds = Credentials.from_authorized_user_file(TOKEN_PATH, SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not os.path.exists(CREDENTIALS_PATH):
                raise FileNotFoundError(
                    f"Fichier '{CREDENTIALS_PATH}' introuvable. "
                    "Suis les instructions du README.md pour créer tes "
                    "identifiants OAuth Google Cloud."
                )
            flow = InstalledAppFlow.from_client_secrets_file(
                CREDENTIALS_PATH, SCOPES
            )
            print(
                "\n>>> Ouvre ce lien dans TON navigateur (sur ta machine "
                "Windows, pas dans le conteneur) pour autoriser l'accès à "
                "Gmail : le script attend que tu valides.\n"
            )
            # open_browser=False : indispensable dans Docker/Linux headless,
            # il n'y a pas de navigateur à l'intérieur du conteneur. Le port
            # est fixe (OAUTH_PORT) pour pouvoir être publié avec
            # "-p 8080:8080" et joignable depuis le navigateur de l'hôte.
            creds = flow.run_local_server(
                port=OAUTH_PORT, open_browser=False,
                bind_addr="0.0.0.0",
            )

        with open(TOKEN_PATH, "w") as token_file:
            token_file.write(creds.to_json())

    return build("gmail", "v1", credentials=creds)


# ---------------------------------------------------------------------------
# Etat (dernière exécution) pour ne traiter que les nouveaux emails
# ---------------------------------------------------------------------------
def load_state():
    if os.path.exists(STATE_PATH):
        with open(STATE_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"last_run": None, "processed_ids": []}


def save_state(state):
    with open(STATE_PATH, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)


# ---------------------------------------------------------------------------
# Construction de la requête de recherche Gmail
# ---------------------------------------------------------------------------
def build_query(last_run):
    keywords = config.MOTS_CLES_RECHERCHE
    or_clause = " OR ".join(f'"{k}"' if " " in k else k for k in keywords)
    query = f"({or_clause})"

    if last_run:
        # Gmail n'accepte que des dates (jour près), on part de la veille
        # du dernier passage pour ne rien rater, les doublons étant filtrés
        # ensuite via processed_ids.
        since = (datetime.fromisoformat(last_run) - timedelta(days=1)).strftime("%Y/%m/%d")
        query += f" after:{since}"
    else:
        since = (datetime.now() - timedelta(days=config.JOURS_HISTORIQUE_INITIAL)).strftime("%Y/%m/%d")
        query += f" after:{since}"

    return query


# ---------------------------------------------------------------------------
# Classification du statut de la candidature
# ---------------------------------------------------------------------------
def detect_statut(texte):
    texte = texte.lower()
    for statut, mots in config.CLASSIFICATION.items():
        for mot in mots:
            if mot.lower() in texte:
                return statut
    return "Candidature envoyée"


def ressemble_a_un_nom_de_personne(nom):
    # Détecte la convention française "Prénom NOM" (nom de famille tout en
    # MAJUSCULES, ex. "Agathe COLLINET") pour éviter de le prendre pour
    # l'entreprise. Ne matche pas "Sopra Steria" ou "Hellowork Candidatu."
    # car leur second mot n'est pas entièrement en majuscules.
    return bool(re.match(r"^[A-ZÀ-Ý][a-zà-ÿ\-]+\s+[A-ZÀ-Ý]{2,}(\s|$|\.)", nom.strip()))


def extraire_entreprise(sender_email, sender_name, subject):
    subject = subject.strip()

    # 1) pattern très fiable "... chez X" en fin de sujet (HelloWork,
    #    Sopra Steria, Basic-Fit...). Prioritaire car il donne le vrai nom
    #    de l'entreprise même quand l'expéditeur est une plateforme tierce.
    m = re.search(r"chez\s+([A-Za-zÀ-ÿ0-9&'\.\- ]{2,40})\s*[!\.]?$", subject)
    if m:
        return m.group(1).strip(" .,-!")

    # 2) pattern "Votre candidature {Entreprise}" (ex : "Votre candidature
    #    Hardis Group") utilisé par certains ATS sans "chez" ni "poste de"
    m = re.search(r"^votre candidature\s+([A-Za-zÀ-ÿ0-9&'\.\- ]{2,40})$", subject, re.IGNORECASE)
    if m:
        return m.group(1).strip(" .,-!")

    est_plateforme = sender_name and any(
        p.lower() in sender_name.lower() for p in config.PLATEFORMES_A_IGNORER
    )
    est_nom_personne = sender_name and ressemble_a_un_nom_de_personne(sender_name)

    # 3) nom affiché de l'expéditeur, sauf si c'est une plateforme tierce
    #    ou le nom d'une personne (un recruteur n'est pas l'entreprise)
    if sender_name and not est_plateforme and not est_nom_personne:
        nom = re.sub(
            r"(?i)\b(recrutement|recruiting|talent|rh|hr|careers?|jobs?|"
            r"no[- ]?reply|équipe|team|candidature)\b",
            "",
            sender_name,
        ).strip(" -|,")
        if len(nom) > 1:
            return nom.title() if nom.islower() else nom

    # 4) sinon déduit du domaine de l'email (utile aussi pour les noms de
    #    personnes : ex. a.collinet@soprasteria.com -> Soprasteria)
    domaine = sender_email.split("@")[-1].split(".")[0]
    domaine = re.sub(r"(?i)^(mail|no-?reply|jobs|careers|talent|smart|welcome)", "", domaine)
    if domaine and domaine.lower() not in ("gmail", "outlook", "yahoo", "hotmail"):
        return domaine.capitalize()

    # 5) dernier recours : le nom affiché tel quel (ex. plateforme)
    return sender_name or sender_email


def extraire_poste(subject):
    # on retire d'abord le suffixe "chez X" pour ne pas le confondre
    # avec l'intitulé du poste
    sujet = re.sub(r"\s*chez\s+[A-Za-zÀ-ÿ0-9&'\.\- ]{2,40}\s*[!\.]?$", "", subject.strip())

    patterns = [
        r"pour\s+être\s+([A-Za-zÀ-ÿ0-9\-/' ]{3,60})",
        r"(?:au|pour le)\s+poste\s+(?:de|d')\s*([A-Za-zÀ-ÿ0-9\-/'&() ]{3,60})",
        r"candidature\s+pour\s+([A-Za-zÀ-ÿ0-9\-/'() ]{3,60})",
        r"pour\s+le\s+poste\s+d[e']\s*([A-Za-zÀ-ÿ0-9\-/'&() ]{3,60})",
        r"candidature\s+de\s+([A-Za-zÀ-ÿ0-9\-/'() ]{3,60}?)(?:\s+n['e]|\s*$)",
    ]
    for p in patterns:
        m = re.search(p, sujet, re.IGNORECASE)
        if m:
            return m.group(1).strip(" .,-")
    return ""


# ---------------------------------------------------------------------------
# Lecture des messages Gmail
# ---------------------------------------------------------------------------
def doit_exclure(sender_email, subject):
    """True si l'email est une alerte/newsletter d'offres d'emploi ou une
    notification de compte, PAS une vraie candidature — malgré le mot
    "candidature"/"recrutement" présent dans le texte."""
    sender_email = (sender_email or "").lower()
    domaine = sender_email.split("@")[-1]

    if sender_email in [e.lower() for e in config.EXPEDITEURS_EXACTS_A_EXCLURE]:
        return True

    if any(domaine == d or domaine.endswith("." + d) for d in config.EXPEDITEURS_DOMAINES_A_EXCLURE):
        return True

    subj = (subject or "").lower()
    if any(mot.lower() in subj for mot in config.MOTS_CLES_EXCLUSION_OBJET):
        return True

    return False


def get_header(headers, name):
    for h in headers:
        if h["name"].lower() == name.lower():
            return h["value"]
    return ""


def parse_sender(raw_from):
    m = re.match(r'^"?([^"<]*)"?\s*<?([\w\.\-+]+@[\w\.\-]+)>?$', raw_from.strip())
    if m:
        return m.group(1).strip(), m.group(2).strip()
    return "", raw_from


def fetch_new_messages(service, query, processed_ids):
    nouveaux = []
    ids_vus = []
    page_token = None

    while True:
        resp = service.users().messages().list(
            userId="me", q=query, pageToken=page_token, maxResults=100
        ).execute()

        for m in resp.get("messages", []):
            if m["id"] in processed_ids:
                continue

            msg = service.users().messages().get(
                userId="me", id=m["id"], format="metadata",
                metadataHeaders=["Subject", "From", "Date"],
            ).execute()

            headers = msg["payload"]["headers"]
            subject = get_header(headers, "Subject")
            raw_from = get_header(headers, "From")
            date_str = get_header(headers, "Date")
            snippet = msg.get("snippet", "")

            sender_name, sender_email = parse_sender(raw_from)

            ids_vus.append(m["id"])

            if doit_exclure(sender_email, subject):
                continue

            try:
                date_obj = datetime.strptime(
                    date_str.split(" (")[0].strip(), "%a, %d %b %Y %H:%M:%S %z"
                )
                # On retire le fuseau horaire après coup : il ne sert qu'au
                # tri chronologique entre emails, et certains en-têtes Date
                # n'en fournissent pas -> mélanger naive/aware plante le tri.
                date_obj = date_obj.replace(tzinfo=None)
                date_email = date_obj.strftime("%d/%m/%Y")
            except ValueError:
                date_obj = datetime.min
                date_email = date_str

            nouveaux.append({
                "id": m["id"],
                "date": date_email,
                "date_obj": date_obj,
                "entreprise": extraire_entreprise(sender_email, sender_name, subject),
                "poste": extraire_poste(subject),
                "statut": detect_statut(subject + " " + snippet),
                "objet": subject,
                "expediteur": sender_email,
                "lien": f"https://mail.google.com/mail/u/0/#all/{m['id']}",
            })

        page_token = resp.get("nextPageToken")
        if not page_token:
            break

    # Tri chronologique croissant : indispensable pour que, lors de la
    # fusion, le DERNIER email traité pour une candidature soit bien le
    # plus récent (sinon un statut plus ancien pourrait écraser un plus
    # récent si Gmail les avait renvoyés dans le désordre).
    nouveaux.sort(key=lambda e: e["date_obj"])
    return nouveaux, ids_vus


# ---------------------------------------------------------------------------
# Excel : création / mise à jour
# ---------------------------------------------------------------------------
def ouvrir_ou_creer_classeur():
    if os.path.exists(OUTPUT_PATH):
        wb = load_workbook(OUTPUT_PATH)
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Candidatures"
        ws.append(HEADERS)
        for col_idx, header in enumerate(HEADERS, start=1):
            cell = ws.cell(row=1, column=col_idx)
            cell.font = Font(bold=True, color="FFFFFF", name="Arial")
            cell.fill = PatternFill("solid", fgColor="4472C4")
            cell.alignment = Alignment(horizontal="center")
        widths = [14, 22, 28, 18, 16, 32, 40, 28, 38, 24]
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
        ws.freeze_panes = "A2"
    return wb, ws


STATUT_COULEURS = {
    "Candidature envoyée": "FFF2CC",
    "Reçue / en cours": "DDEBF7",
    "Entretien": "D9EAD3",
    "Refus": "F4CCCC",
    "Offre / Accepté": "B6D7A8",
}


def construire_index(ws):
    """Associe chaque candidature déjà présente (entreprise+poste normalisés)
    à son numéro de ligne, pour savoir si un nouvel email correspond à une
    candidature déjà suivie."""
    index = {}
    for r in range(2, ws.max_row + 1):
        entreprise = ws.cell(row=r, column=2).value or ""
        poste = ws.cell(row=r, column=3).value or ""
        cle = normaliser(entreprise) + "||" + normaliser(poste)
        index[cle] = r
    return index


def parser_date_fr(date_str):
    """Convertit une date au format 'dd/mm/YYYY' en objet datetime pour le
    tri ; si le format est inattendu, renvoie la date la plus ancienne
    possible pour que la ligne finisse en bas (le tri étant décroissant)
    plutôt que de planter."""
    try:
        return datetime.strptime(date_str, "%d/%m/%Y")
    except (ValueError, TypeError):
        return datetime.min


def trier_par_date(ws):
    """Retrie toutes les lignes du classeur par 'Date candidature'
    décroissante (les plus récentes en haut), puis par 'Dernière mise à
    jour' décroissante en cas d'égalité. Conserve la mise en forme (couleur
    du statut, police, retour à la ligne de l'historique)."""
    lignes = [[cell.value for cell in row] for row in ws.iter_rows(min_row=2)]
    if len(lignes) <= 1:
        return

    lignes.sort(key=lambda r: (parser_date_fr(r[0]), parser_date_fr(r[4])), reverse=True)

    ws.delete_rows(2, ws.max_row - 1)
    for row in lignes:
        ws.append(row)
        r = ws.max_row
        couleur = STATUT_COULEURS.get(row[3], "FFFFFF")
        ws.cell(row=r, column=4).fill = PatternFill("solid", fgColor=couleur)
        for c in range(1, len(HEADERS) + 1):
            ws.cell(row=r, column=c).font = Font(name="Arial", size=10)
        ws.cell(row=r, column=6).alignment = Alignment(wrap_text=True, vertical="top")


def ajouter_lignes(ws, entrees):
    index = construire_index(ws)
    nb_nouvelles, nb_mises_a_jour = 0, 0

    for e in entrees:
        cle = normaliser(e["entreprise"]) + "||" + normaliser(e["poste"])

        if cle in index:
            # Candidature déjà suivie -> on met à jour la ligne existante
            r = index[cle]
            ancien_statut = ws.cell(row=r, column=4).value

            if ancien_statut != e["statut"]:
                ws.cell(row=r, column=4).value = e["statut"]
                couleur = STATUT_COULEURS.get(e["statut"], "FFFFFF")
                ws.cell(row=r, column=4).fill = PatternFill("solid", fgColor=couleur)

            ws.cell(row=r, column=5).value = e["date"]  # dernière mise à jour
            historique = ws.cell(row=r, column=6).value or ""
            ws.cell(row=r, column=6).value = (
                historique + f"\n{e['date']} : {e['statut']}" if historique else f"{e['date']} : {e['statut']}"
            )
            ws.cell(row=r, column=7).value = e["objet"]
            ws.cell(row=r, column=8).value = e["expediteur"]
            ws.cell(row=r, column=9).value = e["lien"]
            ids_existants = ws.cell(row=r, column=10).value or ""
            ws.cell(row=r, column=10).value = (
                ids_existants + "," + e["id"] if ids_existants else e["id"]
            )
            nb_mises_a_jour += 1
        else:
            # Nouvelle candidature -> nouvelle ligne
            row = [
                e["date"], e["entreprise"], e["poste"], e["statut"],
                e["date"], f"{e['date']} : {e['statut']}",
                e["objet"], e["expediteur"], e["lien"], e["id"],
            ]
            ws.append(row)
            r = ws.max_row
            couleur = STATUT_COULEURS.get(e["statut"], "FFFFFF")
            ws.cell(row=r, column=4).fill = PatternFill("solid", fgColor=couleur)
            index[cle] = r
            nb_nouvelles += 1

        for c in range(1, len(HEADERS) + 1):
            ws.cell(row=r, column=c).font = Font(name="Arial", size=10)
        ws.cell(row=r, column=6).alignment = Alignment(wrap_text=True, vertical="top")

    return nb_nouvelles, nb_mises_a_jour


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
def main():
    if "--trier-seulement" in sys.argv:
        # Retrie le fichier existant sans se connecter à Gmail : utile pour
        # remettre de l'ordre dans un fichier déjà rempli avant ce correctif.
        if not os.path.exists(OUTPUT_PATH):
            print(f"Aucun fichier trouvé à {OUTPUT_PATH}, rien à trier.")
            return
        wb = load_workbook(OUTPUT_PATH)
        ws = wb.active
        trier_par_date(ws)
        wb.save(OUTPUT_PATH)
        print(f"Fichier trié par date : {OUTPUT_PATH}")
        return

    print(f"[{datetime.now()}] Démarrage du suivi des candidatures…")

    state = load_state()
    service = get_gmail_service()
    query = build_query(state["last_run"])
    print(f"Requête Gmail : {query}")

    nouveaux, ids_vus = fetch_new_messages(service, query, set(state["processed_ids"]))
    nb_exclus = len(ids_vus) - len(nouveaux)
    print(f"{len(nouveaux)} nouvel(le)(s) email(s) de candidature trouvé(s) "
          f"({nb_exclus} email(s) ignoré(s) car alerte/newsletter).")

    if nouveaux:
        wb, ws = ouvrir_ou_creer_classeur()
        nb_nouvelles, nb_maj = ajouter_lignes(ws, nouveaux)
        trier_par_date(ws)
        wb.save(OUTPUT_PATH)
        print(f"{nb_nouvelles} nouvelle(s) candidature(s) ajoutée(s), "
              f"{nb_maj} statut(s) mis à jour.")
        print(f"Fichier mis à jour : {OUTPUT_PATH}")

    state["last_run"] = datetime.now().isoformat()
    state["processed_ids"] = list(set(state["processed_ids"]) | set(ids_vus))
    save_state(state)

    print(f"[{datetime.now()}] Terminé.")


if __name__ == "__main__":
    main()