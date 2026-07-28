# -*- coding: utf-8 -*-
"""
config.py — Personnalise ici la détection sans toucher au script principal.
"""

# Nom du fichier Excel généré (dans le même dossier que le script)
OUTPUT_FILENAME = "candidatures.xlsx"

# Lors du tout premier lancement, on regarde les X derniers jours de mails
JOURS_HISTORIQUE_INITIAL = 90

# ---------------------------------------------------------------------------
# Mots-clés utilisés pour la RECHERCHE Gmail (pour ne récupérer que les
# emails potentiellement liés à une candidature). Reste large : la
# classification fine se fait ensuite dans CLASSIFICATION ci-dessous.
# ---------------------------------------------------------------------------
MOTS_CLES_RECHERCHE = [
    "candidature",
    "postulé",
    "postulation",
    "votre candidature",
    "recrutement",
    "entretien",
    "process de recrutement",
    "accusé de réception",
    "offre d'emploi",
    "nous avons bien reçu",
]

# Plateformes / intermédiaires qui apparaissent comme "expéditeur" mais qui
# ne sont PAS l'entreprise qui recrute (HelloWork, France Travail…). Quand
# l'expéditeur correspond à l'une d'elles, on essaie d'abord d'extraire le
# vrai nom de l'entreprise depuis le sujet (ex. "... chez Safran"), sinon on
# garde le nom de la plateforme faute de mieux.
PLATEFORMES_A_IGNORER = [
    "hellowork",
    "indeed",
    "linkedin",
    "welcome to the jungle",
    "wttj",
    "apec",
    "monster",
    "pôle emploi",
    "pole emploi",
    "france travail",
    "glassdoor",
    "jobteaser",
]

# ---------------------------------------------------------------------------
# Classification du statut : la première catégorie dont un mot-clé est
# trouvé dans l'objet + l'aperçu de l'email est retenue.
# L'ORDRE compte : on vérifie "Offre" et "Refus" avant "Entretien",
# eux-mêmes avant "Reçue", pour éviter les faux positifs.
# ---------------------------------------------------------------------------
CLASSIFICATION = {
    "Offre / Accepté": [
        "félicitations",
        "avons le plaisir de vous",
        "proposition d'embauche",
        "offre d'embauche",
        "bienvenue dans l'équipe",
        "bienvenue chez",
        "avez rejoint notre équipe",
        "contrat de travail",
    ],
    "Refus": [
        "ne pouvons donner suite",
        "ne pouvons pas donner suite",
        "n'a pas été retenue",
        "n'avons pas retenu",
        "ne donnerons pas suite",
        "poste a été pourvu",
        "malheureusement",
        "à notre plus grand regret",
        "sommes au regret",
        "nous sommes au regret",
        "autre candidat",
    ],
    "Entretien": [
        "entretien",
        "entrevue",
        "convocation",
        "échange téléphonique",
        "rendez-vous",
        "disponibilités",
        "visioconférence",
        "call découverte",
    ],
    "Reçue / en cours": [
        "avons bien reçu",
        "a bien été reçue",
        "bien reçu votre candidature",
        "accusé de réception",
        "confirmation de candidature",
        "candidature bien reçue",
        "en cours d'examen",
        "étudions votre profil",
        "bonne réception",
        "accusons",
        "étude de votre candidature",
        "étudier votre candidature",
        "prendra contact avec vous",
        "va l'étudier",
        "pour compléter votre candidature",
        "du nouveau sur votre candidature",
    ],
}