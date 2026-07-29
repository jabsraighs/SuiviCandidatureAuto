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
        "avons le plaisir de vous proposer",
        "avons le plaisir de vous annoncer votre",
        "avons le plaisir de vous compter parmi",
        "avons le plaisir de vous accueillir",
        "proposition d'embauche",
        "offre d'embauche",
        "bienvenue dans l'équipe",
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

# ---------------------------------------------------------------------------
# EXCLUSION : évite les faux positifs qui ne sont pas de vraies candidatures
# (alertes emploi, newsletters, notifications de compte...) mais qui
# contiennent quand même un mot comme "candidature" ou "recrutement".
# ---------------------------------------------------------------------------

# Domaines d'expéditeur qui n'envoient QUE des alertes/newsletters d'offres
# d'emploi, jamais de vrai suivi de candidature -> exclusion totale du domaine.
EXPEDITEURS_DOMAINES_A_EXCLURE = [
    "jobleads.com",
    "talent.com",
    "tekkit.io",
]

# Adresses précises à exclure (utile quand un même domaine envoie à la fois
# de vraies confirmations de candidature ET des newsletters, ex. HelloWork :
# on exclut seulement l'adresse newsletter, pas les confirmations réelles).
EXPEDITEURS_EXACTS_A_EXCLURE = [
    "newsletter@emails.hellowork.com",
]

# Si l'objet contient une de ces tournures, l'email est ignoré quel que soit
# l'expéditeur : ce sont des formulations typiques d'alertes/newsletters ou
# de notifications de compte, jamais d'un vrai suivi de candidature.
MOTS_CLES_EXCLUSION_OBJET = [
    "nouvelles offres",
    "nouvelle offre",
    "correspondant à votre recherche",
    "recherche sauvegardée",
    "à ne pas manquer",
    "à découvrir aujourd'hui",
    "à pourvoir cette semaine",
    "offres d'emploi pour vous",
    "emplois pour vous",
    "mot de passe à usage unique",
    "code de vérification",
    "vérifiez votre adresse",
    "confirmez votre adresse",
    "confirmez votre e-mail",
]