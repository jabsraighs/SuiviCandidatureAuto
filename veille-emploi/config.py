# -*- coding: utf-8 -*-
"""
config.py — Personnalise ici les recherches sans toucher au script principal.
"""

# Chaque recherche = (mot-clé, ville). Ajoute/retire des lignes librement.
RECHERCHES = [
    ("DevOps junior", "Île-de-France"),
    ("Administrateur Linux junior", "Île-de-France"),
    ("Développeur PHP Symfony", "Paris"),
    ("DevSecOps junior", "Île-de-France"),
    ("Ingénieur déploiement", "Île-de-France"),
]


 
# ---------------------------------------------------------------------------
# ALERTES EMAIL (source principale) — nécessite d'avoir créé les alertes sur
# Indeed / HelloWork au préalable, envoyées vers la boîte Gmail surveillée.
# ---------------------------------------------------------------------------
 
# Domaines d'expéditeur des emails d'alerte à surveiller.
EXPEDITEURS_ALERTES = [
    "indeed.com",
    "hellowork.com",
    "emails.hellowork.com",
]
 
# Lors du tout premier lancement, on regarde les X derniers jours d'emails.
JOURS_HISTORIQUE_INITIAL_EMAIL = 7
 
# Fragments de texte à ignorer dans les liens extraits des emails (ce sont
# des liens de navigation — désabonnement, préférences — pas des offres).
LIENS_A_IGNORER = [
    "se désabonner",
    "désabonner",
    "gérer mes alertes",
    "gérer les alertes",
    "voir toutes les offres",
    "modifier mes préférences",
    "unsubscribe",
    "voir plus d'offres",
    "créer une alerte",
    "télécharger l'app",
    "politique de confidentialité",
    "mentions légales",
]