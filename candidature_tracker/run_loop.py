#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
run_loop.py
===========
Lance gmail_candidature_tracker.py en boucle, toutes les N heures, pour que
le conteneur Docker fasse le travail d'une tâche planifiée sans dépendre du
Planificateur de tâches Windows.

Variable d'environnement :
  INTERVAL_HOURS  -> intervalle entre deux passages (défaut : 3)
"""

import os
import time
import traceback
from datetime import datetime

import gmail_candidature_tracker as tracker

INTERVAL_HOURS = float(os.environ.get("INTERVAL_HOURS", "3"))

if __name__ == "__main__":
    print(f"Suivi des candidatures lancé — passage toutes les {INTERVAL_HOURS} h.")
    while True:
        try:
            tracker.main()
        except Exception:
            print(f"[{datetime.now()}] Erreur pendant l'exécution :")
            traceback.print_exc()

        print(f"[{datetime.now()}] Prochain passage dans {INTERVAL_HOURS} h.\n")
        time.sleep(INTERVAL_HOURS * 3600)