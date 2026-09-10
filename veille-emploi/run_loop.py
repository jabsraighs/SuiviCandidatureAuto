# -*- coding: utf-8 -*-
"""
run_loop.py — Exécute veille_emploi.py en boucle, toutes les INTERVAL_HOURS heures.
"""

import os
import subprocess
import sys
import time
from datetime import datetime

INTERVAL_HOURS = float(os.environ.get("INTERVAL_HOURS", "3"))


def log(msg):
    print(f"[{datetime.now().isoformat(sep=' ', timespec='seconds')}] {msg}", flush=True)


def main():
    log(f"Démarrage de la veille d'offres (intervalle : {INTERVAL_HOURS}h).")
    while True:
        try:
            subprocess.run([sys.executable, "veille_emploi.py"], check=True)
        except subprocess.CalledProcessError as e:
            log(f"Erreur lors de l'exécution de veille_emploi.py : {e}")

        log(f"Prochain passage dans {INTERVAL_HOURS}h.")
        time.sleep(INTERVAL_HOURS * 3600)


if __name__ == "__main__":
    main()