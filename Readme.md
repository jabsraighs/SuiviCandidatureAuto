# Suivi automatique des candidatures (Gmail → Excel)

Ce script se connecte à ton Gmail, repère les emails liés à tes candidatures
(accusé de réception, entretien, refus, offre…) et alimente un fichier Excel
`candidatures.xlsx` avec une ligne par email détecté, sans jamais créer de
doublon. Il est fait pour tourner **tout seul, régulièrement**, via une tâche
planifiée.

## 1. Créer les identifiants Google (une seule fois — nécessaire dans tous les cas)

1. Va sur https://console.cloud.google.com/ et crée un projet (ou utilise un
   projet existant).
2. Menu **API et services → Bibliothèque** → cherche **Gmail API** → clique
   **Activer**.
3. Menu **API et services → Écran de consentement OAuth** :
   - Type d'utilisateur : **Externe**
   - Renseigne un nom d'app (ex. "Suivi Candidatures"), ton email, sauvegarde.
   - Dans **Utilisateurs test**, ajoute ta propre adresse Gmail.
4. Menu **API et services → Identifiants → Créer des identifiants → ID
   client OAuth** :
   - Type d'application : **Application de bureau**
   - Une fois créé, clique **Télécharger le JSON**.
5. Renomme ce fichier téléchargé en `credentials.json`.

> Ce fichier ne quitte jamais ton ordinateur : l'authentification se fait en
> local, Anthropic/Claude n'a accès à rien de tout ça.

---

## Méthode A — Docker (recommandée si tu n'as pas Python installé)

Il te faut juste **Docker Desktop** (https://www.docker.com/products/docker-desktop/),
installé et lancé sur Windows.

### A.1 Organiser les fichiers

Crée un dossier, ex. `C:\Users\TonNom\gmail_tracker`, avec dedans :
```
gmail_tracker/
├── Dockerfile
├── docker-compose.yml
├── requirements.txt
├── gmail_candidature_tracker.py
├── config.py
├── run_loop.py
└── data/
    └── credentials.json      ← celui téléchargé à l'étape 1
```
Crée le sous-dossier `data` et mets-y ton `credentials.json`.

### A.2 Construire l'image

Ouvre un terminal (PowerShell) dans le dossier `gmail_tracker` :

```powershell
docker compose build
```

### A.3 Première autorisation Gmail (une seule fois, manuellement)

```powershell
docker compose run --rm -p 8080:8080 candidatures python gmail_candidature_tracker.py
```

Le terminal affiche un lien `https://accounts.google.com/...` : **copie-le
dans ton navigateur Windows**, connecte-toi et autorise l'accès **lecture
seule** à Gmail. La page te renverra vers `localhost:8080`, ce qui est
intercepté par le conteneur (grâce au `-p 8080:8080`) → le terminal affiche
que l'autorisation a réussi, et `data/token.json` est créé.

### A.4 Lancer en continu (tourne tout seul, remplace le Planificateur de tâches)

```powershell
docker compose up -d
```

C'est tout : le conteneur tourne en arrière-plan et repasse automatiquement
toutes les **3 heures** (réglable, voir plus bas), même après un redémarrage
de Windows (grâce à `restart: unless-stopped` dans `docker-compose.yml`,
tant que Docker Desktop est configuré pour démarrer avec Windows).

Commandes utiles :
```powershell
docker compose logs -f          # voir ce que fait le script en direct
docker compose stop             # mettre en pause
docker compose up -d            # relancer
docker compose down             # arrêter et supprimer le conteneur (les données dans ./data restent)
```

Pour changer l'intervalle (ex. toutes les 6h), édite `INTERVAL_HOURS` dans
`docker-compose.yml`, puis `docker compose up -d --build`.

Le fichier `candidatures.xlsx` apparaît dans `gmail_tracker\data\` — ouvrable
directement dans Excel depuis l'explorateur Windows, même pendant que le
conteneur tourne.

---

## Méthode B — Python installé directement sur Windows

Si tu préfères ne pas utiliser Docker, installe Python
(https://www.python.org/downloads/, coche **"Add python.exe to PATH"** à
l'installation), puis :

### B.1 Installer les dépendances

```powershell
pip install -r requirements.txt
```

### B.2 Placer `credentials.json`

Dans le même dossier que `gmail_candidature_tracker.py`.

### B.3 Premier lancement (manuel, pour autoriser l'accès)

```powershell
python gmail_candidature_tracker.py
```

Une page de navigateur s'ouvre pour autoriser l'accès **lecture seule** à
Gmail. Accepte : `token.json` est créé et réutilisé pour tous les
lancements suivants.

### B.4 Planifier avec le Planificateur de tâches Windows

**⚠️ Fais d'abord B.3 manuellement** — le Planificateur tourne en
arrière-plan et ne peut pas ouvrir de navigateur.

**Option graphique**
1. Démarrer → **"Planificateur de tâches"**.
2. **Créer une tâche** (pas "tâche de base").
3. Onglet **Général** : nom `Suivi candidatures Gmail`, coche **Exécuter
   que l'utilisateur soit connecté ou non** si besoin.
4. Onglet **Déclencheurs → Nouveau** : Quotidien, **Répéter la tâche toutes
   les** `3 heures`, pendant une durée indéterminée.
5. Onglet **Actions → Nouveau** : Démarrer un programme →
   - Programme : chemin complet vers `pythonw.exe` (trouve-le avec
     `where python` puis remplace `python.exe` par `pythonw.exe`)
   - Arguments : `gmail_candidature_tracker.py`
   - Démarrer dans : le dossier `gmail_tracker`
6. Onglet **Conditions** : décoche l'option secteur si tu es sur portable.
7. **OK**, entre ton mot de passe de session.

**Option ligne de commande** (invite de commande en administrateur) :
```cmd
schtasks /create /tn "Suivi candidatures Gmail" /tr "\"C:\chemin\vers\pythonw.exe\" \"C:\Users\TonNom\gmail_tracker\gmail_candidature_tracker.py\"" /sc hourly /mo 3 /st 09:00
```
Vérifier : `schtasks /query /tn "Suivi candidatures Gmail"`
Tester : `schtasks /run /tn "Suivi candidatures Gmail"`
Supprimer : `schtasks /delete /tn "Suivi candidatures Gmail" /f`

---

## Personnaliser la détection

Ouvre `config.py` :
- `MOTS_CLES_RECHERCHE` : mots utilisés pour repérer les emails candidats.
- `CLASSIFICATION` : mots-clés qui déterminent le statut (Reçue, Entretien,
  Refus, Offre). Ajoute tes propres tournures si tu vois des emails mal
  classés ou non détectés.
- `JOURS_HISTORIQUE_INITIAL` : nombre de jours d'historique scannés au tout
  premier lancement (90 par défaut).

Avec Docker, après modification de `config.py`, reconstruis l'image :
```powershell
docker compose up -d --build
```

## Fichiers générés (ne pas supprimer)

Avec Docker, ils se trouvent dans `data/` ; en Python direct, dans le
dossier du script.

| Fichier | Rôle |
|---|---|
| `token.json` | Jeton d'accès Gmail (généré après la 1ère autorisation) |
| `state.json` | Mémorise la date du dernier passage + emails déjà traités |
| `candidatures.xlsx` | Le suivi de tes candidatures |

## Colonnes du fichier Excel

Le script suit **une ligne par candidature** (pas une ligne par email) : si
plusieurs emails arrivent pour la même entreprise/poste (accusé de
réception, puis entretien, puis refus...), la ligne existante est **mise à
jour automatiquement** au lieu d'être dupliquée, avec un historique complet.

| Colonne | Contenu |
|---|---|
| Date candidature | Date du tout premier email détecté pour cette candidature |
| Entreprise | Déduite du sujet ("... chez X") ou du nom/domaine de l'expéditeur |
| Poste | Extrait de l'objet quand c'est détectable |
| Statut actuel | Le statut le plus récent : Candidature envoyée / Reçue-en cours / Entretien / Refus / Offre |
| Dernière mise à jour | Date du dernier email reçu pour cette candidature |
| Historique statuts | Toutes les étapes horodatées (ex. "01/07 : Reçue — 15/07 : Entretien — 25/07 : Refus") |
| Objet du dernier email | Objet du dernier email pris en compte |
| Expéditeur | Adresse email source du dernier email |
| Lien Gmail (dernier) | Lien direct vers le dernier email |
| ID messages | Identifiants techniques de tous les emails liés (évite les doublons) |

**Comment est reconnue "la même candidature" ?** Par entreprise + poste
(normalisés, insensibles à la casse/accents). Limite à connaître : si le
poste n'est pas détecté pour deux candidatures différentes chez la même
entreprise, elles pourraient être fusionnées à tort — dans ce cas, sépare-les
à la main dans Excel (ça n'affecte pas les prochains passages du script,
qui continueront à les traiter comme des lignes distinctes une fois séparées
tant que les postes diffèrent).

L'extraction d'entreprise/poste est faite par heuristique (regex) : elle ne
sera pas parfaite à 100 %, tu peux corriger à la main dans Excel.


supprimer le fichier token.json
docker rm -f $(docker ps -aq)
docker compose up --build -d
docker compose exec candidatures python gmail_candidature_tracker.py --trier-seulement
docker compose run --rm -p 8080:8080 candidatures python gmail_candidature_tracker.py