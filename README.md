# Thrive·Track — Tableau de bord CSSVD (ECAM & UPAS)

Tableau de bord Streamlit du suivi des producteurs de cacao : engagement des ménages,
inspection CSSVD (swollen shoot), promesses de taille, symptômes avec photos, carte des
parcelles et export Excel des producteurs touchés.

Les données proviennent d'un **export KoboToolbox** (.xlsx ou .csv).

---

## 1. Lancer en local

```bash
pip install -r requirements.txt
streamlit run app.py
```

Ouvrir http://localhost:8501.

Pour charger les données :
- soit déposer l'export Kobo dans le dossier `data/` (le plus récent est lu automatiquement) ;
- soit utiliser le bouton **⚙ Paramètres → Export KoboToolbox**.

Sans fichier, l'application affiche des **données de démonstration**.

---

## 2. Déployer sur Streamlit Community Cloud

### a) Mettre le code sur GitHub

Créer un dépôt **privé** sur github.com, puis dans ce dossier :

```bash
git init
git add .
git status          # vérifier : AUCUN fichier .xlsx / .csv / secrets.toml ne doit apparaître
git commit -m "Tableau de bord CSSVD"
git branch -M main
git remote add origin https://github.com/<votre-compte>/<votre-depot>.git
git push -u origin main
```

Fichiers qui doivent être publiés :

| Fichier | Rôle |
|---|---|
| `app.py` | l'application |
| `requirements.txt` | bibliothèques Python (versions testées) |
| `.streamlit/config.toml` | thème et réglages |
| `.streamlit/secrets.toml.example` | modèle pour la clé Kobo (sans la vraie clé) |
| `.gitignore` | empêche la publication des données et des secrets |
| `README.md` | ce guide |

### b) Créer l'application

1. Aller sur **https://share.streamlit.io** et se connecter avec GitHub.
2. **Create app → Deploy a public app from GitHub**.
3. Renseigner :
   - **Repository** : votre dépôt
   - **Branch** : `main`
   - **Main file path** : `app.py`
4. **Advanced settings** :
   - **Python version** : `3.12`
   - **Secrets** : coller (facultatif, pour afficher les photos des symptômes)
     ```toml
     KOBO_TOKEN = "votre_cle_api_kobo"
     ```
     La clé se trouve sur Kobo : *Paramètres du compte → Sécurité → Clé API*.
5. **Deploy**. L'adresse ressemble à `https://<nom>.streamlit.app`.

### c) Restreindre l'accès

L'application contient des **données personnelles** (noms des producteurs, positions GPS
des parcelles). Avec un dépôt **privé**, l'application est privée par défaut :
*App → Settings → Sharing* → inviter les personnes autorisées par e-mail.

---

## 3. Publier les données pour tout le monde (administrateur)

L'administrateur charge l'export Kobo ; **tous les utilisateurs voient la mise à jour
automatiquement** (vérification toutes les 30 secondes, sans rafraîchir la page).

### Configuration (une seule fois)

1. Créer un **second dépôt GitHub, PRIVÉ**, pour les données, par ex.
   `ouehi-mario/dashboard-thrive-data` (cocher *Add a README* pour qu'il ne soit pas vide).
2. Créer un jeton : GitHub → *Settings → Developer settings → Personal access tokens →
   Fine-grained tokens → Generate new token*
   - *Repository access* : **Only select repositories** → le dépôt de données
   - *Permissions → Contents* : **Read and write**
3. Dans Streamlit Cloud → *App → Settings → Secrets* :
   ```toml
   ADMIN_PASSWORD = "un_mot_de_passe_solide"
   GITHUB_TOKEN   = "github_pat_..."
   DATA_REPO      = "ouehi-mario/dashboard-thrive-data"
   KOBO_TOKEN     = "..."          # facultatif (photos)
   ```

Sans `GITHUB_TOKEN` / `DATA_REPO`, la publication fonctionne aussi mais les données sont
gardées sur le serveur uniquement : elles disparaissent à chaque redémarrage de l'application.

### Publier

1. **⚙ Paramètres → Espace administrateur** → saisir le mot de passe → *Se connecter*.
2. Déposer un ou plusieurs fichiers (.xlsx / .csv). Chaque fichier est **reconnu automatiquement** :

   | Fichier | Coopérative | Reconnu grâce à la colonne |
   |---|---|---|
   | Export KoboToolbox | ECAM | `Producteur/…`, `_submission_time` |
   | `thrive_engagements_menages_…` | UPAS | `Date Engagement` |
   | `thrive_inspections_cssvd_…` | UPAS | `Date Observation` |
   | `thrive_promesses_taille_…` | UPAS | `Date Promesse` |

3. Vérifier la ligne « ✓ fichier → type · N lignes », puis **📤 Publier … pour tous**.

On peut republier un seul fichier (par ex. seulement les promesses UPAS) : les autres restent inchangés.

### ECAM : fusion avec les données déjà publiées

Le bouton **« ECAM : ajouter aux données déjà publiées (fusion) »** est activé par défaut :

- les fiches déjà publiées sont **conservées**, même si elles ne sont plus dans le nouvel export
  (export partiel, fiches supprimées sur Kobo…) ;
- les nouvelles fiches sont **ajoutées** ;
- une fiche présente dans les deux (même identifiant Kobo `_uuid`) est **mise à jour** avec la version du nouvel export.

Avant de publier, l'écran affiche le bilan : « N déjà publiées + N dans le fichier → N nouvelles,
N mises à jour, N conservées · N au total ». Le nouvel export doit avoir le même format que le précédent
(même choix *libellés / noms XML* dans Kobo), sinon la fusion est refusée.
Désactivez le bouton pour **remplacer** complètement les données. Les fichiers UPAS sont toujours remplacés.

**💾 Tout l'ECAM (.xlsx)** télécharge l'ensemble des fiches fusionnées (sauvegarde).

### Archive des photos

À chaque publication ECAM, les photos de l'export (symptômes, documents, photo du producteur)
sont téléchargées depuis Kobo et **enregistrées dans le dépôt de données** (`photos/`).
Elles restent visibles même si la soumission est supprimée sur Kobo. Nécessite `KOBO_TOKEN`.

Le bouton **🖼 Archiver N photo(s)** relance l'archivage des photos manquantes (par ex. après un échec réseau).
Compter environ 200 Ko par photo : prévoir de surveiller la taille du dépôt (GitHub recommande moins de 1 à 5 Go).

## Coopératives

Le menu en haut à gauche permet de passer de **ECAM** à **UPAS**.

| | ECAM | UPAS |
|---|---|---|
| Source | 1 export KoboToolbox | 3 fichiers liés par le code producteur |
| Pages | Tableau de bord, Producteurs, Coachs, Symptômes, Carte, Données | Tableau de bord, Suivi producteurs, Inspections, Promesses, Carte, Données |
| Objectif | 1 000 producteurs | 1 000 producteurs |

Les numéros de téléphone et les noms des conjointes (UPAS) ne sont jamais affichés.

Chaque publication est enregistrée comme une version dans le dépôt de données
(historique consultable et restaurable sur GitHub).

## 3 bis. Utilisation

- **Mode clair / sombre** : automatique, selon le réglage du téléphone ou du PC (application et rapport hors ligne).
- Les utilisateurs n'ont rien à charger : ils voient la dernière version publiée
  (date affichée dans ⚙ Paramètres).
- **Objectif** : 1 000 producteurs par défaut (modifiable dans ⚙ Paramètres).
- **Photos des symptômes** : cliquer sur un symptôme. Si les photos ne s'affichent pas,
  vérifier le secret `KOBO_TOKEN` (ou le saisir dans ⚙ Paramètres).
- **Export** : page *Symptômes → Producteurs avec swollen shoot → Exporter Excel / CSV*.

## 4. Mettre à jour l'application

Toute modification poussée sur GitHub est redéployée automatiquement :

```bash
git add app.py
git commit -m "Description de la modification"
git push
```

## 5. Bon à savoir

- Sur l'offre gratuite, l'application se met en veille après quelques jours sans visite ;
  la première ouverture prend alors ~30 secondes.
- Ne jamais écrire la clé Kobo dans `app.py` ni la publier sur GitHub.
- Ne jamais pousser l'export Kobo sur GitHub (le `.gitignore` l'empêche par défaut).
