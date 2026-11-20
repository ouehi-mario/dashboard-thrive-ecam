# Ecam·Track — Tableau de bord CSSVD

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

## 3. Utilisation en ligne

- Charger l'export Kobo via **⚙ Paramètres**. Le fichier reste dans la session du navigateur ;
  il n'est **pas** enregistré sur le serveur et doit être rechargé après un rafraîchissement.
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
