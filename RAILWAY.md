# Déployer Webora sur Railway

## 1. Mettre le projet sur GitHub

Créez un nouveau dépôt GitHub, décompressez `webora-docker.zip`, puis placez
le contenu du dossier `webora` à la racine du dépôt.

Vous pouvez également utiliser Git depuis le dossier du projet :

```bash
git init
git add .
git commit -m "Deploy Webora"
git branch -M main
git remote add origin https://github.com/VOTRE-COMPTE/webora.git
git push -u origin main
```

## 2. Créer le service Railway

1. Connectez-vous sur <https://railway.app/>.
2. Choisissez **New Project** puis **Deploy from GitHub repo**.
3. Autorisez Railway à accéder au dépôt Webora.
4. Sélectionnez le dépôt.

Railway détecte `railway.toml` et utilise automatiquement le `Dockerfile`.
La variable `PORT` fournie par Railway est prise en charge par Nginx.

## 3. Générer l'adresse publique

Dans le service Railway :

1. Ouvrez **Settings**.
2. Accédez à la section **Networking**.
3. Cliquez sur **Generate Domain**.

Railway fournit alors une adresse HTTPS publique.

## 4. Déployer les mises à jour

Chaque nouveau `git push` sur la branche connectée déclenche une nouvelle
construction et un nouveau déploiement.

## Données de l'application

Railway héberge ici uniquement l'interface Webora. Les contacts et imports
Excel restent dans le `localStorage` de chaque navigateur. Aucune base de
données Railway n'est nécessaire pour cette version, mais les données ne sont
pas synchronisées entre appareils.
