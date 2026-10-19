# Héberger Webora avec Docker

L'application est compilée avec Node.js et pnpm, puis servie par Nginx dans
une image de production légère.

## Lancement rapide

Depuis la racine du projet :

```bash
docker compose up -d --build
```

Webora est ensuite disponible sur :

```text
http://localhost:8080
```

Pour utiliser un autre port :

```bash
WEBORA_PORT=3000 docker compose up -d --build
```

## Commandes utiles

```bash
# Afficher l'état et la santé du conteneur
docker compose ps

# Afficher les journaux
docker compose logs -f webora

# Reconstruire après une mise à jour
docker compose up -d --build

# Arrêter l'application
docker compose down
```

## Déploiement sur un serveur

1. Installer Docker et le plugin Docker Compose sur le serveur.
2. Copier ou cloner le projet.
3. Exécuter `docker compose up -d --build`.
4. Autoriser le port choisi dans le pare-feu, ou placer un reverse proxy
   HTTPS devant `http://127.0.0.1:8080`.

Le conteneur expose Nginx sur le port `80`. Le port public `8080` est
configurable avec la variable `WEBORA_PORT`.

Pour un déploiement sur Railway, consultez également `RAILWAY.md`. L'image
écoute automatiquement le port dynamique fourni par la plateforme.

## Import Excel et persistance

Les fichiers `.xlsx`, `.xls` et `.csv` sont lus directement dans le navigateur.
Ils ne sont pas envoyés au conteneur.

Les comptes, contacts et imports restent enregistrés dans le `localStorage` du
navigateur. Les données survivent donc au redémarrage du conteneur, mais elles
ne sont pas partagées entre plusieurs navigateurs ou appareils. Pour une
utilisation multi-utilisateur réelle, il faudra ajouter une API, une base de
données et une authentification côté serveur.
