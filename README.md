# Template TypeScript sans front-end Numericoach pour Google Apps Script

Ce template permet de créer un projet Google Apps Script sans partie HTML en TypeScript.

## Installation

Installer les dépendances

```bash
npm i
```

Ou [Bun](https://bun.sh/) (beaucoup plus rapide)

```bash
bun i
```

## Utilisation

Il vous faudra changer dans .clasp.json le scriptId que vous pouvez retrouver dans votre projet Google Apps Script en ligne dans les paramètres du projet.

Une fois que tout est prêt, vous pouvez utiliser la commande suivante pour déployer votre projet.

```bash
clasp push -w
```

## Configurer les pipelines CI/CD

Pour configurer les pipelines CI/CD (Production et QA), il vous suffira d'ajouter l'identifiant du script, et d'aller créer les secrets dans GitHub avec les JSON d'authentification au compte qui sont pour la plupart disponibles dans Dashlane.

### Production

Pour la production, il vous faudra ajouter le `script-id` dans le fichier [main.yml](/.github/workflows/main.yml) à la ligne 13.

Ensuite, il vous faut aller sur le repository du projet dans GitHub, puis dans `Settings` > `Secrets and variables` > `Actions` > `New repository secret` et ajouter le secret suivant :

-   `CLASPRC_JSON_PROD` : Le JSON d'authentification disponible sur Dashlane, pour le site GitHub avec la mention `CLASPRC_JSON` et copiez le nom d'utilisateur secondaire.

### QA

Pour la QA, c'est la même chose, il vous faudra ajouter le `script-id` dans le fichier [qa.yml](/.github/workflows/qa.yml) à la ligne 13.

Ensuite, il vous faut aller sur le repository du projet dans GitHub, puis dans `Settings` > `Secrets and variables` > `Actions` > `New repository secret` et ajouter le secret suivant :

-   `CLASPRC_JSON_QA` : Le JSON d'authentification disponible sur Dashlane, pour le site GitHub avec la mention `CLASPRC_JSON` et copiez le nom d'utilisateur secondaire.
