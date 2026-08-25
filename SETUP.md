# Guide d'Installation et d'Architecture - Inscriptions Leroy Merlin

Ce document explique le fonctionnement du projet, l'architecture simplifiée, la sécurisation contre les accès simultanés via `LockService`, et la procédure d'installation sur un nouveau Google Sheet.

---

## 1. Architecture Simplifiée

Le système est conçu autour de 4 composants essentiels :
1. **Google Sheets :** Pilote les sessions de formation, enregistre les inscriptions et stocke les paramètres (`PARAMETRES`).
2. **Google Forms :** Permet aux directeurs de s'inscrire ou de se désinscrire avec sélection dynamique de la session.
3. **Google Agenda :** Crée automatiquement 1 événement unique par session et y inscrit le participant.
4. **Google Apps Script (TypeScript) :** Exécute l'automatisation avec protection de concurrence.

---

## 2. Protection contre la Concurrence (`LockService`)

Pour éviter qu'en cas d'inscriptions simultanées deux personnes ne dépassent la capacité d'une session ou qu'il y ait doublon d'événement Google Agenda, le script utilise **`LockService`** :

- **Lors de l'inscription (`onSubmit`) :** Le script acquiert un verrou exclusif (`LockService.getScriptLock().tryLock(10000)`). Il vérifie les places disponibles et l'anti-doublon *sous le verrou* avant d'écrire l'inscription.
- **Lors de la création d'événement (`getOrCreateSessionEventId`) :** La recherche/création de l'événement Agenda est verrouillée pour garantir **1 Session = 1 Événement Agenda unique**.
- **Découplage Agenda / E-mail :** L'ajout de l'invité dans l'agenda s'exécute en priorité absolue. Tout échec éventuel d'envoi d'e-mail est capturé dans un bloc `try/catch` isolé afin de ne pas bloquer l'inscription.

---

## 3. Procédure d'Installation sur un Nouveau Classeur

### Étape 1 : Cloner le projet avec Clasp
1. Récupérez l'ID du nouveau script Google Apps Script depuis votre Google Sheet (*Extension > Apps Script*).
2. Dans le fichier `.clasp.json`, remplacez le `scriptId` :
   ```json
   {"scriptId": "VOTRE_NOUVEAU_SCRIPT_ID", "rootDir": "./src/"}
   ```
3. Envoyez le code TypeScript vers Google Apps Script :
   ```bash
   npx clasp push
   ```

### Étape 2 : Installer les Déclencheurs (Triggers)
Ouvrez le Google Sheet lié, cliquez sur le menu sur mesure **`NUMERICOACH > Installer les déclencheurs (Triggers)`**.
Ceci exécutera la fonction `setupTriggers()` qui installe le déclencheur `onFormSubmit` automatiquement.

### Étape 3 : Récupérer le champ réél Google Forms (`entry.XXXXX`)
1. Ouvrez votre Google Form en mode pré-remplissage (*Obtenir le lien pré-rempli*).
2. Saisissez un texte dans la question de sélection de session et cliquez sur *Obtenir le lien*.
3. Copiez l'URL générée et repérez l'ID de champ réel, par exemple `entry.2116080188`.
4. Renseignez cet ID dans la plage nommée `PARAMETRE_ENTRY_SESSION` de l'onglet `PARAMETRES`.

---

## 4. Structure des Fichiers TypeScript

- [src/00 - SETTINGS.ts](file:///c:/Project/Leroy-Merlin-Project/src/00%20-%20SETTINGS.ts) : Initialisation du menu NUMERICOACH et gestion des plages nommées.
- [src/05 - INSCRIPTIONS.ts](file:///c:/Project/Leroy-Merlin-Project/src/05%20-%20INSCRIPTIONS.ts) : Gestion des soumissions `onSubmit`, contrôle `LockService`, mise à jour des choix `updateFormChoices`.
- [src/08 - AGENDA.ts](file:///c:/Project/Leroy-Merlin-Project/src/08%20-%20AGENDA.ts) : Synchronisation thread-safe des événements Google Agenda.
- [src/triggers.ts](file:///c:/Project/Leroy-Merlin-Project/src/triggers.ts) : Installation automatisée des déclencheurs Apps Script.
