# Projet UCPA - Leroy Merlin - Tableau de Bord Formations

Ce projet permet d'automatiser la gestion des inscriptions aux formations UCPA pour Leroy Merlin en liant Google Sheets, Google Forms et Google Apps Script.

---

## Fonctionnement général

Le système repose sur quatre piliers principaux :

### 1. Inscription via Google Forms
L'utilisateur remplit le formulaire d'inscription en choisissant sa formation. Chaque choix de session intègre l'identifiant de la session entre crochets, par exemple : `Gmail Agenda Meet [SES-0001]`.

### 2. Traitement automatique (Google Apps Script)
Lorsqu'une réponse au formulaire est envoyée, le script déclencheur `onSubmit(e)` réalise les actions suivantes :
- Extraction de l'identifiant de session (`SES-XXXX`) et de l'e-mail de l'inscrit.
- Vérification des doublons : si l'adresse e-mail est déjà inscrite à la même session, l'inscription est ignorée.
- Enregistrement de l'inscription dans la feuille `INSCRIPTIONS` (Colonne A : Horodateur, Colonne B : ID Session, Colonne C : E-mail).
- Génération d'une convocation officielle au format PDF à partir d'un modèle Google Docs et sauvegarde dans la GED Google Drive.
- Envoi automatique d'un e-mail de confirmation contenant la convocation en pièce jointe, les liens d'accès (Google Meet) et un lien de désinscription.
- Envoi d'une invitation Google Calendar.
- Mise à jour automatique des choix dans le Google Form pour retirer les sessions complètes.

### 3. Calcul du nombre d'inscrits dans le Tableau de Bord
Dans la feuille `SESSIONS`, le nombre de participants inscrits (colonne `Nb Inscrits`) est calculé dynamiquement :
- La formule compare les identifiants de session (`B:B`) avec les inscriptions enregistrées dans la colonne B de la feuille `INSCRIPTIONS`.
- Elle déduit également les demandes enregistrées dans la feuille `DESINSCRIPTION FORM`.

Formule automatique globale recommandée sur la cellule `L1` (en-tête de la colonne L) :
```excel
=ARRAYFORMULA(IF(B:B<>""; IF(B:B=B1; "Nb Inscrits"; IFNA(COUNTIF(INSCRIPTIONS!B:B; "="&B:B); 0) - IFNA(COUNTIF('DESINSCRIPTION FORM'!G:G; "="&B:B); 0)); ""))
```

### 4. Interface Web (Web App)
Une interface Web HTML/JS servie par `doGet.js` permet aux utilisateurs de consulter la liste des formations disponibles et les places restantes en temps réel. Les données sont conservées dans le cache du script pour assurer un chargement rapide, avec vidage automatique lors de chaque nouvelle inscription.

---

## Structure des fichiers du projet

Le projet est écrit en **TypeScript** et structuré comme suit sous le dossier `src/` :

- `src/config/`
  - `Settings.ts` : Configuration globale, constantes et accès aux feuilles de calcul / plages nommées.
- `src/classes/`
  - `Query.ts` & `Query_utils.ts` : Classes utilitaires pour requêter les données de Sheets de façon robuste.
- `src/handlers/`
  - `FormHandler.ts` : Logique de soumission de formulaires (inscriptions, désinscriptions, gestion de la liste d'attente), acquisition des verrous avec `LockService` et mise à jour des options des formulaires.
  - `triggers.ts` : Fonctions déclencheurs du projet (`onSubmit`, `onEditTrigger`, `autoUpdateFormChoicesTrigger`).
- `src/services/`
  - `AgendaService.ts` : Synchronisation thread-safe avec les événements Google Agenda.
  - `DocGeneratorService.ts` : Génération des convocations à partir du modèle de document Google Docs.
  - `FormService.ts` : Service facilitant la manipulation de formulaires Google Forms.
  - `MailService.ts` : Construction et envoi des notifications par e-mail.
  - `PdfService.ts` : Conversion des fichiers Google Docs en PDF pour envoi en pièce jointe.
- `src/web/`
  - `WebApp.ts` : Contrôleur d'application web (`doGet`) servant la Web App.
  - `CatalogTemplate.html` : Interface web (catalogue des formations disponibles).
  - `MESSAGE.html` : Contenu HTML pour les e-mails de notification.
- `src/utils.ts` : Fonctions utilitaires globales.

---

## Déploiement des modifications

Les fichiers TypeScript du dossier `src/` sont compilés en JavaScript dans le dossier `dist/` avant d'être poussés vers Google Apps Script.

Pour compiler et pousser les modifications :

```bash
npm run push
```
