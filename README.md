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

- `src/00 - SETTINGS.js` : Initialisation des accès aux feuilles, plages nommées et gestionnaires de cache.
- `src/05 - INSCRIPTIONS.js` : Logique de réception des formulaires, enregistrement des inscriptions, génération des convocations PDF et envois des e-mails.
- `src/08 - AGENDA.js` : Synchronisation des sessions avec Google Calendar.
- `src/10 - ARCHIVES.js` : Fonctions d'archivage.
- `src/ALERTE.js` : Gestion des alertes e-mail (ex: nombre d'inscrits insuffisant à 3 jours de l'échéance).
- `src/doGet.js` : Contrôleur de la Web App fournissant l'interface utilisateur.
- `src/web.html` et `src/MESSAGE.html` : Templates HTML de la Web App et des e-mails.
- `UCPA - LEROY MERLIN - TABLEAU DE BORD - FORMATIONS (1).xlsx` : Classeur de suivi.

---

## Déploiement des modifications

Pour pousser les mises à jour des fichiers du dossier `src/` vers l'instance Google Apps Script :

```bash
npx clasp push
```
