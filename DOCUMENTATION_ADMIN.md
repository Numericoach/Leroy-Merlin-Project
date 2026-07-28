# 📘 Documentation Administrateur - Système de Gestion des Inscriptions Leroy Merlin x Numericoach

Ce document est la notice technique et administrative complète du système d'automatisation des inscriptions aux formations Leroy Merlin. Il est spécialement structuré pour être copié-collé directement dans un Google Doc.

---

## 📋 Table des Matières
1. [Présentation Générale & Architecture du Système](#1-présentation-générale--architecture-du-système)
2. [Guide d'Installation & Configuration Initiale (Setup Pas à Pas)](#2-guide-dinstallation--configuration-initiale-setup-pas-à-pas)
3. [Description Détaillée des Onglets Google Sheets](#3-description-détaillée-des-onglets-google-sheets)
4. [Référentiel des Paramètres Système (Onglet PARAMETRES)](#4-référentiel-des-paramètres-système-onglet-parametres)
5. [Fonctionnement des Automatismes & Services](#5-fonctionnement-des-automatismes--services)
6. [Guide d'Utilisation du Menu Administrateur NUMERICOACH](#6-guide-dutilisation-du-menu-administrateur-numericoach)
7. [Dépannage & Résolution des Incidents Courants](#7-dépannage--résolution-des-incidents-courants)

---

## 1. Présentation Générale & Architecture du Système

Le système orchestre l'ensemble du cycle de vie d'une formation : de la publication des sessions sur le catalogue WebApp jusqu'à la convocation des participants et la synchronisation de l'agenda Google.

```
                    ┌─────────────────────────┐
                    │ Catalogue WebApp / Form │
                    └────────────┬────────────┘
                                 │ Inscription
                                 ▼
                    ┌─────────────────────────┐
                    │  Feuille Google Sheets  │
                    └────────────┬────────────┘
                                 │ Déclencheur onSubmit
         ┌───────────────────────┼───────────────────────┐
         ▼                       ▼                       ▼
┌──────────────────┐    ┌──────────────────┐    ┌──────────────────┐
│  Google Agenda   │    │ Convocation PDF  │    │ Email Confirmation│
│  (Lien Meet +    │    │ (Génération Drive│    │  (Avec PDF & Lien │
│  Liste Inscrits) │    │   depuis Modèle) │    │  Désinscription)  │
└──────────────────┘    └──────────────────┘    └──────────────────┘
```

---

## 2. Guide d'Installation & Configuration Initiale (Setup Pas à Pas)

Pour initialiser le système sur un nouveau compte ou un nouveau classeur, vous devez suivre rigoureusement les 5 étapes ci-dessous :

### Étape 1 : Préparation du Classeur Google Sheets
1. Ouvrez le classeur maître Google Sheets.
2. Vérifiez que les 6 onglets obligatoires sont présents : `SESSIONS`, `INSCRIPTIONS`, `INSCRIPTIONSS`, `PARAMETRES`, `SESSION AGENDA` et `GED`.

📷 **[EMPLACEMENT CAPTURE N°1 : Onglets du Classeur Google Sheets]**  
*👉 Insérez ici une capture montrant la barre d'onglets en bas du classeur Google Sheets (SESSIONS, INSCRIPTIONS, PARAMETRES, etc.).*

---

### Étape 2 : Récupération des IDs d'Édition des Formulaires Google Forms
Vous devez utiliser **3 formulaires Google Forms** :
- Formulaire n°1 : **Inscription principale**
- Formulaire n°2 : **Liste d'attente**
- Formulaire n°3 : **Désinscription**

> [!IMPORTANT]
> Pour chacun de ces formulaires, vous devez ouvrir la page de modification (l'URL se termine par `/edit`) et copier l'identifiant présent dans l'adresse URL entre `/d/` et `/edit`.  
> *Exemple d'URL d'édition :* `https://docs.google.com/forms/d/1LOsvh4spCORP-xkR8dhTFlYvmXtXRvjNnYkTrmJCbQo/edit`  
> *L'ID d'Édition est :* `1LOsvh4spCORP-xkR8dhTFlYvmXtXRvjNnYkTrmJCbQo`

📷 **[EMPLACEMENT CAPTURE N°2 : URL d'Édition d'un Google Formulaire]**  
*👉 Insérez ici une capture de la barre d'adresse de votre navigateur sur Google Forms, en surbrillance sur l'ID d'édition.*

---

### Étape 3 : Configuration de l'Agenda Google
1. Rendez-vous sur [Google Agenda](https://calendar.google.com).
2. Dans les paramètres de l'agenda dédié aux formations, copiez l'**Identifiant de l'agenda** (ex: `c_xxxxxx@group.calendar.google.com`).
3. Assurez-vous que le compte Google qui exécute le script dispose des droits d'accès **"Modifier les événements"** sur cet agenda.

---

### Étape 4 : Préparation du Modèle de Convocation PDF & Dossier Drive
1. Créez un modèle de convocation au format Google Docs dans votre Google Drive.
2. Intégrez-y les balises entre crochets qui seront remplacées dynamiquement : `[NOM]`, `[PRENOM]`, `[DATE]`, `[HEURE_DEBUT]`, `[HEURE_FIN]`, `[LIEU]`, `[LIEN_MEET]`.
3. Copiez l'ID du Google Docs modèle ainsi que l'ID du dossier Google Drive cible recevant les convocations PDF générées.

---

### Étape 5 : Installation des Déclencheurs (Triggers)
1. Ouvrez votre classeur Google Sheets.
2. Dans la barre de menu supérieure, cliquez sur **NUMERICOACH** > **Installer / Réinitialiser les déclencheurs (Triggers)**.
3. Autorisez les permissions Google Apps Script demandées.

---

## 3. Description Détaillée des Onglets Google Sheets

### A. Onglet `SESSIONS`
C'est l'onglet de pilotage des formations. Vous devez y saisir chaque session de formation.

| Colonne | Nom de l'En-tête | Rôle & Format |
| :--- | :--- | :--- |
| **Col A** | `ID SESSION` | Identifiant unique généré (ex: `SES-0001`, `SES-0002`). |
| **Col B** | `FORMATION` | Nom du module de formation (ex: `Gemini 2H30`). |
| **Col C** | `DATE` | Date de la session (Format `JJ/MM/AAAA`). |
| **Col D** | `HEURE DEBUT` | Heure de début (ex: `09:00`). |
| **Col E** | `HEURE FIN` | Heure de fin (ex: `11:30`). |
| **Col H** | `NB DE PLACES` | Capacité maximale de la session (ex: `4`). |
| **Col K** | `Publier` | Case à cocher. Cochez `TRUE` pour publier la session sur le catalogue et les formulaires. |
| **Col L** | `Nb Inscrits` | Calculé automatiquement par le script. |
| **Col M** | `Places Restantes` | Calculé par la formule `=Col H - Col L`. |

📷 **[EMPLACEMENT CAPTURE N°3 : Onglet SESSIONS]**  
*👉 Insérez ici une capture du tableau de l'onglet SESSIONS avec des lignes de sessions saisies et des cases cochées dans la colonne Publier.*

---

### B. Onglet `PARAMETRES`
Cet onglet contient l'ensemble des clés de configuration utilisées par les automatismes.

📷 **[EMPLACEMENT CAPTURE N°4 : Onglet PARAMETRES]**  
*👉 Insérez ici une capture complète de l'onglet PARAMETRES montrant la colonne B (Clé) et la colonne C (Valeur).*

---

## 4. Référentiel des Paramètres Système (Onglet `PARAMETRES`)

Voici le tableau exhaustif des paramètres requis dans la feuille `PARAMETRES` :

| Clé (Colonne B) | Description & Instructions | Exemple de Valeur (Colonne C) |
| :--- | :--- | :--- |
| `ID Agenda >` | Identifiant de l'agenda Google récepteur. | `c_7da3b16909fd...@group.calendar.google.com` |
| `Mail Expediteur` | Adresse e-mail utilisée pour l'envoi des convocations. | `thierry.vanoffe@ext.leroymerlin.fr` |
| `Id Form Edit Inscriptions` | ID d'Édition du Google Form Inscription. | `1LOsvh4spCORP-xkR8dhTFlYvmXtXRvjNnYkTrmJCbQo` |
| `Id Form Edit Liste Attente` | ID d'Édition du Google Form Liste d'Attente. | `1KzGeU5UDYmUlw8EgeT1GrGCldZZPbHpc7ZPGpinssk` |
| `Id Form Edit Desinscription` | ID d'Édition du Google Form Désinscription. | `1UIYYprxFckHFDQYeARtj1An08k0nTUZ6hhA6TRSI0` |
| `ID du docs Modèle de convocation >` | ID du fichier Google Docs modèle. | `1bJfgjsants-waATPS9Ks5C9ctPG6cu0BKopr5Ek0Gw0` |
| `ID du dossier recueillant les convocations >` | ID du dossier Google Drive cible. | `1N0Vipl0kto3RhcDW-3AAD-8IDfIZqqfB` |

---

## 5. Fonctionnement des Automatismes & Services

### 1. Synchronisation des Formulaires (`updateFormChoices`)
Dès que vous modifiez une session dans l'onglet `SESSIONS` ou que vous cliquez sur le menu de mise à jour :
- Le script analyse les places disponibles.
- Le formulaire **Inscription** est mis à jour avec les sessions publiées ayant au moins 1 place libre.
- Les formulaires **Liste d'Attente** et **Désinscription** sont mis à jour avec la totalité des sessions publiées.
- Les questions de type **Liste déroulante**, **Choix multiple** et **Cases à cocher** sont gérées automatiquement.

📷 **[EMPLACEMENT CAPTURE N°5 : Formulaire d'Inscription à jour]**  
*👉 Insérez ici une capture de votre formulaire d'inscription Google Forms montrant la liste des sessions à jour.*

---

### 2. Gestion de l'Agenda Google (`AgendaService`)
Lorsqu'un participant s'inscrit :
- L'événement correspondant à la session est recherché ou créé dans Google Agenda.
- Un lien **Google Meet** est généré.
- Le titre de l'événement s'actualise avec le nombre d'inscrits :  
  - `(1 participant) Gemini 2H30 [SES-0002]`
  - `(3 participants) Gemini 2H30 [SES-0002]`
- La description de l'événement est enrichie avec la liste nominative complète des participants.
- L'invité reçoit la notification Google Calendar.

---

### 3. Catalogue WebApp (`CatalogTemplate.html`)
L'application Web affiche les cartes de formations en direct :
- Si 1 seule place est disponible ➔ `Plus que 1 place.` *(Gestion dynamique du singulier)*.
- Si 2 places ou plus ➔ `Plus que 2 places.` *(Gestion dynamique du pluriel)*.

📷 **[EMPLACEMENT CAPTURE N°6 : Catalogue WebApp Public]**  
*👉 Insérez ici une capture du catalogue WebApp montrant les cartes de sessions colorées.*

---

## 6. Guide d'Utilisation du Menu Administrateur OUTILS

Dans la barre supérieure de votre classeur Google Sheets, vous disposez du menu **OUTILS** :

1. **`Mettre à jour les sessions dans le Formulaire`** :  
   À exécuter après la saisie de nouvelles sessions pour forcer la synchronisation immédiate des 3 formulaires.
2. **`Synchroniser les événements Agenda`** :  
   Permet de vérifier et de créer l'ensemble des créneaux dans Google Agenda.
3. **`Installer / Réinitialiser les déclencheurs (Triggers)`** :  
   À exécuter si une soumission de formulaire ne déclenche plus l'envoi d'e-mail.
4. **`Retraiter les inscriptions non traitées`** :  
   Permet de relancer l'envoi des convocations pour les inscriptions enregistrées lors d'une panne réseau.
5. **`Générer les Documentations (Google Docs)`** :  
   Crée automatiquement les deux guides officiels au format Google Docs dans votre Google Drive.

---

## 7. Dépannage & Résolution des Incidents Courants

| Symptôme | Cause Probable | Solution Administrateur |
| :--- | :--- | :--- |
| **Notification `0 formulaire(s) mis à jour`** | L'ID présent dans `PARAMETRES` est le lien public (`/viewform`) au lieu de l'ID d'Édition. | Ouvrez le formulaire en mode édition, copiez l'ID présent dans l'URL `/edit` et collez-le dans la feuille `PARAMETRES`. |
| **Les événements n'apparaissent pas dans Google Agenda** | Le compte Google exécutant le script n'a pas les droits sur l'agenda cible. | Dans Google Agenda, partagez l'agenda avec l'adresse e-mail de l'administrateur en donnant l'autorisation *"Modifier les événements"*. |
| **E-mail de confirmation non reçu** | Le déclencheur `onFormSubmit` s'est désactivé suite à une mise à jour. | Rendez-vous sur le menu **NUMERICOACH** > **Installer / Réinitialiser les déclencheurs**. |
