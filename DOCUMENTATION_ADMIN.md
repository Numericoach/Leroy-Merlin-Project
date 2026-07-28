# Documentation Administrateur - Système de Gestion des Inscriptions Leroy Merlin x Numericoach

Ce document constitue le guide de référence complet pour l'administration, le paramétrage, le suivi et la maintenance du système d'automatisation des inscriptions aux formations Leroy Merlin.

---

## 📋 Table des Matières
1. [Vue d'Ensemble & Architecture du Code](#1-vue-densemble--architecture-du-code)
2. [Structure de la Base de Données (Google Sheets)](#2-structure-de-la-base-de-données-google-sheets)
3. [Configuration des Paramètres (Onglet `PARAMETRES`)](#3-configuration-des-paramètres-onglet-parametres)
4. [Gestion des Formulaires Google Forms](#4-gestion-des-formulaires-google-forms)
5. [Gestion de l'Agenda Google & des Visioconférences](#5-gestion-de-lagenda-google--des-visioconférences)
6. [Génération des Convocations PDF & Envois d'E-mails](#6-génération-des-convocations-pdf--envois-d-e-mails)
7. [Catalogue WebApp Dynamique](#7-catalogue-webapp-dynamique)
8. [Procédures d'Administration & Menu NUMERICOACH](#8-procédures-dadministration--menu-numericoach)
9. [Guide de Déploiement & Maintenance Technique](#9-guide-de-déploiement--maintenance-technique)

---

## 1. Vue d'Ensemble & Architecture du Code

Le projet est développé en TypeScript et déployé sur Google Apps Script via la CLI `clasp`. Le code est modulaire et structuré dans le dossier `src/` :

* **[src/config/Settings.ts](file:///c:/Project/Leroy-Merlin-Project/src/config/Settings.ts)** : Gestion de l'accès dynamique et de la mise en cache des clés de configuration de l'onglet `PARAMETRES`.
* **[src/services/FormService.ts](file:///c:/Project/Leroy-Merlin-Project/src/services/FormService.ts)** : Gestion de la synchronisation automatique des choix de sessions (listes déroulantes, choix multiples, cases à cocher) dans les formulaires Google Forms.
* **[src/services/AgendaService.ts](file:///c:/Project/Leroy-Merlin-Project/src/services/AgendaService.ts)** : Gestion de la création et de la mise à jour des événements Google Agenda, ajout des invités, gestion du nombre d'inscrits dans le titre et mise à jour de la description nominative.
* **[src/services/MailService.ts](file:///c:/Project/Leroy-Merlin-Project/src/services/MailService.ts)** : Envoi automatisé des convocations par e-mail avec pièce jointe PDF et liens d'action (désinscription, liste d'attente).
* **[src/services/PdfService.ts](file:///c:/Project/Leroy-Merlin-Project/src/services/PdfService.ts)** : Génération automatique des convocations au format PDF à partir d'un modèle Google Docs dans Google Drive.
* **[src/handlers/FormHandler.ts](file:///c:/Project/Leroy-Merlin-Project/src/handlers/FormHandler.ts)** : Traitement principal déclenché lors de la soumission d'une inscription ou d'une désinscription.
* **[src/handlers/triggers.ts](file:///c:/Project/Leroy-Merlin-Project/src/handlers/triggers.ts)** : Installation et exécution des déclencheurs automatiques (`onFormSubmit`, `onEditTrigger`).
* **[src/web/WebApp.ts](file:///c:/Project/Leroy-Merlin-Project/src/web/WebApp.ts)** : Contrôleur du catalogue de formations publié en application Web Google.
* **[src/web/CatalogTemplate.html](file:///c:/Project/Leroy-Merlin-Project/src/web/CatalogTemplate.html)** : Interface utilisateur (HTML/CSS) du catalogue public des sessions disponibles.
* **[src/utils.ts](file:///c:/Project/Leroy-Merlin-Project/src/utils.ts)** : Fonctions utilitaires (extraction d'ID, formatage propre des dates et des heures).

---

## 2. Structure de la Base de Données (Google Sheets)

Le classeur Google Sheets fait office de base de données centrale. Il comporte 6 onglets principaux :

| Onglet | Rôle Administrateur |
| :--- | :--- |
| **`SESSIONS`** | Saisie des sessions de formation (Dates, Heures, Intitulés, Lieu, Nombre de places, État de publication). |
| **`INSCRIPTIONS`** | Liste consolidée et nettoyée des inscrits validés avec leurs informations et liens de désinscription. |
| **`INSCRIPTIONSS`** | Onglet de réception brute des réponses soumises via le Google Formulaire d'inscription. |
| **`PARAMETRES`** | Table des variables de configuration système (IDs de formulaires, ID d'agenda, e-mail expéditeur). |
| **`SESSION AGENDA`** | Table d'association entre l'identifiant de la session (`SES-XXXX`) et l'ID de l'événement Google Agenda correspondant. |
| **`GED`** | Historique et archivage des liens vers les documents PDF de convocation générés. |

---

## 3. Configuration des Paramètres (Onglet `PARAMETRES`)

Pour modifier le comportement du système sans toucher au code, il faut accéder à l'onglet **`PARAMETRES`** du Google Sheets.

![Onglet Paramètres Système](file:///C:/Users/Arthu/.gemini/antigravity-ide/brain/9ff38e37-2448-4fa8-bb45-69f018aa355c/media__1785237973899.png)

![IDs de Formulaires et Textes de Convocations](file:///C:/Users/Arthu/.gemini/antigravity-ide/brain/9ff38e37-2448-4fa8-bb45-69f018aa355c/media__1785237993748.png)

### Tableau des clés de configuration principales :

| Clé (Colonne B) | Description & Instructions |
| :--- | :--- |
| **`ID Agenda >`** | Rentrez l'identifiant complet de l'agenda Google cible (ex: `c_xxxxxx@group.calendar.google.com`). Le compte exécutant le script doit disposer des droits d'écriture sur cet agenda. |
| **`Mail Expediteur`** | Définissez l'adresse e-mail qui apparaîtra en expéditeur des convocations (ex: `thierry.vanoffe@ext.leroymerlin.fr`). |
| **`Id Form Edit Inscriptions`** | Rentrez l'ID d'Édition du Google Form d'inscription principal (trouvable dans l'URL `/edit` du formulaire). |
| **`Id Form Edit Liste Attente`** | Rentrez l'ID d'Édition du Google Form de Liste d'Attente. |
| **`Id Form Edit Desinscription`** | Rentrez l'ID d'Édition du Google Form de Désinscription. |
| **`ID du docs Modèle de convocation >`** | ID du modèle Google Docs servant de trame pour générer les convocations PDF. |
| **`ID du dossier recueillant les convocations >`** | ID du dossier Google Drive dans lequel les convocations PDF sont enregistrées. |

---

## 4. Gestion des Formulaires Google Forms

Le système synchronise dynamiquement les options de choix de sessions sur 3 formulaires Google Forms :

1. **Formulaire d'Inscription Principal** : Ne présente que les sessions **publiées et ayant des places restantes**.
2. **Formulaire de Liste d'Attente** : Présente l'ensemble des sessions publiées afin d'accepter les demandes lorsque la session est complète.
3. **Formulaire de Désinscription** : Présente l'ensemble des sessions pour permettre à un participant d'annuler sa présence.

![Formulaire d'Édition Liste d'Attente](file:///C:/Users/Arthu/.gemini/antigravity-ide/brain/9ff38e37-2448-4fa8-bb45-69f018aa355c/media__1785238995373.png)

![Formulaire de Désinscription](file:///C:/Users/Arthu/.gemini/antigravity-ide/brain/9ff38e37-2448-4fa8-bb45-69f018aa355c/media__1785238733259.png)

> [!IMPORTANT]
> Pour que la synchronisation fonctionne, vous devez toujours renseigner les **IDs d'Édition** des formulaires (se terminant par `/edit` dans le navigateur) dans l'onglet `PARAMETRES`, et non les liens de réponse publique (`/viewform` ou `1FAIpQL...`).

---

## 5. Gestion de l'Agenda Google & des Visioconférences

Le module [src/services/AgendaService.ts](file:///c:/Project/Leroy-Merlin-Project/src/services/AgendaService.ts) orchestre la création et le rafraîchissement des événements d'agenda :

* **Titre dynamique de l'événement** : Se met à jour au format :
  - `(1 participant) Titre de la Formation [SES-XXXX]` *(si 1 seul inscrit)*
  - `(X participants) Titre de la Formation [SES-XXXX]` *(si plusieurs inscrits)*
* **Lien Google Meet de visioconférence** : Un lien Google Meet unique est automatiquement rattaché à l'événement lors de sa création.
* **Description nominative** : La description de l'événement est enrichie en temps réel avec la liste exhaustive des participants (Nom, Prénom, Magasin, Email et nombre de places réservées).
* **Gestion des annulations** : Lorsqu'un inscrit se désinscrit, il est automatiquement retiré des invités de l'événement et le titre ainsi que la description sont mis à jour instantanément.

---

## 6. Génération des Convocations PDF & Envois d'E-mails

Lors de la validation d'une inscription :
1. Le service [src/services/PdfService.ts](file:///c:/Project/Leroy-Merlin-Project/src/services/PdfService.ts) copie le modèle Google Docs référencé dans `PARAMETRES`, remplace les balises (Nom, Prénom, Session, Date, Heure, Lien Meet) puis convertit le document en PDF.
2. Le fichier PDF est sauvegardé dans le dossier Google Drive défini dans `PARAMETRES` et son lien est archivé dans l'onglet `GED`.
3. Le service [src/services/MailService.ts](file:///c:/Project/Leroy-Merlin-Project/src/services/MailService.ts) expédie la convocation avec le PDF en pièce jointe ainsi qu'un lien personnalisé de désinscription.

---

## 7. Catalogue WebApp Dynamique

Le catalogue interactif est hébergé via Google Apps Script (fichiers [src/web/WebApp.ts](file:///c:/Project/Leroy-Merlin-Project/src/web/WebApp.ts) et [src/web/CatalogTemplate.html](file:///c:/Project/Leroy-Merlin-Project/src/web/CatalogTemplate.html)).

![Catalogue Public WebApp](file:///C:/Users/Arthu/.gemini/antigravity-ide/brain/9ff38e37-2448-4fa8-bb45-69f018aa355c/media__1785226539788.png)

* **Règles d'affichage des places** :
  - S'il reste **1 place** disponible : la carte affiche la mention exacte `Plus que 1 place.` *(au singulier)*.
  - S'il reste **2 places ou plus** : la carte affiche `Plus que X places.` *(au pluriel)*.
* **Couleurs dynamiques** : La couleur d'en-tête et du bouton d'inscription s'adapte à la thématique de la formation enregistrée dans la feuille `SESSIONS`.

---

## 8. Procédures d'Administration & Menu NUMERICOACH

Un menu sur-mesure intitulé **NUMERICOACH** est présent dans la barre supérieure du classeur Google Sheets.

### Actions disponibles dans le menu :

1. **`Mettre à jour les sessions dans le Formulaire`** :
   Recalcule les places et met à jour instantanément la liste des créneaux dans les 3 formulaires Google Forms.

2. **`Installer / Réinitialiser les déclencheurs (Triggers)`** :
   Recrée les déclencheurs automatiques Google Apps Script en cas de perte de synchronisation des formulaires ou d'édition.

3. **`Retraiter les inscriptions non traitées`** :
   Permet d'exécuter un traitement de rattrapage sur l'onglet `INSCRIPTIONSS` si des soumissions n'ont pas pu être traitées suite à un incident réseau Google.

---

## 9. Guide de Déploiement & Maintenance Technique

Pour apporter des modifications au code source du projet :

1. **Ouvrir le projet dans VS Code** :
   Rendez-vous dans le répertoire `c:\Project\Leroy-Merlin-Project`.

2. **Modifier le code TypeScript ou HTML** :
   Toutes les sources se trouvent sous le répertoire `src/`.

3. **Compiler et déployer sur Google Apps Script** :
   Exécutez la commande suivante dans le terminal :
   ```bash
   npm run push
   ```
   *Cette commande compile le code TypeScript en JavaScript dans le dossier `dist/` puis effectue un `clasp push --force` vers le projet Google Apps Script lié.*

4. **Vérification dans Google Apps Script** :
   Rendez-vous sur la console Apps Script pour vérifier que la dernière version a bien été poussée sans erreur.
