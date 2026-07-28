# Documentation Administrateur - Système de Gestion des Inscriptions Leroy Merlin x Numericoach

Ce document décrit le fonctionnement, la configuration et la maintenance du système d'automatisation des inscriptions aux formations pour les directeurs de magasins Leroy Merlin.

---

## 📋 Table des Matières
1. [Architecture de la BDD Google Sheets](#1-architecture-de-la-bdd-google-sheets)
2. [Configuration des Paramètres (Onglet PARAMETRES)](#2-configuration-des-paramètres-onglet-parametres)
3. [Fonctionnement des Flux d'Inscriptions](#3-fonctionnement-des-flux-dinscriptions)
4. [Gestion de l'Agenda Google](#4-gestion-de-lagenda-google)
5. [Maintenance & Dépannage](#5-maintenance--dépannage)

---

## 1. Architecture de la BDD Google Sheets

Le classeur Google Sheets pilote l'intégralité du système. Il contient les onglets suivants :

| Onglet | Rôle administrateur |
| :--- | :--- |
| **`SESSIONS`** | Liste des sessions de formation programmées (dates, heures, lieux, places disponibles, titre). |
| **`INSCRIPTIONS`** | Base de données centrale des directeurs inscrits (Gérée automatiquement par script). |
| **`INSCRIPTIONSS`** | Réception des réponses brutes du Google Formulaire d'inscription. |
| **`PARAMETRES`** | Variables de configuration globales (IDs de formulaires, e-mails, modèles). |
| **`SESSION AGENDA`** | Table de correspondance entre les IDs de sessions et les événements Google Agenda créés. |
| **`GED`** | Historique des liens vers les documents PDF de convocation générés automatiquement. |

---

## 2. Configuration des Paramètres (Onglet `PARAMETRES`)

Pour personnaliser l'expéditeur, l'agenda ou les formulaires de redirection, modifiez simplement les valeurs dans l'onglet **`PARAMETRES`** :

| Clé Paramètre (Col A) | Rôle (Col B) | Exemple de valeur |
| :--- | :--- | :--- |
| `PARAMETRE_EXPEDITEUR_EMAIL` | Adresse e-mail d'envoi des confirmations (doit être votre e-mail ou un alias Gmail validé). | `formation@numericoach.fr` |
| `PARAMETRE_NOM_EXPEDITEUR` | Nom d'affichage dans la boîte de réception des directeurs. | `Service Formation Leroy Merlin` |
| `PARAMETRE_ID_AGENDA` | ID de l'agenda cible Google Agenda où créer les événements. | `c_xxxxxxxxx@group.calendar.google.com` |
| `PARAMETRE_ID_MODELE_CONVOC` | ID du fichier Google Docs servant de template pour la convocation PDF. | `1bJfgjsants-waATPS9Ks5C9ctPG...` |
| `PARAMETRE_ID_DOSSIER_CONVOC` | ID du dossier Google Drive où ranger les PDF générés. | `1Hm1L-GAqArtLPgE53ygVUTSe_...` |
| `PARAMETRE_ID_FORMS_INSCRIPTION` | ID du formulaire Google Forms principal d'inscription. | `1FAIpQLSfxxxxxxxxx` |
| `PARAMETRE_ID_FORMS_DESINSCRIPTION` | ID du formulaire Google Forms de désinscription. | `1FAIpQLSdyyyyyyyyy` |
| `PARAMETRE_ID_FORMS_LISTE_ATTENTE` | ID du formulaire Google Forms de liste d'attente. | `1FAIpQLSzzzzzzzzz` |

---

## 3. Fonctionnement des Flux d'Inscriptions

### A. Inscription Standard
1. Le directeur répond au Google Formulaire d'inscription.
2. Le déclencheur `onSubmit` intercepte la soumission :
   - Vérifie s'il reste des places disponibles dans l'onglet `SESSIONS`.
   - Si **oui** : Enregistre le directeur dans `INSCRIPTIONS`, l'ajoute à l'Agenda Google, génère et lui envoie la convocation PDF.
   - Si **non** : Envoie un e-mail lui indiquant que la session est complète avec un lien vers la **Liste d'attente**.

### B. Désinscription (Annulation)
1. Le directeur clique sur le lien personnalisé de désinscription présent dans son e-mail de confirmation.
2. Le formulaire pré-rempli s'ouvre, il valide sa désinscription.
3. Le script :
   - Supprime sa ligne dans l'onglet `INSCRIPTIONS`.
   - Le retire de l'invitation Google Agenda.
   - Déclenche la promotion automatique du premier inscrit en liste d'attente (si configuré) ou notifie les personnes en attente.

---

## 4. Gestion de l'Agenda Google

* **Synchronisation du Nombre de Participants** : 
  Le titre de l'événement se met automatiquement à jour au format `[X participants] Titre formation [ID-SESSION]`.
* **Liste nominative en temps réel** :
  La description de l'événement de l'agenda contient la liste complète des directeurs inscrits (Prénom Nom, Magasin, Email et nombre de participants renseignés).

---

## 5. Maintenance & Dépannage

Un menu personnalisé **`NUMERICOACH`** est ajouté en haut du Google Sheets pour gérer le système de manière autonome :

### Actions Rapides du Menu Sheets :
* **Installer / Réinitialiser les déclencheurs (Triggers)** : À exécuter si le formulaire ne semble plus envoyer d'e-mails automatique lors d'une soumission. Cela recrée les déclencheurs automatiques Google.
* **Mettre à jour les sessions dans le Formulaire** : Recalcule les places restantes et rafraîchit la liste déroulante des sessions disponibles dans le Google Forms.
* **Retraiter les inscriptions non traitées** : Permet de scanner l'onglet brut du formulaire `INSCRIPTIONSS` et de rattraper manuellement les inscriptions qui n'auraient pas été enregistrées (en cas de coupure de service de Google).

### Déploiement du Code source :
Si vous modifiez le code localement en TypeScript :
1. Compilez et poussez les changements avec `npm run push` (qui exécute le build et `clasp push --force`).
2. Tentez de tester d'abord en navigation privée si vous rencontrez des blocages de droits liés à des comptes multiples sur votre navigateur.
