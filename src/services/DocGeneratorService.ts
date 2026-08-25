/**
 * Génère automatiquement les documentations officielles au format Google Docs dans Google Drive
 */
function createGoogleDocsDocumentation(): void {
  const ui = (typeof SpreadsheetApp !== 'undefined' && SpreadsheetApp.getUi) ? SpreadsheetApp.getUi() : null;
  
  try {
    if (ss) ss.toast("⏳ Génération des documentations Google Docs dans votre Google Drive...", "OUTILS", 10);

    // -------------------------------------------------------------
    // 1. Création de la Documentation Administrateur
    // -------------------------------------------------------------
    const adminDoc = DocumentApp.create("Documentation Administrateur - Leroy Merlin x Numericoach");
    const adminBody = adminDoc.getBody();

    adminBody.appendParagraph("Documentation Administrateur").setHeading(DocumentApp.ParagraphHeading.TITLE);
    adminBody.appendParagraph("Système de Gestion des Inscriptions Leroy Merlin x Numericoach\n").setHeading(DocumentApp.ParagraphHeading.SUBTITLE);

    adminBody.appendParagraph("Ce document constitue le guide de référence complet pour l'administration, le paramétrage, le suivi et la maintenance du système d'automatisation des inscriptions aux formations.\n").setHeading(DocumentApp.ParagraphHeading.NORMAL);

    adminBody.appendParagraph("1. Vue d'Ensemble & Architecture du Code").setHeading(DocumentApp.ParagraphHeading.HEADING1);
    adminBody.appendParagraph("Le projet est développé en TypeScript et déployé sur Google Apps Script via la CLI clasp. Les fichiers sources principaux sont :");
    
    adminBody.appendListItem("src/config/Settings.ts : Gestion dynamique des paramètres de configuration et du menu.").setGlyphType(DocumentApp.GlyphType.BULLET);
    adminBody.appendListItem("src/services/FormService.ts : Synchronisation des choix de sessions dans les 3 formulaires Google Forms.").setGlyphType(DocumentApp.GlyphType.BULLET);
    adminBody.appendListItem("src/services/AgendaService.ts : Gestion des événements Google Agenda et création des liens Meet.").setGlyphType(DocumentApp.GlyphType.BULLET);
    adminBody.appendListItem("src/services/MailService.ts : Expédition automatisée des convocations e-mail.").setGlyphType(DocumentApp.GlyphType.BULLET);
    adminBody.appendListItem("src/services/PdfService.ts : Génération des convocations au format PDF.").setGlyphType(DocumentApp.GlyphType.BULLET);
    adminBody.appendListItem("src/web/WebApp.ts & CatalogTemplate.html : Application Web du catalogue public.").setGlyphType(DocumentApp.GlyphType.BULLET);

    adminBody.appendParagraph("\n2. Structure de la Base de Données (Google Sheets)").setHeading(DocumentApp.ParagraphHeading.HEADING1);
    adminBody.appendTable([
      ["Onglet", "Rôle Administrateur"],
      ["SESSIONS", "Saisie des sessions de formation (Dates, Heures, Lieux, Places disponibles)."],
      ["INSCRIPTIONS", "Base de données consolidée des participants inscrits."],
      ["INSCRIPTIONSS", "Réception des réponses brutes des formulaires d'inscription."],
      ["PARAMETRES", "Table des variables de configuration système."],
      ["SESSION AGENDA", "Table d'association entre l'ID de session et l'événement Google Agenda."],
      ["GED", "Historique et liens vers les convocations PDF générées."]
    ]);

    adminBody.appendParagraph("\n3. Configuration des Paramètres (Onglet PARAMETRES)").setHeading(DocumentApp.ParagraphHeading.HEADING1);
    adminBody.appendParagraph("Vous devez modifier les clés de configuration dans l'onglet PARAMETRES pour ajuster le comportement du système :");
    adminBody.appendTable([
      ["Clé Paramètre", "Description & Instructions"],
      ["ID Agenda >", "Identifiant complet du calendrier Google Agenda récepteur."],
      ["Mail Expediteur", "Adresse e-mail d'envoi des confirmations."],
      ["Id Form Edit Inscriptions", "ID d'Édition du Google Form d'Inscription."],
      ["Id Form Edit Liste Attente", "ID d'Édition du Google Form de Liste d'Attente."],
      ["Id Form Edit Desinscription", "ID d'Édition du Google Form de Désinscription."],
      ["ID du docs Modèle de convocation >", "ID du modèle Google Docs pour la convocation PDF."],
      ["ID du dossier recueillant les convocations >", "ID du dossier Google Drive récepteur des PDF."]
    ]);

    adminBody.appendParagraph("\n4. Menu NUMERICOACH").setHeading(DocumentApp.ParagraphHeading.HEADING1);
    adminBody.appendParagraph("Vous disposez d'un menu personnalisé dans Google Sheets pour gérer le système :");
    adminBody.appendListItem("Mettre à jour les sessions dans le Formulaire : Synchronise les 3 formulaires Google Forms.").setGlyphType(DocumentApp.GlyphType.BULLET);
    adminBody.appendListItem("Installer / Réinitialiser les déclencheurs (Triggers) : Recrée les triggers automatiques Google.").setGlyphType(DocumentApp.GlyphType.BULLET);
    adminBody.appendListItem("Retraiter les inscriptions non traitées : Rattrape les soumissions brutes non traitées.").setGlyphType(DocumentApp.GlyphType.BULLET);

    adminDoc.saveAndClose();

    // -------------------------------------------------------------
    // 2. Création du Guide Utilisateur
    // -------------------------------------------------------------
    const userDoc = DocumentApp.create("Guide Utilisateur - Inscriptions Leroy Merlin x Numericoach");
    const userBody = userDoc.getBody();

    userBody.appendParagraph("Guide Utilisateur").setHeading(DocumentApp.ParagraphHeading.TITLE);
    userBody.appendParagraph("Inscription aux Formations Leroy Merlin x Numericoach\n").setHeading(DocumentApp.ParagraphHeading.SUBTITLE);

    userBody.appendParagraph("Bienvenue dans le guide utilisateur. Ce document vous explique pas à pas comment consulter les formations, vous inscrire et gérer votre participation.\n").setHeading(DocumentApp.ParagraphHeading.NORMAL);

    userBody.appendParagraph("1. Consultation du Catalogue").setHeading(DocumentApp.ParagraphHeading.HEADING1);
    userBody.appendParagraph("Vous devez vous rendre sur le catalogue en ligne pour visualiser les sessions disponibles. Chaque fiche indique le titre, la date, l'horaire et le nombre de places restantes (ex: Plus que 1 place. ou Plus que 4 places.).");

    userBody.appendParagraph("\n2. Procédure d'Inscription").setHeading(DocumentApp.ParagraphHeading.HEADING1);
    userBody.appendListItem("Cliquez sur le bouton 'Je m'inscris' sous la fiche souhaitée.").setGlyphType(DocumentApp.GlyphType.NUMBER);
    userBody.appendListItem("Remplissez vos informations (Nom, Prénom, E-mail professionnel, Magasin).").setGlyphType(DocumentApp.GlyphType.NUMBER);
    userBody.appendListItem("Sélectionnez le créneau horaire et validez le formulaire.").setGlyphType(DocumentApp.GlyphType.NUMBER);

    userBody.appendParagraph("\n3. Réception de la Convocation & Agenda").setHeading(DocumentApp.ParagraphHeading.HEADING1);
    userBody.appendParagraph("Dès la validation de votre inscription, vous recevez un e-mail de confirmation contenant votre convocation PDF en pièce jointe. Un événement Google Agenda est automatiquement ajouté à votre calendrier avec le lien de visioconférence Google Meet.");

    userBody.appendParagraph("\n4. Désinscription & Liste d'Attente").setHeading(DocumentApp.ParagraphHeading.HEADING1);
    userBody.appendParagraph("En cas d'imprévu, vous devez cliquer sur le lien 'Se désinscrire' situé au bas de votre e-mail de confirmation. Si une session est complète, vous devez vous inscrire sur le formulaire de Liste d'Attente pour être recontacté en priorité.");

    userDoc.saveAndClose();

    const adminUrl = adminDoc.getUrl();
    const userUrl = userDoc.getUrl();

    Logger.log("Google Doc Admin créé : " + adminUrl);
    Logger.log("Google Doc Utilisateur créé : " + userUrl);

    if (ui) {
      ui.alert(
        "✅ Documentations Google Docs générées !",
        "Les 2 documentations ont été créées dans votre Google Drive :\n\n" +
        "1. Documentation Administrateur :\n" + adminUrl + "\n\n" +
        "2. Guide Utilisateur :\n" + userUrl,
        ui.ButtonSet.OK
      );
    }
  } catch (err) {
    Logger.log("Erreur lors de la création des Google Docs : " + err);
    if (ui) {
      ui.alert("❌ Erreur lors de la génération des Google Docs : " + err);
    }
  }
}
