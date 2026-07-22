/**
 * Déclencheur sur soumission du formulaire d'inscription
 * Protégé par LockService contre les soumissions simultanées
 */
function onSubmit(e?: GoogleAppsScript.Events.SheetsOnFormSubmit): void {
  if (!e || !e.range) return;

  const lock = LockService.getScriptLock();
  // Attendre jusqu'à 10 secondes pour acquérir le verrou exclusif
  const hasLock = lock.tryLock(10000);
  if (!hasLock) {
    Logger.log("Impossible d'acquérir le verrou. Soumission reportée ou ignorée.");
    return;
  }

  try {
    const thisResponses = e.range.getValues()[0];
    const thisSheetName = e.range.getSheet().getSheetName();

    if (thisSheetName === "INSCRIPTIONSS" || thisSheetName === "INSCRIPTIONS FORM") {
      const thisTime = thisResponses[0];
      const thisEmail = (thisResponses[9] || thisResponses[1] || "").toString().trim(); // Colonne E-mail
      const thisSession = (thisResponses[5] || thisResponses[3] || "").toString().trim(); // Colonne Session
      if (!thisSession || !thisEmail) return;
      
      const match = thisSession.match(/\[(.*)\]/);
      if (!match) return;
      const thisSessionid = match[1];
      
      // 1. VÉRIFIER LES PLACES RESTANTES SOUS VERROU
      const sheetSessions = ss.getSheetByName("SESSIONS");
      let remainingSeats = 1; // Défaut à 1 si non trouvé
      if (sheetSessions) {
        const lastRowSessions = sheetSessions.getLastRow();
        if (lastRowSessions >= 2) {
          const sessionsData = sheetSessions.getRange(2, 2, lastRowSessions - 1, 12).getValues();
          for (let i = 0; i < sessionsData.length; i++) {
            if (sessionsData[i][0] === thisSessionid) {
              remainingSeats = Number(sessionsData[i][11]);
              break;
            }
          }
        }
      }

      if (remainingSeats <= 0) {
        Logger.log("Inscription refusée : plus de places disponibles pour " + thisSessionid);
        return;
      }

      // 2. VÉRIFIER SI DÉJÀ INSCRIT SOUS VERROU
      if (!sheetInscriptions) return;
      const maxRows = sheetInscriptions.getMaxRows();
      let verif: any[][] = [];
      if (maxRows > 1) {
        const sessionsEmail = sheetInscriptions.getRange(2, 2, maxRows - 1, 2).getValues();
        verif = sessionsEmail.filter(row => (row[0] === thisSessionid && row[1] === thisEmail));
      }
      
      if (verif.length > 0) {
        Logger.log("Déjà inscrit : " + thisEmail + " à " + thisSessionid);
        return;
      }

      // 3. INSCRIPTION DANS LE SHEETS
      inscription(thisTime, thisSessionid, thisEmail);

      // 4. FONCTION PRINCIPALE : AJOUT DANS GOOGLE AGENDA (Priorité absolue)
      addParticipantToCalendar(thisSessionid, thisEmail);

      // 5. FONCTION SECONDAIRE : ENVOI DE LA CONVOCATION / CONFIRMATION (Découplé)
      try {
        sendConfirmationMail(thisSessionid, thisEmail);
      } catch (mailErr) {
        Logger.log("Avertissement : échec de l'envoi d'e-mail (n'impacte pas l'inscription ni l'agenda) : " + mailErr);
      }
      
      // 6. MISE À JOUR DYNAMIQUE DES CHOIX DU FORMULAIRE
      try {
        updateFormChoices();
      } catch (formErr) {
        Logger.log("Erreur lors de la mise à jour des choix du formulaire : " + formErr);
      }
    }
  } catch (err) {
    Logger.log("Erreur critique dans onSubmit : " + err);
  } finally {
    if (hasLock) {
      lock.releaseLock();
    }
  }
}

/**
 * Enregistrer l'inscription dans la feuille de calcul
 */
function inscription(time: any, sessionId: string, email: string): void {
  if (sheetInscriptions) {
    sheetInscriptions.appendRow([time, sessionId, email]);
  }
}

/**
 * Mettre à jour dynamiquement la liste déroulante des sessions disponibles dans le Google Form
 */
function updateFormChoices(): void {
  if (!sheetParametres) return;
  const formId = sheetParametres.getRange(pnParaCel["PARAMETRE_ID_FORMS_INSCRIPTION"]).getValue() as string;
  if (!formId) {
    Logger.log("ID Forms Inscription non trouvé dans les paramètres.");
    return;
  }

  const form = FormApp.openById(formId);
  const items = form.getItems(FormApp.ItemType.LIST);
  let sessionItem: GoogleAppsScript.Forms.ListItem | null = null;

  for (let i = 0; i < items.length; i++) {
    if (items[i].getTitle().indexOf("Inscription à la formation suivante") > -1 || items[i].getTitle().indexOf("Session") > -1) {
      sessionItem = items[i].asListItem();
      break;
    }
  }

  if (!sessionItem) {
    Logger.log("Question de sélection de session non trouvée dans le formulaire.");
    return;
  }

  const sheetSessions = ss.getSheetByName("SESSIONS");
  if (!sheetSessions) return;

  const lastRow = sheetSessions.getLastRow();
  if (lastRow < 2) return;

  const values = sheetSessions.getRange(2, 2, lastRow - 1, 14).getValues();
  const choices: string[] = [];

  values.forEach(function(row) {
    const sessionId = row[0];
    const publish = row[9];
    const remaining = row[11];
    const formationTitle = row[13];

    if (sessionId && publish && Number(remaining) > 0) {
      const dateObj = new Date(row[2]);
      const dateStr = dateObj.getDate() + "/" + (dateObj.getMonth() + 1) + "/" + dateObj.getFullYear();
      choices.push(formationTitle + " - " + dateStr + " [" + sessionId + "]");
    }
  });

  if (choices.length > 0) {
    sessionItem.setChoiceValues(choices);
    Logger.log("Formulaire mis à jour avec " + choices.length + " sessions disponibles.");
  } else {
    sessionItem.setChoiceValues(["Aucune session disponible pour le moment"]);
    Logger.log("Aucune session disponible.");
  }
}

/**
 * Envoi de l'e-mail de convocation / confirmation avec gestion optionnelle du PDF si modèle présent
 */
function sendConfirmationMail(sessionId: string, email: string): void {
  if (!sheetParametres) return;

  const urlDesinscription = sheetParametres.getRange(pnParaCel["PARAMETRE_ID_FORMS_DESINSCRIPTION"]).getValue() || "";
  
  let entrySessionId = "entry.2116080188";
  if (pnParaCel["PARAMETRE_ENTRY_SESSION"]) {
    const customEntry = sheetParametres.getRange(pnParaCel["PARAMETRE_ENTRY_SESSION"]).getValue();
    if (customEntry) entrySessionId = customEntry.toString().trim();
  }

  let entryEmailId = "entry.193822625";
  if (pnParaCel["PARAMETRE_ENTRY_EMAIL"]) {
    const customEmailEntry = sheetParametres.getRange(pnParaCel["PARAMETRE_ENTRY_EMAIL"]).getValue();
    if (customEmailEntry) entryEmailId = customEmailEntry.toString().trim();
  }

  const desinscriptionLink = "https://docs.google.com/forms/d/e/" + urlDesinscription + 
    "/viewform?usp=pp_url&" + entryEmailId + "=" + encodeURIComponent(email) + 
    "&" + entrySessionId + "=[" + encodeURIComponent(sessionId) + "]";

  const connexionInfo = sheetParametres.getRange(pnParaCel["PARAMETRE_CONNEXION_1"]).getValue() || "Lien Meet inclus dans votre invitation Agenda";

  // Optionnel : Générer le PDF de convocation si l'ID du modèle est renseigné dans PARAMETRES
  let pdfAttachment: GoogleAppsScript.Base.Blob | null = null;
  let pdfUrl = "";

  try {
    if (pnParaCel["PARAMETRE_ID_MODELE_CONVOC"]) {
      const modeleConvocationId = sheetParametres.getRange(pnParaCel["PARAMETRE_ID_MODELE_CONVOC"]).getValue() as string;
      if (modeleConvocationId && modeleConvocationId.length > 10) {
        const modeleConvocation = DriveApp.getFileById(modeleConvocationId);
        const folderConvocationId = pnParaCel["PARAMETRE_ID_DOSSIER_CONVOC"] ? sheetParametres.getRange(pnParaCel["PARAMETRE_ID_DOSSIER_CONVOC"]).getValue() as string : "";
        const folderConvocation = folderConvocationId ? DriveApp.getFolderById(folderConvocationId) : DriveApp.getRootFolder();
        
        const convocationName = "Convocation_" + sessionId + "_" + email;
        const convocationDoc = modeleConvocation.makeCopy(convocationName, folderConvocation);
        const doc = DocumentApp.openById(convocationDoc.getId());
        const body = doc.getBody();

        body.replaceText("{{DATE AUJOURDHUI}}", Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy'));
        body.replaceText("{{SESSION ID}}", sessionId);
        body.replaceText("{{EMAIL}}", email);
        doc.saveAndClose();

        pdfAttachment = convocationDoc.getAs('application/pdf');
        pdfAttachment.setName(convocationName + ".pdf");
        const pdfFile = folderConvocation.createFile(pdfAttachment);
        pdfUrl = pdfFile.getUrl();

        try { convocationDoc.setTrashed(true); } catch (e) {}
      }
    }
  } catch (pdfErr) {
    Logger.log("Avertissement : la génération du PDF n'a pas pu être effectuée (envoi sans pièce jointe) : " + pdfErr);
  }

  const subject = "Convocation & Confirmation d'inscription - Formation Leroy Merlin [" + sessionId + "]";
  
  let htmlBody = "<div style='font-family: Arial, sans-serif; color: #333; max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 8px; overflow: hidden;'>"
    + "<div style='background-color: #0596DE; padding: 20px; text-align: center; color: white;'>"
    + "<h2 style='margin: 0;'>Confirmation & Convocation de Formation</h2>"
    + "</div>"
    + "<div style='padding: 24px;'>"
    + "<p>Bonjour,</p>"
    + "<p>Votre inscription à la session de formation <b>[" + sessionId + "]</b> a bien été confirmée.</p>"
    + "<p><b>Invitation Agenda :</b> Une invitation Google Agenda contenant la date, l'heure et le lien de connexion vous a été envoyée.</p>"
    + "<div style='background-color: #f4f7f6; padding: 15px; border-radius: 6px; margin: 15px 0;'>"
    + "<b>Informations de connexion :</b><br>" + connexionInfo
    + "</div>";

  if (pdfUrl !== "") {
    htmlBody += "<p><a href='" + pdfUrl + "' style='display:inline-block; background-color:#0596DE; color:white; padding:10px 18px; text-decoration:none; border-radius:5px;'>Télécharger votre Convocation PDF</a></p>";
  }

  htmlBody += "<p style='margin-top: 25px;'><a href='" + desinscriptionLink + "' style='color:#CC3C25;'>Demander une désinscription</a></p>"
    + "</div>"
    + "<div style='background-color: #f9f9f9; padding: 12px; text-align: center; font-size: 12px; color: #777;'>"
    + "Numericoach &bull; Gestion des Formations Leroy Merlin"
    + "</div></div>";

  const mailOptions: GoogleAppsScript.Mail.MailAdvancedParameters = {
    to: email,
    subject: subject,
    htmlBody: htmlBody
  };

  if (pdfAttachment) {
    mailOptions.attachments = [pdfAttachment];
  }

  MailApp.sendEmail(mailOptions);
  Logger.log("Mail de convocation envoyé à " + email + " pour la session " + sessionId);
}

/**
 * Helper pour générer l'URL d'un formulaire pré-rempli avec l'entrée réelle
 */
function getPreFilledFormUrl(formId: string, entryId: string, selectedValue: string): string {
  return "https://docs.google.com/forms/d/e/" + formId + "/viewform?usp=pp_url&" + entryId + "=" + encodeURIComponent(selectedValue);
}
