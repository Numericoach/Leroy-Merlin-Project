function onSubmit(e) {
  if (!e || !e.range) return;

  var lock = LockService.getScriptLock();
  var hasLock = lock.tryLock(10000);
  if (!hasLock) {
    Logger.log("Impossible d'acquérir le verrou. Soumission reportée ou ignorée.");
    return;
  }

  try {
    var thisResponses = e.range.getValues()[0];
    var thisSheetName = e.range.getSheet().getSheetName();

    if (thisSheetName === "INSCRIPTIONSS" || thisSheetName === "INSCRIPTIONS FORM") {
      var thisTime = thisResponses[0];
      var thisEmail = (thisResponses[9] || thisResponses[1] || "").toString().trim();
      var thisSession = (thisResponses[5] || thisResponses[3] || "").toString().trim();
      if (!thisSession || !thisEmail) return;
      
      var match = thisSession.match(/\[(.*)\]/);
      if (!match) return;
      var thisSessionid = match[1];
      
      var sheetSessions = ss.getSheetByName("SESSIONS");
      var remainingSeats = 1;
      if (sheetSessions) {
        var lastRowSessions = sheetSessions.getLastRow();
        if (lastRowSessions >= 2) {
          var sessionsData = sheetSessions.getRange(2, 2, lastRowSessions - 1, 12).getValues();
          for (var i = 0; i < sessionsData.length; i++) {
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

      if (!sheetInscriptions) return;
      var maxRows = sheetInscriptions.getMaxRows();
      var verif = [];
      if (maxRows > 1) {
        var sessionsEmail = sheetInscriptions.getRange(2, 2, maxRows - 1, 2).getValues();
        verif = sessionsEmail.filter(function(row) { return (row[0] === thisSessionid && row[1] === thisEmail); });
      }
      
      if (verif.length > 0) {
        Logger.log("Déjà inscrit : " + thisEmail + " à " + thisSessionid);
        return;
      }

      inscription(thisTime, thisSessionid, thisEmail);

      addParticipantToCalendar(thisSessionid, thisEmail);

      try {
        sendConfirmationMail(thisSessionid, thisEmail);
      } catch (mailErr) {
        Logger.log("Avertissement : échec de l'envoi d'e-mail : " + mailErr);
      }
      
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

function inscription(time, sessionId, email) {
  if (sheetInscriptions) {
    sheetInscriptions.appendRow([time, sessionId, email]);
  }
}

function updateFormChoices() {
  if (!sheetParametres) return;
  var formId = sheetParametres.getRange(pnParaCel["PARAMETRE_ID_FORMS_INSCRIPTION"]).getValue();
  if (!formId) {
    Logger.log("ID Forms Inscription non trouvé dans les paramètres.");
    return;
  }

  var form = FormApp.openById(formId);
  var items = form.getItems(FormApp.ItemType.LIST);
  var sessionItem = null;

  for (var i = 0; i < items.length; i++) {
    if (items[i].getTitle().indexOf("Inscription à la formation suivante") > -1 || items[i].getTitle().indexOf("Session") > -1) {
      sessionItem = items[i].asListItem();
      break;
    }
  }

  if (!sessionItem) {
    Logger.log("Question de sélection de session non trouvée dans le formulaire.");
    return;
  }

  var sheetSessions = ss.getSheetByName("SESSIONS");
  if (!sheetSessions) return;

  var lastRow = sheetSessions.getLastRow();
  if (lastRow < 2) return;

  var values = sheetSessions.getRange(2, 2, lastRow - 1, 14).getValues();
  var choices = [];

  values.forEach(function(row) {
    var sessionId = row[0];
    var publish = row[9];
    var remaining = row[11];
    var formationTitle = row[13];

    if (sessionId && publish && Number(remaining) > 0) {
      var dateObj = new Date(row[2]);
      var dateStr = dateObj.getDate() + "/" + (dateObj.getMonth() + 1) + "/" + dateObj.getFullYear();
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

function sendConfirmationMail(sessionId, email) {
  if (!sheetParametres) return;

  var urlDesinscription = sheetParametres.getRange(pnParaCel["PARAMETRE_ID_FORMS_DESINSCRIPTION"]).getValue() || "";
  
  var entrySessionId = "entry.2116080188";
  if (pnParaCel["PARAMETRE_ENTRY_SESSION"]) {
    var customEntry = sheetParametres.getRange(pnParaCel["PARAMETRE_ENTRY_SESSION"]).getValue();
    if (customEntry) entrySessionId = customEntry.toString().trim();
  }

  var entryEmailId = "entry.193822625";
  if (pnParaCel["PARAMETRE_ENTRY_EMAIL"]) {
    var customEmailEntry = sheetParametres.getRange(pnParaCel["PARAMETRE_ENTRY_EMAIL"]).getValue();
    if (customEmailEntry) entryEmailId = customEmailEntry.toString().trim();
  }

  var desinscriptionLink = "https://docs.google.com/forms/d/e/" + urlDesinscription + 
    "/viewform?usp=pp_url&" + entryEmailId + "=" + encodeURIComponent(email) + 
    "&" + entrySessionId + "=[" + encodeURIComponent(sessionId) + "]";

  var connexionInfo = sheetParametres.getRange(pnParaCel["PARAMETRE_CONNEXION_1"]).getValue() || "Lien Meet inclus dans votre invitation Agenda";

  var pdfAttachment = null;
  var pdfUrl = "";

  try {
    if (pnParaCel["PARAMETRE_ID_MODELE_CONVOC"]) {
      var modeleConvocationId = sheetParametres.getRange(pnParaCel["PARAMETRE_ID_MODELE_CONVOC"]).getValue();
      if (modeleConvocationId && modeleConvocationId.length > 10) {
        var modeleConvocation = DriveApp.getFileById(modeleConvocationId);
        var folderConvocationId = pnParaCel["PARAMETRE_ID_DOSSIER_CONVOC"] ? sheetParametres.getRange(pnParaCel["PARAMETRE_ID_DOSSIER_CONVOC"]).getValue() : "";
        var folderConvocation = folderConvocationId ? DriveApp.getFolderById(folderConvocationId) : DriveApp.getRootFolder();
        
        var convocationName = "Convocation_" + sessionId + "_" + email;
        var convocationDoc = modeleConvocation.makeCopy(convocationName, folderConvocation);
        var doc = DocumentApp.openById(convocationDoc.getId());
        var body = doc.getBody();

        body.replaceText("{{DATE AUJOURDHUI}}", Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy'))
            .replaceText("{{SESSION ID}}", sessionId)
            .replaceText("{{EMAIL}}", email);
        doc.saveAndClose();

        pdfAttachment = convocationDoc.getAs('application/pdf');
        pdfAttachment.setName(convocationName + ".pdf");
        var pdfFile = folderConvocation.createFile(pdfAttachment);
        pdfUrl = pdfFile.getUrl();

        try { convocationDoc.setTrashed(true); } catch (e) {}
      }
    }
  } catch (pdfErr) {
    Logger.log("Avertissement : la génération du PDF n'a pas pu être effectuée (envoi sans pièce jointe) : " + pdfErr);
  }

  var subject = "Convocation & Confirmation d'inscription - Formation Leroy Merlin [" + sessionId + "]";
  
  var htmlBody = "<div style='font-family: Arial, sans-serif; color: #333; max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 8px; overflow: hidden;'>"
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

  var mailOptions = {
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

function getPreFilledFormUrl(formId, entryId, selectedValue) {
  return "https://docs.google.com/forms/d/e/" + formId + "/viewform?usp=pp_url&" + entryId + "=" + encodeURIComponent(selectedValue);
}
