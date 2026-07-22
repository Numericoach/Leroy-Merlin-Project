function onSubmit(e) {
  var lock = LockService.getScriptLock();
  var hasLock = lock.tryLock(10000);
  if (!hasLock) {
    Logger.log("Impossible d'acquérir le verrou. Soumission reportée ou ignorée.");
    return;
  }

  try {
    var thisTime = new Date();
    var thisEmail = "";
    var thisSession = "";
    var thisSheetName = "";

    if (e && e.range) {
      try {
        thisSheetName = e.range.getSheet().getName();
      } catch (err) {}
    }

    if (e && e.namedValues) {
      for (var key in e.namedValues) {
        var keyLower = key.toLowerCase();
        if (!thisEmail && (keyLower.indexOf("mail") > -1 || keyLower.indexOf("courriel") > -1 || keyLower.indexOf("email") > -1)) {
          thisEmail = (e.namedValues[key][0] || "").toString().trim();
        }
        if (!thisSession && (keyLower.indexOf("session") > -1 || keyLower.indexOf("formation") > -1)) {
          thisSession = (e.namedValues[key][0] || "").toString().trim();
        }
      }
    }

    if (e && e.values && Array.isArray(e.values)) {
      if (!thisTime) thisTime = e.values[0];
      e.values.forEach(function(val) {
        var valStr = (val || "").toString().trim();
        if (!thisEmail && valStr.indexOf("@") > -1) {
          thisEmail = valStr;
        }
        if (!thisSession && (valStr.indexOf("[") > -1 || valStr.toLowerCase().indexOf("formation") > -1 || valStr.toLowerCase().indexOf("session") > -1)) {
          thisSession = valStr;
        }
      });
    }

    if ((!thisEmail || !thisSession) && e && e.range) {
      var rowValues = e.range.getValues()[0];
      if (rowValues && rowValues.length > 0) {
        if (!thisTime) thisTime = rowValues[0];
        rowValues.forEach(function(val) {
          var valStr = (val || "").toString().trim();
          if (!thisEmail && valStr.indexOf("@") > -1) {
            thisEmail = valStr;
          }
          if (!thisSession && (valStr.indexOf("[") > -1 || valStr.toLowerCase().indexOf("formation") > -1 || valStr.toLowerCase().indexOf("session") > -1)) {
            thisSession = valStr;
          }
        });
      }
    }

    if (!thisSession || !thisEmail) {
      Logger.log("Données manquantes (Email: " + thisEmail + ", Session: " + thisSession + ")");
      return;
    }

    var thisSessionid = thisSession;
    var match = thisSession.match(/\[(.*?)\]/);
    if (match && match[1]) {
      thisSessionid = match[1].trim();
    }

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
  var formId = getParamValue("PARAMETRE_ID_FORMS_INSCRIPTION");
  if (!formId) {
    Logger.log("ID Forms Inscription non trouvé dans les paramètres.");
    return;
  }

  var form = FormApp.openById(formId);
  var items = form.getItems(FormApp.ItemType.LIST);
  var sessionItem = null;

  for (var i = 0; i < items.length; i++) {
    var title = items[i].getTitle();
    if (title.indexOf("Inscription à la formation suivante") > -1 || title.indexOf("Session") > -1 || title.indexOf("formation") > -1) {
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
  var urlDesinscription = getParamValue("PARAMETRE_ID_FORMS_DESINSCRIPTION");
  
  var entrySessionId = "entry.2116080188";
  var customEntry = getParamValue("PARAMETRE_ENTRY_SESSION");
  if (customEntry) entrySessionId = customEntry;

  var entryEmailId = "entry.193822625";
  var customEmailEntry = getParamValue("PARAMETRE_ENTRY_EMAIL");
  if (customEmailEntry) entryEmailId = customEmailEntry;

  var desinscriptionLink = "https://docs.google.com/forms/d/e/" + urlDesinscription + 
    "/viewform?usp=pp_url&" + entryEmailId + "=" + encodeURIComponent(email) + 
    "&" + entrySessionId + "=[" + encodeURIComponent(sessionId) + "]";

  var connexionInfo = getParamValue("PARAMETRE_CONNEXION_1") || "Lien Meet inclus dans votre invitation Agenda";

  var pdfAttachment = null;
  var pdfUrl = "";

  try {
    var modeleConvocationId = getParamValue("PARAMETRE_ID_MODELE_CONVOC");
    if (modeleConvocationId && modeleConvocationId.length > 10) {
      var modeleConvocation = DriveApp.getFileById(modeleConvocationId);
      var folderConvocationId = getParamValue("PARAMETRE_ID_DOSSIER_CONVOC");
      var folderConvocation = folderConvocationId ? DriveApp.getFolderById(folderConvocationId) : DriveApp.getRootFolder();
      
      var convocationName = "Convocation_" + sessionId + "_" + email;
      var convocationDoc = modeleConvocation.makeCopy(convocationName, folderConvocation);
      var doc = DocumentApp.openById(convocationDoc.getId());
      var body = doc.getBody();

      body.replaceText("{{DATE AUJOURDHUI}}", Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy'));
      body.replaceText("{{SESSION ID}}", sessionId);
      body.replaceText("{{EMAIL}}", email);
      doc.saveAndClose();

      pdfAttachment = convocationDoc.getAs('application/pdf');
      pdfAttachment.setName(convocationName + ".pdf");
      var pdfFile = folderConvocation.createFile(pdfAttachment);
      pdfUrl = pdfFile.getUrl();

      try { convocationDoc.setTrashed(true); } catch (e) {}
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
    + "<p>Votre inscription à la session de formation <b>[" + sessionId + "]</b> a bien été enregistrée.</p>"
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
