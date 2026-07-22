/**
 * Déclencheur sur soumission du formulaire d'inscription
 * Protégé par LockService contre les soumissions simultanées
 */
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
      
      // 1. VÉRIFIER LES PLACES RESTANTES SOUS VERROU
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

      // 2. VÉRIFIER SI DÉJÀ INSCRIT SOUS VERROU
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

      // 3. INSCRIPTION DANS LE SHEETS
      inscription(thisTime, thisSessionid, thisEmail);

      // 4. FONCTION PRINCIPALE : AJOUT DANS GOOGLE AGENDA (Priorité absolue)
      addParticipantToCalendar(thisSessionid, thisEmail);

      // 5. FONCTION SECONDAIRE : ENVOI DE L'E-MAIL DE CONFIRMATION (Découplé)
      try {
        sendConfirmationMail(thisSessionid, thisEmail);
      } catch (mailErr) {
        Logger.log("Avertissement : échec de l'envoi d'e-mail : " + mailErr);
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

  var subject = "Confirmation d'inscription à votre session de formation Leroy Merlin";
  var htmlBody = "<h3>Bonjour,</h3>"
    + "<p>Votre inscription à la session de formation <b>[" + sessionId + "]</b> a bien été enregistrée.</p>"
    + "<p>Une invitation Google Agenda vous a été envoyée avec le lien de connexion et le programme.</p>"
    + "<p><a href='" + desinscriptionLink + "' style='color:#CC3C25;'>Cliquer ici pour vous désinscrire</a></p>"
    + "<br><p>Cordialement,<br>L’équipe Numericoach</p>";

  MailApp.sendEmail({
    to: email,
    subject: subject,
    htmlBody: htmlBody
  });
}

function getPreFilledFormUrl(formId, entryId, selectedValue) {
  return "https://docs.google.com/forms/d/e/" + formId + "/viewform?usp=pp_url&" + entryId + "=" + encodeURIComponent(selectedValue);
}
