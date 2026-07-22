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

      // 5. FONCTION SECONDAIRE : ENVOI DE L'E-MAIL DE CONFIRMATION (Découplé)
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
 * Envoi d'un e-mail de confirmation léger sans dépendance aux PDF
 */
function sendConfirmationMail(sessionId: string, email: string): void {
  if (!sheetParametres) return;

  const urlDesinscription = sheetParametres.getRange(pnParaCel["PARAMETRE_ID_FORMS_DESINSCRIPTION"]).getValue() || "";
  
  // Utiliser l'ID de champ réel si configuré dans PARAMETRES
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

  const subject = "Confirmation d'inscription à votre session de formation Leroy Merlin";
  const htmlBody = "<h3>Bonjour,</h3>"
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

/**
 * Helper pour générer l'URL d'un formulaire pré-rempli avec l'entrée réelle
 */
function getPreFilledFormUrl(formId: string, entryId: string, selectedValue: string): string {
  return "https://docs.google.com/forms/d/e/" + formId + "/viewform?usp=pp_url&" + entryId + "=" + encodeURIComponent(selectedValue);
}
