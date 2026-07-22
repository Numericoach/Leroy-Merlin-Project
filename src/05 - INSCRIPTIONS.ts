/**
 * Déclencheur sur soumission du formulaire d'inscription
 * Protégé par LockService contre les soumissions simultanées
 */
function onSubmit(e?: any): void {
  const lock = LockService.getScriptLock();
  const hasLock = lock.tryLock(10000);
  if (!hasLock) {
    Logger.log("Impossible d'acquérir le verrou. Soumission reportée ou ignorée.");
    return;
  }

  try {
    let thisTime: any = null;
    let thisEmail = "";
    let thisSession = "";
    let thisCivilite = "";
    let thisPrenom = "";
    let thisNom = "";
    let thisNbParticipants = 1;
    let thisSheetName = "";

    if (e && e.range) {
      try {
        thisSheetName = e.range.getSheet().getName();
        thisTime = e.range.getValues()[0][0]; // Timestamp exact de la cellule soumise par le Form
      } catch (err) {}
    }

    if (!thisTime && e && e.namedValues && e.namedValues["Horodateur"]) {
      thisTime = e.namedValues["Horodateur"][0];
    }
    if (!thisTime && e && e.values && e.values[0]) {
      thisTime = e.values[0];
    }
    if (!thisTime) {
      thisTime = new Date();
    }

    // 1. EXTRACTION INTELLIGENTE DES DONNÉES DE LA SOUMISSION
    if (e && e.namedValues) {
      for (const key in e.namedValues) {
        const keyLower = key.toLowerCase();
        const val = (e.namedValues[key][0] || "").toString().trim();
        if (!thisEmail && (keyLower.indexOf("mail") > -1 || keyLower.indexOf("courriel") > -1 || keyLower.indexOf("email") > -1)) {
          thisEmail = val;
        }
        if (!thisSession && (keyLower.indexOf("session") > -1 || keyLower.indexOf("formation") > -1)) {
          thisSession = val;
        }
        if (!thisCivilite && keyLower.indexOf("civilite") > -1) {
          thisCivilite = val;
        }
        if (!thisPrenom && keyLower.indexOf("prenom") > -1) {
          thisPrenom = val;
        }
        if (!thisNom && keyLower.indexOf("nom") > -1 && keyLower.indexOf("prenom") === -1) {
          thisNom = val;
        }
        if (keyLower.indexOf("participant") > -1 || keyLower.indexOf("nombre") > -1 || keyLower.indexOf("nb") > -1) {
          const parsed = parseInt(val, 10);
          if (!isNaN(parsed) && parsed > 0) {
            thisNbParticipants = parsed;
          }
        }
      }
    }

    if (e && e.values && Array.isArray(e.values)) {
      if (!thisEmail && e.values[1] && e.values[1].indexOf("@") > -1) thisEmail = e.values[1];
      if (!thisCivilite && e.values[2]) thisCivilite = e.values[2];
      if (!thisPrenom && e.values[3]) thisPrenom = e.values[3];
      if (!thisNom && e.values[4]) thisNom = e.values[4];

      e.values.forEach((val: any) => {
        const valStr = (val || "").toString().trim();
        if (!thisEmail && valStr.indexOf("@") > -1) {
          thisEmail = valStr;
        }
        if (!thisSession && (valStr.indexOf("[") > -1 || valStr.toLowerCase().indexOf("formation") > -1 || valStr.toLowerCase().indexOf("session") > -1)) {
          thisSession = valStr;
        }
      });
    }

    if ((!thisEmail || !thisSession) && e && e.range) {
      const rowValues = e.range.getValues()[0];
      if (rowValues && rowValues.length > 0) {
        if (!thisEmail && rowValues[1] && rowValues[1].toString().indexOf("@") > -1) thisEmail = rowValues[1].toString().trim();
        if (!thisCivilite && rowValues[2]) thisCivilite = rowValues[2].toString().trim();
        if (!thisPrenom && rowValues[3]) thisPrenom = rowValues[3].toString().trim();
        if (!thisNom && rowValues[4]) thisNom = rowValues[4].toString().trim();

        rowValues.forEach((val: any) => {
          const valStr = (val || "").toString().trim();
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

    // EXTRACTION SÉCURISÉE DE L'ID DE SESSION
    let thisSessionid = thisSession;
    const match = thisSession.match(/\[(.*?)\]/);
    if (match && match[1]) {
      thisSessionid = match[1].trim();
    }

    // 2. VÉRIFIER LES PLACES RESTANTES SOUS VERROU AVEC LE NOMBRE DE PARTICIPANTS
    const sheetSessions = ss ? ss.getSheetByName("SESSIONS") : null;
    let remainingSeats = 1;
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

    if (remainingSeats < thisNbParticipants) {
      Logger.log("Inscription refusée : pas assez de places disponibles (" + remainingSeats + " restantes, " + thisNbParticipants + " demandées) pour " + thisSessionid);
      return;
    }

    // 3. VÉRIFIER SI DÉJÀ INSCRIT SOUS VERROU
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

    // 4. INSCRIPTION DANS LE SHEETS AVEC L'HORODATEUR EXACT DU FORMULAIRE
    inscription(thisTime, thisSessionid, thisEmail);

    // 5. FONCTION PRINCIPALE : AJOUT DANS GOOGLE AGENDA (Priorité absolue)
    addParticipantToCalendar(thisSessionid, thisEmail);

    // 6. FONCTION SECONDAIRE : ENVOI DE LA CONVOCATION / CONFIRMATION (Découplé)
    try {
      sendConfirmationMail(thisSessionid, thisEmail, thisPrenom, thisNom, thisCivilite, thisNbParticipants);
    } catch (mailErr) {
      Logger.log("Avertissement : échec de l'envoi d'e-mail (n'impacte pas l'inscription ni l'agenda) : " + mailErr);
    }
    
    // 7. MISE À JOUR DYNAMIQUE DES CHOIX DU FORMULAIRE
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

/**
 * Enregistrer l'inscription dans la feuille de calcul avec le timestamp exact du formulaire
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
  const formId = getParamValue("PARAMETRE_ID_FORMS_INSCRIPTION");
  if (!formId) {
    Logger.log("ID Forms Inscription non trouvé dans les paramètres.");
    return;
  }

  const form = FormApp.openById(formId);
  const items = form.getItems(FormApp.ItemType.LIST);
  let sessionItem: GoogleAppsScript.Forms.ListItem | null = null;

  for (let i = 0; i < items.length; i++) {
    const title = items[i].getTitle();
    if (title.indexOf("Inscription à la formation suivante") > -1 || title.indexOf("Session") > -1 || title.indexOf("formation") > -1) {
      sessionItem = items[i].asListItem();
      break;
    }
  }

  if (!sessionItem) {
    Logger.log("Question de sélection de session non trouvée dans le formulaire.");
    return;
  }

  const sheetSessions = ss ? ss.getSheetByName("SESSIONS") : null;
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
 * Envoi de l'e-mail de convocation / confirmation personnalisé avec le PDF complet
 */
function sendConfirmationMail(sessionId: string, email: string, prenom?: string, nom?: string, civilite?: string, nbParticipants?: number): void {
  const urlDesinscription = getParamValue("PARAMETRE_ID_FORMS_DESINSCRIPTION");
  
  let entrySessionId = "entry.2116080188";
  const customEntry = getParamValue("PARAMETRE_ENTRY_SESSION");
  if (customEntry) entrySessionId = customEntry;

  let entryEmailId = "entry.193822625";
  const customEmailEntry = getParamValue("PARAMETRE_ENTRY_EMAIL");
  if (customEmailEntry) entryEmailId = customEmailEntry;

  const desinscriptionLink = "https://docs.google.com/forms/d/e/" + urlDesinscription + 
    "/viewform?usp=pp_url&" + entryEmailId + "=" + encodeURIComponent(email) + 
    "&" + entrySessionId + "=[" + encodeURIComponent(sessionId) + "]";

  const connexionInfo = getParamValue("PARAMETRE_CONNEXION_1") || "Lien Meet inclus dans votre invitation Agenda";

  // Récupérer les détails de la session depuis SESSIONS
  let formationTitle = "Formation Leroy Merlin";
  let dateStr = "";
  let heureDebutStr = "";
  let heureFinStr = "";
  let lieuStr = "";
  let descriptionStr = "";

  const sheetSessions = ss ? ss.getSheetByName("SESSIONS") : null;
  if (sheetSessions) {
    const lastRow = sheetSessions.getLastRow();
    if (lastRow >= 2) {
      const sessionsValues = sheetSessions.getRange(2, 2, lastRow - 1, 14).getValues();
      for (let i = 0; i < sessionsValues.length; i++) {
        if (sessionsValues[i][0] === sessionId) {
          formationTitle = sessionsValues[i][13] || formationTitle;
          if (sessionsValues[i][2]) {
            const d = new Date(sessionsValues[i][2]);
            dateStr = d.getDate() + "/" + (d.getMonth() + 1) + "/" + d.getFullYear();
          }
          if (sessionsValues[i][3]) {
            const hd = new Date(sessionsValues[i][3]);
            heureDebutStr = hd.getHours() + "h" + (hd.getMinutes() < 10 ? "0" : "") + hd.getMinutes();
          }
          if (sessionsValues[i][4]) {
            const hf = new Date(sessionsValues[i][4]);
            heureFinStr = hf.getHours() + "h" + (hf.getMinutes() < 10 ? "0" : "") + hf.getMinutes();
          }
          lieuStr = sessionsValues[i][7] || "";
          break;
        }
      }
    }
  }

  // Optionnel : Générer le PDF de convocation si l'ID du modèle est renseigné dans PARAMETRES
  let pdfAttachment: GoogleAppsScript.Base.Blob | null = null;
  let pdfUrl = "";

  try {
    const modeleConvocationId = getParamValue("PARAMETRE_ID_MODELE_CONVOC");
    if (modeleConvocationId && modeleConvocationId.length > 10) {
      const modeleConvocation = DriveApp.getFileById(modeleConvocationId);
      const folderConvocationId = getParamValue("PARAMETRE_ID_DOSSIER_CONVOC");
      const folderConvocation = folderConvocationId ? DriveApp.getFolderById(folderConvocationId) : DriveApp.getRootFolder();
      
      const convocationName = "Convocation_" + sessionId + "_" + email;
      const convocationDoc = modeleConvocation.makeCopy(convocationName, folderConvocation);
      const doc = DocumentApp.openById(convocationDoc.getId());
      const body = doc.getBody();

      const nbPartStr = (nbParticipants || 1).toString();

      body.replaceText("{{DATE AUJOURDHUI}}", Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy'));
      body.replaceText("{{SESSION ID}}", sessionId);
      body.replaceText("{{EMAIL}}", email);
      body.replaceText("{{CIVILITE}}", civilite || "");
      body.replaceText("{{PRENOM}}", prenom || "");
      body.replaceText("{{NOM}}", nom || "");
      body.replaceText("{{TITRE FORMATION}}", formationTitle);
      body.replaceText("{{DATE}}", dateStr);
      body.replaceText("{{HEURE DEBUT}}", heureDebutStr);
      body.replaceText("{{HEURE FIN}}", heureFinStr);
      body.replaceText("{{LIEU}}", lieuStr);
      body.replaceText("{{CONNEXION}}", connexionInfo);
      body.replaceText("{{DESCRIPTION}}", descriptionStr);
      body.replaceText("{{APPLI}}", formationTitle);
      body.replaceText("{{NB PARTICIPANTS}}", nbPartStr);
      body.replaceText("{{PARTICIPANTS}}", nbPartStr);

      doc.saveAndClose();

      pdfAttachment = convocationDoc.getAs('application/pdf');
      pdfAttachment.setName(convocationName + ".pdf");
      const pdfFile = folderConvocation.createFile(pdfAttachment);
      pdfUrl = pdfFile.getUrl();

      try { convocationDoc.setTrashed(true); } catch (e) {}
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
    + "<p>Bonjour " + (prenom ? prenom + " " + (nom || "") : "") + ",</p>"
    + "<p>Votre inscription pour <b>" + (nbParticipants || 1) + " participant(s)</b> à la session de formation <b>" + formationTitle + " [" + sessionId + "]</b> a bien été confirmée.</p>"
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
