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
    const thisNbParticipants = 1; // 1 Inscription = 1 Entreprise
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

    // 1. EXTRACTION INTELLIGENTE ET RIGOUREUSE DES DONNÉES DE LA SOUMISSION
    if (e && e.namedValues) {
      for (const key in e.namedValues) {
        const keyLower = key.toLowerCase();
        const val = (e.namedValues[key][0] || "").toString().trim();
        
        if (!thisEmail && (keyLower.indexOf("mail") > -1 || keyLower.indexOf("courriel") > -1 || keyLower.indexOf("email") > -1)) {
          thisEmail = val;
        }
        
        if (val.indexOf("[") > -1 || val.indexOf("SES-") > -1 || keyLower.indexOf("inscription à la formation") > -1 || keyLower.indexOf("session") > -1) {
          if (val.indexOf("[") > -1 || val.indexOf("SES-") > -1 || !thisSession) {
            thisSession = val;
          }
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
        if (valStr.indexOf("[") > -1 || valStr.indexOf("SES-") > -1) {
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
          if (valStr.indexOf("[") > -1 || valStr.indexOf("SES-") > -1) {
            thisSession = valStr;
          }
        });
      }
    }

    if (!thisSession || !thisEmail) {
      Logger.log("Données manquantes (Email: " + thisEmail + ", Session: " + thisSession + ")");
      return;
    }

    // EXTRACTION PRÉCISE DE L'ID DE SESSION (ex: SES-0002)
    let thisSessionid = "";
    const matchBracket = thisSession.match(/\[(.*?)\]/);
    if (matchBracket && matchBracket[1]) {
      thisSessionid = matchBracket[1].trim();
    } else {
      const matchSes = thisSession.match(/(SES-\d+)/i);
      if (matchSes && matchSes[1]) {
        thisSessionid = matchSes[1].trim();
      } else {
        thisSessionid = thisSession.trim();
      }
    }

    // 2. VÉRIFIER LES PLACES RESTANTES DANS L'ONGLET SESSIONS (APRÈS AJOUT PAR GOOGLE FORM)
    const sheetSessions = ss ? ss.getSheetByName("SESSIONS") : null;
    let remainingSeatsAfterForm = 0;
    let sessionFound = false;

    if (sheetSessions) {
      const lastRowSessions = sheetSessions.getLastRow();
      if (lastRowSessions >= 2) {
        const sessionsData = sheetSessions.getRange(2, 2, lastRowSessions - 1, 13).getValues();
        for (let i = 0; i < sessionsData.length; i++) {
          const idInSheet = (sessionsData[i][0] || "").toString().trim();
          if (idInSheet === thisSessionid) {
            remainingSeatsAfterForm = Number(sessionsData[i][11]); // Colonne M (Places Restantes)
            sessionFound = true;
            break;
          }
        }

        // Recherche par comparaison si l'ID exact n'est pas dans des crochets
        if (!sessionFound) {
          const searchLower = thisSession.toLowerCase();
          for (let i = 0; i < sessionsData.length; i++) {
            const idInSheet = (sessionsData[i][0] || "").toString().trim();
            const formationInSheet = (sessionsData[i][1] || "").toString().trim().toLowerCase();
            const infoCompInSheet = (sessionsData[i][8] || "").toString().trim().toLowerCase();

            if (idInSheet && (
              searchLower.indexOf(idInSheet.toLowerCase()) > -1 ||
              (infoCompInSheet && searchLower.indexOf("avec pratique") > -1 && infoCompInSheet.indexOf("avec pratique") > -1) ||
              (infoCompInSheet && searchLower.indexOf("sans pratique") > -1 && infoCompInSheet.indexOf("sans pratique") > -1) ||
              (formationInSheet && searchLower.indexOf(formationInSheet) > -1)
            )) {
              thisSessionid = idInSheet;
              remainingSeatsAfterForm = Number(sessionsData[i][11]);
              sessionFound = true;
              break;
            }
          }
        }
      }
    }

    if (!sessionFound) {
      Logger.log("Avertissement : ID Session " + thisSessionid + " non trouvé dans SESSIONS. Passage par défaut.");
      remainingSeatsAfterForm = 0;
    }

    // Si remainingSeatsAfterForm < 0, cela signifie que la session est déjà complète !
    if (remainingSeatsAfterForm < 0) {
      Logger.log("Inscription refusée : session complète pour " + thisSessionid);
      
      // SUPPRIMER LA LIGNE EN TROP DE INSCRIPTIONSS POUR CONSERVER UN COMPTEUR PROPRE
      if (e && e.range) {
        try {
          e.range.getSheet().deleteRow(e.range.getRow());
          Logger.log("Ligne d'inscription refusée supprimée de " + e.range.getSheet().getName() + " à la ligne " + e.range.getRow());
        } catch (delErr) {
          Logger.log("Erreur lors de la suppression de la ligne refusée : " + delErr);
        }
      }

      try {
        sendWaitingListMail(thisSessionid, thisEmail, thisPrenom, thisNom);
      } catch (waitErr) {
        Logger.log("Erreur lors de l'envoi de l'e-mail de liste d'attente : " + waitErr);
      }
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
      sendConfirmationMail(thisSessionid, thisEmail, thisPrenom, thisNom, thisCivilite);
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
    try {
      updateFormChoices();
    } catch (fErr) {
      Logger.log("Erreur dans updateFormChoices (finally) : " + fErr);
    }
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
 * Helper pour extraire l'ID propre d'un formulaire Google Forms même si l'utilisateur a collé l'URL entière
 */
function extractFormId(input: string): string {
  if (!input) return "";
  const cleaned = input.trim();
  const match = cleaned.match(/\/d\/(?:e\/)?([a-zA-Z0-9_-]+)/);
  if (match && match[1]) {
    return match[1];
  }
  return cleaned;
}

/**
 * Helper pour formater proprement une date (string, Date, ou timestamp) en DD/MM/YYYY
 */
function formatDateClean(dateVal: any): string {
  if (!dateVal) return "";
  if (dateVal instanceof Date && !isNaN(dateVal.getTime())) {
    const d = dateVal.getDate();
    const m = dateVal.getMonth() + 1;
    const y = dateVal.getFullYear();
    return (d < 10 ? "0" : "") + d + "/" + (m < 10 ? "0" : "") + m + "/" + y;
  }
  const str = dateVal.toString().trim();
  if (str.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/)) {
    return str;
  }
  const parsed = new Date(str);
  if (!isNaN(parsed.getTime())) {
    const d = parsed.getDate();
    const m = parsed.getMonth() + 1;
    const y = parsed.getFullYear();
    return (d < 10 ? "0" : "") + d + "/" + (m < 10 ? "0" : "") + m + "/" + y;
  }
  return str;
}

/**
 * Helper pour formater proprement une heure en XhXX
 */
function formatTimeClean(timeVal: any): string {
  if (!timeVal) return "";
  if (timeVal instanceof Date && !isNaN(timeVal.getTime())) {
    const h = timeVal.getHours();
    const m = timeVal.getMinutes();
    return h + "h" + (m > 0 ? (m < 10 ? "0" : "") + m : "00");
  }
  const str = timeVal.toString().trim();
  if (str.indexOf("h") > -1) return str;
  if (str.indexOf(":") > -1) {
    const parts = str.split(":");
    const h = parseInt(parts[0], 10);
    const m = parseInt(parts[1], 10);
    if (!isNaN(h)) {
      return h + "h" + (!isNaN(m) && m > 0 ? (m < 10 ? "0" : "") + m : "00");
    }
  }
  const parsed = new Date(str);
  if (!isNaN(parsed.getTime())) {
    const h = parsed.getHours();
    const m = parsed.getMinutes();
    return h + "h" + (m > 0 ? (m < 10 ? "0" : "") + m : "00");
  }
  return str;
}

/**
 * Mettre à jour dynamiquement la liste des sessions disponibles dans le Google Form (Liste déroulante ou Choix multiple)
 */
function updateFormChoices(): void {
  try {
    let rawFormId = getParamValue("PARAMETRE_ID_EDITION");
    if (!rawFormId || rawFormId.trim() === "") {
      rawFormId = getParamValue("PARAMETRE_ID_FORMS_INSCRIPTION");
    }
    const formId = extractFormId(rawFormId);
    if (!formId) {
      Logger.log("ID Forms Inscription non trouvé dans les paramètres.");
      if (ss) ss.toast("❌ ID Formulaire non trouvé dans l'onglet PARAMETRES.", "NUMERICOACH", 6);
      return;
    }

    let form: GoogleAppsScript.Forms.Form;
    try {
      form = FormApp.openById(formId);
    } catch (openErr) {
      Logger.log("Avertissement : impossible d'ouvrir le Google Form avec l'ID '" + formId + "' : " + openErr);
      if (ss) ss.toast("❌ Impossible d'ouvrir le Formulaire. Vérifiez l'ID d'édition (" + formId + ").", "NUMERICOACH", 7);
      return;
    }

    const items = form.getItems();
    const targetItems: (GoogleAppsScript.Forms.ListItem | GoogleAppsScript.Forms.MultipleChoiceItem)[] = [];

    for (let i = 0; i < items.length; i++) {
      const title = items[i].getTitle().toLowerCase();
      const type = items[i].getType();
      if (type === FormApp.ItemType.LIST || type === FormApp.ItemType.MULTIPLE_CHOICE) {
        if (
          title.indexOf("inscription") > -1 ||
          title.indexOf("session") > -1 ||
          title.indexOf("formation") > -1 ||
          title.indexOf("créneau") > -1 ||
          title.indexOf("creneau") > -1 ||
          title.indexOf("choix") > -1 ||
          title.indexOf("date") > -1
        ) {
          if (type === FormApp.ItemType.LIST) targetItems.push(items[i].asListItem());
          else targetItems.push(items[i].asMultipleChoiceItem());
        }
      }
    }

    // Fallback: si aucun item ciblé par mot-clé, prendre tous les éléments LIST / MULTIPLE_CHOICE du formulaire
    if (targetItems.length === 0) {
      for (let i = 0; i < items.length; i++) {
        const type = items[i].getType();
        if (type === FormApp.ItemType.LIST) targetItems.push(items[i].asListItem());
        else if (type === FormApp.ItemType.MULTIPLE_CHOICE) targetItems.push(items[i].asMultipleChoiceItem());
      }
    }

    if (targetItems.length === 0) {
      Logger.log("Aucun élément de type Liste déroulante ou Choix multiple trouvé dans le formulaire.");
      if (ss) ss.toast("❌ Aucune question de choix trouvée dans le Google Form.", "NUMERICOACH", 6);
      return;
    }

    const sheetSessions = ss ? ss.getSheetByName("SESSIONS") : null;
    if (!sheetSessions) return;

    const lastRow = sheetSessions.getLastRow();
    if (lastRow < 2) return;

    // Lecture depuis la colonne A (1) jusqu'à la colonne 20 (Col P) pour éviter tout décalage
    const values = sheetSessions.getRange(2, 1, lastRow - 1, 20).getValues();
    const choices: string[] = [];
    const addedSessionIds: string[] = [];

    values.forEach(function(row) {
      const sessionId = (row[1] || "").toString().trim(); // Col B (ID SESSION)
      const formationCol = (row[2] || "").toString().trim(); // Col C (FORMATION)
      const dateVal = row[3]; // Col D (DATE)
      const heureDebutVal = row[4]; // Col E (HEURE DEBUT)
      const heureFinVal = row[5]; // Col F (HEURE FIN)
      const infoComp = (row[9] || "").toString().trim(); // Col J (INFORMATION COMPLEMENTAIRE)
      const publish = row[10]; // Col K (Publier)
      const rawRemaining = parseFloat(String(row[12] || "").replace(",", ".")); // Col M (Places Restantes)
      const nbPlaces = parseFloat(String(row[7] || "").replace(",", ".")); // Col H (NB DE PLACES)
      const nbInscrits = parseFloat(String(row[11] || "").replace(",", ".")); // Col L (Nb Inscrits)
      const moduleTitle = row[14] || formationCol || "Formation"; // Col O (MODULE TITRE)

      if (!sessionId || sessionId.toLowerCase().indexOf("ses-") === -1) {
        return;
      }

      // Est publié si coché ou non-vide / non-false
      const isPublished = Boolean(publish) && 
                        String(publish).toUpperCase() !== "FALSE" && 
                        String(publish).toUpperCase() !== "FAUX" && 
                        String(publish) !== "0" && 
                        String(publish).trim() !== "";

      // Calcul robuste des places restantes : utilise Col M ou (Col H - Col L)
      let remainingSeats = 0;
      if (!isNaN(rawRemaining)) {
        remainingSeats = rawRemaining;
      } else if (!isNaN(nbPlaces)) {
        const registered = !isNaN(nbInscrits) ? nbInscrits : 0;
        remainingSeats = nbPlaces - registered;
      } else {
        remainingSeats = 1; // Fallback par défaut si non renseigné
      }

      const hasSeats = remainingSeats > 0;

      if (sessionId && isPublished && hasSeats) {
        const dateStr = formatDateClean(dateVal);
        const heureDebutStr = formatTimeClean(heureDebutVal);
        const heureFinStr = formatTimeClean(heureFinVal);

        let compClean = infoComp;
        if (compClean.toLowerCase().indexOf("avec pratique") > -1) {
          compClean = "(AVEC PRATIQUE)";
        } else if (compClean.toLowerCase().indexOf("sans pratique") > -1) {
          compClean = "(SANS PRATIQUE)";
        } else if (compClean.length > 0) {
          compClean = "(" + compClean + ")";
        }

        let titleClean = moduleTitle.toString().toUpperCase();
        if (titleClean.indexOf("[") > -1) {
          titleClean = titleClean.split("[")[0].trim();
        }

        let label = titleClean;
        if (compClean) {
          label += " " + compClean;
        }
        if (dateStr) {
          label += " - " + dateStr;
        }
        if (heureDebutStr && heureFinStr) {
          label += " de " + heureDebutStr + " à " + heureFinStr;
        }
        label += " [" + sessionId + "]";

        choices.push(label);
        addedSessionIds.push(sessionId);
      }
    });

    if (choices.length > 0) {
      targetItems.forEach(item => {
        if (item.getType() === FormApp.ItemType.LIST) {
          (item as GoogleAppsScript.Forms.ListItem).setChoiceValues(choices);
        } else if (item.getType() === FormApp.ItemType.MULTIPLE_CHOICE) {
          (item as GoogleAppsScript.Forms.MultipleChoiceItem).setChoiceValues(choices);
        }
      });
      Logger.log("Formulaire mis à jour avec " + choices.length + " sessions : " + addedSessionIds.join(", "));
      if (ss) ss.toast("✅ Formulaire mis à jour avec " + choices.length + " session(s) (" + addedSessionIds.join(", ") + ")", "NUMERICOACH", 5);
    } else {
      const defaultMsg = ["Aucune session disponible pour le moment"];
      targetItems.forEach(item => {
        if (item.getType() === FormApp.ItemType.LIST) {
          (item as GoogleAppsScript.Forms.ListItem).setChoiceValues(defaultMsg);
        } else if (item.getType() === FormApp.ItemType.MULTIPLE_CHOICE) {
          (item as GoogleAppsScript.Forms.MultipleChoiceItem).setChoiceValues(defaultMsg);
        }
      });
      Logger.log("Aucune session disponible.");
      if (ss) ss.toast("⚠️ Aucune session disponible. Formulaire réinitialisé.", "NUMERICOACH", 5);
    }
  } catch (err) {
    Logger.log("Erreur dans updateFormChoices : " + err);
    if (ss) ss.toast("❌ Erreur lors de la mise à jour : " + err, "NUMERICOACH", 7);
  }
}

/**
 * Envoi d'un e-mail d'information et d'inscription sur Liste d'Attente lorsque la session est complète
 */
function sendWaitingListMail(sessionId: string, email: string, prenom?: string, nom?: string): void {
  const formListeAttenteId = extractFormId(getParamValue("PARAMETRE_ID_FORMS_LISTE_ATTENTE"));
  const senderEmail = getParamValue("PARAMETRE_EXPEDITEUR_EMAIL");
  const senderName = getParamValue("PARAMETRE_NOM_EXPEDITEUR") || "Formations Leroy Merlin";
  
  let entrySessionId = "entry.2116080188";
  const customEntry = getParamValue("PARAMETRE_ENTRY_SESSION");
  if (customEntry) entrySessionId = customEntry;

  let entryEmailId = "entry.193822625";
  const customEmailEntry = getParamValue("PARAMETRE_ENTRY_EMAIL");
  if (customEmailEntry) entryEmailId = customEmailEntry;

  let listeAttenteLink = "#";
  if (formListeAttenteId) {
    listeAttenteLink = "https://docs.google.com/forms/d/e/" + formListeAttenteId + 
      "/viewform?usp=pp_url&" + entryEmailId + "=" + encodeURIComponent(email) + 
      "&" + entrySessionId + "=[" + encodeURIComponent(sessionId) + "]";
  }

  const subject = "Session complète - Option Liste d'Attente - Formation Leroy Merlin [" + sessionId + "]";
  
  let htmlBody = "<div style='font-family: Arial, sans-serif; color: #333; max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 8px; overflow: hidden;'>"
    + "<div style='background-color: #78BE20; padding: 20px; text-align: center; color: white;'>"
    + "<h2 style='margin: 0; font-weight: bold;'>Session complète - liste d'attente</h2>"
    + "</div>"
    + "<div style='padding: 24px;'>"
    + "<p>Bonjour " + (prenom ? prenom + " " + (nom || "") : "") + ",</p>"
    + "<p>Nous avons bien reçu votre demande d'inscription à la session <b>[" + sessionId + "]</b>.</p>"
    + "<p style='color: #C0392B;'><b>Information importante :</b> Cette session est actuellement complète.</p>"
    + "<p>Afin de ne pas rater les prochaines disponibilités ou une place libérée, vous pouvez vous inscrire sur notre <b>liste d'attente</b> :</p>";

  if (formListeAttenteId) {
    htmlBody += "<p style='text-align: center; margin: 25px 0;'><a href='" + listeAttenteLink + "' style='display:inline-block; background-color:#78BE20; color:white; padding:12px 22px; text-decoration:none; border-radius:5px; font-weight:bold;'>Rejoindre la liste d'attente</a></p>";
  } else {
    htmlBody += "<p><i>Vous serez recontacté dès qu’une nouvelle session sera ouverte.</i></p>";
  }

  htmlBody += "</div></div>";

  const mailOptions: any = {
    to: email,
    subject: subject,
    htmlBody: htmlBody,
    name: senderName
  };

  if (senderEmail && senderEmail.length > 3) {
    mailOptions.from = senderEmail;
    mailOptions.replyTo = senderEmail;
  }

  try {
    MailApp.sendEmail(mailOptions);
  } catch (err) {
    Logger.log("Avertissement expéditeur personnalisé : " + err + ". Tentative d'envoi avec l'expéditeur par défaut...");
    delete mailOptions.from;
    MailApp.sendEmail(mailOptions);
  }
  Logger.log("Mail de liste d'attente envoyé à " + email + " pour la session " + sessionId);
}

/**
 * Envoi de l'e-mail de convocation / confirmation personnalisé avec le PDF complet
 */
function sendConfirmationMail(sessionId: string, email: string, prenom?: string, nom?: string, civilite?: string): void {
  const urlDesinscription = extractFormId(getParamValue("PARAMETRE_ID_FORMS_DESINSCRIPTION"));
  const senderEmail = getParamValue("PARAMETRE_EXPEDITEUR_EMAIL");
  const senderName = getParamValue("PARAMETRE_NOM_EXPEDITEUR") || "Formations Leroy Merlin";
  
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
      body.replaceText("{{NB PARTICIPANTS}}", "1");
      body.replaceText("{{PARTICIPANTS}}", "1");

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

  const subject = "Convocation & confirmation d'inscription - Formation Leroy Merlin [" + sessionId + "]";
  
  let htmlBody = "<div style='font-family: Arial, sans-serif; color: #333; max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 8px; overflow: hidden;'>"
    + "<div style='background-color: #78BE20; padding: 20px; text-align: center; color: white;'>"
    + "<h2 style='margin: 0; font-weight: bold;'>Confirmation & convocation de formation</h2>"
    + "</div>"
    + "<div style='padding: 24px;'>"
    + "<p>Bonjour " + (prenom ? prenom + " " + (nom || "") : "") + ",</p>"
    + "<p>Votre inscription à la session de formation <b>" + formationTitle + " [" + sessionId + "]</b> a bien été confirmée.</p>"
    + "<p><b>Invitation Agenda :</b> Une invitation Google Agenda contenant la date, l'heure et le lien de connexion vous a été envoyée.</p>"
    + "<div style='background-color: #f4f7f6; padding: 15px; border-radius: 6px; margin: 15px 0;'>"
    + "<b>Informations de connexion :</b><br>" + connexionInfo
    + "</div>";

  if (pdfUrl !== "") {
    htmlBody += "<p><a href='" + pdfUrl + "' style='display:inline-block; background-color:#78BE20; color:white; padding:10px 18px; text-decoration:none; border-radius:5px; font-weight:bold;'>Télécharger votre Convocation PDF</a></p>";
  }

  htmlBody += "<p style='margin-top: 25px;'><a href='" + desinscriptionLink + "' style='color:#CC3C25;'>Demander une désinscription</a></p>"
    + "</div></div>";

  const mailOptions: any = {
    to: email,
    subject: subject,
    htmlBody: htmlBody,
    name: senderName
  };

  if (senderEmail && senderEmail.length > 3) {
    mailOptions.from = senderEmail;
    mailOptions.replyTo = senderEmail;
  }

  if (pdfAttachment) {
    mailOptions.attachments = [pdfAttachment];
  }

  try {
    MailApp.sendEmail(mailOptions);
  } catch (err) {
    Logger.log("Avertissement expéditeur personnalisé : " + err + ". Tentative d'envoi avec l'expéditeur par défaut...");
    delete mailOptions.from;
    MailApp.sendEmail(mailOptions);
  }
  Logger.log("Mail de convocation envoyé à " + email + " pour la session " + sessionId);
}

/**
 * Helper pour générer l'URL d'un formulaire pré-rempli avec l'entrée réelle
 */
function getPreFilledFormUrl(formId: string, entryId: string, selectedValue: string): string {
  return "https://docs.google.com/forms/d/e/" + formId + "/viewform?usp=pp_url&" + entryId + "=" + encodeURIComponent(selectedValue);
}
