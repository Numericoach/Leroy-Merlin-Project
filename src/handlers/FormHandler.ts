/**
 * Déclencheur sur soumission du formulaire d'inscription
 * Protégé par LockService contre les soumissions simultanées
 */
function onSubmit(e?: any): void {
  let lock: GoogleAppsScript.Lock.Lock | null = null;
  let hasLock = false;

  try {
    try {
      lock = LockService.getScriptLock();
      hasLock = lock.tryLock(10000);
    } catch (lockErr) {
      Logger.log("Avertissement LockService : " + lockErr);
    }

    // 1. DÉTECTER SI C'EST LA LISTE D'ATTENTE OU LA DÉSINSCRIPTION
    const sheetName = e && e.range ? e.range.getSheet().getName().toUpperCase() : "";

    if (sheetName.indexOf("ATTENTE") > -1) {
      Logger.log("Soumission enregistrée dans la liste d'attente : " + sheetName);
      if (ss) ss.toast("📥 Inscription sur la liste d'attente enregistrée !", "LISTE D'ATTENTE", 5);
      return;
    }

    if (sheetName.indexOf("DESINSCRIPTION") > -1) {
      if (ss) ss.toast("Désinscription détectée. Extraction de la session...", "DÉBOGAGE", 5);
      const data = extractSubmissionData(e);
      if (data.thisSessionid) {
        if (ss) ss.toast("Place libérée pour la session : " + data.thisSessionid + ". Recherche dans la file d'attente...", "DÉBOGAGE", 5);
        Logger.log("Désinscription détectée pour la session " + data.thisSessionid);
        processWaitingList(data.thisSessionid);
      } else {
        if (ss) ss.toast("❌ Impossible de trouver l'ID de session dans la désinscription.", "DÉBOGAGE", 8);
      }
      return;
    }

    const data = extractSubmissionData(e);
    
    if (!data.thisSessionid || !data.thisEmail) {
      Logger.log("Données manquantes (Email: " + data.thisEmail + ", Session: " + data.thisSessionid + ")");
      return;
    }

    // 2. VÉRIFIER LES PLACES RESTANTES DANS L'ONGLET SESSIONS
    const sheetSessions = ss ? ss.getSheetByName("SESSIONS") : null;
    let remainingSeatsAfterForm = 0;
    let sessionFound = false;

    if (sheetSessions) {
      const lastRowSessions = sheetSessions.getLastRow();
      if (lastRowSessions >= 2) {
        const sessionsData = sheetSessions.getRange(2, 2, lastRowSessions - 1, 13).getValues();
        for (let i = 0; i < sessionsData.length; i++) {
          const idInSheet = (sessionsData[i][0] || "").toString().trim();
          if (idInSheet === data.thisSessionid) {
            remainingSeatsAfterForm = Number(sessionsData[i][11]); // Colonne M (Places Restantes)
            sessionFound = true;
            break;
          }
        }

        // Recherche par comparaison si l'ID exact n'est pas dans des crochets
        if (!sessionFound) {
          const searchLower = data.thisSession.toLowerCase();
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
              data.thisSessionid = idInSheet;
              remainingSeatsAfterForm = Number(sessionsData[i][11]);
              sessionFound = true;
              break;
            }
          }
        }
      }
    }

    if (!sessionFound) {
      Logger.log("Avertissement : ID Session " + data.thisSessionid + " non trouvé dans SESSIONS. Passage par défaut.");
      remainingSeatsAfterForm = 999; // Défaut permissif pour ne pas bloquer l'inscription
    }

    // Si remainingSeatsAfterForm <= 0, cela signifie que la session est complète !
    if (remainingSeatsAfterForm <= 0) {
      Logger.log("Inscription refusée : session complète pour " + data.thisSessionid);
      if (ss) ss.toast("⚠️ Session " + data.thisSessionid + " complète ! Redirection vers la liste d'attente...", "INFO", 6);
      
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
        sendWaitingListMail(data.thisSessionid, data.thisEmail, data.thisPrenom, data.thisNom);
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
      verif = sessionsEmail.filter(row => (row[0] === data.thisSessionid && row[1] === data.thisEmail));
    }
    
    if (verif.length > 0) {
      Logger.log("Déjà inscrit : " + data.thisEmail + " à " + data.thisSessionid);
      if (ss) ss.toast("ℹ️ Vous êtes déjà inscrit à la session " + data.thisSessionid + ". Pas de nouveau mail envoyé.", "INFO", 7);
      return;
    }

    // 4. INSCRIPTION DANS LE SHEETS AVEC L'HORODATEUR EXACT DU FORMULAIRE
    inscription(data.thisTime, data.thisSessionid, data.thisEmail, data.thisNbParticipants);

    // 5. FONCTION PRINCIPALE : AJOUT DANS GOOGLE AGENDA
    try {
      addParticipantToCalendar(data.thisSessionid, data.thisEmail);
    } catch (agendaErr) {
      Logger.log("Avertissement : échec de l'ajout à l'agenda : " + agendaErr);
    }

    // 6. ENVOI DE LA CONVOCATION / CONFIRMATION
    try {
      sendConfirmationMail(data.thisSessionid, data.thisEmail, data.thisPrenom, data.thisNom, data.thisCivilite, data.thisNbParticipants);
      if (ss) ss.toast("✅ Inscription validée ! Convocation envoyée à " + data.thisEmail, "SUCCÈS", 7);
    } catch (mailErr) {
      Logger.log("Avertissement : échec de l'envoi d'e-mail : " + mailErr);
      if (ss) ss.toast("⚠️ Inscription enregistrée mais échec d'envoi du mail : " + mailErr, "AVERTISSEMENT", 7);
    }
    
  } catch (err) {
    Logger.log("Erreur critique dans onSubmit : " + err);
  } finally {
    if (hasLock && lock) {
      try {
        lock.releaseLock();
      } catch (relErr) {}
    }
  }
}

/**
 * Enregistrer l'inscription dans la feuille de calcul avec le timestamp exact du formulaire
 */
function inscription(time: any, sessionId: string, email: string, nbParticipants: number = 1): void {
  if (sheetInscriptions) {
    // N'ajouter que 3 colonnes pour ne pas écraser les formules de la colonne D
    sheetInscriptions.appendRow([time, sessionId, email]);
  }
}

/**
 * Extrait intelligemment toutes les données de soumission du formulaire
 */
function extractSubmissionData(e: any): any {
  let thisTime: any = null;
  let thisEmail = "";
  let thisSession = "";
  let thisCivilite = "";
  let thisPrenom = "";
  let thisNom = "";
  let thisNbParticipants = 1;

  if (e && e.range) {
    try {
      thisTime = e.range.getValues()[0][0];
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
      if ((keyLower.indexOf("nombre") > -1 && keyLower.indexOf("participant") > -1) || keyLower.indexOf("combien") > -1) {
        const parsedNb = parseInt(val, 10);
        if (!isNaN(parsedNb) && parsedNb > 0) {
          thisNbParticipants = parsedNb;
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

  let thisSessionid = "";
  const matchSes = thisSession.match(/(SES-[\w-]+)/i);
  if (matchSes && matchSes[1]) {
    thisSessionid = matchSes[1].trim().toUpperCase();
  } else {
    const matchBracket = thisSession.match(/\[(.*?)\]/);
    if (matchBracket && matchBracket[1]) {
      thisSessionid = matchBracket[1].trim();
    } else {
      thisSessionid = thisSession.trim();
    }
  }

  return {
    thisTime,
    thisEmail,
    thisSession,
    thisSessionid,
    thisCivilite,
    thisPrenom,
    thisNom,
    thisNbParticipants
  };
}

/**
 * Vérifie la file d'attente pour une session donnée et notifie la première personne.
 */
function processWaitingList(sessionId: string): void {
  let sheetAttente = ss ? (ss.getSheetByName("INSCRIPTIONS FILE ATTENTE") || ss.getSheetByName("FILE ATTENTE")) : null;
  
  if (!sheetAttente && ss) {
    const sheets = ss.getSheets();
    for (let i = 0; i < sheets.length; i++) {
      const name = sheets[i].getName().toUpperCase();
      if (name.indexOf("ATTENTE") > -1) {
        sheetAttente = sheets[i];
        break;
      }
    }
  }

  if (!sheetAttente) {
    if (ss) ss.toast("❌ Onglet Liste d'attente introuvable !", "DÉBOGAGE", 8);
    return;
  }

  const lastRow = sheetAttente.getLastRow();
  if (lastRow < 2) {
    if (ss) ss.toast("File d'attente vide pour l'instant.", "DÉBOGAGE", 5);
    return;
  }

  const data = sheetAttente.getRange(2, 1, lastRow - 1, 15).getValues(); // Lire 15 colonnes max
  let personNotified = false;

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    
    let isSessionMatch = false;
    let isNotified = false;
    let email = "";
    
    for (let c = 0; c < row.length; c++) {
      const val = (row[c] || "").toString().trim();
      if (val.indexOf(sessionId) > -1) isSessionMatch = true;
      if (val.indexOf("@") > -1 && val.indexOf(".") > -1 && !email) email = val;
      if (val.indexOf("Notifié") > -1 || val.indexOf("Inscrit") > -1) isNotified = true;
    }

    if (isSessionMatch && !isNotified && email) {
      const prenom = (row[2] || "").toString();
      const nom = (row[3] || "").toString();

      Logger.log("Place libérée ! Inscription automatique de " + email + " pour la session " + sessionId);
      if (ss) ss.toast("🚀 Promotion automatique de " + email + " pour la session " + sessionId + "...", "TRAITEMENT", 5);

      // 1. INSCRIPTION AUTOMATIQUE DANS L'ONGLET INSCRIPTIONS
      try {
        inscription(new Date(), sessionId, email, 1);
      } catch (inscErr) {
        Logger.log("Erreur lors de l'inscription automatique : " + inscErr);
      }

      // 2. AJOUT À L'AGENDA GOOGLE
      try {
        addParticipantToCalendar(sessionId, email);
      } catch (agendaErr) {
        Logger.log("Erreur ajout agenda automatique : " + agendaErr);
      }

      // 3. ENVOI DE L'E-MAIL DE CONFIRMATION ET CONVOCATION PDF
      try {
        sendConfirmationMail(sessionId, email, prenom, nom);
        if (ss) ss.toast("✅ " + email + " inscrit automatiquement & convocation envoyée !", "SUCCÈS", 8);
      } catch (err) {
        Logger.log("Erreur envoi mail confirmation : " + err);
        // Fallback si la convocation échoue : envoi du mail simple de place disponible
        try {
          sendSpotAvailableMail(sessionId, email, prenom, nom);
          if (ss) ss.toast("✅ Mail d'information envoyé à " + email, "SUCCÈS", 8);
        } catch (subErr) {
          if (ss) ss.toast("❌ Erreur d'envoi de mail à " + email + " : " + subErr, "ERREUR", 8);
        }
      }
      
      // Marquer comme inscrit automatiquement dans la première colonne vide
      let writeCol = 6;
      for (let c = 0; c < row.length; c++) {
         if ((row[c] || "").toString().trim() === "") {
            writeCol = c + 1;
            break;
         }
      }
      
      const timeStr = Utilities.formatDate(new Date(), "Europe/Paris", "dd/MM/yyyy HH:mm:ss");
      sheetAttente.getRange(i + 2, writeCol).setValue("Inscrit automatiquement le " + timeStr);
      personNotified = true;
      break; // Une seule personne promue par désinscription
    }
  }

  if (!personNotified && ss) {
    ss.toast("Aucun candidat en attente non-notifié trouvé pour " + sessionId, "DÉBOGAGE", 5);
  }
}

/**
 * Rattrapage manuel : scanne l'onglet INSCRIPTIONSS et traite les inscriptions qui ne sont pas encore dans INSCRIPTIONS.
 */
function processUnprocessedInscriptions(): void {
  const sheetFormResps = ss ? (ss.getSheetByName("INSCRIPTIONSS") || ss.getSheetByName("INSCRIPTIONS FORM")) : null;
  if (!sheetFormResps || !sheetInscriptions) {
    if (ss) ss.toast("❌ Onglets d'inscriptions introuvables.", "NUMERICOACH", 6);
    return;
  }

  const lastRow = sheetFormResps.getLastRow();
  if (lastRow < 2) {
    if (ss) ss.toast("Aucune réponse dans l'onglet des formulaires.", "NUMERICOACH", 5);
    return;
  }

  const data = sheetFormResps.getRange(2, 1, lastRow - 1, 10).getValues();
  let countProcessed = 0;

  data.forEach(row => {
    const time = row[0] || new Date();
    const email = (row[1] || "").toString().trim();
    const prenom = (row[2] || "").toString().trim();
    const nom = (row[3] || "").toString().trim();
    const rawSession = (row[5] || "").toString().trim();

    if (!email || !rawSession) return;

    let sessionId = "";
    const matchSes = rawSession.match(/(SES-[\w-]+)/i);
    if (matchSes && matchSes[1]) {
      sessionId = matchSes[1].trim().toUpperCase();
    } else {
      sessionId = rawSession;
    }

    // Vérifier si déjà dans INSCRIPTIONS
    const maxRows = sheetInscriptions.getMaxRows();
    let isAlreadyIn = false;
    if (maxRows > 1) {
      const existing = sheetInscriptions.getRange(2, 2, maxRows - 1, 2).getValues();
      isAlreadyIn = existing.some(r => r[0] === sessionId && r[1] === email);
    }

    if (!isAlreadyIn) {
      inscription(time, sessionId, email, 1);
      try {
        addParticipantToCalendar(sessionId, email);
      } catch (e) {}
      try {
        sendConfirmationMail(sessionId, email, prenom, nom);
      } catch (e) {}
      countProcessed++;
    }
  });

  if (ss) ss.toast("✅ " + countProcessed + " inscription(s) rattrapée(s) et traitée(s) !", "NUMERICOACH", 6);
}
