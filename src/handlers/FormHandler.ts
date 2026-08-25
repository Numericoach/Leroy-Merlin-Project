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
    let isDesinscription = sheetName.indexOf("DESINSCRIPTION") > -1 || sheetName.indexOf("DÉSINSCRIPTION") > -1;

    if (!isDesinscription && e && e.range) {
      try {
        const sh = e.range.getSheet();
        const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map((h: any) => (h || "").toString().toLowerCase());
        isDesinscription = headers.some((h: string) => h.indexOf("désinscription") > -1 || h.indexOf("desinscription") > -1 || h.indexOf("désinscrire") > -1);
      } catch (err) {}
    }

    if (sheetName.indexOf("ATTENTE") > -1) {
      Logger.log("Soumission enregistrée dans la liste d'attente : " + sheetName);
      const data = extractSubmissionData(e);
      if (data.thisEmail && data.thisSessionid) {
        try {
          sendWaitingListMail(data.thisSessionid, data.thisEmail, data.thisPrenom, data.thisNom);
        } catch (waitMailErr) {
          Logger.log("Erreur envoi mail confirmation liste d'attente : " + waitMailErr);
        }
      }
      if (ss) ss.toast("📥 Inscription sur la liste d'attente enregistrée & mail d'attente envoyé !", "LISTE D'ATTENTE", 5);
      return;
    }

    if (isDesinscription) {
      if (ss) ss.toast("🚪 Désinscription détectée. Retrait du participant et mise à jour de l'agenda...", "DÉSINSCRIPTION", 5);
      const data = extractSubmissionData(e);
      if (data.thisSessionid && data.thisEmail) {
        // 1. Supprimer de la feuille INSCRIPTIONS
        if (sheetInscriptions) {
          const lastRowInsc = sheetInscriptions.getLastRow();
          if (lastRowInsc >= 2) {
            const rows = sheetInscriptions.getRange(2, 1, lastRowInsc - 1, 3).getValues();
            for (let i = rows.length - 1; i >= 0; i--) {
              const rowSes = (rows[i][1] || "").toString().trim();
              const rowEmail = (rows[i][2] || "").toString().trim().toLowerCase();
              if (rowSes === data.thisSessionid && rowEmail === data.thisEmail.toLowerCase()) {
                sheetInscriptions.deleteRow(i + 2);
                Logger.log("Ligne d'inscription supprimée pour " + data.thisEmail + " - session " + data.thisSessionid);
              }
            }
          }
        }

        // 2. Retirer de Google Agenda (removeGuest) et rafraîchir la description et le titre (nb d'inscrits)
        try {
          removeParticipantFromCalendar(data.thisSessionid, data.thisEmail);
        } catch (agendaErr) {
          Logger.log("Avertissement retrait agenda : " + agendaErr);
        }

        if (ss) ss.toast("✅ Participant " + data.thisEmail + " retiré de l'agenda pour la session " + data.thisSessionid, "SUCCÈS", 6);
        Logger.log("Désinscription traitée pour la session " + data.thisSessionid);
        processWaitingList(data.thisSessionid);

        try {
          updateFormChoices();
        } catch (formErr) {
          Logger.log("Avertissement mise à jour formulaires : " + formErr);
        }
      } else {
        if (ss) ss.toast("❌ Impossible de trouver l'ID de session ou l'email dans la désinscription.", "AVERTISSEMENT", 8);
      }
      return;
    }

    const data = extractSubmissionData(e);
    
    if (!data.thisSessionid || !data.thisEmail) {
      Logger.log("Données manquantes (Email: " + data.thisEmail + ", Session: " + data.thisSessionid + ")");
      return;
    }

    // 2. VÉRIFIER LES PLACES RESTANTES ET CHARGER LES DÉTAILS DE LA SESSION
    const sheetSessions = ss ? ss.getSheetByName("SESSIONS") : null;
    let remainingSeatsAfterForm = 0;
    let sessionFound = false;
    let sessionDetails: any = null;

    if (sheetSessions) {
      const lastRowSessions = sheetSessions.getLastRow();
      if (lastRowSessions >= 2) {
        const sessionsData = sheetSessions.getRange(2, 2, lastRowSessions - 1, 14).getValues();
        for (let i = 0; i < sessionsData.length; i++) {
          const idInSheet = (sessionsData[i][0] || "").toString().trim();
          if (idInSheet === data.thisSessionid) {
            remainingSeatsAfterForm = Number(sessionsData[i][11]); // Colonne M (Places Restantes)
            sessionFound = true;
            
            const dateVal = sessionsData[i][2];
            const hdVal = sessionsData[i][3];
            const hfVal = sessionsData[i][4];
            let dateStr = "";
            let heureDebutStr = "";
            let heureFinStr = "";
            if (dateVal) {
              const d = new Date(dateVal);
              if (!isNaN(d.getTime())) {
                dateStr = d.getDate() + "/" + (d.getMonth() + 1) + "/" + d.getFullYear();
              }
            }
            if (hdVal) {
              const hd = new Date(hdVal);
              if (!isNaN(hd.getTime())) {
                heureDebutStr = hd.getHours() + "h" + (hd.getMinutes() < 10 ? "0" : "") + hd.getMinutes();
              }
            }
            if (hfVal) {
              const hf = new Date(hfVal);
              if (!isNaN(hf.getTime())) {
                heureFinStr = hf.getHours() + "h" + (hf.getMinutes() < 10 ? "0" : "") + hf.getMinutes();
              }
            }

            sessionDetails = {
              formationTitle: sessionsData[i][13] || "Formation Leroy Merlin",
              dateStr: dateStr,
              heureDebutStr: heureDebutStr,
              heureFinStr: heureFinStr,
              lieuStr: sessionsData[i][7] || "",
              infoCompStr: sessionsData[i][8] || ""
            };
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

              const dateVal = sessionsData[i][2];
              const hdVal = sessionsData[i][3];
              const hfVal = sessionsData[i][4];
              let dateStr = "";
              let heureDebutStr = "";
              let heureFinStr = "";
              if (dateVal) {
                const d = new Date(dateVal);
                if (!isNaN(d.getTime())) dateStr = d.getDate() + "/" + (d.getMonth() + 1) + "/" + d.getFullYear();
              }
              if (hdVal) {
                const hd = new Date(hdVal);
                if (!isNaN(hd.getTime())) heureDebutStr = hd.getHours() + "h" + (hd.getMinutes() < 10 ? "0" : "") + hd.getMinutes();
              }
              if (hfVal) {
                const hf = new Date(hfVal);
                if (!isNaN(hf.getTime())) heureFinStr = hf.getHours() + "h" + (hf.getMinutes() < 10 ? "0" : "") + hf.getMinutes();
              }

              sessionDetails = {
                formationTitle: sessionsData[i][13] || "Formation Leroy Merlin",
                dateStr: dateStr,
                heureDebutStr: heureDebutStr,
                heureFinStr: heureFinStr,
                lieuStr: sessionsData[i][7] || "",
                infoCompStr: sessionsData[i][8] || ""
              };
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
      const targetSes = (data.thisSessionid || "").trim().toUpperCase();
      const targetEmail = (data.thisEmail || "").toString().trim().toLowerCase();
      verif = sessionsEmail.filter(row => {
        const rowSes = (row[0] || "").toString().trim().toUpperCase();
        const rowEmail = (row[1] || "").toString().trim().toLowerCase();
        return rowSes === targetSes && rowEmail === targetEmail;
      });
    }
    
    if (verif.length > 0) {
      Logger.log("Déjà inscrit : " + data.thisEmail + " à " + data.thisSessionid);
      if (ss) ss.toast("ℹ️ " + data.thisEmail + " est déjà inscrit(e) à la session " + data.thisSessionid + ". Ligne en doublon supprimée.", "INFO", 7);
      
      // SUPPRIMER LA LIGNE EN DOUBLON DE LA FEUILLE DE RÉPONSES
      if (e && e.range) {
        try {
          e.range.getSheet().deleteRow(e.range.getRow());
          Logger.log("Ligne d'inscription en doublon supprimée de " + e.range.getSheet().getName() + " à la ligne " + e.range.getRow());
        } catch (delErr) {
          Logger.log("Erreur lors de la suppression de la ligne en doublon : " + delErr);
        }
      }
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
      sendConfirmationMail(data.thisSessionid, data.thisEmail, data.thisPrenom, data.thisNom, data.thisCivilite, data.thisNbParticipants, sessionDetails);
      if (ss) ss.toast("✅ Inscription validée ! Convocation envoyée à " + data.thisEmail, "SUCCÈS", 7);
    } catch (mailErr) {
      Logger.log("Avertissement : échec de l'envoi d'e-mail : " + mailErr);
      if (ss) ss.toast("⚠️ Inscription enregistrée mais échec d'envoi du mail : " + mailErr, "AVERTISSEMENT", 7);
    }

    // 7. RAFRAÎCHIR AUTOMATIQUEMENT LES CHOIX DE SESSIONS DANS LES FORMULAIRES
    try {
      updateFormChoices();
    } catch (syncErr) {
      Logger.log("Avertissement rafraîchissement des formulaires : " + syncErr);
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

function inscription(time: any, sessionId: string, email: string, nbParticipants: number = 1): void {
  if (sheetInscriptions) {
    const cleanSessionId = (sessionId || "").toString().trim().toUpperCase();
    const cleanEmail = (email || "").toString().trim().toLowerCase();
    // N'ajouter que 3 colonnes propres pour ne pas écraser les formules de la colonne D
    sheetInscriptions.appendRow([time, cleanSessionId, cleanEmail]);
  }
}

/**
 * Helper pour normaliser les en-têtes (minuscules, sans accents, sans espaces superflus)
 */
function normalizeHeaderString(str: string): string {
  return (str || "")
    .toString()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim();
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

  // 1. Extraction basée sur namedValues (Google Forms)
  if (e && e.namedValues) {
    for (const key in e.namedValues) {
      const keyNorm = normalizeHeaderString(key);
      const val = (e.namedValues[key][0] || "").toString().trim();
      if (!val) continue;

      if (!thisEmail && (keyNorm.indexOf("mail") > -1 || keyNorm.indexOf("courriel") > -1)) {
        thisEmail = val;
      }
      if (val.indexOf("[") > -1 || val.indexOf("SES-") > -1 || keyNorm.indexOf("inscription") > -1 || keyNorm.indexOf("session") > -1) {
        if (val.indexOf("[") > -1 || val.indexOf("SES-") > -1 || !thisSession) {
          thisSession = val;
        }
      }
      if (!thisCivilite && (keyNorm.indexOf("civilite") > -1 || keyNorm.indexOf("titre") > -1)) {
        thisCivilite = val;
      }
      if (!thisPrenom && (keyNorm.indexOf("prenom") > -1 || keyNorm.indexOf("first") > -1)) {
        thisPrenom = val;
      }
      if (!thisNom && keyNorm.indexOf("prenom") === -1 && keyNorm.indexOf("nombre") === -1 && (keyNorm.indexOf("nom") > -1 || keyNorm.indexOf("last") > -1)) {
        thisNom = val;
      }
      if ((keyNorm.indexOf("nombre") > -1 && keyNorm.indexOf("participant") > -1) || keyNorm.indexOf("combien") > -1) {
        const parsedNb = parseInt(val, 10);
        if (!isNaN(parsedNb) && parsedNb > 0) {
          thisNbParticipants = parsedNb;
        }
      }
    }
  }

  // 2. Extraction basée sur les en-têtes réels de la feuille de réponse (e.range.getSheet())
  if (e && e.range) {
    try {
      const sheet = e.range.getSheet();
      const lastCol = sheet.getLastColumn();
      if (lastCol > 0) {
        const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => normalizeHeaderString(h ? h.toString() : ""));
        const rowVals = e.range.getValues()[0];

        for (let col = 0; col < headers.length; col++) {
          const h = headers[col];
          const val = (rowVals[col] || "").toString().trim();
          if (!val) continue;

          if (!thisEmail && (h.indexOf("mail") > -1 || h.indexOf("courriel") > -1)) {
            thisEmail = val;
          }
          if (!thisPrenom && (h.indexOf("prenom") > -1 || h.indexOf("first") > -1)) {
            thisPrenom = val;
          }
          if (!thisNom && h.indexOf("prenom") === -1 && h.indexOf("nombre") === -1 && (h.indexOf("nom") > -1 || h.indexOf("last") > -1)) {
            thisNom = val;
          }
          if (!thisCivilite && (h.indexOf("civilite") > -1 || h.indexOf("titre") > -1)) {
            thisCivilite = val;
          }
          if (!thisSession && (val.indexOf("[") > -1 || val.indexOf("SES-") > -1 || h.indexOf("session") > -1 || h.indexOf("inscription") > -1)) {
            thisSession = val;
          }
        }
      }
    } catch (err) {
      Logger.log("Erreur lors de l'extraction par la feuille de réponse : " + err);
    }
  }

  // 3. Extraction de secours sur e.values pour e-mail et session uniquement
  if (e && e.values && Array.isArray(e.values)) {
    e.values.forEach((val: any) => {
      const valStr = (val || "").toString().trim();
      if (!thisEmail && valStr.indexOf("@") > -1) {
        thisEmail = valStr;
      }
      if (!thisSession && (valStr.indexOf("[") > -1 || valStr.indexOf("SES-") > -1)) {
        thisSession = valStr;
      }
    });
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
function processWaitingList(sessionId: string): boolean {
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
    return false;
  }

  const lastRow = sheetAttente.getLastRow();
  if (lastRow < 2) {
    return false;
  }

  const lastColAttente = sheetAttente.getLastColumn();
  const headersAttente = sheetAttente.getRange(1, 1, 1, lastColAttente).getValues()[0].map(h => normalizeHeaderString(h ? h.toString() : ""));
  const prenomColIdxAtt = headersAttente.findIndex(h => h.indexOf("prenom") > -1 || h.indexOf("first") > -1);
  const nomColIdxAtt = headersAttente.findIndex(h => h.indexOf("prenom") === -1 && h.indexOf("nombre") === -1 && (h.indexOf("nom") > -1 || h.indexOf("last") > -1));

  const data = sheetAttente.getRange(2, 1, lastRow - 1, lastColAttente).getValues();
  let personNotified = false;
  const cleanSessionId = (sessionId || "").toString().trim().toUpperCase();

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    
    let isSessionMatch = false;
    let isNotified = false;
    let email = "";
    
    for (let c = 0; c < row.length; c++) {
      const val = (row[c] || "").toString().trim().toUpperCase();
      if (val.indexOf(cleanSessionId) > -1) isSessionMatch = true;
      if (val.indexOf("@") > -1 && val.indexOf(".") > -1 && !email) email = row[c].toString().trim();
      if (val.indexOf("NOTIFIÉ") > -1 || val.indexOf("INSCRIT") > -1) isNotified = true;
    }

    if (isSessionMatch && !isNotified && email) {
      const prenom = prenomColIdxAtt > -1 ? (row[prenomColIdxAtt] || "").toString().trim() : "";
      const nom = nomColIdxAtt > -1 ? (row[nomColIdxAtt] || "").toString().trim() : "";

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
        try {
          sendSpotAvailableMail(sessionId, email, prenom, nom);
          if (ss) ss.toast("✅ Mail d'information envoyé à " + email, "SUCCÈS", 8);
        } catch (subErr) {
          if (ss) ss.toast("❌ Erreur d'envoi de mail à " + email + " : " + subErr, "ERREUR", 8);
        }
      }
      
      let writeCol = Math.max(lastColAttente, 6);
      for (let c = 0; c < row.length; c++) {
         const strCell = (row[c] || "").toString().trim();
         if (strCell.indexOf("Inscrit") > -1 || strCell.indexOf("Notifié") > -1) {
            writeCol = c + 1;
            break;
         }
         if (strCell === "" && c >= 5) {
            writeCol = c + 1;
            break;
         }
      }
      
      const timeStr = Utilities.formatDate(new Date(), "Europe/Paris", "dd/MM/yyyy HH:mm:ss");
      sheetAttente.getRange(i + 2, writeCol).setValue("Inscrit automatiquement le " + timeStr);
      personNotified = true;
      break;
    }
  }

  return personNotified;
}

/**
 * Parcourt toutes les sessions publiées ayant des places disponibles
 * et promeut automatiquement les candidats inscrits en liste d'attente.
 */
function processAllWaitingLists(): void {
  const sheetSessions = ss ? ss.getSheetByName("SESSIONS") : null;
  if (!sheetSessions) return;

  const lastRow = sheetSessions.getLastRow();
  if (lastRow < 2) return;

  const sessionsData = sheetSessions.getRange(2, 1, lastRow - 1, 14).getValues();
  let totalPromoted = 0;

  sessionsData.forEach(row => {
    const sessionId = (row[1] || "").toString().trim(); // Col B (ID SESSION)
    const publish = row[10]; // Col K (Publier)
    const rawRemaining = parseFloat(String(row[12] || "").replace(",", ".")); // Col M (Places Restantes)
    const nbPlaces = parseFloat(String(row[7] || "").replace(",", ".")); // Col H (NB PLACES)
    const nbInscrits = parseFloat(String(row[11] || "").replace(",", ".")); // Col L (Nb Inscrits)

    const isPublished = Boolean(publish) && 
                      String(publish).toUpperCase() !== "FALSE" && 
                      String(publish).toUpperCase() !== "FAUX" && 
                      String(publish) !== "0" && 
                      String(publish).trim() !== "";

    let remainingSeats = 0;
    if (!isNaN(rawRemaining)) {
      remainingSeats = rawRemaining;
    } else if (!isNaN(nbPlaces)) {
      remainingSeats = nbPlaces - (!isNaN(nbInscrits) ? nbInscrits : 0);
    }

    if (sessionId && isPublished && remainingSeats > 0) {
      for (let p = 0; p < remainingSeats; p++) {
        const promoted = processWaitingList(sessionId);
        if (promoted) {
          totalPromoted++;
        } else {
          break;
        }
      }
    }
  });

  if (totalPromoted > 0 && ss) {
    ss.toast("✅ " + totalPromoted + " candidat(s) promu(s) depuis la liste d'attente !", "LISTE D'ATTENTE", 6);
  }
}

/**
 * Rattrapage manuel : scanne l'onglet INSCRIPTIONSS et traite les inscriptions qui ne sont pas encore dans INSCRIPTIONS.
 */
function processUnprocessedInscriptions(): void {
  const sheetFormResps = ss ? (ss.getSheetByName("INSCRIPTIONSS") || ss.getSheetByName("INSCRIPTIONS FORM")) : null;
  if (!sheetFormResps || !sheetInscriptions) {
    if (ss) ss.toast("❌ Onglets d'inscriptions introuvables.", "OUTILS", 6);
    return;
  }

  const lastRow = sheetFormResps.getLastRow();
  if (lastRow < 2) {
    if (ss) ss.toast("Aucune réponse dans l'onglet des formulaires.", "OUTILS", 5);
    return;
  }

  const lastCol = sheetFormResps.getLastColumn();
  const rawHeaders = sheetFormResps.getRange(1, 1, 1, lastCol).getValues()[0].map(h => normalizeHeaderString(h ? h.toString() : ""));
  
  const emailColIdx = rawHeaders.findIndex(h => h.indexOf("mail") > -1 || h.indexOf("courriel") > -1);
  const prenomColIdx = rawHeaders.findIndex(h => h.indexOf("prenom") > -1 || h.indexOf("first") > -1);
  const nomColIdx = rawHeaders.findIndex(h => h.indexOf("prenom") === -1 && h.indexOf("nombre") === -1 && (h.indexOf("nom") > -1 || h.indexOf("last") > -1));
  const civiliteColIdx = rawHeaders.findIndex(h => h.indexOf("civilite") > -1 || h.indexOf("titre") > -1);
  const sessionColIdx = rawHeaders.findIndex(h => h.indexOf("session") > -1 || h.indexOf("inscription") > -1 || h.indexOf("formation") > -1);

  const data = sheetFormResps.getRange(2, 1, lastRow - 1, lastCol).getValues();
  let countProcessed = 0;

  // Charger toutes les inscriptions existantes une seule fois avant la boucle pour un contrôle en O(1)
  const existingSet = new Set<string>();
  const maxRows = sheetInscriptions.getMaxRows();
  if (maxRows > 1) {
    const existing = sheetInscriptions.getRange(2, 2, maxRows - 1, 2).getValues();
    existing.forEach(r => {
      const ses = (r[0] || "").toString().trim();
      const em = (r[1] || "").toString().trim().toLowerCase();
      if (ses && em) {
        existingSet.add(ses + "|" + em);
      }
    });
  }

  data.forEach(row => {
    const time = row[0] || new Date();
    const email = emailColIdx > -1 ? (row[emailColIdx] || "").toString().trim() : "";
    const prenom = prenomColIdx > -1 ? (row[prenomColIdx] || "").toString().trim() : "";
    const nom = nomColIdx > -1 ? (row[nomColIdx] || "").toString().trim() : "";
    const civilite = civiliteColIdx > -1 ? (row[civiliteColIdx] || "").toString().trim() : "";
    
    let rawSession = "";
    if (sessionColIdx > -1) {
      rawSession = (row[sessionColIdx] || "").toString().trim();
    } else {
      row.forEach((cell: any) => {
        const str = (cell || "").toString().trim();
        if (!rawSession && (str.indexOf("SES-") > -1 || str.indexOf("[") > -1)) {
          rawSession = str;
        }
      });
    }

    if (!email || !rawSession) return;

    let sessionId = "";
    const matchSes = rawSession.match(/(SES-[\w-]+)/i);
    if (matchSes && matchSes[1]) {
      sessionId = matchSes[1].trim().toUpperCase();
    } else {
      sessionId = rawSession;
    }

    const key = sessionId + "|" + email.toLowerCase();
    if (!existingSet.has(key)) {
      existingSet.add(key);
      inscription(time, sessionId, email, 1);
      try {
        addParticipantToCalendar(sessionId, email);
      } catch (e) {}
      try {
        sendConfirmationMail(sessionId, email, prenom, nom, civilite);
      } catch (e) {}
      countProcessed++;
    }
  });

  if (ss) ss.toast("✅ " + countProcessed + " inscription(s) rattrapée(s) et traitée(s) !", "OUTILS", 6);
}
