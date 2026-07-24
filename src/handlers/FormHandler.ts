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
      remainingSeatsAfterForm = 0;
    }

    // Si remainingSeatsAfterForm < 0, cela signifie que la session est déjà complète !
    if (remainingSeatsAfterForm < 0) {
      Logger.log("Inscription refusée : session complète pour " + data.thisSessionid);
      
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
      return;
    }

    // 4. INSCRIPTION DANS LE SHEETS AVEC L'HORODATEUR EXACT DU FORMULAIRE
    inscription(data.thisTime, data.thisSessionid, data.thisEmail, data.thisNbParticipants);

    // 5. FONCTION PRINCIPALE : AJOUT DANS GOOGLE AGENDA (Priorité absolue)
    addParticipantToCalendar(data.thisSessionid, data.thisEmail);

    // 6. FONCTION SECONDAIRE : ENVOI DE LA CONVOCATION / CONFIRMATION (Découplé)
    try {
      sendConfirmationMail(data.thisSessionid, data.thisEmail, data.thisPrenom, data.thisNom, data.thisCivilite, data.thisNbParticipants);
    } catch (mailErr) {
      Logger.log("Avertissement : échec de l'envoi d'e-mail (n'impacte pas l'inscription ni l'agenda) : " + mailErr);
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
