/**
 * Helper pour obtenir l'agenda cible (ID renseigné dans PARAMETRES ou Agenda Principal par défaut)
 */
function getTargetCalendar(): GoogleAppsScript.Calendar.Calendar | null {
  const agendaId = getParamValue("PARAMETRE_ID_AGENDA");
  if (agendaId && agendaId.trim() !== "") {
    const cleanId = agendaId.trim();
    try {
      const agenda = CalendarApp.getCalendarById(cleanId);
      if (agenda) return agenda;
      Logger.log("ATTENTION : L'agenda Google ID '" + cleanId + "' est introuvable ou n'est pas partagé avec le compte exécutant le script (Droits de modification requis).");
    } catch (err) {
      Logger.log("Avertissement getCalendarById(" + cleanId + ") : " + err + ". Utilisation de l'agenda principal.");
    }
  }
  try {
    return CalendarApp.getDefaultCalendar();
  } catch (err) {
    Logger.log("Erreur lors de la récupération de l'agenda par défaut : " + err);
    return null;
  }
}

/**
 * Helper robuste pour récupérer un événement Google Agenda par son ID via CalendarApp
 */
function getEventByIdRobust(eventId: string): GoogleAppsScript.Calendar.CalendarEvent | null {
  if (!eventId) return null;
  let event: GoogleAppsScript.Calendar.CalendarEvent | null = null;
  const targetCal = getTargetCalendar();

  if (targetCal) {
    try {
      event = targetCal.getEventById(eventId);
    } catch (e) {}
    if (!event && eventId.indexOf("@") === -1) {
      try {
        event = targetCal.getEventById(eventId + "@google.com");
      } catch (e) {}
    }
  }

  if (!event) {
    try {
      event = CalendarApp.getEventById(eventId);
    } catch (e) {}
    if (!event && eventId.indexOf("@") === -1) {
      try {
        event = CalendarApp.getEventById(eventId + "@google.com");
      } catch (e) {}
    }
  }

  return event;
}

/**
 * Nettoie le titre du module pour éviter la répétition 'Accompagnement Leroy Merlin - Formation Leroy Merlin'
 */
function getCleanFormationTitle(rawTitle: string, fallbackModule: string = ""): string {
  let str = (rawTitle || "").toString().trim();
  if (!str || str.toLowerCase() === "formation leroy merlin" || str.toLowerCase() === "accompagnement leroy merlin") {
    str = (fallbackModule || "").toString().trim();
  }
  
  let cleaned = str
    .replace(/\[FOR-.*?\]/gi, "")
    .replace(/^Accompagnement Leroy Merlin\s*-\s*/i, "")
    .replace(/^Accompagnement Leroy Merlin/i, "")
    .replace(/^Formation Leroy Merlin\s*-\s*/i, "")
    .replace(/^Formation Leroy Merlin/i, "")
    .replace(/^Formation\s+/i, "")
    .trim();

  if (!cleaned || cleaned.toLowerCase() === "formation leroy merlin" || cleaned.toLowerCase() === "accompagnement leroy merlin") {
    if (fallbackModule) {
      cleaned = fallbackModule
        .replace(/\[FOR-.*?\]/gi, "")
        .replace(/^Accompagnement Leroy Merlin\s*-\s*/i, "")
        .replace(/^Accompagnement Leroy Merlin/i, "")
        .replace(/^Formation Leroy Merlin\s*-\s*/i, "")
        .replace(/^Formation Leroy Merlin/i, "")
        .replace(/^Formation\s+/i, "")
        .trim();
    }
  }

  if (cleaned.toLowerCase() === "formation leroy merlin" || cleaned.toLowerCase() === "accompagnement leroy merlin") {
    cleaned = "";
  }

  return cleaned;
}

/**
 * Extrait le nombre de participants spécifié dans la réponse du formulaire pour cet e-mail
 */
function getNbParticipantsForEmailAndSession(sessionId: string, email: string): number {
  if (!ss) return 1;
  const cleanEmail = (email || "").toLowerCase().trim();
  if (!cleanEmail) return 1;
  
  try {
    const sheets = ss.getSheets();
    for (let s = 0; s < sheets.length; s++) {
      const name = sheets[s].getName();
      const lowerName = name.toLowerCase();
      if (
        lowerName.indexOf("form") > -1 ||
        lowerName.indexOf("répons") > -1 ||
        lowerName.indexOf("repons") > -1 ||
        lowerName.indexOf("inscriptions") > -1
      ) {
        const lastRow = sheets[s].getLastRow();
        const lastCol = sheets[s].getLastColumn();
        if (lastRow >= 2 && lastCol >= 2) {
          const headers = sheets[s].getRange(1, 1, 1, lastCol).getValues()[0].map(h => (h || "").toString().toLowerCase());
          const emailColIdx = headers.findIndex(h => h.indexOf("mail") > -1 || h.indexOf("email") > -1 || h.indexOf("courriel") > -1);
          const nbColIdx = headers.findIndex(h => h.indexOf("nombre") > -1 || h.indexOf("combien") > -1 || h.indexOf("participant") > -1);
          
          if (emailColIdx > -1) {
            const values = sheets[s].getRange(2, 1, lastRow - 1, lastCol).getValues();
            for (let r = values.length - 1; r >= 0; r--) {
              const rowEmail = (values[r][emailColIdx] || "").toString().toLowerCase().trim();
              if (rowEmail === cleanEmail) {
                if (nbColIdx > -1 && values[r][nbColIdx] !== undefined && values[r][nbColIdx] !== null && values[r][nbColIdx] !== "") {
                  const rawVal = values[r][nbColIdx].toString().trim();
                  const parsed = parseInt(rawVal.replace(/[^0-9]/g, ""), 10);
                  if (!isNaN(parsed) && parsed > 0) {
                    return parsed;
                  }
                }
              }
            }
          }
        }
      }
    }
  } catch (e) {
    Logger.log("Erreur lors de la lecture du nombre de participants : " + e);
  }
  return 1;
}

/**
 * Extrait le lien Google Meet configuré ou généré automatiquement pour une session (priorité à l'événement Agenda réel)
 */
function getMeetUrlForSession(sessionId: string): string {
  if (!sessionId) return "https://meet.google.com/apv-qoem-zpc";
  
  // 1. Chercher le lien Meet réel rattaché à l'événement Google Agenda de cette session
  try {
    const sheetSessionEvenements = ss ? ss.getSheetByName("SESSION AGENDA") : null;
    if (sheetSessionEvenements) {
      const lastRowEvt = sheetSessionEvenements.getLastRow();
      if (lastRowEvt >= 2) {
        const sessionEvenementValues = sheetSessionEvenements.getRange(2, 1, lastRowEvt - 1, 3).getValues();
        for (let i = sessionEvenementValues.length - 1; i >= 0; i--) {
          if ((sessionEvenementValues[i][1] || "").toString().trim().toUpperCase() === sessionId.trim().toUpperCase()) {
            const storedId = sessionEvenementValues[i][2] as string;
            if (storedId) {
              const event = getEventByIdRobust(storedId);
              if (event) {
                // a) Via l'API Google Apps Script native
                try {
                  const hangout = (event as any).getHangoutLink ? (event as any).getHangoutLink() : null;
                  if (hangout && hangout.indexOf("meet.google.com") > -1) {
                    return hangout;
                  }
                } catch (e) {}

                // b) Via la description de l'événement Agenda
                const desc = event.getDescription() || "";
                const matchDesc = desc.match(/(https:\/\/meet\.google\.com\/[a-z0-9\-]+)/i);
                if (matchDesc && matchDesc[1]) {
                  return matchDesc[1];
                }

                // c) Via Calendar v3 Advanced Service
                try {
                  const agenda = getTargetCalendar();
                  const calendarId = agenda ? agenda.getId() : "primary";
                  const cleanEvtId = storedId.replace("@google.com", "");
                  if (typeof (globalThis as any).Calendar !== "undefined" && (globalThis as any).Calendar.Events) {
                    const fetchedEvt = (globalThis as any).Calendar.Events.get(calendarId, cleanEvtId);
                    if (fetchedEvt && fetchedEvt.hangoutLink) {
                      return fetchedEvt.hangoutLink;
                    }
                  }
                } catch (e) {}
              }
            }
          }
        }
      }
    }
  } catch (e) {
    Logger.log("Erreur lors de la récupération du lien Meet de l'événement : " + e);
  }

  // 2. Chercher dans l'onglet SESSIONS / LIEUX
  try {
    const sheetSessions = ss ? ss.getSheetByName("SESSIONS") : null;
    if (sheetSessions) {
      const lastRow = sheetSessions.getLastRow();
      if (lastRow >= 2) {
        const sessionsValues = sheetSessions.getRange(2, 2, lastRow - 1, 14).getValues();
        for (let i = 0; i < sessionsValues.length; i++) {
          if ((sessionsValues[i][0] || "").toString().trim().toUpperCase() === sessionId.trim().toUpperCase()) {
            const lieuStr = (sessionsValues[i][6] || sessionsValues[i][7] || "").toString();
            const match = lieuStr.match(/\[(.*?)\]/);
            if (match && match[1]) {
              const lieuId = match[1];
              const sheetLieux = ss ? ss.getSheetByName("LIEUX") : null;
              if (sheetLieux) {
                const maxLieux = sheetLieux.getLastRow();
                if (maxLieux >= 2) {
                  const lieuxData = sheetLieux.getRange(2, 2, maxLieux - 1, 7).getValues();
                  for (let j = 0; j < lieuxData.length; j++) {
                    if (lieuxData[j][0] === lieuId) {
                      const addr = (lieuxData[j][2] || "").toString();
                      const acces = (lieuxData[j][5] || "").toString();
                      if (addr.indexOf("meet.google.com") > -1) return addr;
                      const meetMatch = acces.match(/(https:\/\/meet\.google\.com\/[a-z0-9\-]+)/i);
                      if (meetMatch) return meetMatch[1];
                    }
                  }
                }
              }
            }
            break;
          }
        }
      }
    }
  } catch (e) {}

  // 3. Fallback sur PARAMETRES / Défaut
  const connInfo = getParamValue("PARAMETRE_CONNEXION_1");
  if (connInfo) {
    const meetMatch = connInfo.match(/(https:\/\/meet\.google\.com\/[a-z0-9\-]+)/i);
    if (meetMatch) return meetMatch[1];
  }

  return "https://meet.google.com/apv-qoem-zpc";
}

/**
 * Obtenir ou créer l'événement Google Agenda pour une session donnée (protégé par LockService)
 */
function getOrCreateSessionEventId(sessionId: string): string | null {
  if (!sessionId) return null;

  const lock = LockService.getScriptLock();
  // Lever une exception ou échouer proprement si le lock ne peut pas être obtenu
  const hasLock = lock.tryLock(Constants.LOCK.TIMEOUT_MS);
  if (!hasLock) {
    Logger.log("Échec getOrCreateSessionEventId : impossible d'acquérir le verrou pour la session " + sessionId);
    return null;
  }

  try {
    let sheetSessionEvenements = ss ? ss.getSheetByName(Constants.SHEETS.SESSION_AGENDA || "SESSION AGENDA") : null;
    if (!sheetSessionEvenements && ss) {
      try {
        sheetSessionEvenements = ss.insertSheet("SESSION AGENDA");
        sheetSessionEvenements.appendRow(["Horodateur", "ID SESSION", "ID EVENT AGENDA"]);
      } catch (e) {
        Logger.log("Erreur lors de la création de l'onglet SESSION AGENDA : " + e);
      }
    }

    const agenda = getTargetCalendar();
    if (!agenda) {
      Logger.log("Aucun agenda Google disponible.");
      return null;
    }

    const cleanSessionId = (sessionId || "").toString().trim().toUpperCase();

    // 1. Recherche prioritaire : vérifier si l'ID d'événement est déjà renseigné dans l'onglet SESSIONS
    const sheetSessions = ss ? ss.getSheetByName("SESSIONS") : null;
    if (!sheetSessions) return null;
    const lastRow = sheetSessions.getLastRow();
    if (lastRow < 2) return null;

    // Lire de la colonne B (ID SESSION) à la colonne Y (ID EVENEMENT CALENDAR, colonne 25)
    // 24 colonnes à partir de la colonne 2 (B) permet d'inclure Y (2 + 24 - 1 = 25)
    const sessionsRange = sheetSessions.getRange(2, 2, lastRow - 1, 24);
    const sessionsValues = sessionsRange.getValues();
    let targetRowIndex = -1;
    let existingEventId = "";
    let targetSession: any[] | null = null;

    for (let i = 0; i < sessionsValues.length; i++) {
      if ((sessionsValues[i][0] || "").toString().trim().toUpperCase() === cleanSessionId) {
        targetRowIndex = i + 2; // Index de ligne physique dans la feuille
        existingEventId = (sessionsValues[i][23] || "").toString().trim(); // Index 23 de la ligne lue (Col Y)
        targetSession = sessionsValues[i];
        break;
      }
    }

    if (!targetSession) return null;

    const dateDebut = parseDateTime(targetSession[2], targetSession[3]);
    const dateFin = parseDateTime(targetSession[2], targetSession[4]);

    // Si on a déjà un ID d'événement dans SESSIONS, on vérifie s'il existe toujours dans Google Agenda
    if (existingEventId) {
      let existingEvent: GoogleAppsScript.Calendar.CalendarEvent | null = null;
      try {
        existingEvent = agenda.getEventById(existingEventId);
      } catch (e) {}
      if (!existingEvent && existingEventId.indexOf("@") === -1) {
        try {
          existingEvent = agenda.getEventById(existingEventId + "@google.com");
        } catch (e) {}
      }
      if (existingEvent) {
        // Mettre à jour la date/heure si elle a changé dans le Sheets
        if (existingEvent.getStartTime().getTime() !== dateDebut.getTime() ||
            existingEvent.getEndTime().getTime() !== dateFin.getTime()) {
          existingEvent.setTime(dateDebut, dateFin);
          Logger.log("Mise à jour de la date/heure de l'événement existant " + existingEventId + " pour la session " + sessionId);
        }
        return existingEventId;
      }
    }

    // 2. Recherche secondaire dans la table de correspondance SESSION AGENDA
    if (sheetSessionEvenements) {
      const lastRowEvt = sheetSessionEvenements.getLastRow();
      if (lastRowEvt >= 2) {
        const sessionEvenementValues = sheetSessionEvenements.getRange(2, 1, lastRowEvt - 1, 3).getValues();
        for (let i = sessionEvenementValues.length - 1; i >= 0; i--) {
          const storedSessionId = (sessionEvenementValues[i][1] || "").toString().trim().toUpperCase();
          if (storedSessionId === cleanSessionId) {
            const storedId = (sessionEvenementValues[i][2] || "").toString().trim();
            if (storedId) {
              let existingEvent: GoogleAppsScript.Calendar.CalendarEvent | null = null;
              try {
                existingEvent = agenda.getEventById(storedId);
              } catch (e) {}
              if (!existingEvent && storedId.indexOf("@") === -1) {
                try {
                  existingEvent = agenda.getEventById(storedId + "@google.com");
                } catch (e) {}
              }
              if (existingEvent) {
                // Mettre à jour la date/heure si elle a changé dans le Sheets
                if (existingEvent.getStartTime().getTime() !== dateDebut.getTime() ||
                    existingEvent.getEndTime().getTime() !== dateFin.getTime()) {
                  existingEvent.setTime(dateDebut, dateFin);
                  Logger.log("Mise à jour de la date/heure de l'événement existant " + storedId + " pour la session " + sessionId);
                }
                // NE PAS écrire dans l'onglet SESSIONS car la colonne N est gérée par une formule ArrayFormula
                return storedId;
              }
            }
          }
        }
      }
    }

    // 3. Recherche de sécurité sur Google Agenda par tag [sessionId]
    try {
      const dateDebutSearch = parseDateTime(targetSession[2], targetSession[3]);
      const searchTag = "[" + cleanSessionId + "]";
      const timeMin = new Date(dateDebutSearch.getTime() - 24 * 3600 * 1000);
      const timeMax = new Date(dateDebutSearch.getTime() + 24 * 3600 * 1000);
      const foundEvents = agenda.getEvents(timeMin, timeMax, { search: searchTag });
      if (foundEvents && foundEvents.length > 0) {
        const foundEvt = foundEvents[0];
        const foundId = foundEvt.getId();
        if (sheetSessionEvenements) {
          sheetSessionEvenements.appendRow([new Date(), sessionId, foundId]);
        }
        // NE PAS écrire dans l'onglet SESSIONS car la colonne N est gérée par une formule ArrayFormula
        Logger.log("Événement Agenda existant réutilisé par recherche de tag : " + searchTag + " (ID: " + foundId + ")");
        return foundId;
      }
    } catch (searchErr) {
      Logger.log("Avertissement recherche événement par tag : " + searchErr);
    }

    // 4. Création du nouvel événement
    const sheetFormations = ss ? ss.getSheetByName("FORMATIONS") : null;
    let formationTitle = targetSession[13] || "Formation Leroy Merlin";
    let formationDescription = "";
    let formationCompetences = "";

    if (sheetFormations && targetSession[1]) {
      const match = targetSession[1].toString().match(/\[(.*)\]/);
      if (match) {
        const formationId = match[1];
        const lastRowForm = sheetFormations.getLastRow();
        if (lastRowForm >= 2) {
          const formationsValues = sheetFormations.getRange(2, 1, lastRowForm - 1, 7).getValues();
          for (let i = 0; i < formationsValues.length; i++) {
            if (formationsValues[i][1] === formationId) {
              formationTitle = formationsValues[i][2] || formationTitle;
              formationDescription = formationsValues[i][4] || "";
              formationCompetences = formationsValues[i][6] || "";
              break;
            }
          }
        }
      }
    }



    const sessionModuleColC = targetSession[1] || "";
    const sessionModuleColO = targetSession[13] || "";
    const cleanFormTitle = getCleanFormationTitle(sessionModuleColO, sessionModuleColC);
    const eventTitle = "Accompagnement Leroy Merlin" + (cleanFormTitle ? " - " + cleanFormTitle : "") + " [" + sessionId + "]";
    const textAgenda = getParamValue("PARAMETRE_TEXTE_AGENDA") || "";
    const connexionInfo = getParamValue("PARAMETRE_CONNEXION_1") || "";
    const meetUrl = getMeetUrlForSession(sessionId);
    const meetHeader = meetUrl ? "<p style='font-size:14px;'>📹 <b>Visioconférence Google Meet :</b> <a href='" + meetUrl + "'>" + meetUrl + "</a></p><p></p>" : "";

    const description = "<h3>👤 Inscrits et participants (Total : 0) :</h3><ul><li>Aucun participant inscrit pour le moment.</li></ul>";

    const newEvent = agenda.createEvent(eventTitle, dateDebut, dateFin, { description: description, sendInvites: true });
    newEvent.setGuestsCanInviteOthers(false).setGuestsCanModify(false).setGuestsCanSeeGuests(true);

    let generatedMeetUrl = "";
    try {
      const calendarId = agenda.getId() || "primary";
      const cleanEvtId = newEvent.getId().replace("@google.com", "");
      if (typeof (globalThis as any).Calendar !== "undefined" && (globalThis as any).Calendar.Events) {
        const patched = (globalThis as any).Calendar.Events.patch({
          conferenceData: {
            createRequest: {
              requestId: Utilities.getUuid(),
              conferenceSolutionKey: { type: "hangoutsMeet" }
            }
          }
        }, calendarId, cleanEvtId, { conferenceDataVersion: 1 });

        if (patched && patched.hangoutLink) {
          generatedMeetUrl = patched.hangoutLink;
        }
      }
    } catch (meetErr) {
      Logger.log("Information création visioconférence Google Meet : " + meetErr);
    }

    if (generatedMeetUrl) {
      try {
        const updatedDesc = "<b>📹 Visioconférence Google Meet :</b> <a href='" + generatedMeetUrl + "'>" + generatedMeetUrl + "</a><p></p>" + description;
        newEvent.setDescription(updatedDesc);
      } catch (e) {}
    }

    const newEventId = newEvent.getId();
    
    // Écrire l'ID d'événement créé uniquement dans SESSION AGENDA (pas dans SESSIONS car géré par formule)
    if (sheetSessionEvenements) {
      sheetSessionEvenements.appendRow([new Date(), sessionId, newEventId]);
    }

    Logger.log("Créneau d'accompagnement pré-réservé dans Google Agenda : " + sessionId + " (ID: " + newEventId + ")");
    return newEventId;
  } catch (err) {
    Logger.log("Erreur dans getOrCreateSessionEventId : " + err);
    return null;
  } finally {
    lock.releaseLock();
  }
}

/**
 * Ajouter un participant comme invité dans l'événement Google Agenda (Découplé et sécurisé)
 */
function addParticipantToCalendar(sessionId: string, email: string): boolean {
  if (!sessionId || !email) return false;
  try {
    const agenda = getTargetCalendar();
    if (!agenda) {
      Logger.log("Agenda introuvable pour ajouter l'invité.");
      return false;
    }

    const eventId = getOrCreateSessionEventId(sessionId);
    if (!eventId) {
      Logger.log("Impossible d'obtenir l'ID d'événement pour la session : " + sessionId);
      return false;
    }

    const event = getEventByIdRobust(eventId);

    if (event) {
      event.addGuest(email);
      Logger.log("Invité ajouté avec succès à Google Agenda : " + email + " pour session " + sessionId);

      // Déclencher l'envoi forcé du mail d'invitation Google Agenda via Calendar Advanced API
      try {
        const calendarId = agenda.getId() || "primary";
        const cleanEvtId = eventId.replace("@google.com", "");
        if (typeof (globalThis as any).Calendar !== "undefined" && (globalThis as any).Calendar.Events) {
          const currentEvt = (globalThis as any).Calendar.Events.get(calendarId, cleanEvtId);
          if (currentEvt) {
            const attendees = currentEvt.attendees || [];
            if (!attendees.some((a: any) => a.email && a.email.toLowerCase() === email.toLowerCase())) {
              attendees.push({ email: email });
            }
            (globalThis as any).Calendar.Events.patch({ attendees: attendees }, calendarId, cleanEvtId, { sendUpdates: "all" });
            Logger.log("Notification d'invitation Google Agenda envoyée (sendUpdates: all) à " + email);
          }
        }
      } catch (calApiErr) {
        Logger.log("Information envoi mise à jour Google Calendar API : " + calApiErr);
      }

      try {
        updateEventAttendeeListAndDescription(sessionId);
      } catch (updateErr) {
        Logger.log("Erreur lors de la mise à jour des détails de l'événement Agenda : " + updateErr);
      }
      return true;
    } else {
      Logger.log("Événement Agenda introuvable avec ID: " + eventId);
    }
  } catch (err) {
    Logger.log("Erreur lors de l'ajout de l'invité à l’agenda (" + email + ") : " + err);
  }
  return false;
}

/**
 * Retirer un participant d'un événement Google Agenda lors d'une désinscription
 */
function removeParticipantFromCalendar(sessionId: string, email: string): boolean {
  if (!sessionId || !email) return false;
  try {
    const agenda = getTargetCalendar();
    if (!agenda) return false;

    const eventId = getOrCreateSessionEventId(sessionId);
    if (!eventId) return false;

    const event = getEventByIdRobust(eventId);

    if (event) {
      event.removeGuest(email);
      Logger.log("Invité retiré avec succès de Google Agenda : " + email + " pour session " + sessionId);

      try {
        const calendarId = agenda.getId() || "primary";
        const cleanEvtId = eventId.replace("@google.com", "");
        if (typeof (globalThis as any).Calendar !== "undefined" && (globalThis as any).Calendar.Events) {
          const currentEvt = (globalThis as any).Calendar.Events.get(calendarId, cleanEvtId);
          if (currentEvt && currentEvt.attendees) {
            const updatedAttendees = currentEvt.attendees.filter((a: any) => !a.email || a.email.toLowerCase() !== email.toLowerCase().trim());
            (globalThis as any).Calendar.Events.patch({ attendees: updatedAttendees }, calendarId, cleanEvtId, { sendUpdates: "all" });
            Logger.log("Notification de retrait Google Agenda envoyée (sendUpdates: all) pour " + email);
          }
        }
      } catch (calApiErr) {
        Logger.log("Information envoi retrait Google Calendar API : " + calApiErr);
      }

      try {
        updateEventAttendeeListAndDescription(sessionId);
      } catch (updateErr) {
        Logger.log("Erreur lors de la mise à jour des détails de l'événement Agenda : " + updateErr);
      }
      return true;
    }
  } catch (err) {
    Logger.log("Erreur lors du retrait de l'invité agenda : " + err);
  }
  return false;
}

/**
 * Met à jour le titre et la description de l'événement Google Agenda pour refléter le nombre total de participants
 * et la liste nominative des directeurs inscrits avec leur nombre respectif.
 */
function updateEventAttendeeListAndDescription(sessionId: string): boolean {
  if (!sessionId) return false;
  try {
    const agenda = getTargetCalendar();
    if (!agenda) return false;

    const eventId = getOrCreateSessionEventId(sessionId);
    if (!eventId) return false;

    const event = getEventByIdRobust(eventId);

    if (!event) {
      Logger.log("Événement introuvable pour la mise à jour de la description : " + eventId);
      return false;
    }

    // 1. Lire toutes les inscriptions actives depuis l'onglet INSCRIPTIONS
    const sheetInscriptions = ss ? ss.getSheetByName("INSCRIPTIONS") : null;
    if (!sheetInscriptions) return false;
    const lastRowInsc = sheetInscriptions.getLastRow();
    if (lastRowInsc < 2) {
      resetEventToDefault(event, sessionId);
      return true;
    }

    const maxCols = sheetInscriptions.getLastColumn();
    const dataInsc = sheetInscriptions.getRange(2, 1, lastRowInsc - 1, maxCols).getValues();
    const targetSessionId = (sessionId || "").toString().trim().toUpperCase();
    const activeInscriptions = dataInsc.filter(row => (row[1] || "").toString().trim().toUpperCase() === targetSessionId);

    let totalParticipants = 0;
    let participantListHtml = "";
    
    activeInscriptions.forEach(row => {
      const email = (row[2] || "").toString().trim();
      const civilite = (row[3] || "").toString().trim();
      const prenom = (row[4] || "").toString().trim();
      const nom = (row[5] || "").toString().trim();
      const magasin = (row[6] || "").toString().trim();
      
      let nbPart = getNbParticipantsForEmailAndSession(sessionId, email);
      
      totalParticipants += nbPart;
      
      const displayName = (prenom || nom) ? (civilite ? civilite + " " : "") + prenom + " " + nom : email;
      const storeName = magasin ? " (" + magasin + ")" : "";
      
      participantListHtml += "<li><b>" + displayName + "</b>" + storeName + " : " + nbPart + " participant(s)</li>";
    });

    if (participantListHtml === "") {
      const sheetRaw = ss ? (ss.getSheetByName("INSCRIPTIONSS") || ss.getSheetByName("INSCRIPTIONS FORM")) : null;
      if (sheetRaw && sheetRaw.getLastRow() >= 2) {
        const rawData = sheetRaw.getRange(2, 1, sheetRaw.getLastRow() - 1, sheetRaw.getLastColumn()).getValues();
        rawData.forEach(r => {
          const rSes = (r[1] || "").toString().trim().toUpperCase();
          const rEmail = (r[2] || "").toString().trim().toLowerCase();
          if (rSes === targetSessionId && rEmail) {
            let nbPart = getNbParticipantsForEmailAndSession(sessionId, rEmail);
            totalParticipants += nbPart;
            participantListHtml += "<li><b>" + rEmail + "</b> : " + nbPart + " participant(s)</li>";
          }
        });
      }
    }

    if (participantListHtml === "") {
      participantListHtml = "<li>Aucun participant inscrit pour le moment.</li>";
    }

    // 2. Récupérer les informations de base de la session depuis SESSIONS
    const sheetSessions = ss ? ss.getSheetByName("SESSIONS") : null;
    let formationTitle = "Formation Leroy Merlin";
    let formationDescription = "";
    let formationCompetences = "";
    let cleanTitle = "Accompagnement Leroy Merlin [" + sessionId + "]";

    if (sheetSessions) {
      const lastRow = sheetSessions.getLastRow();
      if (lastRow >= 2) {
        const sessionsValues = sheetSessions.getRange(2, 2, lastRow - 1, 14).getValues();
        for (let i = 0; i < sessionsValues.length; i++) {
          if ((sessionsValues[i][0] || "").toString().trim().toUpperCase() === targetSessionId) {
            formationTitle = sessionsValues[i][13] || sessionsValues[i][1] || formationTitle;
            formationDescription = sessionsValues[i][8] || "";
            const sessionModuleColC = sessionsValues[i][1] || "";
            const sessionModuleColO = sessionsValues[i][13] || "";
            const cleanFormTitle = getCleanFormationTitle(sessionModuleColO, sessionModuleColC);
            cleanTitle = "Accompagnement Leroy Merlin" + (cleanFormTitle ? " - " + cleanFormTitle : "") + " [" + sessionId + "]";
            break;
          }
        }
      }
    }

    const textAgenda = getParamValue("PARAMETRE_TEXTE_AGENDA") || "";
    let connexionInfo = getParamValue("PARAMETRE_CONNEXION_1") || "";
    connexionInfo = connexionInfo.replace(/Connectez-vous le jour J.*?https:\/\/meet\.google\.com\/[a-z0-9-]+/gi, "").replace(/https:\/\/meet\.google\.com\/[a-z0-9-]+/gi, "").trim();

    const countTag = totalParticipants > 0 ? " (" + totalParticipants + " inscrit" + (totalParticipants > 1 ? "s" : "") + ")" : "";
    cleanTitle = cleanTitle + countTag;

    const meetUrl = getMeetUrlForSession(sessionId);
    const meetHeader = meetUrl ? "<p style='font-size:14px;'>📹 <b>Visioconférence Google Meet :</b> <a href='" + meetUrl + "'>" + meetUrl + "</a></p><p></p>" : "";

    const newDescription = "<h3>👤 Inscrits et participants (Total : " + totalParticipants + ") :</h3>"
      + "<ul>" + participantListHtml + "</ul>";

    event.setTitle(cleanTitle);
    event.setDescription(newDescription);
    
    Logger.log("Événement Agenda mis à jour avec succès (description enrichie) : " + cleanTitle);
    return true;
  } catch (err) {
    Logger.log("Erreur lors de la mise à jour des détails de l'événement Agenda : " + err);
  }
  return false;
}

/**
 * Remet la description et le titre de l'événement à leur état initial si aucune inscription
 */
function resetEventToDefault(event: GoogleAppsScript.Calendar.CalendarEvent, sessionId: string): void {
  try {
    const sheetSessions = ss ? ss.getSheetByName("SESSIONS") : null;
    let formationTitle = "Formation Leroy Merlin";
    let formationDescription = "";
    
    if (sheetSessions) {
      const lastRow = sheetSessions.getLastRow();
      if (lastRow >= 2) {
        const sessionsValues = sheetSessions.getRange(2, 2, lastRow - 1, 13).getValues();
        for (let i = 0; i < sessionsValues.length; i++) {
          if (sessionsValues[i][0] === sessionId) {
            formationTitle = sessionsValues[i][13] || formationTitle;
            formationDescription = sessionsValues[i][8] || "";
            break;
          }
        }
      }
    }
    
    const textAgenda = getParamValue("PARAMETRE_TEXTE_AGENDA") || "";
    const connexionInfo = getParamValue("PARAMETRE_CONNEXION_1") || "";
    
    const cleanFormTitle = getCleanFormationTitle(formationTitle);
    const cleanTitle = "Accompagnement Leroy Merlin" + (cleanFormTitle ? " - " + cleanFormTitle : "") + " [" + sessionId + "]";
    const newDescription = "<h3>👤 Inscrits et participants (Total : 0) :</h3><ul><li>Aucun participant inscrit pour le moment.</li></ul>";
      
    event.setTitle(cleanTitle);
    event.setDescription(newDescription);
  } catch (e) {
    Logger.log("Erreur lors de la réinitialisation de l'événement : " + e);
  }
}


/**
 * Parcourir les sessions pour créer les événements manquants
 */
function createEventSession(): void {
  const sheetSessions = ss ? ss.getSheetByName("SESSIONS") : null;
  if (!sheetSessions) return;

  const lastRow = sheetSessions.getLastRow();
  if (lastRow < 2) return;

  if (ss) ss.toast("📅 Synchronisation/Pré-réservation des événements Google Agenda...", "OUTILS", 5);

  let count = 0;
  const sessionsValues = sheetSessions.getRange(2, 2, lastRow - 1, 10).getValues();
  sessionsValues.forEach(function (row) {
    const sessionId = (row[0] || "").toString().trim();
    const dateVal = row[2]; // Col D (DATE)
    const publish = row[9]; // Col K (Publier)

    if (!sessionId || sessionId.toLowerCase().indexOf("ses-") === -1) return;

    const isPublished = Boolean(publish) && 
                      String(publish).toUpperCase() !== "FALSE" && 
                      String(publish).toUpperCase() !== "FAUX" && 
                      String(publish) !== "0" && 
                      String(publish).trim() !== "";

    let isPastDate = false;
    if (dateVal) {
      let d: Date | null = null;
      if (dateVal instanceof Date && !isNaN(dateVal.getTime())) {
        d = dateVal;
      } else if (typeof dateVal === 'string' && dateVal.trim() !== '') {
        const parts = dateVal.split('/');
        if (parts.length === 3) {
          d = new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
        } else {
          d = new Date(dateVal);
        }
      }
      if (d && !isNaN(d.getTime())) {
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        if (d.getTime() < today.getTime()) {
          isPastDate = true;
        }
      }
    }

    if (isPublished && !isPastDate) {
      const evtId = getOrCreateSessionEventId(sessionId);
      if (evtId) count++;
    }
  });

  if (ss) ss.toast("✅ " + count + " événement(s) à venir synchronisé(s) dans Google Agenda !", "OUTILS", 7);
}

/**
 * Vérifier les statuts des invités Google Agenda
 */
function checkGuests(): void {
  const sheetSessions = ss ? ss.getSheetByName("SESSIONS") : null;
  if (!sheetSessions) return;
  
  const lastRow = sheetSessions.getRange("Y1").getDataRegion().getLastRow();
  if (lastRow < 2) return;

  const eventIdsValues = sheetSessions.getRange(2, 25, lastRow - 1, 1).getValues();
  const eventIds = eventIdsValues.map(r => r[0]);

  let datas: any[][] = [];

  const agenda = getTargetCalendar();
  if (!agenda) return;

  eventIds.forEach(function (eventId) {
    if (eventId !== "") {
      try {
        const thisEvent = getEventByIdRobust(eventId);
        if (!thisEvent) return;
        const thisEventTitle = thisEvent.getTitle();
        const thisGuests = thisEvent.getGuestList();

        thisGuests.forEach(function (guest) {
          const thisStatus = guest.getGuestStatus();
          const thisEmail = guest.getEmail();
          datas.push([thisEventTitle, thisEmail, thisStatus]);
        });
      } catch (err) {
        Logger.log("Erreur lors de la vérification de l'événement " + eventId + " : " + err);
      }
    }
  });

  const sheetRecapGuests = ss ? ss.getSheetByName("RECAP REPONSES INVITATIONS") : null;
  if (sheetRecapGuests) {
    const maxRows = sheetRecapGuests.getMaxRows();
    if (maxRows > 1) {
      sheetRecapGuests.getRange(2, 1, maxRows - 1, 3).clearContent();
    }
    if (datas.length > 0) {
      sheetRecapGuests.getRange(2, 1, datas.length, 3).setValues(datas);
    }
  }
}

/**
 * Fonction de test complète pour vérifier l'intégration Google Agenda.
 * Exécutable depuis le menu NUMERICOACH > Tester l'intégration Google Agenda
 */
function testAgendaIntegration(): void {
  const userEmail = Session.getActiveUser().getEmail() || "test@example.com";
  if (ss) ss.toast("📅 Début du test Google Agenda...", "OUTILS", 5);

  try {
    // 1. Vérifier l'accès à l'agenda cible
    const agenda = getTargetCalendar();
    if (!agenda) {
      if (ss) ss.toast("❌ Aucun agenda disponible ! Vérifiez l'ID dans PARAMETRES ou les autorisations.", "OUTILS", 8);
      Logger.log("Test Agenda échec : getTargetCalendar() a retourné null.");
      return;
    }

    const calName = agenda.getName();
    const calId = agenda.getId();
    Logger.log("Agenda cible détecté : " + calName + " (ID: " + calId + ")");
    if (ss) ss.toast("✅ Agenda détecté : " + calName, "OUTILS", 5);

    // 2. Chercher une session réelle dans l'onglet SESSIONS
    const sheetSessions = ss ? ss.getSheetByName("SESSIONS") : null;
    let testSessionId = "SES-TEST";

    if (sheetSessions) {
      const lastRow = sheetSessions.getLastRow();
      if (lastRow >= 2) {
        const sessions = sheetSessions.getRange(2, 2, lastRow - 1, 1).getValues();
        for (let i = 0; i < sessions.length; i++) {
          const sid = (sessions[i][0] || "").toString().trim();
          if (sid && sid.indexOf("SES-") > -1) {
            testSessionId = sid;
            break;
          }
        }
      }
    }

    Logger.log("Session de test sélectionnée : " + testSessionId);
    if (ss) ss.toast("🔄 Test de création/récupération d'événement pour " + testSessionId + "...", "OUTILS", 5);

    // 3. Obtenir ou créer l'événement pour cette session
    const eventId = getOrCreateSessionEventId(testSessionId);
    if (!eventId) {
      if (ss) ss.toast("❌ Échec de la création/récupération de l'événement pour " + testSessionId, "OUTILS", 8);
      return;
    }

    Logger.log("Événement Agenda obtenu : " + eventId);

    // 4. Test d'ajout d'invité
    if (ss) ss.toast("👤 Test d'ajout d'invité : " + userEmail + "...", "OUTILS", 5);
    const added = addParticipantToCalendar(testSessionId, userEmail);

    if (added) {
      if (ss) ss.toast("✅ Succès ! Événement et invité (" + userEmail + ") synchronisés dans Google Agenda (" + calName + ") !", "OUTILS", 8);
      Logger.log("Test Agenda RÉUSSI avec succès pour " + userEmail + " sur " + testSessionId);
    } else {
      if (ss) ss.toast("⚠️ Événement trouvé mais échec lors de l'ajout de l'invité. Consultez les journaux.", "OUTILS", 8);
    }

  } catch (err: any) {
    Logger.log("Erreur dans testAgendaIntegration : " + err);
    if (ss) ss.toast("❌ Erreur Agenda : " + err.toString(), "OUTILS", 8);
  }
}
