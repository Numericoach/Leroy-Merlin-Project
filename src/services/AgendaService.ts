/**
 * Helper pour obtenir l'agenda cible (ID renseigné dans PARAMETRES ou Agenda Principal par défaut)
 */
function getTargetCalendar(): GoogleAppsScript.Calendar.Calendar | null {
  const agendaId = getParamValue("PARAMETRE_ID_AGENDA");
  if (agendaId && agendaId.trim() !== "") {
    try {
      const agenda = CalendarApp.getCalendarById(agendaId.trim());
      if (agenda) return agenda;
    } catch (err) {
      Logger.log("Avertissement getCalendarById(" + agendaId + ") : " + err + ". Utilisation de l'agenda principal.");
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
 * Obtenir ou créer l'événement Google Agenda pour une session donnée (protégé par LockService)
 */
function getOrCreateSessionEventId(sessionId: string): string | null {
  if (!sessionId) return null;

  const lock = LockService.getScriptLock();
  const hasLock = lock.tryLock(10000); // Attendre max 10 secondes pour acquérir le verrou

  try {
    let sheetSessionEvenements = ss ? ss.getSheetByName("SESSION AGENDA") : null;
    if (!sheetSessionEvenements && ss) {
      try {
        sheetSessionEvenements = ss.insertSheet("SESSION AGENDA");
        sheetSessionEvenements.appendRow(["Horodateur", "ID SESSION", "ID EVENT AGENDA"]);
      } catch (e) {
        Logger.log("Erreur lors de la création de l'onglet SESSION AGENDA : " + e);
      }
    }
    if (!sheetSessionEvenements) return null;

    const agenda = getTargetCalendar();
    if (!agenda) {
      Logger.log("Aucun agenda Google disponible.");
      return null;
    }

    // 1. Vérifier si l'événement existe déjà dans la BDD Sheets
    const lastRowEvt = sheetSessionEvenements.getLastRow();
    if (lastRowEvt >= 2) {
      const sessionEvenementValues = sheetSessionEvenements.getRange(2, 1, lastRowEvt - 1, 3).getValues();
      for (let i = 0; i < sessionEvenementValues.length; i++) {
        if (sessionEvenementValues[i][1] === sessionId) {
          const storedId = sessionEvenementValues[i][2] as string;
          if (storedId) {
            // Vérifier que l'événement existe bien toujours dans Google Agenda
            try {
              let existingEvent = agenda.getEventById(storedId);
              if (!existingEvent && storedId.indexOf("@") === -1) {
                existingEvent = agenda.getEventById(storedId + "@google.com");
              }
              if (existingEvent) {
                return storedId;
              }
            } catch (checkErr) {
              Logger.log("Avertissement vérification événement existant : " + checkErr);
            }
          }
        }
      }
    }

    // 2. Si pas trouvé ou supprimé, créer l'événement dans Google Agenda
    const sheetSessions = ss ? ss.getSheetByName("SESSIONS") : null;
    if (!sheetSessions) return null;
    const lastRow = sheetSessions.getLastRow();
    if (lastRow < 2) return null;

    const sessionsValues = sheetSessions.getRange(2, 2, lastRow - 1, 14).getValues();
    let targetSession: any[] | null = null;
    for (let i = 0; i < sessionsValues.length; i++) {
      if (sessionsValues[i][0] === sessionId) {
        targetSession = sessionsValues[i];
        break;
      }
    }

    if (!targetSession) return null;

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

    const dateDebut = parseDateTime(targetSession[2], targetSession[3]);
    const dateFin = parseDateTime(targetSession[2], targetSession[4]);

    const eventTitle = "Formation Leroy Merlin - " + formationTitle + " [" + sessionId + "]";
    const textAgenda = getParamValue("PARAMETRE_TEXTE_AGENDA") || "";
    const connexionInfo = getParamValue("PARAMETRE_CONNEXION_1") || "";

    const description = textAgenda
      + "<p></p><b>" + formationDescription + "</b><p></p>"
      + connexionInfo
      + "<p>Programme de la formation :</p>" + formationCompetences;

    const newEvent = agenda.createEvent(eventTitle, dateDebut, dateFin, { description: description });
    newEvent.setGuestsCanInviteOthers(false).setGuestsCanModify(false).setGuestsCanSeeGuests(false);

    const newEventId = newEvent.getId();
    sheetSessionEvenements.appendRow([new Date(), sessionId, newEventId]);
    Logger.log("Nouvel événement créé pour la session : " + sessionId + " (ID: " + newEventId + ")");

    return newEventId;
  } catch (err) {
    Logger.log("Erreur dans getOrCreateSessionEventId : " + err);
    return null;
  } finally {
    if (hasLock) lock.releaseLock();
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

    let event: GoogleAppsScript.Calendar.CalendarEvent | null = null;
    try {
      event = agenda.getEventById(eventId);
    } catch (e) {}

    if (!event && eventId.indexOf("@") === -1) {
      try {
        event = agenda.getEventById(eventId + "@google.com");
      } catch (e) {}
    }

    if (event) {
      event.addGuest(email);
      Logger.log("Invité ajouté avec succès à Google Agenda : " + email + " pour session " + sessionId);
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

    let event: GoogleAppsScript.Calendar.CalendarEvent | null = null;
    try {
      event = agenda.getEventById(eventId);
    } catch (e) {}

    if (!event && eventId.indexOf("@") === -1) {
      try {
        event = agenda.getEventById(eventId + "@google.com");
      } catch (e) {}
    }

    if (event) {
      event.removeGuest(email);
      Logger.log("Invité retiré avec succès de Google Agenda : " + email + " pour session " + sessionId);
      return true;
    }
  } catch (err) {
    Logger.log("Erreur lors du retrait de l'invité agenda : " + err);
  }
  return false;
}

/**
 * Parcourir les sessions pour créer les événements manquants
 */
function createEventSession(): void {
  const sheetSessions = ss ? ss.getSheetByName("SESSIONS") : null;
  if (!sheetSessions) return;

  const lastRow = sheetSessions.getLastRow();
  if (lastRow < 2) return;

  const sessionsValues = sheetSessions.getRange(2, 2, lastRow - 1, 1).getValues();
  sessionsValues.forEach(function (row) {
    const sessionId = row[0];
    if (sessionId && sessionId !== "") {
      getOrCreateSessionEventId(sessionId);
    }
  });
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
        let thisEvent = agenda.getEventById(eventId);
        if (!thisEvent && eventId.indexOf("@") === -1) {
          thisEvent = agenda.getEventById(eventId + "@google.com");
        }
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
      sheetRecapGuests.deleteRows(2, maxRows - 1);
    }
    if (datas.length > 0) {
      sheetRecapGuests.getRange(2, 1, datas.length, 3).setValues(datas);
    }
  }
}
