/**
  * Obtenir ou créer l'événement Google Agenda pour une session donnée (protégé par LockService)
  */
function getOrCreateSessionEventId(sessionId: string): string | null {
  if (!sessionId) return null;

  const lock = LockService.getScriptLock();
  const hasLock = lock.tryLock(10000); // Attendre max 10 secondes pour acquérir le verrou

  try {
    const sheetSessionEvenements = ss ? ss.getSheetByName("SESSION AGENDA") : null;
    if (!sheetSessionEvenements) return null;

    // 1. Vérifier si l'événement existe déjà dans le BDD Sheets
    const sessionEvenementValues = sheetSessionEvenements.getDataRange().getValues();
    for (let i = 0; i < sessionEvenementValues.length; i++) {
      if (sessionEvenementValues[i][1] === sessionId) {
        return sessionEvenementValues[i][2] as string; // eventId existant
      }
    }

    // 2. Si pas trouvé, créer l'événement
    const agendaId = getParamValue("PARAMETRE_ID_AGENDA");
    if (!agendaId) {
      Logger.log("PARAMETRE_ID_AGENDA non renseigné.");
      return null;
    }

    const agenda = CalendarApp.getCalendarById(agendaId);
    if (!agenda) {
      Logger.log("Agenda introuvable ID: " + agendaId);
      return null;
    }

    const sheetSessions = ss ? ss.getSheetByName("SESSIONS") : null;
    if (!sheetSessions) return null;
    const lastRow = sheetSessions.getLastRow();
    if (lastRow < 2) return null;

    const sessionsValues = sheetSessions.getRange(2, 2, lastRow - 1, 5).getValues();
    let targetSession: any[] | null = null;
    for (let i = 0; i < sessionsValues.length; i++) {
      if (sessionsValues[i][0] === sessionId) {
        targetSession = sessionsValues[i];
        break;
      }
    }

    if (!targetSession) return null;

    const sheetFormations = ss ? ss.getSheetByName("FORMATIONS") : null;
    let formationTitle = "Formation Leroy Merlin";
    let formationDescription = "";
    let formationCompetences = "";

    if (sheetFormations && targetSession[1]) {
      const match = targetSession[1].match(/\[(.*)\]/);
      if (match) {
        const formationId = match[1];
        const formationsValues = sheetFormations.getDataRange().getValues();
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

    const thisSessionDate = new Date(targetSession[2]);
    const thisSessionHeureDebut = new Date(targetSession[3]);
    const dateDebut = new Date(
      thisSessionDate.getFullYear(),
      thisSessionDate.getMonth(),
      thisSessionDate.getDate(),
      thisSessionHeureDebut.getHours(),
      thisSessionHeureDebut.getMinutes(),
      0
    );

    const thisSessionHeureFin = new Date(targetSession[4]);
    const dateFin = new Date(
      thisSessionDate.getFullYear(),
      thisSessionDate.getMonth(),
      thisSessionDate.getDate(),
      thisSessionHeureFin.getHours(),
      thisSessionHeureFin.getMinutes(),
      0
    );

    const eventTitle = "Formation Leroy Merlin - " + formationTitle + " [" + sessionId + "]";
    const textAgenda = getParamValue("PARAMETRE_TEXTE_AGENDA");
    const connexionInfo = getParamValue("PARAMETRE_CONNEXION_1");

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
  try {
    const eventId = getOrCreateSessionEventId(sessionId);
    if (!eventId) return false;

    const agendaId = getParamValue("PARAMETRE_ID_AGENDA");
    if (!agendaId) return false;

    const agenda = CalendarApp.getCalendarById(agendaId);
    if (!agenda) return false;

    const event = agenda.getEventById(eventId);
    if (event) {
      event.addGuest(email);
      Logger.log("Invité ajouté avec succès à Google Agenda : " + email + " pour session " + sessionId);
      return true;
    }
  } catch (err) {
    Logger.log("Erreur lors de l'ajout de l'invité à l’agenda : " + err);
  }
  return false;
}

/**
 * Retirer un participant d'un événement Google Agenda lors d'une désinscription
 */
function removeParticipantFromCalendar(sessionId: string, email: string): boolean {
  try {
    const eventId = getOrCreateSessionEventId(sessionId);
    if (!eventId) return false;

    const agendaId = getParamValue("PARAMETRE_ID_AGENDA");
    if (!agendaId) return false;

    const agenda = CalendarApp.getCalendarById(agendaId);
    if (!agenda) return false;

    const event = agenda.getEventById(eventId);
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

  const agendaId = getParamValue("PARAMETRE_ID_AGENDA");
  if (!agendaId) return;

  const agenda = CalendarApp.getCalendarById(agendaId);
  if (!agenda) return;

  eventIds.forEach(function (eventId) {
    if (eventId !== "") {
      try {
        const thisEvent = agenda.getEventById(eventId);
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
