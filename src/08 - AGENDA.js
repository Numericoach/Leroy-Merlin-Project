function getOrCreateSessionEventId(sessionId) {
  if (!sessionId) return null;

  var lock = LockService.getScriptLock();
  var hasLock = lock.tryLock(10000);

  try {
    var sheetSessionEvenements = ss ? ss.getSheetByName("SESSION AGENDA") : null;
    if (!sheetSessionEvenements) return null;

    var sessionEvenementValues = sheetSessionEvenements.getDataRange().getValues();
    for (var i = 0; i < sessionEvenementValues.length; i++) {
      if (sessionEvenementValues[i][1] === sessionId) {
        return sessionEvenementValues[i][2];
      }
    }

    var agendaId = getParamValue("PARAMETRE_ID_AGENDA");
    if (!agendaId) {
      Logger.log("PARAMETRE_ID_AGENDA non renseigné.");
      return null;
    }

    var agenda = CalendarApp.getCalendarById(agendaId);
    if (!agenda) {
      Logger.log("Agenda introuvable ID: " + agendaId);
      return null;
    }

    var sheetSessions = ss ? ss.getSheetByName("SESSIONS") : null;
    if (!sheetSessions) return null;
    var lastRow = sheetSessions.getLastRow();
    if (lastRow < 2) return null;

    var sessionsValues = sheetSessions.getRange(2, 2, lastRow - 1, 5).getValues();
    var targetSession = null;
    for (var j = 0; j < sessionsValues.length; j++) {
      if (sessionsValues[j][0] === sessionId) {
        targetSession = sessionsValues[j];
        break;
      }
    }

    if (!targetSession) return null;

    var sheetFormations = ss ? ss.getSheetByName("FORMATIONS") : null;
    var formationTitle = "Formation Leroy Merlin";
    var formationDescription = "";
    var formationCompetences = "";

    if (sheetFormations && targetSession[1]) {
      var match = targetSession[1].match(/\[(.*)\]/);
      if (match) {
        var formationId = match[1];
        var formationsValues = sheetFormations.getDataRange().getValues();
        for (var k = 0; k < formationsValues.length; k++) {
          if (formationsValues[k][1] === formationId) {
            formationTitle = formationsValues[k][2] || formationTitle;
            formationDescription = formationsValues[k][4] || "";
            formationCompetences = formationsValues[k][6] || "";
            break;
          }
        }
      }
    }

    var thisSessionDate = new Date(targetSession[2]);
    var thisSessionHeureDebut = new Date(targetSession[3]);
    var dateDebut = new Date(
      thisSessionDate.getFullYear(),
      thisSessionDate.getMonth(),
      thisSessionDate.getDate(),
      thisSessionHeureDebut.getHours(),
      thisSessionHeureDebut.getMinutes(),
      0
    );

    var thisSessionHeureFin = new Date(targetSession[4]);
    var dateFin = new Date(
      thisSessionDate.getFullYear(),
      thisSessionDate.getMonth(),
      thisSessionDate.getDate(),
      thisSessionHeureFin.getHours(),
      thisSessionHeureFin.getMinutes(),
      0
    );

    var eventTitle = "Formation Leroy Merlin - " + formationTitle + " [" + sessionId + "]";
    var textAgenda = getParamValue("PARAMETRE_TEXTE_AGENDA");
    var connexionInfo = getParamValue("PARAMETRE_CONNEXION_1");

    var description = textAgenda
      + "<p></p><b>" + formationDescription + "</b><p></p>"
      + connexionInfo
      + "<p>Programme de la formation :</p>" + formationCompetences;

    var newEvent = agenda.createEvent(eventTitle, dateDebut, dateFin, { description: description });
    newEvent.setGuestsCanInviteOthers(false).setGuestsCanModify(false).setGuestsCanSeeGuests(false);

    var newEventId = newEvent.getId();
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

function addParticipantToCalendar(sessionId, email) {
  try {
    var eventId = getOrCreateSessionEventId(sessionId);
    if (!eventId) return false;

    var agendaId = getParamValue("PARAMETRE_ID_AGENDA");
    if (!agendaId) return false;

    var agenda = CalendarApp.getCalendarById(agendaId);
    if (!agenda) return false;

    var event = agenda.getEventById(eventId);
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

function createEventSession() {
  var sheetSessions = ss ? ss.getSheetByName("SESSIONS") : null;
  if (!sheetSessions) return;

  var lastRow = sheetSessions.getLastRow();
  if (lastRow < 2) return;

  var sessionsValues = sheetSessions.getRange(2, 2, lastRow - 1, 1).getValues();
  sessionsValues.forEach(function (row) {
    var sessionId = row[0];
    if (sessionId && sessionId !== "") {
      getOrCreateSessionEventId(sessionId);
    }
  });
}

function checkGuests() {
  var sheetSessions = ss ? ss.getSheetByName("SESSIONS") : null;
  if (!sheetSessions) return;
  
  var lastRow = sheetSessions.getRange("Y1").getDataRegion().getLastRow();
  if (lastRow < 2) return;

  var eventIdsValues = sheetSessions.getRange(2, 25, lastRow - 1, 1).getValues();
  var eventIds = eventIdsValues.map(function(r) { return r[0]; });

  var datas = [];

  var agendaId = getParamValue("PARAMETRE_ID_AGENDA");
  if (!agendaId) return;

  var agenda = CalendarApp.getCalendarById(agendaId);
  if (!agenda) return;

  eventIds.forEach(function (eventId) {
    if (eventId !== "") {
      try {
        var thisEvent = agenda.getEventById(eventId);
        if (!thisEvent) return;
        var thisEventTitle = thisEvent.getTitle();
        var thisGuests = thisEvent.getGuestList();

        thisGuests.forEach(function (guest) {
          var thisStatus = guest.getGuestStatus();
          var thisEmail = guest.getEmail();
          datas.push([thisEventTitle, thisEmail, thisStatus]);
        });
      } catch (err) {
        Logger.log("Erreur lors de la vérification de l'événement " + eventId + " : " + err);
      }
    }
  });

  var sheetRecapGuests = ss.getSheetByName("RECAP REPONSES INVITATIONS");
  if (sheetRecapGuests) {
    var maxRows = sheetRecapGuests.getMaxRows();
    if (maxRows > 1) {
      sheetRecapGuests.deleteRows(2, maxRows - 1);
    }
    if (datas.length > 0) {
      sheetRecapGuests.getRange(2, 1, datas.length, 3).setValues(datas);
    }
  }
}
