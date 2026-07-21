function createEventSession() {
  
    const agendaId = sheetParametres.getRange(pnParaCel["PARAMETRE_ID_AGENDA"]).getValue();
    const agenda = CalendarApp.getCalendarById(agendaId);
    if (!agenda) {
      Logger.log("Agenda introuvable.");
      return;
    }

    //FORMATIONS
    const sheetFormations = ss.getSheetByName("FORMATIONS");
    const formationsValues = sheetFormations.getDataRange().getValues();
    let formations = {};
    formationsValues.forEach(function (formation) {
      formations[formation[1]] = formation;
    });

    // SESSIONS AGENDA
    const sheetSessionEvenements = ss.getSheetByName("SESSION AGENDA");
    const sessionEvenementValues = sheetSessionEvenements.getDataRange().getValues();
    let sessionEvenement = {};
    sessionEvenementValues.forEach(function (value) {
      sessionEvenement[value[1]] = value[2];
    });

    // SESSIONS
    const sheetSessions = ss.getSheetByName("SESSIONS");
    const sessionsValues = sheetSessions.getRange(2, 2, sheetSessions.getLastRow() - 1, 5).getValues();

    sessionsValues.forEach(function (session ) {
      const thisSessionId = session[0];
      if (thisSessionId != "") {
        Logger.log(session);
        const thisSessionDate = new Date(session[2]);
        const thisSessionHeureDebut = new Date(session[3]);
        const thisSessionDateDebut =
          new Date(thisSessionDate.getFullYear(), thisSessionDate.getMonth(), thisSessionDate.getDate(), thisSessionHeureDebut.getHours(), thisSessionHeureDebut.getMinutes(), 0);

        const thisSessionHeureFin = new Date(session[4]);
        const thisSessionDateFin =
          new Date(thisSessionDate.getFullYear(), thisSessionDate.getMonth(), thisSessionDate.getDate(), thisSessionHeureFin.getHours(), thisSessionHeureFin.getMinutes(), 0);

        const match = session[1].match(/\[(.*)\]/);
        if (!match) return;
        const thisFormationId = match[1];
        if (!formations[thisFormationId]) return;
        const thisFormationTitle = formations[thisFormationId][2];
        const thisFormationDescription = formations[thisFormationId][4]; 
        const thisFormationCompetences = formations[thisFormationId][6];

        const thisEventTitle = "UCPA - Formation " + thisFormationTitle + " [" + thisSessionId + "]";

        const thisTextAgenda = sheetParametres.getRange(pnParaCel["PARAMETRE_TEXTE_AGENDA"]).getValue();

        const thisDescription = thisTextAgenda
          + "<p></p>"
          + "<b>"+thisFormationDescription+"</b>"
          + "<p></p>"
          +  sheetParametres.getRange(pnParaCel["PARAMETRE_CONNEXION_1"]).getValue()
          + "<p>Voici le programme de la formation du jour :</p>"
          + thisFormationCompetences
          + "<p></p>"
          + "<p>-----------</p>"
          + "Pour vous désinscrire utilisez exclusivement le lien que vous avez reçu dans le mail d'inscription.";

        const thisEventId = sessionEvenement[thisSessionId];

        if (thisEventId == undefined) {
          // CREATION
          Logger.log("Création événement");
          const newEvent = agenda.createEvent(thisEventTitle, thisSessionDateDebut, thisSessionDateFin,
            {
              description: thisDescription
            });
          newEvent.setGuestsCanInviteOthers(false).setGuestsCanModify(false).setGuestsCanSeeGuests(false);

          const newEventId = newEvent.getId();

          // AJOUT DANS LA BDD
          sheetSessionEvenements.appendRow([new Date(), thisSessionId, newEventId]);
          Logger.log("Event créé : " + thisSessionId);
        }
      }
    });
}

// VERIFIER LES STATUTS DES INVITES
function checkGuests() {
  const sheetSessions = ss.getSheetByName("SESSIONS");
  if (!sheetSessions) return;
  
  const lastRow = sheetSessions.getRange("Y1").getDataRegion().getLastRow();
  if (lastRow < 2) return;

  const eventIdsValues = sheetSessions.getRange(2, 25, lastRow - 1, 1).getValues();
  const eventIds = eventIdsValues.map(function (r) {
    return r[0];
  });

  let datas = [];

  const agendaId = sheetParametres.getRange(pnParaCel["PARAMETRE_ID_AGENDA"]).getValue();
  const agenda = CalendarApp.getCalendarById(agendaId);
  if (!agenda) return;

  eventIds.forEach(function (eventId) {
    if (eventId != "") {
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

  const sheetRecapGuests = ss.getSheetByName("RECAP REPONSES INVITATIONS");
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
