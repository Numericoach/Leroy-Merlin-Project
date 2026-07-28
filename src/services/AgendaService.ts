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
 * Helper robuste pour récupérer un événement Google Agenda par son ID via CalendarApp
 */
function getEventByIdRobust(eventId: string): GoogleAppsScript.Calendar.CalendarEvent | null {
  if (!eventId) return null;
  let event: GoogleAppsScript.Calendar.CalendarEvent | null = null;
  try {
    event = CalendarApp.getEventById(eventId);
  } catch (e) {}

  if (!event && eventId.indexOf("@") === -1) {
    try {
      event = CalendarApp.getEventById(eventId + "@google.com");
    } catch (e) {}
  }
  return event;
}

/**
 * Nettoie le titre du module pour éviter la répétition 'Accompagnement Leroy Merlin - Formation Leroy Merlin'
 */
function getCleanFormationTitle(rawTitle: string): string {
  if (!rawTitle) return "";
  const cleaned = rawTitle
    .replace(/^Accompagnement Leroy Merlin\s*-\s*/i, "")
    .replace(/^Accompagnement Leroy Merlin/i, "")
    .replace(/^Formation Leroy Merlin\s*-\s*/i, "")
    .replace(/^Formation Leroy Merlin/i, "")
    .trim();
  return cleaned || rawTitle;
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
            const existingEvent = getEventByIdRobust(storedId);
            if (existingEvent) {
              return storedId;
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

    const cleanFormTitle = getCleanFormationTitle(formationTitle);
    const eventTitle = "Accompagnement Leroy Merlin" + (cleanFormTitle ? " - " + cleanFormTitle : "") + " [" + sessionId + "]";
    const textAgenda = getParamValue("PARAMETRE_TEXTE_AGENDA") || "";
    const connexionInfo = getParamValue("PARAMETRE_CONNEXION_1") || "";

    const description = textAgenda
      + "<p></p><b>" + formationDescription + "</b><p></p>"
      + connexionInfo
      + "<p>Programme de l'accompagnement :</p>" + formationCompetences;

    const newEvent = agenda.createEvent(eventTitle, dateDebut, dateFin, { description: description, sendInvites: true });
    newEvent.setGuestsCanInviteOthers(false).setGuestsCanModify(false).setGuestsCanSeeGuests(false);

    // Tenter la génération du lien Google Meet via l'API Calendar v3
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

    // Si un lien Google Meet a été généré, l'ajouter à la description
    if (generatedMeetUrl) {
      try {
        const updatedDesc = "<b>📹 Visioconférence Google Meet :</b> <a href='" + generatedMeetUrl + "'>" + generatedMeetUrl + "</a><p></p>" + description;
        newEvent.setDescription(updatedDesc);
      } catch (e) {}
    }

    const newEventId = newEvent.getId();
    sheetSessionEvenements.appendRow([new Date(), sessionId, newEventId]);
    Logger.log("Créneau d'accompagnement pré-réservé dans Google Agenda : " + sessionId + " (ID: " + newEventId + ")");

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

    const event = getEventByIdRobust(eventId);

    if (event) {
      event.addGuest(email);
      Logger.log("Invité ajouté avec succès à Google Agenda : " + email + " pour session " + sessionId);
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
    const activeInscriptions = dataInsc.filter(row => (row[1] || "").toString().trim() === sessionId);

    let totalParticipants = 0;
    let participantListHtml = "";
    
    activeInscriptions.forEach(row => {
      const email = (row[2] || "").toString().trim();
      const civilite = (row[3] || "").toString().trim();
      const prenom = (row[4] || "").toString().trim();
      const nom = (row[5] || "").toString().trim();
      const magasin = (row[6] || "").toString().trim();
      
      let nbPart = 1;
      if (row[7] !== undefined && row[7] !== "") {
        const parsed = parseInt(row[7].toString(), 10);
        if (!isNaN(parsed) && parsed > 0) {
          nbPart = parsed;
        }
      }
      
      totalParticipants += nbPart;
      
      const displayName = (prenom || nom) ? (civilite ? civilite + " " : "") + prenom + " " + nom : email;
      const storeName = magasin ? " (" + magasin + ")" : "";
      
      participantListHtml += "<li><b>" + displayName + "</b>" + storeName + " : " + nbPart + " participant(s) (" + email + ")</li>";
    });

    if (participantListHtml === "") {
      participantListHtml = "<li>Aucun participant inscrit pour le moment.</li>";
    }

    // 2. Récupérer les informations de base de la session depuis SESSIONS
    const sheetSessions = ss ? ss.getSheetByName("SESSIONS") : null;
    let formationTitle = "Formation Leroy Merlin";
    let formationDescription = "";
    let formationCompetences = "";

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

    // 3. Titre propre sans répétition et affichage des inscrits + lien Google Meet dans la description
    const cleanFormTitle = getCleanFormationTitle(formationTitle);
    const cleanTitle = "Accompagnement Leroy Merlin" + (cleanFormTitle ? " - " + cleanFormTitle : "") + " [" + sessionId + "]";
    
    let meetHeader = "";
    try {
      const hangout = event.getHangoutLink();
      if (hangout) {
        meetHeader = "<b>📹 Visioconférence Google Meet :</b> <a href='" + hangout + "'>" + hangout + "</a><p></p>";
      }
    } catch (e) {}

    const newDescription = meetHeader
      + textAgenda
      + "<p></p><b>" + formationDescription + "</b><p></p>"
      + connexionInfo
      + "<p>Programme de l'accompagnement :</p>" + formationCompetences
      + "<br><hr><br>"
      + "<h3>👤 Inscrits et participants (Total : " + totalParticipants + ") :</h3>"
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
    const newDescription = textAgenda
      + "<p></p><b>" + formationDescription + "</b><p></p>"
      + connexionInfo;
      
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
  if (ss) ss.toast("📅 Début du test Google Agenda...", "NUMERICOACH", 5);

  try {
    // 1. Vérifier l'accès à l'agenda cible
    const agenda = getTargetCalendar();
    if (!agenda) {
      if (ss) ss.toast("❌ Aucun agenda disponible ! Vérifiez l'ID dans PARAMETRES ou les autorisations.", "NUMERICOACH", 8);
      Logger.log("Test Agenda échec : getTargetCalendar() a retourné null.");
      return;
    }

    const calName = agenda.getName();
    const calId = agenda.getId();
    Logger.log("Agenda cible détecté : " + calName + " (ID: " + calId + ")");
    if (ss) ss.toast("✅ Agenda détecté : " + calName, "NUMERICOACH", 5);

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
    if (ss) ss.toast("🔄 Test de création/récupération d'événement pour " + testSessionId + "...", "NUMERICOACH", 5);

    // 3. Obtenir ou créer l'événement pour cette session
    const eventId = getOrCreateSessionEventId(testSessionId);
    if (!eventId) {
      if (ss) ss.toast("❌ Échec de la création/récupération de l'événement pour " + testSessionId, "NUMERICOACH", 8);
      return;
    }

    Logger.log("Événement Agenda obtenu : " + eventId);

    // 4. Test d'ajout d'invité
    if (ss) ss.toast("👤 Test d'ajout d'invité : " + userEmail + "...", "NUMERICOACH", 5);
    const added = addParticipantToCalendar(testSessionId, userEmail);

    if (added) {
      if (ss) ss.toast("✅ Succès ! Événement et invité (" + userEmail + ") synchronisés dans Google Agenda (" + calName + ") !", "NUMERICOACH", 8);
      Logger.log("Test Agenda RÉUSSI avec succès pour " + userEmail + " sur " + testSessionId);
    } else {
      if (ss) ss.toast("⚠️ Événement trouvé mais échec lors de l'ajout de l'invité. Consultez les journaux.", "NUMERICOACH", 8);
    }

  } catch (err: any) {
    Logger.log("Erreur dans testAgendaIntegration : " + err);
    if (ss) ss.toast("❌ Erreur Agenda : " + err.toString(), "NUMERICOACH", 8);
  }
}
