// Soumission formulaire 
function onSubmit(e?: GoogleAppsScript.Events.SheetsOnFormSubmit): void {
  if (!e || !e.range) return;

  createEventSession();

  const thisResponses = e.range.getValues()[0];
  const thisSheetName = e.range.getSheet().getSheetName();

  if (thisSheetName === "INSCRIPTIONSS") {
    const thisTime = thisResponses[0];
    const thisEmail = thisResponses[9] as string;    // Colonne J (10ème colonne, index 9)
    const thisSession = thisResponses[5] as string;  // Colonne F (6ème colonne, index 5)
    if (!thisSession) return;
    
    const match = thisSession.match(/\[(.*)\]/);
    if (!match) return;
    const thisSessionid = match[1];
    
    // VERIFIER SI DEJA INSCRIT ? 
    if (!sheetInscriptions) return;
    const maxRows = sheetInscriptions.getMaxRows();
    let verif: any[][] = [];
    if (maxRows > 1) {
      const sessionsEmail = sheetInscriptions.getRange(2, 2, maxRows - 1, 2).getValues();
      verif = sessionsEmail.filter(row => (row[0] === thisSessionid && row[1] === thisEmail));
    }
    
    if (verif.length > 0) {
      // DEJA INSCRIT
      Logger.log("Déjà inscrit : " + thisEmail + " à " + thisSessionid);
    } else {
      // INSCRIPTION 
      inscription(thisTime, thisSessionid, thisEmail);

      // Créer convocation et Envoyer le mail 
      sendConvocationMail(thisSessionid, thisEmail);
      
      // Mettre à jour les choix du formulaire
      try {
        updateFormChoices();
      } catch (err) {
        Logger.log("Erreur lors de la mise à jour des choix du formulaire : " + err);
      }
    }
  }
}

// Mettre à jour dynamiquement les choix du Google Form
function updateFormChoices(): void {
  if (!sheetParametres) return;
  const formId = sheetParametres.getRange(pnParaCel["PARAMETRE_ID_FORMS_INSCRIPTION"]).getValue() as string;
  if (!formId) {
    Logger.log("ID Forms Inscription non trouvé dans les paramètres.");
    return;
  }

  const form = FormApp.openById(formId);
  const items = form.getItems(FormApp.ItemType.LIST);
  let sessionItem: GoogleAppsScript.Forms.ListItem | null = null;

  for (let i = 0; i < items.length; i++) {
    if (items[i].getTitle().indexOf("Inscription à la formation suivante") > -1) {
      sessionItem = items[i].asListItem();
      break;
    }
  }

  if (!sessionItem) {
    Logger.log("Question 'Inscription à la formation suivante' non trouvée.");
    return;
  }

  const sheetSessions = ss.getSheetByName("SESSIONS");
  if (!sheetSessions) return;

  const lastRow = sheetSessions.getLastRow();
  if (lastRow < 2) return;

  const values = sheetSessions.getRange(2, 2, lastRow - 1, 14).getValues();
  const choices: string[] = [];

  values.forEach(function(row) {
    const sessionId = row[0];
    const publish = row[9];
    const remaining = row[11];
    const formationTitle = row[13];

    if (sessionId && publish && Number(remaining) > 0) {
      const dateObj = new Date(row[2]);
      const dateStr = dateObj.getDate() + "/" + (dateObj.getMonth() + 1) + "/" + dateObj.getFullYear();
      choices.push(formationTitle + " - " + dateStr + " [" + sessionId + "]");
    }
  });

  if (choices.length > 0) {
    sessionItem.setChoiceValues(choices);
    Logger.log("Formulaire mis à jour avec " + choices.length + " sessions.");
  } else {
    sessionItem.setChoiceValues(["Aucune session disponible pour le moment"]);
    Logger.log("Aucune session disponible.");
  }
}

// INSCRIPTION
function inscription(time: any, sessionId: string, email: string): void {
  if (sheetInscriptions) {
    sheetInscriptions.appendRow([time, sessionId, email]);
  }
}

function reparePdf(): void {
  if (!sheetInscription) return;
  const valuesInscriptions = sheetInscription.getDataRange().getValues();
  valuesInscriptions.forEach(function (row, i) {
    if (i > 0) {
      const thisTime = row[0];
      const thisEmail = row[1];
      const thisSessionid = row[10];
      const thisConvocation = row[12];
      const thisDateSession = row[13];

      if (thisConvocation === "" && thisEmail !== "") {
        if (thisDateSession >= 44470) {
          Logger.log("à réparer : " + thisEmail + " " + thisSessionid );
          inscription(thisTime, thisSessionid, thisEmail);
          sendConvocationMail(thisSessionid, thisEmail);
        }
      }
    }
  });
}

// Helper pour remplacer plusieurs variables dans un Element Google Doc
function replaceDocVariables(docElement: { replaceText(searchPattern: string, replacement: string): any }, replacements: Record<string, string>): void {
  for (const [key, value] of Object.entries(replacements)) {
    docElement.replaceText(key, value || "");
  }
}

// CONVOCATION + MAIL 
function sendConvocationMail(thisSessionId: string, thisEmail: string): void {
  SpreadsheetApp.flush();
  
  if (!sheetInscriptions || !sheetParametres) return;

  // RECUPERATION DES INFORMATIONS
  const inscriptionsValues = sheetInscriptions.getDataRange().getDisplayValues();
  let thisInscriptions = inscriptionsValues.filter(row => (row[1] === thisSessionId && row[2] === thisEmail));
  if (thisInscriptions.length === 0) return;

  const thisInscription = thisInscriptions[0];

  // FABRICATION DE LA CONVOCATION
  const dateBrute = new Date();
  Logger.log("Donnees : " + thisSessionId + " " + thisEmail + "  " + thisInscription[3]);

  const dateAujourdhui = dateBrute.getDate() + "/" + (dateBrute.getMonth() + 1) + "/" + dateBrute.getFullYear();
  const thisCivilite = thisInscription[3];
  const thisPrenom = thisInscription[4];
  const thisNom = thisInscription[5];
  const thisTitreFormation = thisInscription[7];
  const thisDate = thisInscription[8];
  const thisHeureDebut = thisInscription[9];
  const thisHeureFin = thisInscription[10];
  const thisDuree = thisInscription[11];
  const thisDescription = thisInscription[12];
  const thisAppli = thisInscription[13];
  const thisCompetences = thisInscription[14];

  const thisLieu = thisInscription[16];
  const thisLieuAcces = thisInscription[17];

  const parametresConnexion = sheetParametres.getRange(pnParaCel["PARAMETRE_CONNEXION_1"]).getValue();

  const thisLieuAdresse = thisInscription[18];
  const thisLieuCp = thisInscription[19];
  const thisLieuVille = thisInscription[20];
  
  const modeleConvocationId = sheetParametres.getRange(pnParaCel["PARAMETRE_ID_MODELE_CONVOC"]).getValue() as string;
  const modeleConvocation = DriveApp.getFileById(modeleConvocationId);

  const folderConvocationId = sheetParametres.getRange(pnParaCel["PARAMETRE_ID_DOSSIER_CONVOC"]).getValue() as string;
  const folderConvocation = DriveApp.getFolderById(folderConvocationId);

  const convocationName = "convocation " + thisSessionId + " " + thisPrenom + " " + thisNom;
  const convocation = modeleConvocation.makeCopy(convocationName, folderConvocation);
  const convocationUrl = convocation.getUrl();

  const docConvocation = DocumentApp.openByUrl(convocationUrl);
  const docConvocationBody = docConvocation.getBody();

  const replacements: Record<string, string> = {
    "{{DATE AUJOURDHUI}}": dateAujourdhui,
    "{{CIVILITE}}": thisCivilite,
    "{{PRENOM}}": thisPrenom,
    "{{NOM}}": thisNom,
    "{{SESSION ID}}": thisSessionId,
    "{{TITRE FORMATION}}": thisTitreFormation,
    "{{DATE}}": thisDate,
    "{{HEURE DEBUT}}": thisHeureDebut,
    "{{HEURE FIN}}": thisHeureFin,
    "{{LIEU}}": thisLieu,
    "{{ADRESSE LIEU}}": thisLieuAdresse,
    "{{CP}}": thisLieuCp,
    "{{VILLE}}": thisLieuVille,
    "{{INFOS COMPLEMENTAIRES}}": thisLieuAcces,
    "{{DESCRIPTION}}": thisDescription,
    "{{APPLI}}": thisAppli,
    "{{COMPETENCES}}": thisCompetences
  };

  replaceDocVariables(docConvocationBody, replacements);

  const docConvocationFooter = docConvocation.getFooter();
  if (docConvocationFooter) {
    replaceDocVariables(docConvocationFooter, replacements);
  }

  docConvocation.saveAndClose();

  let pdfConvocation = docConvocation.getAs('application/pdf');
  pdfConvocation.setName(convocationName + ".pdf");
  let pdfFileConvocation = DriveApp.createFile(pdfConvocation);
  pdfFileConvocation.moveTo(folderConvocation);
  const pdfFileUrl = pdfFileConvocation.getUrl();

  try {
    convocation.setTrashed(true);
  } catch (err) {
    Logger.log("Impossible de supprimer le document temporaire : " + err);
  }

  // CopieR L'URL DANS LA GED
  const sheetGed = ss.getSheetByName("GED");
  if (sheetGed) {
    sheetGed.appendRow([new Date(), "CONVOCATION", thisSessionId, thisEmail, pdfFileUrl, convocationName]);
  }

  // ENVOI DU MAIL 
  let messageinterne = "<h3>Bonjour " + thisCivilite + " " + thisPrenom + " " + thisNom + " </h3>"
    + "<p>Votre inscription à la formation suivante a bien été enregistrée :</p>"
    + "<p style='font-size:16px; font-weight:bold ; text-align:center;'>" + thisTitreFormation + "</p>"
    + "<p style='font-size:16px; text-align:center;'><b>" + thisDate + "</b> de <b>" + thisHeureDebut + "</b> à <b>" + thisHeureFin + "</b> en " + thisLieu + "</p>";

  const messageConnexion = sheetParametres.getRange(pnParaCel["PARAMETRE_CONNEXION_1"]).getValue();
  const urlDesinscription = sheetParametres.getRange(pnParaCel["PARAMETRE_ID_FORMS_DESINSCRIPTION"]).getValue();
  const thisDesinscriptionLink = "https://docs.google.com/forms/d/e/" + urlDesinscription + "/viewform?usp=pp_url&entry.193822625=" + thisEmail + "&entry.2116080188=" + thisTitreFormation + " " + thisDate + "[" + thisSessionId + "]";

  if (messageConnexion !== "") {
    messageinterne += "<h3>Pour participer à la formation</h3>" + messageConnexion;
  }

  messageinterne += "<p></p>"
    + "<table  border='0' cellpadding='0' cellspacing='0' width='100%' align='center' ><tr><td style='text-align:center ; vertical-align:middle; width:50%;'>"
    + "<a href='" + pdfFileUrl + "' style='display:inline-block; color:#ffffff; text-decoration:none; background-color:#1155CC; padding:15px; text-align:center; width:60%; min-width:200px; border-radius:5px; '><img  width='20' src='https://drive.google.com/uc?export=view&id=1mmefwTlbrl9fROL59li0tDGR74IXNlER' style='vertical-align:middle'> Ouvrez votre convocation </a> "
    + "</td><td style='text-align:center; vertical-align:middle;  width:50%; '>"
    + "<a href='" + thisDesinscriptionLink + "'  style='display:inline-block; color:#ffffff; text-decoration:none; background-color:#CC3C25; padding:15px; text-align:center; width:60%; min-width:200px;border-radius:5px; '><img  width='20' src='https://drive.google.com/uc?export=view&id=1y0wQDqiAmW3Ehz1zI-tihJJIYka3oZkt' style='vertical-align:middle'> Désinscription </a>"
    + "</td></tr></table>"
    + "<p></p>"
    + "<p>L’équipe vous souhaite une bonne formation.</p>";

  let message = "<div bgcolor='#EEF2F6' marginheight='0' marginwidth='0' style='font-family:Arial,sans-serif'>"
    + "<table align='center' bgcolor='#efefef' border='0' cellpadding='0' cellspacing='0' width='100%'>"
    + "<tr height='15'><td></td><td></td><td></td></tr>"
    + "<tr><td width='15'><td>"
    + "<table width='100%' border='0' cellpadding='0' cellspacing='0' style='max-width:650px  ' align='center'>"
    + "<tbody>"
    + "<tr><td>"
    + "<table id='header' border='0' cellpadding='0' cellspacing='0' style=' min-width: 100%; border-radius: 8px 8px 0 0; background-color: #1155cc; background: linear-gradient(306deg,#07C2D9 0%,#1155cc 70%);'>"
    + "<tbody><tr><td>"
    + "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    + "<tbody><tr><td>"
    + "<img  alt='numericoach' src='https://drive.google.com/uc?export=view&id=1lflKUgwpMqU8wB81CYc3UTfIKBJdgf19' style='margin:15px'>"
    + "</td>"
    + "<td style='font-size: 16px; line-height: 35px;  color: #ffffff; text-align:right ; padding:15px; vertical-align: middle ;'>"
    + "<span style='vertical-align: middle; '>Inscription confirmée <img src='https://drive.google.com/uc?export=view&id=16AdXFfqUdSaI_GGvCqT3T3clMRl1YJNW' width='20' style='vertical-align:middle'></span>"
    + "</td></tr></tbody>"
    + "</table>"
    + "</td></tr></tbody>"
    + "</table>"
    + "</td></tr>"
    + "<tr><td>"
    + "<table border='0' cellpadding='0' cellspacing='0' width='100%' bgcolor='#ffffff' style='color:#1D2E4D'><tbody><tr><td width='15'></td><td>"
    + messageinterne
    + "</td><td width='15'></td></tr></tbody></table>"
    + "</td></tr>"
    + "<tr><td>"
    + "<table id='footer' border='0' cellpadding='0' cellspacing='0' style=' min-width: 100%; border-radius:  0 0 8px 8px; background-color: #1155cc; background: linear-gradient(306deg,#07C2D9 0%,#1155cc 70%);'>"
    + "<tbody><tr><td style='color:#ffffff; padding:15px'>"
    + "<span style='vertical-align: middle; '>Numericoach - 2020 | <a href='https://www.thierryvanoffe.com' style='color:#ffffff'>Numeriblog</a></span> | <a href='https://www.youtube.com/channel/UCuVW0pl83QgawEaNepFlh7A' style='color:#ffffff'>Numeritube</a></span>"
    + "</td></tr></tbody>"
    + "</table>"
    + "</td></tr>"
    + "</tbody>"
    + "</table>"
    + "</td><td width='15'></td></tr>"
    + "<tr height='15'><td></td><td></td><td></td></tr>"
    + "<table>"
    + "</div>";

  // ENVOI DU MESSAGE
  MailApp.sendEmail({
    to: thisEmail,
    subject: thisPrenom + " " + thisNom + " : votre inscription à la formation " + thisTitreFormation + " ",
    htmlBody: message
  });

  // AJOUT INVITATION 
  const ssSessionEvent = ss.getSheetByName("SESSION AGENDA");
  if (ssSessionEvent) {
    const values = ssSessionEvent.getDataRange().getValues();
    let sessionEvent: Record<string, string> = {};
    values.forEach(function (row) {
      sessionEvent[row[1]] = row[2];
    });

    const idEventAgenda = sessionEvent[thisSessionId];
    if (!sheetParametres) return;
    const agendaId = sheetParametres.getRange(pnParaCel["PARAMETRE_ID_AGENDA"]).getValue() as string;
    const agenda = CalendarApp.getCalendarById(agendaId);
    if (agenda && idEventAgenda) {
      let event = agenda.getEventById(idEventAgenda);
      if (event) {
        event.addGuest(thisEmail);
      }
    }
  }
}

function creerPdfMailAgenda(i: number, thisRange: any[]): void {
  if (!sheetParametres || !sheetInscription) return;

  const mailFormation = sheetParametres.getRange(pnParaCel["PARAMETRE_EMAIL"]).getValue() as string;
  const agendaId = sheetParametres.getRange(pnParaCel["PARAMETRE_ID_AGENDA"]).getValue() as string;

  const thisMail = thisRange[1];
  const thisSessionId = thisRange[2];
  const thisCivilite = thisRange[3];
  const thisPrenom = thisRange[4];
  const thisNom = thisRange[5];

  const thisTitreFormation = thisRange[9];
  const thisCompetences = thisRange[21];
  const thisDescription = thisRange[22];
  const thisAppli = thisRange[23];

  const thisDate = thisRange[10];
  const thisHeureDebut = thisRange[11];
  const thisHeureFin = thisRange[12];
  const thisDuree = thisRange[13];

  const thisLieu = thisRange[15];
  const thisLieuAdresse = thisRange[16];
  const thisLieuCp = thisRange[17];
  const thisLieuVille = thisRange[18];
  const thisLieuAcces = thisRange[19];

  const thisDesinscriptionLink = thisRange[29];
  const thisAppliImageLink = thisRange[30];

  const dateAujourdhui = Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy');

  // CREATION DE LA CONVOCATION 
  const modeleConvocationId = sheetParametres.getRange(pnParaCel["PARAMETRE_ID_MODELE_CONVOC"]).getValue() as string;
  const modeleConvocation = DriveApp.getFileById(modeleConvocationId);

  const folderConvocationId = sheetParametres.getRange(pnParaCel["PARAMETRE_ID_DOSSIER_CONVOC"]).getValue() as string;
  const folderConvocation = DriveApp.getFolderById(folderConvocationId);

  const convocationName = "convocation " + thisSessionId + " " + thisPrenom + " " + thisNom;
  const convocation = modeleConvocation.makeCopy(convocationName, folderConvocation);
  const convocationUrl = convocation.getUrl();

  const docConvocation = DocumentApp.openByUrl(convocationUrl);
  const docConvocationBody = docConvocation.getBody();

  // Rechercher Element Image
  if (thisAppliImageLink && thisAppliImageLink !== "") {
    let idImage = thisAppliImageLink.match(/[-\w]{25,}/);
    if (idImage) {
      let imageAppliBlob = DriveApp.getFileById(idImage[0]).getBlob();
      let imageAppli = docConvocationBody.insertImage(15, imageAppliBlob);
      let style: Record<string, any> = {};
      style[DocumentApp.Attribute.WIDTH as any] = 40;
      style[DocumentApp.Attribute.HEIGHT as any] = 40;
      style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT as any] = DocumentApp.HorizontalAlignment.CENTER;
      imageAppli.setAttributes(style as any);
    }
  }

  const replacements: Record<string, string> = {
    "{{DATE AUJOURDHUI}}": dateAujourdhui,
    "{{CIVILITE}}": thisCivilite,
    "{{PRENOM}}": thisPrenom,
    "{{NOM}}": thisNom,
    "{{SESSION ID}}": thisSessionId,
    "{{TITRE FORMATION}}": thisTitreFormation,
    "{{DATE}}": thisDate,
    "{{HEURE DEBUT}}": thisHeureDebut,
    "{{HEURE FIN}}": thisHeureFin,
    "{{LIEU}}": thisLieu,
    "{{ADRESSE LIEU}}": thisLieuAdresse,
    "{{CP}}": thisLieuCp,
    "{{VILLE}}": thisLieuVille,
    "{{INFOS COMPLEMENTAIRES}}": thisLieuAcces,
    "{{DESCRIPTION}}": thisDescription,
    "{{APPLI}}": thisAppli,
    "{{COMPETENCES}}": thisCompetences
  };

  replaceDocVariables(docConvocationBody, replacements);

  docConvocation.saveAndClose();

  let pdfConvocation = docConvocation.getAs('application/pdf');
  pdfConvocation.setName(convocationName + ".pdf");
  let pdfFileConvocation = DriveApp.createFile(pdfConvocation);
  pdfFileConvocation.moveTo(folderConvocation);

  let pdfUrl = pdfFileConvocation.getUrl();
  sheetInscription.getRange(i, 26).setValue(pdfUrl);

  try {
    convocation.setTrashed(true);
  } catch (err) {
    Logger.log("Impossible de supprimer le document temporaire : " + err);
  }

  const messageConnexion = sheetParametres.getRange(pnParaCel["PARAMETRE_CONNEXION_1"]).getValue();

  // CREATION ET ENVOI DU MESSAGE
  let messageinterne = "<h3>Bonjour " + thisCivilite + " " + thisPrenom + " " + thisNom + " </h3>"
    + "<p>Votre inscription à la formation suivante a bien été enregistrée :</p>"
    + "<p style='font-size:16px; font-weight:bold ; text-align:center;'>" + thisTitreFormation + "</p>"
    + "<p style='font-size:16px; text-align:center;'><b>" + thisDate + "</b> de <b>" + thisHeureDebut + "</b> à <b>" + thisHeureFin + "</b> en " + thisLieu + "</p>";

  if (messageConnexion !== "") {
    messageinterne += "<h3>Pour participer à la formation</h3>" + messageConnexion;
  }

  messageinterne += "<p></p>"
    + "<table  border='0' cellpadding='0' cellspacing='0' width='100%' align='center' ><tr><td style='text-align:center ; vertical-align:middle; width:50%;'>"
    + "<a href='" + pdfUrl + "' style='display:inline-block; color:#ffffff; text-decoration:none; background-color:#1155CC; padding:15px; text-align:center; width:60%; min-width:200px; border-radius:5px; '><img  width='20' src='https://drive.google.com/uc?export=view&id=1mmefwTlbrl9fROL59li0tDGR74IXNlER' style='vertical-align:middle'> Ouvrez votre convocation </a> "
    + "</td><td style='text-align:center; vertical-align:middle;  width:50%; '>"
    + "<a href='" + thisDesinscriptionLink + "'  style='display:inline-block; color:#ffffff; text-decoration:none; background-color:#CC3C25; padding:15px; text-align:center; width:60%; min-width:200px;border-radius:5px; '><img  width='20' src='https://drive.google.com/uc?export=view&id=1y0wQDqiAmW3Ehz1zI-tihJJIYka3oZkt' style='vertical-align:middle'> Désinscription </a>"
    + "</td></tr></table>"
    + "<p></p>"
    + "<p>L’équipe vous souhaite une bonne formation.</p>";

  let message = "<div bgcolor='#EEF2F6' marginheight='0' marginwidth='0' style='font-family:Arial,sans-serif'>"
    + "<table align='center' bgcolor='#efefef' border='0' cellpadding='0' cellspacing='0' width='100%'>"
    + "<tr height='15'><td></td><td></td><td></td></tr>"
    + "<tr><td width='15'><td>"
    + "<table width='100%' border='0' cellpadding='0' cellspacing='0' style='max-width:650px  ' align='center'>"
    + "<tbody>"
    + "<tr><td>"
    + "<table id='header' border='0' cellpadding='0' cellspacing='0' style=' min-width: 100%; border-radius: 8px 8px 0 0; background-color: #1155cc; background: linear-gradient(306deg,#07C2D9 0%,#1155cc 70%);'>"
    + "<tbody><tr><td>"
    + "<table width='100%' border='0' cellpadding='0' cellspacing='0'>"
    + "<tbody><tr><td>"
    + "<img  alt='numericoach' src='https://drive.google.com/uc?export=view&id=1lflKUgwpMqU8wB81CYc3UTfIKBJdgf19' style='margin:15px'>"
    + "</td>"
    + "<td style='font-size: 16px; line-height: 35px;  color: #ffffff; text-align:right ; padding:15px; vertical-align: middle ;'>"
    + "<span style='vertical-align: middle; '>Inscription confirmée <img src='https://drive.google.com/uc?export=view&id=16AdXFfqUdSaI_GGvCqT3T3clMRl1YJNW' width='20' style='vertical-align:middle'></span>"
    + "</td></tr></tbody>"
    + "</table>"
    + "</td></tr></tbody>"
    + "</table>"
    + "</td></tr>"
    + "<tr><td>"
    + "<table border='0' cellpadding='0' cellspacing='0' width='100%' bgcolor='#ffffff' style='color:#1D2E4D'><tbody><tr><td width='15'></td><td>"
    + messageinterne
    + "</td><td width='15'></td></tr></tbody></table>"
    + "</td></tr>"
    + "<tr><td>"
    + "<table id='footer' border='0' cellpadding='0' cellspacing='0' style=' min-width: 100%; border-radius:  0 0 8px 8px; background-color: #1155cc; background: linear-gradient(306deg,#07C2D9 0%,#1155cc 70%);'>"
    + "<tbody><tr><td style='color:#ffffff; padding:15px'>"
    + "<span style='vertical-align: middle; '>Numericoach - 2020 | <a href='https://www.thierryvanoffe.com' style='color:#ffffff'>Numeriblog</a></span> | <a href='https://www.youtube.com/channel/UCuVW0pl83QgawEaNepFlh7A' style='color:#ffffff'>Numeritube</a></span>"
    + "</td></tr></tbody>"
    + "</table>"
    + "</td></tr>"
    + "</tbody>"
    + "</table>"
    + "</td><td width='15'></td></tr>"
    + "<tr height='15'><td></td><td></td><td></td></tr>"
    + "<table>"
    + "</div>";

  // ENVOI DU MESSAGE
  MailApp.sendEmail({
    to: thisMail,
    subject: thisPrenom + " " + thisNom + " : votre inscription à la formation " + thisTitreFormation + " ",
    htmlBody: message
  });

  // Recherche Id Event 
  const ssSessionEvent = ss.getSheetByName("SESSION AGENDA");
  if (ssSessionEvent) {
    const values = ssSessionEvent.getDataRange().getValues();
    let sessionEvent: Record<string, string> = {};
    values.forEach(function (row) {
      sessionEvent[row[1]] = row[2];
    });

    const idEventAgenda = sessionEvent[thisSessionId];
    const agenda = CalendarApp.getCalendarById(agendaId);
    if (agenda && idEventAgenda) {
      let event = agenda.getEventById(idEventAgenda);
      if (event) {
        event.addGuest(thisMail);
      }
    }
  }
}
