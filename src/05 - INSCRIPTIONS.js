// Soumission formulaire 
function onSubmit(e) {
  // Vider le cache pour actualiser le nombre de places restantes sur le site
  try {
    clearSessionsCache();
  } catch (err) {
    Logger.log("Erreur lors du vidage du cache : " + err);
  }

  createEventSession();

  const thisResponses = e.range.getValues()[0];
  const thisSheetName = e.range.getSheet().getSheetName();

  if (thisSheetName == "INSCRIPTIONSS") {
    const thisTime = thisResponses[0];
    const thisEmail = thisResponses[9];    // Colonne J (10ème colonne, index 9)
    const thisSession = thisResponses[5];  // Colonne F (6ème colonne, index 5)
    if (!thisSession) return;
    const thisSessionid = thisSession.match(/\[(.*)\]/)[1];
    
    // VERIFIER SI DEJA INSCRIT ? 
    const maxRows = sheetInscriptions.getMaxRows();
    let verif = [];
    if (maxRows > 1) {
      const sessionsEmail = sheetInscriptions.getRange(2, 2, maxRows - 1, 2).getValues();
      verif = sessionsEmail.filter(row => (row[0] == thisSessionid && row[1] == thisEmail));
    }
    
    if (verif.length > 0) {
      // DEJA INCRIT
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
function updateFormChoices() {
  const formId = sheetParametres.getRange(pnParaCel["PARAMETRE_ID_FORMS_INSCRIPTION"]).getValue();
  if (!formId) {
    Logger.log("ID Forms Inscription non trouvé dans les paramètres.");
    return;
  }

  const form = FormApp.openById(formId);
  const items = form.getItems(FormApp.ItemType.LIST);
  let sessionItem = null;

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
  const choices = [];

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
function inscription(time, sessionId, email) {
  // ajouter l'inscription 
  sheetInscriptions.appendRow([time, sessionId, email]);
}

function reparePdf() {
  const valuesInscriptions = sheetInscription.getDataRange().getValues();
  valuesInscriptions.forEach(function (row, i) {
    if (i > 0) {
      const thisTime = row[0];
      const thisEmail = row[1];
      const thisSessionid = row[10];
      const thisConvocation = row[12];
      const thisDateSession = row[13];

      if (thisConvocation == "" && thisEmail != "") {
        // Date 1/10/2021 = 44470
        if (thisDateSession >= 44470) {
          Logger.log("à réparer : " + thisEmail + " " + thisSessionid );
          inscription(thisTime, thisSessionid, thisEmail);
          sendConvocationMail(thisSessionid, thisEmail);
        }
      }
    }
  });
}

// CONVOCATION + MAIL 
function sendConvocationMail(thisSessionId, thisEmail) {
  SpreadsheetApp.flush();
  
  // RECUPERATION DES INFORMATIONS
  const inscriptionsValues = sheetInscriptions.getDataRange().getDisplayValues();
  let thisInscriptions = inscriptionsValues.filter(row => (row[1] == thisSessionId && row[2] == thisEmail));
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
  
  const modeleConvocationId = sheetParametres.getRange(pnParaCel["PARAMETRE_ID_MODELE_CONVOC"]).getValue();
  const modeleConvocation = DriveApp.getFileById(modeleConvocationId);

  const folderConvocationId = sheetParametres.getRange(pnParaCel["PARAMETRE_ID_DOSSIER_CONVOC"]).getValue();
  const folderConvocation = DriveApp.getFolderById(folderConvocationId);

  const convocationName = "convocation " + thisSessionId + " " + thisPrenom + " " + thisNom;
  const convocation = modeleConvocation.makeCopy(convocationName, folderConvocation);
  const convocationUrl = convocation.getUrl();

  const docConvocation = DocumentApp.openByUrl(convocationUrl);
  const docConvocationBody = docConvocation.getBody();

  // remplacer les champs :
  docConvocationBody
    .replaceText("{{DATE AUJOURDHUI}}", dateAujourdhui)
    .replaceText("{{CIVILITE}}", thisCivilite)
    .replaceText("{{PRENOM}}", thisPrenom)
    .replaceText("{{NOM}}", thisNom)
    .replaceText("{{SESSION ID}}", thisSessionId)
    .replaceText("{{TITRE FORMATION}}", thisTitreFormation)
    .replaceText("{{DATE}}", thisDate)
    .replaceText("{{HEURE DEBUT}}", thisHeureDebut)
    .replaceText("{{HEURE FIN}}", thisHeureFin)
    .replaceText("{{LIEU}}", thisLieu)
    .replaceText("{{ADRESSE LIEU}}", thisLieuAdresse)
    .replaceText("{{CP}}", thisLieuCp)
    .replaceText("{{VILLE}}", thisLieuVille)
    .replaceText("{{INFOS COMPLEMENTAIRES}}", thisLieuAcces)
    .replaceText("{{DESCRIPTION}}", thisDescription)
    .replaceText("{{APPLI}}", thisAppli)
    .replaceText("{{COMPETENCES}}", thisCompetences);

  const docConvocationFooter = docConvocation.getFooter();
  if (docConvocationFooter) {
    docConvocationFooter
      .replaceText("{{DATE AUJOURDHUI}}", dateAujourdhui)
      .replaceText("{{CIVILITE}}", thisCivilite)
      .replaceText("{{PRENOM}}", thisPrenom)
      .replaceText("{{NOM}}", thisNom)
      .replaceText("{{SESSION ID}}", thisSessionId)
      .replaceText("{{TITRE FORMATION}}", thisTitreFormation)
      .replaceText("{{DATE}}", thisDate)
      .replaceText("{{HEURE DEBUT}}", thisHeureDebut)
      .replaceText("{{HEURE FIN}}", thisHeureFin)
      .replaceText("{{LIEU}}", thisLieu)
      .replaceText("{{ADRESSE LIEU}}", thisLieuAdresse)
      .replaceText("{{CP}}", thisLieuCp)
      .replaceText("{{VILLE}}", thisLieuVille)
      .replaceText("{{INFOS COMPLEMENTAIRES}}", thisLieuAcces)
      .replaceText("{{DESCRIPTION}}", thisDescription)
      .replaceText("{{APPLI}}", thisAppli)
      .replaceText("{{COMPETENCES}}", thisCompetences);
  }

  docConvocation.saveAndClose();

  let pdfConvocation = docConvocation.getAs('application/pdf');
  pdfConvocation.setName(convocationName + ".pdf");
  let pdfFileConvocation = DriveApp.createFile(pdfConvocation);
  pdfFileConvocation.moveTo(folderConvocation);
  const pdfFileUrl = pdfFileConvocation.getUrl();

  // Mettre à la corbeille le document de travail temporaire
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

  if (messageConnexion != "") {
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
    let sessionEvent = {};
    values.forEach(function (row) {
      sessionEvent[row[1]] = row[2];
    });

    const idEventAgenda = sessionEvent[thisSessionId];
    const agendaId = sheetParametres.getRange(pnParaCel["PARAMETRE_ID_AGENDA"]).getValue();
    const agenda = CalendarApp.getCalendarById(agendaId);
    let event = agenda.getEventById(idEventAgenda);
    if (event) {
      event.addGuest(thisEmail);
    }
  }
}

function creerPdfMailAgenda(i, thisRange) {
  const mailFormation = sheetParametres.getRange(pnParaCel["PARAMETRE_EMAIL"]).getValue();
  const agendaId = sheetParametres.getRange(pnParaCel["PARAMETRE_ID_AGENDA"]).getValue();

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
  const modeleConvocationId = sheetParametres.getRange(pnParaCel["PARAMETRE_ID_MODELE_CONVOC"]).getValue();
  const modeleConvocation = DriveApp.getFileById(modeleConvocationId);

  const folderConvocationId = sheetParametres.getRange(pnParaCel["PARAMETRE_ID_DOSSIER_CONVOC"]).getValue();
  const folderConvocation = DriveApp.getFolderById(folderConvocationId);

  const convocationName = "convocation " + thisSessionId + " " + thisPrenom + " " + thisNom;
  const convocation = modeleConvocation.makeCopy(convocationName, folderConvocation);
  const convocationUrl = convocation.getUrl();

  const docConvocation = DocumentApp.openByUrl(convocationUrl);
  const docConvocationBody = docConvocation.getBody();

  // Rechercher Element Image
  if (thisAppliImageLink != "") {
    let idImage = thisAppliImageLink.match(/[-\w]{25,}/);
    if (idImage) {
      let imageAppliBlob = DriveApp.getFileById(idImage[0]).getBlob();
      let imageAppli = docConvocationBody.insertImage(15, imageAppliBlob);
      let style = {};
      style[DocumentApp.Attribute.WIDTH] = 40;
      style[DocumentApp.Attribute.HEIGHT] = 40;
      style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = "CENTER";
      imageAppli.setAttributes(style);
    }
  }

  // remplacer les champs :
  docConvocationBody
    .replaceText("{{DATE AUJOURDHUI}}", dateAujourdhui)
    .replaceText("{{CIVILITE}}", thisCivilite)
    .replaceText("{{PRENOM}}", thisPrenom)
    .replaceText("{{NOM}}", thisNom)
    .replaceText("{{SESSION ID}}", thisSessionId)
    .replaceText("{{TITRE FORMATION}}", thisTitreFormation)
    .replaceText("{{DATE}}", thisDate)
    .replaceText("{{HEURE DEBUT}}", thisHeureDebut)
    .replaceText("{{HEURE FIN}}", thisHeureFin)
    .replaceText("{{LIEU}}", thisLieu)
    .replaceText("{{ADRESSE LIEU}}", thisLieuAdresse)
    .replaceText("{{CP}}", thisLieuCp)
    .replaceText("{{VILLE}}", thisLieuVille)
    .replaceText("{{INFOS COMPLEMENTAIRES}}", thisLieuAcces)
    .replaceText("{{DESCRIPTION}}", thisDescription)
    .replaceText("{{APPLI}}", thisAppli)
    .replaceText("{{COMPETENCES}}", thisCompetences);

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

  if (messageConnexion != "") {
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
    let sessionEvent = {};
    values.forEach(function (row) {
      sessionEvent[row[1]] = row[2];
    });

    const idEventAgenda = sessionEvent[thisSessionId];
    const agenda = CalendarApp.getCalendarById(agendaId);
    let event = agenda.getEventById(idEventAgenda);
    if (event) {
      event.addGuest(thisMail);
    }
  }
}
