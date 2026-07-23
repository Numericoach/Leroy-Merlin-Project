/**
 * Envoi d'un e-mail d'information et d'inscription sur Liste d'Attente lorsque la session est complète
 */
function sendWaitingListMail(sessionId: string, email: string, prenom?: string, nom?: string): void {
  const formListeAttenteId = extractFormId(getParamValue("PARAMETRE_ID_FORMS_LISTE_ATTENTE"));
  const senderEmail = getParamValue("PARAMETRE_EXPEDITEUR_EMAIL");
  const senderName = getParamValue("PARAMETRE_NOM_EXPEDITEUR") || "Formations Leroy Merlin";
  
  let entrySessionId = "entry.2116080188";
  const customEntry = getParamValue("PARAMETRE_ENTRY_SESSION");
  if (customEntry) entrySessionId = customEntry;

  let entryEmailId = "entry.193822625";
  const customEmailEntry = getParamValue("PARAMETRE_ENTRY_EMAIL");
  if (customEmailEntry) entryEmailId = customEmailEntry;

  const listeAttenteLink = "https://sites.google.com/numericoach.fr/testleroymerlin/liste-attente";

  const subject = "Session complète - Option Liste d'Attente - Formation Leroy Merlin [" + sessionId + "]";
  
  let htmlBody = "<div style='font-family: Arial, sans-serif; color: #333; max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 8px; overflow: hidden;'>"
    + "<div style='background-color: #78BE20; padding: 20px; text-align: center; color: white;'>"
    + "<h2 style='margin: 0; font-weight: bold;'>Session complète - liste d'attente</h2>"
    + "</div>"
    + "<div style='padding: 24px;'>"
    + "<p>Bonjour " + (prenom ? prenom + " " + (nom || "") : "") + ",</p>"
    + "<p>Nous avons bien reçu votre demande d'inscription à la session <b>[" + sessionId + "]</b>.</p>"
    + "<p style='color: #C0392B;'><b>Information importante :</b> Cette session est actuellement complète.</p>"
    + "<p>Afin de ne pas rater les prochaines disponibilités ou une place libérée, vous pouvez vous inscrire sur notre <b>liste d'attente</b> :</p>"
    + "<p style='text-align: center; margin: 25px 0;'><a href='" + listeAttenteLink + "' style='display:inline-block; background-color:#78BE20; color:white; padding:12px 22px; text-decoration:none; border-radius:5px; font-weight:bold;'>Rejoindre la liste d'attente</a></p>"
    + "</div></div>";

  const mailOptions: any = {
    to: email,
    subject: subject,
    htmlBody: htmlBody,
    name: senderName
  };

  if (senderEmail && senderEmail.length > 3) {
    mailOptions.from = senderEmail;
    mailOptions.replyTo = senderEmail;
  }

  try {
    MailApp.sendEmail(mailOptions);
  } catch (err) {
    Logger.log("Avertissement expéditeur personnalisé : " + err + ". Tentative d'envoi avec l'expéditeur par défaut...");
    delete mailOptions.from;
    MailApp.sendEmail(mailOptions);
  }
  Logger.log("Mail de liste d'attente envoyé à " + email + " pour la session " + sessionId);
}

/**
 * Envoi de l'e-mail de convocation / confirmation personnalisé avec le PDF complet
 */
function sendConfirmationMail(sessionId: string, email: string, prenom?: string, nom?: string, civilite?: string, nbParticipants: number = 1): void {
  const urlDesinscription = extractFormId(getParamValue("PARAMETRE_ID_FORMS_DESINSCRIPTION"));
  const senderEmail = getParamValue("PARAMETRE_EXPEDITEUR_EMAIL");
  const senderName = getParamValue("PARAMETRE_NOM_EXPEDITEUR") || "Formations Leroy Merlin";
  
  let entrySessionId = "entry.2116080188";
  const customEntry = getParamValue("PARAMETRE_ENTRY_SESSION");
  if (customEntry) entrySessionId = customEntry;

  let entryEmailId = "entry.193822625";
  const customEmailEntry = getParamValue("PARAMETRE_ENTRY_EMAIL");
  if (customEmailEntry) entryEmailId = customEmailEntry;

  const desinscriptionLink = "https://docs.google.com/forms/d/e/" + urlDesinscription + 
    "/viewform?usp=pp_url&" + entryEmailId + "=" + encodeURIComponent(email) + 
    "&" + entrySessionId + "=[" + encodeURIComponent(sessionId) + "]";

  const connexionInfo = getParamValue("PARAMETRE_CONNEXION_1") || "Lien Meet inclus dans votre invitation Agenda";

  // Récupérer les détails de la session depuis SESSIONS
  let formationTitle = "Formation Leroy Merlin";
  let dateStr = "";
  let heureDebutStr = "";
  let heureFinStr = "";
  let lieuStr = "";
  let descriptionStr = "";

  const sheetSessions = ss ? ss.getSheetByName("SESSIONS") : null;
  if (sheetSessions) {
    const lastRow = sheetSessions.getLastRow();
    if (lastRow >= 2) {
      const sessionsValues = sheetSessions.getRange(2, 2, lastRow - 1, 14).getValues();
      for (let i = 0; i < sessionsValues.length; i++) {
        if (sessionsValues[i][0] === sessionId) {
          formationTitle = sessionsValues[i][13] || formationTitle;
          if (sessionsValues[i][2]) {
            const d = new Date(sessionsValues[i][2]);
            dateStr = d.getDate() + "/" + (d.getMonth() + 1) + "/" + d.getFullYear();
          }
          if (sessionsValues[i][3]) {
            const hd = new Date(sessionsValues[i][3]);
            heureDebutStr = hd.getHours() + "h" + (hd.getMinutes() < 10 ? "0" : "") + hd.getMinutes();
          }
          if (sessionsValues[i][4]) {
            const hf = new Date(sessionsValues[i][4]);
            heureFinStr = hf.getHours() + "h" + (hf.getMinutes() < 10 ? "0" : "") + hf.getMinutes();
          }
          lieuStr = sessionsValues[i][7] || "";
          break;
        }
      }
    }
  }

  // Appeler le service PDF pour générer la pièce jointe
  const pdfResult = generateConvocationPdf(
    sessionId, email, prenom || "", nom || "", civilite || "", 
    formationTitle, dateStr, heureDebutStr, heureFinStr, lieuStr, 
    connexionInfo, descriptionStr, nbParticipants
  );

  const subject = "Convocation & confirmation d'inscription - Formation Leroy Merlin [" + sessionId + "]";
  
  let htmlBody = "<div style='font-family: Arial, sans-serif; color: #333; max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 8px; overflow: hidden;'>"
    + "<div style='background-color: #78BE20; padding: 20px; text-align: center; color: white;'>"
    + "<h2 style='margin: 0; font-weight: bold;'>Confirmation & convocation de formation</h2>"
    + "</div>"
    + "<div style='padding: 24px;'>"
    + "<p>Bonjour " + (prenom ? prenom + " " + (nom || "") : "") + ",</p>"
    + "<p>Votre inscription à la session de formation <b>" + formationTitle + " [" + sessionId + "]</b> a bien été confirmée.</p>"
    + "<p><b>Invitation Agenda :</b> Une invitation Google Agenda contenant la date, l'heure et le lien de connexion vous a été envoyée.</p>"
    + "<div style='background-color: #f4f7f6; padding: 15px; border-radius: 6px; margin: 15px 0;'>"
    + "<b>Informations de connexion :</b><br>" + connexionInfo
    + "</div>";

  if (pdfResult.pdfUrl && pdfResult.pdfUrl !== "") {
    htmlBody += "<p><a href='" + pdfResult.pdfUrl + "' style='display:inline-block; background-color:#78BE20; color:white; padding:10px 18px; text-decoration:none; border-radius:5px; font-weight:bold;'>Télécharger votre Convocation PDF</a></p>";
  }

  htmlBody += "<p style='margin-top: 25px;'><a href='" + desinscriptionLink + "' style='color:#CC3C25;'>Demander une désinscription</a></p>"
    + "</div></div>";

  const mailOptions: any = {
    to: email,
    subject: subject,
    htmlBody: htmlBody,
    name: senderName
  };

  if (senderEmail && senderEmail.length > 3) {
    mailOptions.from = senderEmail;
    mailOptions.replyTo = senderEmail;
  }

  if (pdfResult.pdfAttachment) {
    mailOptions.attachments = [pdfResult.pdfAttachment];
  }

  try {
    MailApp.sendEmail(mailOptions);
  } catch (err) {
    Logger.log("Avertissement expéditeur personnalisé : " + err + ". Tentative d'envoi avec l'expéditeur par défaut...");
    delete mailOptions.from;
    MailApp.sendEmail(mailOptions);
  }
  Logger.log("Mail de convocation envoyé à " + email + " pour la session " + sessionId);
}
