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

  const listeAttenteLink = "https://sites.google.com/ext.leroymerlin.fr/inscriptionenligne/liste-attente";

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
 * Envoi d'un e-mail lorsqu'une place se libère suite à une désinscription.
 */
function sendSpotAvailableMail(sessionId: string, email: string, prenom?: string, nom?: string): void {
  const senderEmail = getParamValue("PARAMETRE_EXPEDITEUR_EMAIL");
  const senderName = getParamValue("PARAMETRE_NOM_EXPEDITEUR") || "Formations Leroy Merlin";
  
  let siteUrl = getParamValue("PARAMETRE_URL_SITE_INSCRIPTION");
  if (!siteUrl || siteUrl.trim() === "") {
    siteUrl = "https://sites.google.com/ext.leroymerlin.fr/inscriptionenligne";
  }

  const subject = "Place libérée ! Inscrivez-vous vite - Formation Leroy Merlin [" + sessionId + "]";
  
  let htmlBody = "<div style='font-family: Arial, sans-serif; color: #333; max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 8px; overflow: hidden;'>"
    + "<div style='background-color: #00B140; padding: 20px; text-align: center; color: white;'>"
    + "<h2 style='margin: 0; font-weight: bold;'>🎉 Une place s'est libérée !</h2>"
    + "</div>"
    + "<div style='padding: 24px;'>"
    + "<p>Bonjour " + (prenom ? prenom + " " + (nom || "") : "") + ",</p>"
    + "<p>Bonne nouvelle ! Suite à un désistement, <b>une place vient de se libérer</b> pour la session de formation <b>[" + sessionId + "]</b> que vous attendiez.</p>"
    + "<p style='color: #00B140;'><b>Attention, premier arrivé, premier servi !</b></p>"
    + "<p>Ne tardez pas, cliquez sur le bouton ci-dessous pour retourner sur le catalogue et valider votre inscription définitive :</p>"
    + "<p style='text-align: center; margin: 25px 0;'><a href='" + siteUrl + "' style='display:inline-block; background-color:#00B140; color:white; padding:12px 22px; text-decoration:none; border-radius:5px; font-weight:bold;'>M'inscrire à la formation</a></p>"
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
}

/**
 * Envoi de l'e-mail de convocation / confirmation personnalisé avec le PDF complet
 */
function sendConfirmationMail(
  sessionId: string, 
  email: string, 
  prenom?: string, 
  nom?: string, 
  civilite?: string, 
  nbParticipants: number = 1,
  sessionDetails?: {
    formationTitle?: string,
    dateStr?: string,
    heureDebutStr?: string,
    heureFinStr?: string,
    lieuStr?: string,
    infoCompStr?: string,
    adresseLieu?: string,
    cpLieu?: string,
    villeLieu?: string
  }
): void {
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

  const actualMeetUrl = getMeetUrlForSession(sessionId);
  let connexionInfo = "";
  if (actualMeetUrl) {
    connexionInfo = "Connectez-vous le jour J, 5 minutes avant le début du webinaire sur ce lien Meet :\n" + actualMeetUrl;
  } else {
    connexionInfo = getParamValue("PARAMETRE_CONNEXION_1") || "Lien Meet inclus dans votre invitation Agenda";
  }

  // Récupérer les détails de la session (depuis sessionDetails pré-chargé ou SESSIONS)
  let formationTitle = sessionDetails?.formationTitle || "Formation Leroy Merlin";
  let dateStr = sessionDetails?.dateStr || "";
  let heureDebutStr = sessionDetails?.heureDebutStr || "";
  let heureFinStr = sessionDetails?.heureFinStr || "";
  let lieuStr = sessionDetails?.lieuStr || "";
  let infoCompStr = sessionDetails?.infoCompStr || "";
  let adresseLieu = sessionDetails?.adresseLieu || "";
  let cpLieu = sessionDetails?.cpLieu || "";
  let villeLieu = sessionDetails?.villeLieu || "";

  if (!sessionDetails) {
    const sheetSessions = ss ? ss.getSheetByName("SESSIONS") : null;
    if (sheetSessions) {
      const lastRow = sheetSessions.getLastRow();
      if (lastRow >= 2) {
        const sessionsValues = sheetSessions.getRange(2, 2, lastRow - 1, 14).getValues();
        for (let i = 0; i < sessionsValues.length; i++) {
          const currentSesId = (sessionsValues[i][0] || "").toString().trim().toUpperCase();
          const targetSesId = (sessionId || "").toString().trim().toUpperCase();
          if (currentSesId === targetSesId) {
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
            infoCompStr = sessionsValues[i][8] || "";
            break;
          }
        }
      }
    }

    const lieuMatch = lieuStr.match(/\[(.*?)\]/);
    if (lieuMatch && lieuMatch[1]) {
      const lieuId = lieuMatch[1];
      const sheetLieux = ss ? ss.getSheetByName("LIEUX") : null;
      if (sheetLieux) {
        const lieuxLastRow = sheetLieux.getLastRow();
        if (lieuxLastRow >= 2) {
          const lieuxValues = sheetLieux.getRange(2, 2, lieuxLastRow - 1, 6).getValues();
          for (let j = 0; j < lieuxValues.length; j++) {
            if (lieuxValues[j][0] === lieuId) {
              adresseLieu = lieuxValues[j][2] || "";
              cpLieu = lieuxValues[j][3] || "";
              villeLieu = lieuxValues[j][4] || "";
              break;
            }
          }
        }
      }
    }
  }

  // Appeler le service PDF pour générer la pièce jointe
  const pdfResult = generateConvocationPdf(
    sessionId, email, prenom || "", nom || "", civilite || "", 
    formationTitle, dateStr, heureDebutStr, heureFinStr, lieuStr, 
    connexionInfo, "", nbParticipants,
    adresseLieu, cpLieu, villeLieu, infoCompStr
  );

  const subject = "Convocation & confirmation d'inscription - Accompagnement Leroy Merlin [" + sessionId + "]";
  
  let htmlBody = "<div style='font-family: Arial, sans-serif; color: #333; max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 8px; overflow: hidden;'>"
    + "<div style='background-color: #78BE20; padding: 20px; text-align: center; color: white;'>"
    + "<h2 style='margin: 0; font-weight: bold;'>Confirmation & convocation d'accompagnement</h2>"
    + "</div>"
    + "<div style='padding: 24px;'>"
    + "<p>Bonjour " + (prenom ? prenom + " " + (nom || "") : "") + ",</p>"
    + "<p>Votre inscription à la session d'accompagnement <b>" + formationTitle + " [" + sessionId + "]</b> a bien été confirmée.</p>"
    + "<p><b>Invitation Agenda :</b> Une invitation Google Agenda contenant la date, l'heure et le lien de connexion vous a été envoyée.</p>"
    + "<div style='background-color: #f4f7f6; padding: 15px; border-radius: 6px; margin: 15px 0;'>"
    + "<b>Informations de connexion :</b><br>" + connexionInfo
    + "</div>";

  if (pdfResult.pdfUrl && pdfResult.pdfUrl !== "") {
    htmlBody += "<p><a href='" + pdfResult.pdfUrl + "' style='display:inline-block; background-color:#78BE20; color:white; padding:10px 18px; text-decoration:none; border-radius:5px; font-weight:bold;'>📄 Télécharger votre Convocation PDF</a></p>";
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
    const trimmedSender = senderEmail.trim();
    mailOptions.replyTo = trimmedSender;
    mailOptions.from = trimmedSender;
  }

  if (pdfResult.pdfAttachment) {
    mailOptions.attachments = [pdfResult.pdfAttachment];
  }

  try {
    MailApp.sendEmail(mailOptions);
  } catch (err) {
    Logger.log("Avertissement expéditeur personnalisé (" + senderEmail + ") : " + err + ". L'e-mail a été envoyé avec l'adresse d'exécution principale par défaut.");
    delete mailOptions.from;
    MailApp.sendEmail(mailOptions);
  }
  Logger.log("Mail de convocation envoyé à " + email + " pour la session " + sessionId);
}
