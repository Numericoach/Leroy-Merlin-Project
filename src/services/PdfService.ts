/**
 * Génère le document de convocation PDF pour un participant
 * @returns {GoogleAppsScript.Base.Blob | null} Le blob du PDF généré, ou null en cas d'erreur
 */
function generateConvocationPdf(
  sessionId: string,
  email: string,
  prenom: string,
  nom: string,
  civilite: string,
  formationTitle: string,
  dateStr: string,
  heureDebutStr: string,
  heureFinStr: string,
  lieuStr: string,
  connexionInfo: string,
  descriptionStr: string,
  nbParticipants: number,
  adresseLieu: string = "",
  cpLieu: string = "",
  villeLieu: string = "",
  infoCompStr: string = ""
): { pdfAttachment: GoogleAppsScript.Base.Blob | null, pdfUrl: string } {
  let pdfAttachment: GoogleAppsScript.Base.Blob | null = null;
  let pdfUrl = "";

  try {
    let modeleConvocationId = getParamValue("PARAMETRE_ID_MODELE_CONVOC");
    if (!modeleConvocationId || modeleConvocationId.length < 10) {
      modeleConvocationId = "1bJfgjsants-waATPS9Ks5C9ctPG6cu0BKopr5Ek0Gw0"; // Hardcoded fallback based on client's sheet
    }
    if (modeleConvocationId && modeleConvocationId.length > 10) {
      const modeleConvocation = DriveApp.getFileById(modeleConvocationId);
      const folderConvocationId = getParamValue("PARAMETRE_ID_DOSSIER_CONVOC");
      const folderConvocation = folderConvocationId ? DriveApp.getFolderById(folderConvocationId) : DriveApp.getRootFolder();
      
      const convocationName = "Convocation_" + sessionId + "_" + email;
      const convocationDoc = modeleConvocation.makeCopy(convocationName, folderConvocation);
      const doc = DocumentApp.openById(convocationDoc.getId());
      
      const replaceVariables = (element: any) => {
        if (!element) return;
        try {
          element.replaceText("{{DATE AUJOURDHUI}}", Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy'));
          element.replaceText("{{SESSION ID}}", sessionId);
          element.replaceText("{{EMAIL}}", email);
          element.replaceText("{{CIVILITE}}", civilite || "");
          element.replaceText("{{PRENOM}}", prenom || "");
          element.replaceText("{{NOM}}", nom || "");
          element.replaceText("{{TITRE FORMATION}}", formationTitle);
          element.replaceText("{{DATE}}", dateStr);
          element.replaceText("{{HEURE DEBUT}}", heureDebutStr);
          element.replaceText("{{HEURE FIN}}", heureFinStr);
          element.replaceText("{{LIEU}}", lieuStr);
          element.replaceText("{{CONNEXION}}", connexionInfo);
          element.replaceText("{{DESCRIPTION}}", descriptionStr);
          element.replaceText("{{APPLI}}", formationTitle);
          element.replaceText("{{NB PARTICIPANTS}}", nbParticipants.toString());
          element.replaceText("{{PARTICIPANTS}}", nbParticipants.toString());
          element.replaceText("{{ADRESSE LIEU}}", adresseLieu || "");
          element.replaceText("{{CP}}", cpLieu || "");
          element.replaceText("{{VILLE}}", villeLieu || "");
          element.replaceText("{{INFOS COMPLEMENTAIRES}}", infoCompStr || "");
        } catch (e) {
          // ignore
        }
      };

      replaceVariables(doc.getHeader());
      replaceVariables(doc.getFooter());
      replaceVariables(doc.getBody());

      doc.saveAndClose();

      pdfAttachment = convocationDoc.getAs('application/pdf');
      pdfAttachment.setName(convocationName + ".pdf");
      const pdfFile = folderConvocation.createFile(pdfAttachment);
      pdfUrl = pdfFile.getUrl();

      try { convocationDoc.setTrashed(true); } catch (e) {}
    }
  } catch (pdfErr: any) {
    Logger.log("Avertissement : la génération du PDF n'a pas pu être effectuée (envoi sans pièce jointe) : " + pdfErr);
    try {
      MailApp.sendEmail(
        email, 
        "🛠️ LEROY MERLIN - ERREUR PDF (DEBUG)", 
        "Bonjour,\nLa génération du PDF a échoué. Voici l'erreur technique capturée par le script :\n\n" + 
        pdfErr.toString() + "\n\nStack:\n" + pdfErr.stack + 
        "\n\nModele ID lu: '" + getParamValue("PARAMETRE_ID_MODELE_CONVOC") + "'"
      );
    } catch (e) {}
  }

  return { pdfAttachment, pdfUrl };
}
