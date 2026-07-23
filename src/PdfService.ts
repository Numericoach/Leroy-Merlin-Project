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
  nbParticipants: number
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
      const body = doc.getBody();

      body.replaceText("{{DATE AUJOURDHUI}}", Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy'));
      body.replaceText("{{SESSION ID}}", sessionId);
      body.replaceText("{{EMAIL}}", email);
      body.replaceText("{{CIVILITE}}", civilite || "");
      body.replaceText("{{PRENOM}}", prenom || "");
      body.replaceText("{{NOM}}", nom || "");
      body.replaceText("{{TITRE FORMATION}}", formationTitle);
      body.replaceText("{{DATE}}", dateStr);
      body.replaceText("{{HEURE DEBUT}}", heureDebutStr);
      body.replaceText("{{HEURE FIN}}", heureFinStr);
      body.replaceText("{{LIEU}}", lieuStr);
      body.replaceText("{{CONNEXION}}", connexionInfo);
      body.replaceText("{{DESCRIPTION}}", descriptionStr);
      body.replaceText("{{APPLI}}", formationTitle);
      body.replaceText("{{NB PARTICIPANTS}}", nbParticipants.toString());
      body.replaceText("{{PARTICIPANTS}}", nbParticipants.toString());

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
