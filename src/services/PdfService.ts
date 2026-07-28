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
          const todayStr = Utilities.formatDate(new Date(), 'Europe/Paris', 'dd/MM/yyyy');
          try { element.replaceText("{{DATE AUJOURDHUI}}", todayStr); } catch (e) {}
          try { element.replaceText("\\{\\{DATE AUJOURDHUI\\}\\}", todayStr); } catch (e) {}
          try { element.replaceText("{{DATE_AUJOURDHUI}}", todayStr); } catch (e) {}
          try { element.replaceText("\\{\\{DATE_AUJOURDHUI\\}\\}", todayStr); } catch (e) {}

          const targets: Record<string, string> = {
            "SESSION ID": sessionId,
            "EMAIL": email,
            "CIVILITE": civilite || "",
            "PRENOM": prenom || "",
            "NOM": nom || "",
            "TITRE FORMATION": formationTitle,
            "DATE": dateStr,
            "HEURE DEBUT": heureDebutStr,
            "HEURE FIN": heureFinStr,
            "LIEU": lieuStr,
            "CONNEXION": connexionInfo,
            "DESCRIPTION": descriptionStr,
            "APPLI": formationTitle,
            "NB PARTICIPANTS": nbParticipants.toString(),
            "PARTICIPANTS": nbParticipants.toString(),
            "ADRESSE LIEU": adresseLieu || "",
            "CP": cpLieu || "",
            "VILLE": villeLieu || "",
            "INFOS COMPLEMENTAIRES": infoCompStr || ""
          };

          for (const key in targets) {
            const val = targets[key];
            try { element.replaceText("{{" + key + "}}", val); } catch (e) {}
            try { element.replaceText("\\{\\{" + key + "\\}\\}", val); } catch (e) {}
          }
        } catch (e) {
          // ignore
        }
      };

      const body = doc.getBody();

      replaceVariables(body);
      replaceVariables(doc.getHeader());
      replaceVariables(doc.getFooter());

      try {
        const parent = body.getParent();
        for (let i = 0; i < parent.getNumChildren(); i++) {
          const child = parent.getChild(i);
          const childType = child.getType();
          if (childType === DocumentApp.ElementType.HEADER_SECTION || childType === DocumentApp.ElementType.FOOTER_SECTION) {
            replaceVariables(child);
          }
        }
      } catch (e) {}

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

/**
 * Fonction de test déclenchée depuis le menu NUMERICOACH > Tester la génération de PDF
 */
function testPDFGeneration(): void {
  const activeUserEmail = Session.getActiveUser().getEmail() || "test@example.com";
  if (ss) ss.toast("🛠️ Génération d'un PDF de test en cours...", "NUMERICOACH", 5);

  try {
    const result = generateConvocationPdf(
      "SES-TEST",
      activeUserEmail,
      "Jean",
      "DUPONT",
      "M.",
      "Formation Test Leroy Merlin",
      Utilities.formatDate(new Date(), "Europe/Paris", "dd/MM/yyyy"),
      "09h00",
      "12h00",
      "DISTANCIEL [LIEU-001]",
      "https://meet.google.com/test-meet",
      "Formation de démonstration et test de convocation PDF",
      1,
      "1 Rue de Test",
      "75001",
      "Paris",
      "Test de convocation"
    );

    if (result.pdfUrl) {
      if (ss) ss.toast("✅ PDF de test généré avec succès ! Lien dans le journal.", "NUMERICOACH", 7);
      Logger.log("PDF de test créé avec succès : " + result.pdfUrl);
    } else {
      if (ss) ss.toast("❌ Échec de génération du PDF. Vérifiez l'ID du modèle dans PARAMETRES.", "NUMERICOACH", 7);
    }
  } catch (err) {
    Logger.log("Erreur dans testPDFGeneration : " + err);
    if (ss) ss.toast("❌ Erreur : " + err, "NUMERICOACH", 7);
  }
}
