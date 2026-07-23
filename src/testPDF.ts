/**
 * Test manuel de la génération de PDF
 */
function testPDFGeneration() {
  const ui = SpreadsheetApp.getUi();
  const email = Session.getActiveUser().getEmail();
  
  try {
    const pdfResult = generateConvocationPdf(
      "TEST-001", email, "John", "Doe", "Monsieur", 
      "Formation Test", "01/01/2026", "09h00", "17h00", "En ligne", 
      "Lien de test", "Description de test", 1
    );

    if (pdfResult.pdfAttachment) {
      ui.alert("✅ SUCCÈS", "Le PDF a été généré avec succès !\nURL : " + pdfResult.pdfUrl, ui.ButtonSet.OK);
    } else {
      ui.alert("❌ ÉCHEC", "Le PDF n'a pas été généré. Aucune erreur n'a été levée, mais le retour est vide. Vérifiez les IDs dans PARAMETRES.", ui.ButtonSet.OK);
    }
  } catch (err) {
    ui.alert("❌ ERREUR CRITIQUE", "Erreur lors de la génération :\n\n" + err, ui.ButtonSet.OK);
  }
}
