/**
 * Installer automatiquement les déclencheurs (Triggers) Apps Script pour le projet.
 * Cette fonction peut être exécutée depuis le menu "NUMERICOACH > Installer les déclencheurs".
 */
function setupTriggers(): void {
  const allTriggers = ScriptApp.getProjectTriggers();
  
  // 1. Nettoyer les anciens déclencheurs existants pour éviter les doublons
  allTriggers.forEach(function (trigger) {
    if (trigger.getHandlerFunction() === "onSubmit" || trigger.getHandlerFunction() === "onEdit") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 2. Créer le déclencheurs installable sur soumission de formulaire (onFormSubmit)
  ScriptApp.newTrigger("onSubmit")
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  Logger.log("Déclencheur 'onSubmit' (onFormSubmit) installé avec succès !");
  
  const ui = SpreadsheetApp.getUi();
  if (ui) {
    ui.alert("Installation réussie", "Le déclencheur automatique d'inscription a été correctement configuré.", ui.ButtonSet.OK);
  }
}
