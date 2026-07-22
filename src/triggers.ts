/**
 * Installer automatiquement les déclencheurs (Triggers) Apps Script pour le projet.
 * Cette fonction peut être exécutée depuis le menu "NUMERICOACH > Installer les déclencheurs".
 */
function setupTriggers(): void {
  const allTriggers = ScriptApp.getProjectTriggers();
  
  // 1. Nettoyer les anciens déclencheurs existants pour éviter les doublons
  allTriggers.forEach(function (trigger) {
    const handler = trigger.getHandlerFunction();
    if (handler === "onSubmit" || handler === "onEditTrigger" || handler === "onEdit") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // 2. Créer le déclencheur installable sur soumission de formulaire (onFormSubmit)
  ScriptApp.newTrigger("onSubmit")
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  // 3. Créer le déclencheur installable sur modification du Sheets (onEdit)
  ScriptApp.newTrigger("onEditTrigger")
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  Logger.log("Déclencheurs 'onSubmit' et 'onEditTrigger' installés avec succès !");
  
  try {
    const ui = SpreadsheetApp.getUi();
    if (ui) {
      ui.alert("Installation réussie", "Les déclencheurs automatiques d'inscription et de mise à jour des sessions ont été correctement configurés.", ui.ButtonSet.OK);
    }
  } catch (e) {}
}

/**
 * Fonction déclenchée lors d'une modification manuelle dans la feuille de calcul
 */
function onEditTrigger(e?: any): void {
  if (!e || !e.range) return;
  try {
    const sheetName = e.range.getSheet().getName();
    if (sheetName === "SESSIONS" || sheetName === "PARAMETRES") {
      Logger.log("Modification détectée dans : " + sheetName + " (Mise à jour du Formulaire...)");
      updateFormChoices();
    }
  } catch (err) {
    Logger.log("Erreur dans onEditTrigger : " + err);
  }
}
