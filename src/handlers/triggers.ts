/**
 * Installer automatiquement les déclencheurs (Triggers) Apps Script pour le projet.
 * Cette fonction peut être exécutée depuis le menu "NUMERICOACH > Installer les déclencheurs".
 */
function setupTriggers(): void {
  const allTriggers = ScriptApp.getProjectTriggers();
  
  // 1. Nettoyer les anciens déclencheurs existants pour éviter les doublons
  allTriggers.forEach(function (trigger) {
    const handler = trigger.getHandlerFunction();
    if (
      handler === "onSubmit" || 
      handler === "onEditTrigger" || 
      handler === "onEdit" || 
      handler === "autoUpdateFormChoicesTrigger"
    ) {
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

  // 4. Créer le déclencheur temporel récurrent (toutes les 15 minutes) pour mettre à jour les choix automatiquement
  ScriptApp.newTrigger("autoUpdateFormChoicesTrigger")
    .timeBased()
    .everyMinutes(15)
    .create();

  Logger.log("Déclencheurs 'onSubmit', 'onEditTrigger' et 'autoUpdateFormChoicesTrigger' installés avec succès !");
  if (ss) ss.toast("✅ Déclencheurs automatiques (Soumission, Modification & Temporel 15 min) installés !", "OUTILS", 7);
}

/**
 * Déclencheur temporel récurrent (toutes les 15 minutes)
 */
function autoUpdateFormChoicesTrigger(): void {
  try {
    Logger.log("Synchronisation automatique récurrente des choix dans les Google Forms...");
    updateFormChoices();
  } catch (err) {
    Logger.log("Erreur dans autoUpdateFormChoicesTrigger : " + err);
  }
}

/**
 * Fonction déclenchée lors de TOUTE modification manuelle dans la feuille de calcul
 * (suppression de ligne dans INSCRIPTIONSS, modification de SESSIONS, PARAMETRES, FILE ATTENTE, etc.)
 */
function onEditTrigger(e?: any): void {
  if (!e || !e.range) return;
  try {
    const sheetName = e.range.getSheet().getName();
    if (sheetName === "SESSIONS" || sheetName === "PARAMETRES") {
      Logger.log("Modification dans l'onglet " + sheetName + " : synchronisation des événements d'agenda et du formulaire...");
      if (sheetName === "SESSIONS") createEventSession();
      try {
        updateFormChoices();
      } catch (formErr) {
        Logger.log("Erreur lors de la mise à jour des choix du formulaire : " + formErr);
      }
    }
  } catch (err) {
    Logger.log("Erreur dans onEditTrigger : " + err);
  }
}
