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
 * Restaure la formule dynamique ArrayFormula dans la cellule B1 de l'onglet SESSIONS.
 * Génère automatiquement les IDs "SES-0001", "SES-0002"... de manière propre et sans "SES-0000".
 */
function ensureSessionIds(): void {
  const sheetSessions = ss ? ss.getSheetByName("SESSIONS") : null;
  if (!sheetSessions) return;

  // 1. Nettoyer B2:B pour permettre à l'ARRAYFORMULA de B1 de se développer sans l'erreur "ne pas écraser les données de B2"
  const lastRow = sheetSessions.getLastRow();
  if (lastRow >= 2) {
    try {
      sheetSessions.getRange(2, 2, lastRow - 1, 1).clearContent();
    } catch (e) {}
  }

  // 2. Placer/Restaurer la formule propre ArrayFormula en B1 (ID SESSION)
  try {
    const cellB1 = sheetSessions.getRange(1, 2);
    try {
      cellB1.setFormula('=ARRAYFORMULA(IF(ROW(C:C)=1; "ID SESSION"; IF(C:C<>""; "SES-" & TEXT(ROW(C:C)-1; "0000"); "")))');
    } catch (e1) {
      cellB1.setFormula('=ARRAYFORMULA(IF(ROW(C:C)=1, "ID SESSION", IF(C:C<>"", "SES-" & TEXT(ROW(C:C)-1, "0000"), "")))');
    }
    if (ss) ss.toast("✅ Formule ArrayFormula activée (B2:B nettoyés pour laisser la formule s'étendre) !", "SESSIONS", 5);
  } catch (e) {
    Logger.log("Erreur application formule B1 : " + e);
  }

  // 3. S'assurer que G1 (DUREE) et M1 (PLACES RESTANTES) ne soient pas en #REF! ou #NAME?
  try {
    const cellG1 = sheetSessions.getRange(1, 7);
    try {
      cellG1.setFormula('=ARRAYFORMULA(IF(C:C<>""; IF(C:C=C1; "DUREE"; TEXT(F:F-E:E; "h:mm")); ""))');
    } catch (e1) {
      cellG1.setFormula('=ARRAYFORMULA(IF(C:C<>"", IF(C:C=C1, "DUREE", TEXT(F:F-E:E, "h:mm")), ""))');
    }

    const cellM1 = sheetSessions.getRange(1, 13);
    try {
      cellM1.setFormula('=ARRAYFORMULA(IF(C:C<>""; IF(C:C=C1; "PLACES RESTANTES"; H:H-L:L); ""))');
    } catch (e1) {
      cellM1.setFormula('=ARRAYFORMULA(IF(C:C<>"", IF(C:C=C1, "PLACES RESTANTES", H:H-L:L), ""))');
    }
  } catch (e) {}
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
      if (sheetName === "SESSIONS") {
        ensureSessionIds();
        createEventSession();
        try {
          processAllWaitingLists();
        } catch (waitErr) {
          Logger.log("Erreur traitement liste d'attente onEdit : " + waitErr);
        }
      }
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
