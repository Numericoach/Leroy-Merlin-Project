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
 * Assure la présence d'identifiants uniques et stables (SES-XXXX) dans la colonne B de l'onglet SESSIONS.
 * Au lieu d'utiliser une ArrayFormula qui change si les lignes sont triées ou déplacées, cette fonction
 * écrit en dur un ID persistant.
 */
function ensureSessionIds(): void {
  const sheetSessions = ss ? ss.getSheetByName("SESSIONS") : null;
  if (!sheetSessions) return;

  const lastRow = sheetSessions.getLastRow();
  if (lastRow < 2) return;

  const colId = 2; // Colonne B
  const colTitre = 3; // Colonne C

  const rangeB = sheetSessions.getRange(2, colId, lastRow - 1, 1);
  const rangeC = sheetSessions.getRange(2, colTitre, lastRow - 1, 1);

  const ids = rangeB.getValues();
  const titles = rangeC.getValues();

  let maxIdNum = 0;

  // 1. Déterminer le numéro d'ID le plus élevé existant
  for (let i = 0; i < ids.length; i++) {
    const val = (ids[i][0] || "").toString().trim();
    const match = val.match(/^SES-(\d+)$/i);
    if (match) {
      const num = parseInt(match[1], 10);
      if (num > maxIdNum) {
        maxIdNum = num;
      }
    }
  }

  let hasChanges = false;

  // 2. Générer des IDs séquentiels stables pour les lignes sans identifiants
  for (let i = 0; i < ids.length; i++) {
    const currentId = (ids[i][0] || "").toString().trim();
    const currentTitle = (titles[i][0] || "").toString().trim();

    if (currentTitle && !currentId) {
      maxIdNum++;
      const formattedNum = ("0000" + maxIdNum).slice(-4);
      ids[i][0] = `SES-${formattedNum}`;
      hasChanges = true;
    }
  }

  if (hasChanges) {
    rangeB.setValues(ids);
    if (ss) ss.toast("✅ Identifiants de session stables générés !", "SESSIONS", 5);
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
    const isSessionsOrParams = (sheetName === "SESSIONS" || sheetName === "PARAMETRES");
    const isInscriptionsOrDesinscriptions = (sheetName === "INSCRIPTIONS" || sheetName.indexOf("DESINSCRIPTION") > -1);

    if (isSessionsOrParams || isInscriptionsOrDesinscriptions) {
      Logger.log("Modification dans l'onglet " + sheetName + " : traitement des déclencheurs...");
      
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
