function setupTriggers() {
  var allTriggers = ScriptApp.getProjectTriggers();
  
  allTriggers.forEach(function (trigger) {
    if (trigger.getHandlerFunction() === "onSubmit" || trigger.getHandlerFunction() === "onEdit") {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("onSubmit")
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  Logger.log("Déclencheur 'onSubmit' (onFormSubmit) installé avec succès !");
  
  var ui = SpreadsheetApp.getUi();
  if (ui) {
    ui.alert("Installation réussie", "Le déclencheur automatique d'inscription a été correctement configuré.", ui.ButtonSet.OK);
  }
}
