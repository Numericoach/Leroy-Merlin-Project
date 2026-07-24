/**
 * Mettre à jour dynamiquement la liste des sessions disponibles dans le Google Form (Liste déroulante ou Choix multiple)
 */
function updateFormChoices(): void {
  try {
    let rawFormId = getParamValue("PARAMETRE_ID_EDITION");
    if (!rawFormId || rawFormId.trim() === "") {
      rawFormId = getParamValue("PARAMETRE_ID_FORMS_INSCRIPTION");
    }
    const formId = extractFormId(rawFormId);
    if (!formId) {
      Logger.log("ID Forms Inscription non trouvé dans les paramètres.");
      if (ss) ss.toast("❌ ID Formulaire non trouvé dans l'onglet PARAMETRES.", "NUMERICOACH", 6);
      return;
    }

    let form: GoogleAppsScript.Forms.Form;
    try {
      form = FormApp.openById(formId);
    } catch (openErr) {
      Logger.log("Avertissement : impossible d'ouvrir le Google Form avec l'ID '" + formId + "' : " + openErr);
      if (ss) ss.toast("❌ Impossible d'ouvrir le Formulaire. Vérifiez l'ID d'édition (" + formId + ").", "NUMERICOACH", 7);
      return;
    }

    const items = form.getItems();
    const targetItems: (GoogleAppsScript.Forms.ListItem | GoogleAppsScript.Forms.MultipleChoiceItem)[] = [];

    for (let i = 0; i < items.length; i++) {
      const title = items[i].getTitle().toLowerCase();
      const type = items[i].getType();
      if (type === FormApp.ItemType.LIST || type === FormApp.ItemType.MULTIPLE_CHOICE) {
        if (
          title.indexOf("inscription") > -1 ||
          title.indexOf("session") > -1 ||
          title.indexOf("formation") > -1 ||
          title.indexOf("créneau") > -1 ||
          title.indexOf("creneau") > -1 ||
          title.indexOf("choix") > -1 ||
          title.indexOf("date") > -1
        ) {
          if (type === FormApp.ItemType.LIST) targetItems.push(items[i].asListItem());
          else targetItems.push(items[i].asMultipleChoiceItem());
        }
      }
    }

    // Fallback: si aucun item ciblé par mot-clé, prendre tous les éléments LIST / MULTIPLE_CHOICE du formulaire
    if (targetItems.length === 0) {
      for (let i = 0; i < items.length; i++) {
        const type = items[i].getType();
        if (type === FormApp.ItemType.LIST) targetItems.push(items[i].asListItem());
        else if (type === FormApp.ItemType.MULTIPLE_CHOICE) targetItems.push(items[i].asMultipleChoiceItem());
      }
    }

    if (targetItems.length === 0) {
      Logger.log("Aucun élément de type Liste déroulante ou Choix multiple trouvé dans le formulaire.");
      if (ss) ss.toast("❌ Aucune question de choix trouvée dans le Google Form.", "NUMERICOACH", 6);
      return;
    }

    const sheetSessions = ss ? ss.getSheetByName("SESSIONS") : null;
    if (!sheetSessions) return;

    const lastRow = sheetSessions.getLastRow();
    if (lastRow < 2) return;

    // Lecture depuis la colonne A (1) jusqu'à la colonne 20 (Col P) pour éviter tout décalage
    const values = sheetSessions.getRange(2, 1, lastRow - 1, 20).getValues();
    const choices: string[] = [];
    const addedSessionIds: string[] = [];

    values.forEach(function(row) {
      const sessionId = (row[1] || "").toString().trim(); // Col B (ID SESSION)
      const formationCol = (row[2] || "").toString().trim(); // Col C (FORMATION)
      const dateVal = row[3]; // Col D (DATE)
      const heureDebutVal = row[4]; // Col E (HEURE DEBUT)
      const heureFinVal = row[5]; // Col F (HEURE FIN)
      const infoComp = (row[9] || "").toString().trim(); // Col J (INFORMATION COMPLEMENTAIRE)
      const publish = row[10]; // Col K (Publier)
      const rawRemaining = parseFloat(String(row[12] || "").replace(",", ".")); // Col M (Places Restantes)
      const nbPlaces = parseFloat(String(row[7] || "").replace(",", ".")); // Col H (NB DE PLACES)
      const nbInscrits = parseFloat(String(row[11] || "").replace(",", ".")); // Col L (Nb Inscrits)
      const moduleTitle = row[14] || formationCol || "Formation"; // Col O (MODULE TITRE)

      if (!sessionId || sessionId.toLowerCase().indexOf("ses-") === -1) {
        return;
      }

      // Est publié si coché ou non-vide / non-false
      const isPublished = Boolean(publish) && 
                        String(publish).toUpperCase() !== "FALSE" && 
                        String(publish).toUpperCase() !== "FAUX" && 
                        String(publish) !== "0" && 
                        String(publish).trim() !== "";

      // Calcul robuste des places restantes : utilise Col M ou (Col H - Col L)
      let remainingSeats = 0;
      if (!isNaN(rawRemaining)) {
        remainingSeats = rawRemaining;
      } else if (!isNaN(nbPlaces)) {
        const registered = !isNaN(nbInscrits) ? nbInscrits : 0;
        remainingSeats = nbPlaces - registered;
      } else {
        remainingSeats = 1; // Fallback par défaut si non renseigné
      }

      const hasSeats = remainingSeats > 0;

      if (sessionId && isPublished && hasSeats) {
        const dateStr = formatDateClean(dateVal);
        const heureDebutStr = formatTimeClean(heureDebutVal);
        const heureFinStr = formatTimeClean(heureFinVal);

        let compClean = infoComp;
        if (compClean.toLowerCase().indexOf("avec pratique") > -1) {
          compClean = "(AVEC PRATIQUE)";
        } else if (compClean.toLowerCase().indexOf("sans pratique") > -1) {
          compClean = "(SANS PRATIQUE)";
        } else if (compClean.length > 0) {
          compClean = "(" + compClean + ")";
        }

        let titleClean = moduleTitle.toString().toUpperCase();
        if (titleClean.indexOf("[") > -1) {
          titleClean = titleClean.split("[")[0].trim();
        }

        let label = titleClean;
        if (compClean) {
          label += " " + compClean;
        }
        if (dateStr) {
          label += " - " + dateStr;
        }
        if (heureDebutStr && heureFinStr) {
          label += " de " + heureDebutStr + " à " + heureFinStr;
        }
        label += " [" + sessionId + "]";

        choices.push(label);
        addedSessionIds.push(sessionId);
      }
    });

    if (choices.length > 0) {
      targetItems.forEach(item => {
        if (item.getType() === FormApp.ItemType.LIST) {
          (item as GoogleAppsScript.Forms.ListItem).setChoiceValues(choices);
        } else if (item.getType() === FormApp.ItemType.MULTIPLE_CHOICE) {
          (item as GoogleAppsScript.Forms.MultipleChoiceItem).setChoiceValues(choices);
        }
      });
      Logger.log("Formulaire mis à jour avec " + choices.length + " sessions : " + addedSessionIds.join(", "));
      if (ss) ss.toast("✅ Formulaire mis à jour avec " + choices.length + " session(s) (" + addedSessionIds.join(", ") + ")", "NUMERICOACH", 5);
    } else {
      const defaultMsg = ["Aucune session disponible pour le moment"];
      targetItems.forEach(item => {
        if (item.getType() === FormApp.ItemType.LIST) {
          (item as GoogleAppsScript.Forms.ListItem).setChoiceValues(defaultMsg);
        } else if (item.getType() === FormApp.ItemType.MULTIPLE_CHOICE) {
          (item as GoogleAppsScript.Forms.MultipleChoiceItem).setChoiceValues(defaultMsg);
        }
      });
      Logger.log("Aucune session disponible.");
      if (ss) ss.toast("⚠️ Aucune session disponible. Formulaire réinitialisé.", "NUMERICOACH", 5);
    }
  } catch (err) {
    Logger.log("Erreur dans updateFormChoices : " + err);
    if (ss) ss.toast("❌ Erreur lors de la mise à jour : " + err, "NUMERICOACH", 7);
  }
}
