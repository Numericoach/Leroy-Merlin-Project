/**
 * Mettre à jour dynamiquement la liste des sessions disponibles dans le Google Form (Liste déroulante ou Choix multiple)
 */
function updateFormChoices(): void {
  try {
    let rawMainFormId = getParamValue("PARAMETRE_ID_FORMS_EDIT");
    if (!rawMainFormId || rawMainFormId.trim() === "" || rawMainFormId.indexOf("1FAIpQL") > -1) {
      rawMainFormId = getParamValue("PARAMETRE_ID_EDITION");
    }
    if (!rawMainFormId || rawMainFormId.trim() === "" || rawMainFormId.indexOf("1FAIpQL") > -1) {
      rawMainFormId = getParamValue("PARAMETRE_ID_FORMS_INSCRIPTION");
    }
    // Hardcoded fallback d'urgence sur l'ID d'Édition officiel si seul le lien de vue publique est présent dans PARAMETRES
    if (!rawMainFormId || rawMainFormId.indexOf("1FAIpQL") > -1) {
      rawMainFormId = "1LOsvh4spCORP-xkR8dhTFlYvmXtXRvjNnYkTrmJCbQo";
    }

    const mainFormId = extractFormId(rawMainFormId);

    let rawWaitingFormId = getParamValue("PARAMETRE_ID_FORMS_EDIT_WAITING");
    if (!rawWaitingFormId || rawWaitingFormId.indexOf("1FAIpQL") > -1) {
      rawWaitingFormId = getParamValue("PARAMETRE_ID_FORMS_LISTE_ATTENTE");
    }
    const waitingFormId = (rawWaitingFormId && rawWaitingFormId.indexOf("1FAIpQL") === -1) ? extractFormId(rawWaitingFormId) : "";

    let rawUnsubFormId = getParamValue("PARAMETRE_ID_FORMS_EDIT_UNSUB");
    if (!rawUnsubFormId || rawUnsubFormId.indexOf("1FAIpQL") > -1) {
      rawUnsubFormId = getParamValue("PARAMETRE_ID_FORMS_DESINSCRIPTION");
    }
    const unsubFormId = (rawUnsubFormId && rawUnsubFormId.indexOf("1FAIpQL") === -1) ? extractFormId(rawUnsubFormId) : "";

    if (!mainFormId) {
      Logger.log("ID Forms Inscription non trouvé dans les paramètres.");
      if (ss) ss.toast("❌ ID Formulaire d'édition non trouvé.", "OUTILS", 6);
      return;
    }

    const formsToUpdate: { id: string; type: string }[] = [];
    if (mainFormId) formsToUpdate.push({ id: mainFormId, type: "MAIN" });
    if (waitingFormId) formsToUpdate.push({ id: waitingFormId, type: "WAITING" });
    if (unsubFormId) formsToUpdate.push({ id: unsubFormId, type: "UNSUB" });

    const sheetSessions = ss ? ss.getSheetByName("SESSIONS") : null;
    if (!sheetSessions) return;

    const lastRow = sheetSessions.getLastRow();
    if (lastRow < 2) return;

    // Lecture depuis la colonne A (1) jusqu'à la colonne 20 (Col P) pour éviter tout décalage
    const values = sheetSessions.getRange(2, 1, lastRow - 1, 20).getValues();
    
    const choicesAvailable: string[] = [];
    const choicesFull: string[] = [];
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

      // Filtrer les dates passées
      let isPastDate = false;
      if (dateVal) {
        let d: Date | null = null;
        if (dateVal instanceof Date && !isNaN(dateVal.getTime())) {
          d = dateVal;
        } else if (typeof dateVal === 'string' && dateVal.trim() !== '') {
          const parts = dateVal.split('/');
          if (parts.length === 3) {
            d = new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
          } else {
            d = new Date(dateVal);
          }
        }
        if (d && !isNaN(d.getTime())) {
          const today = new Date();
          today.setHours(0, 0, 0, 0);
          if (d.getTime() < today.getTime()) {
            isPastDate = true;
          }
        }
      }

      if (sessionId && isPublished && !isPastDate) {
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

        choicesFull.push(label);
        
        if (hasSeats) {
          choicesAvailable.push(label);
          addedSessionIds.push(sessionId);
        }
      }
    });

    const targetItemId = getParamValue("PARAMETRE_ITEM_ID_SESSION");
    let totalUpdated = 0;

    formsToUpdate.forEach(formInfo => {
      try {
        if (!formInfo.id) return;
        const form = FormApp.openById(formInfo.id);
        const items = form.getItems();
        const targetItems: (GoogleAppsScript.Forms.ListItem | GoogleAppsScript.Forms.MultipleChoiceItem | GoogleAppsScript.Forms.CheckboxItem)[] = [];

        for (let i = 0; i < items.length; i++) {
          const type = items[i].getType();
          if (
            type !== FormApp.ItemType.LIST &&
            type !== FormApp.ItemType.MULTIPLE_CHOICE &&
            type !== FormApp.ItemType.CHECKBOX
          ) {
            continue;
          }

          const itemIdStr = items[i].getId().toString();
          if (targetItemId && itemIdStr === targetItemId.trim()) {
            if (type === FormApp.ItemType.LIST) targetItems.push(items[i].asListItem());
            else if (type === FormApp.ItemType.MULTIPLE_CHOICE) targetItems.push(items[i].asMultipleChoiceItem());
            else if (type === FormApp.ItemType.CHECKBOX) targetItems.push(items[i].asCheckboxItem());
            continue;
          }

          const rawTitle = items[i].getTitle();
          const title = (rawTitle || "").toString().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();

          // EXCLURE FORMELLEMENT les questions de confirmation ("Confirmez votre désinscription") et de motif
          if (
            title.indexOf("confirm") > -1 ||
            title.indexOf("raison") > -1 ||
            title.indexOf("pour quelle") > -1
          ) {
            continue;
          }

          if (
            title.indexOf("voici les sessions") > -1 ||
            title.indexOf("place") > -1 ||
            title.indexOf("selection") > -1 ||
            title.indexOf("sélection") > -1 ||
            title.indexOf("inscription") > -1 ||
            title.indexOf("session") > -1 ||
            title.indexOf("formation") > -1 ||
            title.indexOf("créneau") > -1 ||
            title.indexOf("creneau") > -1 ||
            title.indexOf("choix") > -1 ||
            title.indexOf("date") > -1 ||
            title.indexOf("suivante") > -1 ||
            title.indexOf("concernée") > -1 ||
            title.indexOf("concernee") > -1 ||
            title.indexOf("désinscription") > -1 ||
            title.indexOf("desinscription") > -1 ||
            title.indexOf("attente") > -1
          ) {
            if (type === FormApp.ItemType.LIST) targetItems.push(items[i].asListItem());
            else if (type === FormApp.ItemType.MULTIPLE_CHOICE) targetItems.push(items[i].asMultipleChoiceItem());
            else if (type === FormApp.ItemType.CHECKBOX) targetItems.push(items[i].asCheckboxItem());
          }
        }

        // Fallback si aucun titre spécifique n'a été trouvé (en conservant les exclusions)
        if (targetItems.length === 0) {
          for (let i = 0; i < items.length; i++) {
            const rawTitle = items[i].getTitle();
            const title = (rawTitle || "").toString().toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
            const type = items[i].getType();

            if (
              title.indexOf("confirm") > -1 ||
              title.indexOf("raison") > -1 ||
              title.indexOf("pour quelle") > -1
            ) {
              continue;
            }

            if (type === FormApp.ItemType.LIST) targetItems.push(items[i].asListItem());
            else if (type === FormApp.ItemType.MULTIPLE_CHOICE) targetItems.push(items[i].asMultipleChoiceItem());
            else if (type === FormApp.ItemType.CHECKBOX) targetItems.push(items[i].asCheckboxItem());
          }
        }

        const choicesToApply = (formInfo.type === "MAIN") ? choicesAvailable : choicesFull;

        if (choicesToApply.length > 0) {
          targetItems.forEach(item => {
            if (item.getType() === FormApp.ItemType.LIST) {
              (item as GoogleAppsScript.Forms.ListItem).setChoiceValues(choicesToApply);
            } else if (item.getType() === FormApp.ItemType.MULTIPLE_CHOICE) {
              (item as GoogleAppsScript.Forms.MultipleChoiceItem).setChoiceValues(choicesToApply);
            } else if (item.getType() === FormApp.ItemType.CHECKBOX) {
              (item as GoogleAppsScript.Forms.CheckboxItem).setChoiceValues(choicesToApply);
            }
          });
          totalUpdated++;
        } else {
          const defaultMsg = ["Aucune session disponible pour le moment"];
          targetItems.forEach(item => {
            if (item.getType() === FormApp.ItemType.LIST) {
              (item as GoogleAppsScript.Forms.ListItem).setChoiceValues(defaultMsg);
            } else if (item.getType() === FormApp.ItemType.MULTIPLE_CHOICE) {
              (item as GoogleAppsScript.Forms.MultipleChoiceItem).setChoiceValues(defaultMsg);
            }
          });
        }
      } catch (err) {
        Logger.log("Erreur maj formulaire " + formInfo.type + " : " + err);
      }
    });

    Logger.log("Formulaires mis à jour avec succès : " + totalUpdated);
    if (ss) ss.toast("✅ " + totalUpdated + " formulaire(s) mis à jour avec les sessions.", "OUTILS", 5);

  } catch (err) {
    Logger.log("Erreur globale dans updateFormChoices : " + err);
    if (ss) ss.toast("❌ Erreur lors de la mise à jour : " + err, "OUTILS", 7);
  }
}
