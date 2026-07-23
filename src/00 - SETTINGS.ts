// Configuration globale et variables réutilisables
const ss: GoogleAppsScript.Spreadsheet.Spreadsheet | null = typeof SpreadsheetApp !== 'undefined' && SpreadsheetApp.getActiveSpreadsheet ? SpreadsheetApp.getActiveSpreadsheet() : null;
const sheetParametres: GoogleAppsScript.Spreadsheet.Sheet | null = ss ? ss.getSheetByName("PARAMETRES") : null;
const sheetInscriptions: GoogleAppsScript.Spreadsheet.Sheet | null = ss ? ss.getSheetByName("INSCRIPTIONS") : null; 

/**
 * Récupérer la valeur d'un paramètre dans la feuille PARAMETRES (Recherche dynamique et flexible)
 */
function getParamValue(paramKey: string): string {
  const activeSs = ss || (typeof SpreadsheetApp !== 'undefined' && SpreadsheetApp.getActiveSpreadsheet ? SpreadsheetApp.getActiveSpreadsheet() : null);
  const sheetParam = activeSs ? activeSs.getSheetByName("PARAMETRES") : null;
  if (!sheetParam) return "";

  const lastRow = sheetParam.getLastRow();
  if (lastRow < 2) return "";

  const data = sheetParam.getRange(2, 1, lastRow - 1, 2).getValues();
  const searchClean = paramKey.replace(/^PARAMETRE_/i, "").replace(/_/g, " ").trim().toLowerCase();

  for (let i = 0; i < data.length; i++) {
    const rawKey = (data[i][0] || "").toString().trim();
    if (!rawKey) continue;

    const keyClean = rawKey.replace(/^PARAMETRE_/i, "").replace(/_/g, " ").replace(/[>]/g, "").trim().toLowerCase();

    if (rawKey === paramKey || keyClean === searchClean || rawKey.toLowerCase() === paramKey.toLowerCase()) {
      return data[i][1] ? data[i][1].toString().trim() : "";
    }
    
    // Fallbacks très permissifs pour les paramètres souvent mal orthographiés ou avec apostrophes/accents
    const lowerRaw = rawKey.toLowerCase();
    if (paramKey === "PARAMETRE_ID_FORMS_LISTE_ATTENTE" && lowerRaw.indexOf("liste d'attente") > -1) {
      return data[i][1] ? data[i][1].toString().trim() : "";
    }
    if (paramKey === "PARAMETRE_ID_FORMS_DESINSCRIPTION" && lowerRaw.indexOf("désinscription") > -1) {
      return data[i][1] ? data[i][1].toString().trim() : "";
    }
    if (paramKey === "PARAMETRE_ENTRY_SESSION" && lowerRaw.indexOf("entry session") > -1) {
      return data[i][1] ? data[i][1].toString().trim() : "";
    }
    if (paramKey === "PARAMETRE_ENTRY_EMAIL" && lowerRaw.indexOf("entry email") > -1) {
      return data[i][1] ? data[i][1].toString().trim() : "";
    }
    if (paramKey === "PARAMETRE_ID_MODELE_CONVOC" && lowerRaw.indexOf("modèle de convocation") > -1) {
      return data[i][1] ? data[i][1].toString().trim() : "";
    }
    if (paramKey === "PARAMETRE_ID_DOSSIER_CONVOC" && lowerRaw.indexOf("dossier") > -1 && lowerRaw.indexOf("convocation") > -1) {
      return data[i][1] ? data[i][1].toString().trim() : "";
    }
    if (paramKey === "PARAMETRE_ID_AGENDA" && lowerRaw.indexOf("id agenda") > -1) {
      return data[i][1] ? data[i][1].toString().trim() : "";
    }
  }
  return "";
}

/**
 * Menu personnalisé dans Google Sheets
 */
function onOpen(): void {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('NUMERICOACH')
    .addItem('Mettre à jour les sessions dans le Formulaire', 'updateFormChoices')
    .addItem('Installer les déclencheurs (Triggers)', 'setupTriggers')
    .addToUi();
}

/**
 * Parser une date et une heure provenant de Google Sheets
 */
function parseDateTime(dateVal: Date | string | number | null | undefined, timeVal: Date | string | number | null | undefined): Date {
  let d = new Date();
  if (dateVal instanceof Date) {
    d = new Date(dateVal.getTime());
  } else if (typeof dateVal === 'string' && dateVal.trim() !== '') {
    const parts = dateVal.split('/');
    if (parts.length === 3) {
      d = new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
    } else {
      d = new Date(dateVal);
    }
  }

  let hours = 9;
  let minutes = 0;

  if (timeVal instanceof Date) {
    hours = timeVal.getHours();
    minutes = timeVal.getMinutes();
  } else if (typeof timeVal === 'string' && timeVal.trim() !== '') {
    const timeParts = timeVal.split(':');
    if (timeParts.length >= 2) {
      hours = parseInt(timeParts[0], 10);
      minutes = parseInt(timeParts[1], 10);
    }
  } else if (typeof timeVal === 'number') {
    const totalMinutes = Math.round(timeVal * 24 * 60);
    hours = Math.floor(totalMinutes / 60);
    minutes = totalMinutes % 60;
  }
  
  return new Date(d.getFullYear(), d.getMonth(), d.getDate(), hours, minutes, 0);
}

/**
 * Déclencheur automatique lors d'une modification manuelle du Sheets
 */
function onEdit(e?: GoogleAppsScript.Events.SheetsOnEdit): void {
  if (!e || !e.range) return;
  try {
    const sheetName = e.range.getSheet().getName();
    Logger.log("Modification détectée dans l'onglet : " + sheetName + " (Mise à jour automatique...)");
    updateFormChoices();
  } catch (err) {
    Logger.log("Erreur dans onEdit : " + err);
  }
}
