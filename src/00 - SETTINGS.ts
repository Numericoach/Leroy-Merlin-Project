/* NUMERICOACH - INSCRIPTIONS LEROY MERLIN - @2026 */

const ss: GoogleAppsScript.Spreadsheet.Spreadsheet = SpreadsheetApp.getActiveSpreadsheet(); 
const sheetParametres: GoogleAppsScript.Spreadsheet.Sheet | null = ss ? ss.getSheetByName("PARAMETRES") : null;
const sheetInscriptions: GoogleAppsScript.Spreadsheet.Sheet | null = ss ? ss.getSheetByName("INSCRIPTIONS") : null; 

// Plages nommées des paramètres
const plagesNommeesParametres: GoogleAppsScript.Spreadsheet.NamedRange[] = sheetParametres ? sheetParametres.getNamedRanges() : []; 
const pnParaCel: Record<string, string> = {}; 
const celParaPn: Record<string, string> = {}; 

if (plagesNommeesParametres && plagesNommeesParametres.length > 0) {
  for (let i = 0; i < plagesNommeesParametres.length; i++) {
    const thisNamedRangeName = plagesNommeesParametres[i].getName(); 
    const thisNamedRangeNotation = plagesNommeesParametres[i].getRange().getA1Notation(); 
    pnParaCel[thisNamedRangeName] = thisNamedRangeNotation; 
    celParaPn[thisNamedRangeNotation] = thisNamedRangeName; 
  }
}

/**
 * Helper sécurisé pour récupérer la valeur d'un paramètre sans jamais faire crasher le script
 */
function getParamValue(paramKey: string): string {
  try {
    if (!sheetParametres || !pnParaCel) return "";
    const cellA1 = pnParaCel[paramKey];
    if (!cellA1) return "";
    const val = sheetParametres.getRange(cellA1).getValue();
    return val !== null && val !== undefined ? val.toString().trim() : "";
  } catch (err) {
    Logger.log("Avertissement getParamValue(" + paramKey + ") : " + err);
    return "";
  }
}

/**
 * Trigger à l'ouverture de la feuille de calcul
 */
function onOpen(): void {
  try {
    const ui = SpreadsheetApp.getUi(); 
    ui.createMenu("NUMERICOACH")
      .addItem("Installer les déclencheurs (Triggers)", "setupTriggers")
      .addItem("Mettre à jour les sessions dans le Formulaire", "updateFormChoices")
      .addToUi(); 
  } catch (e) {
    Logger.log("onOpen non interactif ou environnement batch.");
  }
}

/**
 * Fonction d'aide globale pour fusionner proprement Date et Heure
 */
function parseDateTime(dateVal: Date | string | number | null | undefined, timeVal: Date | string | number | null | undefined): Date {
  if (!dateVal) return new Date(0);
  const d = new Date(dateVal);
  let hours = 0;
  let minutes = 0;
  
  if (timeVal instanceof Date) {
    hours = timeVal.getHours();
    minutes = timeVal.getMinutes();
  } else if (typeof timeVal === 'string') {
    const parts = timeVal.split(':');
    if (parts.length >= 2) {
      hours = parseInt(parts[0], 10);
      minutes = parseInt(parts[1], 10);
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
    if (sheetName === "SESSIONS" || sheetName === "PARAMETRES") {
      Logger.log("Modification détectée dans : " + sheetName);
      updateFormChoices();
    }
  } catch (err) {
    Logger.log("Erreur dans onEdit : " + err);
  }
}
