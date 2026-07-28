// Configuration globale et variables réutilisables
function getActiveSpreadsheetRobust(): GoogleAppsScript.Spreadsheet.Spreadsheet | null {
  if (typeof SpreadsheetApp === 'undefined') return null;
  return SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
}

const ss: GoogleAppsScript.Spreadsheet.Spreadsheet | null = getActiveSpreadsheetRobust();
const sheetParametres: GoogleAppsScript.Spreadsheet.Sheet | null = ss ? ss.getSheetByName("PARAMETRES") : null;
const sheetInscriptions: GoogleAppsScript.Spreadsheet.Sheet | null = ss ? ss.getSheetByName("INSCRIPTIONS") : null;

let paramCache: Map<string, string> | null = null;

/**
 * Réinitialise le cache mémoire des paramètres
 */
function clearParamCache(): void {
  paramCache = null;
}

/**
 * Récupérer la valeur d'un paramètre dans la feuille PARAMETRES (Recherche dynamique, flexible et mise en cache mémoire)
 */
function getParamValue(paramKey: string): string {
  if (paramCache && paramCache.has(paramKey)) {
    return paramCache.get(paramKey)!;
  }

  const activeSs = ss || (typeof SpreadsheetApp !== 'undefined' && SpreadsheetApp.getActiveSpreadsheet ? SpreadsheetApp.getActiveSpreadsheet() : null);
  const sheetParam = activeSs ? activeSs.getSheetByName("PARAMETRES") : null;
  if (!sheetParam) return "";

  const lastRow = sheetParam.getLastRow();
  if (lastRow < 2) return "";

  // Lire les colonnes B (Clé) et C (Valeur) de l'onglet PARAMETRES
  const data = sheetParam.getRange(2, 2, lastRow - 1, 2).getValues();

  if (!paramCache) {
    paramCache = new Map<string, string>();
  }

  const searchClean = paramKey.replace(/^PARAMETRE_/i, "").replace(/_/g, " ").trim().toLowerCase();

  for (let i = 0; i < data.length; i++) {
    const rawKey = (data[i][0] || "").toString().trim();
    if (!rawKey) continue;

    const val = data[i][1] ? data[i][1].toString().trim() : "";
    paramCache.set(rawKey, val);

    const keyClean = rawKey.replace(/^PARAMETRE_/i, "").replace(/_/g, " ").replace(/[>]/g, "").trim().toLowerCase();

    if (rawKey === paramKey || keyClean === searchClean || rawKey.toLowerCase() === paramKey.toLowerCase()) {
      paramCache.set(paramKey, val);
      return val;
    }
    
    // Fallbacks très permissifs pour les paramètres souvent mal orthographiés ou avec apostrophes/accents
    const lowerRaw = rawKey.toLowerCase();
    const noAccentRaw = lowerRaw.normalize("NFD").replace(/[\u0300-\u036f]/g, "");

    if (paramKey === "PARAMETRE_ID_FORMS_INSCRIPTION" && (noAccentRaw.indexOf("id form edit") > -1 || noAccentRaw.indexOf("forms inscription") > -1)) {
      paramCache.set(paramKey, val);
      return val;
    }
    if (paramKey === "PARAMETRE_ID_FORMS_LISTE_ATTENTE" && noAccentRaw.indexOf("liste d'attente") > -1) {
      paramCache.set(paramKey, val);
      return val;
    }
    if (paramKey === "PARAMETRE_ID_FORMS_DESINSCRIPTION" && noAccentRaw.indexOf("desinscription") > -1) {
      paramCache.set(paramKey, val);
      return val;
    }
    if (paramKey === "PARAMETRE_ENTRY_SESSION" && noAccentRaw.indexOf("entry session") > -1) {
      paramCache.set(paramKey, val);
      return val;
    }
    if (paramKey === "PARAMETRE_ENTRY_EMAIL" && noAccentRaw.indexOf("entry email") > -1) {
      paramCache.set(paramKey, val);
      return val;
    }
    if (paramKey === "PARAMETRE_ID_MODELE_CONVOC" && noAccentRaw.indexOf("modele") > -1 && noAccentRaw.indexOf("convoc") > -1) {
      paramCache.set(paramKey, val);
      return val;
    }
    if (paramKey === "PARAMETRE_ID_DOSSIER_CONVOC" && noAccentRaw.indexOf("dossier") > -1 && noAccentRaw.indexOf("convoc") > -1) {
      paramCache.set(paramKey, val);
      return val;
    }
    if (paramKey === "PARAMETRE_ID_AGENDA" && noAccentRaw.indexOf("id agenda") > -1) {
      paramCache.set(paramKey, val);
      return val;
    }
    if (paramKey === "PARAMETRE_CONNEXION_1" && (noAccentRaw.indexOf("lien connexion") > -1 || noAccentRaw.indexOf("connexion") > -1)) {
      paramCache.set(paramKey, val);
      return val;
    }
    if (paramKey === "PARAMETRE_TEXTE_AGENDA" && (noAccentRaw.indexOf("texte pour tous les evenements") > -1 || noAccentRaw.indexOf("texte pour tous") > -1)) {
      paramCache.set(paramKey, val);
      return val;
    }
    if (paramKey === "PARAMETRE_EXPEDITEUR_EMAIL" && (noAccentRaw.indexOf("mail expediteur") > -1 || noAccentRaw.indexOf("expediteur") > -1)) {
      paramCache.set(paramKey, val);
      return val;
    }
  }

  paramCache.set(paramKey, "");
  return "";
}

/**
 * Menu personnalisé dans Google Sheets
 */
function onOpen(): void {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('NUMERICOACH')
    .addItem('Synchroniser les événements Agenda (Pré-réservation)', 'createEventSession')
    .addItem('Tester la génération de PDF', 'testPDFGeneration')
    .addItem('Tester l\'intégration Google Agenda', 'testAgendaIntegration')
    .addItem('Mettre à jour les sessions dans le Formulaire', 'updateFormChoices')
    .addItem('Installer / Réinitialiser les déclencheurs (Triggers)', 'setupTriggers')
    .addItem('Retraiter les inscriptions non traitées', 'processUnprocessedInscriptions')
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
