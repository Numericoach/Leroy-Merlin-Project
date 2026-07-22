/* NUMERICOACH - INSCRIPTIONS LEROY MERLIN - @2026 */

var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheetParametres = ss ? ss.getSheetByName("PARAMETRES") : null;
var sheetInscriptions = ss ? ss.getSheetByName("INSCRIPTIONS") : null;

var plagesNommeesParametres = sheetParametres ? sheetParametres.getNamedRanges() : [];
var pnParaCel = {};
var celParaPn = {};

if (plagesNommeesParametres && plagesNommeesParametres.length > 0) {
  for (var i = 0; i < plagesNommeesParametres.length; i++) {
    var thisNamedRangeName = plagesNommeesParametres[i].getName();
    var thisNamedRangeNotation = plagesNommeesParametres[i].getRange().getA1Notation();
    pnParaCel[thisNamedRangeName] = thisNamedRangeNotation;
    celParaPn[thisNamedRangeNotation] = thisNamedRangeName;
  }
}

function getParamValue(paramKey) {
  try {
    if (!sheetParametres || !pnParaCel) return "";
    var cellA1 = pnParaCel[paramKey];
    if (!cellA1) return "";
    var val = sheetParametres.getRange(cellA1).getValue();
    return val !== null && val !== undefined ? val.toString().trim() : "";
  } catch (err) {
    Logger.log("Avertissement getParamValue(" + paramKey + ") : " + err);
    return "";
  }
}

function onOpen() {
  try {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu("NUMERICOACH")
      .addItem("Installer les déclencheurs (Triggers)", "setupTriggers")
      .addItem("Mettre à jour les sessions dans le Formulaire", "updateFormChoices")
      .addToUi();
  } catch (e) {
    Logger.log("onOpen non interactif ou environnement batch.");
  }
}

function parseDateTime(dateVal, timeVal) {
  if (!dateVal) return new Date(0);
  var d = new Date(dateVal);
  var hours = 0;
  var minutes = 0;
  
  if (timeVal instanceof Date) {
    hours = timeVal.getHours();
    minutes = timeVal.getMinutes();
  } else if (typeof timeVal === 'string') {
    var parts = timeVal.split(':');
    if (parts.length >= 2) {
      hours = parseInt(parts[0], 10);
      minutes = parseInt(parts[1], 10);
    }
  } else if (typeof timeVal === 'number') {
    var totalMinutes = Math.round(timeVal * 24 * 60);
    hours = Math.floor(totalMinutes / 60);
    minutes = totalMinutes % 60;
  }
  
  return new Date(d.getFullYear(), d.getMonth(), d.getDate(), hours, minutes, 0);
}

function onEdit(e) {
  if (!e || !e.range) return;
  try {
    var sheetName = e.range.getSheet().getName();
    if (sheetName === "SESSIONS" || sheetName === "PARAMETRES") {
      Logger.log("Modification détectée dans : " + sheetName);
      updateFormChoices();
    }
  } catch (err) {
    Logger.log("Erreur dans onEdit : " + err);
  }
}
