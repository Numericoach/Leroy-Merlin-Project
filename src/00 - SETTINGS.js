/* NUM IEN -  INSCRIPTION EN LIGNE - code soumis aux droits d'auteurs - @2020 - 2021

tous droits de distribution et de commercialisation reservés à l'auteur propriétaire du code : NUMERICOACH 
contact@numericoach.com

*/

const ss = SpreadsheetApp.getActiveSpreadsheet(); 
const sheetParametres = ss.getSheetByName("PARAMETRES");
const sheetInscription = ss.getSheetByName("INSCRIPTIONSS"); 

const sheetInscriptions = ss.getSheetByName("INSCRIPTIONS"); 

// plages nommées 
const plagesNommeesParametres = sheetParametres.getNamedRanges(); 
let pnParaCel  = {} ; 
let celParaPn = {} ; 

for (let i=0; i<plagesNommeesParametres.length ; i++)
{
  const thisNamedRangeName = plagesNommeesParametres[i].getName(); 
  const thisNamedRangeNotation = plagesNommeesParametres[i].getRange().getA1Notation(); 
  pnParaCel[thisNamedRangeName] = thisNamedRangeNotation ; 
  celParaPn[thisNamedRangeNotation] = thisNamedRangeName ; 
  
}
const pnParaValues = Object.values(pnParaCel); 
const pnParaKeys = Object.keys(pnParaCel); 

//Logger.log(pnParaCel); 
// 
function onOpen()
{
  const ui = SpreadsheetApp.getUi(); 
  ui.createMenu("NUMERICOACH")
    .addItem("Réparer les PDF manquants", "reparePdf")
    .addToUi(); 
}

// Fonction d'aide globale pour fusionner proprement Date et Heure
function parseDateTime(dateVal, timeVal) {
  if (!dateVal) return new Date(0);
  const d = new Date(dateVal);
  let hours = 0;
  let minutes = 0;
  
  if (timeVal instanceof Date) {
    hours = timeVal.getHours();
    minutes = timeVal.getMinutes();
  } else if (typeof timeVal === 'string' || timeVal instanceof String) {
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

// Déclencheur automatique lors d'une modification manuelle du Sheets
function onEdit(e) {
  if (!e) return;
  try {
    const sheetName = e.range.getSheet().getName();
    // Vider le cache si on modifie les sessions ou les paramètres
    if (sheetName === "SESSIONS" || sheetName === "PARAMETRES") {
      clearSessionsCache();
      Logger.log("Cache vidé suite à une modification manuelle dans : " + sheetName);
    }
  } catch (err) {
    Logger.log("Erreur dans onEdit : " + err);
  }
}
