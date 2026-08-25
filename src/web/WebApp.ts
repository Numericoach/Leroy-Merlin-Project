/**
 * Point d'entrée pour l'application Web Google Apps Script
 */
function doGet(e: any): GoogleAppsScript.HTML.HtmlOutput {
  const template = HtmlService.createTemplateFromFile("CatalogTemplate");
  template.sessionsData = getAvailableSessionsForWeb();
  
  return template.evaluate()
    .setTitle("Formations Leroy Merlin")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // Permet l'intégration dans Google Sites
}

/**
 * Récupère et formate les sessions disponibles depuis l'onglet SESSIONS
 */
function getAvailableSessionsForWeb(): any[] {
  const activeSs = ss;
  const sheetSessions = activeSs ? activeSs.getSheetByName("SESSIONS") : null;
  if (!sheetSessions) return [];

  const lastRow = sheetSessions.getLastRow();
  if (lastRow < 2) return [];

  // Lecture jusqu'à la colonne 25 (Y) pour avoir toutes les nouvelles informations
  const values = sheetSessions.getRange(2, 1, lastRow - 1, 25).getValues();
  const sessions: any[] = [];

  const months = ["janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août", "septembre", "octobre", "novembre", "décembre"];
  const days = ["dimanche", "lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi"];

  values.forEach(function(row) {
    const sessionId = (row[1] || "").toString().trim(); // Col B
    const dateVal = row[3]; // Col D
    const heureDebutVal = row[4]; // Col E
    const heureFinVal = row[5]; // Col F
    const publish = row[10]; // Col K
    const placesRestantes = parseFloat(String(row[12] || "").replace(",", ".")); // Col M
    const titre = (row[14] || "").toString().trim(); // Col O
    const imageUrl = (row[17] || "").toString().trim(); // Col R
    const color = (row[18] || "2E75FF").toString().trim().replace("#", ""); // Col S
    const lienInscription = (row[19] || "").toString().trim(); // Col T
    const description = (row[20] || "").toString().trim(); // Col U
    const lieuNom = (row[23] || "").toString().trim(); // Col X

    if (!sessionId || sessionId.toLowerCase().indexOf("ses-") === -1) {
      return;
    }

    const isPublished = Boolean(publish) && 
                      String(publish).toUpperCase() !== "FALSE" && 
                      String(publish).toUpperCase() !== "FAUX" && 
                      String(publish) !== "0" && 
                      String(publish).trim() !== "";

    // Filtrer les dates passées pour ne pas afficher de sessions expirées
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

    if (isPublished && !isPastDate && !isNaN(placesRestantes) && placesRestantes > 0) {
      // Formatage de la date (ex: Le mardi 28 juillet 2026)
      let dateText = "";
      if (dateVal instanceof Date && !isNaN(dateVal.getTime())) {
        dateText = "Le " + days[dateVal.getDay()] + " " + dateVal.getDate() + " " + months[dateVal.getMonth()] + " " + dateVal.getFullYear();
      } else {
        dateText = formatDateClean(dateVal); // Fallback
      }

      // Formatage de l'heure (ex: 09:00)
      let heureDebutText = "";
      if (heureDebutVal instanceof Date && !isNaN(heureDebutVal.getTime())) {
        const h = heureDebutVal.getHours();
        const m = heureDebutVal.getMinutes();
        heureDebutText = (h < 10 ? "0" + h : h) + ":" + (m < 10 ? "0" + m : m);
      } else {
        heureDebutText = formatTimeClean(heureDebutVal).replace("h", ":");
      }

      let heureFinText = "";
      if (heureFinVal instanceof Date && !isNaN(heureFinVal.getTime())) {
        const h = heureFinVal.getHours();
        const m = heureFinVal.getMinutes();
        heureFinText = (h < 10 ? "0" + h : h) + ":" + (m < 10 ? "0" + m : m);
      } else {
        heureFinText = formatTimeClean(heureFinVal).replace("h", ":");
      }

      sessions.push({
        id: sessionId,
        titre: titre,
        date: dateText,
        horaire: "de " + heureDebutText + " à " + heureFinText,
        lieu: lieuNom,
        description: description,
        places: placesRestantes,
        imageUrl: imageUrl,
        couleur: "#" + color,
        lienInscription: lienInscription
      });
    }
  });

  return sessions;
}
