function doGet(e) {
  const template = HtmlService.createTemplateFromFile('web');
  template.sessions = getSessionsJsonData();
  
  // Récupérer l'ID du formulaire
  let formId = "";
  try {
    formId = sheetParametres.getRange(pnParaCel['PARAMETRE_ID_FORMS_INSCRIPTION']).getValue();
  } catch (err) {
    Logger.log("Erreur de récupération du formulaire ID : " + err);
  }
  template.formId = formId;

  return template.evaluate()
    .setTitle("Sessions de Formation")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSessionsJsonData() {
  const cache = CacheService.getScriptCache();
  const cachedData = cache.get("sessions_data");
  
  if (cachedData != null) {
    try {
      const parsed = JSON.parse(cachedData);
      // Auto-invalidation si le cache est ancien et ne contient pas timeRange
      if (parsed.length > 0 && parsed[0].timeRange !== undefined) {
        Logger.log("Données récupérées depuis le cache valide.");
        return parsed;
      }
    } catch (e) {
      Logger.log("Erreur lors de la lecture du cache, rechargement.");
    }
  }
  
  Logger.log("Cache vide ou obsolète. Récupération des données depuis le Sheets.");
  const sheetSessions = ss.getSheetByName("SESSIONS");
  if (!sheetSessions) return [];
  const lastRow = sheetSessions.getLastRow();
  if (lastRow < 2) return [];
  
  const values = sheetSessions.getRange(2, 2, lastRow - 1, 14).getValues();
  const sessions = [];
  const now = new Date();
  
  // Helper pour formater proprement l'heure en HH:MM
  const formatTime = function(timeVal) {
    if (timeVal instanceof Date) {
      const h = timeVal.getHours().toString().padStart(2, '0');
      const m = timeVal.getMinutes().toString().padStart(2, '0');
      return h + ":" + m;
    }
    if (typeof timeVal === 'string' || timeVal instanceof String) {
      const parts = timeVal.split(':');
      if (parts.length >= 2) return parts[0].padStart(2, '0') + ":" + parts[1].padStart(2, '0');
    }
    if (typeof timeVal === 'number') {
      const totalMinutes = Math.round(timeVal * 24 * 60);
      const h = Math.floor(totalMinutes / 60).toString().padStart(2, '0');
      const m = (totalMinutes % 60).toString().padStart(2, '0');
      return h + ":" + m;
    }
    return "";
  };

  values.forEach(function(row) {
    const sessionId = row[0];
    const publish = row[9];
    const remaining = row[11];
    const formationTitle = row[13];
    const dateRaw = row[2];
    const location = row[7];
    const duration = row[5];
    const details = row[8];
    
    const startTimeRaw = row[3]; // Col E: HEURE DEBUT
    const endTimeRaw = row[4];   // Col F: HEURE FIN
    
    const sessionEnd = parseDateTime(dateRaw, endTimeRaw);

    if (sessionId && publish && Number(remaining) > 0 && sessionEnd > now) {
      const dateObj = new Date(dateRaw);
      const dateStr = dateObj.getDate() + "/" + (dateObj.getMonth() + 1) + "/" + dateObj.getFullYear();
      
      const startTimeStr = formatTime(startTimeRaw);
      const endTimeStr = formatTime(endTimeRaw);
      const timeRangeStr = startTimeStr && endTimeStr ? startTimeStr + " à " + endTimeStr : (startTimeStr || "Non spécifiée");

      sessions.push({
        id: sessionId,
        title: formationTitle,
        date: dateStr,
        timeRange: timeRangeStr,
        location: location || "En ligne (Google Meet)",
        duration: duration || "Non spécifiée",
        details: details || "",
        remaining: remaining
      });
    }
  });
  
  // Mettre en cache pour 10 minutes (600 secondes)
  try {
    cache.put("sessions_data", JSON.stringify(sessions), 600);
  } catch (e) {
    Logger.log("Impossible d'écrire dans le cache: " + e);
  }
  
  return sessions;
}

// Fonction pour effacer le cache des sessions
function clearSessionsCache() {
  const cache = CacheService.getScriptCache();
  cache.remove("sessions_data");
  Logger.log("Cache des sessions vidé.");
}
