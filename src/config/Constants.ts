/**
 * Constantes de configuration globale pour le projet Leroy Merlin - Tableau de Bord Formations.
 * Regroupe les valeurs magiques pour éviter la duplication et faciliter la maintenance.
 */
const Constants = {
  // Onglets (Sheets)
  SHEETS: {
    SESSIONS: "SESSIONS",
    INSCRIPTIONS: "INSCRIPTIONS",
    PARAMETRES: "PARAMETRES",
    DESINSCRIPTION_FORM: "DESINSCRIPTION FORM",
    FILE_ATTENTE: "FILE ATTENTE",
    SESSION_AGENDA: "SESSION AGENDA"
  },

  // Index de colonnes pour la feuille SESSIONS (1-based pour getRange, 0-based pour tableau JS)
  COLUMNS_SESSIONS: {
    ID: 2,          // Colonne B (ID SESSION)
    TITRE: 3,       // Colonne C
    DATE: 4,        // Colonne D
    HEURE: 5,       // Colonne E
    HEURE_FIN: 6,   // Colonne F
    DUREE: 7,       // Colonne G
    PLACES: 8,      // Colonne H (Capacité)
    INSCRITS: 12,   // Colonne L (Nb Inscrits)
    PLACES_RESTANTES: 13, // Colonne M
    EVENT_ID: 14    // Colonne N (ID EVENEMENT CALENDAR)
  },

  // Configuration LockService
  LOCK: {
    TIMEOUT_MS: 10000
  }
};
