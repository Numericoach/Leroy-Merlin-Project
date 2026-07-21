function alerteSessionVide() {
  // envoyer un message 72 heures avant si pas 7 inscrits sur une CV
  /***
   * 
   * (email thierry@numericoach.fr, gaia.jublou@numericoach.fr et criva@ucpasl.com)
   * 
   * Message : ALERTE : PAS ASSEZ D'iNSCRITS SUR LA SESSION X du Y (Z inscrits).
   * 
   */

  const sheetSessions = ss.getSheetByName("SESSIONS DANS 3 JOURS");
  if (!sheetSessions) return;
  
  const valuesSessions = sheetSessions.getDataRange().getDisplayValues(); 
  if (valuesSessions.length > 1)
  {
    valuesSessions.forEach(function(session, i)
    {
      if (i > 0)
      {
        const message = "<p style='color:red'>ALERTE</p> "
        + "<p> PAS ASSEZ D'INSCRITS SUR LA SESSION " + session[1] + " du " + session[2] + " (" + session[6] + " inscrits)</p>"
        + "<p>Inscrits à ce jour : </p>"
        + session[8]; 

        MailApp.sendEmail({
          to: "thierry@numericoach.fr,gaia.jublou@numericoach.fr,criva@ucpasl.com,antoine.martin@amcms.net",
          subject: "ALERTE UCPA INSCRIPTIONS INSUFFISANTES",
          htmlBody: message
        });
      }
    });
  }
}
