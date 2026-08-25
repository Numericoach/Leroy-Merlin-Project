/**
 * Helper pour extraire l'ID propre d'un formulaire Google Forms même si l'utilisateur a collé l'URL entière
 */
function extractFormId(input: string): string {
  if (!input) return "";
  let cleaned = input.trim();
  if (cleaned.indexOf("/edit") > -1) {
    cleaned = cleaned.split("/edit")[0];
  }
  if (cleaned.indexOf("/viewform") > -1) {
    cleaned = cleaned.split("/viewform")[0];
  }
  const match = cleaned.match(/\/d\/(?:e\/)?([a-zA-Z0-9_-]+)/);
  if (match && match[1]) {
    return match[1];
  }
  if (cleaned.indexOf("/") > -1) {
    const parts = cleaned.split("/").filter(p => p.trim() !== "");
    return parts[parts.length - 1];
  }
  return cleaned;
}

/**
 * Helper pour formater proprement une date (string, Date, ou timestamp) en DD/MM/YYYY
 */
function formatDateClean(dateVal: any): string {
  if (!dateVal) return "";
  if (dateVal instanceof Date && !isNaN(dateVal.getTime())) {
    const d = dateVal.getDate();
    const m = dateVal.getMonth() + 1;
    const y = dateVal.getFullYear();
    return (d < 10 ? "0" : "") + d + "/" + (m < 10 ? "0" : "") + m + "/" + y;
  }
  const str = dateVal.toString().trim();
  if (str.match(/^\d{1,2}\/\d{1,2}\/\d{4}$/)) {
    return str;
  }
  const parsed = new Date(str);
  if (!isNaN(parsed.getTime())) {
    const d = parsed.getDate();
    const m = parsed.getMonth() + 1;
    const y = parsed.getFullYear();
    return (d < 10 ? "0" : "") + d + "/" + (m < 10 ? "0" : "") + m + "/" + y;
  }
  return str;
}

/**
 * Helper pour formater proprement une heure en XhXX
 */
function formatTimeClean(timeVal: any): string {
  if (!timeVal) return "";
  if (timeVal instanceof Date && !isNaN(timeVal.getTime())) {
    const h = timeVal.getHours();
    const m = timeVal.getMinutes();
    return h + "h" + (m > 0 ? (m < 10 ? "0" : "") + m : "00");
  }
  const str = timeVal.toString().trim();
  if (str.indexOf("h") > -1) return str;
  if (str.indexOf(":") > -1) {
    const parts = str.split(":");
    const h = parseInt(parts[0], 10);
    const m = parseInt(parts[1], 10);
    if (!isNaN(h)) {
      return h + "h" + (!isNaN(m) && m > 0 ? (m < 10 ? "0" : "") + m : "00");
    }
  }
  const parsed = new Date(str);
  if (!isNaN(parsed.getTime())) {
    const h = parsed.getHours();
    const m = parsed.getMinutes();
    return h + "h" + (m > 0 ? (m < 10 ? "0" : "") + m : "00");
  }
  return str;
}

/**
 * Helper pour générer l'URL d'un formulaire pré-rempli avec l'entrée réelle
 */
function getPreFilledFormUrl(formId: string, entryId: string, selectedValue: string): string {
  return "https://docs.google.com/forms/d/e/" + formId + "/viewform?usp=pp_url&" + entryId + "=" + encodeURIComponent(selectedValue);
}
