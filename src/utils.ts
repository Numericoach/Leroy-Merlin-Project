export namespace Utils {
	/**
	 * Exemple de fonction Apps Script
	 *
	 * @returns L'adresse email de l'utilisateur actif
	 */
	export function getUserEmail() {
		return Session.getActiveUser().getEmail()
	}
}
