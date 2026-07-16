export default class QueryUtils {
	static gasFetch({url, accessToken}: {url: string; accessToken: string}) {
		if (!url) throw new Error("URL parameter is required")
		if (!accessToken) throw new Error("accessToken is required")
		return UrlFetchApp.fetch(url, {
			muteHttpExceptions: false,
			headers: {
				Authorization: `Bearer ${accessToken}`
			}
		})
	}

	static gasFetchAll({URLS, accessToken}: {URLS: string[]; accessToken: string}) {
		if (!URLS) throw new Error("URLS parameter is required")
		if (!accessToken) throw new Error("accessToken is required")
		return UrlFetchApp.fetchAll(
			URLS.map((url) => {
				return {
					url,
					muteHttpExceptions: false,
					headers: {
						Authorization: `Bearer ${accessToken}`
					}
				}
			})
		)
	}

	static parseCsv(csv: string) {
		return Utilities.parseCsv(csv)
	}

	static getOAuthToken() {
		return ScriptApp.getOAuthToken()
	}
}
