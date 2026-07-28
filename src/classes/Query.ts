// DOC : https://ai2.metricrat.co.uk/guides/use-gviz-to-get-and-query-google-sheet-data
// If it doesn't work, try to authorize the script to access the spreadsheet by uncommenting the SpreadsheetApp.openById(ssId) line in the constructor

import QueryUtils from "./Query_utils"

// Définition d'un type pour les paramètres de chaque requête

type QueryParam<H, F extends "json" | "doubleArray"> = {
	query: string
	sheetName: string
	headers?: H
	format: F
	sheetId?: string
}

// Type générique pour transformer toutes les clés de H en un type uniforme.
type TransformToBasicTypes<H> = {[K in keyof H]: number | string | boolean}

/**
 * Type helper pour transformer une requête en un type
 *
 * @param query La requête complète avec la demande en SQL, le format, headers, etc.
 * @returns Un objet vide car on cette fonction ne sert qu'à définir le type de retour de processQuery
 */
function processQuery<H, F extends "json" | "doubleArray">(
	query: QueryParam<H, F>
): F extends "json"
	? {data: Array<TransformToBasicTypes<H>>; numRows: number; headers: Record<string, number>}
	: F extends "doubleArray"
	? {data: string[][]; numRows: number; headers: Record<string, number>}
	: never {
	// Ici irait la logique pour traiter la requête.
	// eslint-disable-next-line
	return {} as any
}

/**
 * @class SHEETS_QUERY
 * @classdesc Class to query a Google Spreadsheet with the Query Language
 * @param {string} ssId The Spreadsheet ID
 * @example
 * const spreadsheet = new SHEETS_QUERY("1X2X3X4X5X6X7X8X9X0X")
 * const data = spreadsheet.query("SELECT A, B, C WHERE A > 0", "Sheet1")
 * console.log(data)
 * // {numRows: 2, data: [{A: 1, B: 2, C: 3}, {A: 2, B: 3, C: 4}], headers: ["A", "B", "C"]}
 */
export class SHEETS_QUERY {
	ssId: string

	/**
	 * If it doesn't work, try to authorize the script to access the spreadsheet by uncommenting the SpreadsheetApp.openById(ssId) line in the constructor
	 *
	 * @param {string} ssId The Spreadsheet ID
	 */
	constructor(ssId: string) {
		this.ssId = ssId
		// To make the authorization
		// SpreadsheetApp.openById(ssId)
	}

	/**
	 * Returns the data from the sheet in JSON format with the headers or in a double array
	 * If you use something like GROUP BY, you have to remove the headers from the parameters
	 *
	 * @param {object} payload The payload
	 * @param {string} payload.query The query in the Query Language (similar to SQL)
	 * @param {string} payload.sheetName The sheet name
	 * @param {string} payload.format Type of the returned data ("json" or "doubleArray")
	 * @param {object} [payload.headers] Object with headers like so {key: Sheet Header}
	 * @param {boolean} [payload.debug] If you want to see the debug logs
	 * @returns {{numRows: number, data: object[] | any[][], headers?: object}} An object with the data, the number of rows and the headers
	 */
	// eslint-disable-next-line
	query<T extends {[key: string]: any}, H extends {[key: string]: any} = {[key: string]: any}, F extends "json" | "doubleArray" = "json">(payload: {
		query: string
		sheetName: string
		format?: F | "json" | "doubleArray"
		headers?: H
		debug?: boolean
	}) {
		const {headers} = payload
		let {query, sheetName, format, debug} = payload
		if (debug === undefined) debug = false
		// Handling format
		if (!format) format = "json" as F
		if (format == "json" || format == "doubleArray") {
			if (debug) console.log(`Returned type set to ${format}`)
		} else {
			throw new Error("The returned type should either be 'doubleArray' or 'json'(by default)")
		}

		// Fetching from API
		// Transforming the query and the sheet name to something the url can understand
		query = encodeURIComponent(query)
		sheetName = encodeURIComponent(sheetName)

		const urlEndpoint = `https://docs.google.com/a/google.com/spreadsheets/d/${this.ssId}/gviz/tq?tq=${query}&sheet=${sheetName}&tqx=out:csv&headers=1`

		let response: string
		let request: GoogleAppsScript.URL_Fetch.HTTPResponse

		try {
			request = QueryUtils.gasFetch({
				url: urlEndpoint,
				accessToken: QueryUtils.getOAuthToken()
			})

			response = request.getContentText()
		} catch (error) {
			console.log("Veuillez vérifier que l'ID du spreadsheet est correct et que vous avez accès à ce spreadsheet")
			throw `Erreur lors de la récupération des données via l'API : ${error.message}`
		}

		// Error handler
		SHEETS_QUERY.errorHandler(response)

		// Double Array
		const doubleArrayData = QueryUtils.parseCsv(response)
		if (format === "doubleArray") {
			const headers = doubleArrayData.shift()
			if (!headers) throw new Error("No data headers found")
			return {
				headers: SHEETS_QUERY.getColIdX(headers),
				numRows: doubleArrayData.length,
				data: doubleArrayData
			} as F extends "doubleArray" ? {headers: {[key: string]: number}; numRows: number; data: string[][]} : never
		}

		// Transform it to JSON
		const headersIndex = 0
		const dataHeaders = doubleArrayData[headersIndex]
		const data = SHEETS_QUERY.getDoubleArraytoJSONData<T | H>(doubleArrayData, headers) as T extends H ? T[] : {[K in keyof H]: string | number | boolean}[]
		if (!dataHeaders) throw new Error("No data headers found")

		return {
			data,
			numRows: data.length,
			headers: SHEETS_QUERY.getColIdX(dataHeaders)
		} as F extends "json" ? {headers: {[key: string]: number}; numRows: number; data: typeof data} : never
	}

	// eslint-disable-next-line
	queryMany<Qs extends QueryParam<any, any>[]>({
		queries,
		debug
	}: {
		queries: [...Qs]
		debug?: boolean
	}): {[K in keyof Qs]: ReturnType<typeof processQuery<Qs[K]["headers"], Qs[K]["format"]>>} {
		if (debug === undefined) debug = false
		// Fetching from API
		// Transforming the query and the sheet name to something the url can understand
		const URLS = queries.map(({query, sheetName, sheetId}) => {
			query = encodeURIComponent(query)
			sheetName = encodeURIComponent(sheetName)
			return `https://docs.google.com/a/google.com/spreadsheets/d/${
				sheetId ? sheetId : this.ssId
			}/gviz/tq?tq=${query}&sheet=${sheetName}&tqx=out:csv&headers=1`
		})

		let responses: string[]
		let requests: GoogleAppsScript.URL_Fetch.HTTPResponse[]

		try {
			requests = QueryUtils.gasFetchAll({
				URLS,
				accessToken: QueryUtils.getOAuthToken()
			})

			responses = requests.map((request) => request.getContentText())
			// eslint-disable-next-line
		} catch (error: any) {
			console.error("Veuillez vérifier que l'ID du spreadsheet est correct et que vous avez accès à ce spreadsheet")
			throw `Erreur lors de la récupération des données via l'API : ${error.message}`
		}

		// Error handler
		responses.forEach((response) => SHEETS_QUERY.errorHandler(response))

		// Double Array
		const doubleArrayData = responses.map((response) => QueryUtils.parseCsv(response))

		return doubleArrayData.map((csvData, index) => {
			const {format, headers} = queries[index]
			if (format === "doubleArray") {
				const headers = csvData.shift()
				if (!headers) throw new Error("No headers found")
				return {
					headers: SHEETS_QUERY.getColIdX(headers),
					numRows: csvData.length,
					data: csvData
				}
			} else {
				const headersIndex = 0
				const dataHeaders = csvData[headersIndex]
				const data = SHEETS_QUERY.getDoubleArraytoJSONData(csvData, headers)
				if (!dataHeaders) throw new Error("No headers found")
				return {
					data,
					numRows: data.length,
					headers: SHEETS_QUERY.getColIdX(dataHeaders)
				}
			}
			// eslint-disable-next-line
		}) as any
	}

	/**
	 * Returns the data from the sheet either in JSON format or in a double array
	 * It uses the range to get the data
	 * If you use something like GROUP BY, you have to remove the headers from the parameters
	 *
	 * @param {object} payload The payload
	 * @param {string} payload.query The query in the Query Language (similar to SQL)
	 * @param {string} payload.sheetName The sheet name
	 * @param {string} payload.range The range of the data (ex: A1:B10)
	 * @param {string} [payload.format] Type of the returned data ("json" or "doubleArray")
	 * @param {object} [payload.headers] Object with headers like so {key: Sheet Header}
	 * @param {boolean} [payload.debug] If you want to see the debug logs
	 * @returns {{numRows: number, data: object[] | any[][], headers?: string[]}} An object with the data, the number of rows and the headers
	 */
	// eslint-disable-next-line
	queryWithRange<T extends {[key: string]: any}, H extends {[key: string]: any} = {[key: string]: any}, F extends "json" | "doubleArray" = "json">(payload: {
		query: string
		sheetName: string
		range: string
		format?: F | "json" | "doubleArray"
		headers?: H
		debug?: boolean
	}) {
		const {headers} = payload
		let {query, sheetName, range, format, debug} = payload
		if (debug === undefined) debug = false
		// Handling format
		if (!format) format = "json" as F
		if (format == "json" || format == "doubleArray") {
			if (debug) console.log(`Returned type set to ${format}`)
		} else {
			throw new Error("The returned type should either be 'doubleArray' or 'json'(by default)")
		}

		// Check range to be a Google Sheets range
		// It can has multiple letters like so : A1:AA10
		// It can also have no number
		// It can also have no letter
		// It can also have only one letter
		if (!range.match(/^[A-Z]+[0-9]*:[A-Z]+[0-9]*$/)) throw new Error("The range should be a Google Sheets range like so : A1:AA10")

		// Fetching from API
		// Transforming the query and the sheet name to something the url can understand
		query = encodeURIComponent(query)
		sheetName = encodeURIComponent(sheetName)
		range = encodeURIComponent(range)

		const urlEndpoint = `https://docs.google.com/a/google.com/spreadsheets/d/${this.ssId}/gviz/tq?tq=${query}&sheet=${sheetName}&range=${range}&tqx=out:csv&headers=1`

		let response: string

		try {
			response = QueryUtils.gasFetch({
				url: urlEndpoint,
				accessToken: QueryUtils.getOAuthToken()
			}).getContentText()
		} catch (error) {
			console.error("Veuillez vérifier que l'ID du spreadsheet est correct et que vous avez accès à ce spreadsheet")
			throw `Erreur lors de la récupération des données via l'API : ${error.message}`
		}

		// Error handler
		SHEETS_QUERY.errorHandler(response)

		// Double Array
		const doubleArrayData = QueryUtils.parseCsv(response)
		if (format === "doubleArray") {
			const dataHeaders = doubleArrayData.shift()
			if (!dataHeaders) throw new Error("No headers in data found")
			return {
				headers: SHEETS_QUERY.getColIdX(dataHeaders),
				numRows: doubleArrayData.length,
				data: doubleArrayData
			} as F extends "doubleArray" ? {headers: {[key: string]: number}; numRows: number; data: string[][]} : never
		}

		// Transform it to JSON
		const headersIndex = 0
		const dataHeaders = doubleArrayData[headersIndex]
		const data = SHEETS_QUERY.getDoubleArraytoJSONData<T | H>(doubleArrayData, headers) as T extends H ? T[] : {[K in keyof H]: string | number | boolean}[]
		if (!dataHeaders) throw new Error("No headers found")

		return {
			data,
			numRows: data.length,
			headers: SHEETS_QUERY.getColIdX(dataHeaders)
		} as F extends "json" ? {headers: {[key: string]: number}; numRows: number; data: typeof data} : never
	}

	/**
	 * Returns the data from the sheet in HTML Table format
	 *
	 * @param {object} payload The payload
	 * @param {string} payload.query The query in the Query Language (similar to SQL)
	 * @param {string} payload.sheetName The sheet name
	 * @param {string} [payload.range] The range of the data (ex: A1:B10)
	 * @param {boolean} [payload.debug] If you want to see the debug logs
	 * @returns {string} The HTML Table
	 */
	queryToHTML(payload: {query: string; sheetName: string; range?: string; debug?: boolean}): string {
		let {query, sheetName, range, debug} = payload
		if (debug === undefined) debug = false
		if (debug) console.log(`Returned type set to html`)

		// Check range to be a Google Sheets range
		// It can has multiple letters like so : A1:AA10
		// It can also have no number
		// It can also have no letter
		// It can also have only one letter
		if (range && !range.match(/^[A-Z]+[0-9]*:[A-Z]+[0-9]*$/)) throw new Error("The range should be a Google Sheets range like so : A1:AA10")

		// Fetching from API
		// Transforming the query and the sheet name to something the url can understand
		query = encodeURIComponent(query)
		sheetName = encodeURIComponent(sheetName)
		if (range) range = encodeURIComponent(range)

		let urlEndpoint = `https://docs.google.com/a/google.com/spreadsheets/d/${this.ssId}/gviz/tq?tq=${query}&sheet=${sheetName}&tqx=out:html`
		if (range) urlEndpoint += `&range=${range}`

		let response: string

		try {
			response = QueryUtils.gasFetch({
				url: urlEndpoint,
				accessToken: QueryUtils.getOAuthToken()
			}).getContentText()
		} catch (error) {
			console.error("Veuillez vérifier que l'ID du spreadsheet est correct et que vous avez accès à ce spreadsheet")
			throw `Erreur lors de la récupération des données via l'API : ${error.message}`
		}

		// Error handler
		SHEETS_QUERY.errorHandler(response)

		return response
	}

	/**
	 * Returns the data from the sheet either in JSON format or in a double array using a skip and a limit
	 *
	 * @param {object} payload The payload
	 * @param {string} payload.query The query in the Query Language (similar to SQL)
	 * @param {string} payload.sheetName The sheet name
	 * @param {number} payload.skip The number of rows to skip
	 * @param {number} payload.limit The number of rows to return
	 * @param {string} [payload.range] The range of the data (ex: A1:B10)
	 * @param {string} [payload.format] Type of the returned data ("json" or "doubleArray")
	 * @param {object} [payload.headers] Object with headers like so {key: Sheet Header}
	 * @param {boolean} [payload.debug] If you want to see the debug logs
	 * @returns {{numRows: number, data: object[] | any[][], headers?: string[]}} An object with the data, the number of rows and the headers
	 */
	// eslint-disable-next-line
	queryWithSkip<T extends {[key: string]: any}, H extends {[key: string]: any} = {[key: string]: any}, F extends "json" | "doubleArray" = "json">(payload: {
		query: string
		sheetName: string
		skip: number
		limit: number
		range?: string
		format?: F | "json" | "doubleArray"
		headers?: H
		debug?: boolean
	}) {
		const {headers, limit, skip} = payload
		let {query, sheetName, range, format, debug} = payload
		if (debug === undefined) debug = false

		// Handling format
		if (!format) format = "json" as F
		if (format == "json" || format == "doubleArray") {
			if (debug) console.log(`Returned type set to ${format}`)
		} else {
			throw new Error("The returned type should either be 'doubleArray' or 'json'(by default)")
		}
		// Check range to be a Google Sheets range
		// It can has multiple letters like so : A1:AA10
		// It can also have no number
		// It can also have no letter
		// It can also have only one letter
		if (range && !range.match(/^[A-Z]+[0-9]*:[A-Z]+[0-9]*$/)) throw new Error("The range should be a Google Sheets range like so : A1:AA10")

		// To add the limit and the skip to the query
		// For limiting its at the end of the query like so : LIMIT 10
		// For skipping its after the limit like so : LIMIT 10 OFFSET 10
		query += ` LIMIT ${limit} OFFSET ${skip}`
		if (debug) console.log(`Query with skip and limit : ${query}`)

		// Fetching from API
		// Transforming the query and the sheet name to something the url can understand
		query = encodeURIComponent(query)
		sheetName = encodeURIComponent(sheetName)
		if (range) range = encodeURIComponent(range)
		if (range) range = `&range=${range}`
		const urlEndpoint = `https://docs.google.com/a/google.com/spreadsheets/d/${this.ssId}/gviz/tq?tq=${query}&sheet=${sheetName}${range}&tqx=out:csv&headers=1`

		let response: string
		try {
			response = QueryUtils.gasFetch({
				url: urlEndpoint,
				accessToken: QueryUtils.getOAuthToken()
			}).getContentText()
		} catch (error) {
			console.error("Veuillez vérifier que l'ID du spreadsheet est correct et que vous avez accès à ce spreadsheet")
			throw `Erreur lors de la récupération des données via l'API : ${error.message}`
		}

		// Error handler
		SHEETS_QUERY.errorHandler(response)

		// Double Array
		const doubleArrayData = QueryUtils.parseCsv(response)
		if (format === "doubleArray") {
			const headers = doubleArrayData.shift()
			if (!headers) throw new Error("No headers found")
			return {
				headers: SHEETS_QUERY.getColIdX(headers),
				numRows: doubleArrayData.length,
				data: doubleArrayData
			} as F extends "doubleArray" ? {headers: {[key: string]: number}; numRows: number; data: string[][]} : never
		}

		// Transform it to JSON
		const headersIndex = 0
		const dataHeaders = doubleArrayData[headersIndex]
		const data = SHEETS_QUERY.getDoubleArraytoJSONData<T | H>(doubleArrayData, headers) as T extends H ? T[] : {[K in keyof H]: string | number | boolean}[]
		if (!headers) throw new Error("No headers found")
		return {
			headers: SHEETS_QUERY.getColIdX(dataHeaders),
			numRows: data.length,
			data
		} as F extends "json" ? {headers: {[key: string]: number}; numRows: number; data: typeof data} : never
	}

	/**
	 * Expressions régulières pré-compilées pour de meilleures performances dans returnGoodType
	 */
	private static readonly REGEX_STARTS_ZERO = /^0/
	private static readonly REGEX_FRENCH_PHONE = /^(0\d{9})$|^(?:\+33|0)\d*(?:\d{2}){4}$|^\+\d{1,3}\s*(\d+\s*)+$/
	private static readonly REGEX_LEADING_ZERO_NUM = /^0\d+$/
	private static readonly REGEX_INTEGER = /^-?\d+$/
	private static readonly REGEX_FLOAT = /^-?\d+[\.,]\d+$/
	private static readonly REGEX_EXPONENT = /^(\d+(\.\d+)?)(e[+-]\d+)$/

	/**
	 * Transforms the double array to JSON
	 *
	 * @param {any[][]} allData The data we get in the double array format such as Range.getValues()
	 * @param {object} headers Object with headers like so {key: Sheet Header}
	 * @returns {object[]} The JSON Format Data
	 */
	// eslint-disable-next-line
	static getDoubleArraytoJSONData<H extends {[key: string]: any}>(allData: any[][], headers?: H | {[key: string]: string}) {
		let headersRow = allData.shift()
		let headersFromEntries:
			| {
					// eslint-disable-next-line
					[k: string]: any
					// eslint-disable-next-line
			  }
			| undefined
		headersRow = headersRow?.map((header: string) => header.trim())
		if (!headersRow) throw new Error("No data headers found")
		const colIdx = this.getColIdX(headersRow)
		let headersKeys: string[] | undefined
		if (headers) {
			// Filter headers to only have the headers that are in the sheet
			headersFromEntries = Object.fromEntries(Object.entries(headers).filter(([, value]) => headersRow?.includes(value.trim())))
			headersKeys = Object.keys(headersFromEntries)
			headersRow = Object.values(headersFromEntries)
		}

		if (!headersRow) throw new Error("No data headers found")

		// Pré-calcul des correspondances colonnes -> clés pour éviter l'itération superflue dans les boucles
		const mapEntries: Array<{key: string; colIndex: number}> = []
		if (headersFromEntries && headersKeys && headersRow) {
			for (let i = 0; i < headersKeys.length; i++) {
				const key = headersKeys[i]
				const headerVal = headersRow[i]
				const idx = colIdx[headerVal]
				if (idx !== undefined) {
					mapEntries.push({key, colIndex: idx})
				}
			}
		} else {
			for (let i = 0; i < headersRow.length; i++) {
				mapEntries.push({key: headersRow[i], colIndex: i})
			}
		}

		// Transformation directe sans reduce inutile
		const data: {[K in keyof H]: string | number | boolean}[] = allData.map((row) => {
			const rowObj: any = {}
			for (let i = 0; i < mapEntries.length; i++) {
				const entry = mapEntries[i]
				rowObj[entry.key] = this.returnGoodType(row[entry.colIndex])
			}
			return rowObj
		})

		return data
	}

	/**
	 * Returns an object with the headers as keys and the index as value
	 *
	 * @param {string[]} headers The first row of the sheet
	 * @returns {object} Object with the headers as keys and the index as value
	 */
	static getColIdX<T extends string>(headers: T[]): {[K in T[number]]: number} {
		return headers.reduce((object, value, index) => {
			object[value] = index
			return object
		}, {} as {[K in T[number]]: number})
	}

	/**
	 * Returns the value with the correct type
	 *
	 * @param {string} value Value to replace
	 * @returns {number | string | boolean} The value with the correct type
	 */
	static returnGoodType(value: string): number | string | boolean {
		if (!value || typeof value !== "string") return value
		if (value === "TRUE" || value === "true") return true
		if (value === "FALSE" || value === "false") return false

		if (this.REGEX_STARTS_ZERO.test(value)) {
			if (this.REGEX_FRENCH_PHONE.test(value) && !this.REGEX_LEADING_ZERO_NUM.test(value)) {
				return value
			}
			return value
		}

		if (this.REGEX_INTEGER.test(value) && !value.includes(".") && !value.includes(",")) {
			return parseInt(value, 10)
		}
		if (this.REGEX_FLOAT.test(value)) {
			return parseFloat(value.replace(",", "."))
		}
		if (this.REGEX_EXPONENT.test(value) && !this.REGEX_LEADING_ZERO_NUM.test(value)) {
			return parseFloat(value)
		}

		return value
	}

	/**
	 * Error handler for the API
	 *
	 * @param {string} response The response from the API
	 * @returns {void}
	 */
	static errorHandler(response: string): void {
		let errorMessage = {
			status: "ok",
			errors: []
		}
		try {
			errorMessage = JSON.parse(response)
		} catch (error) {
			// console.log("No error message")
		}
		// Error handler
		if (errorMessage && errorMessage.status == "error") {
			// console.error("Error Response : ", errorMessage)
			throw new Error(errorMessage.errors.map((error: {detailed_message: string}) => error.detailed_message).join(", "))
		}
	}
}
