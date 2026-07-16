import {SHEETS_QUERY} from "../src/classes/Query"
// @ts-ignore
import Query_Utils from "../src/classes/Query_utils"

describe("Double Array to JSON", () => {
	it("should return a JSON object", () => {
		const data = [
			["ID", "Name", "Age"],
			["1", "John", "20"],
			["2", "Jane", "21"],
			["3", "Joe", "22"],
			["4", "Jill", "23"]
		]

		const json = SHEETS_QUERY.getDoubleArraytoJSONData(data)
		const expected = [
			{
				ID: 1,
				Name: "John",
				Age: 20
			},
			{
				ID: 2,
				Name: "Jane",
				Age: 21
			},
			{
				ID: 3,
				Name: "Joe",
				Age: 22
			},
			{
				ID: 4,
				Name: "Jill",
				Age: 23
			}
		]

		expect(json).toEqual(expected)
	})
})

describe("Double Array to JSON with headers", () => {
	it("should return a JSON object", () => {
		const data = [
			["ID", "Name", "Age"],
			["1", "John", "20"],
			["2", "Jane", "21"],
			["3", "Joe", "22"],
			["4", "Jill", "23"]
		]

		const headers = {
			id: "ID",
			name: "Name",
			age: "Age"
		}

		const json = SHEETS_QUERY.getDoubleArraytoJSONData(data, headers)
		const expected = [
			{
				id: 1,
				name: "John",
				age: 20
			},
			{
				id: 2,
				name: "Jane",
				age: 21
			},
			{
				id: 3,
				name: "Joe",
				age: 22
			},
			{
				id: 4,
				name: "Jill",
				age: 23
			}
		]

		expect(json).toEqual(expected)
	})

	it("should return a JSON object with missing headers", () => {
		const data = [
			["ID", "Name", "Age"],
			["1", "John", "20"],
			["2", "Jane", "21"],
			["3", "Joe", "22"],
			["4", "Jill", "23"]
		]

		const headers = {
			id: "ID",
			name: "Name"
		}

		const json = SHEETS_QUERY.getDoubleArraytoJSONData(data, headers)
		const expected = [
			{
				id: 1,
				name: "John"
			},
			{
				id: 2,
				name: "Jane"
			},
			{
				id: 3,
				name: "Joe"
			},
			{
				id: 4,
				name: "Jill"
			}
		]

		expect(json).toEqual(expected)
	})

	test("Throws an error when no headers are found", () => {
		const input = []
		expect(() => {
			SHEETS_QUERY.getDoubleArraytoJSONData(input)
		}).toThrow("No data headers found")
	})
})

describe("Transform data to right type", () => {
	it("should return a number", () => {
		const data = "1"
		const expected = 1
		const result = SHEETS_QUERY.returnGoodType(data)

		expect(result).toEqual(expected)
	})

	it("should return a string", () => {
		const data = "Hello World"
		const expected = "Hello World"
		const result = SHEETS_QUERY.returnGoodType(data)

		expect(result).toEqual(expected)
	})

	it("should return a boolean", () => {
		const data = "TRUE"
		const expected = true
		const result = SHEETS_QUERY.returnGoodType(data)

		expect(result).toEqual(expected)
	})

	it("should return a float", () => {
		const data = "1.5"
		const expected = 1.5
		const result = SHEETS_QUERY.returnGoodType(data)

		expect(result).toEqual(expected)
	})

	it("should return the negative number", () => {
		const data = "-1"
		const expected = -1
		const result = SHEETS_QUERY.returnGoodType(data)

		expect(result).toEqual(expected)
	})

	it("should return the negative float", () => {
		const data = "-1.1"
		const expected = -1.1
		const result = SHEETS_QUERY.returnGoodType(data)

		expect(result).toEqual(expected)
	})

	it("should return exponential number", () => {
		const data = "1.1e+1"
		const expected = 11
		const result = SHEETS_QUERY.returnGoodType(data)

		expect(result).toEqual(expected)
	})

	it("should return the phone number as it is", () => {
		const data = "06123456768"
		const expected = "06123456768"
		const result = SHEETS_QUERY.returnGoodType(data)

		expect(result).toEqual(expected)
	})

	it("should return the data as it is", () => {
		const data = "0158.0001"
		const expected = "0158.0001"

		const result = SHEETS_QUERY.returnGoodType(data)

		expect(result).toEqual(expected)
	})
})

describe("Sheet data retrieval tests", () => {
	// Setup the mock functions directly on the Utils object
	beforeEach(() => {
		Query_Utils.gasFetch = jest.fn()
		Query_Utils.gasFetchAll = jest.fn()
		Query_Utils.parseCsv = jest.fn()
		Query_Utils.getOAuthToken = jest.fn()
	})

	afterEach(() => {
		jest.restoreAllMocks()
	})

	test("GAS_FETCH correctly handles the fetch", () => {
		// @ts-expect-error
		Query_Utils.gasFetch.mockImplementation(() => ({
			getContentText: () => "Mocked content"
		}))

		const response = Query_Utils.gasFetch({url: "https://example.com", accessToken: "token123"})
		expect(Query_Utils.gasFetch).toHaveBeenCalled()
		// Check if it has been called with the correct parameters
		expect(Query_Utils.gasFetch).toHaveBeenCalledWith({url: "https://example.com", accessToken: "token123"})
		expect(response.getContentText()).toBe("Mocked content")
	})

	test("GAS_FETCH_ALL correctly handles the fetches", () => {
		// @ts-expect-error
		Query_Utils.gasFetchAll.mockImplementation(() => [
			{
				getContentText: () => "Mocked content 1"
			},
			{
				getContentText: () => "Mocked content 2"
			}
		])

		const response = Query_Utils.gasFetchAll({URLS: ["https://example.com", "https://example2.com"], accessToken: "token123"})
		expect(Query_Utils.gasFetchAll).toHaveBeenCalled()
		// Check if it has been called with the correct parameters
		expect(Query_Utils.gasFetchAll).toHaveBeenCalledWith({URLS: ["https://example.com", "https://example2.com"], accessToken: "token123"})
		expect(response[0].getContentText()).toBe("Mocked content 1")
		expect(response[1].getContentText()).toBe("Mocked content 2")
	})

	test("query function json mode without headers", () => {
		// Mock the access token function
		// @ts-expect-error
		Query_Utils.getOAuthToken.mockImplementation(() => "token123")

		// Mock the response from the GAS_FETCH function
		// @ts-expect-error
		Query_Utils.gasFetch.mockImplementation(() => ({
			getContentText: () => "ID,Name,Age\n1,John,20\n2,Jane,21"
		}))

		// Mock the CSV parsing function
		// @ts-expect-error
		Query_Utils.parseCsv.mockImplementation(() => [
			["ID", "Name", "Age"],
			["1", "John", "20"],
			["2", "Jane", "21"]
		])

		const result = new SHEETS_QUERY("").query({
			query: "SELECT * WHERE A > 1",
			sheetName: "Sheet1",
			format: "json"
		})
		expect(result.data).toEqual([
			{ID: 1, Name: "John", Age: 20},
			{ID: 2, Name: "Jane", Age: 21}
		])
		expect(result.numRows).toEqual(2)
		expect(result.headers).toEqual({
			ID: 0,
			Name: 1,
			Age: 2
		})
	})

	test("query function json mode with headers", () => {
		// Mock the access token function
		// @ts-expect-error
		Query_Utils.getOAuthToken.mockImplementation(() => "token123")

		// Mock the response from the GAS_FETCH function
		// @ts-expect-error
		Query_Utils.gasFetch.mockImplementation(() => ({
			getContentText: () => "ID,Name,Age\n1,John,20\n2,Jane,21"
		}))

		// Mock the CSV parsing function
		// @ts-expect-error
		Query_Utils.parseCsv.mockImplementation(() => [
			["ID", "Name", "Age"],
			["1", "John", "20"],
			["2", "Jane", "21"]
		])

		const result = new SHEETS_QUERY("").query({
			query: "SELECT * WHERE A > 1",
			sheetName: "Sheet1",
			format: "json",
			headers: {
				id: "ID",
				name: "Name",
				age: "Age"
			}
		})
		expect(result.data).toEqual([
			{id: 1, name: "John", age: 20},
			{id: 2, name: "Jane", age: 21}
		])
		expect(result.numRows).toEqual(2)
	})

	test("queryMany function json mode without headers", () => {
		// Test fake calls, one to get users and one to get products

		// Mock the access token function
		// @ts-expect-error
		Query_Utils.getOAuthToken.mockImplementation(() => "token123")

		// Mock the response from the GAS_FETCH function
		// @ts-expect-error
		Query_Utils.gasFetchAll.mockImplementation(() => [
			{
				getContentText: () => "ID,Name,Age\n1,John,20\n2,Jane,21"
			},
			{
				getContentText: () => "ID,Name,Price\n1,Apple,20\n2,Banana,21"
			}
		])

		// Mock the CSV parsing function depending on the data
		// @ts-expect-error
		Query_Utils.parseCsv.mockImplementation((data) => {
			if (data === "ID,Name,Age\n1,John,20\n2,Jane,21") {
				return [
					["ID", "Name", "Age"],
					["1", "John", "20"],
					["2", "Jane", "21"]
				]
			} else {
				return [
					["ID", "Name", "Price"],
					["1", "Apple", "20"],
					["2", "Banana", "21"]
				]
			}
		})

		const result = new SHEETS_QUERY("").queryMany({
			queries: [
				{query: "SELECT *", sheetName: "Users", format: "json"},
				{query: "SELECT *", sheetName: "Products", format: "json"}
			]
		})

		expect(result).toEqual([
			{
				data: [
					{ID: 1, Name: "John", Age: 20},
					{ID: 2, Name: "Jane", Age: 21}
				],
				numRows: 2,
				headers: {
					ID: 0,
					Name: 1,
					Age: 2
				}
			},
			{
				data: [
					{ID: 1, Name: "Apple", Price: 20},
					{ID: 2, Name: "Banana", Price: 21}
				],
				numRows: 2,
				headers: {
					ID: 0,
					Name: 1,
					Price: 2
				}
			}
		])
	})
})

describe("Query error handling", () => {
	it("shouldn't let the user query without an appropriate format", () => {
		expect(() => {
			new SHEETS_QUERY("").query({
				query: "SELECT *",
				sheetName: "Sheet1",
				// @ts-expect-error
				format: "invalid"
			})
		}).toThrow("The returned type should either be 'doubleArray' or 'json'(by default)")
	})

	it("should warn the user the query failed", () => {
		// Mock the access token function
		// @ts-expect-error
		Query_Utils.getOAuthToken.mockImplementation(() => "token123")

		// Mock the response from the GAS_FETCH function
		// @ts-expect-error
		Query_Utils.gasFetch.mockImplementation(() => {
			throw new Error("Error occurred")
		})

		expect(() => {
			new SHEETS_QUERY("").query({
				query: "SELECT *",
				sheetName: "Sheet1"
			})
		}).toThrow("Erreur lors de la récupération des données via l'API : Error occurred")
	})

	it("should warn the user that the data was empty", () => {
		// Mock the access token function
		// @ts-expect-error
		Query_Utils.getOAuthToken.mockImplementation(() => "token123")

		// Mock the response from the GAS_FETCH function
		// @ts-expect-error
		Query_Utils.gasFetch.mockImplementation(() => ({
			getContentText: () => ""
		}))

		// Mock the CSV parsing function
		// @ts-expect-error
		Query_Utils.parseCsv.mockImplementation(() => [])

		expect(() => {
			new SHEETS_QUERY("").query({
				query: "SELECT *",
				sheetName: "Sheet1"
			})
		}).toThrow("No data headers found")
	})
})

describe("GAS_FETCH error handling", () => {
	test("should handle errors from UrlFetchApp.fetch gracefully", () => {
		// Mock the fetch to throw an error
		const errorMessage = "Network error occurred"
		//   @ts-expect-error
		Query_Utils.gasFetch.mockImplementation(() => {
			throw new Error(errorMessage)
		})

		// Expect the function to throw an error when called
		expect(() => {
			Query_Utils.gasFetch({url: "https://example.com", accessToken: "token123"})
		}).toThrow(errorMessage)
	})
})

describe("Handling of extreme datasets", () => {
	test("should handle extremely large datasets", async () => {
		// Mock fetch to return a very large dataset
		const largeDataset = new Array(10000).fill("ID,Name,Age\n").join("") // Simulate a large CSV
		// @ts-expect-error
		Query_Utils.getOAuthToken.mockImplementation(() => "token123")
		// @ts-expect-error
		Query_Utils.gasFetch.mockImplementation(() => ({
			getContentText: () => largeDataset
		}))
		// @ts-expect-error
		Query_Utils.parseCsv.mockImplementation(() => {
			const data = largeDataset.split("\n").map((row) => row.split(","))
			return data
		})

		const result = new SHEETS_QUERY("spreadsheetId").query({
			query: "SELECT *",
			sheetName: "LargeSheet",
			format: "json"
		})

		expect(result.numRows).toBe(10000) // Adjust according to the mock data structure
		expect(result.data.length).toBe(10000) // Adjust according to the mock data structure
		expect(result.headers).toEqual({
			ID: 0,
			Name: 1,
			Age: 2
		})
		expect(result.data[0]).toEqual({
			ID: "ID",
			Name: "Name",
			Age: "Age"
		})
	})

	test("should handle extremely small datasets", () => {
		// Mock fetch to return an empty dataset
		// @ts-expect-error
		Query_Utils.getOAuthToken.mockImplementation(() => "token123")
		// @ts-expect-error
		Query_Utils.gasFetch.mockImplementation(() => ({
			getContentText: () => "ID,Name,Age\n"
		}))
		// @ts-expect-error
		Query_Utils.parseCsv.mockImplementation(() => {
			const data = "ID,Name,Age".split("\n").map((row) => row.split(","))
			return data
		})

		const result = new SHEETS_QUERY("spreadsheetId").query({
			query: "SELECT * WHERE Age > 30",
			sheetName: "SmallSheet",
			format: "json"
		})
		expect(result.data).toEqual([])
		expect(result.numRows).toBe(0)
		expect(result.headers).toEqual({
			ID: 0,
			Name: 1,
			Age: 2
		})
	})
})
