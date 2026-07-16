import Utils from "../src/classes/Utils"

describe("always true", () => {
	it("should be true", () => {
		const first = 1
		const second = 2
		const result = new Utils().addition(first, second)
		const expected = 3
		expect(result).toBe(expected)
	})
})
