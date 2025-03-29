import * as fs from "fs/promises"
import { after, before, describe, it } from "mocha"
import * as os from "os"
import * as path from "path"
import "should"
import { convertXlsxToCsv, convertCsvToXlsx } from "./xlsx-csv-converter"
import xlsx from "node-xlsx"

describe("Excel/CSV Converter", () => {
	const tmpDir = path.join(os.tmpdir(), "cline-xlsx-test-" + Math.random().toString(36).slice(2))
	const xlsxTestFile = path.join(tmpDir, "test.xlsx")
	const csvTestFile = path.join(tmpDir, "test.csv")

	// Setup test environment
	before(async () => {
		await fs.mkdir(tmpDir, { recursive: true })

		// Create test XLSX file
		const data = [
			["Name", "Age", "City"],
			["John", 30, "New York"],
			["Alice", 25, "London"],
		]
		const buffer = xlsx.build([{ name: "sheet1", data, options: {} }])
		await fs.writeFile(xlsxTestFile, buffer)

		// Create test CSV file
		const csvContent = "Name,Age,City\nJohn,30,New York\nAlice,25,London"
		await fs.writeFile(csvTestFile, csvContent, "utf8")
	})

	// Clean up after tests
	after(async () => {
		try {
			await fs.rm(tmpDir, { recursive: true, force: true })
		} catch {
			// Ignore cleanup errors
		}
	})

	describe("convertXlsxToCsv", () => {
		it("should convert XLSX to CSV with default output path", async () => {
			const outputPath = await convertXlsxToCsv(xlsxTestFile)

			// Verify output path
			path.extname(outputPath).should.equal(".csv")

			// Verify file exists
			const stat = await fs.stat(outputPath)
			stat.isFile().should.be.true()

			// Verify content
			const content = await fs.readFile(outputPath, "utf8")
			content.should.match(/Name,Age,City/)
			content.should.match(/John,30,New York/)
		})

		it("should convert XLSX to CSV with custom output path", async () => {
			const customOutputPath = path.join(tmpDir, "custom-output.csv")
			const outputPath = await convertXlsxToCsv(xlsxTestFile, customOutputPath)

			// Verify output path
			outputPath.should.equal(customOutputPath)

			// Verify file exists
			const stat = await fs.stat(outputPath)
			stat.isFile().should.be.true()
		})

		it("should throw error for non-existent input file", async () => {
			const nonExistentFile = path.join(tmpDir, "does-not-exist.xlsx")

			try {
				await convertXlsxToCsv(nonExistentFile)
				throw new Error("Should have thrown an error")
			} catch (error) {
				error.message.should.match(/Error converting XLSX to CSV/)
			}
		})
	})

	describe("convertCsvToXlsx", () => {
		it("should convert CSV to XLSX with default output path", async () => {
			const outputPath = await convertCsvToXlsx(csvTestFile)

			// Verify output path
			path.extname(outputPath).should.equal(".xlsx")

			// Verify file exists
			const stat = await fs.stat(outputPath)
			stat.isFile().should.be.true()

			// Verify content by reading back the XLSX
			const workbook = xlsx.parse(outputPath)
			workbook.length.should.be.greaterThan(0)
			workbook[0].data.length.should.be.greaterThan(0)
			workbook[0].data[0][0].should.equal("Name")
		})

		it("should convert CSV to XLSX with custom output path", async () => {
			const customOutputPath = path.join(tmpDir, "custom-output.xlsx")
			const outputPath = await convertCsvToXlsx(csvTestFile, customOutputPath)

			// Verify output path
			outputPath.should.equal(customOutputPath)

			// Verify file exists
			const stat = await fs.stat(outputPath)
			stat.isFile().should.be.true()
		})

		it("should throw error for non-existent input file", async () => {
			const nonExistentFile = path.join(tmpDir, "does-not-exist.csv")

			try {
				await convertCsvToXlsx(nonExistentFile)
				throw new Error("Should have thrown an error")
			} catch (error) {
				error.message.should.match(/Error converting CSV to XLSX/)
			}
		})
	})
})
