import * as fs from "fs/promises"
import { after, before, describe, it } from "mocha"
import * as os from "os"
import * as path from "path"
import "should"
import * as sinon from "sinon"
import { convertDirectory, convertFile, ConversionMode } from "./directory-converter"
import * as xlsxCsvConverter from "./xlsx-csv-converter"
import * as docxMdConverter from "./docx-md-converter"
import * as pptxMdConverter from "./pptx-md-converter"

// Create stubs for the converter functions
let convertXlsxToCsvStub: sinon.SinonStub
let convertCsvToXlsxStub: sinon.SinonStub
let convertDocxToMdStub: sinon.SinonStub
let convertMdToDocxStub: sinon.SinonStub
let convertPptxToMdStub: sinon.SinonStub
let convertMdToPptxStub: sinon.SinonStub

describe("Directory Converter", () => {
	const tmpDir = path.join(os.tmpdir(), "cline-dir-test-" + Math.random().toString(36).slice(2))
	const outputDir = path.join(tmpDir, "output")

	// Setup test environment
	before(async () => {
		// Set up stubs before each test
		convertXlsxToCsvStub = sinon
			.stub(xlsxCsvConverter, "convertXlsxToCsv")
			.callsFake(async (inputPath: string, outputPath?: string) => {
				const actualOutput = outputPath || path.join(path.dirname(inputPath), `${path.basename(inputPath, ".xlsx")}.csv`)
				await fs.writeFile(actualOutput, "Mock CSV content")
				return actualOutput
			})

		convertCsvToXlsxStub = sinon
			.stub(xlsxCsvConverter, "convertCsvToXlsx")
			.callsFake(async (inputPath: string, outputPath?: string) => {
				const actualOutput = outputPath || path.join(path.dirname(inputPath), `${path.basename(inputPath, ".csv")}.xlsx`)
				await fs.writeFile(actualOutput, "Mock XLSX content")
				return actualOutput
			})

		convertDocxToMdStub = sinon
			.stub(docxMdConverter, "convertDocxToMd")
			.callsFake(async (inputPath: string, outputPath?: string) => {
				const actualOutput = outputPath || path.join(path.dirname(inputPath), `${path.basename(inputPath, ".docx")}.md`)
				await fs.writeFile(actualOutput, "Mock Markdown content")
				return actualOutput
			})

		convertMdToDocxStub = sinon
			.stub(docxMdConverter, "convertMdToDocx")
			.callsFake(async (inputPath: string, outputPath?: string) => {
				const actualOutput = outputPath || path.join(path.dirname(inputPath), `${path.basename(inputPath, ".md")}.docx`)
				await fs.writeFile(actualOutput, "Mock DOCX content")
				return actualOutput
			})

		convertPptxToMdStub = sinon
			.stub(pptxMdConverter, "convertPptxToMd")
			.callsFake(async (inputPath: string, outputPath?: string) => {
				const actualOutput = outputPath || path.join(path.dirname(inputPath), `${path.basename(inputPath, ".pptx")}.md`)
				await fs.writeFile(actualOutput, "Mock Markdown content")
				return actualOutput
			})

		convertMdToPptxStub = sinon
			.stub(pptxMdConverter, "convertMdToPptx")
			.callsFake(async (inputPath: string, outputPath?: string) => {
				const actualOutput = outputPath || path.join(path.dirname(inputPath), `${path.basename(inputPath, ".md")}.pptx`)
				await fs.writeFile(actualOutput, "Mock PPTX content")
				return actualOutput
			})

		await fs.mkdir(tmpDir, { recursive: true })
		await fs.mkdir(path.join(tmpDir, "subdir"), { recursive: true })
		await fs.mkdir(outputDir, { recursive: true })

		// Create test files
		await fs.writeFile(path.join(tmpDir, "test1.xlsx"), "mock xlsx content", "utf8")
		await fs.writeFile(path.join(tmpDir, "test2.xlsx"), "mock xlsx content", "utf8")
		await fs.writeFile(path.join(tmpDir, "subdir", "test3.xlsx"), "mock xlsx content", "utf8")

		await fs.writeFile(path.join(tmpDir, "document1.docx"), "mock docx content", "utf8")
		await fs.writeFile(path.join(tmpDir, "subdir", "document2.docx"), "mock docx content", "utf8")

		await fs.writeFile(path.join(tmpDir, "presentation1.pptx"), "mock pptx content", "utf8")

		await fs.writeFile(path.join(tmpDir, "markdown1.md"), "mock markdown content", "utf8")

		// Create non-convertible files
		await fs.writeFile(path.join(tmpDir, "text.txt"), "plain text", "utf8")
	})

	// Clean up after tests
	after(async () => {
		try {
			await fs.rm(tmpDir, { recursive: true, force: true })
		} catch {
			// Ignore cleanup errors
		}

		// Restore all stubs
		convertXlsxToCsvStub.restore()
		convertCsvToXlsxStub.restore()
		convertDocxToMdStub.restore()
		convertMdToDocxStub.restore()
		convertPptxToMdStub.restore()
		convertMdToPptxStub.restore()
	})

	describe("convertFile", () => {
		it("should convert XLSX to CSV", async () => {
			const inputPath = path.join(tmpDir, "test1.xlsx")
			const outputPath = await convertFile(inputPath, "xlsx-to-csv")

			// Verify output path
			path.extname(outputPath).should.equal(".csv")

			// Verify file exists
			const stat = await fs.stat(outputPath)
			stat.isFile().should.be.true()
		})

		it("should convert DOCX to MD", async () => {
			const inputPath = path.join(tmpDir, "document1.docx")
			const outputPath = await convertFile(inputPath, "docx-to-md")

			// Verify output path
			path.extname(outputPath).should.equal(".md")

			// Verify file exists
			const stat = await fs.stat(outputPath)
			stat.isFile().should.be.true()
		})

		it("should throw error for unsupported conversion mode", async () => {
			const inputPath = path.join(tmpDir, "test1.xlsx")

			try {
				// @ts-ignore - testing invalid conversion mode
				await convertFile(inputPath, "unsupported-mode")
				throw new Error("Should have thrown an error")
			} catch (error) {
				error.message.should.match(/Unsupported conversion mode/)
			}
		})
	})

	describe("convertDirectory", () => {
		it("should convert all XLSX files in directory to CSV", async () => {
			const result = await convertDirectory(tmpDir, "xlsx-to-csv")

			// Verify result contains success information
			result.should.match(/Converted 2 files, failed 0/)
			result.should.match(/test1.csv/)
			result.should.match(/test2.csv/)

			// Check if files exist
			const stat1 = await fs.stat(path.join(tmpDir, "test1.csv"))
			stat1.isFile().should.be.true()

			const stat2 = await fs.stat(path.join(tmpDir, "test2.csv"))
			stat2.isFile().should.be.true()
		})

		it("should convert all XLSX files recursively when recursive option is true", async () => {
			const result = await convertDirectory(tmpDir, "xlsx-to-csv", { recursive: true })

			// Verify result contains success information
			result.should.match(/Converted 3 files, failed 0/)
			result.should.match(/test1.csv/)
			result.should.match(/test2.csv/)
			result.should.match(/test3.csv/)

			// Check if files exist
			const stat3 = await fs.stat(path.join(tmpDir, "subdir", "test3.csv"))
			stat3.isFile().should.be.true()
		})

		it("should use output directory when specified", async () => {
			const result = await convertDirectory(tmpDir, "docx-to-md", { outputDir })

			// Verify result contains success information
			result.should.match(/Converted 1 files, failed 0/)

			// Check if file exists in output directory
			const stat = await fs.stat(path.join(outputDir, "document1.md"))
			stat.isFile().should.be.true()
		})

		it("should handle non-existent directory", async () => {
			const nonExistentDir = path.join(tmpDir, "does-not-exist")

			try {
				await convertDirectory(nonExistentDir, "xlsx-to-csv")
				throw new Error("Should have thrown an error")
			} catch (error) {
				error.message.should.match(/Error converting directory/)
			}
		})
	})
})
