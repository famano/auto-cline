import * as fs from "fs/promises"
import { after, before, describe, it } from "mocha"
import * as os from "os"
import * as path from "path"
import "should"
import { convertDocxToMd, convertMdToDocx } from "./docx-md-converter"
import * as sinon from "sinon"

import * as docxMdConverter from "./docx-md-converter"

// Create a stub for nodePandoc
const nodePandocStub = sinon.stub()

// Setup the stub to simulate pandoc execution
nodePandocStub.callsFake((input: string, args: string[], callback?: (err: Error | null, result: string) => void) => {
	return new Promise<string>((resolve, reject) => {
		// Extract output file path from args (should be after -o)
		const outputIndex = args.indexOf("-o")
		if (outputIndex === -1 || outputIndex >= args.length - 1) {
			return reject(new Error("Invalid arguments: missing output path"))
		}

		const outputPath = args[outputIndex + 1]

		// Create an empty file at the output path
		fs.writeFile(outputPath, "Mock conversion output")
			.then(() => {
				if (callback) {
					callback(null, "Success")
				}
				resolve("Success")
			})
			.catch((err) => {
				if (callback) {
					callback(err, "")
				}
				reject(err)
			})
	})
})

// Replace the real nodePandoc with our stub
sinon.stub(docxMdConverter, "nodePandoc" as any).value(nodePandocStub)

describe("DOCX/Markdown Converter", () => {
	const tmpDir = path.join(os.tmpdir(), "cline-docx-test-" + Math.random().toString(36).slice(2))
	const docxTestFile = path.join(tmpDir, "test.docx")
	const mdTestFile = path.join(tmpDir, "test.md")
	const refDocFile = path.join(tmpDir, "reference.docx")

	// Setup test environment
	before(async () => {
		await fs.mkdir(tmpDir, { recursive: true })

		// Create test files
		await fs.writeFile(docxTestFile, "mock docx content", "utf8")
		await fs.writeFile(mdTestFile, "# Test Markdown\n\nThis is a test.", "utf8")
		await fs.writeFile(refDocFile, "mock reference docx content", "utf8")
	})

	// Clean up after tests
	after(async () => {
		try {
			await fs.rm(tmpDir, { recursive: true, force: true })
		} catch {
			// Ignore cleanup errors
		}
	})

	describe("convertDocxToMd", () => {
		it("should convert DOCX to Markdown with default output path", async () => {
			const outputPath = await convertDocxToMd(docxTestFile)

			// Verify output path
			path.extname(outputPath).should.equal(".md")

			// Verify file exists
			const stat = await fs.stat(outputPath)
			stat.isFile().should.be.true()
		})

		it("should convert DOCX to Markdown with custom output path", async () => {
			const customOutputPath = path.join(tmpDir, "custom-output.md")
			const outputPath = await convertDocxToMd(docxTestFile, customOutputPath)

			// Verify output path
			outputPath.should.equal(customOutputPath)

			// Verify file exists
			const stat = await fs.stat(outputPath)
			stat.isFile().should.be.true()
		})

		it("should throw error for non-existent input file", async () => {
			const nonExistentFile = path.join(tmpDir, "does-not-exist.docx")

			try {
				await convertDocxToMd(nonExistentFile)
				throw new Error("Should have thrown an error")
			} catch (error) {
				error.message.should.match(/Error converting DOCX to Markdown/)
			}
		})
	})

	describe("convertMdToDocx", () => {
		it("should convert Markdown to DOCX with default output path", async () => {
			const outputPath = await convertMdToDocx(mdTestFile)

			// Verify output path
			path.extname(outputPath).should.equal(".docx")

			// Verify file exists
			const stat = await fs.stat(outputPath)
			stat.isFile().should.be.true()
		})

		it("should convert Markdown to DOCX with custom output path", async () => {
			const customOutputPath = path.join(tmpDir, "custom-output.docx")
			const outputPath = await convertMdToDocx(mdTestFile, customOutputPath)

			// Verify output path
			outputPath.should.equal(customOutputPath)

			// Verify file exists
			const stat = await fs.stat(outputPath)
			stat.isFile().should.be.true()
		})

		it("should convert Markdown to DOCX with reference doc", async () => {
			const outputPath = await convertMdToDocx(mdTestFile, undefined, refDocFile)

			// Verify output path
			path.extname(outputPath).should.equal(".docx")

			// Verify file exists
			const stat = await fs.stat(outputPath)
			stat.isFile().should.be.true()
		})

		it("should throw error for non-existent input file", async () => {
			const nonExistentFile = path.join(tmpDir, "does-not-exist.md")

			try {
				await convertMdToDocx(nonExistentFile)
				throw new Error("Should have thrown an error")
			} catch (error) {
				error.message.should.match(/Error converting Markdown to DOCX/)
			}
		})

		it("should throw error for non-existent reference doc", async () => {
			const nonExistentRefDoc = path.join(tmpDir, "does-not-exist-ref.docx")

			try {
				await convertMdToDocx(mdTestFile, undefined, nonExistentRefDoc)
				throw new Error("Should have thrown an error")
			} catch (error) {
				error.message.should.match(/Error converting Markdown to DOCX/)
			}
		})
	})
})
