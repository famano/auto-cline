import * as fs from "fs/promises"
import { after, before, describe, it } from "mocha"
import * as os from "os"
import * as path from "path"
import "should"
import { convertPptxToMd, convertMdToPptx } from "./pptx-md-converter"
import * as sinon from "sinon"

import * as pptxMdConverter from "./pptx-md-converter"

// Create a stub for nodePandoc
const nodePandocStub = sinon.stub()

// Setup the stub to simulate pandoc execution
nodePandocStub.callsFake((input: string | string[], args?: string[], callback?: (err: Error | null, result: string) => void) => {
	return new Promise<string>((resolve, reject) => {
		// Handle case when input is args array (missing first arg)
		let outputPath = ""
		if (Array.isArray(input)) {
			args = input
			input = ""
		}

		// Extract output file path from args (should be after -o)
		const outputIndex = args?.indexOf("-o") ?? -1
		if (outputIndex === -1 || !args || outputIndex >= args.length - 1) {
			return reject(new Error("Invalid arguments: missing output path"))
		}

		outputPath = args[outputIndex + 1]

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
sinon.stub(pptxMdConverter, "nodePandoc" as any).value(nodePandocStub)

describe("PPTX/Markdown Converter", () => {
	const tmpDir = path.join(os.tmpdir(), "cline-pptx-test-" + Math.random().toString(36).slice(2))
	const pptxTestFile = path.join(tmpDir, "test.pptx")
	const mdTestFile = path.join(tmpDir, "test.md")
	const refDocFile = path.join(tmpDir, "reference.pptx")

	// Setup test environment
	before(async () => {
		await fs.mkdir(tmpDir, { recursive: true })

		// Create test files
		await fs.writeFile(pptxTestFile, "mock pptx content", "utf8")
		await fs.writeFile(mdTestFile, "# Slide 1\n\nThis is a test slide.", "utf8")
		await fs.writeFile(refDocFile, "mock reference pptx content", "utf8")
	})

	// Clean up after tests
	after(async () => {
		try {
			await fs.rm(tmpDir, { recursive: true, force: true })
		} catch {
			// Ignore cleanup errors
		}
	})

	describe("convertPptxToMd", () => {
		it("should convert PPTX to Markdown with default output path", async () => {
			const outputPath = await convertPptxToMd(pptxTestFile)

			// Verify output path
			path.extname(outputPath).should.equal(".md")

			// Verify file exists
			const stat = await fs.stat(outputPath)
			stat.isFile().should.be.true()
		})

		it("should convert PPTX to Markdown with custom output path", async () => {
			const customOutputPath = path.join(tmpDir, "custom-output.md")
			const outputPath = await convertPptxToMd(pptxTestFile, customOutputPath)

			// Verify output path
			outputPath.should.equal(customOutputPath)

			// Verify file exists
			const stat = await fs.stat(outputPath)
			stat.isFile().should.be.true()
		})

		it("should throw error for non-existent input file", async () => {
			const nonExistentFile = path.join(tmpDir, "does-not-exist.pptx")

			try {
				await convertPptxToMd(nonExistentFile)
				throw new Error("Should have thrown an error")
			} catch (error) {
				error.message.should.match(/Error converting PPTX to Markdown/)
			}
		})
	})

	describe("convertMdToPptx", () => {
		it("should convert Markdown to PPTX with default output path", async () => {
			const outputPath = await convertMdToPptx(mdTestFile)

			// Verify output path
			path.extname(outputPath).should.equal(".pptx")

			// Verify file exists
			const stat = await fs.stat(outputPath)
			stat.isFile().should.be.true()
		})

		it("should convert Markdown to PPTX with custom output path", async () => {
			const customOutputPath = path.join(tmpDir, "custom-output.pptx")
			const outputPath = await convertMdToPptx(mdTestFile, customOutputPath)

			// Verify output path
			outputPath.should.equal(customOutputPath)

			// Verify file exists
			const stat = await fs.stat(outputPath)
			stat.isFile().should.be.true()
		})

		it("should convert Markdown to PPTX with reference doc", async () => {
			const outputPath = await convertMdToPptx(mdTestFile, undefined, refDocFile)

			// Verify output path
			path.extname(outputPath).should.equal(".pptx")

			// Verify file exists
			const stat = await fs.stat(outputPath)
			stat.isFile().should.be.true()
		})

		it("should throw error for non-existent input file", async () => {
			const nonExistentFile = path.join(tmpDir, "does-not-exist.md")

			try {
				await convertMdToPptx(nonExistentFile)
				throw new Error("Should have thrown an error")
			} catch (error) {
				error.message.should.match(/Error converting Markdown to PPTX/)
			}
		})

		it("should throw error for non-existent reference doc", async () => {
			const nonExistentRefDoc = path.join(tmpDir, "does-not-exist-ref.pptx")

			try {
				await convertMdToPptx(mdTestFile, undefined, nonExistentRefDoc)
				throw new Error("Should have thrown an error")
			} catch (error) {
				error.message.should.match(/Error converting Markdown to PPTX/)
			}
		})
	})
})
