import * as fs from "fs/promises"
import * as path from "path"

const nodePandoc = require("node-pandoc")

/**
 * Converts a DOCX file to Markdown format using pandoc
 * @param inputPath Path to the DOCX file
 * @param outputPath Path where the Markdown file will be saved (optional, defaults to same directory with .md extension)
 * @param referenceDocPath Optional path to a reference doc for styling
 * @returns Path to the created Markdown file
 */
export async function convertDocxToMd(inputPath: string, outputPath?: string): Promise<string> {
	try {
		// Validate input file exists
		await fs.access(inputPath)

		// If output path is not provided, use the same directory with .md extension
		if (!outputPath) {
			const parsedPath = path.parse(inputPath)
			outputPath = path.join(parsedPath.dir, `${parsedPath.name}.md`)
		}

		// Build pandoc arguments
		const args = ["-f", "docx", "-t", "markdown", "-o", outputPath]

		// Execute pandoc command
		await nodePandoc(inputPath, args)

		return outputPath
	} catch (error) {
		throw new Error(`Error converting DOCX to Markdown: ${error instanceof Error ? error.message : String(error)}`)
	}
}

/**
 * Converts a Markdown file to DOCX format using pandoc
 * @param inputPath Path to the Markdown file
 * @param outputPath Path where the DOCX file will be saved (optional, defaults to same directory with .docx extension)
 * @param referenceDocPath Optional path to a reference doc for styling
 * @returns Path to the created DOCX file
 */
export async function convertMdToDocx(inputPath: string, outputPath?: string, referenceDocPath?: string): Promise<string> {
	try {
		// Validate input file exists
		await fs.access(inputPath)

		// If output path is not provided, use the same directory with .docx extension
		if (!outputPath) {
			const parsedPath = path.parse(inputPath)
			outputPath = path.join(parsedPath.dir, `${parsedPath.name}.docx`)
		}

		// Build pandoc arguments
		const args = ["-f", "markdown", "-t", "docx", "-o", outputPath]

		// Add reference-doc if provided
		if (referenceDocPath) {
			await fs.access(referenceDocPath) // Validate reference doc exists
			args.push(`--reference-doc=${referenceDocPath}`)
		}

		// Execute pandoc command
		await nodePandoc(inputPath, args)

		return outputPath
	} catch (error) {
		throw new Error(`Error converting Markdown to DOCX: ${error instanceof Error ? error.message : String(error)}`)
	}
}
