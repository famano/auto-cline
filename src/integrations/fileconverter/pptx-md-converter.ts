import * as fs from "fs/promises"
import * as path from "path"

const nodePandoc = require("node-pandoc")

/**
 * Converts a PPTX file to Markdown format using pandoc
 * @param inputPath Path to the PPTX file
 * @param outputPath Path where the Markdown file will be saved (optional, defaults to same directory with .md extension)
 * @returns Path to the created Markdown file
 */
export async function convertPptxToMd(inputPath: string, outputPath?: string): Promise<string> {
	try {
		// Validate input file exists
		await fs.access(inputPath)

		// If output path is not provided, use the same directory with .md extension
		if (!outputPath) {
			const parsedPath = path.parse(inputPath)
			outputPath = path.join(parsedPath.dir, `${parsedPath.name}.md`)
		}

		// Build pandoc arguments
		const args = ["-f", "pptx", "-t", "markdown", "-o", outputPath]

		// Execute pandoc command
		await nodePandoc(inputPath, args)

		return outputPath
	} catch (error) {
		throw new Error(`Error converting PPTX to Markdown: ${error instanceof Error ? error.message : String(error)}`)
	}
}

/**
 * Converts a Markdown file to PPTX format using pandoc
 * @param inputPath Path to the Markdown file
 * @param outputPath Path where the PPTX file will be saved (optional, defaults to same directory with .pptx extension)
 * @param referenceDocPath Optional path to a reference doc for styling
 * @returns Path to the created PPTX file
 */
export async function convertMdToPptx(inputPath: string, outputPath?: string, referenceDocPath?: string): Promise<string> {
	try {
		// Validate input file exists
		await fs.access(inputPath)

		// If output path is not provided, use the same directory with .pptx extension
		if (!outputPath) {
			const parsedPath = path.parse(inputPath)
			outputPath = path.join(parsedPath.dir, `${parsedPath.name}.pptx`)
		}

		// Build pandoc arguments
		const args = ["-f", "markdown", "-t", "pptx", "-o", outputPath]

		// Add reference-doc if provided
		if (referenceDocPath) {
			await fs.access(referenceDocPath) // Validate reference doc exists
			args.push(`--reference-doc=${referenceDocPath}`)
		}

		// Execute pandoc command
		await nodePandoc(args)

		return outputPath
	} catch (error) {
		throw new Error(`Error converting Markdown to PPTX: ${error instanceof Error ? error.message : String(error)}`)
	}
}
