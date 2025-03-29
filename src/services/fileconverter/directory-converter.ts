import * as fs from "fs/promises"
import * as path from "path"
import { convertXlsxToCsv, convertCsvToXlsx } from "./xlsx-csv-converter"
import { convertDocxToMd, convertMdToDocx } from "./docx-md-converter"
import { convertPptxToMd, convertMdToPptx } from "./pptx-md-converter"

/**
 * Conversion mode types
 */
export type ConversionMode = "xlsx-to-csv" | "csv-to-xlsx" | "docx-to-md" | "md-to-docx" | "pptx-to-md" | "md-to-pptx"

/**
 * Options for directory conversion
 */
export interface DirectoryConversionOptions {
	/**
	 * Path to the reference document for pandoc conversions (optional)
	 */
	referenceDocPath?: string

	/**
	 * Output directory path (optional, defaults to same directory as input)
	 */
	outputDir?: string

	/**
	 * Whether to process subdirectories recursively (default: false)
	 */
	recursive?: boolean
}

/**
 * Result of a directory conversion operation
 */
export interface ConversionResult {
	/**
	 * Number of files successfully converted
	 */
	successCount: number

	/**
	 * Number of files that failed to convert
	 */
	failCount: number

	/**
	 * List of paths to successfully converted files
	 */
	convertedFiles: string[]

	/**
	 * Map of failed files with error messages
	 */
	failedFiles: Map<string, string>
}

/**
 * Converts all files of a specific type in a directory
 * @param dirPath Path to the directory containing files to convert
 * @param mode Conversion mode
 * @param options Additional conversion options
 * @returns Conversion result statistics
 */
export async function convertDirectory(
	dirPath: string,
	mode: ConversionMode,
	options: DirectoryConversionOptions = {},
): Promise<string> {
	// Initialize result
	const result: ConversionResult = {
		successCount: 0,
		failCount: 0,
		convertedFiles: [],
		failedFiles: new Map(),
	}

	try {
		// Validate directory exists
		await fs.access(dirPath)

		// Get source file extension based on mode
		const sourceExt = getSourceExtension(mode)

		// Process directory
		await processDirectory(dirPath, mode, options, result, sourceExt)

		return formatResult(result)
	} catch (error) {
		throw new Error(`Error converting directory: ${error instanceof Error ? error.message : String(error)}`)
	}
}

function formatResult(result: ConversionResult): string {
	let formatted = `Converted ${result.successCount} files, failed ${result.failCount}.\n\n`
	formatted += "success:\n"
	for (const file of result.convertedFiles) {
		formatted += `${file}\n`
	}
	formatted += "\nfailed:\n"
	result.failedFiles.forEach((error, file) => {
		formatted += `${file}: ${error}\n`
	})
	return formatted
}

/**
 * Process a directory for conversion
 */
async function processDirectory(
	dirPath: string,
	mode: ConversionMode,
	options: DirectoryConversionOptions,
	result: ConversionResult,
	sourceExt: string,
): Promise<void> {
	// Read directory contents
	const entries = await fs.readdir(dirPath, { withFileTypes: true })

	// Process each entry
	for (const entry of entries) {
		const entryPath = path.join(dirPath, entry.name)

		if (entry.isDirectory() && options.recursive) {
			// Recursively process subdirectory
			await processDirectory(entryPath, mode, options, result, sourceExt)
		} else if (entry.isFile() && path.extname(entry.name).toLowerCase() === sourceExt) {
			// Process matching file
			let outputPath: string | undefined

			// Determine output path if outputDir is specified
			if (options.outputDir) {
				const relativePath = path.relative(dirPath, entryPath)
				const targetExt = getTargetExtension(mode)
				const outputFileName = path.basename(relativePath, sourceExt) + targetExt
				outputPath = path.join(options.outputDir, outputFileName)

				// Ensure output directory exists
				await fs.mkdir(path.dirname(outputPath), { recursive: true })
			}

			try {
				// Convert file based on mode
				const convertedPath = await convertFile(entryPath, mode, outputPath, options.referenceDocPath)
				result.successCount++
				result.convertedFiles.push(convertedPath)
			} catch (error) {
				result.failCount++
				result.failedFiles.set(entryPath, error instanceof Error ? error.message : String(error))
			}
		}
	}
}

/**
 * Convert a single file based on the specified mode
 */
export async function convertFile(
	filePath: string,
	mode: ConversionMode,
	outputPath?: string,
	referenceDocPath?: string,
): Promise<string> {
	switch (mode) {
		case "xlsx-to-csv":
			return await convertXlsxToCsv(filePath, outputPath)
		case "csv-to-xlsx":
			return await convertCsvToXlsx(filePath, outputPath)
		case "docx-to-md":
			return await convertDocxToMd(filePath, outputPath)
		case "md-to-docx":
			return await convertMdToDocx(filePath, outputPath, referenceDocPath)
		case "pptx-to-md":
			return await convertPptxToMd(filePath, outputPath)
		case "md-to-pptx":
			return await convertMdToPptx(filePath, outputPath, referenceDocPath)
		default:
			throw new Error(`Unsupported conversion mode: ${mode}`)
	}
}

/**
 * Get source file extension based on conversion mode
 */
function getSourceExtension(mode: ConversionMode): string {
	switch (mode) {
		case "xlsx-to-csv":
			return ".xlsx"
		case "csv-to-xlsx":
			return ".csv"
		case "docx-to-md":
			return ".docx"
		case "md-to-docx":
			return ".md"
		case "pptx-to-md":
			return ".pptx"
		case "md-to-pptx":
			return ".md"
		default:
			throw new Error(`Unsupported conversion mode: ${mode}`)
	}
}

/**
 * Get target file extension based on conversion mode
 */
function getTargetExtension(mode: ConversionMode): string {
	switch (mode) {
		case "xlsx-to-csv":
			return ".csv"
		case "csv-to-xlsx":
			return ".xlsx"
		case "docx-to-md":
			return ".md"
		case "md-to-docx":
			return ".docx"
		case "pptx-to-md":
			return ".md"
		case "md-to-pptx":
			return ".pptx"
		default:
			throw new Error(`Unsupported conversion mode: ${mode}`)
	}
}
