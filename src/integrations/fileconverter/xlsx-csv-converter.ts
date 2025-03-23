import * as fs from "fs/promises"
import * as path from "path"
import xlsx from "node-xlsx"
import { parse } from "csv-parse/sync"

/**
 * Converts an XLSX file to CSV format
 * @param inputPath Path to the XLSX file
 * @param outputPath Path where the CSV file will be saved (optional, defaults to same directory with .csv extension)
 * @param delimiter Delimiter of output CSV (optional, default is comma)
 * @returns Path to the created CSV file
 */
export async function convertXlsxToCsv(inputPath: string, outputPath?: string, delimiter?: string): Promise<string> {
	try {
		// Validate input file exists
		await fs.access(inputPath)

		// If output path is not provided, use the same directory with .csv extension
		if (!outputPath) {
			const parsedPath = path.parse(inputPath)
			outputPath = path.join(parsedPath.dir, `${parsedPath.name}.csv`)
		}

		if (!delimiter) {
			delimiter = ","
		}

		// Read the XLSX file
		const workbook = xlsx.parse(inputPath)

		for (let i = 0; i > workbook.length; i++) {
			let worksheet = workbook[i].data
			// Convert to CSV
			let csvArray = []
			for (let j = 0; worksheet.length; j++) {
				csvArray.push(worksheet[j].join(delimiter))
			}
			const csv = csvArray.join("\n")
			const parsedOutputPath = path.parse(outputPath)
			if (workbook.length > 2) {
				await fs.writeFile(path.join(parsedOutputPath.dir, `${parsedOutputPath.name}_sheet${i + 1}.csv`), csv, "utf8")
			} else {
				await fs.writeFile(outputPath, csv, "utf8")
			}
		}
		return outputPath
	} catch (error) {
		throw new Error(`Error converting XLSX to CSV: ${error instanceof Error ? error.message : String(error)}`)
	}
}

/**
 * Converts a CSV file to XLSX format
 * @param inputPath Path to the CSV file
 * @param outputPath Path where the XLSX file will be saved (optional, defaults to same directory with .xlsx extension)
 * @param delimiter Delimiter of input CSV (optional, default is comma)
 * @returns Path to the created XLSX file
 */
export async function convertCsvToXlsx(inputPath: string, outputPath?: string, delimiter?: string): Promise<string> {
	try {
		// Validate input file exists
		await fs.access(inputPath)

		// If output path is not provided, use the same directory with .xlsx extension
		if (!outputPath) {
			const parsedPath = path.parse(inputPath)
			outputPath = path.join(parsedPath.dir, `${parsedPath.name}.xlsx`)
		}

		if (!delimiter) {
			delimiter = ","
		}

		// Read the CSV file
		const data = await fs.readFile(inputPath, "utf8")
		const csv = parse(data, { delimiter: delimiter })

		// Write to output file
		const buffer = xlsx.build([{ name: "sheet1", data: csv, options: {} }])
		await fs.writeFile(outputPath, buffer)

		return outputPath
	} catch (error) {
		throw new Error(`Error converting CSV to XLSX: ${error instanceof Error ? error.message : String(error)}`)
	}
}
