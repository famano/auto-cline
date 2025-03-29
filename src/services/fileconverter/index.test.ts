import * as fs from "fs/promises"
import { after, before, describe, it } from "mocha"
import * as os from "os"
import * as path from "path"
import "should"
import * as sinon from "sinon"
import {
	convertXlsxToCsv,
	convertCsvToXlsx,
	convertDocxToMd,
	convertMdToDocx,
	convertPptxToMd,
	convertMdToPptx,
	convertDirectory,
} from "./index"
import * as xlsxCsvConverter from "./xlsx-csv-converter"
import * as docxMdConverter from "./docx-md-converter"
import * as pptxMdConverter from "./pptx-md-converter"
import * as directoryConverter from "./directory-converter"

// This is an integration test to ensure exports are properly set up

// Create stubs for the converter functions
let xlsxToCsvStub: sinon.SinonStub
let csvToXlsxStub: sinon.SinonStub
let docxToMdStub: sinon.SinonStub
let mdToDocxStub: sinon.SinonStub
let pptxToMdStub: sinon.SinonStub
let mdToPptxStub: sinon.SinonStub
let dirConversionStub: sinon.SinonStub

describe("File Converter Module", () => {
	before(() => {
		// Setup stubs
		xlsxToCsvStub = sinon.stub(xlsxCsvConverter, "convertXlsxToCsv").resolves("xlsx-to-csv mock")
		csvToXlsxStub = sinon.stub(xlsxCsvConverter, "convertCsvToXlsx").resolves("csv-to-xlsx mock")

		docxToMdStub = sinon.stub(docxMdConverter, "convertDocxToMd").resolves("docx-to-md mock")
		mdToDocxStub = sinon.stub(docxMdConverter, "convertMdToDocx").resolves("md-to-docx mock")

		pptxToMdStub = sinon.stub(pptxMdConverter, "convertPptxToMd").resolves("pptx-to-md mock")
		mdToPptxStub = sinon.stub(pptxMdConverter, "convertMdToPptx").resolves("md-to-pptx mock")

		dirConversionStub = sinon.stub(directoryConverter, "convertDirectory").resolves("directory mock")
	})

	after(() => {
		// Restore all stubs
		xlsxToCsvStub.restore()
		csvToXlsxStub.restore()
		docxToMdStub.restore()
		mdToDocxStub.restore()
		pptxToMdStub.restore()
		mdToPptxStub.restore()
		dirConversionStub.restore()
	})

	it("should export XLSX/CSV converter functions", async () => {
		const xlsxToCsv = await convertXlsxToCsv("test.xlsx")
		xlsxToCsv.should.equal("xlsx-to-csv mock")

		const csvToXlsx = await convertCsvToXlsx("test.csv")
		csvToXlsx.should.equal("csv-to-xlsx mock")
	})

	it("should export DOCX/MD converter functions", async () => {
		const docxToMd = await convertDocxToMd("test.docx")
		docxToMd.should.equal("docx-to-md mock")

		const mdToDocx = await convertMdToDocx("test.md")
		mdToDocx.should.equal("md-to-docx mock")
	})

	it("should export PPTX/MD converter functions", async () => {
		const pptxToMd = await convertPptxToMd("test.pptx")
		pptxToMd.should.equal("pptx-to-md mock")

		const mdToPptx = await convertMdToPptx("test.md")
		mdToPptx.should.equal("md-to-pptx mock")
	})

	it("should export directory converter function", async () => {
		const dirConversion = await convertDirectory("test-dir", "xlsx-to-csv")
		dirConversion.should.equal("directory mock")
	})
})
