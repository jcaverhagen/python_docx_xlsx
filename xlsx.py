#!/usr/bin/python
import zipfile
from xlsx_items.document import (
	ContentTypeFile, RelationshipFile, AppFile, CoreFile, DocumentRelationshipFile, ThemeFile,
	StylesFile, WorkbookFile, SheetFile
)

class Document :

	sheets = []

	def __init__(self) :
		self.files = {}

		self.files['[Content_Types].xml'] = ContentTypeFile()
		self.files['xl/workbook.xml'] = WorkbookFile()
		self.files['xl/_rels/workbook.xml.rels'] = DocumentRelationshipFile()

	def addSheet(self) :
		sheetId = len(self.sheets) + 1

		#adding to [Content_Types].xml
		self.files['[Content_Types].xml'].addOverride('sheet', sheetId)

		#adding to worksheet relation file
		rel_id = self.files['xl/_rels/workbook.xml.rels'].addRelation('sheet', sheetId=sheetId)

		#adding to workbook.xml
		self.files['xl/workbook.xml'].addSheet(sheetId, rel_id)

		sheet = SheetFile(str(sheetId))
		self.sheets.append(sheet)
		return sheet

	def save(self, filename) :
		xlsxFile = zipfile.ZipFile(filename, mode='w', compression=zipfile.ZIP_DEFLATED)

		xlsxFile.writestr(RelationshipFile().path, RelationshipFile().getXml())
		xlsxFile.writestr(AppFile().path, AppFile().getXml())
		xlsxFile.writestr(CoreFile().path, CoreFile().getXml())
		xlsxFile.writestr(ThemeFile().path, ThemeFile().getXml())
		xlsxFile.writestr(StylesFile().path, StylesFile().getXml())
		
		#add every sheet to xlsx
		for sheet in self.sheets :
			xlsxFile.writestr(sheet.getPath(), sheet.getXml())

		for key, value in self.files.items() :
			xlsxFile.writestr(key, value.getXml())		
