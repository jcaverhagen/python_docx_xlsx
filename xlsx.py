#!/usr/bin/python
import zipfile
from xlsx_items.document import (
	ContentTypeFile, RelationshipFile, AppFile, CoreFile, DocumentRelationshipFile, ThemeFile,
	StylesFile, WorkbookFile, SheetFile
)

class Document :

	def __init__(self) :
		self.files = {}

		self.files['[Content_Types].xml'] = ContentTypeFile()

	def save(self, filename) :
		xlsxFile = zipfile.ZipFile(filename, mode='w', compression=zipfile.ZIP_DEFLATED)

		xlsxFile.writestr(RelationshipFile().path, RelationshipFile().getXml())
		xlsxFile.writestr(AppFile().path, AppFile().getXml())
		xlsxFile.writestr(CoreFile().path, CoreFile().getXml())
		xlsxFile.writestr(DocumentRelationshipFile().path, DocumentRelationshipFile().getXml())
		xlsxFile.writestr(ThemeFile().path, ThemeFile().getXml())
		xlsxFile.writestr(StylesFile().path, StylesFile().getXml())
		xlsxFile.writestr(WorkbookFile().path, WorkbookFile().getXml())
		xlsxFile.writestr(SheetFile('1').getPath(), SheetFile('1').getXml())

		for key, value in self.files.items() :
			xlsxFile.writestr(key, value.getXml())		
