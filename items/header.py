
class HeaderFile :

	path = 'word/header{filenumber}.xml'
	
	def __init__(self, text, filenumber) :

		self.path = self.path.format(filenumber=filenumber)
		self.text = text

		self.xmlString = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
			<w:hdr xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml">
			<w:p>
			<w:pPr>
			<w:pStyle w:val="Header"/>
			</w:pPr>
			<w:r>
			<w:t>{text}</w:t>
			</w:r>
			</w:p>
			</w:hdr>"""

	def getXml(self) :
		return self.xmlString.format(text=self.text)