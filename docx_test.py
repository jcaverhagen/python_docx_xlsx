from docx import Document

document = Document()
openDoc = document.openFile("test.docx")

docData = document.readHeader(openDoc)
for d in docData :
	print d
