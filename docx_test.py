from docx import Document

document = Document("test.docx")

docData = document.readDocument()
for d in docData :
	print (d)

document.save()