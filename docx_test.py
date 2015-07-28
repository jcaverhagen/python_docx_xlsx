from docx import Document

document = Document("test.docx")

#docData = document.readDocument()
#for d in docData :
#	print (d)

#document.searchAndReplace("klaas", "jan")
document.addParagraph("vanuit python toegevoegd")

document.save("test1.docx")