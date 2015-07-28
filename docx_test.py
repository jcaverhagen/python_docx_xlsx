from docx import Document

document = Document("test.docx")

#docData = document.readDocument()
#for d in docData :
#	print (d)

#document.searchAndReplace("klaas", "jan")
document.addParagraph("vanuit python toegevoegd onderaan")
document.addParagraph("vanuit python toegevoegd bovenaan", position='first')
document.addParagraph("vanuit python toegevoegd na klaas", aftertext='klaas')

document.save("test1.docx")