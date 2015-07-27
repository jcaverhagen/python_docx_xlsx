from docx import Document

document = Document("test.docx")

#docData = document.readDocument()
#for d in docData :
#	print (d)

document.searchAndReplace("klaas", "jan")
document.searchAndReplace("Test", "Henk")

document.save("test1.docx")