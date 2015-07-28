from docx import Document

document = Document("test.docx")

document.searchAndReplace("klaas", "jan")
document.addParagraph("vanuit python toegevoegd onderaan")
document.addParagraph("vanuit python toegevoegd bovenaan", position='first')
document.addParagraph("vanuit python toegevoegd na klaas", aftertext='klaas')
document.addHyperlink("Google vanuit python toegevoegd test", "http://www.google.nl")

document.save("test1.docx")