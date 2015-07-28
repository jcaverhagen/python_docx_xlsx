from docx import Document

document = Document("test.docx")

document.searchAndReplace("klaas", "jan")
document.addParagraph("vanuit python toegevoegd onderaan")
document.addParagraph("vanuit python toegevoegd bovenaan", position='first')
document.addParagraph("vanuit python toegevoegd na klaas", aftertext='klaas') #not working properly
document.addHyperlink("Google vanuit python toegevoegd test", "http://www.google.nl", position='first')
document.addHyperlink("link before test", "http://www.google.nl", beforetext='Test')
document.addHyperlink("link after test", "http://www.google.nl", aftertext='Test')

document.save("test1.docx")