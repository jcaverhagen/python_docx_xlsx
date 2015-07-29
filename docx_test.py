from docx import Document

document = Document("test.docx")

document.searchAndReplace("klaas", "jan")
document.addParagraph("vanuit python toegevoegd onderaan")
document.addParagraph("vanuit python toegevoegd bovenaan", 'first')
document.addParagraph("vanuit python toegevoegd na klaas", 'aftertext:klaas')

document.addHyperlink("Google vanuit python toegevoegd test", "http://www.google.nl", 'first')
document.addHyperlink("link before test", "http://www.google.nl", 'beforetext:Test')
document.addHyperlink("link after test", "http://www.google.nl", 'aftertext:Test')

document.makeTextHyperlink("klaas", "http://www.klaas.nl")

document.addTable(5000, 3)
document.addRow({'test1', 'test2', 'test3'})
document.addRow({'test4', 'test5', 'test6'})
document.addRow({'test7', 'test8', 'test9'})
document.addTableToDoc()

document.addTable(5000, 2)
document.addRow({'test1', 'test2'})
document.addRow({'test4', 'test5'})
document.addRow({'test7', 'test8'})
document.addTableToDoc('first')

document.save("test1.docx")