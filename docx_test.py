from docx import Document

document = Document("test.docx")

#search and replace
document.searchAndReplace("klaas", "jan")

#add paragraph
document.addParagraph("vanuit python toegevoegd onderaan")
document.addParagraph("vanuit python toegevoegd bovenaan", 'first')
document.addParagraph("vanuit python toegevoegd na klaas", 'aftertext:klaas')

#add hyperlink
document.addHyperlink("Google vanuit python toegevoegd test", "http://www.google.nl", 'first')
document.addHyperlink("link before test", "http://www.google.nl", 'beforetext:Test')
document.addHyperlink("link after test", "http://www.google.nl", 'aftertext:Test')

#replace text with hyperlink
document.makeTextHyperlink("klaas", "http://www.klaas.nl")

#add table with 3 columns
document.addTable(5000, 3)
document.addRow({'test1', 'test2', 'test3'})
document.addRow({'test4', 'test5', 'test6'})
document.addRow({'test7', 'test8', 'test9'})
document.addTableToDoc()

#add table with 2 columns at top
document.addTable(5000, 2)
document.addRow({'test1', 'test2'})
document.addRow({'test4', 'test5'})
document.addRow({'test7', 'test8'})
document.addTableToDoc('first')

#list
listItem = document.addList('first', type='bullet')
listItem.addItem('piet')
listItem.addItem('henk')
listItem.addItem('Jan')
listItem.addItem('Frits')
listItem.addItem('Bert')
document.closeList(listItem)

document.save("test1.docx")