import items.styles
from docx import Document

#document = Document("test.docx")

"""
#search and replace
document.searchAndReplace("klaas", "jan")

#add paragraph
document.addParagraph("vanuit python toegevoegd onderaan")
document.addParagraph("vanuit python toegevoegd na klaas", 'aftertext:klaas')

#add hyperlink
document.addHyperlink("Google vanuit python toegevoegd test", "http://www.google.nl", 'first')
document.addHyperlink("link before test", "http://www.google.nl", 'beforetext:Test')
document.addHyperlink("link after test", "http://www.google.nl", 'aftertext:Test')

#replace text with hyperlink
document.makeTextHyperlink("klaas", "http://www.klaas.nl")

#add table with 3 columns
table = document.addTable(5000, 3)
table.addRow({'test1', 'test2', 'test3'})
table.addRow({'test4', 'test5', 'test6'})
table.addRow({'test7', 'test8', 'test9'})
document.closeTable(table)

#add table with 2 columns at top
table = document.addTable(5000, 2)
table.addRow({'test1', 'test2'})
table.addRow({'test4', 'test5'})
table.addRow({'test7', 'test8'})
document.closeTable(table, position='first')

#list
listItem = document.addList('first', type='bullet')
listItem.addItem('piet')
listItem.addItem('henk')
listItem.addItem('Jan')
listItem.addItem('Frits')
listItem.addItem('Bert')
document.closeList(listItem)
"""

#add heading
#document.addParagraph("HEADING 1", 'first', 'Heading1')
document = Document()
document.save("test1.docx")