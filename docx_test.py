from docx import Document

"""
DOCUMENT WITH INPUT FILE
"""
#using old file and add text
document = Document('test.docx')
#paragraph
document.addParagraph("vanuit python toegevoegd onderaan")
document.addParagraph("vanuit python toegevoegd na klaas", 'aftertext:klaas')
#search and replace
document.searchAndReplace("klaas", "jan")
#hyperlink
document.addHyperlink("Google vanuit python toegevoegd test", "http://www.google.nl", 'first')
document.addHyperlink("link before test", "http://www.google.nl", 'beforetext:Test')
document.addHyperlink("link after test", "http://www.google.nl", 'aftertext:Test')
#replace text with hyperlink
document.makeTextHyperlink("Test", "http://www.test.nl")

#add anchor
document.makeTextHyperlink("jan", "anker", anchor='anker')

#add table with 3 columns
table = document.addTable(5000, 3)
table.addRow(['test1', 'test2', 'test3'])
table.addRow(['test4', 'test5', 'test6'])
table.addRow(['test7', 'test8', 'test9'])
document.closeTable(table)

#list
numberedList = document.addList('first')
numberedList.addItem('piet')
numberedList.addItem('henk')
numberedList.addItem('Jan')
numberedList.addItem('Klaas', level=1)
numberedList.addItem('Rene', level=1)
numberedList.addItem('Frits')
numberedList.addItem('Bert')
document.closeList(numberedList)

document.addImage('image.jpg', width='20%', height='20%', url='http://www.nu.nl')

document.addHeader('Header tekst eerste pagina', 'first')

document.save("test1.docx")

"""
NEW DOCUMENT WITH NO INPUT FILE, CREATING FROM SCRATCH
"""
#creating complete file
document = Document()
#paragraph
document.addParagraph("vanuit python toegevoegd onderaan", styles={'bold' : True, 'italic' : True, 'underline' : 'red', 
					'uppercase' : True, 'color' : 'red', 'font' : 'Times New Roman'})
document.addParagraph("vanuit python toegevoegd na klaas", 'aftertext:klaas')

#search and replace
document.searchAndReplace("klaas", "jan")
#hyperlink
document.addHyperlink("Google vanuit python toegevoegd test", "http://www.google.nl", 'first')
document.addHyperlink("link before test", "http://www.google.nl", 'beforetext:Test')
document.addHyperlink("link after test", "http://www.google.nl", 'aftertext:Test')
#replace text with hyperlink
document.makeTextHyperlink("jan", "http://www.klaas.nl")

#pagebreak
document.insertBreak('page')

#add table with 3 columns
table = document.addTable(5000, 3)
table.addRow(['test1', 'test2', 'test3'])
table.addRow(['test4', 'test5', 'test6'])
table.addRow(['test7', 'test8', 'test9'])
document.closeTable(table)
#add table with 2 columns at top
table = document.addTable(5000, 2, position='first')
table.addRow(['test1', 'test2'])
table.addRow(['test4', 'test5'])
table.addRow(['test7', 'test8'])
document.closeTable(table)

#insert page break
document.insertBreak('page')

#list
numberedList = document.addList('first')
numberedList.addItem('piet')
numberedList.addItem('henk')
numberedList.addItem('Jan')
numberedList.addItem('Klaas', level=1)
numberedList.addItem('Rene', level=1)
numberedList.addItem('Klaas', level=2)
numberedList.addItem('Rene', level=2)
numberedList.addItem('Frits')
numberedList.addItem('Bert')
document.closeList(numberedList)

#add heading
document.addParagraph("This file is created through python for testing", 'first', 'Heading1', styles={'bold' : True})

#add image to end of file
document.addImage('image.jpg', position='first', width='10%', height='10%')
document.addImage('image1.jpg', width='30%', height='30%', url='http://www.tweakers.net')

#pagebreak
document.insertBreak('page')
document.addParagraph("vanuit python toegevoegd onderaan", styles={'bold' : True, 'italic' : True, 'underline' : 'red', 
					'uppercase' : True, 'color' : 'red', 'font' : 'Times New Roman'})

# adding headers
document.addHeader('Header tekst even pages', 'even')
document.addHeader('Header tekst odd pages', 'default')

document.addFooter('Default footer text', 'default')
document.addFooter('Different footer text', 'first')

#search and replace on headers
document.searchAndReplace("odd pages", "oneven paginas")
document.searchAndReplace("even pages", "even paginas")

#add anchor
document.makeTextHyperlink("vanuit", "python", anchor='anker')

document.save("test2.docx")