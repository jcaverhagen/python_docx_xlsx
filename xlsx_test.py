from xlsx import Document

document = Document()

sheet1 = document.addSheet()
sheet2 = document.addSheet()

sheet1.addData('A2', '200', type='number')
sheet1.addData('A3', '200', type='number')

sheet1.addData('D2', '100')
sheet1.addData('D3', '120')

sheet1.addData('D4', '120')
sheet1.addData('E7', 'tekst', type='text')

document.save('test_excel1.xlsx')