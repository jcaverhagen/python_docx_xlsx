# python docx and xlsx

Simple docx and xlsx editor written in python.

Changelog:
##### 13-08-2015
- Implemented dashed borders for tables and paragraph fontsize
- Fix: test and reference in paragraph.py
- Keeping style file from old document
- searchAndReplace optimize

##### 12-08-2015
- Improved searchAndReplace (fix for replacing in different t elements)

##### 11-08-2015
- Improved searchAndReplace (footer,header,document)
- Fix: looping through paragraphs in searchAndReplace
- Improved search and replace function in headerfile

##### 10-08-2015
- Fix: double reference when adding footer[].xml
- Improved search and replace function in FooterFile

##### 05-08-2015
- Implemented method to add formules to column
- Implemented adding data to sheet per column, types numbers and text
- Implemented adding work sheets to xlsx file
- Moved classes from docx_items to universal folder, element and default prefixes

##### 04-08-2015
- Implemented basic functionality to create xlsx files
- Implemented hyperlinks on images
- Refactored add image method, only add necessary image extension to [Content_Types].xml
- Implemented anchors in document with hyperlink
- Fix: Make text hyperlink, hold the positions after adding hyperlink
- Fix: When making text hyperlink leave the spaces
- Refactoring

##### 03-08-2015
- Implemented Footers (even, odd, first and default settings)
- Fixed search and replace for header files
- Implemented different headers for odd and even pages
- Implemented different first header on page 1
- Add default header implemented, option for different first and even pages coming soon
- Added option to have multiple levels in list (default structure file, future giving style with it)
- Reformated code

##### 31-07-2015
- Fixed list position and order
- Fixed table column order
- Implemented breaks and page breaks as elements
- Fixed issue when add an other image extension to existed document. (temp fix add as default all images extensions to [Content_Types].xml)
- option to post width and height with new image (percentages). Default 100%
- Implemented method to add images to document (just the basics)

##### 30-07-2015
- Added primary colors for font color and underline color
- Implemented style for paragraph, bold, italic, underline, uppercase, color and font
- Refactoring bunch of code, improvements and working with file and creating from scratch
- Option to create an empty word document from scratch with default settings
- Added default xml files, for creating new docx files
- Added option to give an Heading with an paragraph, implemented heading 1 to 5

##### 29-07-2015
- Code improvement
- Added type of list, choose between numeric list or bullets
- Option to add list to custom position in document, first/last/aftertext/beforetext.
- Implemented numeric list to end of document
- Universal method to add element to document with position (first, last, aftertext, beforetext)
- Option to add an table to end of document, create table, add rows, add to document

##### 28-07-2015
- Method to make specific text an hyperlink
- Function to add an hyperlink into an new paragraph with custom text as url
- Added default style and spacing to paragraph element
- Add an paragraph after or before paragraph containing specific text
- Add paragraph to end or beginning of file

##### 27-07-2015
- Including header / footer search and replace
- Search and replace function
- Copy function
