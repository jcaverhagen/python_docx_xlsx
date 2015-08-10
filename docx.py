#!/usr/bin/python
import zipfile
from lxml import etree
from docx_items.paragraph import Paragraph, Break
from docx_items.hyperlink import Hyperlink
from docx_items.table import Table
from docx_items.list import List
from docx_items.image import Image
from docx_items.document import (
 StyleFile, AppFile, RelationshipFile, DocumentRelationshipFile, CoreFile, DocumentFile, HeaderFile, FooterFile,
 ContentTypeFile, NumberingFile, SettingsFile, FontTableFile, StylesWithEffectsFile, WebSettingsFile, ThemeFile
)
from PIL import Image as PILImage

class Document :
    
    def __init__(self, filename=None) :
        self.files = {}
        self.images = {}
        self.filename = filename

        if filename :
            if not zipfile.is_zipfile(filename) :
                raise TypeError("Not an correct docx file.")

            self._doc = zipfile.ZipFile(filename)

            #scan zipfile for header/document and footer xml files
            for path in self._doc.namelist() :
                if path == 'word/document.xml' : 
                    self.files[path] = DocumentFile(etree.fromstring(self._doc.read(path)))
                if path == 'word/_rels/document.xml.rels' :
                    self.files[path] = DocumentRelationshipFile(etree.fromstring(self._doc.read(path)))
                if path == '[Content_Types].xml' :
                    self.files[path] = ContentTypeFile(etree.fromstring(self._doc.read(path)))
                if path == 'word/settings.xml' :
                    self.files[path] = SettingsFile(etree.fromstring(self._doc.read(path)))
                
                if 'footer' in path :
                    self.files[path] = FooterFile(path, etree.fromstring(self._doc.read(path)))
        else :
            self.files['word/document.xml'] = DocumentFile()
            self.files['word/_rels/document.xml.rels'] = DocumentRelationshipFile()
            self.files['[Content_Types].xml'] = ContentTypeFile()
            self.files['word/settings.xml'] = SettingsFile()

        self.doc = self.files['word/document.xml']

    #add paragraph as first
    def addParagraph(self, text, position='last', style='NormalWeb', styles=None) :
        paragraph = Paragraph(text, style, styles).get()
        self.doc.addElement(paragraph, position)

    #add hyperlink to document
    def addHyperlink(self, text, url, position='last', anchor=None) :
        rel_id = 0
        if anchor == None :
            rel_id = self.files['word/_rels/document.xml.rels'].addRelation('hyperlink', url)

        paraElement = Paragraph().get()
        hyperlink = Hyperlink(text, str(rel_id), url, anchor)
        paraElement.append(hyperlink.get())

        self.doc.addElement(paraElement, position)

    #method to make specific test an hyperlink
    def makeTextHyperlink(self, text, url, anchor=None) :
        rel_id = 0
        if anchor == None :
            rel_id = self.files['word/_rels/document.xml.rels'].addRelation('hyperlink', url=url)

        hyperlink = Hyperlink(text, str(rel_id), url, anchor)
        
        self.doc.makeTextHyperlink(text, hyperlink.get())
        if anchor != None :
            self.doc.addAnchorToText(url, anchor)

    #init an new table and returning it to caller
    def addTable(self, width, columns, position='last') :
        return Table(width, columns, position=position)

    #close table and add it to document
    def closeTable(self, table) :
        self.doc.addElement(table.get(), table.getPosition())

    #init an new list and return object to caller
    def addList(self, position='last') :
        return List(position)
        
    #close list and inserting it in document
    def closeList(self, listItem) :
        listItems = listItem.get()
        
        if listItem.getPosition() == 'first' :
             for item in reversed(listItems) :
                self.doc.addElement(item, listItem.getPosition())
        else :
            for item in listItems :
                self.doc.addElement(item, listItem.getPosition())

    #add header
    def addHeader(self, text, headertype='default') :
        if headertype == 'first' :
            filenumber = 3
        elif headertype == 'even' :
            filenumber = 1
            self.files['word/settings.xml'].enableEvenAndOddHeaders()
        else :
            filenumber = 2

        rel_id = self.files['word/_rels/document.xml.rels'].addRelation('header', headerfootertype=headertype)
        self.files['word/header' + str(filenumber) + '.xml'] = HeaderFile(text, str(filenumber))

        self.files['[Content_Types].xml'].addOverride('header', filenumber)

        self.doc.addReference('header', headertype, rel_id)

    #add footer
    def addFooter(self, text, footertype='Default') :
        if footertype == 'first' :
            filenumber = 3
        elif footertype == 'even' :
            filenumber = 1
            self.files['word/settings.xml'].enableEvenAndOddHeaders()
        else :
            filenumber = 2

        rel_id = self.files['word/_rels/document.xml.rels'].addRelation('footer', headerfootertype=footertype)
        if rel_id > 0 :
            self.doc.addReference('footer', footertype, rel_id)
        self.files['word/footer' + str(filenumber) + '.xml'] = FooterFile(str(filenumber))

        self.files['[Content_Types].xml'].addOverride('footer', filenumber)

        self.files['word/footer' + str(filenumber) + '.xml'].addText(text) 

        return self.files['word/footer' + str(filenumber) + '.xml']

    def addImage(self, image, position='last', width='100%', height='100%', url=None) :
        count = 1
        if self.filename is not None :
            for path in self._doc.namelist() :
                if 'media' in path :
                    count = count + 1
            id = count
        else :
            id = len(self.images)
        
        extension = image.split('.')[1]
        imagename = 'image' + str(id) + '.' + extension

        self.images[image] = imagename

        rel_id = self.files['word/_rels/document.xml.rels'].addRelation('image', imagename=imagename)
        self.files['[Content_Types].xml'].addImageExtension(extension)

        url_id = None
        if url is not None :
            url_id = self.files['word/_rels/document.xml.rels'].addRelation('hyperlink', url=url)

        image = Image(image, id, rel_id, width, height, url_id)
        self.doc.addElement(image.get(), position)

    def insertBreak(self, type, position='last') :
        breakEl = Break(type)
        self.doc.addElement(breakEl.get(), position)

    #save document with new values
    def save(self, filename) :
        docxFile = zipfile.ZipFile(filename, mode='w', compression=zipfile.ZIP_DEFLATED)
        if not self.filename :
            docxFile.writestr(RelationshipFile().path, RelationshipFile().getXml())
            docxFile.writestr(AppFile().path, AppFile().getXml())
            docxFile.writestr(CoreFile().path, CoreFile().getXml())
            docxFile.writestr(NumberingFile().path, NumberingFile().getXml())
            docxFile.writestr(FontTableFile().path, FontTableFile().getXml())
            docxFile.writestr(WebSettingsFile().path, WebSettingsFile().getXml())
            docxFile.writestr(StyleFile().path, StyleFile().getXml())
            docxFile.writestr(ThemeFile().path, ThemeFile().getXml())
            docxFile.writestr(StylesWithEffectsFile().path, StylesWithEffectsFile().getXml())
            
        else :
            
            #copy from old docx every file except the files that are in files list
            for path in self._doc.namelist() :
                if path not in self.files :
                    if path == 'word/styles.xml' :
                        styleFile = etree.fromstring(StyleFile().getXml())
                        docxFile.writestr('word/styles.xml', etree.tostring(styleFile))
                    else : docxFile = self.copyToXML(docxFile, path)

        #add files from file list to docx
        for key, value in self.files.items() :
            docxFile.writestr(key, value.getXml())

        for key, value in self.images.items() :
            image = open(key, 'rb')
            docxFile.writestr('word/media/' + value, image.read())
            image.close()

        docxFile.close()

    #search and replace function
    def searchAndReplace(self, regex, replacement) :
        for key, value in self.files.items() :
            if key != 'word/_rels/document.xml.rels' and key != '[Content_Types].xml' and key != 'word/settings.xml': 
                value.searchAndReplace(regex, replacement)

    #copying file from old zip to new zip
    def copyToXML(self, docx, path) :
        docx.writestr(path, self._doc.read(path))
        return docx
