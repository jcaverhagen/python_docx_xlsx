#!/usr/bin/python
import zipfile
from lxml import etree
from items.paragraph import Paragraph
from items.hyperlink import Hyperlink
from items.table import Table
from items.list import List
from items.image import Image
from items.document import (
 StyleFile, AppFile, RelationshipFile, DocumentRelationshipFile, CoreFile, DocumentFile, 
 ContentTypeFile, SettingsFile, FontTableFile, WebSettingsFile, ThemeFile
)
from PIL import Image as PILImage

WPREFIXES = {
        'w' : '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    }

class Document :
    
    _doc = ''
    files = {}
    images = {}

    def __init__(self, filename=None) :
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
        else :
            self.files['word/document.xml'] = DocumentFile()
            self.files['word/_rels/document.xml.rels'] = DocumentRelationshipFile()
    #read document header as xml and returning text as list
    def readHeader(self) :
        return self._readTextFromXML(self._header)

    #read document body as xml and returning text as list
    def readDocument(self) :
        return self._readTextFromXML(self._body)

    #add paragraph as first
    def addParagraph(self, text, position='last', style='NormalWeb', bold=False, italic=False, 
                        underline=False, uppercase=False, color=False, font=False) :
        doc = self.files['word/document.xml']
        paragraph = Paragraph(text, style, bold, italic, underline, uppercase, color, font).get()
        doc.addElement(paragraph, position)

    #add hyperlink to document
    def addHyperlink(self, text, url, position='last') :
        doc = self.files['word/document.xml']
        rel_id = self.files['word/_rels/document.xml.rels'].addRelation('hyperlink', url)

        paraElement = Paragraph().get()
        hyperlink = Hyperlink(text, str(rel_id), url)
        paraElement.append(hyperlink.get())

        doc.addElement(paraElement, position)

    #method to make specific test an hyperlink
    def makeTextHyperlink(self, text, url) :
        doc = self.files['word/document.xml']
        rel_id = self.files['word/_rels/document.xml.rels'].addRelation('hyperlink', url=url)

        hyperlink = Hyperlink(text, str(rel_id), url)
        doc.makeTextHyperlink(text, hyperlink.get())

    #init an new table and returning it to caller
    def addTable(self, width, columns, position='last') :
        return Table(width, columns, position=position)

    #close table and add it to document
    def closeTable(self, table) :
        doc = self.files['word/document.xml']
        doc.addElement(table.get(), table.getPosition())

    #init an new list and return object to caller
    def addList(self, position, type) :
        return List(position, type)
        
    #close list and inserting it in document
    def closeList(self, listItem) :
        doc = self.files['word/document.xml']
        listItems = listItem.get()
        
        if listItem.get() == 'last' :
             for item in listItems :
                doc.addElement(item, listItem.get())
        else :
            for item in reversed(listItems) :
                doc.addElement(item, listItem.get())

    def addImage(self, image, position='last', width='100%', height='100%') :
        doc = self.files['word/document.xml']

        count = 1
        if self.filename is not None :
            for path in self._doc.namelist() :
                if 'media' in path :
                    count = count + 1

            imagename = 'image' + str(count) + '.jpg'
        else :
            imagename = 'image' + str(len(self.images) + 1) + '.jpeg'

        self.images[image] = imagename

        rel_id = self.files['word/_rels/document.xml.rels'].addRelation('image', imagename=imagename)

        image = Image(image, rel_id, width, height)
        doc.addElement(image.get(), position)

    #save document with new values
    def save(self, filename) :
        docxFile = zipfile.ZipFile(filename, mode='w', compression=zipfile.ZIP_DEFLATED)
        if not self.filename :
            docxFile.writestr(RelationshipFile().path, RelationshipFile().getXml())
            docxFile.writestr(AppFile().path, AppFile().getXml())
            docxFile.writestr(CoreFile().path, CoreFile().getXml())
            docxFile.writestr(ContentTypeFile().path, ContentTypeFile().getXml())
            docxFile.writestr(SettingsFile().path, SettingsFile().getXml())
            docxFile.writestr(FontTableFile().path, FontTableFile().getXml())
            docxFile.writestr(WebSettingsFile().path, WebSettingsFile().getXml())
            docxFile.writestr(StyleFile().path, StyleFile().getXml())
            docxFile.writestr(ThemeFile().path, ThemeFile().getXml())
            
        else :
            
            #copy from old docx every file except the files that are in files list
            for path in self._doc.namelist() :
                if path not in self.files :
                    if path == 'word/styles.xml' :
                        styleFile = etree.fromstring(StyleFile().getXml())
                        docxFile.writestr('word/styles.xml', etree.tostring(styleFile, pretty_print=True))
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
            if key != 'word/_rels/document.xml.rels' :
                value.searchAndReplace(regex, replacement)

    #copying file from old zip to new zip
    def copyToXML(self, docx, path) :
        docx.writestr(path, self._doc.read(path))
        return docx

    #search for text tag in xml files
    def _readTextFromXML(self, xml) :
        returnList = []
        for el in xml.iter() :
            if el.tag == WPREFIXES['w'] + 'p' :
                for e in el.iter() :
                    if e.tag == WPREFIXES['w'] + 't' :
                        returnList.append(e.text)

        return returnList