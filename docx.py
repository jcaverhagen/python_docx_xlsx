#!/usr/bin/python

import zipfile
from lxml import etree
from items.paragraph import Paragraph
from items.hyperlink import Hyperlink
from items.table import Table
from items.list import List

WPREFIXES = {
        'w' : '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    }

class Document :
    
    _doc = ''
    _table = ''
    files = {}

    def __init__(self, filename) :
        if not zipfile.is_zipfile(filename) :
            raise TypeError("Not an correct docx file.")

        self._doc = zipfile.ZipFile(filename)
        #scan zipfile for header/document and footer xml files
        for path in self._doc.namelist() :
            if path == 'word/document.xml' or path == 'word/_rels/document.xml.rels' or 'header' in path or 'footer' in path :
                self.files[path] = etree.fromstring(self._doc.read(path))
        
    #read document header as xml and returning text as list
    def readHeader(self) :
        return self._readTextFromXML(self._header)

    #read document body as xml and returning text as list
    def readDocument(self) :
        return self._readTextFromXML(self._body)

    #add paragraph as first
    def addParagraph(self, text, position='last') :
        doc = self.files['word/document.xml']
        for el in doc.iter() :
            if el.tag == WPREFIXES['w'] + 'body' :
                paragraph = Paragraph()
                paragraph.setText(text)
                paraElement = paragraph.get()
                
                self._addToDoc(paraElement, position)

    #add hyperlink to document
    def addHyperlink(self, text, url, position='last') :

        newRelationID = self._getHighestRelationId() + 1
        relations = self.files['word/_rels/document.xml.rels']

        doc = self.files['word/document.xml']
        for el in doc.iter() :
            if el.tag == WPREFIXES['w'] + 'body' :
                paragraph = Paragraph().get()
                hyperlink = Hyperlink(text, str(newRelationID), url)
                relations.append(hyperlink.getRelation())
                paragraph.append(hyperlink.get())

                self._addToDoc(paragraph, position)

    #method to make specific test an hyperlink
    def makeTextHyperlink(self, text, url) :
        for el in self.files['word/document.xml'].iter() :
            if el.tag == WPREFIXES['w'] + 'p' :
                for e in el.iter() :
                    addLink = False
                    if e.tag == WPREFIXES['w'] + 't' :
                        if e.text :
                            if text in e.text :
                                e.text = e.text.replace(text, '')
                                addLink = True
                    if addLink :
                        newRelationID = self._getHighestRelationId() + 1
                        relations = self.files['word/_rels/document.xml.rels']
                        
                        hyperlink = Hyperlink(text, str(newRelationID), url)
                        relations.append(hyperlink.getRelation())
                        el.append(hyperlink.get())

    #search for highest id in relations xml
    def _getHighestRelationId(self) :
        highest = 0;
        relations = self.files['word/_rels/document.xml.rels']
        for rel in relations :
            if int(rel.attrib['Id'].replace('rId', '')) > highest :
                highest = int(rel.attrib['Id'].replace('rId', ''))
        return highest

    def addTable(self, width, columns) :
        self._table = Table(width, columns)

    def addRow(self, val) :
        self._table.addRow(val)

    def addTableToDoc(self, position='last') :
        self._addToDoc(self._table.get(), position)

    #init an new list and return object to caller
    def addList(self, position) :
        listItem = List(position)
        return listItem

    #close list and inserting it in document
    def closeList(self, listItem) :
        position = listItem.getPosition()
        listItems = listItem.get()
        
        if position == 'last' :
             for item in listItems :
                self._addToDoc(item, position)
        else :
            for item in reversed(listItems) :
                self._addToDoc(item, position)

    #method to add element to document file
    def _addToDoc(self, element, position='last') :
        doc = self.files['word/document.xml']
        for el in doc.iter() :
            if el.tag == WPREFIXES['w'] + 'body' :
                if position == 'first' : el.insert(0, element)
                else : 
                    if 'aftertext:' in position :
                        position = self._searchParagraphPosition(position.replace('aftertext:', ''))
                        el.insert(position, element)    
                    elif 'beforetext:' in position :
                        position = self._searchParagraphPosition(position.replace('beforetext:', ''))
                        el.insert(position-1, element)
                    else :
                        el.append(element)
                    

    #search position of paragraph
    def _searchParagraphPosition(self, text) :
        position = 0
        for el in self.files['word/document.xml'].iter() :
            if el.tag == WPREFIXES['w'] + 'p' :
                for e in el.iter() :
                    if e.tag == WPREFIXES['w'] + 't' :
                        position = position + 1
                        if e.text :
                            if text in e.text :
                                return position
        return position

    #save document with new values
    def save(self, filename) :
        docxFile = zipfile.ZipFile(filename, mode='w', compression=zipfile.ZIP_DEFLATED)

        #copy from old docx every file except the files that are in files list
        for path in self._doc.namelist() :
            if path not in self.files :
                docxFile = self.copyToXML(docxFile, path)

        #add files from file list to docx
        for key, value in self.files.items() :
            docxFile.writestr(key, etree.tostring(value, pretty_print=True))

        docxFile.close()

    #search and replace function
    def searchAndReplace(self, regex, replacement) :
        for key, value in self.files.items() :
            if key != 'word/_rels/document.xml.rels' :
                for el in value.iter() :
                    if el.tag == WPREFIXES['w'] + 'p' :
                        for e in el.iter() :
                            if e.tag == WPREFIXES['w'] + 't' :
                                e.text = e.text.replace(regex, replacement)

    #copy function of docx
    def copyFile(self, filename) :

        newFile = zipfile.ZipFile(filename, mode="w", compression=zipfile.ZIP_DEFLATED)
        for path in self._doc.namelist() :
            newFile = self.copyToXML(newFile, path)

        newFile.close()

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