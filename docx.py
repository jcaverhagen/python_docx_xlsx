#!/usr/bin/python

import zipfile
from lxml import etree
from items.paragraph import Paragraph
from items.hyperlink import Hyperlink

WPREFIXES = {
        'w' : '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    }

class Document :
    
    _doc = ''
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
    def addParagraph(self, text, position='last', beforetext=None, aftertext=None) :
        doc = self.files['word/document.xml']
        for el in doc.iter() :
            if el.tag == WPREFIXES['w'] + 'body' :
                paragraph = Paragraph()
                paragraph.setText(text)
                paraElement = paragraph.get()
                
                if aftertext or beforetext :
                    position = self._searchParagraphPosition(aftertext)
                    if beforetext :
                        el.insert(position, paraElement)
                    elif aftertext :
                        el.insert(position + 1, paraElement)
                else :
                    if position == 'first' : el.insert(0, paraElement)
                    else : el.append(paraElement)

    #search position of paragraph
    def _searchParagraphPosition(self, text) :
        position = 0
        for el in self.files['word/document.xml'].iter() :
            if el.tag == WPREFIXES['w'] + 'p' :
                for e in el.iter() :
                    if e.tag == WPREFIXES['w'] + 't' :
                        position = position + 1
                        if text in e.text :
                            return position
        return position

    #add hyperlink to document
    def addHyperlink(self, text, url) :

        newRelationID = self._getHighestRelationId() + 1
        relations = self.files['word/_rels/document.xml.rels']

        doc = self.files['word/document.xml']
        for el in doc.iter() :
            if el.tag == WPREFIXES['w'] + 'body' :
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