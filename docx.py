#!/usr/bin/python

import zipfile
from lxml import etree

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
            if path == 'word/document.xml' or 'header' in path or 'footer' in path :
                self.files[path] = etree.fromstring(self._doc.read(path))
        
    #read document header as xml and returning text as list
    def readHeader(self) :
        return self._readTextFromXML(self._header)

    #read document body as xml and returning text as list
    def readDocument(self) :
        return self._readTextFromXML(self._body)

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
                paraList.append(el)
                for e in el.iter() :
                    if e.tag == WPREFIXES['w'] + 'p' :
                        returnList.append(e.text)

        return returnList