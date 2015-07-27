import zipfile
from lxml import etree

WPREFIXES = {
        'w' : '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    }

class Document :
    
    _doc = ''
    _header = ''
    _body = ''
    _footer = ''

    def __init__(self, filename) :
        if not zipfile.is_zipfile(filename) :
            raise TypeError("Not an correct docx file.")

        self._doc = zipfile.ZipFile(filename)
        #self._header = etree.fromstring(self._doc.read('word/header1.xml'))
        self._body = etree.fromstring(self._doc.read('word/document.xml'))
        #self._footer = etree.fromstring(self._doc.read('word/footer2.xml'))
        
    #read document header as xml and returning text as list
    def readHeader(self) :
        return self._readTextFromXML(self._header)

    #read document body as xml and returning text as list
    def readDocument(self) :
        return self._readTextFromXML(self._body)

    def save(self, filename) :
        docxFile = zipfile.ZipFile(filename, mode='w', compression=zipfile.ZIP_DEFLATED)

        #headerString = etree.tostring(self._header, pretty_print=True)
        #docxFile.writestr('word/header1.xml', headerString)

        bodyString = etree.tostring(self._body, pretty_print=True)
        docxFile.writestr('word/document.xml', bodyString)

        for path in self._doc.namelist() :
            if path != 'word/document.xml' :
                docxFile = self.copyToXML(docxFile, path)

        docxFile.close()

    #search and replace function
    def searchAndReplace(self, regex, replacement) :
        
        for el in self._body.iter() :
            if el.tag == WPREFIXES['w'] + 'p' :
                for e in el.iter() :
                    if e.tag == WPREFIXES['w'] + 't' :
                        e.text = e.text.replace(regex, replacement)

    #copy function of docx
    def copyFile(self, filename) :

        newFile = zipfile.ZipFile(filename, mode="w", compression=zipfile.ZIP_DEFLATED)
        for path in self._doc.namelist() :
            if path != 'word/document.xml' :
                newFile = self.copyToXML(newFile, path)

        bodyString = etree.tostring(self._body, pretty_print=True)
        newFile.writestr('word/document.xml', bodyString)

        newFile.close()

    #copying file from old zip to new zip
    def copyToXML(self, docx, path) :
        docx.writestr(path, self._doc.read(path))
        return docx

    #search for text tag in xml files
    def _readTextFromXML(self, xml) :
        paraList = []
        for el in xml.iter() :
            if el.tag == WPREFIXES['w'] + 'p' :
                paraList.append(el)

        returnList = []
        for para in paraList :
            for p in para.iter() :
                if p.tag == WPREFIXES['w'] + 't' :
                    returnList.append(p.text)

        return returnList