import zipfile
from lxml import etree

WPREFIXES = {
        'w' : '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    }

class Document :
    
    def __init__(self) :
        print "constructor"

    #opening docx file and returning zip archive
    def openFile(self, fileName) :
        if not zipfile.is_zipfile(fileName) :
            raise TypeError("Not an correct docx file.")
        
        return zipfile.ZipFile(fileName) 

    #read header file as xml and returning text from it
    def readHeader(self, document) :
        xml = etree.fromstring(document.read('word/header2.xml'))
        
        return self._readTextFromXML(xml)

    #read document file as xml and returning text from it
    def readDocument(self, document) :
        xml = etree.fromstring(document.read('word/document.xml'))

        return self._readTextFromXML(xml)

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
