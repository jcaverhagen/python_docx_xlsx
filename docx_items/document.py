from datetime import datetime
from universal.element import Element
from lxml import etree
from universal import defaults

class RelationshipFile() :

	path = '_rels/.rels'

	def __init__(self) :
		self.xmlString = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
			<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
			<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
			<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
			<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
			</Relationships>"""

	def getXml(self) :
		return self.xmlString

class DocumentRelationshipFile() :

	path = 'word/_rels/document.xml.rels'
	_rels = ''

	def __init__(self, xml=None) :
		if xml is not None :
			self._rels = xml
		else :
			self._rels = etree.fromstring("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
				<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
				<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" Target="webSettings.xml"/>
				<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
				<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
				<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
				<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
				<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
				</Relationships>""")

	def addRelation(self, type, url=None, imagename='', headerfootertype='default') :
		new_id = self._getHighestRelationId() + 1
		if type == 'hyperlink' :
			attr = {'Id' : 'rId' + str(new_id), 'Type' : defaults.WPREFIXES['r'] + '/hyperlink', 'Target' : url, 'TargetMode' : 'External'}
			rel = Element().createElement('Relationship', prefix=None, attr=attr)
			self._rels.append(rel)
			
		if type == 'image' :
			attr = {'Id' : 'rId' + str(new_id), 'Type' : defaults.WPREFIXES['r'] + '/image', 'Target' : 'media/' + imagename}
			rel = Element().createElement('Relationship', prefix=None, attr=attr)
			self._rels.append(rel)
		
		if type == 'header' :
			if headerfootertype == 'even' :
				file = 'header1.xml'
			elif headerfootertype == 'first' :
				file = 'header3.xml'
			else :
				file = 'header2.xml'

			attr = {'Id' : 'rId' + str(new_id), 'Type' : defaults.WPREFIXES['r'] + '/header', 'Target' : file}
			rel = Element().createElement('Relationship', prefix=None, attr=attr)
			self._rels.append(rel)

		if type == 'footer' :
			if headerfootertype == 'even' :
				file = 'footer1.xml'
			elif headerfootertype == 'first' :
				file = 'footer3.xml'
			else :
				file = 'footer2.xml'

			attr = {'Id' : 'rId' + str(new_id), 'Type' : defaults.WPREFIXES['r'] + '/footer', 'Target' : file}
			rel = Element().createElement('Relationship', prefix=None, attr=attr)
			self._rels.append(rel)

		return new_id

	#search for highest id in relations xml
	def _getHighestRelationId(self) :
		highest = 0
		for rel in self._rels :
			if int(rel.attrib['Id'].replace('rId', '')) > highest :
				highest = int(rel.attrib['Id'].replace('rId', ''))

		return highest

	def getXml(self) :
		return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + etree.tostring(self._rels, pretty_print=True)

class AppFile :

	path = 'docProps/app.xml'
	props = {}

	def __init__(self) :

		self.props['company'] = ''

		self.xmlString = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
			<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
			<Template>Normal.dotm</Template>
			<TotalTime>0</TotalTime>
			<Pages>1</Pages>
			<Words>0</Words>
			<Characters>0</Characters>
			<Application>Microsoft Office Word</Application>
			<DocSecurity>0</DocSecurity>
			<Lines>1</Lines>
			<Paragraphs>1</Paragraphs>
			<ScaleCrop>false</ScaleCrop>
			<Company>{company}</Company>
			<LinksUpToDate>false</LinksUpToDate>
			<CharactersWithSpaces>0</CharactersWithSpaces>
			<SharedDoc>false</SharedDoc>
			<HyperlinksChanged>false</HyperlinksChanged>
			<AppVersion>12.0000</AppVersion>
			</Properties>"""

	def getXml(self) :
		return self.xmlString.format(company=self.props.get('company', ''))

class CoreFile :

	path = 'docProps/core.xml'
	props = {}

	def __init__(self) :
		self.props['creator'] = ''
		self.xmlString = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
			<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
			<dc:creator>{creator}</dc:creator>
			<cp:lastModifiedBy>{creator}</cp:lastModifiedBy>
			<cp:revision>1</cp:revision>
			<dcterms:created xsi:type="dcterms:W3CDTF">{created}</dcterms:created>
			<dcterms:modified xsi:type="dcterms:W3CDTF">{created}</dcterms:modified>
			</cp:coreProperties>"""

	def getXml(self) :
		created = datetime.strftime(datetime.today(), '%Y-%m-%dT%H:%M:%SZ')
		return self.xmlString.format(creator=self.props.get('creator', ''),
									created=created)

class DocumentFile :

	path = 'word/document.xml'
	_doc = ''

	def __init__(self, xml=None) :
		if xml is not None :
			self._doc = xml
		else :
			self._doc = etree.fromstring("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
				<w:document xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml">
				<w:body>
					<w:sectPr>
						<w:pgSz w:w="11906" w:h="16838"/>
						<w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/>
						<w:cols w:space="708"/>
						<w:docGrid w:linePitch="360"/>
					</w:sectPr>
				</w:body>
				</w:document>""")

	#adding an element to document.xml
	def addElement(self, element, position='last') :
		for el in self._doc.iter() :
			if el.tag == '{' + defaults.WPREFIXES['w'] + '}body' :
				if position == 'first' : el.insert(0, element)
				else :
					if 'beforetext:' in position :
						position = self._searchParagraphPosition(position.replace('aftertext:', ''))
						el.insert(position, element)
					elif 'aftertext:' in position :
						position = self._searchParagraphPosition(position.replace('aftertext:', ''))
						el.insert(position + 1, element)
					else :
						el.append(element)

	def makeTextHyperlink(self, text, element) :
		addLink = False
		for el in self._doc.iter() :
			if el.tag == '{' + defaults.WPREFIXES['w'] + '}'  + 'p' :
				for e in el.iter() :
					if e.tag == '{' + defaults.WPREFIXES['w'] + '}'  + 't' :
						if e.text :
							if text in e.text :
								splittedText = e.text.split(text)
								e.set('{' + defaults.WPREFIXES['ns'] + '}space', 'preserve')
								e.text = splittedText[0]
								addLink = True

				if addLink :
					el.append(element)
					
					run = Element().createElement('r')
					textEl = Element().createElement('t', attr={'space' : 'preserve'}, attrprefix='ns')
					textEl.text = splittedText[1]
					run.append(textEl)

					el.append(run)
					break

	def addAnchorToText(self, text, anchor) :
		addBookmark = False
		for el in self._doc.iter() :
			if el.tag == '{' + defaults.WPREFIXES['w'] + '}'  + 'p' :
				for e in el.iter() :
					if e.tag == '{' + defaults.WPREFIXES['w'] + '}'  + 't' :
						if e.text :
							if text in e.text :
								splittedText = e.text.split(text)
								e.set('{' + defaults.WPREFIXES['ns'] + '}space', 'preserve')
								e.text = splittedText[0]
								addBookmark = True

				if addBookmark :
					#create bookmark elements
					bookmarkStart = Element().createElement('bookmarkStart', attr={'id' : '0', 'name' : anchor})
					bookmarkEnd = Element().createElement('bookmarkEnd', attr={'id' : '0'})

					#add start of bookmark
					el.append(bookmarkStart)

					#create run with bookmark text
					run = Element().createElement('r')
					textEl = Element().createElement('t', attr={'space' : 'preserve'}, attrprefix='ns')
					textEl.text = text
					run.append(textEl)
					el.append(run)

					#add end of bookmark
					el.append(bookmarkEnd)

					#create run with bookmark text
					run = Element().createElement('r')
					textEl = Element().createElement('t', attr={'space' : 'preserve'}, attrprefix='ns')
					textEl.text = splittedText[1]
					run.append(textEl)
					el.append(run)
					break

	def searchAndReplace(self, regex, replacement) :
		for el in self._doc.iter() :
			if el.tag == '{' + defaults.WPREFIXES['w'] + '}'  + 'p' :
				for e in el.iter() :
					if e.tag == '{' + defaults.WPREFIXES['w'] + '}'  + 't' :
						e.text = e.text.replace(regex, replacement)

	#search position of paragraph
	def _searchParagraphPosition(self, text):
		position = 0
		for el in self._doc.iter() :
			if el.tag == '{' + defaults.WPREFIXES['w'] + '}'  + 'p' :
				for e in el.iter() :
					if e.tag == '{' + defaults.WPREFIXES['w'] + '}'  + 't' :
						position = position + 1
						if e.text :
							if text in e.text :
								return position
		return position

	def addReference(self, filetype, type, id) :
		if filetype == 'header' :
			Reference = Element().createElement('headerReference', attr={'type' : type, 'rel_id' : 'rId' + str(id)})
		else : 
			Reference = Element().createElement('footerReference', attr={'type' : type, 'rel_id' : 'rId' + str(id)})

		added = False
		for el in self._doc.iter() :
			if el.tag == '{' + defaults.WPREFIXES['w'] + '}'  + 'body' :
				for l in el.iter() :
					if l.tag == '{' + defaults.WPREFIXES['w'] + '}'  + 'sectPr' :
						if added == False :
							l.append(Reference)
		
							if type == 'first' :
								l.append(Element().createElement('titlePg'))
							added = True


	def getXml(self) :
		return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + etree.tostring(self._doc, pretty_print=True)

class ContentTypeFile() :

	path = '[Content_Types].xml'
	imageExtensions = {
		'jpg' : Element().createElement('Default', attr={'Extension' : 'jpg', 'ContentType' : 'image/jpeg'}, prefix=None, attrprefix=None),
		'jpeg' : Element().createElement('Default', attr={'Extension' : 'jpeg', 'ContentType' : 'image/jpeg'}, prefix=None, attrprefix=None),
		'gif' : Element().createElement('Default', attr={'Extension' : 'gif', 'ContentType' : 'image/gif'}, prefix=None, attrprefix=None),
		'png' : Element().createElement('Default', attr={'Extension' : 'png', 'ContentType' : 'image/png'}, prefix=None, attrprefix=None)
	}

	def __init__(self, xml=None) :
		if xml is not None :
			self.xmlString = xml
		else :
			self.xmlString = etree.fromstring("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
				<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
				<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
				<Default Extension="xml" ContentType="application/xml"/>
				<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
				<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
				<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
				<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
				<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
				<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
				<Override PartName="/word/webSettings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"/>
				<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
				<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
				</Types>""")

	def addImageExtension(self, extension) :
		alreadyExists = False
		#check if extension already in Content_Types
		for type in self.xmlString.iter() :
			if type.tag == '{' + defaults.WPREFIXES['ct'] + '}' + 'Default' :
				if type.attrib['Extension'] == extension :
					alreadyExists = True

		if alreadyExists == False :
			self.xmlString.append(self.imageExtensions[extension])

	def addOverride(self, type, filenumber) :
		if type == 'header' :
			PartName = '/word/header' + str(filenumber) + '.xml'
			ContentType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'

			notInXml = True
			for over in self.xmlString.iter() :
				if over.tag == '{' + defaults.WPREFIXES['ct'] + '}' + 'Override' :
					if over.attrib['PartName'] == PartName :
						notInXml = False

			if notInXml :
				self.xmlString.append(Element().createElement('Override', attr={'PartName' : PartName, 'ContentType' : ContentType}, prefix=None, attrprefix=None))
		elif type == 'footer' :
			PartName = '/word/footer' + str(filenumber) + '.xml'
			ContentType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml'

			notInXml = True
			for over in self.xmlString.iter() :
				if over.tag == '{' + defaults.WPREFIXES['ct'] + '}' + 'Override' :
					if over.attrib['PartName'] == PartName :
						notInXml = False

			if notInXml :
				self.xmlString.append(Element().createElement('Override', attr={'PartName' : PartName, 'ContentType' : ContentType}, prefix=None, attrprefix=None))
		
	def getXml(self) :
		return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + etree.tostring(self.xmlString, pretty_print=True)

class SettingsFile() :

	path = 'word/settings.xml'
	enableEvenAndOddHeaders = False
	
	def __init__(self, xml=None) :
		if xml is not None :
			self.settings = xml
		else :
			self.settings = etree.fromstring("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
				<w:settings xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:sl="http://schemas.openxmlformats.org/schemaLibrary/2006/main">
				<w:zoom w:percent="100"/>
				<w:proofState w:spelling="clean" w:grammar="clean" />
				<w:defaultTabStop w:val="708"/>
				<w:hyphenationZone w:val="425"/>
				<w:characterSpacingControl w:val="doNotCompress"/>
				<w:compat/>
				<w:rsids>
				<w:rsidRoot w:val="00C113BC"/>
				<w:rsid w:val="00C113BC"/>
				<w:rsid w:val="00FB4BB8"/>
				</w:rsids>
				<m:mathPr>
				<m:mathFont m:val="Cambria Math"/>
				<m:brkBin m:val="before"/>
				<m:brkBinSub m:val="--"/>
				<m:smallFrac m:val="off"/>
				<m:dispDef/>
				<m:lMargin m:val="0"/>
				<m:rMargin m:val="0"/>
				<m:defJc m:val="centerGroup"/>
				<m:wrapIndent m:val="1440"/>
				<m:intLim m:val="subSup"/>
				<m:naryLim m:val="undOvr"/>
				</m:mathPr>
				<w:themeFontLang w:val="nl-NL"/>
				<w:clrSchemeMapping w:bg1="light1" w:t1="dark1" w:bg2="light2" w:t2="dark2" w:accent1="accent1" w:accent2="accent2" w:accent3="accent3" w:accent4="accent4" w:accent5="accent5" w:accent6="accent6" w:hyperlink="hyperlink" w:followedHyperlink="followedHyperlink"/>
				<w:shapeDefaults>
				<o:shapedefaults v:ext="edit" spidmax="2050"/>
				<o:shapelayout v:ext="edit">
				<o:idmap v:ext="edit" data="1"/>
				</o:shapelayout>
				</w:shapeDefaults>
				<w:decimalSymbol w:val=","/>
				<w:listSeparator w:val=";"/>
				</w:settings>""")

	def enableEvenAndOddHeaders(self) :
		if self.enableEvenAndOddHeaders == False :
			self.settings.append(Element().createElement('evenAndOddHeaders'))
			self.enableEvenAndOddHeaders = True

	def getXml(self) :
		return etree.tostring(self.settings, pretty_print=True)

class HeaderFile :

	path = 'word/header{filenumber}.xml'
	
	def __init__(self, text, filenumber, xml=None) :

		self.path = self.path.format(filenumber=filenumber)
		self.text = text

		self.xmlString = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
			<w:hdr xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml">
			<w:p>
			<w:pPr>
			<w:pStyle w:val="Header"/>
			</w:pPr>
			<w:r>
			<w:t>{text}</w:t>
			</w:r>
			</w:p>
			</w:hdr>"""

	def searchAndReplace(self, regex, replacement) :
		self.text = self.text.replace(regex, replacement)

	def getXml(self) :
		return self.xmlString.format(text=self.text)

class FooterFile :

	path = 'word/footer{filenumber}.xml'

	def __init__(self, path, xml=None) :

		self.path = path

		if xml is not None :
			self.xmlString = xml
		else :
			self.xmlString = etree.fromstring("""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
				<w:ftr xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml">
				</w:ftr>""")

	def addText(self, text) :
		p = Element().createElement('p')
		pPr = Element().createElement('pPr')
		pStyle = Element().createElement('pStyle', attr={'val' : 'Footer'})
		pPr.append(pStyle)
		r = Element().createElement('r')
		t = Element().createElement('t')
		t.text = text
		r.append(t)

		p.append(pPr)
		p.append(r)
		
		self.xmlString.append(p)
		
	def searchAndReplace(self, regex, replacement) :
		for elem in self.xmlString.iter() :
			paragraph = ''
			if elem.tag == '{' + defaults.WPREFIXES['w'] + '}p' :
				for el in elem.iter() :
					if el.tag == '{' + defaults.WPREFIXES['w'] + '}t' :
						paragraph = paragraph + el.text
				if paragraph != '' :
					self.xmlString.remove(elem)
					self.addText(paragraph.replace(regex, replacement))

	def getXml(self) :
		return etree.tostring(self.xmlString, pretty_print=True)

class FontTableFile() :

	path = 'word/fontTable.xml'

	def __init__(self) :
		self.xmlString = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
			<w:fonts xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
			<w:font w:name="Calibri">
			<w:panose1 w:val="020F0502020204030204"/>
			<w:charset w:val="00"/>
			<w:family w:val="swiss"/>
			<w:pitch w:val="variable"/>
			<w:sig w:usb0="E00002FF" w:usb1="4000ACFF" w:usb2="00000001" w:usb3="00000000" w:csb0="0000019F" w:csb1="00000000"/>
			</w:font>
			<w:font w:name="Times New Roman">
			<w:panose1 w:val="02020603050405020304"/>
			<w:charset w:val="00"/>
			<w:family w:val="roman"/>
			<w:pitch w:val="variable"/>
			<w:sig w:usb0="E0002AFF" w:usb1="C0007841" w:usb2="00000009" w:usb3="00000000" w:csb0="000001FF" w:csb1="00000000"/>
			</w:font>
			<w:font w:name="Cambria">
			<w:panose1 w:val="02040503050406030204"/>
			<w:charset w:val="00"/>
			<w:family w:val="roman"/>
			<w:pitch w:val="variable"/>
			<w:sig w:usb0="E00002FF" w:usb1="400004FF" w:usb2="00000000" w:usb3="00000000" w:csb0="0000019F" w:csb1="00000000"/>
			</w:font>
			<w:font w:name="Tahoma">
			<w:panose1 w:val="020B0604030504040204"/>
			<w:charset w:val="00"/>
			<w:family w:val="swiss"/>
			<w:pitch w:val="variable"/>
			<w:sig w:usb0="E1002EFF" w:usb1="C000605B" w:usb2="00000029" w:usb3="00000000" w:csb0="000101FF" w:csb1="00000000"/>
			</w:font>
			</w:fonts>"""

	def getXml(self) :
		return self.xmlString

class NumberingFile() :

	path = 'word/numbering.xml'

	def __init__(self) :
		self.xmlString = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
			<w:numbering xmlns:ve="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml">
			<w:abstractNum w:abstractNumId="0">
			<w:nsid w:val="56BA774D"/>
			<w:multiLevelType w:val="hybridMultilevel"/>
			<w:tmpl w:val="388265EE"/>
			<w:lvl w:ilvl="0" w:tplc="0413000F">
			<w:start w:val="1"/>
			<w:numFmt w:val="decimal"/>
			<w:lvlText w:val="%1."/>
			<w:lvlJc w:val="left"/>
			<w:pPr>
			<w:ind w:left="720" w:hanging="360"/>
			</w:pPr>
			</w:lvl>
			<w:lvl w:ilvl="1" w:tplc="04130019">
			<w:start w:val="1"/>
			<w:numFmt w:val="lowerLetter"/>
			<w:lvlText w:val="%2."/>
			<w:lvlJc w:val="left"/>
			<w:pPr>
			<w:ind w:left="1440" w:hanging="360"/>
			</w:pPr>
			</w:lvl>
			<w:lvl w:ilvl="2" w:tplc="0413001B" w:tentative="1">
			<w:start w:val="1"/>
			<w:numFmt w:val="lowerRoman"/>
			<w:lvlText w:val="%3."/>
			<w:lvlJc w:val="right"/>
			<w:pPr>
			<w:ind w:left="2160" w:hanging="180"/>
			</w:pPr>
			</w:lvl>
			<w:lvl w:ilvl="3" w:tplc="0413000F" w:tentative="1">
			<w:start w:val="1"/>
			<w:numFmt w:val="decimal"/>
			<w:lvlText w:val="%4."/>
			<w:lvlJc w:val="left"/>
			<w:pPr>
			<w:ind w:left="2880" w:hanging="360"/>
			</w:pPr>
			</w:lvl>
			<w:lvl w:ilvl="4" w:tplc="04130019" w:tentative="1">
			<w:start w:val="1"/>
			<w:numFmt w:val="lowerLetter"/>
			<w:lvlText w:val="%5."/>
			<w:lvlJc w:val="left"/>
			<w:pPr>
			<w:ind w:left="3600" w:hanging="360"/>
			</w:pPr>
			</w:lvl>
			<w:lvl w:ilvl="5" w:tplc="0413001B" w:tentative="1">
			<w:start w:val="1"/>
			<w:numFmt w:val="lowerRoman"/>
			<w:lvlText w:val="%6."/>
			<w:lvlJc w:val="right"/>
			<w:pPr>
			<w:ind w:left="4320" w:hanging="180"/>
			</w:pPr>
			</w:lvl>
			<w:lvl w:ilvl="6" w:tplc="0413000F" w:tentative="1">
			<w:start w:val="1"/>
			<w:numFmt w:val="decimal"/>
			<w:lvlText w:val="%7."/>
			<w:lvlJc w:val="left"/>
			<w:pPr>
			<w:ind w:left="5040" w:hanging="360"/>
			</w:pPr>
			</w:lvl>
			<w:lvl w:ilvl="7" w:tplc="04130019" w:tentative="1">
			<w:start w:val="1"/>
			<w:numFmt w:val="lowerLetter"/>
			<w:lvlText w:val="%8."/>
			<w:lvlJc w:val="left"/>
			<w:pPr>
			<w:ind w:left="5760" w:hanging="360"/>
			</w:pPr>
			</w:lvl>
			<w:lvl w:ilvl="8" w:tplc="0413001B" w:tentative="1">
			<w:start w:val="1"/>
			<w:numFmt w:val="lowerRoman"/>
			<w:lvlText w:val="%9."/>
			<w:lvlJc w:val="right"/>
			<w:pPr>
			<w:ind w:left="6480" w:hanging="180"/>
			</w:pPr>
			</w:lvl>
			</w:abstractNum>
			<w:num w:numId="1">
			<w:abstractNumId w:val="0"/>
			</w:num>
			</w:numbering>"""

	def getXml(self) :
		return self.xmlString

class WebSettingsFile() :

	path = 'word/webSettings.xml'

	def __init__(self) :
		self.xmlString = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
			<w:webSettings xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
			<w:optimizeForBrowser/>
			</w:webSettings>"""

	def getXml(self) :
		return self.xmlString

class StylesWithEffectsFile() :

	path = 'word/stylesWithEffects.xml'

	def __init__(self) :
		self.xmlString = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
			<w:styles xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 wp14">
			<w:docDefaults>
			<w:rPrDefault>
			<w:rPr>
			<w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>
			<w:sz w:val="22"/>
			<w:szCs w:val="22"/>
			<w:lang w:val="nl-NL" w:eastAsia="en-US" w:bidi="ar-SA"/>
			</w:rPr>
			</w:rPrDefault>
			<w:pPrDefault>
			<w:pPr>
			<w:spacing w:after="200" w:line="276" w:lineRule="auto"/>
			</w:pPr>
			</w:pPrDefault>
			</w:docDefaults>
			<w:latentStyles w:defLockedState="0" w:defUIPriority="99" w:defSemiHidden="1" w:defUnhideWhenUsed="1" w:defQFormat="0" w:count="267">
			<w:lsdException w:name="Normal" w:semiHidden="0" w:uiPriority="0" w:unhideWhenUsed="0" w:qFormat="1"/>
			<w:lsdException w:name="heading 1" w:semiHidden="0" w:uiPriority="9" w:unhideWhenUsed="0" w:qFormat="1"/>
			<w:lsdException w:name="heading 2" w:uiPriority="9" w:qFormat="1"/>
			<w:lsdException w:name="heading 3" w:uiPriority="9" w:qFormat="1"/>
			<w:lsdException w:name="heading 4" w:uiPriority="9" w:qFormat="1"/>
			<w:lsdException w:name="heading 5" w:uiPriority="9" w:qFormat="1"/>
			<w:lsdException w:name="heading 6" w:uiPriority="9" w:qFormat="1"/>
			<w:lsdException w:name="heading 7" w:uiPriority="9" w:qFormat="1"/>
			<w:lsdException w:name="heading 8" w:uiPriority="9" w:qFormat="1"/>
			<w:lsdException w:name="heading 9" w:uiPriority="9" w:qFormat="1"/>
			<w:lsdException w:name="toc 1" w:uiPriority="39"/>
			<w:lsdException w:name="toc 2" w:uiPriority="39"/>
			<w:lsdException w:name="toc 3" w:uiPriority="39"/>
			<w:lsdException w:name="toc 4" w:uiPriority="39"/>
			<w:lsdException w:name="toc 5" w:uiPriority="39"/>
			<w:lsdException w:name="toc 6" w:uiPriority="39"/>
			<w:lsdException w:name="toc 7" w:uiPriority="39"/>
			<w:lsdException w:name="toc 8" w:uiPriority="39"/>
			<w:lsdException w:name="toc 9" w:uiPriority="39"/>
			<w:lsdException w:name="caption" w:uiPriority="35" w:qFormat="1"/>
			<w:lsdException w:name="Title" w:semiHidden="0" w:uiPriority="10" w:unhideWhenUsed="0" w:qFormat="1"/>
			<w:lsdException w:name="Default Paragraph Font" w:uiPriority="1"/>
			<w:lsdException w:name="Subtitle" w:semiHidden="0" w:uiPriority="11" w:unhideWhenUsed="0" w:qFormat="1"/>
			<w:lsdException w:name="Strong" w:semiHidden="0" w:uiPriority="22" w:unhideWhenUsed="0" w:qFormat="1"/>
			<w:lsdException w:name="Emphasis" w:semiHidden="0" w:uiPriority="20" w:unhideWhenUsed="0" w:qFormat="1"/>
			<w:lsdException w:name="Table Grid" w:semiHidden="0" w:uiPriority="59" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Placeholder Text" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="No Spacing" w:semiHidden="0" w:uiPriority="1" w:unhideWhenUsed="0" w:qFormat="1"/>
			<w:lsdException w:name="Light Shading" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Light List" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Light Grid" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Shading 1" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Shading 2" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium List 1" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium List 2" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Grid 1" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Grid 2" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Grid 3" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Dark List" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Colorful Shading" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Colorful List" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Colorful Grid" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Light Shading Accent 1" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Light List Accent 1" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Light Grid Accent 1" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Shading 1 Accent 1" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Shading 2 Accent 1" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium List 1 Accent 1" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Revision" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="List Paragraph" w:semiHidden="0" w:uiPriority="34" w:unhideWhenUsed="0" w:qFormat="1"/>
			<w:lsdException w:name="Quote" w:semiHidden="0" w:uiPriority="29" w:unhideWhenUsed="0" w:qFormat="1"/>
			<w:lsdException w:name="Intense Quote" w:semiHidden="0" w:uiPriority="30" w:unhideWhenUsed="0" w:qFormat="1"/>
			<w:lsdException w:name="Medium List 2 Accent 1" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Grid 1 Accent 1" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Grid 2 Accent 1" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Grid 3 Accent 1" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Dark List Accent 1" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Colorful Shading Accent 1" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Colorful List Accent 1" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Colorful Grid Accent 1" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Light Shading Accent 2" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Light List Accent 2" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Light Grid Accent 2" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Shading 1 Accent 2" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Shading 2 Accent 2" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium List 1 Accent 2" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium List 2 Accent 2" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Grid 1 Accent 2" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Grid 2 Accent 2" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Grid 3 Accent 2" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Dark List Accent 2" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Colorful Shading Accent 2" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Colorful List Accent 2" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Colorful Grid Accent 2" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Light Shading Accent 3" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Light List Accent 3" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Light Grid Accent 3" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Shading 1 Accent 3" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Shading 2 Accent 3" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium List 1 Accent 3" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium List 2 Accent 3" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Grid 1 Accent 3" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Grid 2 Accent 3" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Grid 3 Accent 3" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Dark List Accent 3" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Colorful Shading Accent 3" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Colorful List Accent 3" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Colorful Grid Accent 3" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Light Shading Accent 4" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Light List Accent 4" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Light Grid Accent 4" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Shading 1 Accent 4" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Shading 2 Accent 4" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium List 1 Accent 4" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium List 2 Accent 4" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Grid 1 Accent 4" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Grid 2 Accent 4" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Grid 3 Accent 4" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Dark List Accent 4" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Colorful Shading Accent 4" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Colorful List Accent 4" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Colorful Grid Accent 4" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Light Shading Accent 5" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Light List Accent 5" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Light Grid Accent 5" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Shading 1 Accent 5" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Shading 2 Accent 5" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium List 1 Accent 5" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium List 2 Accent 5" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Grid 1 Accent 5" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Grid 2 Accent 5" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Grid 3 Accent 5" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Dark List Accent 5" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Colorful Shading Accent 5" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Colorful List Accent 5" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Colorful Grid Accent 5" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Light Shading Accent 6" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Light List Accent 6" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Light Grid Accent 6" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Shading 1 Accent 6" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Shading 2 Accent 6" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium List 1 Accent 6" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium List 2 Accent 6" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Grid 1 Accent 6" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Grid 2 Accent 6" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Grid 3 Accent 6" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Dark List Accent 6" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Colorful Shading Accent 6" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Colorful List Accent 6" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Colorful Grid Accent 6" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Subtle Emphasis" w:semiHidden="0" w:uiPriority="19" w:unhideWhenUsed="0" w:qFormat="1"/>
			<w:lsdException w:name="Intense Emphasis" w:semiHidden="0" w:uiPriority="21" w:unhideWhenUsed="0" w:qFormat="1"/>
			<w:lsdException w:name="Subtle Reference" w:semiHidden="0" w:uiPriority="31" w:unhideWhenUsed="0" w:qFormat="1"/>
			<w:lsdException w:name="Intense Reference" w:semiHidden="0" w:uiPriority="32" w:unhideWhenUsed="0" w:qFormat="1"/>
			<w:lsdException w:name="Book Title" w:semiHidden="0" w:uiPriority="33" w:unhideWhenUsed="0" w:qFormat="1"/>
			<w:lsdException w:name="Bibliography" w:uiPriority="37"/>
			<w:lsdException w:name="TOC Heading" w:uiPriority="39" w:qFormat="1"/>
			</w:latentStyles>
			<w:style w:type="paragraph" w:default="1" w:styleId="Standaard">
			<w:name w:val="Normal"/>
			<w:qFormat/>
			</w:style>
			<w:style w:type="character" w:default="1" w:styleId="Standaardalinea-lettertype">
			<w:name w:val="Default Paragraph Font"/>
			<w:uiPriority w:val="1"/>
			<w:semiHidden/>
			<w:unhideWhenUsed/>
			</w:style>
			<w:style w:type="table" w:default="1" w:styleId="Standaardtabel">
			<w:name w:val="Normal Table"/>
			<w:uiPriority w:val="99"/>
			<w:semiHidden/>
			<w:unhideWhenUsed/>
			<w:tblPr>
			<w:tblInd w:w="0" w:type="dxa"/>
			<w:tblCellMar>
			<w:top w:w="0" w:type="dxa"/>
			<w:left w:w="108" w:type="dxa"/>
			<w:bottom w:w="0" w:type="dxa"/>
			<w:right w:w="108" w:type="dxa"/>
			</w:tblCellMar>
			</w:tblPr>
			</w:style>
			<w:style w:type="numbering" w:default="1" w:styleId="Geenlijst">
			<w:name w:val="No List"/>
			<w:uiPriority w:val="99"/>
			<w:semiHidden/>
			<w:unhideWhenUsed/>
			</w:style>
			<w:style w:type="paragraph" w:styleId="Koptekst">
			<w:name w:val="header"/>
			<w:basedOn w:val="Standaard"/>
			<w:link w:val="KoptekstChar"/>
			<w:uiPriority w:val="99"/>
			<w:unhideWhenUsed/>
			<w:rsid w:val="00740D66"/>
			<w:pPr>
			<w:tabs>
			<w:tab w:val="center" w:pos="4536"/>
			<w:tab w:val="right" w:pos="9072"/>
			</w:tabs>
			<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
			</w:pPr>
			</w:style>
			<w:style w:type="character" w:customStyle="1" w:styleId="KoptekstChar">
			<w:name w:val="Koptekst Char"/>
			<w:basedOn w:val="Standaardalinea-lettertype"/>
			<w:link w:val="Koptekst"/>
			<w:uiPriority w:val="99"/>
			<w:rsid w:val="00740D66"/>
			</w:style>
			<w:style w:type="paragraph" w:styleId="Voettekst">
			<w:name w:val="footer"/>
			<w:basedOn w:val="Standaard"/>
			<w:link w:val="VoettekstChar"/>
			<w:uiPriority w:val="99"/>
			<w:unhideWhenUsed/>
			<w:rsid w:val="00740D66"/>
			<w:pPr>
			<w:tabs>
			<w:tab w:val="center" w:pos="4536"/>
			<w:tab w:val="right" w:pos="9072"/>
			</w:tabs>
			<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
			</w:pPr>
			</w:style>
			<w:style w:type="character" w:customStyle="1" w:styleId="VoettekstChar">
			<w:name w:val="Voettekst Char"/>
			<w:basedOn w:val="Standaardalinea-lettertype"/>
			<w:link w:val="Voettekst"/>
			<w:uiPriority w:val="99"/>
			<w:rsid w:val="00740D66"/>
			</w:style>
			</w:styles>"""

	def getXml(self) :
		return self.xmlString

class StyleFile :

	path = 'word/styles.xml'

	def __init__(self) :
		self.xmlString = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
			<w:styles xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
			<w:docDefaults>
			<w:rPrDefault>
			<w:rPr>
			<w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:cstheme="minorBidi"/>
			<w:sz w:val="22"/>
			<w:szCs w:val="22"/>
			<w:lang w:val="nl-NL" w:eastAsia="en-US" w:bidi="ar-SA"/>
			</w:rPr>
			</w:rPrDefault>
			<w:pPrDefault>
			<w:pPr>
			<w:spacing w:after="200" w:line="276" w:lineRule="auto"/>
			</w:pPr>
			</w:pPrDefault>
			</w:docDefaults>
			<w:latentStyles w:defLockedState="0" w:defUIPriority="99" w:defSemiHidden="1" w:defUnhideWhenUsed="1" w:defQFormat="0" w:count="267">
			<w:lsdException w:name="Normal" w:semiHidden="0" w:uiPriority="0" w:unhideWhenUsed="0" w:qFormat="1"/>
			<w:lsdException w:name="heading 1" w:semiHidden="0" w:uiPriority="9" w:unhideWhenUsed="0" w:qFormat="1"/>
			<w:lsdException w:name="heading 2" w:uiPriority="9" w:qFormat="1"/>
			<w:lsdException w:name="heading 3" w:uiPriority="9" w:qFormat="1"/>
			<w:lsdException w:name="heading 4" w:uiPriority="9" w:qFormat="1"/>
			<w:lsdException w:name="heading 5" w:uiPriority="9" w:qFormat="1"/>
			<w:lsdException w:name="heading 6" w:uiPriority="9" w:qFormat="1"/>
			<w:lsdException w:name="heading 7" w:uiPriority="9" w:qFormat="1"/>
			<w:lsdException w:name="heading 8" w:uiPriority="9" w:qFormat="1"/>
			<w:lsdException w:name="heading 9" w:uiPriority="9" w:qFormat="1"/>
			<w:lsdException w:name="toc 1" w:uiPriority="39"/>
			<w:lsdException w:name="toc 2" w:uiPriority="39"/>
			<w:lsdException w:name="toc 3" w:uiPriority="39"/>
			<w:lsdException w:name="toc 4" w:uiPriority="39"/>
			<w:lsdException w:name="toc 5" w:uiPriority="39"/>
			<w:lsdException w:name="toc 6" w:uiPriority="39"/>
			<w:lsdException w:name="toc 7" w:uiPriority="39"/>
			<w:lsdException w:name="toc 8" w:uiPriority="39"/>
			<w:lsdException w:name="toc 9" w:uiPriority="39"/>
			<w:lsdException w:name="caption" w:uiPriority="35" w:qFormat="1"/>
			<w:lsdException w:name="Title" w:semiHidden="0" w:uiPriority="10" w:unhideWhenUsed="0" w:qFormat="1"/>
			<w:lsdException w:name="Default Paragraph Font" w:uiPriority="1"/>
			<w:lsdException w:name="Subtitle" w:semiHidden="0" w:uiPriority="11" w:unhideWhenUsed="0" w:qFormat="1"/>
			<w:lsdException w:name="Strong" w:semiHidden="0" w:uiPriority="22" w:unhideWhenUsed="0" w:qFormat="1"/>
			<w:lsdException w:name="Emphasis" w:semiHidden="0" w:uiPriority="20" w:unhideWhenUsed="0" w:qFormat="1"/>
			<w:lsdException w:name="Table Grid" w:semiHidden="0" w:uiPriority="59" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Placeholder Text" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="No Spacing" w:semiHidden="0" w:uiPriority="1" w:unhideWhenUsed="0" w:qFormat="1"/>
			<w:lsdException w:name="Light Shading" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Light List" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Light Grid" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Shading 1" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Shading 2" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium List 1" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium List 2" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Grid 1" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Grid 2" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Grid 3" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Dark List" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Colorful Shading" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Colorful List" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Colorful Grid" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Light Shading Accent 1" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Light List Accent 1" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Light Grid Accent 1" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Shading 1 Accent 1" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Shading 2 Accent 1" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium List 1 Accent 1" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Revision" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="List Paragraph" w:semiHidden="0" w:uiPriority="34" w:unhideWhenUsed="0" w:qFormat="1"/>
			<w:lsdException w:name="Quote" w:semiHidden="0" w:uiPriority="29" w:unhideWhenUsed="0" w:qFormat="1"/>
			<w:lsdException w:name="Intense Quote" w:semiHidden="0" w:uiPriority="30" w:unhideWhenUsed="0" w:qFormat="1"/>
			<w:lsdException w:name="Medium List 2 Accent 1" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Grid 1 Accent 1" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Grid 2 Accent 1" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Grid 3 Accent 1" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Dark List Accent 1" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Colorful Shading Accent 1" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Colorful List Accent 1" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Colorful Grid Accent 1" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Light Shading Accent 2" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Light List Accent 2" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Light Grid Accent 2" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Shading 1 Accent 2" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Shading 2 Accent 2" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium List 1 Accent 2" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium List 2 Accent 2" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Grid 1 Accent 2" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Grid 2 Accent 2" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Grid 3 Accent 2" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Dark List Accent 2" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Colorful Shading Accent 2" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Colorful List Accent 2" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Colorful Grid Accent 2" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Light Shading Accent 3" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Light List Accent 3" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Light Grid Accent 3" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Shading 1 Accent 3" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Shading 2 Accent 3" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium List 1 Accent 3" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium List 2 Accent 3" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Grid 1 Accent 3" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Grid 2 Accent 3" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Grid 3 Accent 3" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Dark List Accent 3" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Colorful Shading Accent 3" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Colorful List Accent 3" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Colorful Grid Accent 3" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Light Shading Accent 4" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Light List Accent 4" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Light Grid Accent 4" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Shading 1 Accent 4" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Shading 2 Accent 4" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium List 1 Accent 4" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium List 2 Accent 4" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Grid 1 Accent 4" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Grid 2 Accent 4" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Grid 3 Accent 4" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Dark List Accent 4" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Colorful Shading Accent 4" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Colorful List Accent 4" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Colorful Grid Accent 4" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Light Shading Accent 5" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Light List Accent 5" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Light Grid Accent 5" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Shading 1 Accent 5" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Shading 2 Accent 5" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium List 1 Accent 5" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium List 2 Accent 5" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Grid 1 Accent 5" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Grid 2 Accent 5" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Grid 3 Accent 5" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Dark List Accent 5" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Colorful Shading Accent 5" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Colorful List Accent 5" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Colorful Grid Accent 5" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Light Shading Accent 6" w:semiHidden="0" w:uiPriority="60" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Light List Accent 6" w:semiHidden="0" w:uiPriority="61" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Light Grid Accent 6" w:semiHidden="0" w:uiPriority="62" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Shading 1 Accent 6" w:semiHidden="0" w:uiPriority="63" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Shading 2 Accent 6" w:semiHidden="0" w:uiPriority="64" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium List 1 Accent 6" w:semiHidden="0" w:uiPriority="65" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium List 2 Accent 6" w:semiHidden="0" w:uiPriority="66" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Grid 1 Accent 6" w:semiHidden="0" w:uiPriority="67" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Grid 2 Accent 6" w:semiHidden="0" w:uiPriority="68" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Medium Grid 3 Accent 6" w:semiHidden="0" w:uiPriority="69" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Dark List Accent 6" w:semiHidden="0" w:uiPriority="70" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Colorful Shading Accent 6" w:semiHidden="0" w:uiPriority="71" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Colorful List Accent 6" w:semiHidden="0" w:uiPriority="72" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Colorful Grid Accent 6" w:semiHidden="0" w:uiPriority="73" w:unhideWhenUsed="0"/>
			<w:lsdException w:name="Subtle Emphasis" w:semiHidden="0" w:uiPriority="19" w:unhideWhenUsed="0" w:qFormat="1"/>
			<w:lsdException w:name="Intense Emphasis" w:semiHidden="0" w:uiPriority="21" w:unhideWhenUsed="0" w:qFormat="1"/>
			<w:lsdException w:name="Subtle Reference" w:semiHidden="0" w:uiPriority="31" w:unhideWhenUsed="0" w:qFormat="1"/>
			<w:lsdException w:name="Intense Reference" w:semiHidden="0" w:uiPriority="32" w:unhideWhenUsed="0" w:qFormat="1"/>
			<w:lsdException w:name="Book Title" w:semiHidden="0" w:uiPriority="33" w:unhideWhenUsed="0" w:qFormat="1"/>
			<w:lsdException w:name="Bibliography" w:uiPriority="37"/>
			<w:lsdException w:name="TOC Heading" w:uiPriority="39" w:qFormat="1"/>
			</w:latentStyles>
			<w:style w:type="paragraph" w:default="1" w:styleId="Normal">
			<w:name w:val="Normal"/>
			<w:qFormat/>
			<w:rsid w:val="00A42688"/>
			</w:style>
			<w:style w:type="paragraph" w:styleId="Heading1">
			<w:name w:val="heading 1"/>
			<w:basedOn w:val="Normal"/>
			<w:next w:val="Normal"/>
			<w:link w:val="Heading1Char"/>
			<w:uiPriority w:val="9"/>
			<w:qFormat/>
			<w:rsid w:val="00EE745C"/>
			<w:pPr>
			<w:keepNext/>
			<w:keepLines/>
			<w:spacing w:before="480" w:after="0"/>
			<w:outlineLvl w:val="0"/>
			</w:pPr>
			<w:rPr>
			<w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>
			<w:b/>
			<w:bCs/>
			<w:color w:val="365F91" w:themeColor="accent1" w:themeShade="BF"/>
			<w:sz w:val="28"/>
			<w:szCs w:val="28"/>
			</w:rPr>
			</w:style>
			<w:style w:type="paragraph" w:styleId="Heading2">
			<w:name w:val="heading 2"/>
			<w:basedOn w:val="Normal"/>
			<w:next w:val="Normal"/>
			<w:link w:val="Heading2Char"/>
			<w:uiPriority w:val="9"/>
			<w:unhideWhenUsed/>
			<w:qFormat/>
			<w:rsid w:val="00EE745C"/>
			<w:pPr>
			<w:keepNext/>
			<w:keepLines/>
			<w:spacing w:before="200" w:after="0"/>
			<w:outlineLvl w:val="1"/>
			</w:pPr>
			<w:rPr>
			<w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>
			<w:b/>
			<w:bCs/>
			<w:color w:val="4F81BD" w:themeColor="accent1"/>
			<w:sz w:val="26"/>
			<w:szCs w:val="26"/>
			</w:rPr>
			</w:style>
			<w:style w:type="paragraph" w:styleId="Heading3">
			<w:name w:val="heading 3"/>
			<w:basedOn w:val="Normal"/>
			<w:next w:val="Normal"/>
			<w:link w:val="Heading3Char"/>
			<w:uiPriority w:val="9"/>
			<w:unhideWhenUsed/>
			<w:qFormat/>
			<w:rsid w:val="00EE745C"/>
			<w:pPr>
			<w:keepNext/>
			<w:keepLines/>
			<w:spacing w:before="200" w:after="0"/>
			<w:outlineLvl w:val="2"/>
			</w:pPr>
			<w:rPr>
			<w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>
			<w:b/>
			<w:bCs/>
			<w:color w:val="4F81BD" w:themeColor="accent1"/>
			</w:rPr>
			</w:style>
			<w:style w:type="paragraph" w:styleId="Heading4">
			<w:name w:val="heading 4"/>
			<w:basedOn w:val="Normal"/>
			<w:next w:val="Normal"/>
			<w:link w:val="Heading4Char"/>
			<w:uiPriority w:val="9"/>
			<w:unhideWhenUsed/>
			<w:qFormat/>
			<w:rsid w:val="00EE745C"/>
			<w:pPr>
			<w:keepNext/>
			<w:keepLines/>
			<w:spacing w:before="200" w:after="0"/>
			<w:outlineLvl w:val="3"/>
			</w:pPr>
			<w:rPr>
			<w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>
			<w:b/>
			<w:bCs/>
			<w:i/>
			<w:iCs/>
			<w:color w:val="4F81BD" w:themeColor="accent1"/>
			</w:rPr>
			</w:style>
			<w:style w:type="paragraph" w:styleId="Heading5">
			<w:name w:val="heading 5"/>
			<w:basedOn w:val="Normal"/>
			<w:next w:val="Normal"/>
			<w:link w:val="Heading5Char"/>
			<w:uiPriority w:val="9"/>
			<w:unhideWhenUsed/>
			<w:qFormat/>
			<w:rsid w:val="00EE745C"/>
			<w:pPr>
			<w:keepNext/>
			<w:keepLines/>
			<w:spacing w:before="200" w:after="0"/>
			<w:outlineLvl w:val="4"/>
			</w:pPr>
			<w:rPr>
			<w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>
			<w:color w:val="243F60" w:themeColor="accent1" w:themeShade="7F"/>
			</w:rPr>
			</w:style>
			<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">
			<w:name w:val="Default Paragraph Font"/>
			<w:uiPriority w:val="1"/>
			<w:semiHidden/>
			<w:unhideWhenUsed/>
			</w:style>
			<w:style w:type="table" w:default="1" w:styleId="TableNormal">
			<w:name w:val="Normal Table"/>
			<w:uiPriority w:val="99"/>
			<w:semiHidden/>
			<w:unhideWhenUsed/>
			<w:qFormat/>
			<w:tblPr>
			<w:tblInd w:w="0" w:type="dxa"/>
			<w:tblCellMar>
			<w:top w:w="0" w:type="dxa"/>
			<w:left w:w="108" w:type="dxa"/>
			<w:bottom w:w="0" w:type="dxa"/>
			<w:right w:w="108" w:type="dxa"/>
			</w:tblCellMar>
			</w:tblPr>
			</w:style>
			<w:style w:type="numbering" w:default="1" w:styleId="NoList">
			<w:name w:val="No List"/>
			<w:uiPriority w:val="99"/>
			<w:semiHidden/>
			<w:unhideWhenUsed/>
			</w:style>
			<w:style w:type="character" w:customStyle="1" w:styleId="Heading1Char">
			<w:name w:val="Heading 1 Char"/>
			<w:basedOn w:val="DefaultParagraphFont"/>
			<w:link w:val="Heading1"/>
			<w:uiPriority w:val="9"/>
			<w:rsid w:val="00EE745C"/>
			<w:rPr>
			<w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>
			<w:b/>
			<w:bCs/>
			<w:color w:val="365F91" w:themeColor="accent1" w:themeShade="BF"/>
			<w:sz w:val="28"/>
			<w:szCs w:val="28"/>
			</w:rPr>
			</w:style>
			<w:style w:type="character" w:customStyle="1" w:styleId="Heading2Char">
			<w:name w:val="Heading 2 Char"/>
			<w:basedOn w:val="DefaultParagraphFont"/>
			<w:link w:val="Heading2"/>
			<w:uiPriority w:val="9"/>
			<w:rsid w:val="00EE745C"/>
			<w:rPr>
			<w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>
			<w:b/>
			<w:bCs/>
			<w:color w:val="4F81BD" w:themeColor="accent1"/>
			<w:sz w:val="26"/>
			<w:szCs w:val="26"/>
			</w:rPr>
			</w:style>
			<w:style w:type="character" w:customStyle="1" w:styleId="Heading3Char">
			<w:name w:val="Heading 3 Char"/>
			<w:basedOn w:val="DefaultParagraphFont"/>
			<w:link w:val="Heading3"/>
			<w:uiPriority w:val="9"/>
			<w:rsid w:val="00EE745C"/>
			<w:rPr>
			<w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>
			<w:b/>
			<w:bCs/>
			<w:color w:val="4F81BD" w:themeColor="accent1"/>
			</w:rPr>
			</w:style>
			<w:style w:type="character" w:customStyle="1" w:styleId="Heading4Char">
			<w:name w:val="Heading 4 Char"/>
			<w:basedOn w:val="DefaultParagraphFont"/>
			<w:link w:val="Heading4"/>
			<w:uiPriority w:val="9"/>
			<w:rsid w:val="00EE745C"/>
			<w:rPr>
			<w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>
			<w:b/>
			<w:bCs/>
			<w:i/>
			<w:iCs/>
			<w:color w:val="4F81BD" w:themeColor="accent1"/>
			</w:rPr>
			</w:style>
			<w:style w:type="character" w:customStyle="1" w:styleId="Heading5Char">
			<w:name w:val="Heading 5 Char"/>
			<w:basedOn w:val="DefaultParagraphFont"/>
			<w:link w:val="Heading5"/>
			<w:uiPriority w:val="9"/>
			<w:rsid w:val="00EE745C"/>
			<w:rPr>
			<w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hAnsiTheme="majorHAnsi" w:cstheme="majorBidi"/>
			<w:color w:val="243F60" w:themeColor="accent1" w:themeShade="7F"/>
			</w:rPr>
			</w:style>
			<w:style w:type="character" w:styleId="Hyperlink">
			<w:name w:val="Hyperlink"/>
			<w:basedOn w:val="DefaultParagraphFont"/>
			<w:uiPriority w:val="99"/>
			<w:unhideWhenUsed/>
			<w:rsid w:val="00ED0E71"/>
			<w:rPr>
			<w:color w:val="0000FF" w:themeColor="hyperlink"/>
			<w:u w:val="single"/>
			</w:rPr>
			</w:style>
			<w:style w:type="table" w:styleId="TableGrid">
			<w:name w:val="Table Grid"/>
			<w:basedOn w:val="TableNormal"/>
			<w:uiPriority w:val="59"/>
			<w:rsid w:val="007808D7"/>
			<w:pPr>
			<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
			</w:pPr>
			<w:tblPr>
			<w:tblInd w:w="0" w:type="dxa"/>
			<w:tblBorders>
			<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
			<w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
			<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
			<w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
			<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
			<w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
			</w:tblBorders>
			<w:tblCellMar>
			<w:top w:w="0" w:type="dxa"/>
			<w:left w:w="108" w:type="dxa"/>
			<w:bottom w:w="0" w:type="dxa"/>
			<w:right w:w="108" w:type="dxa"/>
			</w:tblCellMar>
			</w:tblPr>
			</w:style>
			<w:style w:type="paragraph" w:styleId="ListParagraph">
			<w:name w:val="List Paragraph"/>
			<w:basedOn w:val="Normal"/>
			<w:uiPriority w:val="34"/>
			<w:qFormat/>
			<w:rsid w:val="001B5762"/>
			<w:pPr>
			<w:ind w:left="720"/>
			<w:contextualSpacing/>
			</w:pPr>
			</w:style>
			<w:style w:type="paragraph" w:styleId="BalloonText">
			<w:name w:val="Balloon Text"/>
			<w:basedOn w:val="Normal"/>
			<w:link w:val="BalloonTextChar"/>
			<w:uiPriority w:val="99"/>
			<w:semiHidden/>
			<w:unhideWhenUsed/>
			<w:rsid w:val="008752CF"/>
			<w:pPr>
			<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
			</w:pPr>
			<w:rPr>
			<w:rFonts w:ascii="Tahoma" w:hAnsi="Tahoma" w:cs="Tahoma"/>
			<w:sz w:val="16"/>
			<w:szCs w:val="16"/>
			</w:rPr>
			</w:style>
			<w:style w:type="character" w:customStyle="1" w:styleId="BalloonTextChar">
			<w:name w:val="Balloon Text Char"/>
			<w:basedOn w:val="DefaultParagraphFont"/>
			<w:link w:val="BalloonText"/>
			<w:uiPriority w:val="99"/>
			<w:semiHidden/>
			<w:rsid w:val="008752CF"/>
			<w:rPr>
			<w:rFonts w:ascii="Tahoma" w:hAnsi="Tahoma" w:cs="Tahoma"/>
			<w:sz w:val="16"/>
			<w:szCs w:val="16"/>
			</w:rPr>
			</w:style>
			</w:styles>"""

	def getXml(self) :
		return self.xmlString

class ThemeFile() :

	path = 'word/theme/theme1.xml'

	def __init__(self) :
		self.xmlString = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
			<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
			<a:themeElements>
			<a:clrScheme name="Office">
			<a:dk1>
			<a:sysClr val="windowText" lastClr="000000"/>
			</a:dk1>
			<a:lt1>
			<a:sysClr val="window" lastClr="FFFFFF"/>
			</a:lt1>
			<a:dk2>
			<a:srgbClr val="1F497D"/>
			</a:dk2>
			<a:lt2>
			<a:srgbClr val="EEECE1"/>
			</a:lt2>
			<a:accent1>
			<a:srgbClr val="4F81BD"/>
			</a:accent1>
			<a:accent2>
			<a:srgbClr val="C0504D"/>
			</a:accent2>
			<a:accent3>
			<a:srgbClr val="9BBB59"/>
			</a:accent3>
			<a:accent4>
			<a:srgbClr val="8064A2"/>
			</a:accent4>
			<a:accent5>
			<a:srgbClr val="4BACC6"/>
			</a:accent5>
			<a:accent6>
			<a:srgbClr val="F79646"/>
			</a:accent6>
			<a:hlink>
			<a:srgbClr val="0000FF"/>
			</a:hlink>
			<a:folHlink>
			<a:srgbClr val="800080"/>
			</a:folHlink>
			</a:clrScheme>
			<a:fontScheme name="Office">
			<a:majorFont>
			<a:latin typeface="Cambria"/>
			<a:ea typeface=""/>
			<a:cs typeface=""/>
			<a:font script="Arab" typeface="Times New Roman"/>
			<a:font script="Hebr" typeface="Times New Roman"/>
			<a:font script="Thai" typeface="Angsana New"/>
			<a:font script="Ethi" typeface="Nyala"/>
			<a:font script="Beng" typeface="Vrinda"/>
			<a:font script="Gujr" typeface="Shruti"/>
			<a:font script="Khmr" typeface="MoolBoran"/>
			<a:font script="Knda" typeface="Tunga"/>
			<a:font script="Guru" typeface="Raavi"/>
			<a:font script="Cans" typeface="Euphemia"/>
			<a:font script="Cher" typeface="Plantagenet Cherokee"/>
			<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
			<a:font script="Tibt" typeface="Microsoft Himalaya"/>
			<a:font script="Thaa" typeface="MV Boli"/>
			<a:font script="Deva" typeface="Mangal"/>
			<a:font script="Telu" typeface="Gautami"/>
			<a:font script="Taml" typeface="Latha"/>
			<a:font script="Syrc" typeface="Estrangelo Edessa"/>
			<a:font script="Orya" typeface="Kalinga"/>
			<a:font script="Mlym" typeface="Kartika"/>
			<a:font script="Laoo" typeface="DokChampa"/>
			<a:font script="Sinh" typeface="Iskoola Pota"/>
			<a:font script="Mong" typeface="Mongolian Baiti"/>
			<a:font script="Viet" typeface="Times New Roman"/>
			<a:font script="Uigh" typeface="Microsoft Uighur"/>
			</a:majorFont>
			<a:minorFont>
			<a:latin typeface="Calibri"/>
			<a:ea typeface=""/>
			<a:cs typeface=""/>
			<a:font script="Arab" typeface="Arial"/>
			<a:font script="Hebr" typeface="Arial"/>
			<a:font script="Thai" typeface="Cordia New"/>
			<a:font script="Ethi" typeface="Nyala"/>
			<a:font script="Beng" typeface="Vrinda"/>
			<a:font script="Gujr" typeface="Shruti"/>
			<a:font script="Khmr" typeface="DaunPenh"/>
			<a:font script="Knda" typeface="Tunga"/>
			<a:font script="Guru" typeface="Raavi"/>
			<a:font script="Cans" typeface="Euphemia"/>
			<a:font script="Cher" typeface="Plantagenet Cherokee"/>
			<a:font script="Yiii" typeface="Microsoft Yi Baiti"/>
			<a:font script="Tibt" typeface="Microsoft Himalaya"/>
			<a:font script="Thaa" typeface="MV Boli"/>
			<a:font script="Deva" typeface="Mangal"/>
			<a:font script="Telu" typeface="Gautami"/>
			<a:font script="Taml" typeface="Latha"/>
			<a:font script="Syrc" typeface="Estrangelo Edessa"/>
			<a:font script="Orya" typeface="Kalinga"/>
			<a:font script="Mlym" typeface="Kartika"/>
			<a:font script="Laoo" typeface="DokChampa"/>
			<a:font script="Sinh" typeface="Iskoola Pota"/>
			<a:font script="Mong" typeface="Mongolian Baiti"/>
			<a:font script="Viet" typeface="Arial"/>
			<a:font script="Uigh" typeface="Microsoft Uighur"/>
			</a:minorFont>
			</a:fontScheme>
			<a:fmtScheme name="Office">
			<a:fillStyleLst>
			<a:solidFill>
			<a:schemeClr val="phClr"/>
			</a:solidFill>
			<a:gradFill rotWithShape="1">
			<a:gsLst>
			<a:gs pos="0">
			<a:schemeClr val="phClr">
			<a:tint val="50000"/>
			<a:satMod val="300000"/>
			</a:schemeClr>
			</a:gs>
			<a:gs pos="35000">
			<a:schemeClr val="phClr">
			<a:tint val="37000"/>
			<a:satMod val="300000"/>
			</a:schemeClr>
			</a:gs>
			<a:gs pos="100000">
			<a:schemeClr val="phClr">
			<a:tint val="15000"/>
			<a:satMod val="350000"/>
			</a:schemeClr>
			</a:gs>
			</a:gsLst>
			<a:lin ang="16200000" scaled="1"/>
			</a:gradFill>
			<a:gradFill rotWithShape="1">
			<a:gsLst>
			<a:gs pos="0">
			<a:schemeClr val="phClr">
			<a:shade val="51000"/>
			<a:satMod val="130000"/>
			</a:schemeClr>
			</a:gs>
			<a:gs pos="80000">
			<a:schemeClr val="phClr">
			<a:shade val="93000"/>
			<a:satMod val="130000"/>
			</a:schemeClr>
			</a:gs>
			<a:gs pos="100000">
			<a:schemeClr val="phClr">
			<a:shade val="94000"/>
			<a:satMod val="135000"/>
			</a:schemeClr>
			</a:gs>
			</a:gsLst>
			<a:lin ang="16200000" scaled="0"/>
			</a:gradFill>
			</a:fillStyleLst>
			<a:lnStyleLst>
			<a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
			<a:solidFill>
			<a:schemeClr val="phClr">
			<a:shade val="95000"/>
			<a:satMod val="105000"/>
			</a:schemeClr>
			</a:solidFill>
			<a:prstDash val="solid"/>
			</a:ln>
			<a:ln w="25400" cap="flat" cmpd="sng" algn="ctr">
			<a:solidFill>
			<a:schemeClr val="phClr"/>
			</a:solidFill>
			<a:prstDash val="solid"/>
			</a:ln>
			<a:ln w="38100" cap="flat" cmpd="sng" algn="ctr">
			<a:solidFill>
			<a:schemeClr val="phClr"/>
			</a:solidFill>
			<a:prstDash val="solid"/>
			</a:ln>
			</a:lnStyleLst>
			<a:effectStyleLst>
			<a:effectStyle>
			<a:effectLst>
			<a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0">
			<a:srgbClr val="000000">
			<a:alpha val="38000"/>
			</a:srgbClr>
			</a:outerShdw>
			</a:effectLst>
			</a:effectStyle>
			<a:effectStyle>
			<a:effectLst>
			<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
			<a:srgbClr val="000000">
			<a:alpha val="35000"/>
			</a:srgbClr>
			</a:outerShdw>
			</a:effectLst>
			</a:effectStyle>
			<a:effectStyle>
			<a:effectLst>
			<a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
			<a:srgbClr val="000000">
			<a:alpha val="35000"/>
			</a:srgbClr>
			</a:outerShdw>
			</a:effectLst>
			<a:scene3d>
			<a:camera prst="orthographicFront">
			<a:rot lat="0" lon="0" rev="0"/>
			</a:camera>
			<a:lightRig rig="threePt" dir="t">
			<a:rot lat="0" lon="0" rev="1200000"/>
			</a:lightRig>
			</a:scene3d>
			<a:sp3d>
			<a:bevelT w="63500" h="25400"/>
			</a:sp3d>
			</a:effectStyle>
			</a:effectStyleLst>
			<a:bgFillStyleLst>
			<a:solidFill>
			<a:schemeClr val="phClr"/>
			</a:solidFill>
			<a:gradFill rotWithShape="1">
			<a:gsLst>
			<a:gs pos="0">
			<a:schemeClr val="phClr">
			<a:tint val="40000"/>
			<a:satMod val="350000"/>
			</a:schemeClr>
			</a:gs>
			<a:gs pos="40000">
			<a:schemeClr val="phClr">
			<a:tint val="45000"/>
			<a:shade val="99000"/>
			<a:satMod val="350000"/>
			</a:schemeClr>
			</a:gs>
			<a:gs pos="100000">
			<a:schemeClr val="phClr">
			<a:shade val="20000"/>
			<a:satMod val="255000"/>
			</a:schemeClr>
			</a:gs>
			</a:gsLst>
			<a:path path="circle">
			<a:fillToRect l="50000" t="-80000" r="50000" b="180000"/>
			</a:path>
			</a:gradFill>
			<a:gradFill rotWithShape="1">
			<a:gsLst>
			<a:gs pos="0">
			<a:schemeClr val="phClr">
			<a:tint val="80000"/>
			<a:satMod val="300000"/>
			</a:schemeClr>
			</a:gs>
			<a:gs pos="100000">
			<a:schemeClr val="phClr">
			<a:shade val="30000"/>
			<a:satMod val="200000"/>
			</a:schemeClr>
			</a:gs>
			</a:gsLst>
			<a:path path="circle">
			<a:fillToRect l="50000" t="50000" r="50000" b="50000"/>
			</a:path>
			</a:gradFill>
			</a:bgFillStyleLst>
			</a:fmtScheme>
			</a:themeElements>
			<a:objectDefaults/>
			<a:extraClrSchemeLst/>
			</a:theme>"""

	def getXml(self) :
		return self.xmlString
