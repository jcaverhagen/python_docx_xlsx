from lxml import etree
from element import Element
from colors import Colors

_basic = """<w:p>
				<w:pPr>
					<w:pStyle w:val="NormalWeb"/>
					<w:spacing w:before="120" w:after="120"/>
				</w:pPr>
				<w:r>
					<w:t xml"space="preserve">I feel that there is much to be said for the Celtic belief that the souls of those whom we have lost are held captive in some inferior being...</w:t>
				</w:r>
			</w:p>"""

class Paragraph :
	
	_prop = ''
	
	def __init__(self, text='', style='NormalWeb', bold=False, italic=False, underline=False, 
					uppercase=False, color=False, font=False) :
		#root element
		self.para = Element().createElement('p')
		
		#style element
		pPr = Element().createElement('pPr')
		self.para.append(pPr)
		pPr.append(Element().createElement('pStyle', attr={'val' : style}))

		#run and text element
		run = Element().createElement('r')
		self._prop = Element().createElement('rPr')
		run.append(self._prop)
		
		if bold is not False :
			self._prop.append(Element().createElement('b', attr={'val' : 'true'}))
		if italic is not False :
			self._prop.append(Element().createElement('i', attr={'val' : 'true'}))
		if underline is not False :
			if underline in Colors().colors :
				self._prop.append(Element().createElement('u', attr={'val' : 'single', 'color' : Colors().colors[color].replace('#', '')}))
			else :
				if underline == True :
					underline = '#000000'
				self._prop.append(Element().createElement('u', attr={'val' : 'single', 'color' : underline.replace('#', '')}))
		if uppercase is not False :
			self._prop.append(Element().createElement('caps', attr={'val' : 'true'}))
		if color is not False :
			if color in Colors().colors :
				self._prop.append(Element().createElement('color', attr={'val' : Colors().colors[color].replace('#', '')}))
			else :
				self._prop.append(Element().createElement('color', attr={'val' : color.replace('#', '')}))
		if font is not False :
			self._prop.append(Element().createElement('rFonts', attr={'ascii' : font, 'hAnsi' : font}))

		textEl = Element().createElement('t')
		textEl.text = text
		run.append(textEl)
		self.para.append(run)
	
	def get(self) :
		return self.para

class Break :

	def __init__(self, type=None) :
		if type == 'page' :
			self.breakEl = Element().createElement('br', attr={'type' : 'page'})
		else :
			self.breakEl = Element().createElement('br')


	def get(self) :
		return self.breakEl