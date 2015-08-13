from lxml import etree
from universal.element import Element
from universal import defaults

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
	
	def __init__(self, text='', style='NormalWeb', styles=None) :
	
		self._styles = {
			'bold' : False,
			'italic' : False,
			'underline' : False,
			'uppercase' : False,
			'color' : False,
			'font' : False,
			'size' : False
		}

		if styles is not None :
			for key, value in styles.items() :
				self._styles[key] = value

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

		#set styles
		self.setStyles()
		
		textEl = Element().createElement('t', attr={'space' : 'preserve'})
		textEl.text = text
		run.append(textEl)
		self.para.append(run)
	
	def setStyles(self) :
		if self._styles['bold'] is not False :
			self._prop.append(Element().createElement('b', attr={'val' : 'true'}))
		if self._styles['italic'] is not False :
			self._prop.append(Element().createElement('i', attr={'val' : 'true'}))
		if self._styles['underline'] is not False :
			if self._styles['underline'] in defaults.colors :
				self._prop.append(Element().createElement('u', attr={'val' : 'single', 'color' : defaults.colors[self._styles['color']].replace('#', '')}))
			else :
				if self._styles['underline'] == True :
					underline = '#000000'
				self._prop.append(Element().createElement('u', attr={'val' : 'single', 'color' : underline.replace('#', '')}))
		if self._styles['uppercase'] is not False :
			self._prop.append(Element().createElement('caps', attr={'val' : 'true'}))
		if self._styles['color'] is not False :
			if self._styles['color'] in defaults.colors :
				self._prop.append(Element().createElement('color', attr={'val' : defaults.colors[self._styles['color']].replace('#', '')}))
			else :
				self._prop.append(Element().createElement('color', attr={'val' : self._styles['color'].replace('#', '')}))
		if self._styles['font'] is not False :
			self._prop.append(Element().createElement('rFonts', attr={'ascii' : self._styles['font'], 'hAnsi' : self._styles['font']}))
		if self._styles['size'] is not False :
			self._prop.append(Element().createElement('sz', attr={'val' : str(self._styles['size'] * 2)}))

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
