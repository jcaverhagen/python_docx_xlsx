from lxml import etree
from element import Element

_basic = """<w:p>
				<w:pPr>
					<w:pStyle w:val="ListParagraph"/>
					<w:numPr>
						<w:ilvl w:val="0"/>
						<w:numId w:val="1"/>
					</w:numPr>
					<w:ind w:start="10"/>
				</w:pPr>
				<w:r>
					<w:t>This is the first numbered paragraph.</w:t>
				</w:r>
			</w:p>"""

list_types = {
	'numeric' : '1',
	'bullet' : '2'
}

class List :

	_list = []
	_position = ''
	_type = ''
	
	def __init__(self, position='last', type='numeric') :
		self._position = position
		if type in list_types :
			self._type = list_types[type]
		else :
			self._type = list_types['numberic']

	def addItem(self, text) :
		p		= Element().createElement('p')
		pPr 	= Element().createElement('pPr')
		p.append(pPr)

		style 	= Element().createElement('pStyle', attr={'val' : 'ListParagraph'})
		numPr	= Element().createElement('numPr')
		ind		= Element().createElement('ind', attr={'left' : '5', 'firstLine' : '400'})
		pPr.append(style)
		pPr.append(numPr)
		pPr.append(ind)

		levels	= Element().createElement('ilvl', attr={'val' : '0'})
		numId	= Element().createElement('numId', attr={'val' : self._type})
		numPr.append(levels)
		numPr.append(numId)

		run = Element().createElement('r')
		p.append(run)

		textEl = Element().createElement('t')
		textEl.text = text
		run.append(textEl)

		self._list.append(p)

	def getPosition(self) :
		return self._position

	def get(self) :
		return self._list