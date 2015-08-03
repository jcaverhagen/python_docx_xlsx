from lxml import etree
from element import Element

_basic = """<w:p>
				<w:pPr>
					<w:pStyle w:val="ListParagraph"/>
					<w:numPr>
						<w:ilvl w:val="0"/>
						<w:numId w:val="1"/>
					</w:numPr>
				</w:pPr>
				<w:r>
					<w:t>This is the first numbered paragraph.</w:t>
				</w:r>
			</w:p>"""

class List :

	def __init__(self, position='last') :
		self._position = position
		
		self._list = []

	def addItem(self, text, level = 0) :
		p		= Element().createElement('p')
		pPr 	= Element().createElement('pPr')
		p.append(pPr)

		style 	= Element().createElement('pStyle', attr={'val' : 'ListParagraph'})
		numPr	= Element().createElement('numPr')
		pPr.append(style)
		pPr.append(numPr)
		
		levels	= Element().createElement('ilvl', attr={'val' : str(level)})
		numId	= Element().createElement('numId', attr={'val' : '1'})
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