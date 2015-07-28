from lxml import etree
from element import Element

_basic = """<w:p>
				<w:pPr>
					<w:pStyle> w:val="NormalWeb"/>
					<w:spacing w:before="120" w:after="120"/>
				</w:pPr>
				<w:r>
					<w:t xml"space="preserve">I feel that there is much to be said for the Celtic belief that the souls of those whom we have lost are held captive in some inferior being...</w:t>
				</w:r>
			</w:p>"""

class Paragraph :
	
	_para = ''
	_run = ''
	_textEl = ''

	def __init__(self, style='NormalWeb') :
		print "init"
		#root element
		self._para = Element().createElement('p')
		
		#style element
		pPr = Element().createElement('pPr')
		style = Element().createElement('pStyle', attr={'val' : style})
		spacing = Element().createElement('spacing', attr={'before' : '120', 'after' : '120'})
		pPr.append(style)
		pPr.append(spacing)
		self._para.append(pPr)

		#run and text element
		self._run = Element().createElement('r')
		self._textEl = Element().createElement('t')
		
	def setText(self, text) :
		self._textEl.text = text

	def get(self) :
		self._run.append(self._textEl)
		self._para.append(self._run)
		return self._para
