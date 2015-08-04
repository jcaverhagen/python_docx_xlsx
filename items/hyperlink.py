from lxml import etree
from element import Element

_basic = """<w:p>
				<w:hyperlink r:id="rId25">
					<w:r>
						<w:rPr>
							<w:rStyle w:val="Hyperlink"/>
						</w:rPr>
						<w:t>
							Google
						</w:t>
					</w:r>
				</w:hyperlink>
			</w:p>"""

HYPERLINK_SCHEMA = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'

class Hyperlink :
	
	_hyperlink = ''
	_rel_id = ''
	_url = ''

	def __init__(self, text, rel_id, url, anchor=None) :
		self._rel_id = rel_id
		self._url = url

		if anchor is None :
			self._hyperlink = Element().createElement('hyperlink', attr={'rel_id' : 'rId' + rel_id})
		else :
			self._hyperlink = Element().createElement('hyperlink', attr={'anchor' : anchor})
		
		run = Element().createElement('r')
		
		rPr = Element().createElement('rPr')
		style = Element().createElement('rStyle', attr={'val' : 'Hyperlink'})
		rPr.append(style)
		run.append(rPr)

		textEl = Element().createElement('t', text=text, attr={'space' : 'preserve'})
		run.append(textEl)

		self._hyperlink.append(run)

	def get(self) :
		return self._hyperlink

	def getRelation(self) :
		attr = {'TargetMode' : 'External', 'Type' : HYPERLINK_SCHEMA, 'Id' : 'rId' + self._rel_id, 'Target' : self._url}
		rel = Element().createElement('Relationship', prefix=None, attr=attr)
		return rel