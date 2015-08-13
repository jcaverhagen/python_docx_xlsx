import math
from lxml import etree
from universal.element import Element

_basic = """<w:tbl>
				<w:tblPr>
					<w:tblStyle w:val="TableGrid"/>
					<w:tblW w:w="5000" w:type="pct"/>
				</w:tblPr>
				<w:tblGrid>
					<w:gridCol w:w="2880"/>
					<w:gridCol w:w="2880"/>
					<w:gridCol w:w="2880"/>
				</w:tblGrid>
				<w:tr>
					<w:tc>
						<w:tcPr>
							<w:tcW w:w="2880" w:type="dxa"/>
						</w:tcPr>
						<w:p>
							<w:r>
								<w:t>AAA</w:t>
							</w:r>
						</w:p>
					</w:tc>
					<w:tc>
						<w:tcPr>
							<w:tcW w:w="2880" w:type="dxa"/>
						</w:tcPr>
						<w:p>
							<w:r>
								<w:t>BBB</w:t>
							</w:r>
						</w:p>
					</w:tc>
					<w:tc>
						<w:tcPr>
							<w:tcW w:w="2880" w:type="dxa"/>
						</w:tcPr>
						<w:p>
							<w:r>
								<w:t>CCC</w:t>
							</w:r>
						</w:p>
					</w:tc>
				</w:tr>
			</w:tbl>"""

class Table :

	_table = ''
	_columnWidth = ''
	_position = ''

	def __init__(self, width, columns=1, style='TableGrid', position='last') :
		self._position = position

		self._table = Element().createElement('tbl')
		tblPr = Element().createElement('tblPr')
		style = Element().createElement('tblStyle', attr={'val' : style})
		tblW = Element().createElement('tblW', attr={'w' : str(width), 'type' : 'pct'})

		tblPr.append(style)
		tblPr.append(tblW)
		self._table.append(tblPr)

		#calculate column width
		self._columnWidth = math.floor(width / columns)
		tblGrid = Element().createElement('tblGrid')
		for i in range(columns) :
			tblGrid.append(Element().createElement('gridCol', attr={'w' : str(self._columnWidth)}))

		self._table.append(tblGrid)

	def addRow(self, val, style='default') :
		row = self._getNewRow()

		for v in val :
			column = self._getNewColumn(style)
			
			textEl = Element().createElement('t')
			textEl.text = v
			para = Element().createElement('p')
			run = Element().createElement('r')
			run.append(textEl)
			para.append(run)
			column.append(para)
			
			row.append(column)

		self._table.append(row)

	def get(self) :
		return self._table

	def getPosition(self) :
		return self._position

	def _getNewRow(self) :
		return Element().createElement('tr')

	def _getNewColumn(self, style) :
		column = Element().createElement('tc')
		tcPr = Element().createElement('tcPr')
		tcPr.append(Element().createElement('tcW', attr={'w' : str(self._columnWidth), 'type' : 'dxa'}))

		if style == 'dashed' :
			tcBorders = Element().createElement('tcBorders')
			attr = {'val' : 'dashed', 'sz' : '4', 'space' : '0', 'color' : 'auto'}
			top = Element().createElement('top', attr=attr)
			left = Element().createElement('left', attr=attr)
			bottom = Element().createElement('bottom', attr=attr)
			right = Element().createElement('right', attr=attr)
			tcBorders.append(top)
			tcBorders.append(left)
			tcBorders.append(bottom)
			tcBorders.append(right)
			tcPr.append(tcBorders)

		column.append(tcPr)

		return column
