from lxml import etree
from element import Element
from os.path import basename
from PIL import Image as PILImage

_basic = """<w:p>
				<w:r>
					<w:rPr>
						<w:noProof/>
						<w:lang w:eastAsia="nl-NL"/>
					</w:rPr>
					<w:drawing>
						<wp:inline distT="0" distB="0" distL="0" distR="0">
							<wp:extent cx="5760720" cy="4320540"/>
							<wp:effectExtent l="19050" t="0" r="0" b="0"/>
							<wp:docPr id="1" name="Picture 0" descr="image_name.jpg"/>
							<wp:cNvGraphicFramePr>
								<a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
							</wp:cNvGraphicFramePr>
							<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
								<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
									<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
										<pic:nvPicPr>
											<pic:cNvPr id="0" name="image_name.jpg"/>
											<pic:cNvPicPr/>
										</pic:nvPicPr>
										<pic:blipFill>
											<a:blip r:embed="rId5"/>
											<a:stretch>
												<a:fillRect/>
											</a:stretch>
										</pic:blipFill>
										<pic:spPr>
											<a:xfrm>
												<a:off x="0" y="0"/>
												<a:ext cx="5760720" cy="4320540"/>
											</a:xfrm>
											<a:prstGeom prst="rect">
												<a:avLst/>
											</a:prstGeom>
										</pic:spPr>
									</pic:pic>
								</a:graphicData>
							</a:graphic>
						</wp:inline>
					</w:drawing>
				</w:r>
			</w:p>"""

class Image :

	_image = ''
	_p = ''

	def __init__(self, image, rel_id, width='100%', height='100%') :
		self._image = open(image)
		image_name =  basename(self._image.name).replace(' ', '-')
		
		img = PILImage.open(self._image.name)
		imgwidth = self.pixelToEmu(img.size[0])
		imgheight = self.pixelToEmu(img.size[1])
		width = (float(width.replace('%', '')) / 100) * imgwidth
		height = (float(height.replace('%', '')) / 100) * imgheight

		rel_id = rel_id
		
		self._p = Element().createElement('p')
		r = Element().createElement('r')
		rPr = Element().createElement('rPr')
		noProof = Element().createElement('noProof')
		lang = Element().createElement('lang', attr={'eastAsia' : 'nl-NL'})
		rPr.append(noProof)
		rPr.append(lang)
		drawing = Element().createElement('drawing')
		inline = Element().createElement('inline', attr={'distT' : '0', 'distR' : '0', 'distL' : '0', 'distB' : '0'}, prefix='wp', attrprefix=None)
		extend = Element().createElement('extent', attr={'cy' : str(int(height)), 'cx' : str(int(width))}, prefix='wp', attrprefix=None)
		effectExtent = Element().createElement('effectExtent', attr={'l' : '19050', 't' : '0', 'r' : '0', 'b' : '0'}, prefix='wp', attrprefix=None)
		docPr = Element().createElement('docPr', attr={'id' : '1', 'descr' : image_name, 'name' : 'Picture 0'}, prefix='wp', attrprefix=None)
		cNvGraphicFramePr = Element().createElement('cNvGraphicFramePr', prefix='wp')
		graphicFrameLocks = Element().createElement('graphicFrameLocks', attr={ 'noChangeAspect' : '1'}, prefix='a', attrprefix=None) 
		cNvGraphicFramePr.append(graphicFrameLocks)
		graphic = Element().createElement('graphic', prefix='a') 
		graphicData = Element().createElement('graphicData', attr={'uri' : 'http://schemas.openxmlformats.org/drawingml/2006/picture'}, prefix='a', attrprefix=None)
		pic = Element().createElement('pic',  prefix='pic')
		nvPicPr = Element().createElement('nvPicPr', prefix='pic')
		cNvPr = Element().createElement('cNvPr', attr={'id' : '0', 'name' : image_name}, prefix='pic', attrprefix=None)
		cNvPicPr = Element().createElement('cNvPicPr', prefix='pic')
		nvPicPr.append(cNvPr)
		nvPicPr.append(cNvPicPr)
		blipFill = Element().createElement('blipFill', prefix='pic')
		blip = Element().createElement('blip', attr={'embed' : 'rId' + str(rel_id)}, prefix='a')
		stretch = Element().createElement('stretch', prefix='a')
		fillRect = Element().createElement('fillRect', prefix='a')
		stretch.append(fillRect)
		blipFill.append(blip)
		blipFill.append(stretch)
		spPr = Element().createElement('spPr', prefix='pic')
		xfrm = Element().createElement('xfrm', prefix='a')
		off = Element().createElement('off', attr={'y' : '0', 'x' : '0'}, prefix='a', attrprefix=None)
		ext = Element().createElement('ext', attr={'cy' : str(int(height)), 'cx' : str(int(width))}, prefix='a', attrprefix=None)
		xfrm.append(off)
		xfrm.append(ext)
		prstGeom = Element().createElement('prstGeom', attr={'prst' : 'rect'}, prefix='a', attrprefix=None)
		avLst = Element().createElement('avLst', prefix='a')
		prstGeom.append(avLst)
		spPr.append(xfrm)
		spPr.append(prstGeom)
		pic.append(nvPicPr)
		pic.append(blipFill)
		pic.append(spPr)
		graphicData.append(pic)
		graphic.append(graphicData)
		inline.append(extend)
		inline.append(effectExtent)
		inline.append(docPr)
		inline.append(cNvGraphicFramePr)
		inline.append(graphic)
		drawing.append(inline)
		r.append(rPr)
		r.append(drawing)
		self._p.append(r)

	def pixelToEmu(self, pixel) :
		return int(round(pixel * 12700))

	def get(self) :
		return self._p