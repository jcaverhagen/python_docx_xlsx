from lxml import etree

WPREFIXES = {
        'w' : 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        #relationships
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        #drawing
        'wp' : 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
        'a' : 'http://schemas.openxmlformats.org/drawingml/2006/main',
        'pic' : 'http://schemas.openxmlformats.org/drawingml/2006/picture'
    }

class Element :

	def createElement(self, tag, text=None, attr=None, prefix='w', attrprefix='') :
		
		if prefix :
			etree.register_namespace(prefix, WPREFIXES[prefix])
			element = etree.Element('{' + WPREFIXES[prefix] + '}' + tag, nsmap=None)
		else :
			element = etree.Element(tag)

		if attr :

			if attrprefix == '' :
				attrprefix = prefix

			for key, value in attr.items() :
				#id when hyperlink (relationship)
				if key == 'rel_id' or key == 'embed' :
					if key == 'rel_id' :
						element.set('{' + WPREFIXES['r'] + '}id', value)
					else :
						element.set('{' + WPREFIXES['r'] + '}' + key, value)
				elif not attrprefix :
					element.set(key, value)
				else :
					element.set('{' + WPREFIXES[attrprefix] + '}' + key, value)

		if text :
			element.text = text

		return element