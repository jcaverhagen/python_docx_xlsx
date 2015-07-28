from lxml import etree

WPREFIXES = {
        'w' : 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        #relationships
        'r':  'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    }

class Element :

	def createElement(self, tag, text=None, attr=None, prefix='w') :
		element = etree.Element('{' + WPREFIXES[prefix] + '}' + tag, nsmap=None)

		if attr :
			for key, value in attr.items() :
				#id when hyperlink (relationship)
				if key == 'id' :
					element.set('{' + WPREFIXES['r'] + '}' + key, value)
				else :
					element.set('{' + WPREFIXES[prefix] + '}' + key, value)

		if text :
			element.text = text

		return element