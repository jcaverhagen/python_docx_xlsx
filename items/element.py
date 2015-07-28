from lxml import etree

WPREFIXES = {
        'w' : '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    }

class Element :

	def createElement(self, tag, text=None, attr=None) :
		element = etree.Element(WPREFIXES['w'] + tag, nsmap=None)

		if attr :
			for key, value in attr.items() :
				element.set(WPREFIXES['w'] + key, value)

		if text :
			element.text = text

		return element