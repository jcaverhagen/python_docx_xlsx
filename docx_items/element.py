from lxml import etree
import defaults

class Element :

	def createElement(self, tag, text=None, attr=None, prefix='w', attrprefix='') :
		
		if prefix :
			etree.register_namespace(prefix, defaults.WPREFIXES[prefix])
			element = etree.Element('{' + defaults.WPREFIXES[prefix] + '}' + tag, nsmap=None)
		else :
			element = etree.Element(tag)

		if attr :

			if attrprefix == '' :
				attrprefix = prefix

			for key, value in attr.items() :
				#id when hyperlink (relationship)
				if key == 'rel_id' or key == 'embed' :
					if key == 'rel_id' :
						element.set('{' + defaults.WPREFIXES['r'] + '}id', value)
					else :
						element.set('{' + defaults.WPREFIXES['r'] + '}' + key, value)
				elif not attrprefix :
					element.set(key, value)
				else :
					element.set('{' + defaults.WPREFIXES[attrprefix] + '}' + key, value)

		if text :
			element.text = text

		return element