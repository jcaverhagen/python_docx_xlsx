from universal import defaults

class Functions :

	def searchAndReplace(self, elem, regex, replacement, startPosition, endPosition) :
		cursorPosition = startPosition
		charCount = 0
		replacePosition = 0
		for el in elem :
			if el.tag == '{' + defaults.WPREFIXES['w'] + '}'  + 'r' :
				for l in el :
					if l.tag == '{' + defaults.WPREFIXES['w'] + '}'  + 't' :
						charCount = charCount + len(l.text)

						if (charCount - len(l.text)) <= (startPosition + replacePosition) <= charCount or (charCount - len(l.text)) <= endPosition <= charCount :
							
							#when regex start on new element this goes wrong, fix @todo
							if regex[replacePosition] not in l.text :
								cursorPosition = 1
							else :
								#create temp list for replacing at specific position
								tempList = []
								for i in l.text :
									tempList.append(i)

								#loop with cursor through text
								while (len(l.text) - cursorPosition  >= 0) :
									#check for replacements with different length
									if len(replacement) > (replacePosition) :
										tempList[cursorPosition - 1] = replacement[replacePosition]
									else :
										tempList[cursorPosition - 1] = ''
									
								 	#only update replacePosition when is not at end of string
								 	if replacePosition < (len(regex) - 1) :
										replacePosition += 1
									else :
										#add other characters after latest cursor position
										if len(regex) < len(replacement) :
											tempList[cursorPosition - 1] = tempList[cursorPosition - 1] + replacement[replacePosition:]
										break

									cursorPosition += 1
									
								
								#clear old text element and fill with temp list
								l.text = ''
								for char in tempList :
									l.text = l.text + char

								#reset cursorPosition
								cursorPosition = 1
						else :
							cursorPosition -= len(l.text)
						
