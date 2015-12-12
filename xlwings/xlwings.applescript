# Enables to run the script from Script Editor
VbaHandler("testString")

on VbaHandler(paramString)
	set {PYTHONPATH, PythonInterpreter, PythonCommand, WORKBOOK_FULLNAME, ApplicationFullName, LOG_FILE} to SplitString(paramString, ",")
	try
		return do shell script "source ~/.bash_profile;" & PythonInterpreter & "python -u -W ignore -c \"import sys;sys.path.extend('" & PYTHONPATH & "'.split(';'));" & ¬
			PythonCommand & " \" \"" & WORKBOOK_FULLNAME & "\" \"from_xl\" \"" & ApplicationFullName & "\" 2>\"" & LOG_FILE & "\" "
	on error errMsg number errNumber
		return 1
	end try
end VbaHandler

on SplitString(TheBigString, fieldSeparator)
	# From Ron de Bruin's "Mail from Excel 2016 with Mac Mail example": www.rondebruin.nl
	tell AppleScript
		set oldTID to text item delimiters
		set text item delimiters to fieldSeparator
		set theItems to text items of TheBigString
		set text item delimiters to oldTID
	end tell
	return theItems
end SplitString
