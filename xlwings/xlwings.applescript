# Allows to run the script from Script Editor for testing
VbaHandler("TestString")

on VbaHandler(ParameterString)
	set {PYTHONPATH, PythonInterpreter, PythonCommand, WORKBOOK_FULLNAME, ApplicationFullName, LOG_FILE} to SplitString(ParameterString, "|")
	set ShellCommand to PythonInterpreter & " -B -u -W ignore -c \"import sys, os;sys.path.extend(os.path.normcase(os.path.expandvars('" & PYTHONPATH & "')).split(';'));" & ¬
		PythonCommand & " \" \"--wb=" & WORKBOOK_FULLNAME & "\" \"--from_xl=1\" \"--app=" & ApplicationFullName & "\" > /dev/null 2>\"" & LOG_FILE & "\" & "
	try
		do shell script "source ~/.bash_profile"
		return do shell script "source ~/.bash_profile;" & ShellCommand
	on error errMsg number errNumber
		try
			# Try again without sourcing .bash_profile
			return do shell script ShellCommand
		on error errMsg number errNumber
			return 1
		end try
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
