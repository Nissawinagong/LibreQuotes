' check if symbol is present, add if not (added symbols will have a 0 value until the next updateQuotes)
Function checkQuote (stockSymbol As String)
	Dim oSheet as Object
	checkQuote = 0

	If thisComponent.Sheets.hasByName("Link") Then
		oSheet = thisComponent.Sheets.getByName("Link")
		
		i = 0
		Do While i<1000 'limit to 1000 cycles
			If "" = oSheet.getCellByPosition(0,i).String Then
				' symbol not found before end (empty cells), add it
				oSheet.getCellByPosition(0,i).String = stockSymbol
				exit do
			ElseIf stockSymbol = oSheet.getCellByPosition(0,i).String Then
				' symbol found, get value
				checkQuote = oSheet.getCellByPosition(1,i).Value 
				exit do
			End If
			i = i+1
		Loop
		
	Else
		' sheet not found, add it
		oSheet = thisComponent.createInstance("com.sun.star.sheet.Spreadsheet")
		thisComponent.Sheets.insertByName("Link", oSheet)
		
		' also add the symbol as the first entry on the sheet
		oSheet.getCellByPosition(0,0).String = stockSymbol
	End If

End Function


Sub updateQuotes
	'iterate to list all symbols	
	sAllSymbols = ""
	
	If thisComponent.Sheets.hasByName("Link") Then
		oSheet = thisComponent.Sheets.getByName("Link")
	Else
		exit Sub
	End If
		
	r = 0
	Do While true
		If "" = oSheet.getCellByPosition(0,r).String Then
			exit do
		ElseIf sAllSymbols = "" Then
			sAllSymbols = oSheet.getCellByPosition(0,r).String
		Else
			sAllSymbols = sAllSymbols & "," & oSheet.getCellByPosition(0,r).String
		End If
		r = r+1
	Loop

	sUrl = "http://finance.yahoo.com/d/quotes.csv?s=" & sAllSymbols & "&f=sl1"
	sFilter = "Text - txt - csv (StarCalc)"
	sOptions = "44,34,SYSTEM,1,1/10/2/10/3/10/4/10/5/10/6/10/7/10/8/10/9/10"

	oSheet.LinkMode = com.sun.star.sheet.SheetLinkMode.NONE
	oSheet.link(sUrl, "", sFilter, sOptions, 1 )  
	
	saveAndReload
End Sub


sub saveAndReload
	rem ----------------------------------------------------------------------
	rem define variables
	dim document   as object
	dim dispatcher as object
	rem ----------------------------------------------------------------------
	rem get access to the document
	document   = ThisComponent.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

	rem ----------------------------------------------------------------------
	dispatcher.executeDispatch(document, ".uno:Save", "", 0, Array())

	rem ----------------------------------------------------------------------
	dispatcher.executeDispatch(document, ".uno:Reload", "", 0, Array())
end sub
