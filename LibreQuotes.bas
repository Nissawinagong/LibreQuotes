Const QUOTE_SHEET 	as String 	= "Link" 
Const START_ROW 	as Integer 	= 0
Const END_ROW		as Integer 	= 1000

Sub updateQuotes
	'iterate to list all symbols	
	Dim allSymbols as String	
	allSymbols = ""
	
	Dim quoteSheet as Object
	quoteSheet = getSheetAdd(QUOTE_SHEET) 
		
	Dim i as Integer
	i = START_ROW
	Do While  i < END_ROW 'this should alway exit on it's own, but set a limit just in case
		Dim newSymbol as String
		newSymbol = quoteSheet.getCellByPosition(0,i).String
		If "" = newSymbol Then
			' found the end (empty cells)
			exit do
		ElseIf allSymbols = "" Then
			' add first symbol to the list (should be i == START_ROW)
			allSymbols = newSymbol
			If Not(i = START_ROW) Then
				Msgbox "The first symbol was not found on the start row"
			End If
		Else
			' add subsequent symbols to the list
			allSymbols = allSymbols & "," & newSymbol
		End If
		i = i+1
	Loop

	If Not("" = allSymbols) Then
		Dim baseUrl as String
		Dim quoteFormat as String
		Dim url as String
		Dim filter as String
		Dim options as String

		' see http://www.seangw.com/wordpress/2010/01/formatting-stock-data-from-yahoo-finance/
		baseUrl = "http://finance.yahoo.com/d/quotes.csv?s=" 
		quoteFormat = "&f=sl1nld1cc1p2va2ipomws7r1qdyj1" 
		url = baseUrl & allSymbols & quoteFormat
		
		filter = "Text - txt - csv (StarCalc)"
		options = "44,34,SYSTEM,1,1/10/2/10/3/10/4/10/5/10/6/10/7/10/8/10/9/10"

		quoteSheet.LinkMode = com.sun.star.sheet.SheetLinkMode.NONE
		quoteSheet.link(url, "", filter, options, 1 )  
	
		saveAndReload
	End If
End Sub


Function checkQuote (stockSymbol as String) as Double
	'checkQuote = 0 	' Having this resets the quote to zero if the symbol is not found.
						' If updating quotes fails, this may happen and reset all values to o.
						' Do nothing here to just leave the prior quotes.
						
	If "" = stockSymbol Then
		' protect against bad cells in the sheet
		exit Function
	End If

	Dim quoteSheet as Object
	quoteSheet = getSheetAdd(QUOTE_SHEET)
	
	Dim i as Integer	
	i = START_ROW
	Do While i < END_ROW 'this should alway exit on it's own, but set a limit just in case
		Dim symbolToCheck as String
		symbolToCheck = quoteSheet.getCellByPosition(0,i).String
		If symbolToCheck = stockSymbol Then
			' symbol found, get value
			checkQuote = quoteSheet.getCellByPosition(1,i).Value 
			exit do
		ElseIf "" = symbolToCheck Then
			' symbol not found before end (empty cells), initialize value to 0 and add it to the sheet
			checkQuote = 0.0
			quoteSheet.getCellByPosition(0,i).String = stockSymbol
			Msgbox "Value not found for Symbol: " & stockSymbol & chr(13) & "You can try running updateQuotes to get a value for it"
			exit do
		End If
		i = i+1
	Loop
End Function


' get sheetName if it exists and add it if not
Function getSheetAdd (sheetName as String) as Object
	If Not thisComponent.Sheets.hasByName(sheetName) Then
		' the code for adding may be broken and needs testing...
		Dim tmpSheet as Object
		tmpSheet = thisComponent.createInstance("com.sun.star.sheet.Spreadsheet")
		thisComponent.Sheets.insertByName(sheetName, tmpSheet)
	End If
	getSheetAdd = thisComponent.Sheets.getByName(sheetName)
End Function


Sub saveAndReload
	Dim document   as Object
	Dim dispatcher as Object

	document   = ThisComponent.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

	dispatcher.executeDispatch(document, ".uno:Save", "", 0, Array())
	dispatcher.executeDispatch(document, ".uno:Reload", "", 0, Array())
End Sub
