Const QUOTE_SHEET 	as String 	= "Link" 
Const START_ROW 	as Integer 	= 1
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
		quoteFormat = "&f=sl1d1pva2qdyn"
		url = baseUrl & allSymbols & quoteFormat
		
		filter = "Text - txt - csv (StarCalc)"
		options = "44,34,SYSTEM,1,1/10/2/10/3/10/4/10/5/10/6/10/7/10/8/10/9/10"

		quoteSheet.LinkMode = com.sun.star.sheet.SheetLinkMode.NONE
		quoteSheet.link(url, "", filter, options, 1 )  
		
		' probably a better way to do this using the filter and options above
		quoteSheet.Rows.insertByIndex(0,1) ' this seems to overwrite rather than insert...

		quoteSheet.getCellByPosition(0,0).String = "Symbol"
		quoteSheet.getCellByPosition(1,0).String = "Value"
		quoteSheet.getCellByPosition(2,0).String = "Last Trade Date"
		quoteSheet.getCellByPosition(3,0).String = "Previous Close"
		quoteSheet.getCellByPosition(4,0).String = "Volume"
		quoteSheet.getCellByPosition(5,0).String = "Average Volume"
		quoteSheet.getCellByPosition(6,0).String = "Ex Div Date"
		quoteSheet.getCellByPosition(7,0).String = "Div / Share"
		quoteSheet.getCellByPosition(8,0).String = "Div Yield"
		quoteSheet.getCellByPosition(9,0).String = "Name"
	
		quoteSheet.getCellByPosition(10,0).String = "Most Recent Quote"
		Dim mostRecent as Double
		mostRecent = 0
		i = START_ROW
		Do While i < END_ROW 'this should alway exit on it's own, but set a limit just in case
			Dim valueToCheck as Double
			valueToCheck = quoteSheet.getCellByPosition(2,i).Value
			If valueToCheck = 0 Then
				exit do
			ElseIf valueToCheck > mostRecent Then
				mostRecent = valueToCheck
			End If
			i = i+1
		Loop
		quoteSheet.getCellByPosition(11,0).Value = mostRecent

		saveAndReload
	End If
End Sub


Function checkQuote (stockSymbol as String) as Double
	checkQuote = getSymbolData (stockSymbol, 1)
End Function

Function getQuoteDate (stockSymbol as String) as Double
	getQuoteDate = getSymbolData (stockSymbol, 2)
End Function

Function getPrevious (stockSymbol as String) as Double
	getPrevious = getSymbolData (stockSymbol, 3)
End Function

Function todayQuote (stockSymbol as String) as Integer ' for some reason AND() doesn't seem to work with Boolean, so use Integer (0/1) instead
	Dim quoteSheet as Object
	quoteSheet = getSheetAdd(QUOTE_SHEET) 

	If quoteSheet.getCellByPosition(11,0).Value = getQuoteDate(stockSymbol) Then
		todayQuote = 1
	Else
		todayQuote = 0
	End If
End Function

Function getSymbolData (stockSymbol as String, dataIndex as Integer) as Double
	'getSymbolData = 0 	' Having this resets the quote to zero if the symbol is not found.
						' If updating quotes fails, this may happen and reset all values to o.
						' Do nothing here to just leave the prior quotes.
						
	If "0" = stockSymbol Then
		' protect against bad cells in the sheet
		Msgbox "No data retrieved for zero symbol"
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
			getSymbolData = quoteSheet.getCellByPosition(dataIndex,i).Value 
			exit do
		ElseIf "" = symbolToCheck Then
			' symbol not found before end (empty cells), add it to the sheet
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
