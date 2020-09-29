!INC Local Scripts.EAConstants-VBScript

'
' Script Name: Excel-Handling Skript
' Author: 
' Purpose: 
' Date: 
'
Class ExcelClass
	
	'Public SHEET-Functions
	'   - currentSheetName == Arbeitsblattname schreiben/anzeigen
	'   - addSheet == Arbeitsblatt erstellen (am Ende) und auswählen
	'	- selectSheet == Arbeitsblatt auswählen (1-x)
	'	- deleteSheet == Arbeitsblatt löschen
	
	'Public CELL-Functions
	'	- currentRow / currentCol == Aktuelle Zeile / Spalte anzeigen
	'	- nextRow / nextCol == Nächste Zeile / Spalte anzeigen
	'	- resetRow / resetCol == Zur ersten Zeile / Spalte springen
	'	- lastRow / lastCol == Letzte befüllte Zeile / Spalte anzeigen
	'	- gotoRow / gotoCol == Zur Zeile / Spalte springen
	'	- value(VALUE) == VALUE in aktuelle Zelle schreiben/lesen
	'   - valueFormula(FORMEL) == FORMEL in aktuelle Zelle schreiben
	
	'Public FORMAT Functions
	'   - wrapTextAtColumn == Text in SPALTE wird umgebrochen
	'   - setGermanCharSet == Deutsche Umlaute lesbar umwandeln
	'   - convertHTMLlist == HTML-Listen-Tags in lesbare Liste konvertieren
	'	- formatRangeAsTable == Range im aktuellen Arbeitsblatt als Tabelle formatieren
	
	'Public Functions
	'	- show == Excel-Arbeitsmappe anzeigen
	'   - open == Excel-Arbeitsmappe öffnen
	'   - newFile == Excel-Arbeitsmappe erstellen
	
	dim objExcelApp, excelRow, excelCol
	dim maxRow, maxCol	
	
	Private Sub Class_Initialize()
		excelRow = 1
		excelCol = 1
		maxRow = 0
		maxCol = 0
		set objExcelApp = CreateObject("Excel.Application")
	End Sub
	
	Public Property Let currentSheetName(name)
		objExcelApp.ActiveWorkbook.ActiveSheet.Name = name
	End Property
	
	Public Property Get currentSheetName
		currentSheetName = objExcelApp.ActiveWorkbook.ActiveSheet.Name
	End Property
	
	Public Property Get addSheet(name)
		objExcelApp.ActiveWorkbook.Sheets.Add
		objExcelApp.ActiveWorkbook.ActiveSheet.Name = name
		
		excelRow = 1
		excelCol = 1
		maxRow = 0
		maxCol = 0
	End Property
	
	Public Property Get selectSheet(number)
		objExcelApp.ActiveWorkbook.Sheets(number).Select
		
		excelRow = 1
		excelCol = 1
		maxRow = 0
		maxCol = 0
	End Property
	
	Public Property Get deleteSheet(number)
		objExcelApp.ActiveWorkbook.Sheets(number).Delete
	End Property
	
	Public Property Get currentRow()
		currentRow = excelRow
	End Property
	
	Public Property Get currentCol()
		currentCol = excelCol
	End Property
	
	Public Property Get nextRow()
		excelRow = excelRow +1
		nextRow = excelRow
	End Property
	
	Public Property Get nextCol()
		excelCol = excelCol +1
		nextCol = excelCol
	End Property
	
	Public Property Get resetRow()
		excelRow = 1
		resetRow = excelRow
	End Property
	
	Public Property Get resetCol()
		excelCol = 1
		resetCol = excelCol
	End Property
	
	Public Property Get lastRow()
		lastRow = maxRow
	End Property
	
	Public Property Get lastCol()
		lastCol = maxCol
	End Property
	
	Public Property Get gotoRow(r)
		excelRow = r
	End Property
	
	Public Property Get gotoCol(c)
		excelCol = c
	End Property
	
	Public Property Get value
		value = objExcelApp.Cells(excelRow,excelCol).Value
	End Property
	
	Public Property Let value(cellValue)
		objExcelApp.Cells(excelRow,excelCol).Value = cellValue
		if excelRow > maxRow then maxRow = excelRow
		if excelCol > maxCol then maxCol = excelCol
	End Property
	
	Public Property Let valueFormula(cellValue)
		objExcelApp.Cells(excelRow,excelCol).FormulaR1C1 = cellValue
		if excelRow > maxRow then maxRow = excelRow
		if excelCol > maxCol then maxCol = excelCol
	End Property
	
	Public Property Get wrapTextAtColumn(columnNumber)
		objExcelApp.ActiveWorkbook.ActiveSheet.Columns(columnNumber).WrapText = True
	End Property
	
	Public Property Get show()
		objExcelApp.Visible = True
	End Property
	
	Public Property Get newFile
		objExcelApp.Workbooks.Add
		objExcelApp.ActiveWorkbook.Sheets(2).Delete
		objExcelApp.ActiveWorkbook.Sheets(2).Delete
	End Property
	
	Public Property Get open(filePath)
		objExcelApp.Workbooks.Open(filePath)
	End Property
	
	Public Sub setGermanCharSetAtSheet()
		objExcelApp.DisplayAlerts = False
	
		const charAEg = "&#196;"
		const charAEgafter = "Ä"
		objExcelApp.ActiveWorkbook.ActiveSheet.Cells.Replace charAEg, charAEgafter, 2, 1, False, False, False, False
		
		const charAEk = "&#228;"
		const charAEkafter = "ä"
		objExcelApp.ActiveWorkbook.ActiveSheet.Cells.Replace charAEk, charAEkafter, 2, 1, False, False, False, False
		
		const charUEg = "&#220;"
		const charUEgafter = "Ü"
		objExcelApp.ActiveWorkbook.ActiveSheet.Cells.Replace charUEg, charUEgafter, 2, 1, False, False, False, False
		
		const charUEk = "&#252;"
		const charUEkafter = "ü"
		objExcelApp.ActiveWorkbook.ActiveSheet.Cells.Replace charUEk, charUEkafter, 2, 1, False, False, False, False
		
		const charOEk = "&#246;"
		const charOEkafter = "ö"
		objExcelApp.ActiveWorkbook.ActiveSheet.Cells.Replace charOEk, charOEkafter, 2, 1, False, False, False, False
		
		const charSZ = "&#223;"
		const charSZafter = "ß"
		objExcelApp.ActiveWorkbook.ActiveSheet.Cells.Replace charSZ, charSZafter, 2, 1, False, False, False, False
		
		const charAND = "&amp;"
		const charANDafter = "&"
		objExcelApp.ActiveWorkbook.ActiveSheet.Cells.Replace charAND, charANDafter, 2, 1, False, False, False, False
		
		objExcelApp.DisplayAlerts = True
	End Sub
	
	Public Sub convertHTMLlistAtSheet()
		objExcelApp.DisplayAlerts = False
		
		const charHListStart = "<ul>"
		const charHListStartAfter = ""
		objExcelApp.ActiveWorkbook.ActiveSheet.Cells.Replace charHListStart, charHListStartAfter, 2, 1, False, False, False, False
		
		const charHListEnd = "</ul>"
		const charHListEndAfter = ""
		objExcelApp.ActiveWorkbook.ActiveSheet.Cells.Replace charHListEnd, charHListEndAfter, 2, 1, False, False, False, False
		
		const charListEnd = "</li>"
		const charListEndAfter = ""
		objExcelApp.ActiveWorkbook.ActiveSheet.Cells.Replace charListEnd, charListEndAfter, 2, 1, False, False, False, False
		
		const charListStart = "<li>"
		const charListStartAfter = " - "
		objExcelApp.ActiveWorkbook.ActiveSheet.Cells.Replace charListStart, charListStartAfter, 2, 1, False, False, False, False
		
		objExcelApp.DisplayAlerts = True
	End Sub
	
	Public Sub formatRangeAsTable(stringRange, tableName)
		objExcelApp.ActiveWorkbook.ActiveSheet.ListObjects.Add(1, objExcelApp.ActiveWorkbook.ActiveSheet.Range(stringRange), , 1).Name = tableName
	End Sub
	
	Private Sub Class_Terminate()
		
	End Sub
End Class
