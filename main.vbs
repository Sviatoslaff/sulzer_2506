Option Explicit
Const xlUp = -4162, xlPasteValues = -4163, xlNone = -4142
Public Const firstCol = 39, lastCol = 45

Dim tblArea
Dim qtn, plant, sorg, template, serno
Dim qtyRows, rowCount, visibleRows, sapRow, goto_pos, grid, cell 
Dim bExit, bAbort, txtStatus
Dim intRow : intRow = 4
Dim iCol
Dim targetCondition, condValue, condS, condD

'1. Запрашиваем файл QTN и получаем номер qtn, массив значений для последующего заполнения SAP Quotation
Dim excelFile
excelFile = selectExcel()

' Объявляем объект FileSystemObject
Dim fso, filePath, fileName, arrWords
' Устанавливаем путь к файлу
filePath = excelFile
' Создаем объект FileSystemObject
Set fso = CreateObject("Scripting.FileSystemObject")
' Получаем имя файла из пути
fileName = fso.GetFileName(filePath)
' Освобождаем объект
Set fso = Nothing
arrWords = Split(fileName, " ")
qtn = arrWords(0)
If Not IsNumeric(qtn) Then
	MsgBox "Первое слово имени файла должно являться номером QTN. " & qtn & " - не числовое значение!", vbSystemModal Or vbExclamation
	WScript.Quit
End If

'WScript.Quit

'2.0 - открываем транзакцию
 session.findById("wnd[0]").maximize
 session.findById("wnd[0]/tbar[0]/okcd").text = "VA22"
 session.findById("wnd[0]").sendVKey 0
 session.findById("wnd[0]/usr/ctxtVBAK-VBELN").text = qtn
 session.findById("wnd[0]").sendVKey 0

'3. Заполняем открытый SAP Quotation
Dim ArticlesExcel, objWorkbook, pmu, TextSheet
Set ArticlesExcel = CreateObject("Excel.Application")
Set objWorkbook = ArticlesExcel.Workbooks.Open(excelFile)
objWorkbook.Sheets("PMU").Activate
Set pmu = objWorkbook.Worksheets("PMU")
Dim iLastRow: iLastRow = CInt(0)
iLastRow =pmu.Range("A" & pmu.Rows.Count).End(xlUp).Row  
'WScript.Echo iLastRow

'4. Вставка цен

'session.findById("wnd[0]/tbar[0]/btn[3]").press()
'session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4426/subSUBSCREEN_TC:SAPMV45A:4908/tblSAPMV45ATCTRL_U_ERF_KONTRAKT/txtVBAP-ARKTX[4,0]").setFocus
'session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4426/subSUBSCREEN_TC:SAPMV45A:4908/tblSAPMV45ATCTRL_U_ERF_KONTRAKT/txtVBAP-ARKTX[4,0]").caretPosition = 2
'session.findById("wnd[0]").sendVKey(2)
'session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05").select()

session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4426/subSUBSCREEN_TC:SAPMV45A:4908/tblSAPMV45ATCTRL_U_ERF_KONTRAKT/ctxtRV45A-MABNR[1,0]").caretPosition = 1
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV45A:4426/subSUBSCREEN_TC:SAPMV45A:4908/tblSAPMV45ATCTRL_U_ERF_KONTRAKT/ctxtRV45A-MABNR[1,0]").setFocus
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_ITEM/tabpT\05").select			'Conditions tab click
WScript.Sleep 300
tblArea = UserArea.findByName("SAPLV69ATCTRL_KONDITIONEN", "GuiTableControl").Id
Set grid = session.findById(tblArea)
objWorkbook.Sheets("PMU").Activate
intRow = 4
Dim iRow
Do Until ArticlesExcel.Cells(intRow, 31).Value = ""			' 31 - for ZLS3 column in the excel file
	' check the value in the Excel file
	condValue = ArticlesExcel.Cells(intRow, 31).Value 
	If Not IsNumeric(condValue)	Then
		MsgBox "Значение в строке" & intRow & "не является числом: " & condValue, vbSystemModal Or vbExclamation
		Continue 	' skip this value 
	End If
	condValue = FormatNumber(condValue,,,,0) 
	condValue = Replace(condValue, ".", ",")
	MsgBox condValue
	' Assign the target condition
	if condValue  >= 0 Then
		targetCondition = "ZLS3"
	else 
		targetCondition = "ZLD3"	
	end if	
	' Цикл для каждой строки в ценовых условиях
	tblArea = UserArea.findByName("SAPLV69ATCTRL_KONDITIONEN", "GuiTableControl").Id
	Set grid = session.findById(tblArea)
	qtyRows = grid.rowCount - 1
	'MsgBox qtyRows
	iRow = 0
	condS = False
	condV = False
	WScript.Sleep 300
	Do Until  condS And condD 'iRow > qtyRows
		'MsgBox "Row: " & intRow
		tblArea = UserArea.findByName("SAPLV69ATCTRL_KONDITIONEN", "GuiTableControl").Id
		Set grid = session.findById(tblArea)
		if grid.GetCell(iRow, 1).Text = "ZLS3" Or grid.GetCell(iRow, 1).Text = "ZLD3" Then
			WScript.Sleep 100
			grid.GetCell(iRow, 1).setFocus()
			grid.GetCell(iRow, 1).caretPosition = 2	
			' get in the condition ZLS3/ZLD3
			session.findById("wnd[0]").sendVKey 2
			session.findById("wnd[0]/usr/txtKOMV-KBETR").text =  "0" 
			session.findById("wnd[0]/usr/txtKOMV-KBETR").caretPosition = 15
			session.findById("wnd[0]").sendVKey(0)
			session.findById("wnd[0]/tbar[0]/btn[3]").press()		
			If grid.GetCell(iRow, 1).Text = "ZLS3" Then
				condS = True
			End If
			If grid.GetCell(iRow, 1).Text = "ZLD3" Then
				condD = True
			End If
		End If	
		tblArea = UserArea.findByName("SAPLV69ATCTRL_KONDITIONEN", "GuiTableControl").Id
		Set grid = session.findById(tblArea)		
		if grid.GetCell(iRow, 1).Text = targetCondition  Then
			WScript.Sleep 100
			grid.GetCell(iRow, 1).setFocus()
			grid.GetCell(iRow, 1).caretPosition = 2	
			' get in the condition ZLS3
			session.findById("wnd[0]").sendVKey 2
			session.findById("wnd[0]/usr/txtKOMV-KBETR").text =  condValue 
			session.findById("wnd[0]/usr/txtKOMV-KBETR").caretPosition = 15
			session.findById("wnd[0]").sendVKey(0)
			session.findById("wnd[0]/tbar[0]/btn[3]").press()		
		End If	
		iRow = iRow + 1
	Loop
	session.findById("wnd[0]").sendVKey(0)
	session.findById("wnd[0]/tbar[1]/btn[19]").press()	
	intRow = intRow + 1

Loop

session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[11]").press
'session.findById("wnd[1]/usr/btnSPOP-VAROPTION1").press

objWorkbook.Close False
ArticlesExcel.Quit
MsgBox "Script finished! ", vbSystemModal Or vbInformation

'====== Functions ans Subs ========

'returns an unique array from an Excel file chosen by a user
Function GetExcelArray()
	'excelFile = "C:\VBScript\articles.xlsx" ' Полный путь к выбранному файлу
	Dim ArticlesExcel, objWorkbook, ws
	Set ArticlesExcel = CreateObject("Excel.Application")
	Set objWorkbook = ArticlesExcel.Workbooks.Open(excelFile)
	objWorkbook.Sheets("PMU").Activate
	Set ws = objWorkbook.Worksheets("PMU")
	Dim collTemp : Set collTemp = CreateObject("Scripting.Dictionary")

	Dim iLastRow: iLastRow = CInt(0)
   	iLastRow = ws.Range("A" & ws.Rows.Count).End(xlUp).Row  
	'WScript.Echo iLastRow

	' Считаем, что в 4 строке - начало таблицы для обработки
	Dim intRow : intRow = 4
	Dim iCol
	' Цикл для каждой строки
	On Error Resume Next
	Do Until ArticlesExcel.Cells(intRow, firstCol).Value = ""
		'ReDim Preserve arrExcel(intRow - 4, 6)
		'WScript.Echo ArticlesExcel.Cells(intRow, firstCol).Value
		For iCol = firstCol to lastCol
			arrExcel(intRow - 4, iCol - firstCol) = ArticlesExcel.Cells(intRow, iCol).Value
		Next 
		WScript.Echo arrExcel(intRow - 4, 0)
		intRow = intRow + 1
	Loop
	objWorkbook.Close False
	ArticlesExcel.Quit
	WScript.Echo Join(arrExcel)
	GetExcelArray = arrExcel
End Function





Sub OutputToExcel
	Dim ReportExcel, objWorkbook
	Set ReportExcel = CreateObject("Excel.Application")
	Set objWorkbook = ReportExcel.Workbooks.Add()
	ReportExcel.Visible = True
	
	arrReport = dicReport.Items
	intRow = 0
	For Each serno In arrSerno
		strReport = strReport & serno & " : " & arrReport(intRow) & VbCrLf
		ReportExcel.cells(intRow + 1, 1).value = serno
		ReportExcel.cells(intRow + 1, 2).value = arrReport(intRow)
		intRow = intRow + 1
	Next
End Sub



Sub notused


'StartTransaction("ZIB07")
session.findById("wnd[0]/tbar[0]/okcd").text = "ZIB07"
session.findById("wnd[0]").sendVKey 0

bAbort = vbFalse
For Each serno In arrSerno
	bExit = vbFalse
	session.findById("wnd[0]/usr/ctxtP_EQUNR").text = serno
	session.findById("wnd[0]/usr/ctxtP_WERKS2").text = plant
	session.findById("wnd[0]/tbar[1]/btn[8]").press
	WScript.Sleep 500 'Delay for SAP processing
	If session.findById("wnd[0]/usr/ctxtP_EQUNR", False) Is Nothing Then
		Do While session.findById("wnd[0]/usr/chkJOB", False) Is Nothing
			If session.findById("wnd[1]/usr/txtLV_MATNR1", False) Is Nothing Then
				dicReport.Add serno, resNoBOM
				bExit = vbTrue
				session.findById("wnd[1]").sendVKey 0
				Exit Do
			Else
				session.findById("wnd[1]/tbar[0]/btn[8]").press 'V
				'session.findById("wnd[1]/tbar[0]/btn[2]").press       'X        
			End If
		Loop
		
		If Not bExit Then
			session.findById("wnd[0]/usr/chkJOB").selected = False
			session.findById("wnd[0]/usr/chkJOB").setFocus
			
			Set grid = session.findById("wnd[0]/usr/cntlEXTEND/shellcont/shell")
			
			qtyRows = grid.rowCount - 1
			'MsgBox "Rows amount: " & qtyRows
			visibleRows = grid.VisibleRowCount

			' Цикл для каждой строки
			'On Error Resume Next
			intRow = 0
			Do Until intRow > qtyRows
				'Err.Clear
				'MsgBox "Row: " & intRow
				grid.modifyCell intRow, "TEMPLATE", template
				grid.currentCellRow = intRow
				intRow = intRow + 1
			Loop
			grid.triggerModified
			session.findById("wnd[0]/tbar[1]/btn[8]").press
			'    MsgBox "Next Control - btn[3]", vbSystemModal Or vbInformation

			' It can be error that mat number not found - If for that
			If session.findById("wnd[1]/tbar[0]/btn[0]", False) Is Nothing Then
			Else
				bAbort = vbTrue
				dicReport.Add serno, template & resNoTemplate
				session.findById("wnd[1]/tbar[0]/btn[0]").press
			End If
			
			If Not bAbort Then
				session.findById("wnd[0]/tbar[0]/btn[3]").press
				'    MsgBox "Next Control - wnd[1]/tbar[0]/btn[0]", vbSystemModal Or vbInformation
				session.findById("wnd[1]/tbar[0]/btn[0]").press
				dicReport.Add serno, resOK
			End If
		End If
	Else
		' Same selection window - check for status bar
		If session.ActiveWindow.findById("sbar", False) Is Nothing Then
			dicReport.Add serno, resExists
		Else
			txtStatus = session.ActiveWindow.findById("sbar").Text
			dicReport.Add serno, txtStatus
		End If
		
	End If
Next

OutputToExcel

End Sub